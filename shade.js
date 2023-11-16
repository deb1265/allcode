const { DocumentProcessorServiceClient } = require('@google-cloud/documentai').v1;
const vision = require('@google-cloud/vision');
const { google } = require('googleapis');
const PDFServicesSdk = require('@adobe/pdfservices-node-sdk');
const path = require('path');
const pdf = require('pdf-poppler');
const axios = require('axios');
const https = require('https');
const qs = require('querystring');
const fs = require('fs');
const util = require('util');
const unlinkAsync = util.promisify(fs.unlink);
const exec = util.promisify(require('child_process').exec);
const GoggleApi = 'AIzaSyDPuTFidqOSRkhWfv1BxJP8o4pFWnxhJzQ';
const Cust_ID = process.argv[2];
//const { Cust_ID } = require('./main');  // Import Cust_ID from main.js
//const Cust_ID = '8535';
const caspioFileURL1 = `https://c1abd578.caspio.com/rest/v2/tables/SolarProcess/records?q.select=shadereport&q.where=CustID=${Cust_ID}`;
const caspioTokenUrl = 'https://c1abd578.caspio.com/oauth/token';
const clientID = 'cae94860c0ae4fb5aa60756fba6ddfbce64964705b73ac1cfb';
const clientSecret = 'e3bb6f52ab914dfeaaa7a3a6814864ddfa3c74b49832163f92';
const delay = (ms) => new Promise((resolve) => setTimeout(resolve, ms));
const unzipper = require('unzipper');
const xlsx = require('xlsx');

async function caspioAuthorization() {
    const requestBody = {
        grant_type: 'client_credentials',
        client_id: clientID,
        client_secret: clientSecret,
    };
    try {
        const response = await axios.post(caspioTokenUrl, qs.stringify(requestBody), {
            headers: {
                'Content-Type': 'application/x-www-form-urlencoded'
            }
        });
        return response.data.access_token;
    } catch (error) {
        console.error(`Error in caspioAuthorization: ${error}`);
        throw error;
    }
}

const downloadshade = async (accessToken, Cust_ID, caspioFileURL) => {
    const appKey = '68107000';

    // Extract utility bill property from URL.
    const match = caspioFileURL.match(/q\.select=(shadereport\d?)/);
    if (!match) {
        console.error("Failed to match utility bill property from URL.");
        return null;
    }
    const shadereportproperty = match[1];

    // Fetch the file URL using the customer ID.
    const result = await axios.get(caspioFileURL, { headers: { 'Authorization': `Bearer ${accessToken}` } });
    if (!result.data || !result.data.Result || result.data.Result.length === 0) {
        console.warn('No file path found for given customer ID.');
        return null;
    }

    const filePath = result.data.Result[0][shadereportproperty];
    if (!filePath) {
        console.warn(`${shadereportproperty} not found or is undefined.`);
        return null;
    }
    const sanitizedFileName = filePath.substring(filePath.lastIndexOf('/') + 1).replace(/\s+/g, '_'); // replace spaces with underscores
    const fileExtension = sanitizedFileName.substring(sanitizedFileName.lastIndexOf('.') + 1);
    const localPath = `/tmp/shadereport.${fileExtension}`;

    return new Promise((resolve, reject) => {
        const file = fs.createWriteStream(localPath);
        // Encode the filepath to ensure spaces and other special characters are correctly formatted for a URL
        const encodedFilePath = encodeURIComponent(filePath);

        const requestOptions = {
            hostname: 'cdn.caspio.com',
            path: `/${appKey}${encodedFilePath}?ver=1`,
            headers: {
                'Authorization': `Bearer ${accessToken}`,
                'User-Agent': 'Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/92.0.4515.159 Safari/537.36',
                'Referer': 'https://c1abd578.caspio.com/'
            }
        };
        const request = https.get(requestOptions, (response) => {
            if (response.statusCode !== 200) {
                reject(new Error(`Failed to download. Status code: ${response.statusCode}`));
                response.resume(); // Consume response to free up memory.
                return;
            }

            response.pipe(file);

            file.on('finish', () => {
                file.close(() => {
                    console.log('File saved successfully');
                    resolve(localPath);
                });
            });
        });

        request.on('error', (error) => {
            fs.unlink(localPath, () => { }); // Delete the file if there's an error (optional).
            reject(error.message);
        });

        file.on('error', (error) => {
            fs.unlink(localPath, () => { }); // Delete the file on any error.
            reject(error.message);
        });
    });
}

const deleteFiles = async (mainFile, arrayOfFiles = []) => {

    try {
        // Delete the main file (like the PDF itself)
        fs.unlinkSync(mainFile);
        console.log(`Deleted: ${mainFile}`);

        // If there are other files in the array, delete them too
        for (let file of arrayOfFiles) {
            fs.unlinkSync(file);
            console.log(`Deleted: ${file}`);
        }
    } catch (error) {
        console.error(`Error deleting files: ${error.message}`);
    }
};

async function fetchCaspioData(accessToken, url, data = null, method = 'GET') {
    try {
        let response;
        const headers = {
            'Authorization': `Bearer ${accessToken}`
        };

        switch (method.toUpperCase()) {
            case 'GET':
                response = await axios.get(url, { headers });
                break;
            case 'POST':
                response = await axios.post(url, data, { headers });
                break;
            case 'PUT':
                response = await axios.put(url, data, { headers });
                break;
            default:
                throw new Error(`Unsupported method: ${method}`);
        }

        return response.data;
    } catch (error) {
        console.error(`Error in fetchCaspioData: ${error}`);
        throw error;
    }
}

async function caspiomain(CIDvalue, accessToken) {
    try {
        const masterTableUrl = `https://c1abd578.caspio.com/rest/v2/tables/MasterTable/records?q.select=Name%2C%20Address1%2C%20Phone1%2C%20email&q.where=CID=${CIDvalue}`;
        const solarProcessUrl = `https://c1abd578.caspio.com/rest/v2/tables/SolarProcess/records?q.select=NYC_LI%2C%20Finance%2C%20System_Size_Sold%2C%20Panel%2C%20Number_Of_Panels%2C%20Contract_Price%2C%20Canopy_Used%2C%20Affordable%2C%20Solarinsure%2C%20number_of_systems&q.where=CustID=${CIDvalue}`;

        const masterDataResponse = await fetchCaspioData(accessToken, masterTableUrl);
        const solarProcessResponse = await fetchCaspioData(accessToken, solarProcessUrl);

        const masterData = masterDataResponse.Result[0];
        const solarData = solarProcessResponse.Result[0];

        const combinedData = { ...masterData, ...solarData };

        const originalAddress = combinedData.Address1;

        const response = await axios.get(`https://maps.googleapis.com/maps/api/geocode/json?address=${encodeURIComponent(originalAddress)}&key=${GoggleApi}`);
        const Address = response.data.results[0].formatted_address;
        const CID = CIDvalue;
        const NYC_LI = combinedData.NYC_LI;

        return { CID,Address,NYC_LI};
    } catch (error) {
        console.error(`Error in main: ${error}`);
        if (error.response) {
            console.error('Error status:', error.response.status);
            console.error('Error data:', error.response.data);
        }
    }
}

const findSheetByColumnCount = (workbook, targetColumnCount) => {
    // Iterate over each sheet in the workbook
    for (const sheetName of workbook.SheetNames) {
        // Get the sheet object
        const sheet = workbook.Sheets[sheetName];
        
        // Convert sheet to JSON
        const data = xlsx.utils.sheet_to_json(sheet);
        
        // If sheet has no data, continue to next sheet
        if (!data.length) continue;

        // Get the first row of the sheet
        const firstRow = data[0];
        
        // Count the number of keys (columns) in the first row
        const columnCount = Object.keys(firstRow).length;

        // Check if the number of columns matches the targetColumnCount
        if (columnCount === targetColumnCount) {
            return sheet;
        }
    }
    // Return null if no matching sheet is found
    return null;
};
async function extractTableFromPDF(pdfPath) {
    // Initialize PDF SDK
    const credentials = PDFServicesSdk.Credentials
        .servicePrincipalCredentialsBuilder()
        .withClientId('552a5bdc47b940c3beae375fa4c854e8')
        .withClientSecret('p8e-4-9JfUCo2AHgXPGbmVFbxk7WlPCUEXV_')
        .build();
        const executionContext = PDFServicesSdk.ExecutionContext.create(credentials);

    // Extract PDF options
    const options = new PDFServicesSdk.ExtractPDF.options.ExtractPdfOptions.Builder()
        .addElementsToExtract(PDFServicesSdk.ExtractPDF.options.ExtractElementType.TEXT, PDFServicesSdk.ExtractPDF.options.ExtractElementType.TABLES)
        .build();

    const extractPDFOperation = PDFServicesSdk.ExtractPDF.Operation.createNew();
    const input = PDFServicesSdk.FileRef.createFromLocalFile(pdfPath);
    extractPDFOperation.setInput(input);
    extractPDFOperation.setOptions(options);

    const outputFilePath = `/tmp/temporary_extract_${Date.now()}.zip`; // Generates a unique filename
    let extractedData = null; // Initialize the extractedData variable here

    try {
        await extractPDFOperation.execute(executionContext)
            .then(result => result.saveAsFile(outputFilePath));

        await fs.createReadStream(outputFilePath)
            .pipe(unzipper.Extract({ path: 'temporary_output/' }))
            .promise();

        const files = fs.readdirSync('temporary_output/tables/');
        
        for (const file of files) {
            const workbook = xlsx.readFile(`temporary_output/tables/${file}`);
            const requiredColumns = [
                'Panel Count \r', 'Azimuth (deg.) \r', 'Pitch (deg.) \r',
                'Annual TOF (%) \r', 'Annual Solar Access(%) \r', 'Annual TSRF (%) \r'
            ];
			const targetColumnCount = 7;
            const sheet = findSheetByColumnCount(workbook, targetColumnCount);

      if (sheet) {
        const data = xlsx.utils.sheet_to_json(sheet);
        extractedData = data.map(row => {
          return {
            panelCount: row['Panel Count \r'] ? Number(row['Panel Count \r'].trim()) : null,
            azimuth: row['Azimuth (deg.) \r'] ? Number(row['Azimuth (deg.) \r'].trim()) : null,
            pitch: row['Pitch (deg.) \r'] ? Number(row['Pitch (deg.) \r'].trim()) : null,
            annualTOF: row['Annual TOF (%) \r'] ? Number(row['Annual TOF (%) \r'].trim()) : null,
            annualSolarAccess: row['Annual Solar Access(%) \r'] ? Number(row['Annual Solar Access(%) \r'].trim()) : null,
            annualTSRF: row['Annual TSRF (%) \r'] ? Number(row['Annual TSRF (%) \r'].trim()) : null
          };
        });
        
        // Ensure extractedData has exactly 8 rows
        while (extractedData.length < 8) {
          extractedData.push({
            panelCount: 0,
            azimuth: 0,
            pitch: 0,
            annualTOF: 0,
            annualSolarAccess: 0,
            annualTSRF: 0
          });
        }
        
        // Stop the loop if a sheet is found
        break;
      }
    }

    // Transform extractedData to the required format
    const transformedData = {
      Panelcount: extractedData.map(row => row.panelCount),
      pitch: extractedData.map(row => row.pitch),
      azimuth: extractedData.map(row => row.azimuth),
      annualTSRF: extractedData.map(row => row.annualTSRF)
    };

    return transformedData;
    //console.log(transformedData);
  } catch (err) {
    console.log('Exception encountered:', err);
    return null;
  }
}

(async () => {
    const accessToken = await caspioAuthorization();
    console.log("Successfully connected to Caspio!");
    const filepath = await downloadshade(accessToken, Cust_ID, caspioFileURL1);  
    await delay(5000); // Delay for 5 seconds
    const fillArray = (arr, length) => {
        while (arr.length < length) {
            arr.push(0);
        }
        return arr;
    };

    let data = await extractTableFromPDF(filepath);
    console.log(data);
    let updatedData;
	
    if (data) {
        const panelCounts = fillArray(data.Panelcount, 8);
        const pitches = fillArray(data.pitch, 8);
        const azimuths = fillArray(data.azimuth, 8);
        const annualTSRFs = fillArray(data.annualTSRF, 8);
    updatedData = {
        "Panel1_": panelCounts[0] || 0,
		"Panel2_": panelCounts[1] || 0,
		"Panel3_": panelCounts[2] || 0,
		"Panel4_": panelCounts[3] || 0,
		"Panel5_": panelCounts[4] || 0,
		"Panel6_": panelCounts[5] || 0,
		"Panel7_": panelCounts[6] || 0,
		"Panel8_": panelCounts[7] || 0,
		"Tilt1_": pitches[0] || 0,
		"Tilt2_": pitches[1] || 0,
		"Tilt3_": pitches[2] || 0,
		"Tilt4_": pitches[3] || 0,
		"Tilt5_": pitches[4] || 0,
		"Tilt6_": pitches[5] || 0,
		"Tilt7_": pitches[6] || 0,
		"Tilt8_":pitches[7] || 0,
        "Az1_": azimuths[0] || 0,
		"Az2_": azimuths[1] || 0,
		"Az3_": azimuths[2] || 0,
		"Az4_": azimuths[3] || 0,
		"Az5_": azimuths[4] || 0,
		"Az6_": azimuths[5] || 0,
		"Az7_": azimuths[6] || 0,
		"Az8_": azimuths[7] || 0,
		"Tsrf1_": annualTSRFs[0] || 0,
		"Tsrf2_": annualTSRFs[1] || 0,
		"Tsrf3_": annualTSRFs[2] || 0,
		"Tsrf4_": annualTSRFs[3] || 0,
		"Tsrf5_": annualTSRFs[4] || 0,
		"Tsrf6_": annualTSRFs[5] || 0,
		"Tsrf7_": annualTSRFs[6] || 0,
		"Tsrf8_": annualTSRFs[7] || 0,
    };
     console.log(updatedData);
    } else {
        console.log('Data is undefined or null');
    }

    const url = `https://c1abd578.caspio.com/rest/v2/tables/SolarProcess/records?q.where=CustID=${Cust_ID}`;
	
try {
    const response = await fetchCaspioData(accessToken, url, updatedData, 'PUT');
    console.log("Successfully updated Caspio data:", response);
  } catch (error) {
    console.error(`Failed to update Caspio data: ${error.message}`);

    // If AxiosError, log the status code as well
    if (error.response) {
      console.error(`AxiosError: Request failed with status code ${error.response.status}`);
    }
  }
})();