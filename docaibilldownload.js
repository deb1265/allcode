const { DocumentProcessorServiceClient } = require('@google-cloud/documentai').v1;
const vision = require('@google-cloud/vision');
const { google } = require('googleapis');
const path = require('path');
const pdf = require('pdf-poppler');
const axios = require('axios');
const https = require('https');
const qs = require('querystring');
const fs = require('fs');
const util = require('util');
const unlinkAsync = util.promisify(fs.unlink);
const exec = util.promisify(require('child_process').exec);
const Cust_ID = process.argv[2];
//const { Cust_ID } = require('./main');  // Import Cust_ID from main.js
//const Cust_ID = '7408';
const caspioFileURL1 = `https://c1abd578.caspio.com/rest/v2/tables/SolarProcess/records?q.select=Utility_bills&q.where=CustID=${Cust_ID}`;
const caspioFileURL2 = `https://c1abd578.caspio.com/rest/v2/tables/SolarProcess/records?q.select=Utility_bills2&q.where=CustID=${Cust_ID}`;
const caspioFileURL3 = `https://c1abd578.caspio.com/rest/v2/tables/SolarProcess/records?q.select=Utility_bills3&q.where=CustID=${Cust_ID}`;
const caspioTokenUrl = 'https://c1abd578.caspio.com/oauth/token';
const clientID = process.env.clientID;
const clientSecret = process.env.clientSecret;
const GoggleApi = process.env.GoggleApi;
const keys =JSON.parse(process.env.SERVICE_KEY);
const delay = (ms) => new Promise((resolve) => setTimeout(resolve, ms));
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

function extract(regex, text) {
    const match = text.match(regex);
    return match ? match[1] : null; // Return the first captured group, or null if not found
}


const downloadBills = async (accessToken, Cust_ID, caspioFileURL) => {
    const appKey = '68107000';
    // Extract utility bill property from URL.
    const match = caspioFileURL.match(/q\.select=(Utility_bills\d?)/);
    if (!match) {
        console.error("Failed to match utility bill property from URL.");
        return null;
    }
    const utilityBillProperty = match[1];
    // Fetch the file URL using the customer ID.
    const result = await axios.get(caspioFileURL, { headers: { 'Authorization': `Bearer ${accessToken}` } });
    if (!result.data || !result.data.Result || result.data.Result.length === 0) {
        console.warn('No file path found for given customer ID.');
        return null;
    }
    const filePath = result.data.Result[0][utilityBillProperty];
    if (!filePath) {
        console.warn(`${utilityBillProperty} not found or is undefined.`);
        return null;
    }
    const sanitizedFileName = filePath.substring(filePath.lastIndexOf('/') + 1).replace(/\s+/g, '_'); // replace spaces with underscores
    const fileExtension = sanitizedFileName.substring(sanitizedFileName.lastIndexOf('.') + 1);
    const localPath = `/tmp/bills.${fileExtension}`;
    //const caspioFileDownloadURL = `https://cdn.caspio.com/${appKey}${filePath}?ver=1`;
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


async function detectTextInImage(imagePath) {
    if (!imagePath) {
        console.warn("Image path is null or undefined in detectTextInImage.");
        return null;
    }

    const fileExtension = path.extname(imagePath).toLowerCase();

    let mimeType;
    switch (fileExtension) {
        case '.pdf':
            mimeType = 'application/pdf';
            break;
        case '.jpg':
        case '.jpeg':
            mimeType = 'image/jpeg';
            break;
        case '.png':
            mimeType = 'image/png';
            break;
        default:
            console.error(`Unsupported file extension: ${fileExtension}`);
            return null;
    }

    // Create a client
    const client = new DocumentProcessorServiceClient({
        keyFilename: './servicekey.json'
    });
    const processorEndpoint = 'projects/206322115501/locations/us/processors/8f5300d55dbf4b4b';

    // Read the file into memory
    const imageFile = await fs.promises.readFile(imagePath);

    // Convert the image data to a Buffer and base64 encode it
    const encodedImage = Buffer.from(imageFile).toString('base64');

    // Create the request
    const request = {
        name: processorEndpoint,
        inlineDocument: {
            content: encodedImage,
            mimeType: mimeType, // Use the determined mime type
        },
    };

    // Process the document
    const [result] = await client.processDocument(request);

    return result; // Modify as needed to return the desired data
console.log(result);
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

async function processResult(caspioFileURL, client1, accessToken, Utility) {
    let imagePath;
    let accountNumber = null;
    let meterNumber = null;
    let name = null;

    try {
        const downloadResult = await downloadBills(accessToken, Cust_ID, caspioFileURL);

        if (!downloadResult) {
            console.warn("Download result is null or undefined. Skipping processing for this result.");
            return { accountNumber, meterNumber, name };
        }

        imagePath = downloadResult;
        const fileExtension = path.extname(imagePath).toLowerCase();

        if (['.pdf', '.jpg', '.jpeg', '.png'].includes(fileExtension)) {
            console.log(`Detected a ${fileExtension} file. Analyzing it directly...`);
            const resultText = await detectTextInImage(imagePath);
            console.log(resultText);
            // Process resultText to extract accountNumber, meterNumber, and name
            // You can replace this with actual logic to parse the text and extract the information
            ({ accountNumber, meterNumber, name } = extractDetailsFromText(resultText,Utility));
        } else {
            console.error("Unsupported file type. Can't process further.");
            return { accountNumber, meterNumber, name };
        }

        // Clean up files
        await deleteFiles(imagePath);
        await delay(7000);
    } catch (error) {
        console.error(`Error processing result: ${error.message}`);
    }
    return { accountNumber, meterNumber, name };
}

function extractDetailsFromText(result, Utility) {
    // Ensure result is in expected format (e.g., a string)
    let text = result.document.text; // Modify this line based on the actual structure of result

    if (typeof text !== 'string') {
        console.error("Unexpected result format. Cannot extract details.");
        return { accountNumber: null, meterNumber: null, name: null };
    }

    console.log("Text is of type string, proceeding to extract details.");

    let accountNumberMatch;
    let meterNumberMatch;
    let nameMatch;
    let meterNumber;

    if (Utility === 'CONED') {
        accountNumberMatch = text.match(/Account Number: (\d+-\d+-\d+-\d+-\d+)/) || text.match(/Account Number: ([\d-]+)/) || text.match(/Account number: (\d{2}-\d{4}-\d{4}-\d{4}-\d)/);
        meterNumberMatch = text.match(/\b\d{9}\b/);

        if (meterNumberMatch) {
            meterNumber = meterNumberMatch[0];
        } else {
            meterNumberMatch = text.match(/Meter #\n(\d+)/) || text.match(/Mater\n(\d{9})/) || text.match(/Moter\n(\d{9})/);
            meterNumber = meterNumberMatch ? meterNumberMatch[1] : null;
        }

        nameMatch = text.match(/Name: (\w+\s\w+)/) || text.match(/EconEdison\n([A-Z\s]+)\nAccount Number:/) || text.match(/Name: ([A-Z&\s]+[A-Z]+)/);
    } else {
        accountNumberMatch = text.match(/Account #: (\d+)/);
        nameMatch = text.match(/Service To: ([^\n]+)/);
    }

    let accountNumber = accountNumberMatch ? accountNumberMatch[1] : null;
    let name = nameMatch ? nameMatch[1] : null;

    console.log("Account Number:", accountNumber);
    console.log("Meter Number:", meterNumber);
    console.log("Name:", name);

    return { accountNumber, meterNumber, name };
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

(async () => {
    const client1 = new DocumentProcessorServiceClient({
        keyFilename: './servicekey.json'
    });
    const accessToken = await caspioAuthorization();
    console.log("Successfully connected to Google Sheets!");
    console.log("Successfully connected to Caspio!");
    let caspioData = await caspiomain(Cust_ID, accessToken);
    const NYC_LI = caspioData.NYC_LI;
    const Utility = NYC_LI === 'NYC' ? 'CONED' : 'PSEG';
    console.log("Utility:", Utility);
    const results = [];
    try {
        results.push(await processResult(caspioFileURL1, client1, accessToken, Utility));
    } catch (error) {
        console.error("Error processing caspioFileURL1:", error.message);
        results.push(null);
    }
    try {
        results.push(await processResult(caspioFileURL2, client1, accessToken, Utility));
    } catch (error) {
        console.error("Error processing caspioFileURL2:", error.message);
        results.push(null);
    }
    try {
        results.push(await processResult(caspioFileURL3, client1, accessToken, Utility));
    } catch (error) {
        console.error("Error processing caspioFileURL3:", error.message);
        results.push(null);
    }
    const cleanedResults = results.map(result => ({
        accountNumber: result.accountNumber ? (result.accountNumber.match(/[\d]/g) || []).join('') : null,
        meterNumber: result.meterNumber ? (result.meterNumber.match(/[\d]/g) || []).join('') : null,
        name: result.name != null ? result.name : null
    }));

    console.log("Original Results Array:", results);
    console.log("Cleaned Results Array:", cleanedResults);

    const updatedData = {
        "InterconnectionAcc1": cleanedResults[0] ? cleanedResults[0].accountNumber : null,
        "InterconnectionAcc2": cleanedResults[1] ? cleanedResults[1].accountNumber : null,
        "InterconnectionAcc3": cleanedResults[2] ? cleanedResults[2].accountNumber : null,
        "InterconnectionMeter1": cleanedResults[0] ? cleanedResults[0].meterNumber : null,
        "InterconnectionMeter2": cleanedResults[1] ? cleanedResults[1].meterNumber : null,
        "InterconnectionMeter3": cleanedResults[2] ? cleanedResults[2].meterNumber : null,
        "InterconnectionName1": cleanedResults[0] ? cleanedResults[0].name : null,
        "InterconnectionName2": cleanedResults[1] ? cleanedResults[1].name : null,
        "InterconnectionName3": cleanedResults[2] ? cleanedResults[2].name : null
    };
    console.log(updatedData);
    const url = `https://c1abd578.caspio.com/rest/v2/tables/SolarProcess/records?q.where=CustID=${Cust_ID}`;
    try {
        const response = await fetchCaspioData(accessToken, url, updatedData, 'PUT');
        console.log("Successfully updated Caspio data:", response);
    } catch (error) {
        console.error("Failed to update Caspio data:", error.message);
    }
})();