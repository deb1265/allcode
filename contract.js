const fs = require('fs');
const fs1 = require('fs').promises;
const puppeteer = require('puppeteer');
const { google } = require('googleapis');
const path = require('path');
//const { PDFDocument } = require('pdf-lib');
const PDFDocument = require('pdf-lib').PDFDocument;
const Promise = require('bluebird');
//const keys = require('./servicekey.json');
const axios = require('axios');
const cheerio = require('cheerio');
const moment = require('moment');
const qs = require('querystring');
const caspioTokenUrl = 'https://c1abd578.caspio.com/oauth/token';
const clientID = process.env.clientID;
const clientSecret = process.env.clientSecret;
const GoggleApi = process.env.GoggleApi;
const keys =JSON.parse(process.env.SERVICE_KEY);
const pdfPoppler = require('pdf-poppler');
const nycsformsfileId = '1SlkgPpX_xCwph-Nlf0ej-vCaMy8sZhv7';
const sharp = require('sharp');
const delay = (ms) => new Promise((resolve) => setTimeout(resolve, ms));
const nycblanksdownloadedpath = path.resolve('/tmp', 'file.pdf'); 
const contractPath = path.resolve('/tmp', 'contract.pdf'); 
async function deleteFileWithRetry(filePath, retryCount = 5) {
    for (let attempt = 0; attempt < retryCount; attempt++) {
        try {
            fs.unlinkSync(filePath);
            fs.unlinkSync('filed.pdf');
            fs.unlinkSync('contract-17.png');
            fs.unlinkSync('contract.pdf');
            console.log(`Successfully deleted ${filePath}`);
            return;  // exit the function if deletion is successful
        } catch (error) {
            console.error(`Error deleting file on attempt ${attempt + 1}:`, error);
            await new Promise(resolve => setTimeout(resolve, 1000));  // wait for 1 second before next attempt
        }
    }
    console.error(`Failed to delete ${filePath} after ${retryCount} attempts`);
}
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
async function fetchSpreadsheetData(auth) {
    const sheets = google.sheets({ version: 'v4', auth });
    const spreadsheetId = '1aXRq3yh7KMwFJmTBWGL74Hp56Vns8XGcsqwznIkO2tc';
    const range = 'Sheet1!A2:A20'; // adjust this range to include all the rows in the column that contain data

    let response = await sheets.spreadsheets.values.get({
        spreadsheetId,
        range,
    });

    let rows = response.data.values;
    let custIDs = [];
    if (rows.length) {
        rows.forEach(row => {
            let CustID = row[0];
            if (CustID) { // Check if CustID is not null, undefined, etc.
                custIDs.push(CustID); // push each CustID into the array
            }
        });
        return custIDs; // returning the array
    } else {
        console.log('No data found.');
        return null;
    }
}

async function fetchCaspioData(accessToken, url, updateData = null) {
    try {
        if (updateData) {
            const response = await axios.put(url, updateData, {
                headers: {
                    'Authorization': `Bearer ${accessToken}`
                }
            });
            return response.data;
        } else {
            const response = await axios.get(url, {
                headers: {
                    'Authorization': `Bearer ${accessToken}`
                }
            });
            return response.data;
        }
    } catch (error) {
        console.error(`Error in fetchCaspioData: ${error}`);
        throw error;
    }
}

async function caspiomain(CIDvalue, accessToken) {
    try {
        const masterTableUrl = `https://c1abd578.caspio.com/rest/v2/tables/MasterTable/records?q.select=Name%2C%20Address1%2C%20Phone1%2C%20email&q.where=CID=${CIDvalue}`;
        const solarProcessUrl = `https://c1abd578.caspio.com/rest/v2/tables/SolarProcess/records?q.select=Contract_Price%2CPanel%2CSystem_Size_Sold%2CBlock%2CLot%2CBin%2CFlood_Zone%2CNYC_LI%2COath_Ecb_Violation%2CFlood_Zone%2CLandmark%2CBorough%2CStories%2CCB&q.where=CustID=${CIDvalue}`;
        const signedcontractsurl = `https://c1abd578.caspio.com/rest/v2/tables/Signed_contracts/records?q.select=DownloadLink&q.where=CustID=${CIDvalue}`;
        const masterDataResponse = await fetchCaspioData(accessToken, masterTableUrl);
		console.log(masterDataResponse);
        const solarProcessResponse = await fetchCaspioData(accessToken, solarProcessUrl);
		console.log(solarProcessResponse);
        const signedcontractResponse = await fetchCaspioData(accessToken, signedcontractsurl);
		console.log(signedcontractResponse);
        const signedData = signedcontractResponse.Result[signedcontractResponse.Result.length - 1];
        const masterData = masterDataResponse.Result[0];
        const solarData = solarProcessResponse.Result[0];
        const combinedData = { ...masterData, ...solarData, ...signedData };
        let contractlink = combinedData.DownloadLink;
        let contractfileid = contractlink.split('id=')[1].split('&export=download')[0];
        const originalAddress = combinedData.Address1;
        const response = await axios.get(`https://maps.googleapis.com/maps/api/geocode/json?address=${encodeURIComponent(originalAddress)}&key=${GoggleApi}`);
        const Address = response.data.results[0].formatted_address;
        const Contract_Price = combinedData.Contract_Price;
        const Panel = combinedData.Panel;
        const Size = combinedData.System_Size_Sold;
        const Block = combinedData.Block;
        const Lot = combinedData.Lot;
        const Bin = combinedData.Bin;
        const CB = combinedData.CB;
        const Flood_Zone = combinedData.Flood_Zone;
        const Oath_Ecb_Violation = combinedData.Oath_Ecb_Violation;
        const Landmark = combinedData.Landmark;
        const Borough = combinedData.Borough;
        const lead = combinedData.Lead;
        const NYC_LI = combinedData.NYC_LI;
        const Stories = combinedData.Stories;
        let FullName = combinedData.Name; // Replace with your actual full name variable
        let nameParts = FullName.split(' ');
        let Firstname;
        let Lastname;
        if (nameParts.length >= 2) {
            Firstname = nameParts[0];
            Lastname = nameParts[nameParts.length - 1];
        } else if (nameParts.length === 1) {
            Firstname = nameParts[0];
            Lastname = "";  // No last name available
        } else {
            console.error(`Could not parse first and last name from ${FullName}`);
        }
        let Addressp = Address;
        let addressParts = Addressp.split(',');
        // Get the house and street
        let houseStreetParts = addressParts[0].trim().split(' ');
        let houseNumberRegex = /^[\d\-A-Za-z]+/;  // Matches digits, hyphens and letters at the start
        let houseNumberMatch = houseStreetParts[0].match(houseNumberRegex);
        let House;
        let Street;
        if (houseNumberMatch) {
            House = houseNumberMatch[0];
            Street = houseStreetParts.slice(1).join(' ');
        } else {
            console.error(`Could not parse house number and street from ${addressParts[0]}`);
        }
        // Get the city
        let City = addressParts[1].trim();
        // Get the zip
        let Zip = addressParts[3].trim();
        // Get the zip
        // Assuming that the address format is always "House, City, State, Zip"
        // Zip would be in the 3rd index position after splitting
        return { Firstname, Lastname, FullName, Address, House, Street, City, lead, Zip, Landmark, CB, Panel, NYC_LI, Size, Stories, Borough, contractfileid, Contract_Price, Block, Lot, Bin, Flood_Zone, Oath_Ecb_Violation };
    } catch (error) {
        console.error(`Error in main: ${error}`);
        if (error.response) {
            console.error('Error status:', error.response.status);
            console.error('Error data:', error.response.data);
        }
    }
}
async function downloadFile(drive, fileId, downloadpath) { 
    //const fileId = '1tSOSVstd8BpGG9mNF5koWzq9M8OQPiGZ';  // replace with your file id
    const dest = fs.createWriteStream(downloadpath); // use nycblanksdownloadedpath here
    const res = await drive.files.get({ fileId, alt: 'media' }, { responseType: 'stream' });
    return new Promise((resolve, reject) => {
        res.data
            .on('end', async () => {
                console.log('File download completed.');
                // Add a delay here to ensure file is fully written to disk
                await new Promise(resolve => setTimeout(resolve, 2000)); // 2 seconds delay
                resolve();
            })
            .on('error', err => {
                console.error('Error downloading file.');
                reject(err);
            })
            .pipe(dest);
    });
}
(async () => {
    const client = new google.auth.JWT(
        keys.client_email,
        null,
        keys.private_key,
        [
            'https://www.googleapis.com/auth/spreadsheets',
            'https://www.googleapis.com/auth/drive',
            'https://www.googleapis.com/auth/drive.readonly'
        ],
    );
    const drive = google.drive({ version: 'v3', auth: client }); // Note the change here
    await client.authorize();
    const accessToken = await caspioAuthorization();  
    console.log("Successfully connected to Google Sheets!");
    console.log("Successfully connected to Caspio!");  
    const sheetData = await fetchSpreadsheetData(client);
    for (let i = 0; i < sheetData.length; i++) {
        const custId = process.argv[2];
        const filePath = `/tmp/contract${custId}.pdf`;  // Use backticks
        const caspioData = await caspiomain(custId, accessToken);
        const name = caspioData.FullName;
        const contractFileId = caspioData.contractfileid;
        console.log(`Starting file download for contractFileId: ${contractFileId}`);
        await downloadFile(drive, contractFileId, filePath);
		        console.log('File download completed.');
        await delay(5000);
        console.log('5 seconds have passed.');
		  await delay(5000);  // Wait for 5 seconds
    const pdfDoc = await PDFDocument.load(fs.readFileSync(filePath)); // make sure filePath is defined
    console.log('5 seconds have passed.');
        console.log('Checking PNG path...');
        const pngPath = './public/boss_sign.PNG'; 
        //console.log(`PNG path is: ${pngPath}`);
		const oldPdfDoc = await PDFDocument.load(await fs1.readFile(filePath));
        const newPdfDoc = await PDFDocument.create(); // Create a new PDF for pages 4 to 7
        // Copy pages 4 to 7
        for (let i = 3; i <= 6; i++) {
        const [copiedPage] = await newPdfDoc.copyPages(oldPdfDoc, [i]);
        newPdfDoc.addPage(copiedPage);
        }
        const pngBuffer = await fs1.readFile('./public/boss_sign.PNG'); // Ensure to read as a binary file
        const imagePage = await newPdfDoc.embedPng(pngBuffer); // Embedding PNG in PDF
		const dateOfSign = moment().subtract(7, 'days').format('MM/DD/YYYY');  // Generating the date
        console.log(pngBuffer.slice(0, 10));

		const pagesData = [
                            //{ pageNum: 1, imageX: 4.5909 * 72, imageY: 6.3707 * 72, imageWidth: 1.5924 * 72, imageHeight: 0.4419 * 72 },
                             { pageNum: 4, imageX: 1.2 * 72, imageY: 2.0707 * 72, imageWidth: 1.0818 * 72, imageHeight: 0.4929 * 72 },
                            //{ pageNum: 4, imageX: 2.6091 * 72, imageY: 7.4525 * 72, imageWidth: 1.0818 * 72, imageHeight: 0.302 * 72 },
                            //{ pageNum: 18, imageX: 0.6091 * 72, imageY: 2.2015 * 72, imageWidth: 1.5924 * 72, imageHeight: 0.4328 * 72 },
                        ];
                        for (let i = 0; i < pagesData.length; i++) {
							
                            const pageData = pagesData[i];
                                const page = newPdfDoc.getPages()[pageData.pageNum - 1];
                            page.drawImage(imagePage, {
                                x: pageData.imageX,
                                y: pageData.imageY, // from bottom of page
                                width: pageData.imageWidth,
                                height: pageData.imageHeight
                            });
							console.log(`Embedding PNG into page ${pageData.pageNum} at X: ${pageData.imageX}, Y: ${pageData.imageY}`);
							console.log('Embedded PNG into PDF');
							    // Draw the date
    const dateX = pageData.imageX + pageData.imageWidth + 5;  // X-coordinate right next to the image
    const dateY = pageData.imageY + (pageData.imageHeight / 2); // Y-coordinate aligned to the middle of image
        page.drawText(dateOfSign, {
        x: dateX,
        y: dateY,
        size: 12  // You can adjust the size
    });
	}
    const pdfBytes = await newPdfDoc.save();
    await fs1.writeFile(`/tmp/${name}_contract.pdf`, pdfBytes);
  }
})().catch(err => {
  console.error(err);
});