const puppeteer = require('puppeteer');
const { google } = require('googleapis');
//const keys = require('./servicekey.json');
const axios = require('axios');
const cheerio = require('cheerio');
const fs = require('fs');
const qs = require('querystring');
const caspioTokenUrl = 'https://c1abd578.caspio.com/oauth/token';
const clientID = process.env.clientID;
const clientSecret = process.env.clientSecret;
const GoggleApi = process.env.GoggleApi;
const keys =JSON.parse(process.env.SERVICE_KEY);
const { spawn } = require('child_process');
const { PDFDocument } = require('pdf-lib');
const path = require('path');
const pdfParse = require('pdf-parse');

function sleep(ms) {
    return new Promise(resolve => setTimeout(resolve, ms));
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
        // Define URLs
        const masterTableUrl = `https://c1abd578.caspio.com/rest/v2/tables/MasterTable/records?q.select=Name%2C%20Address1%2C%20Phone1%2C%20email&q.where=CID=${CIDvalue}`;
        const solarProcessUrl = `https://c1abd578.caspio.com/rest/v2/tables/SolarProcess/records?q.select=Contract_Price%2CPanel%2CSystem_Size_Sold%2CBlock%2CLot%2CBin%2CFlood_Zone%2CNYC_LI%2COath_Ecb_Violation%2CFlood_Zone%2CLandmark%2CBorough%2CStories%2CCB&q.where=CustID=${CIDvalue}`;
        const signedcontractsurl = `https://c1abd578.caspio.com/rest/v2/tables/Signed_contracts/records?q.select=DownloadLink&q.where=CustID=${CIDvalue}`;
const getElectricPermitUrl = `https://c1abd578.caspio.com/rest/v2/tables/electricalpermit/records?q.select=applicationLink%2C%20ApplicationNumber%2C%20passdate&q.where=CID=${CIDvalue}`;
        // Fetch Data
        const masterDataResponse = await fetchCaspioData(accessToken, masterTableUrl);
        const solarProcessResponse = await fetchCaspioData(accessToken, solarProcessUrl);
        const signedcontractResponse = await fetchCaspioData(accessToken, signedcontractsurl);
        const electricpermitResponse = await fetchCaspioData(accessToken, getElectricPermitUrl);
         const electricData = electricpermitResponse.Result[0];
//console.log(masterDataResponse);
//console.log(solarProcessResponse);
//console.log(signedcontractResponse);
        // Initialize default values
        let contractlink = "";
        let contractfileid = "";

        // Check and process signedcontractResponse
        if (signedcontractResponse?.Result?.length > 0) {
    const signedData = signedcontractResponse.Result[signedcontractResponse.Result.length - 1];
    contractlink = signedData?.DownloadLink;

    // Attempt to extract contractfileid using the first format
    if (contractlink) {
        try {
            contractfileid = contractlink.split('id=')[1]?.split('&export=download')[0];
        } catch (e) {
            console.error("Error extracting contractfileid using first format", e.message);
            contractfileid = ""; // Reset to empty if first method fails
        }

        // If the first format fails, try the second format
        if (!contractfileid) {
            const linkParts = contractlink.split('/d/');
            if (linkParts.length > 1) {
                const idPart = linkParts[1].split('/')[0];
                contractfileid = idPart;
            }
        }
    }
}

        // Process masterDataResponse and solarProcessResponse
        const masterData = masterDataResponse?.Result[0];
        const solarData = solarProcessResponse?.Result[0];

        // Combine Data
        const combinedData = { ...masterData, ...solarData, ...{ DownloadLink: contractlink, ContractFileID: contractfileid },...electricData };
//console.log(combinedData);
        // Extract and process further data
        // Extract and process further data
let FullName = masterData?.Name || "Unknown"; // Default to "Unknown" if Name is not available
let Firstname = "";
let Lastname = "";
        let applicationLink = combinedData.applicationLink;
		console.log(applicationLink);
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
		const Contract_Price = combinedData.Contract_Price;
// Split FullName into Firstname and Lastname
if (FullName) {
    let nameParts = FullName.split(' ');
    if (nameParts.length >= 2) {
        Firstname = nameParts[0];
        Lastname = nameParts[nameParts.length - 1];
    } else if (nameParts.length === 1) {
        Firstname = nameParts[0];
        // Lastname remains an empty string as it's not available
    }
}

// Process Address
const originalAddress = masterData?.Address1 || "";
let Address = "";
let House = "";
let Street = "";
let City = "";
let Zip = "";
console.log(originalAddress);
if (originalAddress) {
    try {
        const response = await axios.get(`https://maps.googleapis.com/maps/api/geocode/json?address=${encodeURIComponent(originalAddress)}&key=${GoggleApi}`);
        Address = response.data.results[0]?.formatted_address || originalAddress; // Fallback to originalAddress if no result
        //console.log(response);
        // Further processing for House, Street, etc.
        let addressParts = Address.split(',');
        let houseStreetParts = addressParts[0].trim().split(' ');
        let houseNumberRegex = /^[\d\-A-Za-z]+/; // Matches digits, hyphens, and letters at the start
        let houseNumberMatch = houseStreetParts[0].match(houseNumberRegex);
        if (houseNumberMatch) {
            House = houseNumberMatch[0];
            Street = houseStreetParts.slice(1).join(' ');
        }
        // Get the city
        City = addressParts[1]?.trim() || "";
        // Get the ZIP code
        let zipRegex = /\b\d{5}\b/; // Matches exactly 5 digits
        let zipMatch = Address.match(zipRegex);
        if (zipMatch) {
            Zip = zipMatch[0];
        }
    } catch (error) {
        console.error(`Error while geocoding: ${error.message}`);
    }
}

return { Firstname, Lastname, FullName, Address, applicationLink, House, Street, City, lead, Zip, Landmark, CB, Panel, NYC_LI, Size, Stories, Borough, contractfileid, Contract_Price, Block, Lot, Bin, Flood_Zone, Oath_Ecb_Violation };
    } catch (error) {
        console.error(`Error in main: ${error}`);
        if (error.response) {
            console.error('Error status:', error.response.status);
            console.error('Error data:', error.response.data);
        }
    }
}


(async () => {
        const browser = await puppeteer.launch({
        headless: true, // This will run the browser visibly
        slowMo: 1 // This will slow down Puppeteer operations by 100 milliseconds
    });
const accessToken = await caspioAuthorization();
    const Cust_ID = process.argv[2];
    //const Cust_ID = '4964';  
    let caspioData = await caspiomain(Cust_ID, accessToken);
    let applicationLink = caspioData.applicationLink;
	console.log(applicationLink);
    let name = caspioData.FullName;

    const page = await browser.newPage();
    const timeout = 40000;
    page.setDefaultTimeout(timeout);

    // Set viewport for the screenshot
    await page.setViewport({
        width: 1303,
        height: 931
    });

// Navigate to the application link
    await page.goto(applicationLink, { waitUntil: 'networkidle2' });
	await sleep(5000);
   const selector = '#ctl00_PlaceHolderMain_InspectionList_inspectionUpdatePanel';
    //await page.waitForSelector(selector);
console.log("Waiting for the '#ctl00_PlaceHolderMain_InspectionList_inspectionUpdatePanel' to be visible...");
    await page.waitForSelector('#ctl00_PlaceHolderMain_InspectionList_inspectionUpdatePanel');

    console.log("Printing text of each row in the table...");
    const rowTexts = await page.evaluate(() => {
        const table = document.querySelector('#ctl00_PlaceHolderMain_InspectionList_gvListCompleted');
        return Array.from(table.querySelectorAll('.InspectionListRow')).map(row => row.innerText);
    });

    console.log(rowTexts);
console.log("Attempting to extract details from 'Pass-Final' row...");
const passFinalDetails = await page.evaluate(() => {
    const rows = Array.from(document.querySelectorAll('#ctl00_PlaceHolderMain_InspectionList_gvListCompleted .InspectionListRow'));
    for (let row of rows) {
        if (row.textContent.includes('Pass-Final')) {
            const resultText = row.textContent;
            // Adjusted regular expression to capture only the date
            const dateMatch = resultText.match(/Resulted\s+on\s+(\d{2}\/\d{2}\/\d{4})/);
            const dateText = dateMatch ? dateMatch[1] : null;
            const viewDetailsLink = row.querySelector('a')?.getAttribute('onclick');
            if (dateText && viewDetailsLink) {
                return { dateText, viewDetailsOnClick: viewDetailsLink };
            }
        }
    }
    return null;
});
if (passFinalDetails) {
    console.log("Pass Final Date found:", passFinalDetails.dateText);
	
// Check if viewDetailsOnClick contains a valid function call
    // Click on the 'View Details' link of the 'Pass-Final' row
    await page.evaluate(() => {
        const passFinalRow = Array.from(document.querySelectorAll('#ctl00_PlaceHolderMain_InspectionList_gvListCompleted .InspectionListRow'))
            .find(row => row.textContent.includes('Pass-Final'));
        const viewDetailsLink = passFinalRow ? passFinalRow.querySelector('a') : null;
        if (viewDetailsLink) {
            viewDetailsLink.click();
        }
    });
    await sleep(5000); // Wait for the new page to load after clicking

        // Take a screenshot
        const screenshotBuffer = await page.screenshot();
        const pdfDoc = await PDFDocument.create();
        const pngImage = await pdfDoc.embedPng(screenshotBuffer);
        const { width, height } = pngImage;
        const pdfPage = pdfDoc.addPage([width, height]);
        pdfPage.drawImage(pngImage, { x: 0, y: 0, width, height });
        const pdfBytes = await pdfDoc.save();

        // Saving the PDF
        const directoryPath = '/tmp';
        if (!fs.existsSync(directoryPath)) fs.mkdirSync(directoryPath);
     const pdfPath = path.join(directoryPath, `${name.replace(/\s/g, '_')}_electric_inspection.pdf`);
    fs.writeFileSync(pdfPath, pdfBytes);
    console.log("PDF saved to:", pdfPath);

    // Creating a JSON object
    const outputJson = {
        passdate: passFinalDetails.dateText,
        passfilepath: pdfPath
    };

    // JSON output directory
    const jsonOutputDir = '/tmp';
    if (!fs.existsSync(jsonOutputDir)) {
        fs.mkdirSync(jsonOutputDir);
    }

    // JSON file path
    const jsonFilePath = path.join(jsonOutputDir, `${Cust_ID}_electric.json`);

    // Saving the JSON string to a file
    fs.writeFileSync(jsonFilePath, JSON.stringify(outputJson, null, 4));

    console.log("JSON saved to:", jsonFilePath);

    // Print the JSON object as the last line
    console.log(JSON.stringify(outputJson));
} else {
    console.log("No 'Pass-Final' row found in the specified panel or missing details.");
}
    // Uncomment the following line if you want to close the browser at the end of the script
     await browser.close();
})().catch(err => {
    console.error(err);
    process.exit(1);
});