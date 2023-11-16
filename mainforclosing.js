const { spawn } = require('child_process');
const { google } = require('googleapis');
//const keys = require('./servicekey.json');
const caspioTokenUrl = 'https://c1abd578.caspio.com/oauth/token';
const clientID = process.env.clientID;
const clientSecret = process.env.clientSecret;
const GoggleApi = process.env.GoggleApi;
const keys =JSON.parse(process.env.SERVICE_KEY);
const axios = require('axios');
const qs = require('querystring');


function sleep(ms) {
    return new Promise(resolve => setTimeout(resolve, ms));
}

async function fetchCaspioData(accessToken, url, updateData = null, method = 'GET') {
	
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
        // Define URLs
        const masterTableUrl = `https://c1abd578.caspio.com/rest/v2/tables/MasterTable/records?q.select=Name%2C%20Address1%2C%20Phone1%2C%20email&q.where=CID=${CIDvalue}`;
        const solarProcessUrl = `https://c1abd578.caspio.com/rest/v2/tables/SolarProcess/records?q.select=Contract_Price%2CPanel%2CSystem_Size_Sold%2CBlock%2CLot%2CBin%2CFlood_Zone%2CNYC_LI%2COath_Ecb_Violation%2CFlood_Zone%2CLandmark%2CBorough%2CStories%2CCB&q.where=CustID=${CIDvalue}`;
        const signedcontractsurl = `https://c1abd578.caspio.com/rest/v2/tables/Signed_contracts/records?q.select=DownloadLink&q.where=CustID=${CIDvalue}`;

        // Fetch Data
        const masterDataResponse = await fetchCaspioData(accessToken, masterTableUrl);
        const solarProcessResponse = await fetchCaspioData(accessToken, solarProcessUrl);
        const signedcontractResponse = await fetchCaspioData(accessToken, signedcontractsurl);
console.log(masterDataResponse);
console.log(solarProcessResponse);
console.log(signedcontractResponse);
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
        const combinedData = { ...masterData, ...solarData, ...{ DownloadLink: contractlink, ContractFileID: contractfileid }};
console.log(combinedData);
        // Extract and process further data
        // Extract and process further data
let FullName = masterData?.Name || "Unknown"; // Default to "Unknown" if Name is not available
let Firstname = "";
let Lastname = "";
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

if (originalAddress) {
    try {
        const response = await axios.get(`https://maps.googleapis.com/maps/api/geocode/json?address=${encodeURIComponent(originalAddress)}&key=${GoggleApi}`);
        Address = response.data.results[0]?.formatted_address || originalAddress; // Fallback to originalAddress if no result

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

return { Firstname, Lastname, FullName, Address, House, Street, City, lead, Zip, Landmark, CB, Panel, NYC_LI, Size, Stories, Borough, contractfileid, Contract_Price, Block, Lot, Bin, Flood_Zone, Oath_Ecb_Violation };
    } catch (error) {
        console.error(`Error in main: ${error}`);
        if (error.response) {
            console.error('Error status:', error.response.status);
            console.error('Error data:', error.response.data);
        }
    }
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
    const range = 'Sheet1!A2:A60';
    
    let response = await sheets.spreadsheets.values.get({
        spreadsheetId,
        range,
    });

    let rows = response.data.values;
    let custIDs = [];
    if (rows.length) {
        rows.forEach(row => {
            let CustID = row[0];
            if (CustID) {
                custIDs.push(CustID);
            }
        });
        return custIDs;
        console.log(custIDs);
    } else {
        console.log('No data found.');
        return null;
    }
}
function runScript(scriptPath, ...args) {
    return new Promise((resolve, reject) => {
        let dataString = '';
        const process = spawn('node', [scriptPath, ...args]);

        process.stdout.on('data', (data) => {
            dataString += data.toString();
        });

        process.stderr.on('data', (data) => {
            console.error(`stderr: ${data}`);
        });

        process.on('close', (code) => {
            if (code !== 0) {
                reject(new Error(`child process exited with code ${code}`));
                return;
            }
            resolve(dataString);
        });
    });
}

async function main(Cust_ID, Lastname, FullName, house_street, electrical_ID) {
    try {
        // Starting ptoletter.js
        console.log(`Running ptoletter for cust_id = ${Cust_ID}`);
        let ptoletterOutput;
        try {
            ptoletterOutput = await runScript('./ptoletter.js', Cust_ID, Lastname, FullName, house_street);
            sendProgressUpdate(20, `PTO letter found and saved for ${FullName}`);
        } catch (error) {
            console.error("Error while running ptoletter.js:", error.message);
            sendProgressUpdate(0, `Error: Could not run ptoletter.js for ${FullName}`);
            return;
        }
        let ptoletterData = extractOutputData(ptoletterOutput);
        await sleep(5000);

        // Running electricalpermit.js
        if (electrical_ID && electrical_ID !== "") {
            await runScript('./electricalpermit.js', Cust_ID);
            sendProgressUpdate(40, `Electrical permit number found and updated in Caspio for ${FullName}`);
        } else {
            //console.log(`electricalpermit.js already ran for this customer before. Skipping...`);
            sendProgressUpdate(40, `Electrical permit already updated for ${FullName}`);
        }

        //console.log(`Running electricaltest for cust_id = ${Cust_ID}`);
        const passinspectionOutput = await runScript('./electricaltest.js', Cust_ID, Lastname);
        let passinspectionData = extractOutputData(passinspectionOutput);
        sendProgressUpdate(65, `Electrical inspection result saved for ${FullName}`);
        await sleep(2000);

        //console.log(`Running nyccloseout for cust_id = ${Cust_ID}`);
        await runScript('./nyccloseout.js', Cust_ID, ptoletterData.ptodate, ptoletterData.filepath, passinspectionData.passdate, passinspectionData.passfilepath);
        sendProgressUpdate(100, `Closeout forms filled out and saved in Google Drive for ${FullName}`);
        await sleep(10000);
    } catch (error) {
        console.error(`Error in main: ${error.message}`);
        console.log("Error: Could not complete the process");
    }
}

// This function will be replaced with SSE or another method to communicate with the client
function sendProgressUpdate(percentage, message) {
    // Placeholder for sending progress updates to the client
    console.log(`Progress: ${percentage}% - ${message}`);
}


function extractOutputData(output) {
    try {
        // Split output by new lines and find the line that is a valid JSON
        const outputLines = output.trim().split('\n');
        for (const line of outputLines) {
            if (line.startsWith('{') && line.endsWith('}')) {
                return JSON.parse(line);
            }
        }
        throw new Error('No JSON data found in output.');
    } catch (err) {
        throw new Error(`Error parsing JSON data: ${err.message}`);
    }
}


async function processAllCustomers() {
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
    await client.authorize();
    //console.log("Successfully connected to Caspio!");
    const accessToken = await caspioAuthorization();

    const sheetData = await fetchSpreadsheetData(client);
    for (let i = 0; i < sheetData.length; i++) {
        let Cust_ID = process.argv[2];
        const electricpermitexistsurl = `https://c1abd578.caspio.com/rest/v2/tables/electricalpermit/records?q.select=passdate&q.where=CID=${Cust_ID}`;
        let electrical_IDexist = await fetchCaspioData(accessToken, electricpermitexistsurl);
        //console.log(electrical_IDexist);
        // Check if Result[0] exists and if it does, get the passdate. Otherwise, set it to null.
        let electrical_ID = electrical_IDexist.Result && electrical_IDexist.Result[0] ? electrical_IDexist.Result[0].passdate : null;
        //console.log(electrical_ID);
        let caspioData = await caspiomain(Cust_ID, accessToken);
        console.log(caspioData);
        const Lastname = caspioData.Lastname;
        let FullName = caspioData.FullName;
        let house_street = `${caspioData.House} ${caspioData.Street}`;
        console.log(house_street);
        sendProgressUpdate(10, `Running analysis for ${FullName}`);
        await main(Cust_ID, Lastname, FullName, house_street, electrical_ID);
    }

}

// Start the process
processAllCustomers().catch(error => {
    console.error('An error occurred in processAllCustomers:', error);
});