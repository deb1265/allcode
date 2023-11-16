const { spawn } = require('child_process');
const { google } = require('googleapis');
const keys =JSON.parse(process.env.SERVICE_KEY);

function sleep(ms) {
    return new Promise(resolve => setTimeout(resolve, ms));
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
    } else {
        console.log('No data found.');
        return null;
    }
}
function runScript(scriptPath, Cust_ID) {
    return new Promise((resolve, reject) => {
        const process = spawn('node', [scriptPath, Cust_ID]);

        process.stdout.on('data', (data) => {
            console.log(`stdout: ${data}`);
        });

        process.stderr.on('data', (data) => {
            console.error(`stderr: ${data}`);
        });

        process.on('close', (code) => {
            if (code !== 0) {
                reject(new Error(`child process exited with code ${code}`));
                return;
            }
            resolve(code);
        });
    });
}
async function main(Cust_ID) {
    try {
        // Logging before running billdownload.js
        console.log(`Running billdownload for cust_id = ${Cust_ID}`);

        // Run billdownload.js
       //await runScript('./docaibilldownload.js', Cust_ID);

        // Wait for 30 seconds
        //await sleep(10000);

        //Logging before running interconnect.js
        console.log(`Running interconnect for cust_id = ${Cust_ID}`);

        //Run interconnect.js
       await runScript('./interconnect.js', Cust_ID);

        // Wait for 5 seconds before processing the next Cust_ID
        await sleep(10000);
    } catch (error) {
        console.error(`Error: ${error.message}`);
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
    console.log("Successfully connected to Caspio!");
    const sheetData = await fetchSpreadsheetData(client);
	const Cust_ID = process.argv[2];
	//const Cust_ID = 7124;
	console.log(`Received Cust_ID: ${Cust_ID}`);
	await main(Cust_ID);
/*
    for (let i = 0; i < sheetData.length; i++) {
        let Cust_ID = sheetData[i];
        await main(Cust_ID);  // Ensure that for each Cust_ID, the two scripts run one after the other.
    }
	*/
}

processAllCustomers();