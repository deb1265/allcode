const fs = require('fs');
const puppeteer = require('puppeteer');
const { google } = require('googleapis');
const path = require('path');
const { PDFDocument } = require('pdf-lib');
//const keys = require('./servicekey.json');
const axios = require('axios');
const cheerio = require('cheerio');
const moment = require('moment');
const qs = require('querystring'); 
const util = require('util');
const exec = util.promisify(require('child_process').exec);
const caspioTokenUrl = 'https://c1abd578.caspio.com/oauth/token';
const clientID = process.env.clientID;
const clientSecret = process.env.clientSecret;
const GoggleApi = process.env.GoggleApi;
const keys =JSON.parse(process.env.SERVICE_KEY);
const pdfPoppler = require('pdf-poppler');
const nycsformsfileId = '1A7y68lk0lwDzJRJPKMO26KzrUjV46UUT';
const sharp = require('sharp');
const delay = (ms) => new Promise((resolve) => setTimeout(resolve, ms));
const nycblanksdownloadedpath = path.resolve('/tmp', 'file.pdf'); 
const contractPath = path.resolve('/tmp', 'contract.pdf'); 
//const Cust_ID = 7364;
const Cust_ID = process.argv[2];
const PDFDocument1 = require('pdfkit');
const API_KEY = 'WrEuodSc5lzfQyfHfCCRmPH790rIz8ta7pkrtFnX';
const BASE_URL = 'https://developer.nrel.gov/api/pvwatts/v6.json';
//const { Cust_ID } = require('./main');  // Import Cust_ID from main.js
async function deleteFileWithRetry(filePath, retryCount = 5) {
    for (let attempt = 0; attempt < retryCount; attempt++) {
        try {
            fs.unlinkSync(filePath);
            fs.unlinkSync('filed.pdf');
            //fs.unlinkSync('contract-17.png');
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
async function generatePvWattsReport(address, systemCapacity, losses, tilt, azimuth, index) {


    // Fetch data function
    async function fetchPvWattsData() {
        const params = {
            'api_key': API_KEY,
            'address': address,
            'system_capacity': systemCapacity,
            'azimuth': azimuth,
            'tilt': tilt,
            'array_type': 1,
            'module_type': 0,
            'losses': losses,
            'format': 'json'
        };

        try {
            const response = await axios.get(BASE_URL, { params });
            return response.data.outputs;
        } catch (error) {
            console.error('Error fetching data:', error);
        }
    }

    function generatePDF(values, filename) {
        const doc = new PDFDocument1();
        doc.pipe(fs.createWriteStream(filename));

        // Title
        doc.fontSize(18).fillColor('red').text('PV Watts', { align: 'center' }).fillColor('black');
        doc.fontSize(12).text('Address: 1265 sunrise highway, bayshore, ny ,11706');
        //doc.fontSize(12).text(`Array tilt: ${tilt}`);
        //doc.fontSize(12).text(`Array Azimuth: ${azimuth}`);

        // Table headers
        const columns = ['Month', 'AC (Kwh)', 'POA Monthly', 'Solrad Monthly', 'DC (Kwh)'];
        const colWidths = [100, 100, 100, 100, 100];
        const tableX = 80; // Starting x position for table
        const tableY = doc.y + 10; // Starting y position for table

        // Draw the header
        doc.font('Helvetica-Bold').fontSize(10);
        for (let i = 0; i < columns.length; i++) {
            doc.text(columns[i], tableX + colWidths.slice(0, i).reduce((a, b) => a + b, 0), tableY);
        }

        // Monthly data
        doc.font('Helvetica').fontSize(10);

        let yPosition = tableY + 20; // y position for drawing horizontal lines

        // Data
        // Data
        const months = ['January', 'February', 'March', 'April', 'May', 'June', 'July', 'August', 'September', 'October', 'November', 'December', 'Total'];


        months.forEach((month, index) => {
            doc.text(month, tableX, yPosition);
            if (values[index]) {
                values[index].forEach((value, colIndex) => {
                    let textValue = value.toFixed(2);
                    doc.text(textValue, tableX + colWidths.slice(0, colIndex + 1).reduce((a, b) => a + b, 0), yPosition);
                });
            }
            yPosition += 20;
        });



        // Draw vertical lines for columns
        for (let i = 0; i <= colWidths.length; i++) {
            doc.moveTo(tableX + colWidths.slice(0, i).reduce((a, b) => a + b, 0), tableY)
                .lineTo(tableX + colWidths.slice(0, i).reduce((a, b) => a + b, 0), yPosition)
                .stroke();
        }

        for (let i = 0; i <= months.length + 1; i++) { // +1 to draw boundary for last row
            doc.moveTo(tableX, tableY + 20 * i)
                .lineTo(tableX + colWidths.reduce((a, b) => a + b, 0), tableY + 20 * i)
                .stroke();
        }
        const pvWattsLogoPath = './pvwatts/pvwatts logo.JPG'; // Path to the PV Watts logo
        const nrelLogoPath = './pvwatts/nrel logo.png'; // Path to the NREL logo

        yPosition += 128; // Add desired inches to push the logos down. Adjust this value as needed.

        // Add PV Watts Logo
        doc.image(pvWattsLogoPath, tableX, yPosition, { width: 100 });

        // Add NREL Logo (assuming you want it next to PV Watts logo)
        doc.image(nrelLogoPath, tableX + 120, yPosition, { width: 100 });

        // Adding Footer Note
        yPosition += 120; // Adjust as needed
        doc.fontSize(10).text('Powered by pvwatts V8 API', tableX, yPosition);

        doc.end();
        return filename; // Return the file path
    }

    async function main() {
        const data = await fetchPvWattsData();
        const values = data.ac_monthly.map((value, index) => [
            value,
            data.poa_monthly[index],
            data.solrad_monthly[index],
            data.dc_monthly[index]
        ]);

        // Adding total row
        values.push([
            data.ac_monthly.reduce((sum, val) => sum + val, 0),
            data.poa_monthly.reduce((sum, val) => sum + val, 0),
            data.solrad_monthly.reduce((sum, val) => sum + val, 0),
            data.dc_monthly.reduce((sum, val) => sum + val, 0)
        ]);

        const filename = `PVWatts_As_built.pdf`;
        return generatePDF(values, filename); // Return the file path
    }
    
    // Call main function here
    return await main();
}
async function uploadpvwattsFile(client, filePath, folderId) {
    try {
        const drive = google.drive({ version: 'v3', auth: client });
        const ext = path.extname(filePath);

        if (ext === '.pdf') {
            const uploadPromise = drive.files.create({
                resource: {
                    'name': path.basename(filePath),
                    'parents': [folderId],
                },
                media: {
                    mimeType: 'application/pdf',
                    body: fs.createReadStream(filePath),
                },
                fields: 'id',
            });

            const response = await uploadPromise;
            console.log('File Id: ', response.data.id);

            fs.unlinkSync(filePath);
        } else {
            console.log('Unsupported file type');
        }

        console.log('File exists after deletion: ', fs.existsSync(filePath));
    } catch (error) {
        console.error('Error uploading file: ', error);
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
/*
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

*/
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
async function createFolder(name, parent, drive) {
    var fileMetadata = {
        'name': name,
        'mimeType': 'application/vnd.google-apps.folder',
        'parents': [parent]
    };
    var response = await drive.files.create({
        resource: fileMetadata,
        fields: 'id'
    });
    console.log('Folder Id: ', response.data.id);
    return response.data.id;
}
async function uploadFile(client, filePath, folderId,NYC_LI,Totalstring,Name,METER) {
    try {
        const drive = google.drive({ version: 'v3', auth: client });
        const ext = path.extname(filePath);
        let signedFormsFolderId = await createFolder(`${Name}-PTO-${METER}`, folderId, drive); 
        //console.log('signedFormsFolderId: ', signedFormsFolderId);
        // If the file is a PDF, split it into parts and upload each part
        if (ext === '.pdf') {
            const pdfBytes = fs.readFileSync(filePath);
            const pdfDoc = await PDFDocument.load(pdfBytes);
            let parts;
            if (NYC_LI !== 'NYC') {
                parts = {
                    'Self cert': [0, 0],
                    'Pseg_AppendixB': [1, 3],
                    'Pseg_AppendixL': [4, 4]
                };
				if (Totalstring == 1) {
                    parts['Pseg_ELD'] = [8, 8];
                } else if (Totalstring == 2) {
                    parts['Pseg_ELD'] = [9, 9];
                } else if (Totalstring == 3) {
                    parts['Pseg_ELD'] = [10, 10];
                } else if (Totalstring == 4) {
                    parts['Pseg_ELD'] = [6, 6];
                }
            } else {
                parts = {
                    'contractor_certification': [7, 7]

                };
            }

            for (const [name, range] of Object.entries(parts)) {
                const newPdfDoc = await PDFDocument.create();
                const [start, end] = Array.isArray(range) ? range : [range, range];

                for (let i = start; i <= end; i++) {
                    const [copiedPage] = await newPdfDoc.copyPages(pdfDoc, [i]);
                    newPdfDoc.addPage(copiedPage);
                }

                const newFilePath = path.join(path.dirname(filePath), `${name}.pdf`);
                fs.writeFileSync(newFilePath, await newPdfDoc.save());

                const uploadPromise = drive.files.create({
                    resource: {
                        'name': path.basename(newFilePath),
                        'parents': [signedFormsFolderId],
                    },
                    media: {
                        mimeType: 'application/pdf',
                        body: fs.createReadStream(newFilePath),
                    },
                    fields: 'id',
                });

                const response = await uploadPromise;
                console.log('File Id: ', response.data.id);
			//return signedFormsFolderId;
			//console.log('signedFormsFolderId: ', signedFormsFolderId);
            fs.unlinkSync(newFilePath);
            }
            fs.unlinkSync(filePath);
        } else if (ext === '.png') {
            const uploadPromise = drive.files.create({
                resource: {
                    'name': path.basename(filePath),
                    'parents': [folderId],
                },
                media: {
                    mimeType: 'image/png',
                    body: fs.createReadStream(filePath),
                },
                fields: 'id',
            });

            const response = await uploadPromise;
            console.log('File Id: ', response.data.id);

            fs.unlinkSync(filePath);
            
			//console.log('signedFormsFolderId: ', signedFormsFolderId);
        }
		return (signedFormsFolderId);
        console.log('File exists after deletion: ', fs.existsSync(filePath));
    } catch (error) {
        console.error('Error uploading file: ', error);
    }
    
}

async function caspiomain(CIDvalue, accessToken) {
    try {
        const masterTableUrl = `https://c1abd578.caspio.com/rest/v2/tables/MasterTable/records?q.select=Name%2C%20Address1%2C%20Phone1%2C%20email&q.where=CID=${CIDvalue}`;
        const solarProcessUrl = `https://c1abd578.caspio.com/rest/v2/tables/SolarProcess/records?q.select=PAM%2CPAM_approvedate%2CInterconnectionName2%2CInterconnectionName3%2CRevised_numberofpanels%2CRevised_system_size%2CRevised_panel%2CInterconnectionMeter1%2CBin%2CInterconnectionMeter2%2CNYC_LI%2CInterconnectionMeter3%2CInterconnectionAcc1%2CInterconnectionAcc2%2CInterconnectionAcc3%2CInterconnectionName1%2CCB&q.where=CustID=${CIDvalue}`;
        const signedcontractsurl = `https://c1abd578.caspio.com/rest/v2/tables/Signed_contracts/records?q.select=DownloadLink&q.where=CustID=${CIDvalue}`;
        const masterDataResponse = await fetchCaspioData(accessToken, masterTableUrl);
        const solarProcessResponse = await fetchCaspioData(accessToken, solarProcessUrl);
        const signedcontractResponse = await fetchCaspioData(accessToken, signedcontractsurl);
        const signedData = signedcontractResponse.Result[signedcontractResponse.Result.length - 1];
        const masterData = masterDataResponse.Result[0];
        const solarData = solarProcessResponse.Result[0];
        const combinedData = { ...masterData, ...solarData, ...signedData };
        let contractlink = combinedData.DownloadLink;
        let contractfileid = contractlink.split('id=')[1].split('&export=download')[0];
        const originalAddress = combinedData.Address1;
        const response = await axios.get(`https://maps.googleapis.com/maps/api/geocode/json?address=${encodeURIComponent(originalAddress)}&key=${GoggleApi}`);
        const Address = response.data.results[0].formatted_address;
        const PAM = combinedData.PAM;
        const Panel = combinedData.Revised_panel;
        const Size = combinedData.Revised_system_size;
        const Number_Of_Panels = combinedData.Revised_numberofpanels;
        const Interconnectionname1 = combinedData.InterconnectionName1;
        const Interconnectionname2 = combinedData.InterconnectionName2;
        const Interconnectionname3 = combinedData.InterconnectionName3;
        const InterconnectionMeter1 = combinedData.InterconnectionMeter1;
        const InterconnectionMeter2 = combinedData.InterconnectionMeter2;
        const InterconnectionMeter3 = combinedData.InterconnectionMeter3;
        const InterconnectionAcc1 = combinedData.InterconnectionAcc1;
        const InterconnectionAcc2 = combinedData.InterconnectionAcc2;
        const InterconnectionAcc3 = combinedData.InterconnectionAcc3;
        const NYC_LI = combinedData.NYC_LI;
		const PAM_approvedate = combinedData.PAM_approvedate;
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
        return { Firstname, Lastname, FullName, Address, House, Street, City, Zip, Panel, NYC_LI, Size, contractfileid, PAM,PAM_approvedate, Number_Of_Panels, InterconnectionMeter1, InterconnectionMeter2, InterconnectionMeter3, InterconnectionAcc1, InterconnectionAcc2, InterconnectionAcc3, Interconnectionname1, Interconnectionname2, Interconnectionname3 };
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
    await downloadFile(drive, nycsformsfileId, nycblanksdownloadedpath);
    const accessToken = await caspioAuthorization();
    console.log("Successfully connected to Google Sheets!");
    console.log("Successfully connected to Caspio!");
    //const sheetData = await fetchSpreadsheetData(client);
    let caspioData = await caspiomain(Cust_ID, accessToken);
    console.log("Data: ", caspioData);
    let Numberofpanel = Number(caspioData.Number_Of_Panels);
    let Numberofpanels1 = '';
    let Numberofpanels2 = '';
    let Numberofpanels3 = '';
    let InterconnectionMeter1 = caspioData.InterconnectionMeter1;
    let InterconnectionMeter2 = caspioData.InterconnectionMeter2;
    let InterconnectionMeter3 = caspioData.InterconnectionMeter3;
    let Interconnectionname1 = caspioData.Interconnectionname1;
    let Interconnectionname2 = caspioData.Interconnectionname2;
    let Interconnectionname3 = caspioData.Interconnectionname3;
    let InterconnectionAcc1 = caspioData.InterconnectionAcc1;
    let InterconnectionAcc2 = caspioData.InterconnectionAcc2;
    let InterconnectionAcc3 = caspioData.InterconnectionAcc3;
	let PAM = caspioData.PAM;
	let PAM_approvedate = caspioData.PAM_approvedate;
    let validMeters = [InterconnectionAcc1, InterconnectionAcc2, InterconnectionAcc3].filter(val => val !== 0 && val !== "");
    // Check conditions
    if (InterconnectionMeter3 === 0 || InterconnectionMeter3 === "") {
        validMeters = [InterconnectionMeter1, InterconnectionMeter2];
    }

    if (InterconnectionAcc2 === 0 || InterconnectionAcc2 === "" && (InterconnectionAcc3 === 0 || InterconnectionAcc3 === "")) {
        validMeters = [InterconnectionMeter1];
    }
    let interconnectionAccounts = [InterconnectionAcc1, InterconnectionAcc2, InterconnectionAcc3];
    let interconnectionMeters = [InterconnectionMeter1, InterconnectionMeter2, InterconnectionMeter3];

    if (validMeters.length === 3) {
        Numberofpanels1 = Math.round(Numberofpanel / 3);
        Numberofpanels2 = Math.round(Numberofpanel / 3);
        Numberofpanels3 = Numberofpanel - Numberofpanels1 - Numberofpanels2;
    } else if (validMeters.length === 2) {
        Numberofpanels1 = Math.round(Numberofpanel / 2);
        Numberofpanels2 = Numberofpanel - Numberofpanels1;
    } else if (validMeters.length === 1) {
        Numberofpanels1 = Numberofpanel;
    }
    let Numberofpanellist = [Numberofpanels1, Numberofpanels2, Numberofpanels3];
    let namelist = [Interconnectionname1, Interconnectionname2, Interconnectionname3];
    console.log("Panel array: ", Numberofpanellist);
    console.log("Name array: ", namelist);
    for (let i = 0; i < interconnectionAccounts.length; i++) {
        if (interconnectionAccounts[i]) {                     //interconnectionMeters[i] && 
            let Account = interconnectionAccounts[i];
            let Meter = interconnectionMeters[i];
            let Numberofpanels = Numberofpanellist[i];
            let Name;
            if (namelist[i] === "" || namelist[i] === null) {
                Name = caspioData.FullName;
            } else {
                Name = namelist[i];
            }
            console.log("Panels Meter : ", Numberofpanels);
            // if interconnectionMeters.length=2 then 
            const filePath = './filled.pdf';
            try {
                if (fs.existsSync(filePath)) {
                    await deleteFileWithRetry(filePath);
                }
                //let caspioData = await caspiomain(Cust_ID, accessToken);
                //console.log("Data: ", caspioData);
                const blankpdfpath = await PDFDocument.load(fs.readFileSync(nycblanksdownloadedpath));
                const pdfDoc = blankpdfpath;
                const form = pdfDoc.getForm();
                //const totalCost = Math.round(parseFloat(caspioData.Contract_Price));
                let panel = caspioData.Panel;
                let Eachpanelkw;
                if (panel.includes("400")) {
                    Eachpanelkw = "0.400";
                } else if (panel.includes("365")) {
                    Eachpanelkw = "0.365";
                } else if (panel.includes("405")) {
                    Eachpanelkw = "0.405";
                } else if (panel.includes("370")) {
                    Eachpanelkw = "0.370";
                } else {
                    Eachpanelkw = "0.400"; // default value if none of the other conditions are met
                }
                let contractfileid = caspioData.contractfileid;
                await downloadFile(drive, contractfileid, contractPath);
                console.log("Successfully connected to Caspio!");
                const dateToday = moment().format('MM.DD.YYYY');
				const dateAWeekAgo = moment().subtract(7, 'days').format('MM.DD.YYYY');
                const NYC_LI = caspioData.NYC_LI;
                //const dateofreport = moment().add(-15, 'days').format('MM.DD.YYYY');
                //const dateofreport = moment().subtract(15, 'days').format('MM/DD/YYYY');
                //let systemsize = caspioData.Size.toString();
                // Function to round a number to two decimal places
                function twoDecimalUpto(num) {
                    return parseFloat(num.toFixed(2));
                }
                let system = twoDecimalUpto(Numberofpanels * parseFloat(Eachpanelkw));
                console.log("calculated system size:", system);
                let systemsize = system.toString();
                console.log("calculated system size in string:", systemsize);
                let Listofbreakers = '';
                let sumofbreakers = '';
                let String1 = '';
                let String2 = '';
                let String3 = '';
                let String4 = '';
                let Amp4 = '';
                let Amp3 = '';
                let Amp2 = '';
                let Amp1 = '';
                let Totalstring = '';
                console.log("Initial Number of panels:", Numberofpanels);
                if (Numberofpanels <= 13) {
                    String1 = Numberofpanels;
                    let Amperate1 = Numberofpanels * 1.25 * 1.21;
                    if (Amperate1 <= 10) Amp1 = '10A';
                    else if (Amperate1 > 10 && Amperate1 <= 15) Amp1 = '15A';
                    else if (Amperate1 > 15 && Amperate1 <= 20) Amp1 = '20A';
                    console.log("String1:", String1, "Amperate1:", Amperate1, "Amp1:", Amp1);
                } else if (Numberofpanels > 13 && Numberofpanels <= 26) {
                    String1 = Numberofpanels <= 20 ? 9 : 13;
                    let Amperate1 = String1 * 1.25 * 1.21;
                    if (Amperate1 <= 10) Amp1 = '10A';
                    else if (Amperate1 > 10 && Amperate1 <= 15) Amp1 = '15A';
                    else if (Amperate1 > 15 && Amperate1 <= 20) Amp1 = '20A';
                    String2 = Numberofpanels - String1;
                    let Amperate2 = String2 * 1.25 * 1.21;
                    Amp2 = Amperate2 <= 10 ? '10A' : (Amperate2 <= 15 ? '15A' : '20A');
                    console.log("String1:", String1, "String2:", String2, "Amperate2:", Amperate2, "Amp1:", Amp1, "Amp2:", Amp2);
                } else if (Numberofpanels > 26 && Numberofpanels <= 39) {
                    String1 = 13;
                    Amp1 = '20A';
                    String2 = Numberofpanels <= 36 ? 9 : 13;
                    let Amperate2 = String2 * 1.25 * 1.21;
                    Amp2 = Amperate2 <= 10 ? '10A' : (Amperate2 <= 15 ? '15A' : '20A');
                    String3 = Numberofpanels - String1 - String2;
                    let Amperate3 = String3 * 1.25 * 1.21;
                    Amp3 = Amperate3 <= 10 ? '10A' : (Amperate3 <= 15 ? '15A' : '20A');
                    console.log("String1:", String1, "String2:", String2, "Amperate2:", Amperate2, "Amp1:", Amp1, "Amp2:", Amp2);
                } else if (Numberofpanels > 39 && Numberofpanels <= 52) {
                    String1 = 13;
                    Amp1 = '20A';
                    String2 = 13;
                    Amp2 = '20A';
                    String3 = Numberofpanels < 46 ? 7 : 13;
                    let Amperate3 = String3 * 1.25 * 1.21;
                    Amp3 = Amperate3 <= 10 ? '10A' : (Amperate3 <= 15 ? '15A' : '20A');
                    String4 = Numberofpanels - String1 - String2 - String3;
                    let Amperate4 = String4 * 1.25 * 1.21;
                    Amp4 = Amperate4 <= 10 ? '10A' : (Amperate4 <= 15 ? '15A' : '20A');
                    console.log("String1:", String1, "String2:", String2, "Amp1:", Amp1, "Amp2:", Amp2);
                }
                sumofbreakers = [Amp1, Amp2, Amp3, Amp4].reduce((acc, val) => acc + (Number(val.replace('A', '')) || 0), 0) + 'A';
                console.log("Sum of breakers:", sumofbreakers);
                Listofbreakers = [Amp1, Amp2, Amp3, Amp4].filter(val => val).join(' & ');
                console.log("List of breakers:", Listofbreakers);
                Totalstring = [Amp1, Amp2, Amp3, Amp4].filter(val => val).length;
                console.log("Totalstring (number of non-empty Amp values):", Totalstring);
                // Define these variables somewhere in your code
                //await deleteFileWithRetry(filePath);
                //let PAMapprovaldate = '10/12/2023'; // need to define
                //let PAMnumber = '232-2234-2'; // need to define
                let Totalinvoutput = twoDecimalUpto(0.29 * Numberofpanels);
                const fields = {
                    Allbreakers: String(Listofbreakers),
                    Totalstring: String(Totalstring),
                    String1: String(String1),
                    String2: String(String2),
                    String3: String(String3),
                    String4: String(String4),
                    Amp4: String(Amp4),
                    Amp3: String(Amp3),
                    Amp2: String(Amp2),
                    Amp1: String(Amp1),
                    AmpC: String(sumofbreakers),
                    Name: Name,
                    Address: caspioData.Address,
                    //Phone: caspioData.Phone,
                    //Email: caspioData.Email,
                    Account: Account,
                    Meter: Meter,
                    Signing_date: dateAWeekAgo,
                    testingdate: dateAWeekAgo,
                    //City: caspioData.City,
                    //Kw: systemsize,
                    ACoutput: Totalinvoutput,
                    Paneloutput: systemsize,
                    //Panel: panel,
                    Panelbrand: panel,
                    Panelkw: Eachpanelkw,
                    Panelnum: String(Numberofpanels),
                    Panelnumbers2: String(Numberofpanels),
                    approvedate: PAM_approvedate,
                    PAM: PAM
                };
                // Fill out form
                for (const fieldName in fields) {
                    const field = form.getTextField(fieldName);
                    if (field && fields[fieldName] !== undefined) {
                        field.setText(fields[fieldName].toString());
                    } else {
                        console.error(`Field ${fieldName} or its corresponding form field is undefined.`);
                    }
                }
                //form.flatten();
                const pdfBytes = await pdfDoc.save();
                fs.writeFileSync('/tmp/filed.pdf', pdfBytes);
                delay(3000);
                // Wait for some time to ensure the file is completely written
                await new Promise(resolve => setTimeout(resolve, 10000)); // 5 seconds delay
                let opts = {
                    format: 'png',
                    out_dir: path.dirname(contractPath),
                    out_prefix: path.basename(contractPath, path.extname(contractPath)), 
                    page: 17
                }
                pdfPoppler.convert(contractPath, opts)
                    .then(async res => {
                        let pngPath = path.join(opts.out_dir, `${opts.out_prefix}-${opts.page}.${opts.format}`);
                        delay(5000);
                        if (fs.existsSync(pngPath)) {
                            const inputPng = fs.readFileSync(pngPath);
                            const outputPng = await sharp(inputPng).extract({
                                left: 140,  // x-coordinate
                                top: 614,   // y-coordinate
                                width: 350, // width
                                height: 139 // height
                            }).toBuffer();
                            await new Promise(resolve => setTimeout(resolve, 5000)); // 5 seconds delay
                            const filedPath = path.resolve('/tmp', 'filed.pdf');
                            const pdfPath = await PDFDocument.load(fs.readFileSync(filedPath));
                            const imagePage = await pdfPath.embedPng(outputPng);
							/*
                            const pagesData = [
                                { pageNum: 9, imageX: 5.2637 * 72, imageY: 8.6126 * 72, imageWidth: 1.6833 * 72, imageHeight: 0.4874 * 72 },
                                { pageNum: 10, imageX: 1.7909 * 72, imageY: 3.3126 * 72, imageWidth: 1.6833 * 72, imageHeight: 0.4874 * 72 },
                                { pageNum: 11, imageX: 0.6091 * 72, imageY: 4.9399 * 72, imageWidth: 1.6833 * 72, imageHeight: 0.4874 * 72 },
                                { pageNum: 16, imageX: 5.9182 * 72, imageY: 4.9762 * 72, imageWidth: 1.4742 * 72, imageHeight: 0.3965 * 72 }, // changed position and dimensions
                                { pageNum: 25, imageX: 2.9 * 72, imageY: 6.4581 * 72, imageWidth: 1.6833 * 72, imageHeight: 0.4874 * 72 },
                                { pageNum: 36, imageX: 1.0091 * 72, imageY: 2.2944 * 72, imageWidth: 1.6833 * 72, imageHeight: 0.4874 * 72 },
                            ];
                            for (let i = 0; i < pagesData.length; i++) {
                                const pageData = pagesData[i];
                                const page = pdfPath.getPages()[pageData.pageNum - 1];
                                page.drawImage(imagePage, {
                                    x: pageData.imageX,
                                    y: pageData.imageY, // from bottom of page
                                    width: pageData.imageWidth,
                                    height: pageData.imageHeight
                                });
                            }
                             */
                            const pdfBytes = await pdfPath.save();
                            fs.writeFileSync('/tmp/filled.pdf', pdfBytes);
                        } else {
                            console.error(`PNG file not found at path: ${pngPath}`);
                        }
                    })
                    .catch(err => console.error(err));
                let FolderID = 'Interconnection_PTO';
				//let folderlink;
                // Assuming caspioAuthorization() is defined and returns an access token
                const FolderIdURL = `https://c1abd578.caspio.com/rest/v2/tables/FolderID/records?q.select=${FolderID}&q.where=Cust_ID=${Cust_ID}`;
                let folderData = await fetchCaspioData(accessToken, FolderIdURL);
                // Assuming the PermitfolderID is the first record in the response data
                
				let PermitfolderIDs;
				let METER = `ACC${i + 1}`;
                if (folderData.Result && folderData.Result.length > 0) {
                    PermitfolderIDs = folderData.Result[0][FolderID];
                }
                if (!PermitfolderIDs) {                             
                            PermitfolderIDs = '1cg70dBWgSqTCwO94dmN4Gb9KbWA0v5bk';
                            //folderlink = `https://drive.google.com/drive/u/0/folders/${folderid}`;
                }
                await new Promise(resolve => setTimeout(resolve, 15000)); // 5 seconds delay
                let signedFormsFolderId= await uploadFile(client, filePath, PermitfolderIDs, NYC_LI,Totalstring,Name,METER);
				//console.log('signedFormsFolderId: ', signedFormsFolderId);

                //await deleteFileWithRetry(filePath);
                await new Promise(resolve => setTimeout(resolve, 15000)); // 5 seconds delay
                const arrays = [
                    { systemCapacity: system, losses: 28, tilt: 20, azimuth: 290 }
                ];
                for (const [index, array] of arrays.entries()) {
                    const filePath1 = await generatePvWattsReport(caspioData.Address, array.systemCapacity, array.losses, array.tilt, array.azimuth, index);
  
                    if (NYC_LI !== 'NYC') {
                        try {
                            await new Promise(resolve => setTimeout(resolve, 15000)); // 5 seconds delay
                            await uploadpvwattsFile(client, filePath1, PermitfolderID);
                            
                            // You can handle post-upload logic here if needed.
                        } catch (error) {
                            console.error('Error calling uploadpvwattsFile: ', error);
                        }
                    } else {
                        console.log('Skipping pvwatts upload as NYC_LI is NYC');
                    }
                }
			    let drivelink = `https://drive.google.com/drive/u/0/folders/${signedFormsFolderId}`;
			    let solarprocessPUT = {
                "intercloseoutlink": drivelink
            };
			let sendurl = `https://c1abd578.caspio.com/rest/v2/tables/SolarProcess/records?q.where=CustID=${Cust_ID}`;
            await fetchCaspioData(accessToken, sendurl, solarprocessPUT, 'PUT');			
            } catch (error) {
                console.error(error);
            }
            fs.unlinkSync(filedPath);
            //fs.unlinkSync('contract-17.png');
            fs.unlinkSync(contractPath);
        }
    }
    })().catch(err => {
        console.error(err);
        process.exit(1);
    });
	