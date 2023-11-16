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
const caspioTokenUrl = 'https://c1abd578.caspio.com/oauth/token';
const clientID = process.env.clientID;
const clientSecret = process.env.clientSecret;
const GoggleApi = process.env.GoggleApi;
const keys =JSON.parse(process.env.SERVICE_KEY);
const pdfPoppler = require('pdf-poppler');
const nycsformsfileId = '1hrRvM_HOddeElzdgivcRYXa45jM7sc6D';
const sharp = require('sharp');
const delay = (ms) => new Promise((resolve) => setTimeout(resolve, ms));
//const nycblanksdownloadedpath = path.resolve(__dirname, 'file.pdf'); // define nycblanksdownloadedpath outside the functions
//const contractPath = path.resolve(__dirname, 'contract.pdf'); // define nycblanksdownloadedpath outside the functions
const nycblanksdownloadedpath = path.resolve('/tmp', 'file.pdf'); // Corrected path for nycblanksdownloadedpath
const contractPath = path.resolve('/tmp', 'contract.pdf'); // Corrected path for contractPath
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
async function uploadFile(client, filePath, folderId, Oath_Ecb_Violation, Landmark, Flood_Zone) {
    try {
        const drive = google.drive({ version: 'v3', auth: client });
        const ext = path.extname(filePath);
        const signedFormsFolderId = await createFolder('updated forms filled', folderId, drive);
        // If the file is a PDF, split it into parts and upload each part
        if (ext === '.pdf') {
            const pdfBytes = fs.readFileSync(filePath);
            const pdfDoc = await PDFDocument.load(pdfBytes);

            const parts = {
                'PW3': [0, 1],
                'PTA4': [2, 5],
                'PW1_modified_cost': [6, 10]
            };


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
        }
		return(signedFormsFolderId);
        console.log('File exists after deletion: ', fs.existsSync(filePath));
    } catch (error) {
        console.error('Error uploading file: ', error);
    }
}
async function caspiomain(CIDvalue, accessToken) {
    try {
        const masterTableUrl = `https://c1abd578.caspio.com/rest/v2/tables/MasterTable/records?q.select=Name%2C%20Address1%2C%20Phone1%2C%20email&q.where=CID=${CIDvalue}`;
        const solarProcessUrl = `https://c1abd578.caspio.com/rest/v2/tables/SolarProcess/records?q.select=ReducedDOBcost%2CPanel%2CSystem_Size_Sold%2CBlock%2CLot%2CBin%2CFlood_Zone%2CNYC_LI%2COath_Ecb_Violation%2CFlood_Zone%2CLandmark%2CBorough%2CStories%2CCB&q.where=CustID=${CIDvalue}`;
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
        const Contract_Price = combinedData.ReducedDOBcost;
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
// Use a regex to find the ZIP code in the entire address string
let zipRegex = /\b\d{5}\b/;  // Matches exactly 5 digits
let zipMatch = Addressp.match(zipRegex);
let Zip;
if (zipMatch) {
    Zip = zipMatch[0];
} else {
    console.error(`Could not find ZIP code in ${Addressp}`);
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
    const sheetData = [1];
     const filePath = '/tmp/filled.pdf'; // Correct path for storing a file in the temporary directory on Google App Engine
    for (let i = 0; i < sheetData.length; i++) {
        try {
            if (fs.existsSync(filePath)) {
                await deleteFileWithRetry(filePath);
            }
            let Cust_ID =process.argv[2];
            //const Cust_ID = 8816;
            let caspioData = await caspiomain(Cust_ID, accessToken);
            console.log("Data: ", caspioData);
            const blankpdfpath = await PDFDocument.load(fs.readFileSync(nycblanksdownloadedpath));
            const pdfDoc = blankpdfpath;
            const form = pdfDoc.getForm();
            const totalCost = Math.round(parseFloat(caspioData.Contract_Price));
            const stories = parseFloat(caspioData.Stories);
            let panel = caspioData.Panel;
            let panelbrand;
            if (panel.includes("REC")) {
                panelbrand = "REC";
            } else if (panel.includes("Solaria")) {
                panelbrand = "Solaria";
            } else if (panel.includes("CT")) {
                panelbrand = "Certainteed";
            } else if (panel.includes("S-Energy")) {
                panelbrand = "S-Energy";
            } else {
                panelbrand = "Trina";
            }
            let Borough = caspioData.Borough;
            let boroughNum;
            let boroughName = caspioData.Borough.toLowerCase().replace(/\s/g, '');
            const boroughMapping = {
                'queens': '4',
                'brooklyn': '3',
                'bronx': '2',
                'manhattan': '1',
                'statenisland': '5'
            };
            boroughNum = boroughMapping[boroughName];
            console.log(boroughNum); // outputs the number associated with the borough name
            let block = caspioData.Block;
            let contractfileid = caspioData.contractfileid;
            await downloadFile(drive, contractfileid, contractPath);
            console.log("Successfully connected to Caspio!");
            let lot = caspioData.Lot;
            let Flood_Zone = caspioData.Flood_Zone;
            let Oath_Ecb_Violation = caspioData.Oath_Ecb_Violation;
            const dateToday = moment().format('MM.DD.YYYY');
            let systemsize = caspioData.Size.toString();
            let Landmark = caspioData.Landmark;
            const formatCurrency = async (number) => {
                return new Intl.NumberFormat('en-US', { style: 'currency', currency: 'USD' }).format(number);
            }
            const formatPercent = async (number) => {
                return `${number}%`;
            }



            const fields = {
                Labor: Math.round(0.20 * totalCost),
                Materials: Math.round(0.30 * totalCost),
                Equipment: Math.round(0.25 * totalCost),
                DA_cost: Math.round(0.25 * totalCost),
                DA_Labor: 0,   // Initialize dependent fields
                L_M_E_cost: 0, // Initialize dependent fields
                D_cost: 0,     // Initialize dependent fields
                A_cost: 0,     // Initialize dependent fields
                First: caspioData.Firstname,
                Last: caspioData.Lastname,
                House: caspioData.House,
                Street: caspioData.Street,
                Borough: Borough,
                Block: block,
                Lot: lot,
                Bin: caspioData.Bin,
                Stories: caspioData.Stories,
                CB: caspioData.CB,
                //lead_violation: caspioData.lead === '0' || caspioData.lead === '' || caspioData.lead === undefined ? "No lead violation exist" : "lead violation exist",
                Zip: caspioData.Zip,
                Kw: systemsize,
                Panel_Model: caspioData.Panel,
                Brand: panelbrand,
                Totalcost: await formatCurrency(totalCost),
                Date: dateToday,
                //Fullname: caspioData.FullName,

            };
            // Handling panel address, city, state, zip
            const panelBrand = process.env.panelbrand;
            const panelAddresses = {
                REC: { address: "1820 Gateway Drive Suite 170", city: "San Mateo", state: "CA", zip: "94404" },
                Solaria: { address: "45700 Northport Loop E", city: "Fremont", state: "CA", zip: "94538" },
                'S-energy': { address: "1170 N Gilbert St", city: "Anaheim", state: "CA", zip: "92801" },
                Certainteed: { address: "20 Moores Rd", city: "Malvern", state: "PA", zip: "19355" },
                Trina: { address: "1425 K Street, N.W., Suite 1000", city: "Washington", state: "DC", zip: "20005" }
            };
            const panelDetails = panelAddresses[panelBrand] || { address: "", city: "", state: "", zip: "" };
            // Fill out the fields
            const panelAddressField = form.getTextField('PanelAddress');
            panelAddressField.setText(panelDetails.address);
            const panelCityField = form.getTextField('Panel_city');
            panelCityField.setText(panelDetails.city);
            const panelStateField = form.getTextField('Panel_state');
            panelStateField.setText(panelDetails.state);
            const panelZipField = form.getTextField('Panel_zip');
            panelZipField.setText(panelDetails.zip);
            // Calculate dependent fields
            fields.DA_Labor = Math.round(fields.DA_cost + fields.Labor);
            fields.L_M_E_cost = Math.round(fields.Labor + fields.Materials + fields.Equipment);
            fields.D_cost = Math.round(fields.DA_cost * 0.60);
            fields.A_cost = Math.round(fields.DA_cost - fields.D_cost);
            //fields.Height = 13 * stories;
            fields.Height = '';
            // Convert the fields to string with $ sign
            // Fill out form
            for (const fieldName in fields) {
                const field = form.getTextField(fieldName);
                if (field && fields[fieldName] !== undefined) {
                    field.setText(fields[fieldName].toString());
                } else {
                    console.error(`Field ${fieldName} or its corresponding form field is undefined.`);
                }
            }
            form.flatten();
            const pdfBytes = await pdfDoc.save();
            fs.writeFileSync('/tmp/filed.pdf', pdfBytes);
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

                    if (fs.existsSync(pngPath)) {
                        const inputPng = fs.readFileSync(pngPath);
                        const outputPng = await sharp(inputPng).extract({
                            left: 191,  // x-coordinate
                            top: 660,   // y-coordinate
                            width: 220, // width
                            height: 83 // height
                        }).toBuffer();
                        await new Promise(resolve => setTimeout(resolve, 5000)); // 5 seconds delay
                        const filedPath = path.resolve('/tmp', 'filed.pdf'); // define pdfPath outside the functions
                        const pdfPath = await PDFDocument.load(fs.readFileSync(filedPath));
                        const imagePage = await pdfPath.embedPng(outputPng);
                        const pagesData = [
                            { pageNum: 2, imageX: 1.2947 * 72, imageY: 3.8382 * 72, imageWidth: 1.1926 * 72, imageHeight: 0.314 * 72 },
                            { pageNum: 6, imageX: 5.0931 * 72, imageY: 1.404 * 72, imageWidth: 1.1926 * 72, imageHeight: 0.314 * 72 },
                            { pageNum: 11, imageX: 5.5403 * 72, imageY: 4.2889 * 72, imageWidth: 0.7789 * 72, imageHeight: 0.2826 * 72 },
                           // { pageNum: 11, imageX: 1.4712 * 72, imageY: 8.2763 * 72, imageWidth: 1.0294 * 72, imageHeight: 0.2521 * 72 }, // changed position and dimensions
                           // { pageNum: 18, imageX: 3.6861 * 72, imageY: 3.8435 * 72, imageWidth: 1.1926 * 72, imageHeight: 0.314 * 72 },
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

                        const pdfBytes = await pdfPath.save();
                        fs.writeFileSync('filled.pdf', pdfBytes);
                    } else {
                        console.error(`PNG file not found at path: ${pngPath}`);
                    }
                })
                .catch(err => console.error(err));
        

            let FolderID = 'Building_Initial';
            // Assuming caspioAuthorization() is defined and returns an access token
            const FolderIdURL = `https://c1abd578.caspio.com/rest/v2/tables/FolderID/records?q.select=${FolderID}&q.where=Cust_ID=${Cust_ID}`
            let folderData = await fetchCaspioData(accessToken, FolderIdURL);
            // Assuming the PermitfolderID is the first record in the response data
            let PermitfolderID;
            if (folderData.Result && folderData.Result.length > 0) {
                PermitfolderID = folderData.Result[0][FolderID];
            }
            if (!PermitfolderID) {
                throw new Error('Could not find the target folder');
            }
            await new Promise(resolve => setTimeout(resolve, 15000)); // 5 seconds delay
            let signedFormsFolderId=await uploadFile(client, filePath, PermitfolderID, Oath_Ecb_Violation, Landmark, Flood_Zone);
            //await deleteFileWithRetry(filePath);
            await new Promise(resolve => setTimeout(resolve, 5000)); // 5 seconds delay
				let drivelink = `https://drive.google.com/drive/u/0/folders/${signedFormsFolderId}`;
			    let solarprocessPUT = {
                "updatedcostefilingformslink": drivelink
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
})().catch(err => {
    console.error(err);
    process.exit(1);
});