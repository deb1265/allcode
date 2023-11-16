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
//let Cust_ID = 8752;
let Cust_ID = process.argv[2];
const sendurl = `https://c1abd578.caspio.com/rest/v2/tables/SolarProcess/records?q.where=CustID=${Cust_ID}`;		
async function fetchCaspioDataName(accessToken, Cust_ID) {
    const getCust_ID_url = `https://c1abd578.caspio.com/rest/v2/tables/FolderID/records?q.select=Name&q.where=Cust_ID=${Cust_ID}`;

    try {
        const headers = {
            'Authorization': `Bearer ${accessToken}`
        };

        // Attempt to fetch the data
        const responseGet = await axios.get(getCust_ID_url, { headers });

        if (responseGet.data && responseGet.data.Result.length > 0) {
            console.log('Data found...');
            // return the 'Name' from the first result.
            return responseGet.data.Result[0].Name;
        } else {
            console.log('Data not found...');
            return null;
        }

    } catch (error) {
        console.error(`Error in fetchCaspioDataName: ${error}`);
        throw error;
    }
}
	async function createFoldersInDrive(client, customerName, custID, landmark, floodzone, finance, affordable, NYC_LI) {
    const drive = google.drive({ version: 'v3', auth: client });
    const parentFolderId = '18u4Yx6PwJBilgnqMDeHJJ-lkCQX-72Ha';
    const folderIds = {};

    async function checkIfFolderExists(client, folderName, parentFolderId) {
        const response = await drive.files.list({
            q: `mimeType='application/vnd.google-apps.folder' and '${parentFolderId}' in parents and name='${folderName}' and trashed=false`,
            fields: 'files(id, name)',
        });
        return response.data.files;
    }

    async function createFolder(client, folderName, parentFolderId) {
        const drive = google.drive({ version: 'v3', auth: client });
        var fileMetadata = {
            'name': folderName,
            'mimeType': 'application/vnd.google-apps.folder',
            'parents': [parentFolderId]
        };
        const res = await drive.files.create({
            resource: fileMetadata,
            fields: 'id'
        });
        console.log('Folder Id: ', res.data.id);
        return res.data.id;
    }

    async function checkAndCreateFolder(client, folderName, parentFolderId) {
        const existingFolders = await checkIfFolderExists(client, folderName, parentFolderId);
        if (existingFolders.length > 0) {
            console.log(`Folder ${folderName} already exists`);
            return existingFolders[0].id;
        }
        return createFolder(client, folderName, parentFolderId);
    }

    let mainFolderId = await checkAndCreateFolder(client, `${customerName}-${custID}`, parentFolderId);
    folderIds[`${customerName}-${custID}`] = mainFolderId;

    let subFolders = ["Building Permit", "CAD Design", "Interconnection", "Signed Papers", "Utility Bill"]
    if (finance.includes("Sunnova")) {
        subFolders.push("Sunnova Submission");
    }

    function determineNyserdaFolderName() {
        if (NYC_LI === 'LI' && affordable === false && finance.includes("Nyserda")) {
            return "Nyserdaloan";
        } else if (NYC_LI === 'LI' && affordable === true && finance.includes("Nyserda")) {
            return "Nyserda Loan with affordable";
        } else if (NYC_LI === 'LI' && affordable === true && finance.includes("Nyserda")) {
            return "Nyserda with affordable";
        } else if (affordable === true && finance.includes("Nyserda")) {
            return "Nyserda Loan with affordable";
        } else if (affordable === true) {
            return "Nyserda with affordable";
        } else if (NYC_LI === 'NYC' && !finance.includes("Nyserda")) {
            return "Nyserda Incentive application";
        }
        return null;
    }

    const nyserdaFolderName = determineNyserdaFolderName();
    if (nyserdaFolderName !== null) {
        subFolders.push(nyserdaFolderName);
    }

    let subSubFolders = {
        "Building Permit": ["Initial", "Signoff"],
        "CAD Design": ["CAD file", "Design File"],
        "Interconnection": ["Initial", "PTO"]
    };

    if (landmark.toLowerCase() === 'yes') {
        subSubFolders["Building Permit"].push("Landmark Application");
        subSubFolders["CAD Design"][1] = "Landmark Design File";
    }

    if (floodzone.toLowerCase() === 'yes') {
        subSubFolders["CAD Design"][1] = "Flood zone Design File";
    }

    if (landmark.toLowerCase() === 'yes' && floodzone.toLowerCase() === 'yes') {
        subSubFolders["CAD Design"][1] = "Landmark & Flood zone Design File";
    }
    for (const folder of subFolders) {
        let subFolderId = await checkAndCreateFolder(client, folder, mainFolderId);
        folderIds[folder] = { id: subFolderId };

        if (subSubFolders[folder]) {
            for (const subSubFolder of subSubFolders[folder]) {
                let subSubFolderId = await checkAndCreateFolder(client, subSubFolder, subFolderId);
                folderIds[folder][subSubFolder] = subSubFolderId;
            }
        }
    }

    console.log(`Created folders for ${customerName}-${custID}`);
    return folderIds;
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
	async function uploadFile(client, filePath, folderId, Oath_Ecb_Violation, Landmark, Flood_Zone) {
    try {
        const drive = google.drive({ version: 'v3', auth: client });
        const ext = path.extname(filePath);

        // If the file is a PDF, split it into parts and upload each part
        if (ext === '.pdf') {
            const pdfBytes = fs.readFileSync(filePath);
            const pdfDoc = await PDFDocument.load(pdfBytes);

            const parts = {
                'PW3': [0, 1],
                'PTA4': [2, 5],
                'TPP1': [6, 7],
                'TR1': [8, 10],
                'TR8': [11, 12],
                'L2': [13, 14],
                'PW2': [19, 21],
                'Jamaica_sewer':[22,22]
            };

            if (Oath_Ecb_Violation > 0) {
                parts['L2_form'] = [13, 14];
            }
            if (Landmark === "Yes") {
                parts['Landmark_Application'] = [15, 17];
            }
            if (Flood_Zone !== "NA") {
                parts['Special_Improvement'] = [18,18];
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
                        'parents': [folderId],
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
        console.log('File exists after deletion: ', fs.existsSync(filePath));
    } catch (error) {
        console.error('Error uploading file: ', error);
    }
}
    async function caspiomain(CIDvalue, accessToken) {
    try {
        const masterTableUrl = `https://c1abd578.caspio.com/rest/v2/tables/MasterTable/records?q.select=Name%2C%20Address1%2C%20Phone1%2C%20email&q.where=CID=${CIDvalue}`;
        const solarProcessUrl = `https://c1abd578.caspio.com/rest/v2/tables/SolarProcess/records?q.select=NYC_LI%2C%20Flood_Zone%2C%20Finance%2C%20Bin%2C%20Oath_Ecb_Violation%2C%20Contract_Price%2C%20Canopy_Used%2C%20Affordable%2C%20Landmark%2C%20number_of_systems&q.where=CustID=${CIDvalue}`;

        const masterDataResponse = await fetchCaspioData(accessToken, masterTableUrl);
        const solarProcessResponse = await fetchCaspioData(accessToken, solarProcessUrl);
        const masterData = masterDataResponse.Result[0];
        const solarData = solarProcessResponse.Result[0];
        const combinedData = { ...masterData, ...solarData };
        const originalAddress = combinedData.Address1;
        const response = await axios.get(`https://maps.googleapis.com/maps/api/geocode/json?address=${encodeURIComponent(originalAddress)}&key=${GoggleApi}`);
        const Address = response.data.results[0].formatted_address;
        const CID = CIDvalue;
        const Name = combinedData.Name;
        const Finance = combinedData.Finance;
        const Phone = combinedData.Phone1;
        const email = combinedData.email;
        const Price = combinedData.Contract_Price;
        const NYC_LI = combinedData.NYC_LI;
        const Flood_Zone = combinedData.Flood_Zone;
        const Bin = combinedData.Bin;
        const Oath_Ecb_Violation = combinedData.Oath_Ecb_Violation;
        const permit_type = combinedData.Canopy_Used;
        const Affordable = combinedData.Affordable;
        const Landmark = combinedData.Landmark;
        return { CID, Name, Address, Phone, email, Price, NYC_LI, Flood_Zone, Finance, Bin, Oath_Ecb_Violation, permit_type, Affordable, Landmark };
    } catch (error) {
        console.error(`Error in main: ${error}`);
        if (error.response) {
            console.error('Error status:', error.response.status);
            console.error('Error data:', error.response.data);
        }
    }
}
(async () => {     
    try {
        const client = new google.auth.JWT(
            keys.client_email,
            null,
            keys.private_key,
            [
                'https://www.googleapis.com/auth/spreadsheets',
                'https://www.googleapis.com/auth/drive'   // Added Google Drive scope
            ],
        );
        await client.authorize();
        const accessToken = await caspioAuthorization();
        console.log("Successfully connected to googlesheet!");
        console.log("Successfully connected to caspio!");
        let caspiodata = await caspiomain(Cust_ID, accessToken);
        let Landmark = caspiodata.Landmark;
        let Flood_Zone = caspiodata.Flood_Zone;
        let NYC_LI = caspiodata.NYC_LI;
        let Name = caspiodata.Name;
        let Finance = caspiodata.Finance;
        let Affordable = caspiodata.Affordable;
        let Oath_Ecb_Violation = caspiodata.Oath_Ecb_Violation;
        let Bin1 = caspiodata.Bin;
        let BIN = Bin1.toString();
        let leadnumber;
        let permittingfolder = 'Building_Initial';
        let buildingPermitInitialFolderId;
		let mainFolderId;
        let folderIds;
        let sendfolderidurl = "https://c1abd578.caspio.com/rest/v2/tables/FolderID/records";
        const folderidfinder = `https://c1abd578.caspio.com/rest/v2/tables/FolderID/records?q.select=${permittingfolder}&q.where=Cust_ID=${Cust_ID}`
        //let PUTurl = `https://c1abd578.caspio.com/rest/v2/tables/FolderID/records?q.where=Cust_ID=${Cust_ID}`;
        let name = await fetchCaspioDataName(accessToken, Cust_ID);
        console.log("Data: ", caspiodata);
        if (NYC_LI !== 'NYC') {
            let landmark = 'NA';
            let floodzone = 'NA';
            if (name !== null) {
                console.log("Folder already created. Skipping folder creation process...");
				let Name_IDfoldername = 'Name_ID';
const folderidfinder1 = `https://c1abd578.caspio.com/rest/v2/tables/FolderID/records?q.select=${Name_IDfoldername}&q.where=Cust_ID=${Cust_ID}`;
let Name_IDfolder = await fetchCaspioData(accessToken, folderidfinder1);
//console.log(Name_IDfolder);

// Extract the Name_ID from the response object
let Name_ID = Name_IDfolder.Result[0].Name_ID;

// Create the Google Drive link
let drivelink = `https://drive.google.com/drive/u/0/folders/${Name_ID}`;
//console.log(drivelink);

// Your PUT request payload
let solarprocessPUT = {
  "Folder_created": 'Yes',
  "mainfolderlink": drivelink
};

// Your PUT request
await fetchCaspioData(accessToken, sendurl, solarprocessPUT, 'PUT');

			console.log("Link added to caspio...");
            } else {
                console.log("Creating folders please wait...");
				console.log("Loading...");
                folderIds = await createFoldersInDrive(client, Name, Cust_ID, landmark, floodzone, Finance, Affordable, NYC_LI);
                mainFolderId = folderIds[`${Name}-${Cust_ID}`];
                let buildingPermitFolderId = folderIds['Building Permit']['id'];
                let buildingPermitInitialFolderId = folderIds['Building Permit']['Initial'];
                let buildingPermitSignoffFolderId = folderIds['Building Permit']['Signoff'];
                let buildingPermitLandmarkApplicationFolderId = folderIds['Building Permit']['Landmark Application'];
                let cadDesignFolderId = folderIds['CAD Design']['id'];
                let cadDesignCADFileFolderId = folderIds['CAD Design']['CAD file'];
                let cadDesignDesignFileFolderId = folderIds['CAD Design']['Design File'];
                let cadDesignLandmarkDesignFileFolderId = folderIds['CAD Design']['Landmark Design File'];
                let cadDesignFloodzoneDesignFileFolderId = folderIds['CAD Design']['Flood zone Design File'];
                let cadDesignLandmarkFloodzoneDesignFileFolderId = folderIds['CAD Design']['Landmark & Flood zone Design File'];
                let interconnectionFolderId = folderIds['Interconnection']['id'];
                let interconnectionInitialFolderId = folderIds['Interconnection']['Initial'];
                let interconnectionPTOFolderId = folderIds['Interconnection']['PTO'];
                let signedPapersFolderId = folderIds['Signed Papers']['id'];
                let utilityBillFolderId = folderIds['Utility Bill']['id'];
                let sunnovaSubmissionFolderId = folderIds['Sunnova Submission'] ? folderIds['Sunnova Submission']['id'] : null;
                let nyserdaLoanWithAffordableFolderId = folderIds['Nyserda Loan with affordable'] ? folderIds['Nyserda Loan with affordable']['id'] : null;
                let nyserdaWithAffordableFolderId = folderIds['Nyserda with affordable'] ? folderIds['Nyserda with affordable']['id'] : null;
                let nyserdaIncentiveApplicationFolderId = folderIds['Nyserda Incentive application'] ? folderIds['Nyserda Incentive application']['id'] : null;
                let nyserdaloanApplicationFolderId = folderIds['Nyserdaloan'] ? folderIds['Nyserdaloan']['id'] : null;
                let folderPOST = {
                    "Cust_ID": Cust_ID,
                    "Name": Name,
                    "Name_ID": mainFolderId,
                    "Building_Permit": buildingPermitFolderId,
                    "CAD_Design": cadDesignFolderId,
                    "Interconnection": interconnectionFolderId,
                    //"Nyserda": Violation_numbers,
                    "Signed_Papers": signedPapersFolderId,
                    "Utility_Bill": utilityBillFolderId,
                    "Building_Initial": buildingPermitInitialFolderId,
                    "Building_Signoff": buildingPermitSignoffFolderId,
                    "CAD_file": cadDesignCADFileFolderId,
                    "CAD_Design_pdf_folder": cadDesignDesignFileFolderId,
                    "Interconnection_Initial": interconnectionInitialFolderId,
                    "Interconnection_PTO": interconnectionPTOFolderId,
                    "Sunnova": sunnovaSubmissionFolderId,
                    "Building_P_Landmark": buildingPermitLandmarkApplicationFolderId,
                    "cad_designwithlandmark": cadDesignLandmarkDesignFileFolderId,
                    "cad_designwithfloodzone": cadDesignFloodzoneDesignFileFolderId,
                    "cad_designlandmarkfloodzone": cadDesignLandmarkFloodzoneDesignFileFolderId,
                    "nyserdaloanwithaffordable": nyserdaLoanWithAffordableFolderId,
                    "nyserdawithaffordable": nyserdaWithAffordableFolderId,
                    "nyserdaincentiveapplication": nyserdaIncentiveApplicationFolderId,
                    "nyserdaloanApplicationFolderId": nyserdaloanApplicationFolderId
                };
                let folderPUT = {
                    "Name": Name,
                    "Name_ID": mainFolderId,
                    "Building_Permit": buildingPermitFolderId,
                    "CAD_Design": cadDesignFolderId,
                    "Interconnection": interconnectionFolderId,
                    //"Nyserda": Violation_numbers,
                    "Signed_Papers": signedPapersFolderId,
                    "Utility_Bill": utilityBillFolderId,
                    "Building_Initial": buildingPermitInitialFolderId,
                    "Building_Signoff": buildingPermitSignoffFolderId,
                    "CAD_file": cadDesignCADFileFolderId,
                    "CAD_Design_pdf_folder": cadDesignDesignFileFolderId,
                    "Interconnection_Initial": interconnectionInitialFolderId,
                    "Interconnection_PTO": interconnectionPTOFolderId,
                    "Sunnova": sunnovaSubmissionFolderId,
                    "Building_P_Landmark": buildingPermitLandmarkApplicationFolderId,
                    "cad_designwithlandmark": cadDesignLandmarkDesignFileFolderId,
                    "cad_designwithfloodzone": cadDesignFloodzoneDesignFileFolderId,
                    "cad_designlandmarkfloodzone": cadDesignLandmarkFloodzoneDesignFileFolderId,
                    "nyserdaloanwithaffordable": nyserdaLoanWithAffordableFolderId,
                    "nyserdawithaffordable": nyserdaWithAffordableFolderId,
                    "nyserdaincentiveapplication": nyserdaIncentiveApplicationFolderId,
                    "nyserdaloanApplicationFolderId": nyserdaloanApplicationFolderId
                };
                await fetchCaspioData(accessToken, sendfolderidurl, folderPOST, 'POST');
                console.log("Folder ids uploaded in the caspio table...");
            }
        }
        else {
			
        const browser = await puppeteer.launch({
            headless: false,
            slowMo: 20 // delay in milliseconds
        });
        const context = await browser.createIncognitoBrowserContext();
        const page = await context.newPage();  // Now using context to open the page
        const timeout = 30000;
        page.setDefaultTimeout(timeout);
        await page.setViewport({
            width: 1263,
            height: 937
        });
        const newTab = await browser.newPage();
        await newTab.setViewport({
            width: 1263,
            height: 937
        });
		
        await newTab.goto('https://hpdonline.nyc.gov/hpdonline/', { waitUntil: 'networkidle0' });
		try {
        {
            const targetPage1 = newTab;
            await scrollIntoViewIfNeeded([
                [
                    '#p-tabpanel-1-label'
                ],
                [
                    'xpath///*[@id="p-tabpanel-1-label"]'
                ],
                [
                    'pierce/#p-tabpanel-1-label'
                ],
                [
                    'aria/BIN'
                ]
            ], targetPage1, timeout);
            const element = await waitForSelectors([
                [
                    '#p-tabpanel-1-label'
                ],
                [
                    'xpath///*[@id="p-tabpanel-1-label"]'
                ],
                [
                    'pierce/#p-tabpanel-1-label'
                ],
                [
                    'aria/BIN'
                ]
            ], targetPage1, { timeout, visible: true });
            await element.click({
                offset: {
                    x: 43.296875,
                    y: 17,
                },
            });
        }
        {
            const targetPage1 = newTab;
            await scrollIntoViewIfNeeded([
                [
                    'p-tabpanel:nth-of-type(2) input'
                ],
                [
                    'xpath///*[@id="dashboardBuildingBINSearchDesk"]/span/p-inputnumber/span/input'
                ],
                [
                    'pierce/p-tabpanel:nth-of-type(2) input'
                ],
                [
                    'aria/Search a BIN number here'
                ]
            ], targetPage1, timeout);
            const element = await waitForSelectors([
                [
                    'p-tabpanel:nth-of-type(2) input'
                ],
                [
                    'xpath///*[@id="dashboardBuildingBINSearchDesk"]/span/p-inputnumber/span/input'
                ],
                [
                    'pierce/p-tabpanel:nth-of-type(2) input'
                ],
                [
                    'aria/Search a BIN number here'
                ]
            ], targetPage1, { timeout, visible: true });
            await element.click({
                offset: {
                    x: 236,
                    y: 20,
                },
            });
            await element.type(BIN);  // typing the BIN into the input field 
        }
        {
            const targetPage1 = newTab;
            await scrollIntoViewIfNeeded([
                [
                    'p-tabpanel:nth-of-type(2) button'
                ],
                [
                    'xpath///*[@id="dashboardBuildingBINSearchDesk"]/span/p-button/button'
                ],
                [
                    'pierce/p-tabpanel:nth-of-type(2) button'
                ],
                [
                    'aria/Search'
                ]
            ], targetPage1, timeout);
            const element = await waitForSelectors([
                [
                    'p-tabpanel:nth-of-type(2) button'
                ],
                [
                    'xpath///*[@id="dashboardBuildingBINSearchDesk"]/span/p-button/button'
                ],
                [
                    'pierce/p-tabpanel:nth-of-type(2) button'
                ],
                [
                    'aria/Search'
                ]
            ], targetPage1, { timeout, visible: true });
            await element.click({
                offset: {
                    x: 30,
                    y: 14,
                },
            });
            await targetPage1.waitForTimeout(2000);
        }
        {
            const targetPage1 = newTab;
            await scrollIntoViewIfNeeded([
                [
                    'li:nth-of-type(3) > a'
                ],
                [
                    'xpath///*[@id="buildingInfoTabMenuDesk"]/div/div/div/ul/li[3]/a'
                ],
                [
                    'pierce/li:nth-of-type(3) > a'
                ],
                [
                    'aria/Violations[role="link"]'
                ]
            ], targetPage1, timeout);
            const element = await waitForSelectors([
                [
                    'li:nth-of-type(3) > a'
                ],
                [
                    'xpath///*[@id="buildingInfoTabMenuDesk"]/div/div/div/ul/li[3]/a'
                ],
                [
                    'pierce/li:nth-of-type(3) > a'
                ],
                [
                    'aria/Violations[role="link"]'
                ]
            ], targetPage1, { timeout, visible: true });
            await element.click({
                offset: {
                    x: 42.21875,
                    y: 27,
                },
            });
            await targetPage1.waitForTimeout(2000);
        }
        {
                const targetPage1 = newTab;
                await scrollIntoViewIfNeeded([
                    [
                        '#p-tabpanel-5-label'
                    ],
                    [
                        'xpath///*[@id="p-tabpanel-5-label"]'
                    ],
                    [
                        'pierce/#p-tabpanel-5-label'
                    ]
                ], targetPage1, timeout);
                const element = await waitForSelectors([
                    [
                        '#p-tabpanel-5-label'
                    ],
                    [
                        'xpath///*[@id="p-tabpanel-5-label"]'
                    ],
                    [
                        'pierce/#p-tabpanel-5-label'
                    ]
                ], targetPage1, { timeout, visible: true });
                await element.click({
                    offset: {
                        x: 95.09375,
                        y: 24,
                    },
                });
                await targetPage1.waitForTimeout(2000);
                // After clicking, get the text content from the element and parse out the number.
                const textContent = await element.evaluate(el => el.textContent);
                const numberMatch = textContent.match(/\((\d+)\)/); // This regex matches a number enclosed in parentheses.
                
                if (numberMatch) {
                    leadnumber = Number(numberMatch[1]);
                    console.log('Lead Violations:', leadnumber)
                } else {
                    console.log('No match found for number in parentheses.');
                }

            }

       } catch (error) {
      console.log('Error in first block:', error);
    }
			
            let name = await fetchCaspioDataName(accessToken, Cust_ID);
            
            if (name !== null) {
                console.log("Folder already created. Skipping folder creation process...");
let Name_IDfoldername = 'Name_ID';
let buildingpermitfolder = 'Building_Initial';
const folderidfinder1 = `https://c1abd578.caspio.com/rest/v2/tables/FolderID/records?q.select=${Name_IDfoldername}&q.where=Cust_ID=${Cust_ID}`;
const folderidfinder2 = `https://c1abd578.caspio.com/rest/v2/tables/FolderID/records?q.select=${buildingpermitfolder}&q.where=Cust_ID=${Cust_ID}`;
let Name_IDfolder = await fetchCaspioData(accessToken, folderidfinder1);
let buildingPermitInitialFolder_ = await fetchCaspioData(accessToken, folderidfinder2);
//console.log(Name_IDfolder);

// Extract the Name_ID from the response object
let Name_ID = Name_IDfolder.Result[0].Name_ID;
buildingPermitInitialFolderId = buildingPermitInitialFolder_.Result[0].Building_Initial;

// Create the Google Drive link
let drivelink = `https://drive.google.com/drive/u/0/folders/${Name_ID}`;
//console.log(drivelink);

// Your PUT request payload
let solarprocessPUT = {
  "Folder_created": 'Yes',
  "mainfolderlink": drivelink
};

// Your PUT request
await fetchCaspioData(accessToken, sendurl, solarprocessPUT, 'PUT');
console.log("Link added to caspio...");

            } else {
                console.log("Starting folder creation process....");
                folderIds = await createFoldersInDrive(client, Name, Cust_ID, Landmark, Flood_Zone, Finance, Affordable, NYC_LI);
                mainFolderId = folderIds[`${Name}-${Cust_ID}`];
                let buildingPermitFolderId = folderIds['Building Permit']['id'];
                buildingPermitInitialFolderId = folderIds['Building Permit']['Initial'];
                let buildingPermitSignoffFolderId = folderIds['Building Permit']['Signoff'];
                let buildingPermitLandmarkApplicationFolderId = folderIds['Building Permit']['Landmark Application'];
                let cadDesignFolderId = folderIds['CAD Design']['id'];
                let cadDesignCADFileFolderId = folderIds['CAD Design']['CAD file'];
                let cadDesignDesignFileFolderId = folderIds['CAD Design']['Design File'];
                let cadDesignLandmarkDesignFileFolderId = folderIds['CAD Design']['Landmark Design File'];
                let cadDesignFloodzoneDesignFileFolderId = folderIds['CAD Design']['Flood zone Design File'];
                let cadDesignLandmarkFloodzoneDesignFileFolderId = folderIds['CAD Design']['Landmark & Flood zone Design File'];
                let interconnectionFolderId = folderIds['Interconnection']['id'];
                let interconnectionInitialFolderId = folderIds['Interconnection']['Initial'];
                let interconnectionPTOFolderId = folderIds['Interconnection']['PTO'];
                let signedPapersFolderId = folderIds['Signed Papers']['id'];
                let utilityBillFolderId = folderIds['Utility Bill']['id'];
                let sunnovaSubmissionFolderId = folderIds['Sunnova Submission'] ? folderIds['Sunnova Submission']['id'] : null;
                let nyserdaLoanWithAffordableFolderId = folderIds['Nyserda Loan with affordable'] ? folderIds['Nyserda Loan with affordable']['id'] : null;
                let nyserdaWithAffordableFolderId = folderIds['Nyserda with affordable'] ? folderIds['Nyserda with affordable']['id'] : null;
                let nyserdaIncentiveApplicationFolderId = folderIds['Nyserda Incentive application'] ? folderIds['Nyserda Incentive application']['id'] : null;
                let nyserdaloanApplicationFolderId = folderIds['Nyserdaloan'] ? folderIds['Nyserdaloan']['id'] : null;
                let folderPOST = {
                    "Cust_ID": Cust_ID,
                    "Name": Name,
                    "Name_ID": mainFolderId,
                    "Building_Permit": buildingPermitFolderId,
                    "CAD_Design": cadDesignFolderId,
                    "Interconnection": interconnectionFolderId,
                    //"Nyserda": Violation_numbers,
                    "Signed_Papers": signedPapersFolderId,
                    "Utility_Bill": utilityBillFolderId,
                    "Building_Initial": buildingPermitInitialFolderId,
                    "Building_Signoff": buildingPermitSignoffFolderId,
                    "CAD_file": cadDesignCADFileFolderId,
                    "CAD_Design_pdf_folder": cadDesignDesignFileFolderId,
                    "Interconnection_Initial": interconnectionInitialFolderId,
                    "Interconnection_PTO": interconnectionPTOFolderId,
                    "Sunnova": sunnovaSubmissionFolderId,
                    "Building_P_Landmark": buildingPermitLandmarkApplicationFolderId,
                    "cad_designwithlandmark": cadDesignLandmarkDesignFileFolderId,
                    "cad_designwithfloodzone": cadDesignFloodzoneDesignFileFolderId,
                    "cad_designlandmarkfloodzone": cadDesignLandmarkFloodzoneDesignFileFolderId,
                    "nyserdaloanwithaffordable": nyserdaLoanWithAffordableFolderId,
                    "nyserdawithaffordable": nyserdaWithAffordableFolderId,
                    "nyserdaincentiveapplication": nyserdaIncentiveApplicationFolderId,
                    "nyserdaloanApplicationFolderId": nyserdaloanApplicationFolderId
                };
                await fetchCaspioData(accessToken, sendfolderidurl, folderPOST, 'POST');
                console.log("Caspio CRM portal updated successfully..");           

            let drivelink = `https://drive.google.com/drive/u/0/folders/${mainFolderId}`;
            let solarprocessPUT = {
                "Folder_created": 'Yes',
                "Lead": leadnumber,
                "mainfolderlink": drivelink
            }
			await fetchCaspioData(accessToken, sendurl, solarprocessPUT, 'PUT');  
			}
			{
                const targetPage1 = newTab;
                await targetPage1.waitForTimeout(3000);  // Optional: Wait for 3 seconds for the page to load completely.
                const screenshotPath = 'C:\\Users\\deb\\\Desktop\\py\\No_lead_violation.png';
                await targetPage1.screenshot({ path: screenshotPath });
                await new Promise(resolve => setTimeout(resolve, 5000)); // increased delay to 5 seconds
                // check if file exists and upload it
                if (fs.existsSync(screenshotPath)) {
                    await uploadFile(client, screenshotPath, buildingPermitInitialFolderId, Oath_Ecb_Violation, Landmark, Flood_Zone);
                } else {
                    console.log("Screenshot file does not exist yet.");
                }
            }
            await browser.close();
            console.log('browser closed');         
        }

    } catch (error) {                                                                
                    console.error(error);
                }
    })().catch(err => {
                console.error(err);
                process.exit(1);
    });

async function waitForSelectors(selectors, frame, options) {
    for (const selector of selectors) {
        try {
            return await waitForSelector(selector, frame, options);
        } catch (err) {
            console.error(err);
        }
    }
    throw new Error('Could not find element for selectors: ' + JSON.stringify(selectors));
}

async function scrollIntoViewIfNeeded(selectors, frame, timeout) {
    const element = await waitForSelectors(selectors, frame, { visible: false, timeout });
    if (!element) {
        throw new Error(
            'The element could not be found.'
        );
    }
    await waitForConnected(element, timeout);
    const isInViewport = await element.isIntersectingViewport({ threshold: 0 });
    if (isInViewport) {
        return;
    }
    await element.evaluate(element => {
        element.scrollIntoView({
            block: 'center',
            inline: 'center',
            behavior: 'auto',
        });
    });
    await waitForInViewport(element, timeout);
}

async function waitForConnected(element, timeout) {
    await waitForFunction(async () => {
        return await element.getProperty('isConnected');
    }, timeout);
}

async function waitForInViewport(element, timeout) {
    await waitForFunction(async () => {
        return await element.isIntersectingViewport({ threshold: 0 });
    }, timeout);
}

async function waitForSelector(selector, frame, options) {
    if (!Array.isArray(selector)) {
        selector = [selector];
    }
    if (!selector.length) {
        throw new Error('Empty selector provided to waitForSelector');
    }
    let element = null;
    for (let i = 0; i < selector.length; i++) {
        const part = selector[i];
        if (element) {
            element = await element.waitForSelector(part, options);
        } else {
            element = await frame.waitForSelector(part, options);
        }
        if (!element) {
            throw new Error('Could not find element: ' + selector.join('>>'));
        }
        if (i < selector.length - 1) {
            element = (await element.evaluateHandle(el => el.shadowRoot ? el.shadowRoot : el)).asElement();
        }
    }
    if (!element) {
        throw new Error('Could not find element: ' + selector.join('|'));
    }
    return element;
}

async function waitForElement(step, frame, timeout) {
    const {
        count = 1,
        operator = '>=',
        visible = true,
        properties,
        attributes,
    } = step;
    const compFn = {
        '==': (a, b) => a === b,
        '>=': (a, b) => a >= b,
        '<=': (a, b) => a <= b,
    }[operator];
    await waitForFunction(async () => {
        const elements = await querySelectorsAll(step.selectors, frame);
        let result = compFn(elements.length, count);
        const elementsHandle = await frame.evaluateHandle((...elements) => {
            return elements;
        }, ...elements);
        await Promise.all(elements.map((element) => element.dispose()));
        if (result && (properties || attributes)) {
            result = await elementsHandle.evaluate(
                (elements, properties, attributes) => {
                    for (const element of elements) {
                        if (attributes) {
                            for (const [name, value] of Object.entries(attributes)) {
                                if (element.getAttribute(name) !== value) {
                                    return false;
                                }
                            }
                        }
                        if (properties) {
                            if (!isDeepMatch(properties, element)) {
                                return false;
                            }
                        }
                    }
                    return true;

                    function isDeepMatch(a, b) {
                        if (a === b) {
                            return true;
                        }
                        if ((a && !b) || (!a && b)) {
                            return false;
                        }
                        if (!(a instanceof Object) || !(b instanceof Object)) {
                            return false;
                        }
                        for (const [key, value] of Object.entries(a)) {
                            if (!isDeepMatch(value, b[key])) {
                                return false;
                            }
                        }
                        return true;
                    }
                },
                properties,
                attributes
            );
        }
        await elementsHandle.dispose();
        return result === visible;
    }, timeout);
}

async function querySelectorsAll(selectors, frame) {
    for (const selector of selectors) {
        const result = await querySelectorAll(selector, frame);
        if (result.length) {
            return result;
        }
    }
    return [];
}

async function querySelectorAll(selector, frame) {
    if (!Array.isArray(selector)) {
        selector = [selector];
    }
    if (!selector.length) {
        throw new Error('Empty selector provided to querySelectorAll');
    }
    let elements = [];
    for (let i = 0; i < selector.length; i++) {
        const part = selector[i];
        if (i === 0) {
            elements = await frame.$$(part);
        } else {
            const tmpElements = elements;
            elements = [];
            for (const el of tmpElements) {
                elements.push(...(await el.$$(part)));
            }
        }
        if (elements.length === 0) {
            return [];
        }
        if (i < selector.length - 1) {
            const tmpElements = [];
            for (const el of elements) {
                const newEl = (await el.evaluateHandle(el => el.shadowRoot ? el.shadowRoot : el)).asElement();
                if (newEl) {
                    tmpElements.push(newEl);
                }
            }
            elements = tmpElements;
        }
    }
    return elements;
}

async function waitForFunction(fn, timeout) {
    let isActive = true;
    const timeoutId = setTimeout(() => {
        isActive = false;
    }, timeout);
    while (isActive) {
        const result = await fn();
        if (result) {
            clearTimeout(timeoutId);
            return;
        }
        await new Promise(resolve => setTimeout(resolve, 100));
    }
    throw new Error('Timed out');
}

async function changeSelectElement(element, value) {
    await element.select(value);
    await element.evaluateHandle((e) => {
        e.blur();
        e.focus();
    });
}

async function changeElementValue(element, value) {
    await element.focus();
    await element.evaluate((input, value) => {
        input.value = value;
        input.dispatchEvent(new Event('input', { bubbles: true }));
        input.dispatchEvent(new Event('change', { bubbles: true }));
    }, value);
}

async function typeIntoElement(element, value) {
    const textToType = await element.evaluate((input, newValue) => {
        if (
            newValue.length <= input.value.length ||
            !newValue.startsWith(input.value)
        ) {
            input.value = '';
            return newValue;
        }
        const originalValue = input.value;
        input.value = '';
        input.value = originalValue;
        return newValue.substring(originalValue.length);
    }, value);
    await element.type(textToType);
}








