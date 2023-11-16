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
//const { exec } = require('child_process');
const { spawn } = require('child_process');
const { PDFDocument } = require('pdf-lib');
const path = require('path');
const url = require('url');
// Create the Google API client
// Define the Google API client

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
        const solarProcessUrl = `https://c1abd578.caspio.com/rest/v2/tables/SolarProcess/records?q.select=NYC_LI%2CFinance%2CSystem_Size_Sold%2CPanel%2CNumber_Of_Panels%2CTotalprice%2CNYSUN_incentive_amount%2CNYSUN_Application_type%2CCanopy_Used%2CAffordable%2CCounty%2CTilt1_%2CTilt2_%2CTilt3_%2CTilt4_%2CTilt5_%2CTilt6_%2CTilt7_%2CAz1_%2CAz2_%2CAz3_%2CAz4_%2CAz5_%2CAz6_%2CAz7_%2CTsrf1_%2CTsrf2_%2CTsrf3_%2CTsrf4_%2CTsrf5_%2CTsrf6_%2CTsrf7_%2CPanel1_%2CPanel2_%2CPanel3_%2CPanel4_%2CPanel5_%2CPanel6_%2CPanel7_&q.where=CustID=${CIDvalue}`;

        const masterDataResponse = await fetchCaspioData(accessToken, masterTableUrl);
        const solarProcessResponse = await fetchCaspioData(accessToken, solarProcessUrl);

        const masterData = masterDataResponse.Result[0];
        const solarData = solarProcessResponse.Result[0];

        const combinedData = { ...masterData, ...solarData };

        const originalAddress = combinedData.Address1;

        const response = await axios.get(`https://maps.googleapis.com/maps/api/geocode/json?address=${encodeURIComponent(originalAddress)}&key=${GoggleApi}`);
        const Address = response.data.results[0].formatted_address;
        const addressComponents = response.data.results[0].address_components;
        let County = '';
        for (let i = 0; i < addressComponents.length; i++) {
            if (addressComponents[i].types.includes('administrative_area_level_2')) {
                County = addressComponents[i].long_name;
                break;
            }
        }
        County = County.split(' ')[0];
        const CID = CIDvalue;
        const Name = combinedData.Name;
        const Phone = combinedData.Phone1;
        const email = combinedData.email.trim();
        const Price = combinedData.Totalprice;
        const NYC_LI = combinedData.NYC_LI;
        const Finance = combinedData.Finance;
        const Size = combinedData.System_Size_Sold;
        const Panel = combinedData.Panel;
        const Panel_count = combinedData.Number_Of_Panels;
        const Canopy = combinedData.Canopy_Used;
        const Affordable = combinedData.Affordable;
        //const County = combinedData.County;
        const NYSUN_Application_type = combinedData.NYSUN_Application_type;
        const Tilt1 = combinedData.Tilt1_ ? combinedData.Tilt1_.toString() : '';
        const Tilt2 = combinedData.Tilt2_ ? combinedData.Tilt2_.toString() : '';
        const Tilt3 = combinedData.Tilt3_ ? combinedData.Tilt3_.toString() : '';
        const Tilt4 = combinedData.Tilt4_ ? combinedData.Tilt4_.toString() : '';
        const Tilt5 = combinedData.Tilt5_ ? combinedData.Tilt5_.toString() : '';
        const Tilt6 = combinedData.Tilt6_ ? combinedData.Tilt6_.toString() : '';
        const Tilt7 = combinedData.Tilt7_ ? combinedData.Tilt7_.toString() : '';
        const Az1 = combinedData.Az1_ ? combinedData.Az1_.toString() : '';
        const Az2 = combinedData.Az2_ ? combinedData.Az2_.toString() : '';
        const Az3 = combinedData.Az3_ ? combinedData.Az3_.toString() : '';
        const Az4 = combinedData.Az4_ ? combinedData.Az4_.toString() : '';
        const Az5 = combinedData.Az5_ ? combinedData.Az5_.toString() : '';
        const Az6 = combinedData.Az6_ ? combinedData.Az6_.toString() : '';
        const Az7 = combinedData.Az7_ ? combinedData.Az7_.toString() : '';
        const Tsrf1 = combinedData.Tsrf1_ ? combinedData.Tsrf1_.toString() : '';
        const Tsrf2 = combinedData.Tsrf2_ ? combinedData.Tsrf2_.toString() : '';
        const Tsrf3 = combinedData.Tsrf3_ ? combinedData.Tsrf3_.toString() : '';
        const Tsrf4 = combinedData.Tsrf4_ ? combinedData.Tsrf4_.toString() : '';
        const Tsrf5 = combinedData.Tsrf5_ ? combinedData.Tsrf5_.toString() : '';
        const Tsrf6 = combinedData.Tsrf6_ ? combinedData.Tsrf6_.toString() : '';
        const Tsrf7 = combinedData.Tsrf7_ ? combinedData.Tsrf7_.toString() : '';
        const Panel1 = combinedData.Panel1_ ? combinedData.Panel1_.toString() : '';
        const Panel2 = combinedData.Panel2_ ? combinedData.Panel2_.toString() : '';
        const Panel3 = combinedData.Panel3_ ? combinedData.Panel3_.toString() : '';
        const Panel4 = combinedData.Panel4_ ? combinedData.Panel4_.toString() : '';
        const Panel5 = combinedData.Panel5_ ? combinedData.Panel5_.toString() : '';
        const Panel6 = combinedData.Panel6_ ? combinedData.Panel6_.toString() : '';
        const Panel7 = combinedData.Panel7_ ? combinedData.Panel7_.toString() : '';

        let FullName = Name; // Replace with your actual full name variable
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
        const HouseAndStreet = addressParts[0].trim();
        // Get the city
        const City = addressParts[1].trim();
        //if (City.includes.tolowercase.trim(queens)) {
       //     County = `Queens`;
       // }   else if (City.includes.tolowercase.trim(brooklyn)) {
        //        County = `Brooklyn`;
         //   } else if (City.includes.tolowercase.trim(bronx)) {
         //       County = `Bronx`;
       // } else if (City.includes.tolowercase.trim(manhattan)) {
       //     County = `New York`;
      //  } else if (City.includes.tolowercase.trimanyspacesandspecialcharacters(statenisland)) {
       //     County = `Richmond`;
       // }
        const StateAndZip = addressParts[2].trim();
        let StateZipParts = StateAndZip.split(' ');
        const Zip = StateZipParts[1].trim();
        // Get the zip
        // Assuming that the address format is always "House, City, State, Zip"
        // Zip would be in the 3rd index position after splitting
        const Tilt = [Tilt1, Tilt2, Tilt3, Tilt4, Tilt5, Tilt6, Tilt7];
        const Az = [Az1, Az2, Az3, Az4, Az5, Az6, Az7];
        const Tsrf = [Tsrf1, Tsrf2, Tsrf3, Tsrf4, Tsrf5, Tsrf6, Tsrf7];
        const Panels = [Panel1, Panel2, Panel3, Panel4, Panel5, Panel6, Panel7];
        return { Firstname, Lastname, HouseAndStreet, Address, City, Zip, Panel_count, Phone, email, Price, Panel, NYC_LI, Size, Finance, Canopy, Affordable, County, NYSUN_Application_type, datarow: { Panels, Tilt, Az, Tsrf } };
    } catch (error) {
        console.error(`Error in main: ${error}`);
        if (error.response) {
            console.error('Error status:', error.response.status);
            console.error('Error data:', error.response.data);
        }
    }
}  
async function downloadPDF(page, firstName, lastName) {
    // Extract the ProjectId from the URL
    const currentUrl = page.url();
    const urlObj = new url.URL(currentUrl);
    const projectId = urlObj.searchParams.get('ProjectId');

    // Get cookies from Puppeteer and format them for Axios
    const cookies = await page.cookies();
    const cookieString = cookies.map(cookie => `${cookie.name}=${cookie.value}`).join('; ');

    const response = await axios.get(`https://portal.nyserda.ny.gov/apex/NYSUN_APPINTAKE_ReviewPDF_Page?ProjectId=${projectId}`, {
        responseType: 'arraybuffer', // This is important to handle PDF data
        headers: {
            'User-Agent': 'Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/114.0.0.0 Safari/537.36',
            'Accept': 'application/pdf',
            'Cookie': cookieString
        },
    });

    // Define directory and create it if it doesn't exist
    const directory = './'; // or any other directory
    if (!fs.existsSync(directory)) {
        fs.mkdirSync(directory, { recursive: true });
    }

    const pdfPath = path.join(directory, `${firstName} ${lastName} ${projectId}.pdf`);

    fs.writeFileSync(pdfPath, response.data);
    return pdfPath;
}

async function uploadFile(client, filePath, folderId) {
    try {
        const drive = google.drive({ version: 'v3', auth: client });
        const ext = path.extname(filePath);

        // If the file is a PDF, upload it directly
        if (ext === '.pdf') {
            const response = await drive.files.create({
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

            console.log('File Id: ', response.data.id);
            fs.unlinkSync(filePath);
        }
        // If the file is a PNG, upload it directly
        else if (ext === '.png') {
            const response = await drive.files.create({
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

            console.log('File Id: ', response.data.id);
            fs.unlinkSync(filePath);
        }
    } catch (error) {
        console.error('Error uploading file: ', error);
    }
}

async function processPDF(page, firstName, lastName, client, nyserdaFolderId) {
    // Download the PDF and get the file path
    const pdfPath = await downloadPDF(page, firstName, lastName);

    // Upload the downloaded file to Google Drive
    await uploadFile(client, pdfPath, nyserdaFolderId);
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

async function writeSpreadsheetData(auth, data, cell) {
    const sheets = google.sheets({ version: 'v4', auth });
    const spreadsheetId = '1aXRq3yh7KMwFJmTBWGL74Hp56Vns8XGcsqwznIkO2tc';
    const range = 'Sheet1!' + cell;

    await sheets.spreadsheets.values.update({
        spreadsheetId,
        range,
        valueInputOption: 'USER_ENTERED',
        resource: {
            values: [[data]],
        },
    });
}
async function retryClickElement(page, selector, maxRetries = 3) {
    for (let i = 0; i < maxRetries; i++) {
        try {
            await page.waitForSelector(selector, { visible: true, timeout: 5000 }); // adjust timeout as needed
            const element = await page.$(selector);
            await element.click({ offset: { x: 112, y: 20 } });
            return; // if successful, exit the function
        } catch (error) {
            if ((error.message.includes('Node is detached') || error.message.includes('Node is either not visible')) && i < maxRetries - 1) {
                console.log('Retrying click...');
                await new Promise(resolve => setTimeout(resolve, 1000)); // wait for a second before retrying
            } else {
                throw error; // rethrow the error if it's not about detachment or we've already retried enough times
            }
        }
    }
}
async function SenddatatoCaspio(accessToken, url, data = null, method = 'GET') {
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
        console.error(`Error in SenddatatoCaspio: ${error}`);
        throw error;
    }
}

function sleep(ms) {
    return new Promise(resolve => setTimeout(resolve, ms));
}
(async () => {
    const browser = await puppeteer.launch({
        headless: false,
        slowMo: 1 // delay in milliseconds
    });
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
    //const sheetdata = await fetchSpreadsheetData(client);
    const sheetdata = 1;
    for (let i = 0; i < sheetdata; i++) {
        let Cust_ID = process.argv[2];
        //let Cust_ID = '8659';
        let caspiodata = await caspiomain(Cust_ID, accessToken);
        console.log("Data: ", caspiodata);
        let Firstname = caspiodata.Firstname;
        let Lastname = caspiodata.Lastname;
        let HouseAndStreet = caspiodata.HouseAndStreet;
        let City = caspiodata.City;
        let Canopy = caspiodata.Canopy;
        let Zip = caspiodata.Zip;
        let nyserdafolderid;
        const sendurl = `https://c1abd578.caspio.com/rest/v2/tables/SolarProcess/records?q.where=CustID=${Cust_ID}`;
        //const nyserdaFolderIdURL = `https://c1abd578.caspio.com/rest/v2/tables/FolderID/records?q.select=${nyserdafolderid}&q.where=Cust_ID=${Cust_ID}`
        let Finance = caspiodata.Finance;
        let PurchaseTValue1 = Finance.includes("Lease") ? "Lease" : "Purchase";
        let NYC_LI = caspiodata.NYC_LI;
        let Size = caspiodata.Size;
        let PanelMODEL = caspiodata.Panel;
        let Panel_count = caspiodata.Panel_count;
        let panelbrand;
        let cleanedPanelModel = PanelMODEL.replace(/[\s/-]/g, '').toLowerCase();
        if (cleanedPanelModel.includes('senergy365')) {
            Panel = 'SL40-60MCI-365V';
            panelbrand = "S-Energy";
        } else if (cleanedPanelModel.includes('solaria400')) {
            Panel = 'PowerX-400R';
            panelbrand = "Solaria";
        } else if (cleanedPanelModel.includes('solariax400rt')) {
            Panel = 'PowerX-400R';
            panelbrand = "Solaria";
        }
        else if (cleanedPanelModel.includes('solaria370')) {
            Panel = 'POWERXT-370R-PD';
            panelbrand = "Solaria";
        } else if (cleanedPanelModel.includes('power400')) {
            Panel = 'PowerX-400R';
            panelbrand = "Solaria";
        } else if (cleanedPanelModel.includes('ct370')) {
            Panel = 'CT370HC11-06';
            panelbrand = "Certainteed";
        } else if (cleanedPanelModel.includes('rec365')) {
            Panel = 'REC365NP2 Black';
            panelbrand = "REC Solar";
        } else if (cleanedPanelModel.includes('rec400')) {
            Panel = 'REC400AA';
            panelbrand = "REC Solar";
        } else if (cleanedPanelModel.includes('rec405')) {
            Panel = 'REC405AA';
            panelbrand = "REC Solar";
        } else if (cleanedPanelModel.includes('rec360')) {
            Panel = 'REC360NP2 Black';
            panelbrand = "REC Solar";
        } else if (cleanedPanelModel.includes('seg400')) {
            Panel = 'SEG-400-BMD-HV';
            panelbrand = "SEG Solar";
        }  else {
            Panel = 'TSM-365DE06X.05(II)';
            panelbrand = "Trina";
        }
        let Price = caspiodata.Price;
        let LaborValue1 = Math.round(0.20 * (Price - 3000 - 2000)).toString();
        let BOSValue1 = Math.round(0.25 * (Price - 3000 - 2000)).toString();
        let PermitCValue1 = Math.round(3000).toString();
        let InspectionCValue1 = Math.round(2000).toString();
        let ArrayCValue1 = Math.round(0.35 * (Price - 3000 - 2000)).toString();
        let InvCValue1 = Math.round(0.20 * (Price - 3000 - 2000)).toString();
        let AnnualEValue1 = Math.round(1350 * Size).toString();
        let Phone = caspiodata.Phone;
        let email = caspiodata.email;
        let County = caspiodata.County;
        let Utility = (NYC_LI === "NYC") ? "Con" : "PSEG";
        let NYSUN_Application_type = caspiodata.NYSUN_Application_type;
        let finance = caspiodata.Finance;
        let Affordable;
        let GIGNY;
        switch (NYSUN_Application_type) {
            case "Affordable Only":
                Affordable = true;
                GIGNY = false;
                submission_requiments = '1.Photos 2.Signed Application 3.Shading Report 4.Site Map 5.Electrical diagram 6.Affordable letter';
                break;
            case "Coned Incentive Only":
                Affordable = false;
                GIGNY = false;
                submission_requiments = '1.Photos 2.Signed Application 3.Shading Report 4.Site Map 5.Electrical diagram';
                break;
            case "NYSERDA Loan":
                Affordable = false;
                GIGNY = true;
                submission_requiments = '1.Photos 2.Signed Application 3.Shading Report 4.Site Map 5.Electrical diagram';
                break;
            case "Affordable+NYSERDA Loan":
                Affordable = true;
                GIGNY = true;
                submission_requiments = '1.Photos 2.Signed Application 3.Shading Report 4.Site Map 5.Electrical diagram 6.Affordable letter';
                break;
            case "Coned Incentive+NYSERDA Loan":
                Affordable = false;
                GIGNY = true;
                submission_requiments = '1.Photos 2.Signed Application 3.Shading Report 4.Site Map 5.Electrical diagram';
                break;
            default:
                console.log('No matching NYSUN Application type');
        }
        let datarow = caspiodata.datarow;
        let Sector = "Residential";
        let Arraytype;
        if (Canopy === true) {
            Arraytype = "Fixed Rooftop Canopy";
        } else {
            Arraytype = "Fixed Roof Mount";
        }

        console.log('sendurl:', sendurl);
        // rest of your code
        // delay between each run
        await sleep(5000); // adjust the delay time as needed
        const page = await browser.newPage();
        const timeout = 30000;
        page.setDefaultTimeout(timeout);
        {
            const targetPage = page;
            await targetPage.setViewport({
                width: 1139,
                height: 937
            })
        }
        {
            const targetPage = page;
            const promises = [];
            promises.push(targetPage.waitForNavigation());
            await targetPage.goto('https://portal.nyserda.ny.gov/PortalLoginPage');
            await Promise.all(promises);
        }
        {
            const targetPage = page;
            await scrollIntoViewIfNeeded([
                [
                    '#loginPage\\:loginForm\\:usernameId'
                ],
                [
                    'xpath///*[@id="loginPage:loginForm:usernameId"]'
                ],
                [
                    'pierce/#loginPage\\:loginForm\\:usernameId'
                ],
                [
                    'aria/Username'
                ]
            ], targetPage, timeout);
            const element = await waitForSelectors([
                [
                    '#loginPage\\:loginForm\\:usernameId'
                ],
                [
                    'xpath///*[@id="loginPage:loginForm:usernameId"]'
                ],
                [
                    'pierce/#loginPage\\:loginForm\\:usernameId'
                ],
                [
                    'aria/Username'
                ]
            ], targetPage, { timeout, visible: true });
            await element.click({
                offset: {
                    x: 21.5,
                    y: 16.515625,
                },
            });
        }
        {
            const targetPage = page;
            await scrollIntoViewIfNeeded([
                [
                    '#loginPage\\:loginForm\\:usernameId'
                ],
                [
                    'xpath///*[@id="loginPage:loginForm:usernameId"]'
                ],
                [
                    'pierce/#loginPage\\:loginForm\\:usernameId'
                ],
                [
                    'aria/Username'
                ]
            ], targetPage, timeout);
            const element = await waitForSelectors([
                [
                    '#loginPage\\:loginForm\\:usernameId'
                ],
                [
                    'xpath///*[@id="loginPage:loginForm:usernameId"]'
                ],
                [
                    'pierce/#loginPage\\:loginForm\\:usernameId'
                ],
                [
                    'aria/Username'
                ]
            ], targetPage, { timeout, visible: true });
            const inputType = await element.evaluate(el => el.type);
            if (inputType === 'select-one') {
                await changeSelectElement(element, 'teamdesign@patriotenergysolution.com.nyserda')
            } else if ([
                'textarea',
                'text',
                'url',
                'tel',
                'search',
                'password',
                'number',
                'email'
            ].includes(inputType)) {
                await typeIntoElement(element, 'teamdesign@patriotenergysolution.com.nyserda');
            } else {
                await changeElementValue(element, 'teamdesign@patriotenergysolution.com.nyserda');
            }
        }
        {
            const targetPage = page;
            await scrollIntoViewIfNeeded([
                [
                    '#loginPage\\:loginForm\\:passwordId'
                ],
                [
                    'xpath///*[@id="loginPage:loginForm:passwordId"]'
                ],
                [
                    'pierce/#loginPage\\:loginForm\\:passwordId'
                ],
                [
                    'aria/Password'
                ]
            ], targetPage, timeout);
            const element = await waitForSelectors([
                [
                    '#loginPage\\:loginForm\\:passwordId'
                ],
                [
                    'xpath///*[@id="loginPage:loginForm:passwordId"]'
                ],
                [
                    'pierce/#loginPage\\:loginForm\\:passwordId'
                ],
                [
                    'aria/Password'
                ]
            ], targetPage, { timeout, visible: true });
            await element.click({
                offset: {
                    x: 41.5,
                    y: 18.515625,
                },
            });
        }
        {
            const targetPage = page;
            await scrollIntoViewIfNeeded([
                [
                    '#loginPage\\:loginForm\\:passwordId'
                ],
                [
                    'xpath///*[@id="loginPage:loginForm:passwordId"]'
                ],
                [
                    'pierce/#loginPage\\:loginForm\\:passwordId'
                ],
                [
                    'aria/Password'
                ]
            ], targetPage, timeout);
            const element = await waitForSelectors([
                [
                    '#loginPage\\:loginForm\\:passwordId'
                ],
                [
                    'xpath///*[@id="loginPage:loginForm:passwordId"]'
                ],
                [
                    'pierce/#loginPage\\:loginForm\\:passwordId'
                ],
                [
                    'aria/Password'
                ]
            ], targetPage, { timeout, visible: true });
            const inputType = await element.evaluate(el => el.type);
            if (inputType === 'select-one') {
                await changeSelectElement(element, 'Patriot1265123456!')
            } else if ([
                'textarea',
                'text',
                'url',
                'tel',
                'search',
                'password',
                'number',
                'email'
            ].includes(inputType)) {
                await typeIntoElement(element, 'Patriot1265123456!');
            } else {
                await changeElementValue(element, 'Patriot1265123456!');
            }
        }
        {
            const targetPage = page;
            const promises = [];
            promises.push(targetPage.waitForNavigation());
            await scrollIntoViewIfNeeded([
                [
                    '#loginPage\\:loginForm\\:loginButton'
                ],
                [
                    'xpath///*[@id="loginPage:loginForm:loginButton"]'
                ],
                [
                    'pierce/#loginPage\\:loginForm\\:loginButton'
                ],
                [
                    'aria/Log In'
                ],
                [
                    'text/Log In'
                ]
            ], targetPage, timeout);
            const element = await waitForSelectors([
                [
                    '#loginPage\\:loginForm\\:loginButton'
                ],
                [
                    'xpath///*[@id="loginPage:loginForm:loginButton"]'
                ],
                [
                    'pierce/#loginPage\\:loginForm\\:loginButton'
                ],
                [
                    'aria/Log In'
                ],
                [
                    'text/Log In'
                ]
            ], targetPage, { timeout, visible: true });
            await element.click({
                offset: {
                    x: 57.5,
                    y: 12.515625,
                },
            });
            await Promise.all(promises);
            
        }
    {
        const targetPage = page;
        const promises = [];
        const startWaitingForEvents = () => {
            promises.push(targetPage.waitForNavigation());
        }
        await puppeteer.Locator.race([
            targetPage.locator('#\\30 1r36000000FOKG_Tab > a'),
            targetPage.locator('::-p-xpath(//*[@id=\\"01r36000000FOKG_Tab\\"]/a)'),
            targetPage.locator(':scope >>> #\\30 1r36000000FOKG_Tab > a'),
            targetPage.locator('::-p-aria(Submit a New Application)'),
            targetPage.locator('::-p-text(Submit a New)')
        ])
            .setTimeout(timeout)
            .on('action', () => startWaitingForEvents())
            .click({
              offset: {
                x: 80.90625,
                y: 11,
              },
            });
        await Promise.all(promises);
    }
		await page.waitForTimeout(2000); // wait for 1 seconds
    {
        const targetPage = page;
        const promises = [];
        const startWaitingForEvents = () => {
            promises.push(targetPage.waitForNavigation());
        }
        await puppeteer.Locator.race([
            targetPage.locator('#j_id0\\:j_id2\\:j_id23\\:0\\:j_id24 span.copy'),
            targetPage.locator('::-p-xpath(//*[@id=\\"j_id0:j_id2:j_id23:0:j_id29\\"]/a/div[1]/span[2])'),
            targetPage.locator(':scope >>> #j_id0\\:j_id2\\:j_id23\\:0\\:j_id24 span.copy')
        ])
            .setTimeout(timeout)
            .on('action', () => startWaitingForEvents())
            .click({
              offset: {
                x: 293,
                y: 0.21875,
              },
            });
        await Promise.all(promises);
    }
		    {
        const targetPage = page;
        await puppeteer.Locator.race([
            targetPage.locator('div:nth-of-type(3) i'),
            targetPage.locator('::-p-xpath(//*[@id=\\"j_id0:j_id1:j_id14:j_id15:frm:j_id144\\"]/fieldset/div[3]/label/div/i)'),
            targetPage.locator(':scope >>> div:nth-of-type(3) i')
        ])
            .setTimeout(timeout)
            .click({
              offset: {
                x: 13.5,
                y: 8.40625,
              },
            });
			//await Promise.all(promises);
    }
/*
        {
            const targetPage = page;
            await scrollIntoViewIfNeeded([
                'div:nth-of-type(3) i',
                'xpath///*[@id="j_id0:j_id1:j_id14:j_id15:frm:j_id140"]/fieldset/div[3]/label/div/i',
                'pierce/div:nth-of-type(3) i'
            ], targetPage, timeout);
            const element = await waitForSelectors([
                'div:nth-of-type(3) i',
                'xpath///*[@id="j_id0:j_id1:j_id14:j_id15:frm:j_id140"]/fieldset/div[3]/label/div/i',
                'pierce/div:nth-of-type(3) i'
            ], targetPage, { timeout, visible: true });
            await element.click({
                offset: {
                    x: 10,
                    y: 8.40625,
                },
            });
            await page.waitForTimeout(1000); // wait for 1 seconds
            //await page.waitForNavigation({ waitUntil: 'networkidle0' });

        }
		*/
		            await page.waitForTimeout(1000); // wait for 1 seconds
		{
        const targetPage = page;
        const promises = [];
        const startWaitingForEvents = () => {
            promises.push(targetPage.waitForNavigation());
        }
        await puppeteer.Locator.race([
            targetPage.locator('a.btn-primary'),
            targetPage.locator('::-p-xpath(//*[@id=\\"j_id0:j_id1:j_id14:j_id15:frm\\"]/footer/div/a[3])'),
            targetPage.locator(':scope >>> a.btn-primary'),
            targetPage.locator('::-p-aria(CONTINUE)')
        ])
            .setTimeout(timeout)
            .on('action', () => startWaitingForEvents())
            .click({
              offset: {
                x: 98.84375,
                y: 18.40625,
              },
            });
        await Promise.all(promises);
    }
            await page.waitForTimeout(1000); // wait for 1 seconds
/*
        {
            const targetPage = page;
            const promises = [];
            promises.push(targetPage.waitForNavigation());
            await scrollIntoViewIfNeeded([
                'a.btn-primary',
                'xpath///*[@id="j_id0:j_id1:j_id14:j_id15:frm"]/footer/div/a[3]',
                'pierce/a.btn-primary',
                'aria/CONTINUE'
            ], targetPage, timeout);
            const element = await waitForSelectors([
                'a.btn-primary',
                'xpath///*[@id="j_id0:j_id1:j_id14:j_id15:frm"]/footer/div/a[3]',
                'pierce/a.btn-primary',
                'aria/CONTINUE'
            ], targetPage, { timeout, visible: true });
            await element.click({
                offset: {
                    x: 34.34375,
                    y: 19.40625,
                },
            });
            await Promise.all(promises);
            await page.waitForTimeout(1000); // wait for 1 seconds
            //await page.waitForNavigation({ waitUntil: 'networkidle0' });
        }
				*/
        {
            const targetPage = page;
            await scrollIntoViewIfNeeded([
                'div.application a',
                'xpath///*[@id="j_id0:j_id2:j_id85"]/a',
                'pierce/div.application a',
                'aria/ ADD A SITE',
                'text/Add a Site'
            ], targetPage, timeout);
            const element = await waitForSelectors([
                'div.application a',
                'xpath///*[@id="j_id0:j_id2:j_id85"]/a',
                'pierce/div.application a',
                'aria/ ADD A SITE',
                'text/Add a Site'
            ], targetPage, { timeout, visible: true });
            await element.click({
                offset: {
                    x: 39,
                    y: 4.515625,
                },
            });
            await page.waitForTimeout(2000); // wait for 1 seconds
        }
        {
            const targetPage = page;
            await scrollIntoViewIfNeeded([
                '#j_id0\\:j_id2\\:frm\\:address1',
                'xpath///*[@id="j_id0:j_id2:frm:address1"]',
                'pierce/#j_id0\\:j_id2\\:frm\\:address1',
                'aria/Address 1'
            ], targetPage, timeout);
            const element = await waitForSelectors([
                '#j_id0\\:j_id2\\:frm\\:address1',
                'xpath///*[@id="j_id0:j_id2:frm:address1"]',
                'pierce/#j_id0\\:j_id2\\:frm\\:address1',
                'aria/Address 1'
            ], targetPage, { timeout, visible: true });
            const inputType = await element.evaluate(el => el.type);

            // Get value from Google Spreadsheet from chatgpt
            const addressValue = HouseAndStreet; //////////////////////////////////////////////////////////////datarow Address input

            if (inputType === 'select-one') {
                await changeSelectElement(element, addressValue)
            } else if ([
                'textarea',
                'text',
                'url',
                'tel',
                'search',
                'password',
                'number',
                'email'
            ].includes(inputType)) {
                await typeIntoElement(element, addressValue);
            } else {
                await changeElementValue(element, addressValue);
            }
            await page.waitForTimeout(1000); // wait for 1 seconds
        }
        {
            const targetPage = page;
            await scrollIntoViewIfNeeded([
                [
                    '#j_id0\\:j_id2\\:frm\\:city'
                ],
                [
                    'xpath///*[@id="j_id0:j_id2:frm:city"]'
                ],
                [
                    'pierce/#j_id0\\:j_id2\\:frm\\:city'
                ],
                [
                    'aria/City'
                ]
            ], targetPage, timeout);
            const element = await waitForSelectors([
                [
                    '#j_id0\\:j_id2\\:frm\\:city'
                ],
                [
                    'xpath///*[@id="j_id0:j_id2:frm:city"]'
                ],
                [
                    'pierce/#j_id0\\:j_id2\\:frm\\:city'
                ],
                [
                    'aria/City'
                ]
            ], targetPage, { timeout, visible: true });
            await element.click({
                offset: {
                    x: 76.421875,
                    y: 14,
                },
            });
            await page.waitForTimeout(1000); // wait for 1 seconds
        }
        {
            const targetPage = page;
            await scrollIntoViewIfNeeded([
                [
                    '#j_id0\\:j_id2\\:frm\\:city'
                ],
                [
                    'xpath///*[@id="j_id0:j_id2:frm:city"]'
                ],
                [
                    'pierce/#j_id0\\:j_id2\\:frm\\:city'
                ],
                [
                    'aria/City'
                ]
            ], targetPage, timeout);
            const element = await waitForSelectors([
                [
                    '#j_id0\\:j_id2\\:frm\\:city'
                ],
                [
                    'xpath///*[@id="j_id0:j_id2:frm:city"]'
                ],
                [
                    'pierce/#j_id0\\:j_id2\\:frm\\:city'
                ],
                [
                    'aria/City'
                ]
            ], targetPage, { timeout, visible: true });
            const inputType = await element.evaluate(el => el.type);
            // Get value from Google Spreadsheet from chatgpt
            const CityValue = City; ///////////////////////////////////////////////////////////////datarow city input
            if (inputType === 'select-one') {
                await changeSelectElement(element, CityValue)
            } else if ([
                'textarea',
                'text',
                'url',
                'tel',
                'search',
                'password',
                'number',
                'email'
            ].includes(inputType)) {
                await typeIntoElement(element, CityValue);
            } else {
                await changeElementValue(element, CityValue);
            }
            await page.waitForTimeout(1000); // wait for 1 seconds
        }
        {
            const targetPage = page;
            await scrollIntoViewIfNeeded([
                [
                    '#j_id0\\:j_id2\\:frm\\:zip'
                ],
                [
                    'xpath///*[@id="j_id0:j_id2:frm:zip"]'
                ],
                [
                    'pierce/#j_id0\\:j_id2\\:frm\\:zip'
                ],
                [
                    'aria/Zip'
                ]
            ], targetPage, timeout);
            const element = await waitForSelectors([
                [
                    '#j_id0\\:j_id2\\:frm\\:zip'
                ],
                [
                    'xpath///*[@id="j_id0:j_id2:frm:zip"]'
                ],
                [
                    'pierce/#j_id0\\:j_id2\\:frm\\:zip'
                ],
                [
                    'aria/Zip'
                ]
            ], targetPage, { timeout, visible: true });
            await element.click({
                offset: {
                    x: 61.421875,
                    y: 21,
                },
            });
            await page.waitForTimeout(1000); // wait for 1 seconds
        }
        {
            const targetPage = page;
            await scrollIntoViewIfNeeded([
                [
                    '#j_id0\\:j_id2\\:frm\\:zip'
                ],
                [
                    'xpath///*[@id="j_id0:j_id2:frm:zip"]'
                ],
                [
                    'pierce/#j_id0\\:j_id2\\:frm\\:zip'
                ],
                [
                    'aria/Zip'
                ]
            ], targetPage, timeout);
            const element = await waitForSelectors([
                [
                    '#j_id0\\:j_id2\\:frm\\:zip'
                ],
                [
                    'xpath///*[@id="j_id0:j_id2:frm:zip"]'
                ],
                [
                    'pierce/#j_id0\\:j_id2\\:frm\\:zip'
                ],
                [
                    'aria/Zip'
                ]
            ], targetPage, { timeout, visible: true });
            const inputType = await element.evaluate(el => el.type);
            // Get value from Google Spreadsheet from chatgpt
            const ZipValue = Zip;
            if (inputType === 'select-one') {
                await changeSelectElement(element, ZipValue)
            } else if ([
                'textarea',
                'text',
                'url',
                'tel',
                'search',
                'password',
                'number',
                'email'
            ].includes(inputType)) {
                await typeIntoElement(element, ZipValue);
            } else {
                await changeElementValue(element, ZipValue);
            }
            await page.waitForTimeout(1000); // wait for 1 seconds
        }
        {
            const targetPage = page;
            await scrollIntoViewIfNeeded([
                [
                    '#select2-j_id0j_id2frmcounty-container'
                ],
                [
                    'xpath///*[@id="select2-j_id0j_id2frmcounty-container"]'
                ],
                [
                    'pierce/#select2-j_id0j_id2frmcounty-container'
                ]
            ], targetPage, timeout);
            const element = await waitForSelectors([
                [
                    '#select2-j_id0j_id2frmcounty-container'
                ],
                [
                    'xpath///*[@id="select2-j_id0j_id2frmcounty-container"]'
                ],
                [
                    'pierce/#select2-j_id0j_id2frmcounty-container'
                ]
            ], targetPage, { timeout, visible: true });
            await element.click({
                offset: {
                    x: 82.421875,
                    y: 16,
                },
            });
            await page.waitForTimeout(1000); // wait for 1 seconds
        }
        {
            const targetPage = page;
            await scrollIntoViewIfNeeded([
                [
                    'span.select2-container input'
                ],
                [
                    'xpath//html/body/span[2]/span/span[1]/input'
                ],
                [
                    'pierce/span.select2-container input'
                ]
            ], targetPage, timeout);
            const element = await waitForSelectors([
                [
                    'span.select2-container input'
                ],
                [
                    'xpath//html/body/span[2]/span/span[1]/input'
                ],
                [
                    'pierce/span.select2-container input'
                ]
            ], targetPage, { timeout, visible: true });
            const inputType = await element.evaluate(el => el.type);
            // Get value from Google Spreadsheet from chatgpt
            const CountyValue = County; ////////////////////////////////////////////datarow County input
            if (inputType === 'select-one') {
                await changeSelectElement(element, CountyValue)
            } else if ([
                'textarea',
                'text',
                'url',
                'tel',
                'search',
                'password',
                'number',
                'email'
            ].includes(inputType)) {
                await typeIntoElement(element, CountyValue);
            } else {
                await changeElementValue(element, CountyValue);
            }
        }
        {
            const targetPage = page;
            await targetPage.keyboard.down('Enter');
        }
        {
            const targetPage = page;
            await targetPage.keyboard.up('Enter');
        }
        {
            const targetPage = page;
            await scrollIntoViewIfNeeded([
                [
                    '#select2-j_id0j_id2frmutilityCompany-container'
                ],
                [
                    'xpath///*[@id="select2-j_id0j_id2frmutilityCompany-container"]'
                ],
                [
                    'pierce/#select2-j_id0j_id2frmutilityCompany-container'
                ]
            ], targetPage, timeout);
            const element = await waitForSelectors([
                [
                    '#select2-j_id0j_id2frmutilityCompany-container'
                ],
                [
                    'xpath///*[@id="select2-j_id0j_id2frmutilityCompany-container"]'
                ],
                [
                    'pierce/#select2-j_id0j_id2frmutilityCompany-container'
                ]
            ], targetPage, { timeout, visible: true });
            await element.click({
                offset: {
                    x: 84.421875,
                    y: 16,
                },
            });
            await page.waitForTimeout(1000); // wait for 1 seconds
        }
        {
            const targetPage = page;
            await scrollIntoViewIfNeeded([
                [
                    'span.select2-container input'
                ],
                [
                    'xpath//html/body/span[2]/span/span[1]/input'
                ],
                [
                    'pierce/span.select2-container input'
                ]
            ], targetPage, timeout);
            const element = await waitForSelectors([
                [
                    'span.select2-container input'
                ],
                [
                    'xpath//html/body/span[2]/span/span[1]/input'
                ],
                [
                    'pierce/span.select2-container input'
                ]
            ], targetPage, { timeout, visible: true });
            const inputType = await element.evaluate(el => el.type);
            // Get value from Google Spreadsheet from chatgpt
            const UtilityValue = Utility;        ///////////////////////////////////////datarow Utility company
            if (inputType === 'select-one') {
                await changeSelectElement(element, UtilityValue)
            } else if ([
                'textarea',
                'text',
                'url',
                'tel',
                'search',
                'password',
                'number',
                'email'
            ].includes(inputType)) {
                await typeIntoElement(element, UtilityValue);
            } else {
                await changeElementValue(element, UtilityValue);
            }
        }
        {
            const targetPage = page;
            await targetPage.keyboard.down('Enter');
        }
        {
            const targetPage = page;
            await targetPage.keyboard.up('Enter');
        }
        {
            const targetPage = page;
            await scrollIntoViewIfNeeded([
                [
                    '#select2-j_id0j_id2frmutilitySector-container'
                ],
                [
                    'xpath///*[@id="select2-j_id0j_id2frmutilitySector-container"]'
                ],
                [
                    'pierce/#select2-j_id0j_id2frmutilitySector-container'
                ],
                [
                    'aria/--None--[role="textbox"]'
                ]
            ], targetPage, timeout);
            const element = await waitForSelectors([
                [
                    '#select2-j_id0j_id2frmutilitySector-container'
                ],
                [
                    'xpath///*[@id="select2-j_id0j_id2frmutilitySector-container"]'
                ],
                [
                    'pierce/#select2-j_id0j_id2frmutilitySector-container'
                ],
                [
                    'aria/--None--[role="textbox"]'
                ]
            ], targetPage, { timeout, visible: true });
            await element.click({
                offset: {
                    x: 80.421875,
                    y: 11,
                },
            });
        }
        {
            const targetPage = page;
            await scrollIntoViewIfNeeded([
                [
                    'span.select2-container input'
                ],
                [
                    'xpath//html/body/span[2]/span/span[1]/input'
                ],
                [
                    'pierce/span.select2-container input'
                ]
            ], targetPage, timeout);
            const element = await waitForSelectors([
                [
                    'span.select2-container input'
                ],
                [
                    'xpath//html/body/span[2]/span/span[1]/input'
                ],
                [
                    'pierce/span.select2-container input'
                ]
            ], targetPage, { timeout, visible: true });
            const inputType = await element.evaluate(el => el.type);
            // Get value from Google Spreadsheet from chatgpt
            const SectorValue = Sector; //////////////////////////////////////////////datarow sector 
            if (inputType === 'select-one') {
                await changeSelectElement(element, SectorValue)
            } else if ([
                'textarea',
                'text',
                'url',
                'tel',
                'search',
                'password',
                'number',
                'email'
            ].includes(inputType)) {
                await typeIntoElement(element, SectorValue);
            } else {
                await changeElementValue(element, SectorValue);
            }
        }
        {
            const targetPage = page;
            await targetPage.keyboard.down('Enter');
        }
        {
            const targetPage = page;
            await targetPage.keyboard.up('Enter');
        }
        /////checkbox-checkbox-checkbox////////////////////////////////////////////////////////////////////////////////////
        // Assuming 'Affordable' is the column header for the checkbox column in your spreadsheet
        if (Affordable === true) {
            const targetPage = page;
            await scrollIntoViewIfNeeded([
                [
                    '#j_id0\\:j_id2\\:frm\\:Affordable_Solar'
                ],
                [
                    'xpath///*[@id="j_id0:j_id2:frm:Affordable_Solar"]'
                ],
                [
                    'pierce/#j_id0\\:j_id2\\:frm\\:Affordable_Solar'
                ]
            ], targetPage, timeout);
            const element = await waitForSelectors([
                [
                    '#j_id0\\:j_id2\\:frm\\:Affordable_Solar'
                ],
                [
                    'xpath///*[@id="j_id0:j_id2:frm:Affordable_Solar"]'
                ],
                [
                    'pierce/#j_id0\\:j_id2\\:frm\\:Affordable_Solar'
                ]
            ], targetPage, { timeout, visible: true });
            await element.click({
                offset: {
                    x: 213.921875,
                    y: 11,
                },
            });
            await page.waitForTimeout(5000); // wait for 1 seconds
        }
        /////checkbox-checkbox-checkbox ended before this line////////////////////////////////////////////////////////////////////////////////////
        {
            const targetPage = page;
            const promises = [];
            promises.push(targetPage.waitForNavigation());
            await scrollIntoViewIfNeeded([
                [
                    '#saveBtn'
                ],
                [
                    'xpath///*[@id="saveBtn"]'
                ],
                [
                    'pierce/#saveBtn'
                ],
                [
                    'aria/SAVE'
                ],
                [
                    'text/Save'
                ]
            ], targetPage, timeout);
            const element = await waitForSelectors([
                [
                    '#saveBtn'
                ],
                [
                    'xpath///*[@id="saveBtn"]'
                ],
                [
                    'pierce/#saveBtn'
                ],
                [
                    'aria/SAVE'
                ],
                [
                    'text/Save'
                ]
            ], targetPage, { timeout, visible: true });
            await element.click({
                offset: {
                    x: 46,
                    y: 8,
                },
            });
            await Promise.all(promises);
			await page.waitForTimeout(3000); // wait for 1 seconds
        }
        {
            const targetPage = page;
            const promises = [];
            promises.push(targetPage.waitForNavigation());
            await scrollIntoViewIfNeeded([
                'a.btn-primary',
                'xpath///*[@id="j_id0:j_id2:frm"]/footer/div/a[3]',
                'pierce/a.btn-primary',
                'aria/CONTINUE'
            ], targetPage, timeout);
            const element = await waitForSelectors([
                'a.btn-primary',
                'xpath///*[@id="j_id0:j_id2:frm"]/footer/div/a[3]',
                'pierce/a.btn-primary',
                'aria/CONTINUE'
            ], targetPage, { timeout, visible: true });
            await element.click({
                delay: 2000,
                offset: {
                    x: 48.34375,
                    y: 26.515625,
                },
            });
            await Promise.all(promises);
            //await page.waitForNavigation({ waitUntil: 'networkidle0' });
        }
        {
    const targetPage = page;
    await puppeteer.Locator.race([
        targetPage.locator('div.page div.content > a'),
        targetPage.locator('::-p-xpath(//*[@id=\\"bodyTable\\"]/tbody/tr/td/div[2]/div[3]/div[2]/a)'),
        targetPage.locator(':scope >>> div.page div.content > a'),
        targetPage.locator('::-p-aria( ADD A CONTACT)')
    ])
        .setTimeout(timeout)
        .click({
          offset: {
            x: 78.640625,
            y: 11.015625,
          },
        });
}
            await page.waitForTimeout(5000); // wait for 1 seconds
      
        {
            const targetPage = page;
            await scrollIntoViewIfNeeded([
                [
                    '#j_id0\\:frm\\:firstName'
                ],
                [
                    'xpath///*[@id="j_id0:frm:firstName"]'
                ],
                [
                    'pierce/#j_id0\\:frm\\:firstName'
                ],
                [
                    'aria/First Name *'
                ]
            ], targetPage, timeout);
            const element = await waitForSelectors([
                [
                    '#j_id0\\:frm\\:firstName'
                ],
                [
                    'xpath///*[@id="j_id0:frm:firstName"]'
                ],
                [
                    'pierce/#j_id0\\:frm\\:firstName'
                ],
                [
                    'aria/First Name *'
                ]
            ], targetPage, { timeout, visible: true });
            await element.click({
                offset: {
                    x: 128.5,
                    y: 22,
                },
            });
            await page.waitForTimeout(1000); // wait for 1 seconds
        }
        {
            const targetPage = page;
            await scrollIntoViewIfNeeded([
                [
                    '#j_id0\\:frm\\:firstName'
                ],
                [
                    'xpath///*[@id="j_id0:frm:firstName"]'
                ],
                [
                    'pierce/#j_id0\\:frm\\:firstName'
                ],
                [
                    'aria/First Name *'
                ]
            ], targetPage, timeout);
            const element = await waitForSelectors([
                [
                    '#j_id0\\:frm\\:firstName'
                ],
                [
                    'xpath///*[@id="j_id0:frm:firstName"]'
                ],
                [
                    'pierce/#j_id0\\:frm\\:firstName'
                ],
                [
                    'aria/First Name *'
                ]
            ], targetPage, { timeout, visible: true });
            const inputType = await element.evaluate(el => el.type);
            // Get value from Google Spreadsheet from chatgpt
            const FirstnameValue = Firstname; ////////////////////////datarow Firstname enter//////////////////////////////////////////////
            if (inputType === 'select-one') {
                await changeSelectElement(element, FirstnameValue)
            } else if ([
                'textarea',
                'text',
                'url',
                'tel',
                'search',
                'password',
                'number',
                'email'
            ].includes(inputType)) {
                await typeIntoElement(element, FirstnameValue);
            } else {
                await changeElementValue(element, FirstnameValue);
            }
        }
        {
            const targetPage = page;
            await scrollIntoViewIfNeeded([
                [
                    '#j_id0\\:frm\\:lastName'
                ],
                [
                    'xpath///*[@id="j_id0:frm:lastName"]'
                ],
                [
                    'pierce/#j_id0\\:frm\\:lastName'
                ],
                [
                    'aria/Last Name *'
                ]
            ], targetPage, timeout);
            const element = await waitForSelectors([
                [
                    '#j_id0\\:frm\\:lastName'
                ],
                [
                    'xpath///*[@id="j_id0:frm:lastName"]'
                ],
                [
                    'pierce/#j_id0\\:frm\\:lastName'
                ],
                [
                    'aria/Last Name *'
                ]
            ], targetPage, { timeout, visible: true });
            await element.click({
                offset: {
                    x: 48.5,
                    y: 15,
                },
            });
            await page.waitForTimeout(1000); // wait for 1 seconds
        }
        {
            const targetPage = page;
            await scrollIntoViewIfNeeded([
                [
                    '#j_id0\\:frm\\:lastName'
                ],
                [
                    'xpath///*[@id="j_id0:frm:lastName"]'
                ],
                [
                    'pierce/#j_id0\\:frm\\:lastName'
                ],
                [
                    'aria/Last Name *'
                ]
            ], targetPage, timeout);
            const element = await waitForSelectors([
                [
                    '#j_id0\\:frm\\:lastName'
                ],
                [
                    'xpath///*[@id="j_id0:frm:lastName"]'
                ],
                [
                    'pierce/#j_id0\\:frm\\:lastName'
                ],
                [
                    'aria/Last Name *'
                ]
            ], targetPage, { timeout, visible: true });
            const inputType = await element.evaluate(el => el.type);
            // Get value from Google Spreadsheet from chatgpt
            const LastnameValue = Lastname; /////////////////////////////////////////datarow Lastname enter///////////////////////////////////////////
            if (inputType === 'select-one') {
                await changeSelectElement(element, LastnameValue)
            } else if ([
                'textarea',
                'text',
                'url',
                'tel',
                'search',
                'password',
                'number',
                'email'
            ].includes(inputType)) {
                await typeIntoElement(element, LastnameValue);
            } else {
                await changeElementValue(element, LastnameValue);
            }
            await page.waitForTimeout(1000); // wait for 1 seconds
        }
        {
            const targetPage = page;
            await scrollIntoViewIfNeeded([
                [
                    '#j_id0\\:frm\\:email'
                ],
                [
                    'xpath///*[@id="j_id0:frm:email"]'
                ],
                [
                    'pierce/#j_id0\\:frm\\:email'
                ],
                [
                    'aria/Email *'
                ]
            ], targetPage, timeout);
            const element = await waitForSelectors([
                [
                    '#j_id0\\:frm\\:email'
                ],
                [
                    'xpath///*[@id="j_id0:frm:email"]'
                ],
                [
                    'pierce/#j_id0\\:frm\\:email'
                ],
                [
                    'aria/Email *'
                ]
            ], targetPage, { timeout, visible: true });
            await element.click({
                offset: {
                    x: 30.5,
                    y: 18,
                },
            });
        }
        {
            const targetPage = page;
            await scrollIntoViewIfNeeded([
                [
                    '#j_id0\\:frm\\:email'
                ],
                [
                    'xpath///*[@id="j_id0:frm:email"]'
                ],
                [
                    'pierce/#j_id0\\:frm\\:email'
                ],
                [
                    'aria/Email *'
                ]
            ], targetPage, timeout);
            const element = await waitForSelectors([
                [
                    '#j_id0\\:frm\\:email'
                ],
                [
                    'xpath///*[@id="j_id0:frm:email"]'
                ],
                [
                    'pierce/#j_id0\\:frm\\:email'
                ],
                [
                    'aria/Email *'
                ]
            ], targetPage, { timeout, visible: true });
            const inputType = await element.evaluate(el => el.type);
            // Get value from Google Spreadsheet from chatgpt
            const EmailValue = email; ///////////////////////////////////////////////////datarow Email enter/////////////////////////////////////
            if (inputType === 'select-one') {
                await changeSelectElement(element, EmailValue)
            } else if ([
                'textarea',
                'text',
                'url',
                'tel',
                'search',
                'password',
                'number',
                'email'
            ].includes(inputType)) {
                await typeIntoElement(element, EmailValue);
            } else {
                await changeElementValue(element, EmailValue);
            }
        }
        {
            const targetPage = page;
            await targetPage.keyboard.up('2');
        }
        {
            const targetPage = page;
            await scrollIntoViewIfNeeded([
                [
                    '#j_id0\\:frm\\:phone'
                ],
                [
                    'xpath///*[@id="j_id0:frm:phone"]'
                ],
                [
                    'pierce/#j_id0\\:frm\\:phone'
                ],
                [
                    'aria/Phone *'
                ]
            ], targetPage, timeout);
            const element = await waitForSelectors([
                [
                    '#j_id0\\:frm\\:phone'
                ],
                [
                    'xpath///*[@id="j_id0:frm:phone"]'
                ],
                [
                    'pierce/#j_id0\\:frm\\:phone'
                ],
                [
                    'aria/Phone *'
                ]
            ], targetPage, { timeout, visible: true });
            await element.click({
                offset: {
                    x: 41.5,
                    y: 26,
                },
            });
        }
        {
            const targetPage = page;
            await scrollIntoViewIfNeeded([
                [
                    '#j_id0\\:frm\\:phone'
                ],
                [
                    'xpath///*[@id="j_id0:frm:phone"]'
                ],
                [
                    'pierce/#j_id0\\:frm\\:phone'
                ],
                [
                    'aria/Phone *'
                ]
            ], targetPage, timeout);
            const element = await waitForSelectors([
                [
                    '#j_id0\\:frm\\:phone'
                ],
                [
                    'xpath///*[@id="j_id0:frm:phone"]'
                ],
                [
                    'pierce/#j_id0\\:frm\\:phone'
                ],
                [
                    'aria/Phone *'
                ]
            ], targetPage, { timeout, visible: true });
            const inputType = await element.evaluate(el => el.type);
            // Get value from Google Spreadsheet from chatgpt
            const PhoneValue = Phone.replace(/[\s/()-]/g, '').trim();
            if (inputType === 'select-one') {
                await changeSelectElement(element, PhoneValue)
            } else if ([
                'textarea',
                'text',
                'url',
                'tel',
                'search',
                'password',
                'number',
                'email'
            ].includes(inputType)) {
                await typeIntoElement(element, PhoneValue);
            } else {
                await changeElementValue(element, PhoneValue);
            }
            await page.waitForTimeout(1000); // wait for 1 seconds
        }
        {
            const targetPage = page;
            await scrollIntoViewIfNeeded([
                [
                    '#j_id0\\:frm\\:company'
                ],
                [
                    'xpath///*[@id="j_id0:frm:company"]'
                ],
                [
                    'pierce/#j_id0\\:frm\\:company'
                ],
                [
                    'aria/Company Name *'
                ]
            ], targetPage, timeout);
            const element = await waitForSelectors([
                [
                    '#j_id0\\:frm\\:company'
                ],
                [
                    'xpath///*[@id="j_id0:frm:company"]'
                ],
                [
                    'pierce/#j_id0\\:frm\\:company'
                ],
                [
                    'aria/Company Name *'
                ]
            ], targetPage, { timeout, visible: true });
            await element.click({
                offset: {
                    x: 57.5,
                    y: 19,
                },
            });
            await page.waitForTimeout(1000); // wait for 1 seconds
        }
        {
            const targetPage = page;
            await scrollIntoViewIfNeeded([
                [
                    '#j_id0\\:frm\\:company'
                ],
                [
                    'xpath///*[@id="j_id0:frm:company"]'
                ],
                [
                    'pierce/#j_id0\\:frm\\:company'
                ],
                [
                    'aria/Company Name *'
                ]
            ], targetPage, timeout);
            const element = await waitForSelectors([
                [
                    '#j_id0\\:frm\\:company'
                ],
                [
                    'xpath///*[@id="j_id0:frm:company"]'
                ],
                [
                    'pierce/#j_id0\\:frm\\:company'
                ],
                [
                    'aria/Company Name *'
                ]
            ], targetPage, { timeout, visible: true });
            const inputType = await element.evaluate(el => el.type);
            if (inputType === 'select-one') {
                await changeSelectElement(element, 'na')
            } else if ([
                'textarea',
                'text',
                'url',
                'tel',
                'search',
                'password',
                'number',
                'email'
            ].includes(inputType)) {
                await typeIntoElement(element, 'na');
            } else {
                await changeElementValue(element, 'na');
            }
            await page.waitForTimeout(1000); // wait for 1 seconds
        }
        {
            const targetPage = page;
            await scrollIntoViewIfNeeded([
                [
                    '#j_id0\\:frm\\:selectedRoleId'
                ],
                [
                    'xpath///*[@id="j_id0:frm:selectedRoleId"]'
                ],
                [
                    'pierce/#j_id0\\:frm\\:selectedRoleId'
                ],
                [
                    'aria/Contact Role *'
                ]
            ], targetPage, timeout);
            const element = await waitForSelectors([
                [
                    '#j_id0\\:frm\\:selectedRoleId'
                ],
                [
                    'xpath///*[@id="j_id0:frm:selectedRoleId"]'
                ],
                [
                    'pierce/#j_id0\\:frm\\:selectedRoleId'
                ],
                [
                    'aria/Contact Role *'
                ]
            ], targetPage, { timeout, visible: true });
            await element.click({
                offset: {
                    x: 44.5,
                    y: 14,
                },
            });
            await page.waitForTimeout(1000); // wait for 1 seconds
        }
        {
            const targetPage = page;
            await scrollIntoViewIfNeeded([
                [
                    '#j_id0\\:frm\\:selectedRoleId'
                ],
                [
                    'xpath///*[@id="j_id0:frm:selectedRoleId"]'
                ],
                [
                    'pierce/#j_id0\\:frm\\:selectedRoleId'
                ],
                [
                    'aria/Contact Role *'
                ]
            ], targetPage, timeout);
            const element = await waitForSelectors([
                [
                    '#j_id0\\:frm\\:selectedRoleId'
                ],
                [
                    'xpath///*[@id="j_id0:frm:selectedRoleId"]'
                ],
                [
                    'pierce/#j_id0\\:frm\\:selectedRoleId'
                ],
                [
                    'aria/Contact Role *'
                ]
            ], targetPage, { timeout, visible: true });
            const inputType = await element.evaluate(el => el.type);
            if (inputType === 'select-one') {
                await changeSelectElement(element, 'Customer')
            } else if ([
                'textarea',
                'text',
                'url',
                'tel',
                'search',
                'password',
                'number',
                'email'
            ].includes(inputType)) {
                await typeIntoElement(element, 'Customer');
            } else {
                await changeElementValue(element, 'Customer');
            }
            await page.waitForTimeout(1000); // wait for 1 seconds
        }
        {
            const targetPage = page;
            const promises = [];
            promises.push(targetPage.waitForNavigation());
            await scrollIntoViewIfNeeded([
                [
                    '#saveBtn1'
                ],
                [
                    'xpath///*[@id="saveBtn1"]'
                ],
                [
                    'pierce/#saveBtn1'
                ],
                [
                    'aria/SAVE CHANGES'
                ],
                [
                    'text/Save Changes'
                ]
            ], targetPage, timeout);
            const element = await waitForSelectors([
                [
                    '#saveBtn1'
                ],
                [
                    'xpath///*[@id="saveBtn1"]'
                ],
                [
                    'pierce/#saveBtn1'
                ],
                [
                    'aria/SAVE CHANGES'
                ],
                [
                    'text/Save Changes'
                ]
            ], targetPage, { timeout, visible: true });
            await element.click({
                offset: {
                    x: 73.0625,
                    y: 13,
                },
            });
            await Promise.all(promises);
            await page.waitForTimeout(5000); // wait for 1 seconds
        }
        {
            const targetPage = page;
            const promises = [];
            promises.push(targetPage.waitForNavigation());
            await scrollIntoViewIfNeeded([
                [
                    'a.btn-primary'
                ],
                [
                    'xpath///*[@id="j_id0:frm"]/footer/div/a[3]'
                ],
                [
                    'pierce/a.btn-primary'
                ],
                [
                    'aria/CONTINUE'
                ]
            ], targetPage, timeout);
            const element = await waitForSelectors([
                [
                    'a.btn-primary'
                ],
                [
                    'xpath///*[@id="j_id0:frm"]/footer/div/a[3]'
                ],
                [
                    'pierce/a.btn-primary'
                ],
                [
                    'aria/CONTINUE'
                ]
            ], targetPage, { timeout, visible: true });
            await element.click({
                offset: {
                    x: 91.34375,
                    y: 23.015625,
                },
            });
            await Promise.all(promises);
            await page.waitForTimeout(5000); // wait for 1 seconds
        } 
        {
            const elementHandle = await page.$('[id^="select2-j_id0j_id2frmj_id"]');
            let id = await page.evaluate(element => element.id, elementHandle);
            let match = id.match(/select2-j_id0j_id2frmj_id(\d{3})/);
            let formv = match ? match[1] : null;
            console.log(`Extracted formv: ${formv}`);
            // After constructing selectors, log them to see if they are as expected
            console.log(`Constructed selectors: #j_id0\\:j_id2\\:frm\\:j_id${formv}\\:0\\:qty`);
            /////////////////////////////ARRAY 1 STARTS HERE////////////////////////////////////////////////////////////////////////
            {

                const targetPage = page;
                await scrollIntoViewIfNeeded([
                    [
                        `#j_id0\\:j_id2\\:frm\\:j_id${formv}\\:0\\:qty`
                    ],
                    [
                        `xpath///*[@id="j_id0:j_id2:frm:j_id${formv}:0:qty"]`
                    ],
                    [
                        `pierce/#j_id0\\:j_id2\\:frm\\:j_id${formv}\\:0\\:qty`
                    ]
                ], targetPage, timeout);
                const element = await waitForSelectors([
                    [
                        `#j_id0\\:j_id2\\:frm\\:j_id${formv}\\:0\\:qty`
                    ],
                    [
                        `xpath///*[@id="j_id0:j_id2:frm:j_id${formv}:0:qty"]`
                    ],
                    [
                        `pierce/#j_id0\\:j_id2\\:frm\\:j_id${formv}\\:0\\:qty`
                    ]
                ], targetPage, { timeout, visible: true });
                const inputType = await element.evaluate(el => el.type);
                // Get value from Google Spreadsheet from chatgpt
                const panelValue = datarow.Panels[0]; //Panel input
                if (inputType === 'select-one') {
                    await changeSelectElement(element, panelValue)
                } else if ([
                    'textarea',
                    'text',
                    'url',
                    'tel',
                    'search',
                    'password',
                    'number',
                    'email'
                ].includes(inputType)) {
                    await typeIntoElement(element, panelValue);
                } else {
                    await changeElementValue(element, panelValue);
                }
                await page.waitForTimeout(1000); // wait for 1 seconds
            }
            {
                const targetPage = page;
                await scrollIntoViewIfNeeded([
                    [
                        `#select2-j_id0j_id2frmj_id${formv}0module-container`
                    ],
                    [
                        `xpath///*[@id="select2-j_id0j_id2frmj_id${formv}0module-container"]`
                    ],
                    [
                        `pierce/#select2-j_id0j_id2frmj_id${formv}0module-container`
                    ]
                ], targetPage, timeout);
                const element = await waitForSelectors([
                    [
                        `#select2-j_id0j_id2frmj_id${formv}0module-container`
                    ],
                    [
                        `xpath///*[@id="select2-j_id0j_id2frmj_id${formv}0module-container"]`
                    ],
                    [
                        `pierce/#select2-j_id0j_id2frmj_id${formv}0module-container`
                    ]
                ], targetPage, { timeout, visible: true });
                await element.click({
                    offset: {
                        x: 97,
                        y: 14.015625,
                    },
                });
            }
            {
                const targetPage = page;
                await scrollIntoViewIfNeeded([
                    [
                        'body > span input'
                    ],
                    [
                        'xpath//html/body/span/span/span[1]/input'
                    ],
                    [
                        'pierce/body > span input'
                    ]
                ], targetPage, timeout);
                const element = await waitForSelectors([
                    [
                        'body > span input'
                    ],
                    [
                        'xpath//html/body/span/span/span[1]/input'
                    ],
                    [
                        'pierce/body > span input'
                    ]
                ], targetPage, { timeout, visible: true });
                const inputType = await element.evaluate(el => el.type);
                // Get value from Google Spreadsheet from chatgpt
                const Brand1Value = panelbrand; ////////////////////////////////////Brand1 input
                if (inputType === 'select-one') {
                    await changeSelectElement(element, Brand1Value)
                } else if ([
                    'textarea',
                    'text',
                    'url',
                    'tel',
                    'search',
                    'password',
                    'number',
                    'email'
                ].includes(inputType)) {
                    await typeIntoElement(element, Brand1Value);
                } else {
                    await changeElementValue(element, Brand1Value);
                }
            }
            {
                const targetPage = page;
                await targetPage.keyboard.down('Enter');
            }
            {
                const targetPage = page;
                await targetPage.keyboard.up('Enter');
            }
            {
                const targetPage = page;
                await scrollIntoViewIfNeeded([
                    [
                        `#select2-j_id0j_id2frmj_id${formv}0modelName-container`
                    ],
                    [
                        `xpath///*[@id="select2-j_id0j_id2frmj_id${formv}0modelName-container"]`
                    ],
                    [
                        `pierce/#select2-j_id0j_id2frmj_id${formv}0modelName-container`
                    ]
                ], targetPage, timeout);
                const element = await waitForSelectors([
                    [
                        `#select2-j_id0j_id2frmj_id${formv}0modelName-container`
                    ],
                    [
                        `xpath///*[@id="select2-j_id0j_id2frmj_id${formv}0modelName-container"]`
                    ],
                    [
                        `pierce/#select2-j_id0j_id2frmj_id${formv}0modelName-container`
                    ]
                ], targetPage, { timeout, visible: true });
                await element.click({
                    offset: {
                        x: 110,
                        y: 26.015625,
                    },
                });
                await page.waitForTimeout(3000); // wait for 1 seconds
            }
            {
                const targetPage = page;
                await scrollIntoViewIfNeeded([
                    [
                        'body > span input'
                    ],
                    [
                        'xpath//html/body/span/span/span[1]/input'
                    ],
                    [
                        'pierce/body > span input'
                    ]
                ], targetPage, timeout);
                const element = await waitForSelectors([
                    [
                        'body > span input'
                    ],
                    [
                        'xpath//html/body/span/span/span[1]/input'
                    ],
                    [
                        'pierce/body > span input'
                    ]
                ], targetPage, { timeout, visible: true });
                const inputType = await element.evaluate(el => el.type);
                // Get value from Google Spreadsheet from chatgpt
                const ModelValue = Panel; ////////////////////////////////////Model1 input
                if (inputType === 'select-one') {
                    await changeSelectElement(element, ModelValue)
                } else if ([
                    'textarea',
                    'text',
                    'url',
                    'tel',
                    'search',
                    'password',
                    'number',
                    'email'
                ].includes(inputType)) {
                    await typeIntoElement(element, ModelValue);
                } else {
                    await changeElementValue(element, ModelValue);
                }
            }
            {
                const targetPage = page;
                await targetPage.keyboard.down('Enter');
            }
            {
                const targetPage = page;
                await targetPage.keyboard.up('Enter');
                await page.waitForTimeout(5000); // wait for 1 seconds
            }
            {
                const targetPage = page;
                await scrollIntoViewIfNeeded([
                    [
                        `#j_id0\\:j_id2\\:frm\\:j_id${formv}\\:0\\:arrayType`
                    ],
                    [
                        `xpath///*[@id="j_id0:j_id2:frm:j_id${formv}:0:arrayType"]`
                    ],
                    [
                        `pierce/#j_id0\\:j_id2\\:frm\\:j_id${formv}\\:0\\:arrayType`
                    ],
                    [
                        `aria/Array Type`
                    ]
                ], targetPage, timeout);
                const element = await waitForSelectors([
                    [
                        `#j_id0\\:j_id2\\:frm\\:j_id${formv}\\:0\\:arrayType`
                    ],
                    [
                        `xpath///*[@id="j_id0:j_id2:frm:j_id${formv}:0:arrayType"]`
                    ],
                    [
                        `pierce/#j_id0\\:j_id2\\:frm\\:j_id${formv}\\:0\\:arrayType`
                    ],
                    [
                        `aria/Array Type`
                    ]
                ], targetPage, { timeout, visible: true });
                await element.click({
                    offset: {
                        x: 165,
                        y: 11.015625,
                    },
                });
            }
            {
                const targetPage = page;
                await scrollIntoViewIfNeeded([
                    [
                        `#j_id0\\:j_id2\\:frm\\:j_id${formv}\\:0\\:arrayType`
                    ],
                    [
                        `xpath///*[@id="j_id0:j_id2:frm:j_id${formv}:0:arrayType"]`
                    ],
                    [
                        `pierce/#j_id0\\:j_id2\\:frm\\:j_id${formv}\\:0\\:arrayType`
                    ],
                    [
                        `aria/Array Type`
                    ]
                ], targetPage, timeout);
                const element = await waitForSelectors([
                    [
                        `#j_id0\\:j_id2\\:frm\\:j_id${formv}\\:0\\:arrayType`
                    ],
                    [
                        `xpath///*[@id="j_id0:j_id2:frm:j_id${formv}:0:arrayType"]`
                    ],
                    [
                        `pierce/#j_id0\\:j_id2\\:frm\\:j_id${formv}\\:0\\:arrayType`
                    ],
                    [
                        'aria/Array Type'
                    ]
                ], targetPage, { timeout, visible: true });
                const inputType = await element.evaluate(el => el.type);
                // Get value from Google Spreadsheet from chatgpt
                const AType1Value = Arraytype; ////////////////////////////////////AType1 input
                if (inputType === 'select-one') {
                    await changeSelectElement(element, AType1Value)
                } else if ([
                    'textarea',
                    'text',
                    'url',
                    'tel',
                    'search',
                    'password',
                    'number',
                    'email'
                ].includes(inputType)) {
                    await typeIntoElement(element, AType1Value);
                } else {
                    await changeElementValue(element, AType1Value);
                }
                await page.waitForTimeout(1000); // wait for 1 seconds
            }
            {
                const targetPage = page;
                await scrollIntoViewIfNeeded([
                    [
                        `#j_id0\\:j_id2\\:frm\\:j_id${formv}\\:0\\:arrayAzimuth`
                    ],
                    [
                        `xpath///*[@id="j_id0:j_id2:frm:j_id${formv}:0:arrayAzimuth"]`
                    ],
                    [
                        `pierce/#j_id0\\:j_id2\\:frm\\:j_id${formv}\\:0\\:arrayAzimuth`
                    ],
                    [
                        `aria/Array Azimuth`
                    ]
                ], targetPage, timeout);
                const element = await waitForSelectors([
                    [
                        `#j_id0\\:j_id2\\:frm\\:j_id${formv}\\:0\\:arrayAzimuth`
                    ],
                    [
                        `xpath///*[@id="j_id0:j_id2:frm:j_id${formv}:0:arrayAzimuth"]`
                    ],
                    [
                        `pierce/#j_id0\\:j_id2\\:frm\\:j_id${formv}\\:0\\:arrayAzimuth`
                    ],
                    [
                        `aria/Array Azimuth`
                    ]
                ], targetPage, { timeout, visible: true });
                await element.click({
                    offset: {
                        x: 31,
                        y: 17.515625,
                    },
                });
            }
            {
                const targetPage = page;
                await scrollIntoViewIfNeeded([
                    [
                        `#j_id0\\:j_id2\\:frm\\:j_id${formv}\\:0\\:arrayAzimuth`
                    ],
                    [
                        `xpath///*[@id="j_id0:j_id2:frm:j_id${formv}:0:arrayAzimuth"]`
                    ],
                    [
                        `pierce/#j_id0\\:j_id2\\:frm\\:j_id${formv}\\:0\\:arrayAzimuth`
                    ],
                    [
                        `aria/Array Azimuth`
                    ]
                ], targetPage, timeout);
                const element = await waitForSelectors([
                    [
                        `#j_id0\\:j_id2\\:frm\\:j_id${formv}\\:0\\:arrayAzimuth`
                    ],
                    [
                        `xpath///*[@id="j_id0:j_id2:frm:j_id${formv}:0:arrayAzimuth"]`
                    ],
                    [
                        `pierce/#j_id0\\:j_id2\\:frm\\:j_id${formv}\\:0\\:arrayAzimuth`
                    ],
                    [
                        `aria/Array Azimuth`
                    ]
                ], targetPage, { timeout, visible: true });
                const inputType = await element.evaluate(el => el.type);
                // Get value from Google Spreadsheet from chatgpt
                const Az1Value = datarow.Az[0];        ////////////////////////////////////Az1 input
                if (inputType === 'select-one') {
                    await changeSelectElement(element, Az1Value)
                } else if ([
                    'textarea',
                    'text',
                    'url',
                    'tel',
                    'search',
                    'password',
                    'number',
                    'email'
                ].includes(inputType)) {
                    await typeIntoElement(element, Az1Value);
                } else {
                    await changeElementValue(element, Az1Value);
                }
                await page.waitForTimeout(1000); // wait for 1 seconds
            }
            {
                const targetPage = page;
                await scrollIntoViewIfNeeded([
                    [
                        `#j_id0\\:j_id2\\:frm\\:j_id${formv}\\:0\\:arrayTilt`
                    ],
                    [
                        `xpath///*[@id="j_id0:j_id2:frm:j_id${formv}:0:arrayTilt"]`
                    ],
                    [
                        `pierce/#j_id0\\:j_id2\\:frm\\:j_id${formv}\\:0\\:arrayTilt`
                    ],
                    [
                        `aria/Array Tilt`
                    ]
                ], targetPage, timeout);
                const element = await waitForSelectors([
                    [
                        `#j_id0\\:j_id2\\:frm\\:j_id${formv}\\:0\\:arrayTilt`
                    ],
                    [
                        `xpath///*[@id="j_id0:j_id2:frm:j_id${formv}:0:arrayTilt"]`
                    ],
                    [
                        `pierce/#j_id0\\:j_id2\\:frm\\:j_id${formv}\\:0\\:arrayTilt`
                    ],
                    [
                        `aria/Array Tilt`
                    ]
                ], targetPage, { timeout, visible: true });
                await element.click({
                    offset: {
                        x: 20,
                        y: 18.515625,
                    },
                });
                await page.waitForTimeout(1000); // wait for 1 seconds
            }
            {
                const targetPage = page;
                await scrollIntoViewIfNeeded([
                    [
                        `#j_id0\\:j_id2\\:frm\\:j_id${formv}\\:0\\:arrayTilt`
                    ],
                    [
                        `xpath///*[@id="j_id0:j_id2:frm:j_id${formv}:0:arrayTilt"]`
                    ],
                    [
                        `pierce/#j_id0\\:j_id2\\:frm\\:j_id${formv}\\:0\\:arrayTilt`
                    ],
                    [
                        `aria/Array Tilt`
                    ]
                ], targetPage, timeout);
                const element = await waitForSelectors([
                    [
                        `#j_id0\\:j_id2\\:frm\\:j_id${formv}\\:0\\:arrayTilt`
                    ],
                    [
                        `xpath///*[@id="j_id0:j_id2:frm:j_id${formv}:0:arrayTilt"]`
                    ],
                    [
                        `pierce/#j_id0\\:j_id2\\:frm\\:j_id${formv}\\:0\\:arrayTilt`
                    ],
                    [
                        `aria/Array Tilt`
                    ]
                ], targetPage, { timeout, visible: true });
                const inputType = await element.evaluate(el => el.type);
                // Get value from Google Spreadsheet from chatgpt
                const Tilt1Value = datarow.Tilt[0]; ////////////////////////////////////////Tilt1 enter
                if (inputType === 'select-one') {
                    await changeSelectElement(element, Tilt1Value)
                } else if ([
                    'textarea',
                    'text',
                    'url',
                    'tel',
                    'search',
                    'password',
                    'number',
                    'email'
                ].includes(inputType)) {
                    await typeIntoElement(element, Tilt1Value);
                } else {
                    await changeElementValue(element, Tilt1Value);
                }
                await page.waitForTimeout(1000); // wait for 1 seconds
            }
            {
                const targetPage = page;
                await scrollIntoViewIfNeeded([
                    [
                        `#j_id0\\:j_id2\\:frm\\:j_id${formv}\\:0\\:arrayTilt`
                    ],
                    [
                        `xpath///*[@id="j_id0:j_id2:frm:j_id${formv}:0:arrayTilt"]`
                    ],
                    [
                        `pierce/#j_id0\\:j_id2\\:frm\\:j_id${formv}\\:0\\:arrayTilt`
                    ],
                    [
                        `aria/Array Tilt`
                    ]
                ], targetPage, timeout);
                const element = await waitForSelectors([
                    [
                        `#j_id0\\:j_id2\\:frm\\:j_id${formv}\\:0\\:arrayTilt`
                    ],
                    [
                        `xpath///*[@id="j_id0:j_id2:frm:j_id${formv}:0:arrayTilt"]`
                    ],
                    [
                        `pierce/#j_id0\\:j_id2\\:frm\\:j_id${formv}\\:0\\:arrayTilt`
                    ],
                    [
                        `aria/Array Tilt`
                    ]
                ], targetPage, { timeout, visible: true });
                await element.click({
                    delay: 692.4000000357628,
                    offset: {
                        x: 8,
                        y: 17.515625,
                    },
                });
                await page.waitForTimeout(1000); // wait for 1 seconds
            }
            {
                const targetPage = page;
                await scrollIntoViewIfNeeded([
                    [
                        `div.bodyDiv div:nth-of-type(2) > div:nth-of-type(4)`
                    ],
                    [
                        `xpath///*[@id="array-1-content"]/div/div[2]/div[4]`
                    ],
                    [
                        `pierce/div.bodyDiv div:nth-of-type(2) > div:nth-of-type(4)`
                    ],
                    [
                        `text/Array TSRF0`
                    ]
                ], targetPage, timeout);
                const element = await waitForSelectors([
                    [
                        `div.bodyDiv div:nth-of-type(2) > div:nth-of-type(4)`
                    ],
                    [
                        `xpath///*[@id="array-1-content"]/div/div[2]/div[4]`
                    ],
                    [
                        `pierce/div.bodyDiv div:nth-of-type(2) > div:nth-of-type(4)`
                    ],
                    [
                        `text/Array TSRF0`
                    ]
                ], targetPage, { timeout, visible: true });
                await element.click({
                    delay: 581.6000000238419,
                    offset: {
                        x: 17,
                        y: 46.015625,
                    },
                });
                await page.waitForTimeout(1000); // wait for 1 seconds
            }
            {
                const targetPage = page;
                await scrollIntoViewIfNeeded([
                    [
                        `#j_id0\\:j_id2\\:frm\\:j_id${formv}\\:0\\:tsrf`
                    ],
                    [
                        `xpath///*[@id="j_id0:j_id2:frm:j_id${formv}:0:tsrf"]`
                    ],
                    [
                        `pierce/#j_id0\\:j_id2\\:frm\\:j_id${formv}\\:0\\:tsrf`
                    ],
                    [
                        `aria/Array TSRF`
                    ]
                ], targetPage, timeout);
                const element = await waitForSelectors([
                    [
                        `#j_id0\\:j_id2\\:frm\\:j_id${formv}\\:0\\:tsrf`
                    ],
                    [
                        `xpath///*[@id="j_id0:j_id2:frm:j_id${formv}:0:tsrf"]`
                    ],
                    [
                        `pierce/#j_id0\\:j_id2\\:frm\\:j_id${formv}\\:0\\:tsrf`
                    ],
                    [
                        `aria/Array TSRF`
                    ]
                ], targetPage, { timeout, visible: true });
                const inputType = await element.evaluate(el => el.type);
                // Get value from Google Spreadsheet from chatgpt
                const Tsrf1Value = datarow.Tsrf[0]; //////////////////////////////////////////Tsrf1 ENTER
                if (inputType === 'select-one') {
                    await changeSelectElement(element, Tsrf1Value)
                } else if ([
                    'textarea',
                    'text',
                    'url',
                    'tel',
                    'search',
                    'password',
                    'number',
                    'email'
                ].includes(inputType)) {
                    await typeIntoElement(element, Tsrf1Value);
                } else {
                    await changeElementValue(element, Tsrf1Value);
                }
                await page.waitForTimeout(5000); // wait for 1 seconds
            }
            ////////////////////////////////////////////////ARRAY 2 STARTS HERE//////////////////////////////////////////////////////////////////////////////////////
            if (datarow.Panels[1] !== "" && datarow.Panels[1] !== "0") {
{
    const targetPage = page;
    await scrollIntoViewIfNeeded([
        [
            '#j_id0\\:j_id2\\:frm\\:outerPanel > div.content'
        ],
        [
            'xpath///*[@id="j_id0:j_id2:frm:outerPanel"]/div[2]'
        ],
        [
            'pierce/#j_id0\\:j_id2\\:frm\\:outerPanel > div.content'
        ]
    ], targetPage, timeout);
    const element = await waitForSelectors([
        [
            '#j_id0\\:j_id2\\:frm\\:outerPanel > div.content'
        ],
        [
            'xpath///*[@id="j_id0:j_id2:frm:outerPanel"]/div[2]'
        ],
        [
            'pierce/#j_id0\\:j_id2\\:frm\\:outerPanel > div.content'
        ]
    ], targetPage, { timeout, visible: true });
    await element.click({
      offset: {
        x: 660,
        y: 334.515625,
      },
    });
}

				{
                    const targetPage = page;
                    await scrollIntoViewIfNeeded([
                        [
                            'div.ModulePanel div.add-row > a'
                        ],
                        [
                            'xpath///*[@id="j_id0:j_id2:frm:moduleSection"]/div[2]/a'
                        ],
                        [
                            'pierce/div.ModulePanel div.add-row > a'
                        ],
                        [
                            'aria/ ADD MORE ARRAYS'
                        ],
                        [
                            'text/Add more arrays'
                        ]
                    ], targetPage, timeout);
                    const element = await waitForSelectors([
                        [
                            'div.ModulePanel div.add-row > a'
                        ],
                        [
                            'xpath///*[@id="j_id0:j_id2:frm:moduleSection"]/div[2]/a'
                        ],
                        [
                            'pierce/div.ModulePanel div.add-row > a'
                        ],
                        [
                            'aria/ ADD MORE ARRAYS'
                        ],
                        [
                            'text/Add more arrays'
                        ]
                    ], targetPage, { timeout, visible: true });
                    await element.click({
                        offset: {
                            x: 96,    //111,
                            y: 4.765625, //1.265625,
                        },
                    });
                    await page.waitForTimeout(1000); // wait for 1 seconds
                }

                {
                    const targetPage = page;
                    await scrollIntoViewIfNeeded([
                        [
                            `#j_id0\\:j_id2\\:frm\\:j_id${formv}\\:1\\:qty`
                        ],
                        [
                            `xpath///*[@id="j_id0:j_id2:frm:j_id${formv}:1:qty"]`
                        ],
                        [
                            `pierce/#j_id0\\:j_id2\\:frm\\:j_id${formv}\\:1\\:qty`
                        ]
                    ], targetPage, timeout);
                    const element = await waitForSelectors([
                        [
                            `#j_id0\\:j_id2\\:frm\\:j_id${formv}\\:1\\:qty`
                        ],
                        [
                            `xpath///*[@id="j_id0:j_id2:frm:j_id${formv}:1:qty"]`
                        ],
                        [
                            `pierce/#j_id0\\:j_id2\\:frm\\:j_id${formv}\\:1\\:qty`
                        ]
                    ], targetPage, { timeout, visible: true });
                    await element.click({
                        offset: {
                            x: 43,
                            y: 32.015625,
                        },
                    });
                    await page.waitForTimeout(2000); // wait for 2 seconds
                }
                {
                    const targetPage = page;
                    await scrollIntoViewIfNeeded([
                        [
                            `#j_id0\\:j_id2\\:frm\\:j_id${formv}\\:1\\:qty`
                        ],
                        [
                            `xpath///*[@id="j_id0:j_id2:frm:j_id${formv}:1:qty"]`
                        ],
                        [
                            `pierce/#j_id0\\:j_id2\\:frm\\:j_id${formv}\\:1\\:qty`
                        ]
                    ], targetPage, timeout);
                    const element = await waitForSelectors([
                        [
                            `#j_id0\\:j_id2\\:frm\\:j_id${formv}\\:1\\:qty`
                        ],
                        [
                            `xpath///*[@id="j_id0:j_id2:frm:j_id${formv}:1:qty"]`
                        ],
                        [
                            `pierce/#j_id0\\:j_id2\\:frm\\:j_id${formv}\\:1\\:qty`
                        ]
                    ], targetPage, { timeout, visible: true });
                    const inputType = await element.evaluate(el => el.type);
                    // Get value from Google Spreadsheet from chatgpt
                    const Panel2Value = datarow.Panels[1]; // adjust according to the actual column header in the spreadsheet
                    if (inputType === 'select-one') {
                        await changeSelectElement(element, Panel2Value)
                    } else if ([
                        'textarea',
                        'text',
                        'url',
                        'tel',
                        'search',
                        'password',
                        'number',
                        'email'
                    ].includes(inputType)) {
                        await typeIntoElement(element, Panel2Value);
                    } else {
                        await changeElementValue(element, Panel2Value);
                    }
                    await page.waitForTimeout(2000); // wait for 2 seconds
                }
                {
                    const targetPage = page;
                    await scrollIntoViewIfNeeded([
                        [
                            `#select2-j_id0j_id2frmj_id${formv}1module-container`
                        ],
                        [
                            `xpath///*[@id="select2-j_id0j_id2frmj_id${formv}1module-container"]`
                        ],
                        [
                            `pierce/#select2-j_id0j_id2frmj_id${formv}1module-container`
                        ]
                    ], targetPage, timeout);
                    const element = await waitForSelectors([
                        [
                            `#select2-j_id0j_id2frmj_id${formv}1module-container`
                        ],
                        [
                            `xpath///*[@id="select2-j_id0j_id2frmj_id${formv}1module-container"]`
                        ],
                        [
                            `pierce/#select2-j_id0j_id2frmj_id${formv}1module-container`
                        ]
                    ], targetPage, { timeout, visible: true });
                    await element.click({
                        offset: {
                            x: 74,
                            y: 23.015625,
                        },
                    });
                    await page.waitForTimeout(1000); // wait for 1 seconds
                }
                {
                    const targetPage = page;
                    await scrollIntoViewIfNeeded([
                        [
                            'body > span input'
                        ],
                        [
                            'xpath//html/body/span/span/span[1]/input'
                        ],
                        [
                            'pierce/body > span input'
                        ]
                    ], targetPage, timeout);
                    const element = await waitForSelectors([
                        [
                            'body > span input'
                        ],
                        [
                            'xpath//html/body/span/span/span[1]/input'
                        ],
                        [
                            'pierce/body > span input'
                        ]
                    ], targetPage, { timeout, visible: true });
                    const inputType = await element.evaluate(el => el.type);
                    // Get value from Google Spreadsheet from chatgpt
                    const Brand2Value = panelbrand; // adjust according to the actual column header in the spreadsheet
                    if (inputType === 'select-one') {
                        await changeSelectElement(element, Brand2Value)
                    } else if ([
                        'textarea',
                        'text',
                        'url',
                        'tel',
                        'search',
                        'password',
                        'number',
                        'email'
                    ].includes(inputType)) {
                        await typeIntoElement(element, Brand2Value);
                    } else {
                        await changeElementValue(element, Brand2Value);
                    }
                }
                {
                    const targetPage = page;
                    await targetPage.keyboard.down('Enter');
                }
                {
                    const targetPage = page;
                    await targetPage.keyboard.up('Enter');
                }
                {
                    const targetPage = page;
                    await scrollIntoViewIfNeeded([
                        [
                            `#select2-j_id0j_id2frmj_id${formv}1modelName-container`
                        ],
                        [
                            `xpath///*[@id="select2-j_id0j_id2frmj_id${formv}1modelName-container"]`
                        ],
                        [
                            `pierce/#select2-j_id0j_id2frmj_id${formv}1modelName-container`
                        ]
                    ], targetPage, timeout);
                    const element = await waitForSelectors([
                        [
                            `#select2-j_id0j_id2frmj_id${formv}1modelName-container`
                        ],
                        [
                            `xpath///*[@id="select2-j_id0j_id2frmj_id${formv}1modelName-container"]`
                        ],
                        [
                            `pierce/#select2-j_id0j_id2frmj_id${formv}1modelName-container`
                        ]
                    ], targetPage, { timeout, visible: true });
                    await element.click({
                        offset: {
                            x: 107,
                            y: 20.015625,
                        },
                    });
                }
                {
                    const targetPage = page;
                    await scrollIntoViewIfNeeded([
                        [
                            'body > span input'
                        ],
                        [
                            'xpath//html/body/span/span/span[1]/input'
                        ],
                        [
                            'pierce/body > span input'
                        ]
                    ], targetPage, timeout);
                    const element = await waitForSelectors([
                        [
                            'body > span input'
                        ],
                        [
                            'xpath//html/body/span/span/span[1]/input'
                        ],
                        [
                            'pierce/body > span input'
                        ]
                    ], targetPage, { timeout, visible: true });
                    const inputType = await element.evaluate(el => el.type);
                    // Get value from Google Spreadsheet from chatgpt
                    const Model2Value = Panel; // adjust according to the actual column header in the spreadsheet
                    if (inputType === 'select-one') {
                        await changeSelectElement(element, Model2Value)
                    } else if ([
                        'textarea',
                        'text',
                        'url',
                        'tel',
                        'search',
                        'password',
                        'number',
                        'email'
                    ].includes(inputType)) {
                        await typeIntoElement(element, Model2Value);
                    } else {
                        await changeElementValue(element, Model2Value);
                    }
                }
                {
                    const targetPage = page;
                    await targetPage.keyboard.down('Enter');
                }
                {
                    const targetPage = page;
                    await targetPage.keyboard.up('Enter');
                }
                await page.waitForTimeout(2000); // wait for 1 seconds
                {
                    const targetPage = page;
                    await scrollIntoViewIfNeeded([
                        [
                            `#j_id0\\:j_id2\\:frm\\:j_id${formv}\\:1\\:arrayType`
                        ],
                        [
                            `xpath///*[@id="j_id0:j_id2:frm:j_id${formv}:1:arrayType"]`
                        ],
                        [
                            `pierce/#j_id0\\:j_id2\\:frm\\:j_id${formv}\\:1\\:arrayType`
                        ]
                    ], targetPage, timeout);
                    const element = await waitForSelectors([
                        [
                            `#j_id0\\:j_id2\\:frm\\:j_id${formv}\\:1\\:arrayType`
                        ],
                        [
                            `xpath///*[@id="j_id0:j_id2:frm:j_id${formv}:1:arrayType"]`
                        ],
                        [
                            `pierce/#j_id0\\:j_id2\\:frm\\:j_id${formv}\\:1\\:arrayType`
                        ]
                    ], targetPage, { timeout, visible: true });
                    await element.click({
                        offset: {
                            x: 78,
                            y: 28.015625,
                        },
                    });
                    await page.waitForTimeout(1000); // wait for 1 seconds
                }
                {
                    const targetPage = page;
                    await scrollIntoViewIfNeeded([
                        [
                            `#j_id0\\:j_id2\\:frm\\:j_id${formv}\\:1\\:arrayType`
                        ],
                        [
                            `xpath///*[@id="j_id0:j_id2:frm:j_id${formv}:1:arrayType"]`
                        ],
                        [
                            `pierce/#j_id0\\:j_id2\\:frm\\:j_id${formv}\\:1\\:arrayType`
                        ]
                    ], targetPage, timeout);
                    const element = await waitForSelectors([
                        [
                            `#j_id0\\:j_id2\\:frm\\:j_id${formv}\\:1\\:arrayType`
                        ],
                        [
                            `xpath///*[@id="j_id0:j_id2:frm:j_id${formv}:1:arrayType"]`
                        ],
                        [
                            `pierce/#j_id0\\:j_id2\\:frm\\:j_id${formv}\\:1\\:arrayType`
                        ]
                    ], targetPage, { timeout, visible: true });
                    const inputType = await element.evaluate(el => el.type);
                    // Get value from Google Spreadsheet from chatgpt
                    const AType2Value = Arraytype; // adjust according to the actual column header in the spreadsheet
                    if (inputType === 'select-one') {
                        await changeSelectElement(element, AType2Value)
                    } else if ([
                        'textarea',
                        'text',
                        'url',
                        'tel',
                        'search',
                        'password',
                        'number',
                        'email'
                    ].includes(inputType)) {
                        await typeIntoElement(element, AType2Value);
                    } else {
                        await changeElementValue(element, AType2Value);
                    }
                    await page.waitForTimeout(1000); // wait for 1 seconds
                }
                {
                    const targetPage = page;
                    await scrollIntoViewIfNeeded([
                        [
                            `#j_id0\\:j_id2\\:frm\\:j_id${formv}\\:1\\:arrayAzimuth`
                        ],
                        [
                            `xpath///*[@id="j_id0:j_id2:frm:j_id${formv}:1:arrayAzimuth"]`
                        ],
                        [
                            `pierce/#j_id0\\:j_id2\\:frm\\:j_id${formv}\\:1\\:arrayAzimuth`
                        ]
                    ], targetPage, timeout);
                    const element = await waitForSelectors([
                        [
                            `#j_id0\\:j_id2\\:frm\\:j_id${formv}\\:1\\:arrayAzimuth`
                        ],
                        [
                            `xpath///*[@id="j_id0:j_id2:frm:j_id${formv}:1:arrayAzimuth"]`
                        ],
                        [
                            `pierce/#j_id0\\:j_id2\\:frm\\:j_id${formv}\\:1\\:arrayAzimuth`
                        ]
                    ], targetPage, { timeout, visible: true });
                    await element.click({
                        offset: {
                            x: 24,
                            y: 15.515625,
                        },
                    });
                    await page.waitForTimeout(1000); // wait for 1 seconds
                }
                {
                    const targetPage = page;
                    await scrollIntoViewIfNeeded([
                        [
                            `#j_id0\\:j_id2\\:frm\\:j_id${formv}\\:1\\:arrayAzimuth`
                        ],
                        [
                            `xpath///*[@id="j_id0:j_id2:frm:j_id${formv}:1:arrayAzimuth"]`
                        ],
                        [
                            `pierce/#j_id0\\:j_id2\\:frm\\:j_id${formv}\\:1\\:arrayAzimuth`
                        ]
                    ], targetPage, timeout);
                    const element = await waitForSelectors([
                        [
                            `#j_id0\\:j_id2\\:frm\\:j_id${formv}\\:1\\:arrayAzimuth`
                        ],
                        [
                            `xpath///*[@id="j_id0:j_id2:frm:j_id${formv}:1:arrayAzimuth"]`
                        ],
                        [
                            `pierce/#j_id0\\:j_id2\\:frm\\:j_id${formv}\\:1\\:arrayAzimuth`
                        ]
                    ], targetPage, { timeout, visible: true });
                    const inputType = await element.evaluate(el => el.type);
                    // Get value from Google Spreadsheet from chatgpt
                    const Az2Value = datarow.Az[1]; // adjust according to the actual column header in the spreadsheet
                    if (inputType === 'select-one') {
                        await changeSelectElement(element, Az2Value)
                    } else if ([
                        'textarea',
                        'text',
                        'url',
                        'tel',
                        'search',
                        'password',
                        'number',
                        'email'
                    ].includes(inputType)) {
                        await typeIntoElement(element, Az2Value);
                    } else {
                        await changeElementValue(element, Az2Value);
                    }
                }
                {
                    const targetPage = page;
                    await scrollIntoViewIfNeeded([
                        [
                            `#j_id0\\:j_id2\\:frm\\:j_id${formv}\\:1\\:arrayTilt`
                        ],
                        [
                            `xpath///*[@id="j_id0:j_id2:frm:j_id${formv}:1:arrayTilt"]`
                        ],
                        [
                            `pierce/#j_id0\\:j_id2\\:frm\\:j_id${formv}\\:1\\:arrayTilt`
                        ]
                    ], targetPage, timeout);
                    const element = await waitForSelectors([
                        [
                            `#j_id0\\:j_id2\\:frm\\:j_id${formv}\\:1\\:arrayTilt`
                        ],
                        [
                            `xpath///*[@id="j_id0:j_id2:frm:j_id${formv}:1:arrayTilt"]`
                        ],
                        [
                            `pierce/#j_id0\\:j_id2\\:frm\\:j_id${formv}\\:1\\:arrayTilt`
                        ]
                    ], targetPage, { timeout, visible: true });
                    await element.click({
                        offset: {
                            x: 41,
                            y: 17.515625,
                        },
                    });
                    await page.waitForTimeout(1000); // wait for 1 seconds
                }
                {
                    const targetPage = page;
                    await scrollIntoViewIfNeeded([
                        [
                            `#j_id0\\:j_id2\\:frm\\:j_id${formv}\\:1\\:arrayTilt`
                        ],
                        [
                            `xpath///*[@id="j_id0:j_id2:frm:j_id${formv}:1:arrayTilt"]`
                        ],
                        [
                            `pierce/#j_id0\\:j_id2\\:frm\\:j_id${formv}\\:1\\:arrayTilt`
                        ]
                    ], targetPage, timeout);
                    const element = await waitForSelectors([
                        [
                            `#j_id0\\:j_id2\\:frm\\:j_id${formv}\\:1\\:arrayTilt`
                        ],
                        [
                            `xpath///*[@id="j_id0:j_id2:frm:j_id${formv}:1:arrayTilt"]`
                        ],
                        [
                            `pierce/#j_id0\\:j_id2\\:frm\\:j_id${formv}\\:1\\:arrayTilt`
                        ]
                    ], targetPage, { timeout, visible: true });
                    const inputType = await element.evaluate(el => el.type);
                    // Get value from Google Spreadsheet from chatgpt
                    const Tilt2Value = datarow.Tilt[1]; // adjust according to the actual column header in the spreadsheet
                    if (inputType === 'select-one') {
                        await changeSelectElement(element, Tilt2Value)
                    } else if ([
                        'textarea',
                        'text',
                        'url',
                        'tel',
                        'search',
                        'password',
                        'number',
                        'email'
                    ].includes(inputType)) {
                        await typeIntoElement(element, Tilt2Value);
                    } else {
                        await changeElementValue(element, Tilt2Value);
                    }
                    await page.waitForTimeout(1000); // wait for 1 seconds
                }
                {
                    const targetPage = page;
                    await scrollIntoViewIfNeeded([
                        [
                            `#array-2-content > div > div:nth-of-type(2)`
                        ],
                        [
                            `xpath///*[@id="array-2-content"]/div/div[2]`
                        ],
                        [
                            `pierce/#array-2-content > div > div:nth-of-type(2)`
                        ]
                    ], targetPage, timeout);
                    const element = await waitForSelectors([
                        [
                            `#array-2-content > div > div:nth-of-type(2)`
                        ],
                        [
                            `xpath///*[@id="array-2-content"]/div/div[2]`
                        ],
                        [
                            `pierce/#array-2-content > div > div:nth-of-type(2)`
                        ]
                    ], targetPage, { timeout, visible: true });
                    await element.click({
                        delay: 362.80000001192093,
                        offset: {
                            x: 321,
                            y: 30.015625,
                        },
                    });
                }
                {
                    const targetPage = page;
                    await scrollIntoViewIfNeeded([
                        [
                            `#j_id0\\:j_id2\\:frm\\:j_id${formv}\\:1\\:tsrf`
                        ],
                        [
                            `xpath///*[@id="j_id0:j_id2:frm:j_id${formv}:1:tsrf"]`
                        ],
                        [
                            `pierce/#j_id0\\:j_id2\\:frm\\:j_id${formv}\\:1\\:tsrf`
                        ]
                    ], targetPage, timeout);
                    const element = await waitForSelectors([
                        [
                            `#j_id0\\:j_id2\\:frm\\:j_id${formv}\\:1\\:tsrf`
                        ],
                        [
                            `xpath///*[@id="j_id0:j_id2:frm:j_id${formv}:1:tsrf"]`
                        ],
                        [
                            `pierce/#j_id0\\:j_id2\\:frm\\:j_id${formv}\\:1\\:tsrf`
                        ]
                    ], targetPage, { timeout, visible: true });
                    const inputType = await element.evaluate(el => el.type);
                    // Get value from Google Spreadsheet from chatgpt
                    const Tsrf2Value = datarow.Tsrf[1]; // adjust according to the actual column header in the spreadsheet
                    if (inputType === 'select-one') {
                        await changeSelectElement(element, Tsrf2Value)
                    } else if ([
                        'textarea',
                        'text',
                        'url',
                        'tel',
                        'search',
                        'password',
                        'number',
                        'email'
                    ].includes(inputType)) {
                        await typeIntoElement(element, Tsrf2Value);
                    } else {
                        await changeElementValue(element, Tsrf2Value);
                    }
                    await page.waitForTimeout(1000); // wait for 1 seconds
                }
            }
			
            /////////////////////////////////////////////// ARRAY 3 STARTS BELOW/////////////////////////////////////////////////////////////////////////////////////
            if (datarow.Panels[2] !== "" && datarow.Panels[2] !== "0") {
				{
    const targetPage = page;
    await scrollIntoViewIfNeeded([
        [
            '#array-2-content > div > div:nth-of-type(2)'
        ],
        [
            'xpath///*[@id="array-2-content"]/div/div[2]'
        ],
        [
            'pierce/#array-2-content > div > div:nth-of-type(2)'
        ]
    ], targetPage, timeout);
    const element = await waitForSelectors([
        [
            '#array-2-content > div > div:nth-of-type(2)'
        ],
        [
            'xpath///*[@id="array-2-content"]/div/div[2]'
        ],
        [
            'pierce/#array-2-content > div > div:nth-of-type(2)'
        ]
    ], targetPage, { timeout, visible: true });
    await element.click({
      offset: {
        x: 206,
        y: 80.015625,
      },
    });
}

				
                {
                    const targetPage = page;
                    await scrollIntoViewIfNeeded([
                        [
                            'div.ModulePanel div.add-row > a'
                        ],
                        [
                            'xpath///*[@id="j_id0:j_id2:frm:moduleSection"]/div[3]/a'
                        ],
                        [
                            'pierce/div.ModulePanel div.add-row > a'
                        ],
                        [
                            'aria/ ADD MORE ARRAYS'
                        ],
                        [
                            'text/Add more arrays'
                        ]
                    ], targetPage, timeout);
                    const element = await waitForSelectors([
                        [
                            'div.ModulePanel div.add-row > a'
                        ],
                        [
                            'xpath///*[@id="j_id0:j_id2:frm:moduleSection"]/div[3]/a'
                        ],
                        [
                            'pierce/div.ModulePanel div.add-row > a'
                        ],
                        [
                            'aria/ ADD MORE ARRAYS'
                        ],
                        [
                            'text/Add more arrays'
                        ]
                    ], targetPage, { timeout, visible: true });
                    await element.click({
                        offset: {
                            x: 92,
                            y: 8.265625,
                        },
                    });
                    await page.waitForTimeout(2000); // wait for 1 seconds
                }
                {
                    const targetPage = page;
                    await scrollIntoViewIfNeeded([
                        [
                            `#j_id0\\:j_id2\\:frm\\:j_id${formv}\\:2\\:qty`
                        ],
                        [
                            `xpath///*[@id="j_id0:j_id2:frm:j_id${formv}:2:qty"]`
                        ],
                        [
                            `pierce/#j_id0\\:j_id2\\:frm\\:j_id${formv}\\:2\\:qty`
                        ]
                    ], targetPage, timeout);
                    const element = await waitForSelectors([
                        [
                            `#j_id0\\:j_id2\\:frm\\:j_id${formv}\\:2\\:qty`
                        ],
                        [
                            `xpath///*[@id="j_id0:j_id2:frm:j_id${formv}:2:qty"]`
                        ],
                        [
                            `pierce/#j_id0\\:j_id2\\:frm\\:j_id${formv}\\:2\\:qty`
                        ]
                    ], targetPage, { timeout, visible: true });
                    await element.click({
                        offset: {
                            x: 55,
                            y: 20.015625,
                        },
                    });
                    await page.waitForTimeout(1000); // wait for 1 seconds
                }
                {
                    const targetPage = page;
                    await scrollIntoViewIfNeeded([
                        [
                            `#j_id0\\:j_id2\\:frm\\:j_id${formv}\\:2\\:qty`
                        ],
                        [
                            `xpath///*[@id="j_id0:j_id2:frm:j_id${formv}:2:qty"]`
                        ],
                        [
                            `pierce/#j_id0\\:j_id2\\:frm\\:j_id${formv}\\:2\\:qty`
                        ]
                    ], targetPage, timeout);
                    const element = await waitForSelectors([
                        [
                            `#j_id0\\:j_id2\\:frm\\:j_id${formv}\\:2\\:qty`
                        ],
                        [
                            `xpath///*[@id="j_id0:j_id2:frm:j_id${formv}:2:qty"]`
                        ],
                        [
                            `pierce/#j_id0\\:j_id2\\:frm\\:j_id${formv}\\:2\\:qty`
                        ]
                    ], targetPage, { timeout, visible: true });
                    const inputType = await element.evaluate(el => el.type);
                    // Get value from Google Spreadsheet from chatgpt
                    const Panel3Value = datarow.Panels[2]; // adjust according to the actual column header in the spreadsheet
                    if (inputType === 'select-one') {
                        await changeSelectElement(element, Panel3Value)
                    } else if ([
                        'textarea',
                        'text',
                        'url',
                        'tel',
                        'search',
                        'password',
                        'number',
                        'email'
                    ].includes(inputType)) {
                        await typeIntoElement(element, Panel3Value);
                    } else {
                        await changeElementValue(element, Panel3Value);
                    }
                    await page.waitForTimeout(1000); // wait for 1 seconds
                }
                {
                    const targetPage = page;
                    await scrollIntoViewIfNeeded([
                        [
                            `#select2-j_id0j_id2frmj_id${formv}2module-container`
                        ],
                        [
                            `xpath///*[@id="select2-j_id0j_id2frmj_id${formv}2module-container"]`
                        ],
                        [
                            `pierce/#select2-j_id0j_id2frmj_id${formv}2module-container`
                        ]
                    ], targetPage, timeout);
                    const element = await waitForSelectors([
                        [
                            `#select2-j_id0j_id2frmj_id${formv}2module-container`
                        ],
                        [
                            `xpath///*[@id="select2-j_id0j_id2frmj_id${formv}2module-container"]`
                        ],
                        [
                            `pierce/#select2-j_id0j_id2frmj_id${formv}2module-container`
                        ]
                    ], targetPage, { timeout, visible: true });
                    await element.click({
                        offset: {
                            x: 95,
                            y: 15.015625,
                        },
                    });
                    await page.waitForTimeout(1000); // wait for 1 seconds
                }
                {
                    const targetPage = page;
                    await scrollIntoViewIfNeeded([
                        [
                            'body > span input'
                        ],
                        [
                            'xpath//html/body/span/span/span[1]/input'
                        ],
                        [
                            'pierce/body > span input'
                        ]
                    ], targetPage, timeout);
                    const element = await waitForSelectors([
                        [
                            'body > span input'
                        ],
                        [
                            'xpath//html/body/span/span/span[1]/input'
                        ],
                        [
                            'pierce/body > span input'
                        ]
                    ], targetPage, { timeout, visible: true });
                    const inputType = await element.evaluate(el => el.type);
                    // Get value from Google Spreadsheet from chatgpt
                    const Brand3Value = panelbrand;  // adjust according to the actual column header in the spreadsheet
                    if (inputType === 'select-one') {
                        await changeSelectElement(element, Brand3Value)
                    } else if ([
                        'textarea',
                        'text',
                        'url',
                        'tel',
                        'search',
                        'password',
                        'number',
                        'email'
                    ].includes(inputType)) {
                        await typeIntoElement(element, Brand3Value);
                    } else {
                        await changeElementValue(element, Brand3Value);
                    }
                }
                {
                    const targetPage = page;
                    await targetPage.keyboard.down('Enter');
                }
                {
                    const targetPage = page;
                    await targetPage.keyboard.up('Enter');
                    await page.waitForTimeout(1000); // wait for 1 seconds
                }
                {
                    const targetPage = page;
                    await scrollIntoViewIfNeeded([
                        [
                            `#select2-j_id0j_id2frmj_id${formv}2modelName-container`
                        ],
                        [
                            `xpath///*[@id="select2-j_id0j_id2frmj_id${formv}2modelName-container"]`
                        ],
                        [
                            `pierce/#select2-j_id0j_id2frmj_id${formv}2modelName-container`
                        ]
                    ], targetPage, timeout);
                    const element = await waitForSelectors([
                        [
                            `#select2-j_id0j_id2frmj_id${formv}2modelName-container`
                        ],
                        [
                            `xpath///*[@id="select2-j_id0j_id2frmj_id${formv}2modelName-container"]`
                        ],
                        [
                            `pierce/#select2-j_id0j_id2frmj_id${formv}2modelName-container`
                        ]
                    ], targetPage, { timeout, visible: true });
                    await element.click({
                        offset: {
                            x: 90,
                            y: 17.015625,
                        },
                    });
                    await page.waitForTimeout(1000); // wait for 1 seconds
                }
                {
                    const targetPage = page;
                    await scrollIntoViewIfNeeded([
                        [
                            'body > span input'
                        ],
                        [
                            'xpath//html/body/span/span/span[1]/input'
                        ],
                        [
                            'pierce/body > span input'
                        ]
                    ], targetPage, timeout);
                    const element = await waitForSelectors([
                        [
                            'body > span input'
                        ],
                        [
                            'xpath//html/body/span/span/span[1]/input'
                        ],
                        [
                            'pierce/body > span input'
                        ]
                    ], targetPage, { timeout, visible: true });
                    const inputType = await element.evaluate(el => el.type);
                    // Get value from Google Spreadsheet from chatgpt
                    const Model3Value = Panel; // adjust according to the actual column header in the spreadsheet
                    if (inputType === 'select-one') {
                        await changeSelectElement(element, Model3Value)
                    } else if ([
                        'textarea',
                        'text',
                        'url',
                        'tel',
                        'search',
                        'password',
                        'number',
                        'email'
                    ].includes(inputType)) {
                        await typeIntoElement(element, Model3Value);
                    } else {
                        await changeElementValue(element, Model3Value);
                    }
                }
                {
                    const targetPage = page;
                    await targetPage.keyboard.down('Enter');
                }
                {
                    const targetPage = page;
                    await targetPage.keyboard.up('Enter');
                    await page.waitForTimeout(2000); // wait for 1 seconds
                }
                {
                    const targetPage = page;
                    await scrollIntoViewIfNeeded([
                        [
                            `#j_id0\\:j_id2\\:frm\\:j_id${formv}\\:2\\:arrayType`
                        ],
                        [
                            `xpath///*[@id="j_id0:j_id2:frm:j_id${formv}:2:arrayType"]`
                        ],
                        [
                            `pierce/#j_id0\\:j_id2\\:frm\\:j_id${formv}\\:2\\:arrayType`
                        ]
                    ], targetPage, timeout);
                    const element = await waitForSelectors([
                        [
                            `#j_id0\\:j_id2\\:frm\\:j_id${formv}\\:2\\:arrayType`
                        ],
                        [
                            `xpath///*[@id="j_id0:j_id2:frm:j_id${formv}:2:arrayType"]`
                        ],
                        [
                            `pierce/#j_id0\\:j_id2\\:frm\\:j_id${formv}\\:2\\:arrayType`
                        ]
                    ], targetPage, { timeout, visible: true });
                    await element.click({
                        offset: {
                            x: 44,
                            y: 14.015625,
                        },
                    });
                    await page.waitForTimeout(2000); // wait for 2 seconds
                }
                {
                    const targetPage = page;
                    await scrollIntoViewIfNeeded([
                        [
                            `#j_id0\\:j_id2\\:frm\\:j_id${formv}\\:2\\:arrayType`
                        ],
                        [
                            `xpath///*[@id="j_id0:j_id2:frm:j_id${formv}:2:arrayType"]`
                        ],
                        [
                            `pierce/#j_id0\\:j_id2\\:frm\\:j_id${formv}\\:2\\:arrayType`
                        ]
                    ], targetPage, timeout);
                    const element = await waitForSelectors([
                        [
                            `#j_id0\\:j_id2\\:frm\\:j_id${formv}\\:2\\:arrayType`
                        ],
                        [
                            `xpath///*[@id="j_id0:j_id2:frm:j_id${formv}:2:arrayType"]`
                        ],
                        [
                            `pierce/#j_id0\\:j_id2\\:frm\\:j_id${formv}\\:2\\:arrayType`
                        ]
                    ], targetPage, { timeout, visible: true });
                    const inputType = await element.evaluate(el => el.type);
                    // Get value from Google Spreadsheet from chatgpt
                    const AType3Value = Arraytype; // adjust according to the actual column header in the spreadsheet
                    if (inputType === 'select-one') {
                        await changeSelectElement(element, AType3Value)
                    } else if ([
                        'textarea',
                        'text',
                        'url',
                        'tel',
                        'search',
                        'password',
                        'number',
                        'email'
                    ].includes(inputType)) {
                        await typeIntoElement(element, AType3Value);
                    } else {
                        await changeElementValue(element, AType3Value);
                    }
                    await page.waitForTimeout(2000); // wait for 2 seconds
                }
                {
                    const targetPage = page;
                    await scrollIntoViewIfNeeded([
                        [
                            `#j_id0\\:j_id2\\:frm\\:j_id${formv}\\:2\\:arrayAzimuth`
                        ],
                        [
                            `xpath///*[@id="j_id0:j_id2:frm:j_id${formv}:2:arrayAzimuth"]`
                        ],
                        [
                            `pierce/#j_id0\\:j_id2\\:frm\\:j_id${formv}\\:2\\:arrayAzimuth`
                        ]
                    ], targetPage, timeout);
                    const element = await waitForSelectors([
                        [
                            `#j_id0\\:j_id2\\:frm\\:j_id${formv}\\:2\\:arrayAzimuth`
                        ],
                        [
                            `xpath///*[@id="j_id0:j_id2:frm:j_id${formv}:2:arrayAzimuth"]`
                        ],
                        [
                            `pierce/#j_id0\\:j_id2\\:frm\\:j_id${formv}\\:2\\:arrayAzimuth`
                        ]
                    ], targetPage, { timeout, visible: true });
                    await element.click({
                        offset: {
                            x: 37,
                            y: 22.515625,
                        },
                    });
                    await page.waitForTimeout(1000); // wait for 1 seconds
                }
                {
                    const targetPage = page;
                    await scrollIntoViewIfNeeded([
                        [
                            `#j_id0\\:j_id2\\:frm\\:j_id${formv}\\:2\\:arrayAzimuth`
                        ],
                        [
                            `xpath///*[@id="j_id0:j_id2:frm:j_id${formv}:2:arrayAzimuth"]`
                        ],
                        [
                            `pierce/#j_id0\\:j_id2\\:frm\\:j_id${formv}\\:2\\:arrayAzimuth`
                        ]
                    ], targetPage, timeout);
                    const element = await waitForSelectors([
                        [
                            `#j_id0\\:j_id2\\:frm\\:j_id${formv}\\:2\\:arrayAzimuth`
                        ],
                        [
                            `xpath///*[@id="j_id0:j_id2:frm:j_id${formv}:2:arrayAzimuth"]`
                        ],
                        [
                            `pierce/#j_id0\\:j_id2\\:frm\\:j_id${formv}\\:2\\:arrayAzimuth`
                        ]
                    ], targetPage, { timeout, visible: true });
                    const inputType = await element.evaluate(el => el.type);
                    // Get value from Google Spreadsheet from chatgpt
                    const Az3Value = datarow.Az[2]; // adjust according to the actual column header in the spreadsheet
                    if (inputType === 'select-one') {
                        await changeSelectElement(element, Az3Value)
                    } else if ([
                        'textarea',
                        'text',
                        'url',
                        'tel',
                        'search',
                        'password',
                        'number',
                        'email'
                    ].includes(inputType)) {
                        await typeIntoElement(element, Az3Value);
                    } else {
                        await changeElementValue(element, Az3Value);
                    }
                    await page.waitForTimeout(1000); // wait for 1 seconds
                }
                {
                    const targetPage = page;
                    await scrollIntoViewIfNeeded([
                        [
                            `#j_id0\\:j_id2\\:frm\\:j_id${formv}\\:2\\:arrayTilt`
                        ],
                        [
                            `xpath///*[@id="j_id0:j_id2:frm:j_id${formv}:2:arrayTilt"]`
                        ],
                        [
                            `pierce/#j_id0\\:j_id2\\:frm\\:j_id${formv}\\:2\\:arrayTilt`
                        ]
                    ], targetPage, timeout);
                    const element = await waitForSelectors([
                        [
                            `#j_id0\\:j_id2\\:frm\\:j_id${formv}\\:2\\:arrayTilt`
                        ],
                        [
                            `xpath///*[@id="j_id0:j_id2:frm:j_id${formv}:2:arrayTilt"]`
                        ],
                        [
                            `pierce/#j_id0\\:j_id2\\:frm\\:j_id${formv}\\:2\\:arrayTilt`
                        ]
                    ], targetPage, { timeout, visible: true });
                    await element.click({
                        offset: {
                            x: 255,
                            y: 53.015625,
                        },
                    });
                    await element.click({
                        clickCount: 2,
                        offset: {
                            x: 255,
                            y: 53.015625,
                        },
                    });
                    await page.waitForTimeout(1000); // wait for 1 seconds
                }
                {
                    const targetPage = page;
                    await scrollIntoViewIfNeeded([
                        [
                            `#j_id0\\:j_id2\\:frm\\:j_id${formv}\\:2\\:arrayTilt`
                        ],
                        [
                            `xpath///*[@id="j_id0:j_id2:frm:j_id${formv}:2:arrayTilt"]`
                        ],
                        [
                            `pierce/#j_id0\\:j_id2\\:frm\\:j_id${formv}\\:2\\:arrayTilt`
                        ]
                    ], targetPage, timeout);
                    const element = await waitForSelectors([
                        [
                            `#j_id0\\:j_id2\\:frm\\:j_id${formv}\\:2\\:arrayTilt`
                        ],
                        [
                            `xpath///*[@id="j_id0:j_id2:frm:j_id${formv}:2:arrayTilt"]`
                        ],
                        [
                            `pierce/#j_id0\\:j_id2\\:frm\\:j_id${formv}\\:2\\:arrayTilt`
                        ]
                    ], targetPage, { timeout, visible: true });
                    const inputType = await element.evaluate(el => el.type);
                    // Get value from Google Spreadsheet from chatgpt
                    const Tilt3Value = datarow.Tilt[2]; // adjust according to the actual column header in the spreadsheet
                    if (inputType === 'select-one') {
                        await changeSelectElement(element, Tilt3Value)
                    } else if ([
                        'textarea',
                        'text',
                        'url',
                        'tel',
                        'search',
                        'password',
                        'number',
                        'email'
                    ].includes(inputType)) {
                        await typeIntoElement(element, Tilt3Value);
                    } else {
                        await changeElementValue(element, Tilt3Value);
                    }
                    await page.waitForTimeout(1000); // wait for 1 seconds
                }
                {
                    const targetPage = page;
                    await scrollIntoViewIfNeeded([
                        [
                            `#j_id0\\:j_id2\\:frm\\:j_id${formv}\\:2\\:tsrf`
                        ],
                        [
                            `xpath///*[@id="j_id0:j_id2:frm:j_id${formv}:2:tsrf"]`
                        ],
                        [
                            `pierce/#j_id0\\:j_id2\\:frm\\:j_id${formv}\\:2\\:tsrf`
                        ]
                    ], targetPage, timeout);
                    const element = await waitForSelectors([
                        [
                            `#j_id0\\:j_id2\\:frm\\:j_id${formv}\\:2\\:tsrf`
                        ],
                        [
                            `xpath///*[@id="j_id0:j_id2:frm:j_id${formv}:2:tsrf"]`
                        ],
                        [
                            `pierce/#j_id0\\:j_id2\\:frm\\:j_id${formv}\\:2\\:tsrf`
                        ]
                    ], targetPage, { timeout, visible: true });
                    await element.click({
                        delay: 530.6999999880791,
                        offset: {
                            x: 1,
                            y: 20.515625,
                        },
                    });
                    await page.waitForTimeout(1000); // wait for 1 seconds
                }
                {
                    const targetPage = page;
                    await scrollIntoViewIfNeeded([
                        [
                            `#j_id0\\:j_id2\\:frm\\:j_id${formv}\\:2\\:tsrf`
                        ],
                        [
                            `xpath///*[@id="j_id0:j_id2:frm:j_id${formv}:2:tsrf"]`
                        ],
                        [
                            `pierce/#j_id0\\:j_id2\\:frm\\:j_id${formv}\\:2\\:tsrf`
                        ]
                    ], targetPage, timeout);
                    const element = await waitForSelectors([
                        [
                            `#j_id0\\:j_id2\\:frm\\:j_id${formv}\\:2\\:tsrf`
                        ],
                        [
                            `xpath///*[@id="j_id0:j_id2:frm:j_id${formv}:2:tsrf"]`
                        ],
                        [
                            `pierce/#j_id0\\:j_id2\\:frm\\:j_id${formv}\\:2\\:tsrf`
                        ]
                    ], targetPage, { timeout, visible: true });
                    const inputType = await element.evaluate(el => el.type);
                    // Get value from Google Spreadsheet from chatgpt
                    const Tsrf3Value = datarow.Tsrf[2]; // adjust according to the actual column header in the spreadsheet
                    if (inputType === 'select-one') {
                        await changeSelectElement(element, Tsrf3Value)
                    } else if ([
                        'textarea',
                        'text',
                        'url',
                        'tel',
                        'search',
                        'password',
                        'number',
                        'email'
                    ].includes(inputType)) {
                        await typeIntoElement(element, Tsrf3Value);
                    } else {
                        await changeElementValue(element, Tsrf3Value);
                    }
                    await page.waitForTimeout(1000); // wait for 1 seconds
                }

            }
            /////////////////////////////////////////////// ARRAY 4 STARTS BELOW/////////////////////////////////////////////////////////////////////////////////////
            if (datarow.Panels[3] !== "" && datarow.Panels[3] !== "0") {
				{
    const targetPage = page;
    await scrollIntoViewIfNeeded([
        [
            '#array-3-content > div > div:nth-of-type(2)'
        ],
        [
            'xpath///*[@id="array-3-content"]/div/div[2]'
        ],
        [
            'pierce/#array-3-content > div > div:nth-of-type(2)'
        ]
    ], targetPage, timeout);
    const element = await waitForSelectors([
        [
            '#array-3-content > div > div:nth-of-type(2)'
        ],
        [
            'xpath///*[@id="array-3-content"]/div/div[2]'
        ],
        [
            'pierce/#array-3-content > div > div:nth-of-type(2)'
        ]
    ], targetPage, { timeout, visible: true });
    await element.click({
      offset: {
        x: 426,
        y: 82.015625,
      },
    });
}

                {
                    const targetPage = page;
                    await scrollIntoViewIfNeeded([
                        [
                            'div.ModulePanel div.add-row > a'
                        ],
                        [
                            'xpath///*[@id="j_id0:j_id2:frm:moduleSection"]/div[4]/a'
                        ],
                        [
                            'pierce/div.ModulePanel div.add-row > a'
                        ],
                        [
                            'aria/ ADD MORE ARRAYS'
                        ],
                        [
                            'text/Add more arrays'
                        ]
                    ], targetPage, timeout);
                    const element = await waitForSelectors([
                        [
                            'div.ModulePanel div.add-row > a'
                        ],
                        [
                            'xpath///*[@id="j_id0:j_id2:frm:moduleSection"]/div[4]/a'
                        ],
                        [
                            'pierce/div.ModulePanel div.add-row > a'
                        ],
                        [
                            'aria/ ADD MORE ARRAYS'
                        ],
                        [
                            'text/Add more arrays'
                        ]
                    ], targetPage, { timeout, visible: true });
                    await element.click({
                        offset: {
                            x: 92,
                            y: 8.265625,
                        },
                    });
                    await page.waitForTimeout(2000); // wait for 1 seconds
                }
                {
                    const targetPage = page;
                    await scrollIntoViewIfNeeded([
                        [
                            `#j_id0\\:j_id2\\:frm\\:j_id${formv}\\:3\\:qty`
                        ],
                        [
                            `xpath///*[@id="j_id0:j_id2:frm:j_id${formv}:3:qty"]`
                        ],
                        [
                            `pierce/#j_id0\\:j_id2\\:frm\\:j_id${formv}\\:3\\:qty`
                        ]
                    ], targetPage, timeout);
                    const element = await waitForSelectors([
                        [
                            `#j_id0\\:j_id2\\:frm\\:j_id${formv}\\:3\\:qty`
                        ],
                        [
                            `xpath///*[@id="j_id0:j_id2:frm:j_id${formv}:3:qty"]`
                        ],
                        [
                            `pierce/#j_id0\\:j_id2\\:frm\\:j_id${formv}\\:3\\:qty`
                        ]
                    ], targetPage, { timeout, visible: true });
                    await element.click({
                        offset: {
                            x: 55,
                            y: 20.015625,
                        },
                    });
                    await page.waitForTimeout(1000); // wait for 1 seconds
                }
                {
                    const targetPage = page;
                    await scrollIntoViewIfNeeded([
                        [
                            `#j_id0\\:j_id2\\:frm\\:j_id${formv}\\:3\\:qty`
                        ],
                        [
                            `xpath///*[@id="j_id0:j_id2:frm:j_id${formv}:3:qty"]`
                        ],
                        [
                            `pierce/#j_id0\\:j_id2\\:frm\\:j_id${formv}\\:3\\:qty`
                        ]
                    ], targetPage, timeout);
                    const element = await waitForSelectors([
                        [
                            `#j_id0\\:j_id2\\:frm\\:j_id${formv}\\:3\\:qty`
                        ],
                        [
                            `xpath///*[@id="j_id0:j_id2:frm:j_id${formv}:3:qty"]`
                        ],
                        [
                            `pierce/#j_id0\\:j_id2\\:frm\\:j_id${formv}\\:3\\:qty`
                        ]
                    ], targetPage, { timeout, visible: true });
                    const inputType = await element.evaluate(el => el.type);
                    // Get value from Google Spreadsheet from chatgpt
                    const Panel3Value = datarow.Panels[3]; // adjust according to the actual column header in the spreadsheet
                    if (inputType === 'select-one') {
                        await changeSelectElement(element, Panel3Value)
                    } else if ([
                        'textarea',
                        'text',
                        'url',
                        'tel',
                        'search',
                        'password',
                        'number',
                        'email'
                    ].includes(inputType)) {
                        await typeIntoElement(element, Panel3Value);
                    } else {
                        await changeElementValue(element, Panel3Value);
                    }
                    await page.waitForTimeout(1000); // wait for 1 seconds
                }
                {
                    const targetPage = page;
                    await scrollIntoViewIfNeeded([
                        [
                            `#select2-j_id0j_id2frmj_id${formv}3module-container`
                        ],
                        [
                            `xpath///*[@id="select2-j_id0j_id2frmj_id${formv}3module-container"]`
                        ],
                        [
                            `pierce/#select2-j_id0j_id2frmj_id${formv}3module-container`
                        ]
                    ], targetPage, timeout);
                    const element = await waitForSelectors([
                        [
                            `#select2-j_id0j_id2frmj_id${formv}3module-container`
                        ],
                        [
                            `xpath///*[@id="select2-j_id0j_id2frmj_id${formv}3module-container"]`
                        ],
                        [
                            `pierce/#select2-j_id0j_id2frmj_id${formv}3module-container`
                        ]
                    ], targetPage, { timeout, visible: true });
                    await element.click({
                        offset: {
                            x: 95,
                            y: 15.015625,
                        },
                    });
                    await page.waitForTimeout(1000); // wait for 1 seconds
                }
                {
                    const targetPage = page;
                    await scrollIntoViewIfNeeded([
                        [
                            'body > span input'
                        ],
                        [
                            'xpath//html/body/span/span/span[1]/input'
                        ],
                        [
                            'pierce/body > span input'
                        ]
                    ], targetPage, timeout);
                    const element = await waitForSelectors([
                        [
                            'body > span input'
                        ],
                        [
                            'xpath//html/body/span/span/span[1]/input'
                        ],
                        [
                            'pierce/body > span input'
                        ]
                    ], targetPage, { timeout, visible: true });
                    const inputType = await element.evaluate(el => el.type);
                    // Get value from Google Spreadsheet from chatgpt
                    const Brand3Value = panelbrand; // adjust according to the actual column header in the spreadsheet
                    if (inputType === 'select-one') {
                        await changeSelectElement(element, Brand3Value)
                    } else if ([
                        'textarea',
                        'text',
                        'url',
                        'tel',
                        'search',
                        'password',
                        'number',
                        'email'
                    ].includes(inputType)) {
                        await typeIntoElement(element, Brand3Value);
                    } else {
                        await changeElementValue(element, Brand3Value);
                    }
                }
                {
                    const targetPage = page;
                    await targetPage.keyboard.down('Enter');
                }
                {
                    const targetPage = page;
                    await targetPage.keyboard.up('Enter');
                    await page.waitForTimeout(1000); // wait for 1 seconds
                }
                {
                    const targetPage = page;
                    await scrollIntoViewIfNeeded([
                        [
                            `#select2-j_id0j_id2frmj_id${formv}3modelName-container`
                        ],
                        [
                            `xpath///*[@id="select2-j_id0j_id2frmj_id${formv}3modelName-container"]`
                        ],
                        [
                            `pierce/#select2-j_id0j_id2frmj_id${formv}3modelName-container`
                        ]
                    ], targetPage, timeout);
                    const element = await waitForSelectors([
                        [
                            `#select2-j_id0j_id2frmj_id${formv}3modelName-container`
                        ],
                        [
                            `xpath///*[@id="select2-j_id0j_id2frmj_id${formv}3modelName-container"]`
                        ],
                        [
                            `pierce/#select2-j_id0j_id2frmj_id${formv}3modelName-container`
                        ]
                    ], targetPage, { timeout, visible: true });
                    await element.click({
                        offset: {
                            x: 90,
                            y: 17.015625,
                        },
                    });
                    await page.waitForTimeout(1000); // wait for 1 seconds
                }
                {
                    const targetPage = page;
                    await scrollIntoViewIfNeeded([
                        [
                            'body > span input'
                        ],
                        [
                            'xpath//html/body/span/span/span[1]/input'
                        ],
                        [
                            'pierce/body > span input'
                        ]
                    ], targetPage, timeout);
                    const element = await waitForSelectors([
                        [
                            'body > span input'
                        ],
                        [
                            'xpath//html/body/span/span/span[1]/input'
                        ],
                        [
                            'pierce/body > span input'
                        ]
                    ], targetPage, { timeout, visible: true });
                    const inputType = await element.evaluate(el => el.type);
                    // Get value from Google Spreadsheet from chatgpt
                    const Model3Value = Panel; // adjust according to the actual column header in the spreadsheet
                    if (inputType === 'select-one') {
                        await changeSelectElement(element, Model3Value)
                    } else if ([
                        'textarea',
                        'text',
                        'url',
                        'tel',
                        'search',
                        'password',
                        'number',
                        'email'
                    ].includes(inputType)) {
                        await typeIntoElement(element, Model3Value);
                    } else {
                        await changeElementValue(element, Model3Value);
                    }
                }
                {
                    const targetPage = page;
                    await targetPage.keyboard.down('Enter');
                }
                {
                    const targetPage = page;
                    await targetPage.keyboard.up('Enter');
                    await page.waitForTimeout(2000); // wait for 1 seconds
                }
                {
                    const targetPage = page;
                    await scrollIntoViewIfNeeded([
                        [
                            `#j_id0\\:j_id2\\:frm\\:j_id${formv}\\:3\\:arrayType`
                        ],
                        [
                            `xpath///*[@id="j_id0:j_id2:frm:j_id${formv}:3:arrayType"]`
                        ],
                        [
                            `pierce/#j_id0\\:j_id2\\:frm\\:j_id${formv}\\:3\\:arrayType`
                        ]
                    ], targetPage, timeout);
                    const element = await waitForSelectors([
                        [
                            `#j_id0\\:j_id2\\:frm\\:j_id${formv}\\:3\\:arrayType`
                        ],
                        [
                            `xpath///*[@id="j_id0:j_id2:frm:j_id${formv}:3:arrayType"]`
                        ],
                        [
                            `pierce/#j_id0\\:j_id2\\:frm\\:j_id${formv}\\:3\\:arrayType`
                        ]
                    ], targetPage, { timeout, visible: true });
                    await element.click({
                        offset: {
                            x: 44,
                            y: 14.015625,
                        },
                    });
                    await page.waitForTimeout(2000); // wait for 2 seconds
                }
                {
                    const targetPage = page;
                    await scrollIntoViewIfNeeded([
                        [
                            `#j_id0\\:j_id2\\:frm\\:j_id${formv}\\:3\\:arrayType`
                        ],
                        [
                            `xpath///*[@id="j_id0:j_id2:frm:j_id${formv}:3:arrayType"]`
                        ],
                        [
                            `pierce/#j_id0\\:j_id2\\:frm\\:j_id${formv}\\:3\\:arrayType`
                        ]
                    ], targetPage, timeout);
                    const element = await waitForSelectors([
                        [
                            `#j_id0\\:j_id2\\:frm\\:j_id${formv}\\:3\\:arrayType`
                        ],
                        [
                            `xpath///*[@id="j_id0:j_id2:frm:j_id${formv}:3:arrayType"]`
                        ],
                        [
                            `pierce/#j_id0\\:j_id2\\:frm\\:j_id${formv}\\:3\\:arrayType`
                        ]
                    ], targetPage, { timeout, visible: true });
                    const inputType = await element.evaluate(el => el.type);
                    // Get value from Google Spreadsheet from chatgpt
                    const AType3Value = Arraytype; // adjust according to the actual column header in the spreadsheet
                    if (inputType === 'select-one') {
                        await changeSelectElement(element, AType3Value)
                    } else if ([
                        'textarea',
                        'text',
                        'url',
                        'tel',
                        'search',
                        'password',
                        'number',
                        'email'
                    ].includes(inputType)) {
                        await typeIntoElement(element, AType3Value);
                    } else {
                        await changeElementValue(element, AType3Value);
                    }
                    await page.waitForTimeout(2000); // wait for 2 seconds
                }
                {
                    const targetPage = page;
                    await scrollIntoViewIfNeeded([
                        [
                            `#j_id0\\:j_id2\\:frm\\:j_id${formv}\\:3\\:arrayAzimuth`
                        ],
                        [
                            `xpath///*[@id="j_id0:j_id2:frm:j_id${formv}:3:arrayAzimuth"]`
                        ],
                        [
                            `pierce/#j_id0\\:j_id2\\:frm\\:j_id${formv}\\:3\\:arrayAzimuth`
                        ]
                    ], targetPage, timeout);
                    const element = await waitForSelectors([
                        [
                            `#j_id0\\:j_id2\\:frm\\:j_id${formv}\\:3\\:arrayAzimuth`
                        ],
                        [
                            `xpath///*[@id="j_id0:j_id2:frm:j_id${formv}:3:arrayAzimuth"]`
                        ],
                        [
                            `pierce/#j_id0\\:j_id2\\:frm\\:j_id${formv}\\:3\\:arrayAzimuth`
                        ]
                    ], targetPage, { timeout, visible: true });
                    await element.click({
                        offset: {
                            x: 37,
                            y: 22.515625,
                        },
                    });
                    await page.waitForTimeout(1000); // wait for 1 seconds
                }
                {
                    const targetPage = page;
                    await scrollIntoViewIfNeeded([
                        [
                            `#j_id0\\:j_id2\\:frm\\:j_id${formv}\\:3\\:arrayAzimuth`
                        ],
                        [
                            `xpath///*[@id="j_id0:j_id2:frm:j_id${formv}:3:arrayAzimuth"]`
                        ],
                        [
                            `pierce/#j_id0\\:j_id2\\:frm\\:j_id${formv}\\:3\\:arrayAzimuth`
                        ]
                    ], targetPage, timeout);
                    const element = await waitForSelectors([
                        [
                            `#j_id0\\:j_id2\\:frm\\:j_id${formv}\\:3\\:arrayAzimuth`
                        ],
                        [
                            `xpath///*[@id="j_id0:j_id2:frm:j_id${formv}:3:arrayAzimuth"]`
                        ],
                        [
                            `pierce/#j_id0\\:j_id2\\:frm\\:j_id${formv}\\:3\\:arrayAzimuth`
                        ]
                    ], targetPage, { timeout, visible: true });
                    const inputType = await element.evaluate(el => el.type);
                    // Get value from Google Spreadsheet from chatgpt
                    const Az3Value = datarow.Az[3]; // adjust according to the actual column header in the spreadsheet
                    if (inputType === 'select-one') {
                        await changeSelectElement(element, Az3Value)
                    } else if ([
                        'textarea',
                        'text',
                        'url',
                        'tel',
                        'search',
                        'password',
                        'number',
                        'email'
                    ].includes(inputType)) {
                        await typeIntoElement(element, Az3Value);
                    } else {
                        await changeElementValue(element, Az3Value);
                    }
                    await page.waitForTimeout(1000); // wait for 1 seconds
                }
                {
                    const targetPage = page;
                    await scrollIntoViewIfNeeded([
                        [
                            `#j_id0\\:j_id2\\:frm\\:j_id${formv}\\:3\\:arrayTilt`
                        ],
                        [
                            `xpath///*[@id="j_id0:j_id2:frm:j_id${formv}:3:arrayTilt"]`
                        ],
                        [
                            `pierce/#j_id0\\:j_id2\\:frm\\:j_id${formv}\\:3\\:arrayTilt`
                        ]
                    ], targetPage, timeout);
                    const element = await waitForSelectors([
                        [
                            `#j_id0\\:j_id2\\:frm\\:j_id${formv}\\:3\\:arrayTilt`
                        ],
                        [
                            `xpath///*[@id="j_id0:j_id2:frm:j_id${formv}:3:arrayTilt"]`
                        ],
                        [
                            `pierce/#j_id0\\:j_id2\\:frm\\:j_id${formv}\\:3\\:arrayTilt`
                        ]
                    ], targetPage, { timeout, visible: true });
                    await element.click({
                        offset: {
                            x: 255,
                            y: 53.015625,
                        },
                    });
                    await element.click({
                        clickCount: 2,
                        offset: {
                            x: 255,
                            y: 53.015625,
                        },
                    });
                    await page.waitForTimeout(1000); // wait for 1 seconds
                }
                {
                    const targetPage = page;
                    await scrollIntoViewIfNeeded([
                        [
                            `#j_id0\\:j_id2\\:frm\\:j_id${formv}\\:3\\:arrayTilt`
                        ],
                        [
                            `xpath///*[@id="j_id0:j_id2:frm:j_id${formv}:3:arrayTilt"]`
                        ],
                        [
                            `pierce/#j_id0\\:j_id2\\:frm\\:j_id${formv}\\:3\\:arrayTilt`
                        ]
                    ], targetPage, timeout);
                    const element = await waitForSelectors([
                        [
                            `#j_id0\\:j_id2\\:frm\\:j_id${formv}\\:3\\:arrayTilt`
                        ],
                        [
                            `xpath///*[@id="j_id0:j_id2:frm:j_id${formv}:3:arrayTilt"]`
                        ],
                        [
                            `pierce/#j_id0\\:j_id2\\:frm\\:j_id${formv}\\:3\\:arrayTilt`
                        ]
                    ], targetPage, { timeout, visible: true });
                    const inputType = await element.evaluate(el => el.type);
                    // Get value from Google Spreadsheet from chatgpt
                    const Tilt3Value = datarow.Tilt[3]; // adjust according to the actual column header in the spreadsheet
                    if (inputType === 'select-one') {
                        await changeSelectElement(element, Tilt3Value)
                    } else if ([
                        'textarea',
                        'text',
                        'url',
                        'tel',
                        'search',
                        'password',
                        'number',
                        'email'
                    ].includes(inputType)) {
                        await typeIntoElement(element, Tilt3Value);
                    } else {
                        await changeElementValue(element, Tilt3Value);
                    }
                    await page.waitForTimeout(1000); // wait for 1 seconds
                }
                {
                    const targetPage = page;
                    await scrollIntoViewIfNeeded([
                        [
                            `#j_id0\\:j_id2\\:frm\\:j_id${formv}\\:3\\:tsrf`
                        ],
                        [
                            `xpath///*[@id="j_id0:j_id2:frm:j_id${formv}:3:tsrf"]`
                        ],
                        [
                            `pierce/#j_id0\\:j_id2\\:frm\\:j_id${formv}\\:3\\:tsrf`
                        ]
                    ], targetPage, timeout);
                    const element = await waitForSelectors([
                        [
                            `#j_id0\\:j_id2\\:frm\\:j_id${formv}\\:3\\:tsrf`
                        ],
                        [
                            `xpath///*[@id="j_id0:j_id2:frm:j_id${formv}:3:tsrf"]`
                        ],
                        [
                            `pierce/#j_id0\\:j_id2\\:frm\\:j_id${formv}\\:3\\:tsrf`
                        ]
                    ], targetPage, { timeout, visible: true });
                    await element.click({
                        delay: 530.6999999880791,
                        offset: {
                            x: 1,
                            y: 20.515625,
                        },
                    });
                    await page.waitForTimeout(1000); // wait for 1 seconds
                }
                {
                    const targetPage = page;
                    await scrollIntoViewIfNeeded([
                        [
                            `#j_id0\\:j_id2\\:frm\\:j_id${formv}\\:3\\:tsrf`
                        ],
                        [
                            `xpath///*[@id="j_id0:j_id2:frm:j_id${formv}:3:tsrf"]`
                        ],
                        [
                            `pierce/#j_id0\\:j_id2\\:frm\\:j_id${formv}\\:3\\:tsrf`
                        ]
                    ], targetPage, timeout);
                    const element = await waitForSelectors([
                        [
                            `#j_id0\\:j_id2\\:frm\\:j_id${formv}\\:3\\:tsrf`
                        ],
                        [
                            `xpath///*[@id="j_id0:j_id2:frm:j_id${formv}:3:tsrf"]`
                        ],
                        [
                            `pierce/#j_id0\\:j_id2\\:frm\\:j_id${formv}\\:3\\:tsrf`
                        ]
                    ], targetPage, { timeout, visible: true });
                    const inputType = await element.evaluate(el => el.type);
                    // Get value from Google Spreadsheet from chatgpt
                    const Tsrf3Value = datarow.Tsrf[3]; // adjust according to the actual column header in the spreadsheet
                    if (inputType === 'select-one') {
                        await changeSelectElement(element, Tsrf3Value)
                    } else if ([
                        'textarea',
                        'text',
                        'url',
                        'tel',
                        'search',
                        'password',
                        'number',
                        'email'
                    ].includes(inputType)) {
                        await typeIntoElement(element, Tsrf3Value);
                    } else {
                        await changeElementValue(element, Tsrf3Value);
                    }
                    await page.waitForTimeout(1000); // wait for 1 seconds
                }
            }
            /////////////////////////////////////////////// ARRAY 5 STARTS BELOW/////////////////////////////////////////////////////////////////////////////////////
            if (datarow.Panels[4] !== "" && datarow.Panels[4] !== "0") {
                   {
    const targetPage = page;
    await scrollIntoViewIfNeeded([
        [
            '#array-4-content > div > div:nth-of-type(2)'
        ],
        [
            'xpath///*[@id="array-4-content"]/div/div[2]'
        ],
        [
            'pierce/#array-4-content > div > div:nth-of-type(2)'
        ]
    ], targetPage, timeout);
    const element = await waitForSelectors([
        [
            '#array-4-content > div > div:nth-of-type(2)'
        ],
        [
            'xpath///*[@id="array-4-content"]/div/div[2]'
        ],
        [
            'pierce/#array-4-content > div > div:nth-of-type(2)'
        ]
    ], targetPage, { timeout, visible: true });
    await element.click({
      offset: {
        x: 705,
        y: 71.015625,
      },
    });
}


                {
                    const targetPage = page;
                    await scrollIntoViewIfNeeded([
                        [
                            'div.ModulePanel div.add-row > a'
                        ],
                        [
                            'xpath///*[@id="j_id0:j_id2:frm:moduleSection"]/div[5]/a'
                        ],
                        [
                            'pierce/div.ModulePanel div.add-row > a'
                        ],
                        [
                            'aria/ ADD MORE ARRAYS'
                        ],
                        [
                            'text/Add more arrays'
                        ]
                    ], targetPage, timeout);
                    const element = await waitForSelectors([
                        [
                            'div.ModulePanel div.add-row > a'
                        ],
                        [
                            'xpath///*[@id="j_id0:j_id2:frm:moduleSection"]/div[5]/a'
                        ],
                        [
                            'pierce/div.ModulePanel div.add-row > a'
                        ],
                        [
                            'aria/ ADD MORE ARRAYS'
                        ],
                        [
                            'text/Add more arrays'
                        ]
                    ], targetPage, { timeout, visible: true });
                    await element.click({
                        offset: {
                            x: 92,
                            y: 8.265625,
                        },
                    });
                    await page.waitForTimeout(2000); // wait for 1 seconds
                }
                {
                    const targetPage = page;
                    await scrollIntoViewIfNeeded([
                        [
                            `#j_id0\\:j_id2\\:frm\\:j_id${formv}\\:4\\:qty`
                        ],
                        [
                            `xpath///*[@id="j_id0:j_id2:frm:j_id${formv}:4:qty"]`
                        ],
                        [
                            `pierce/#j_id0\\:j_id2\\:frm\\:j_id${formv}\\:4\\:qty`
                        ]
                    ], targetPage, timeout);
                    const element = await waitForSelectors([
                        [
                            `#j_id0\\:j_id2\\:frm\\:j_id${formv}\\:4\\:qty`
                        ],
                        [
                            `xpath///*[@id="j_id0:j_id2:frm:j_id${formv}:4:qty"]`
                        ],
                        [
                            `pierce/#j_id0\\:j_id2\\:frm\\:j_id${formv}\\:4\\:qty`
                        ]
                    ], targetPage, { timeout, visible: true });
                    await element.click({
                        offset: {
                            x: 55,
                            y: 20.015625,
                        },
                    });
                    await page.waitForTimeout(1000); // wait for 1 seconds
                }
                {
                    const targetPage = page;
                    await scrollIntoViewIfNeeded([
                        [
                            `#j_id0\\:j_id2\\:frm\\:j_id${formv}\\:4\\:qty`
                        ],
                        [
                            `xpath///*[@id="j_id0:j_id2:frm:j_id${formv}:4:qty"]`
                        ],
                        [
                            `pierce/#j_id0\\:j_id2\\:frm\\:j_id${formv}\\:4\\:qty`
                        ]
                    ], targetPage, timeout);
                    const element = await waitForSelectors([
                        [
                            `#j_id0\\:j_id2\\:frm\\:j_id${formv}\\:4\\:qty`
                        ],
                        [
                            `xpath///*[@id="j_id0:j_id2:frm:j_id${formv}:4:qty"]`
                        ],
                        [
                            `pierce/#j_id0\\:j_id2\\:frm\\:j_id${formv}\\:4\\:qty`
                        ]
                    ], targetPage, { timeout, visible: true });
                    const inputType = await element.evaluate(el => el.type);
                    // Get value from Google Spreadsheet from chatgpt
                    const Panel3Value = datarow.Panels[4]; // adjust according to the actual column header in the spreadsheet
                    if (inputType === 'select-one') {
                        await changeSelectElement(element, Panel3Value)
                    } else if ([
                        'textarea',
                        'text',
                        'url',
                        'tel',
                        'search',
                        'password',
                        'number',
                        'email'
                    ].includes(inputType)) {
                        await typeIntoElement(element, Panel3Value);
                    } else {
                        await changeElementValue(element, Panel3Value);
                    }
                    await page.waitForTimeout(1000); // wait for 1 seconds
                }
                {
                    const targetPage = page;
                    await scrollIntoViewIfNeeded([
                        [
                            `#select2-j_id0j_id2frmj_id${formv}4module-container`
                        ],
                        [
                            `xpath///*[@id="select2-j_id0j_id2frmj_id${formv}4module-container"]`
                        ],
                        [
                            `pierce/#select2-j_id0j_id2frmj_id${formv}4module-container`
                        ]
                    ], targetPage, timeout);
                    const element = await waitForSelectors([
                        [
                            `#select2-j_id0j_id2frmj_id${formv}4module-container`
                        ],
                        [
                            `xpath///*[@id="select2-j_id0j_id2frmj_id${formv}4module-container"]`
                        ],
                        [
                            `pierce/#select2-j_id0j_id2frmj_id${formv}4module-container`
                        ]
                    ], targetPage, { timeout, visible: true });
                    await element.click({
                        offset: {
                            x: 95,
                            y: 15.015625,
                        },
                    });
                    await page.waitForTimeout(1000); // wait for 1 seconds
                }
                {
                    const targetPage = page;
                    await scrollIntoViewIfNeeded([
                        [
                            'body > span input'
                        ],
                        [
                            'xpath//html/body/span/span/span[1]/input'
                        ],
                        [
                            'pierce/body > span input'
                        ]
                    ], targetPage, timeout);
                    const element = await waitForSelectors([
                        [
                            'body > span input'
                        ],
                        [
                            'xpath//html/body/span/span/span[1]/input'
                        ],
                        [
                            'pierce/body > span input'
                        ]
                    ], targetPage, { timeout, visible: true });
                    const inputType = await element.evaluate(el => el.type);
                    // Get value from Google Spreadsheet from chatgpt
                    const Brand3Value = panelbrand; // adjust according to the actual column header in the spreadsheet
                    if (inputType === 'select-one') {
                        await changeSelectElement(element, Brand3Value)
                    } else if ([
                        'textarea',
                        'text',
                        'url',
                        'tel',
                        'search',
                        'password',
                        'number',
                        'email'
                    ].includes(inputType)) {
                        await typeIntoElement(element, Brand3Value);
                    } else {
                        await changeElementValue(element, Brand3Value);
                    }
                }
                {
                    const targetPage = page;
                    await targetPage.keyboard.down('Enter');
                }
                {
                    const targetPage = page;
                    await targetPage.keyboard.up('Enter');
                    await page.waitForTimeout(1000); // wait for 1 seconds
                }
                {
                    const targetPage = page;
                    await scrollIntoViewIfNeeded([
                        [
                            `#select2-j_id0j_id2frmj_id${formv}4modelName-container`
                        ],
                        [
                            `xpath///*[@id="select2-j_id0j_id2frmj_id${formv}4modelName-container"]`
                        ],
                        [
                            `pierce/#select2-j_id0j_id2frmj_id${formv}4modelName-container`
                        ]
                    ], targetPage, timeout);
                    const element = await waitForSelectors([
                        [
                            `#select2-j_id0j_id2frmj_id${formv}4modelName-container`
                        ],
                        [
                            `xpath///*[@id="select2-j_id0j_id2frmj_id${formv}4modelName-container"]`
                        ],
                        [
                            `pierce/#select2-j_id0j_id2frmj_id${formv}4modelName-container`
                        ]
                    ], targetPage, { timeout, visible: true });
                    await element.click({
                        offset: {
                            x: 90,
                            y: 17.015625,
                        },
                    });
                    await page.waitForTimeout(1000); // wait for 1 seconds
                }
                {
                    const targetPage = page;
                    await scrollIntoViewIfNeeded([
                        [
                            'body > span input'
                        ],
                        [
                            'xpath//html/body/span/span/span[1]/input'
                        ],
                        [
                            'pierce/body > span input'
                        ]
                    ], targetPage, timeout);
                    const element = await waitForSelectors([
                        [
                            'body > span input'
                        ],
                        [
                            'xpath//html/body/span/span/span[1]/input'
                        ],
                        [
                            'pierce/body > span input'
                        ]
                    ], targetPage, { timeout, visible: true });
                    const inputType = await element.evaluate(el => el.type);
                    // Get value from Google Spreadsheet from chatgpt
                    const Model3Value = Panel; // adjust according to the actual column header in the spreadsheet
                    if (inputType === 'select-one') {
                        await changeSelectElement(element, Model3Value)
                    } else if ([
                        'textarea',
                        'text',
                        'url',
                        'tel',
                        'search',
                        'password',
                        'number',
                        'email'
                    ].includes(inputType)) {
                        await typeIntoElement(element, Model3Value);
                    } else {
                        await changeElementValue(element, Model3Value);
                    }
                }
                {
                    const targetPage = page;
                    await targetPage.keyboard.down('Enter');
                }
                {
                    const targetPage = page;
                    await targetPage.keyboard.up('Enter');
                    await page.waitForTimeout(2000); // wait for 1 seconds
                }
                {
                    const targetPage = page;
                    await scrollIntoViewIfNeeded([
                        [
                            `#j_id0\\:j_id2\\:frm\\:j_id${formv}\\:4\\:arrayType`
                        ],
                        [
                            `xpath///*[@id="j_id0:j_id2:frm:j_id${formv}:4:arrayType"]`
                        ],
                        [
                            `pierce/#j_id0\\:j_id2\\:frm\\:j_id${formv}\\:4\\:arrayType`
                        ]
                    ], targetPage, timeout);
                    const element = await waitForSelectors([
                        [
                            `#j_id0\\:j_id2\\:frm\\:j_id${formv}\\:4\\:arrayType`
                        ],
                        [
                            `xpath///*[@id="j_id0:j_id2:frm:j_id${formv}:4:arrayType"]`
                        ],
                        [
                            `pierce/#j_id0\\:j_id2\\:frm\\:j_id${formv}\\:4\\:arrayType`
                        ]
                    ], targetPage, { timeout, visible: true });
                    await element.click({
                        offset: {
                            x: 44,
                            y: 14.015625,
                        },
                    });
                    await page.waitForTimeout(2000); // wait for 2 seconds
                }
                {
                    const targetPage = page;
                    await scrollIntoViewIfNeeded([
                        [
                            `#j_id0\\:j_id2\\:frm\\:j_id${formv}\\:4\\:arrayType`
                        ],
                        [
                            `xpath///*[@id="j_id0:j_id2:frm:j_id${formv}:4:arrayType"]`
                        ],
                        [
                            `pierce/#j_id0\\:j_id2\\:frm\\:j_id${formv}\\:4\\:arrayType`
                        ]
                    ], targetPage, timeout);
                    const element = await waitForSelectors([
                        [
                            `#j_id0\\:j_id2\\:frm\\:j_id${formv}\\:4\\:arrayType`
                        ],
                        [
                            `xpath///*[@id="j_id0:j_id2:frm:j_id${formv}:4:arrayType"]`
                        ],
                        [
                            `pierce/#j_id0\\:j_id2\\:frm\\:j_id${formv}\\:4\\:arrayType`
                        ]
                    ], targetPage, { timeout, visible: true });
                    const inputType = await element.evaluate(el => el.type);
                    // Get value from Google Spreadsheet from chatgpt
                    const AType3Value = Arraytype; // adjust according to the actual column header in the spreadsheet
                    if (inputType === 'select-one') {
                        await changeSelectElement(element, AType3Value)
                    } else if ([
                        'textarea',
                        'text',
                        'url',
                        'tel',
                        'search',
                        'password',
                        'number',
                        'email'
                    ].includes(inputType)) {
                        await typeIntoElement(element, AType3Value);
                    } else {
                        await changeElementValue(element, AType3Value);
                    }
                    await page.waitForTimeout(2000); // wait for 2 seconds
                }
                {
                    const targetPage = page;
                    await scrollIntoViewIfNeeded([
                        [
                            `#j_id0\\:j_id2\\:frm\\:j_id${formv}\\:4\\:arrayAzimuth`
                        ],
                        [
                            `xpath///*[@id="j_id0:j_id2:frm:j_id${formv}:4:arrayAzimuth"]`
                        ],
                        [
                            `pierce/#j_id0\\:j_id2\\:frm\\:j_id${formv}\\:4\\:arrayAzimuth`
                        ]
                    ], targetPage, timeout);
                    const element = await waitForSelectors([
                        [
                            `#j_id0\\:j_id2\\:frm\\:j_id${formv}\\:4\\:arrayAzimuth`
                        ],
                        [
                            `xpath///*[@id="j_id0:j_id2:frm:j_id${formv}:4:arrayAzimuth"]`
                        ],
                        [
                            `pierce/#j_id0\\:j_id2\\:frm\\:j_id${formv}\\:4\\:arrayAzimuth`
                        ]
                    ], targetPage, { timeout, visible: true });
                    await element.click({
                        offset: {
                            x: 37,
                            y: 22.515625,
                        },
                    });
                    await page.waitForTimeout(1000); // wait for 1 seconds
                }
                {
                    const targetPage = page;
                    await scrollIntoViewIfNeeded([
                        [
                            `#j_id0\\:j_id2\\:frm\\:j_id${formv}\\:4\\:arrayAzimuth`
                        ],
                        [
                            `xpath///*[@id="j_id0:j_id2:frm:j_id${formv}:4:arrayAzimuth"]`
                        ],
                        [
                            `pierce/#j_id0\\:j_id2\\:frm\\:j_id${formv}\\:4\\:arrayAzimuth`
                        ]
                    ], targetPage, timeout);
                    const element = await waitForSelectors([
                        [
                            `#j_id0\\:j_id2\\:frm\\:j_id${formv}\\:4\\:arrayAzimuth`
                        ],
                        [
                            `xpath///*[@id="j_id0:j_id2:frm:j_id${formv}:4:arrayAzimuth"]`
                        ],
                        [
                            `pierce/#j_id0\\:j_id2\\:frm\\:j_id${formv}\\:4\\:arrayAzimuth`
                        ]
                    ], targetPage, { timeout, visible: true });
                    const inputType = await element.evaluate(el => el.type);
                    // Get value from Google Spreadsheet from chatgpt
                    const Az3Value = datarow.Az[4]; // adjust according to the actual column header in the spreadsheet
                    if (inputType === 'select-one') {
                        await changeSelectElement(element, Az3Value)
                    } else if ([
                        'textarea',
                        'text',
                        'url',
                        'tel',
                        'search',
                        'password',
                        'number',
                        'email'
                    ].includes(inputType)) {
                        await typeIntoElement(element, Az3Value);
                    } else {
                        await changeElementValue(element, Az3Value);
                    }
                    await page.waitForTimeout(1000); // wait for 1 seconds
                }
                {
                    const targetPage = page;
                    await scrollIntoViewIfNeeded([
                        [
                            `#j_id0\\:j_id2\\:frm\\:j_id${formv}\\:4\\:arrayTilt`
                        ],
                        [
                            `xpath///*[@id="j_id0:j_id2:frm:j_id${formv}:4:arrayTilt"]`
                        ],
                        [
                            `pierce/#j_id0\\:j_id2\\:frm\\:j_id${formv}\\:4\\:arrayTilt`
                        ]
                    ], targetPage, timeout);
                    const element = await waitForSelectors([
                        [
                            `#j_id0\\:j_id2\\:frm\\:j_id${formv}\\:4\\:arrayTilt`
                        ],
                        [
                            `xpath///*[@id="j_id0:j_id2:frm:j_id${formv}:4:arrayTilt"]`
                        ],
                        [
                            `pierce/#j_id0\\:j_id2\\:frm\\:j_id${formv}\\:4\\:arrayTilt`
                        ]
                    ], targetPage, { timeout, visible: true });
                    await element.click({
                        offset: {
                            x: 255,
                            y: 53.015625,
                        },
                    });
                    await element.click({
                        clickCount: 2,
                        offset: {
                            x: 255,
                            y: 53.015625,
                        },
                    });
                    await page.waitForTimeout(1000); // wait for 1 seconds
                }
                {
                    const targetPage = page;
                    await scrollIntoViewIfNeeded([
                        [
                            `#j_id0\\:j_id2\\:frm\\:j_id${formv}\\:4\\:arrayTilt`
                        ],
                        [
                            `xpath///*[@id="j_id0:j_id2:frm:j_id${formv}:4:arrayTilt"]`
                        ],
                        [
                            `pierce/#j_id0\\:j_id2\\:frm\\:j_id${formv}\\:4\\:arrayTilt`
                        ]
                    ], targetPage, timeout);
                    const element = await waitForSelectors([
                        [
                            `#j_id0\\:j_id2\\:frm\\:j_id${formv}\\:4\\:arrayTilt`
                        ],
                        [
                            `xpath///*[@id="j_id0:j_id2:frm:j_id${formv}:4:arrayTilt"]`
                        ],
                        [
                            `pierce/#j_id0\\:j_id2\\:frm\\:j_id${formv}\\:4\\:arrayTilt`
                        ]
                    ], targetPage, { timeout, visible: true });
                    const inputType = await element.evaluate(el => el.type);
                    // Get value from Google Spreadsheet from chatgpt
                    const Tilt3Value = datarow.Tilt[4]; // adjust according to the actual column header in the spreadsheet
                    if (inputType === 'select-one') {
                        await changeSelectElement(element, Tilt3Value)
                    } else if ([
                        'textarea',
                        'text',
                        'url',
                        'tel',
                        'search',
                        'password',
                        'number',
                        'email'
                    ].includes(inputType)) {
                        await typeIntoElement(element, Tilt3Value);
                    } else {
                        await changeElementValue(element, Tilt3Value);
                    }
                    await page.waitForTimeout(1000); // wait for 1 seconds
                }
                {
                    const targetPage = page;
                    await scrollIntoViewIfNeeded([
                        [
                            `#j_id0\\:j_id2\\:frm\\:j_id${formv}\\:4\\:tsrf`
                        ],
                        [
                            `xpath///*[@id="j_id0:j_id2:frm:j_id${formv}:4:tsrf"]`
                        ],
                        [
                            `pierce/#j_id0\\:j_id2\\:frm\\:j_id${formv}\\:4\\:tsrf`
                        ]
                    ], targetPage, timeout);
                    const element = await waitForSelectors([
                        [
                            `#j_id0\\:j_id2\\:frm\\:j_id${formv}\\:4\\:tsrf`
                        ],
                        [
                            `xpath///*[@id="j_id0:j_id2:frm:j_id${formv}:4:tsrf"]`
                        ],
                        [
                            `pierce/#j_id0\\:j_id2\\:frm\\:j_id${formv}\\:4\\:tsrf`
                        ]
                    ], targetPage, { timeout, visible: true });
                    await element.click({
                        delay: 530.6999999880791,
                        offset: {
                            x: 1,
                            y: 20.515625,
                        },
                    });
                    await page.waitForTimeout(1000); // wait for 1 seconds
                }
                {
                    const targetPage = page;
                    await scrollIntoViewIfNeeded([
                        [
                            `#j_id0\\:j_id2\\:frm\\:j_id${formv}\\:4\\:tsrf`
                        ],
                        [
                            `xpath///*[@id="j_id0:j_id2:frm:j_id${formv}:4:tsrf"]`
                        ],
                        [
                            `pierce/#j_id0\\:j_id2\\:frm\\:j_id${formv}\\:4\\:tsrf`
                        ]
                    ], targetPage, timeout);
                    const element = await waitForSelectors([
                        [
                            `#j_id0\\:j_id2\\:frm\\:j_id${formv}\\:4\\:tsrf`
                        ],
                        [
                            `xpath///*[@id="j_id0:j_id2:frm:j_id${formv}:4:tsrf"]`
                        ],
                        [
                            `pierce/#j_id0\\:j_id2\\:frm\\:j_id${formv}\\:4\\:tsrf`
                        ]
                    ], targetPage, { timeout, visible: true });
                    const inputType = await element.evaluate(el => el.type);
                    // Get value from Google Spreadsheet from chatgpt
                    const Tsrf3Value = datarow.Tsrf[4]; // adjust according to the actual column header in the spreadsheet
                    if (inputType === 'select-one') {
                        await changeSelectElement(element, Tsrf3Value)
                    } else if ([
                        'textarea',
                        'text',
                        'url',
                        'tel',
                        'search',
                        'password',
                        'number',
                        'email'
                    ].includes(inputType)) {
                        await typeIntoElement(element, Tsrf3Value);
                    } else {
                        await changeElementValue(element, Tsrf3Value);
                    }
                    await page.waitForTimeout(1000); // wait for 1 seconds
                }
            }
            ////////////////////////////////////////////// ARRAY 6 STARTS BELOW/////////////////////////////////////////////////////////////////////////////////////
            if (datarow.Panels[5] !== "" && datarow.Panels[5] !== "0") {
				{
    const targetPage = page;
    await scrollIntoViewIfNeeded([
        [
            '#array-5-content > div > div:nth-of-type(2)'
        ],
        [
            'xpath///*[@id="array-5-content"]/div/div[2]'
        ],
        [
            'pierce/#array-5-content > div > div:nth-of-type(2)'
        ]
    ], targetPage, timeout);
    const element = await waitForSelectors([
        [
            '#array-5-content > div > div:nth-of-type(2)'
        ],
        [
            'xpath///*[@id="array-5-content"]/div/div[2]'
        ],
        [
            'pierce/#array-5-content > div > div:nth-of-type(2)'
        ]
    ], targetPage, { timeout, visible: true });
    await element.click({
      offset: {
        x: 670,
        y: 81.515625,
      },
    });
}

                {
                    const targetPage = page;
                    await scrollIntoViewIfNeeded([
                        [
                            'div.ModulePanel div.add-row > a'
                        ],
                        [
                            'xpath///*[@id="j_id0:j_id2:frm:moduleSection"]/div[6]/a'
                        ],
                        [
                            'pierce/div.ModulePanel div.add-row > a'
                        ],
                        [
                            'aria/ ADD MORE ARRAYS'
                        ],
                        [
                            'text/Add more arrays'
                        ]
                    ], targetPage, timeout);
                    const element = await waitForSelectors([
                        [
                            'div.ModulePanel div.add-row > a'
                        ],
                        [
                            'xpath///*[@id="j_id0:j_id2:frm:moduleSection"]/div[6]/a'
                        ],
                        [
                            'pierce/div.ModulePanel div.add-row > a'
                        ],
                        [
                            'aria/ ADD MORE ARRAYS'
                        ],
                        [
                            'text/Add more arrays'
                        ]
                    ], targetPage, { timeout, visible: true });
                    await element.click({
                        offset: {
                            x: 92,
                            y: 8.265625,
                        },
                    });
                    await page.waitForTimeout(2000); // wait for 1 seconds
                }
                {
                    const targetPage = page;
                    await scrollIntoViewIfNeeded([
                        [
                            `#j_id0\\:j_id2\\:frm\\:j_id${formv}\\:5\\:qty`
                        ],
                        [
                            `xpath///*[@id="j_id0:j_id2:frm:j_id${formv}:5:qty"]`
                        ],
                        [
                            `pierce/#j_id0\\:j_id2\\:frm\\:j_id${formv}\\:5\\:qty`
                        ]
                    ], targetPage, timeout);
                    const element = await waitForSelectors([
                        [
                            `#j_id0\\:j_id2\\:frm\\:j_id${formv}\\:5\\:qty`
                        ],
                        [
                            `xpath///*[@id="j_id0:j_id2:frm:j_id${formv}:5:qty"]`
                        ],
                        [
                            `pierce/#j_id0\\:j_id2\\:frm\\:j_id${formv}\\:5\\:qty`
                        ]
                    ], targetPage, { timeout, visible: true });
                    await element.click({
                        offset: {
                            x: 55,
                            y: 20.015625,
                        },
                    });
                    await page.waitForTimeout(1000); // wait for 1 seconds
                }
                {
                    const targetPage = page;
                    await scrollIntoViewIfNeeded([
                        [
                            `#j_id0\\:j_id2\\:frm\\:j_id${formv}\\:5\\:qty`
                        ],
                        [
                            `xpath///*[@id="j_id0:j_id2:frm:j_id${formv}:5:qty"]`
                        ],
                        [
                            `pierce/#j_id0\\:j_id2\\:frm\\:j_id${formv}\\:5\\:qty`
                        ]
                    ], targetPage, timeout);
                    const element = await waitForSelectors([
                        [
                            `#j_id0\\:j_id2\\:frm\\:j_id${formv}\\:5\\:qty`
                        ],
                        [
                            `xpath///*[@id="j_id0:j_id2:frm:j_id${formv}:5:qty"]`
                        ],
                        [
                            `pierce/#j_id0\\:j_id2\\:frm\\:j_id${formv}\\:5\\:qty`
                        ]
                    ], targetPage, { timeout, visible: true });
                    const inputType = await element.evaluate(el => el.type);
                    // Get value from Google Spreadsheet from chatgpt
                    const Panel3Value = datarow.Panels[5]; // adjust according to the actual column header in the spreadsheet
                    if (inputType === 'select-one') {
                        await changeSelectElement(element, Panel3Value)
                    } else if ([
                        'textarea',
                        'text',
                        'url',
                        'tel',
                        'search',
                        'password',
                        'number',
                        'email'
                    ].includes(inputType)) {
                        await typeIntoElement(element, Panel3Value);
                    } else {
                        await changeElementValue(element, Panel3Value);
                    }
                    await page.waitForTimeout(1000); // wait for 1 seconds
                }
                {
                    const targetPage = page;
                    await scrollIntoViewIfNeeded([
                        [
                            `#select2-j_id0j_id2frmj_id${formv}5module-container`
                        ],
                        [
                            `xpath///*[@id="select2-j_id0j_id2frmj_id${formv}5module-container"]`
                        ],
                        [
                            `pierce/#select2-j_id0j_id2frmj_id${formv}5module-container`
                        ]
                    ], targetPage, timeout);
                    const element = await waitForSelectors([
                        [
                            `#select2-j_id0j_id2frmj_id${formv}5module-container`
                        ],
                        [
                            `xpath///*[@id="select2-j_id0j_id2frmj_id${formv}5module-container"]`
                        ],
                        [
                            `pierce/#select2-j_id0j_id2frmj_id${formv}5module-container`
                        ]
                    ], targetPage, { timeout, visible: true });
                    await element.click({
                        offset: {
                            x: 95,
                            y: 15.015625,
                        },
                    });
                    await page.waitForTimeout(1000); // wait for 1 seconds
                }
                {
                    const targetPage = page;
                    await scrollIntoViewIfNeeded([
                        [
                            'body > span input'
                        ],
                        [
                            'xpath//html/body/span/span/span[1]/input'
                        ],
                        [
                            'pierce/body > span input'
                        ]
                    ], targetPage, timeout);
                    const element = await waitForSelectors([
                        [
                            'body > span input'
                        ],
                        [
                            'xpath//html/body/span/span/span[1]/input'
                        ],
                        [
                            'pierce/body > span input'
                        ]
                    ], targetPage, { timeout, visible: true });
                    const inputType = await element.evaluate(el => el.type);
                    // Get value from Google Spreadsheet from chatgpt
                    const Brand3Value = panelbrand; // adjust according to the actual column header in the spreadsheet
                    if (inputType === 'select-one') {
                        await changeSelectElement(element, Brand3Value)
                    } else if ([
                        'textarea',
                        'text',
                        'url',
                        'tel',
                        'search',
                        'password',
                        'number',
                        'email'
                    ].includes(inputType)) {
                        await typeIntoElement(element, Brand3Value);
                    } else {
                        await changeElementValue(element, Brand3Value);
                    }
                }
                {
                    const targetPage = page;
                    await targetPage.keyboard.down('Enter');
                }
                {
                    const targetPage = page;
                    await targetPage.keyboard.up('Enter');
                    await page.waitForTimeout(1000); // wait for 1 seconds
                }
                {
                    const targetPage = page;
                    await scrollIntoViewIfNeeded([
                        [
                            `#select2-j_id0j_id2frmj_id${formv}5modelName-container`
                        ],
                        [
                            `xpath///*[@id="select2-j_id0j_id2frmj_id${formv}5modelName-container"]`
                        ],
                        [
                            `pierce/#select2-j_id0j_id2frmj_id${formv}5modelName-container`
                        ]
                    ], targetPage, timeout);
                    const element = await waitForSelectors([
                        [
                            `#select2-j_id0j_id2frmj_id${formv}5modelName-container`
                        ],
                        [
                            `xpath///*[@id="select2-j_id0j_id2frmj_id${formv}5modelName-container"]`
                        ],
                        [
                            `pierce/#select2-j_id0j_id2frmj_id${formv}5modelName-container`
                        ]
                    ], targetPage, { timeout, visible: true });
                    await element.click({
                        offset: {
                            x: 90,
                            y: 17.015625,
                        },
                    });
                    await page.waitForTimeout(1000); // wait for 1 seconds
                }
                {
                    const targetPage = page;
                    await scrollIntoViewIfNeeded([
                        [
                            'body > span input'
                        ],
                        [
                            'xpath//html/body/span/span/span[1]/input'
                        ],
                        [
                            'pierce/body > span input'
                        ]
                    ], targetPage, timeout);
                    const element = await waitForSelectors([
                        [
                            'body > span input'
                        ],
                        [
                            'xpath//html/body/span/span/span[1]/input'
                        ],
                        [
                            'pierce/body > span input'
                        ]
                    ], targetPage, { timeout, visible: true });
                    const inputType = await element.evaluate(el => el.type);
                    // Get value from Google Spreadsheet from chatgpt
                    const Model3Value = Panel; // adjust according to the actual column header in the spreadsheet
                    if (inputType === 'select-one') {
                        await changeSelectElement(element, Model3Value)
                    } else if ([
                        'textarea',
                        'text',
                        'url',
                        'tel',
                        'search',
                        'password',
                        'number',
                        'email'
                    ].includes(inputType)) {
                        await typeIntoElement(element, Model3Value);
                    } else {
                        await changeElementValue(element, Model3Value);
                    }
                }
                {
                    const targetPage = page;
                    await targetPage.keyboard.down('Enter');
                }
                {
                    const targetPage = page;
                    await targetPage.keyboard.up('Enter');
                    await page.waitForTimeout(3000); // wait for 1 seconds
                }
                {
                    const targetPage = page;
                    await scrollIntoViewIfNeeded([
                        [
                            `#j_id0\\:j_id2\\:frm\\:j_id${formv}\\:5\\:arrayType`
                        ],
                        [
                            `xpath///*[@id="j_id0:j_id2:frm:j_id${formv}:5:arrayType"]`
                        ],
                        [
                            `pierce/#j_id0\\:j_id2\\:frm\\:j_id${formv}\\:5\\:arrayType`
                        ]
                    ], targetPage, timeout);
                    const element = await waitForSelectors([
                        [
                            `#j_id0\\:j_id2\\:frm\\:j_id${formv}\\:5\\:arrayType`
                        ],
                        [
                            `xpath///*[@id="j_id0:j_id2:frm:j_id${formv}:5:arrayType"]`
                        ],
                        [
                            `pierce/#j_id0\\:j_id2\\:frm\\:j_id${formv}\\:5\\:arrayType`
                        ]
                    ], targetPage, { timeout, visible: true });
                    await element.click({
                        offset: {
                            x: 44,
                            y: 14.015625,
                        },
                    });
                    await page.waitForTimeout(3000); // wait for 2 seconds
                }
                {
                    const targetPage = page;
                    await scrollIntoViewIfNeeded([
                        [
                            `#j_id0\\:j_id2\\:frm\\:j_id${formv}\\:5\\:arrayType`
                        ],
                        [
                            `xpath///*[@id="j_id0:j_id2:frm:j_id${formv}:5:arrayType"]`
                        ],
                        [
                            `pierce/#j_id0\\:j_id2\\:frm\\:j_id${formv}\\:5\\:arrayType`
                        ]
                    ], targetPage, timeout);
                    const element = await waitForSelectors([
                        [
                            `#j_id0\\:j_id2\\:frm\\:j_id${formv}\\:5\\:arrayType`
                        ],
                        [
                            `xpath///*[@id="j_id0:j_id2:frm:j_id${formv}:5:arrayType"]`
                        ],
                        [
                            `pierce/#j_id0\\:j_id2\\:frm\\:j_id${formv}\\:5\\:arrayType`
                        ]
                    ], targetPage, { timeout, visible: true });
                    const inputType = await element.evaluate(el => el.type);
                    // Get value from Google Spreadsheet from chatgpt
                    const AType3Value = Arraytype; // adjust according to the actual column header in the spreadsheet
                    if (inputType === 'select-one') {
                        await changeSelectElement(element, AType3Value)
                    } else if ([
                        'textarea',
                        'text',
                        'url',
                        'tel',
                        'search',
                        'password',
                        'number',
                        'email'
                    ].includes(inputType)) {
                        await typeIntoElement(element, AType3Value);
                    } else {
                        await changeElementValue(element, AType3Value);
                    }
                    await page.waitForTimeout(2000); // wait for 2 seconds
                }
                {
                    const targetPage = page;
                    await scrollIntoViewIfNeeded([
                        [
                            `#j_id0\\:j_id2\\:frm\\:j_id${formv}\\:5\\:arrayAzimuth`
                        ],
                        [
                            `xpath///*[@id="j_id0:j_id2:frm:j_id${formv}:5:arrayAzimuth"]`
                        ],
                        [
                            `pierce/#j_id0\\:j_id2\\:frm\\:j_id${formv}\\:5\\:arrayAzimuth`
                        ]
                    ], targetPage, timeout);
                    const element = await waitForSelectors([
                        [
                            `#j_id0\\:j_id2\\:frm\\:j_id${formv}\\:5\\:arrayAzimuth`
                        ],
                        [
                            `xpath///*[@id="j_id0:j_id2:frm:j_id${formv}:5:arrayAzimuth"]`
                        ],
                        [
                            `pierce/#j_id0\\:j_id2\\:frm\\:j_id${formv}\\:5\\:arrayAzimuth`
                        ]
                    ], targetPage, { timeout, visible: true });
                    await element.click({
                        offset: {
                            x: 37,
                            y: 22.515625,
                        },
                    });
                    await page.waitForTimeout(1000); // wait for 1 seconds
                }
                {
                    const targetPage = page;
                    await scrollIntoViewIfNeeded([
                        [
                            `#j_id0\\:j_id2\\:frm\\:j_id${formv}\\:5\\:arrayAzimuth`
                        ],
                        [
                            `xpath///*[@id="j_id0:j_id2:frm:j_id${formv}:5:arrayAzimuth"]`
                        ],
                        [
                            `pierce/#j_id0\\:j_id2\\:frm\\:j_id${formv}\\:5\\:arrayAzimuth`
                        ]
                    ], targetPage, timeout);
                    const element = await waitForSelectors([
                        [
                            `#j_id0\\:j_id2\\:frm\\:j_id${formv}\\:5\\:arrayAzimuth`
                        ],
                        [
                            `xpath///*[@id="j_id0:j_id2:frm:j_id${formv}:5:arrayAzimuth"]`
                        ],
                        [
                            `pierce/#j_id0\\:j_id2\\:frm\\:j_id${formv}\\:5\\:arrayAzimuth`
                        ]
                    ], targetPage, { timeout, visible: true });
                    const inputType = await element.evaluate(el => el.type);
                    // Get value from Google Spreadsheet from chatgpt
                    const Az3Value = datarow.Az[5]; // adjust according to the actual column header in the spreadsheet
                    if (inputType === 'select-one') {
                        await changeSelectElement(element, Az3Value)
                    } else if ([
                        'textarea',
                        'text',
                        'url',
                        'tel',
                        'search',
                        'password',
                        'number',
                        'email'
                    ].includes(inputType)) {
                        await typeIntoElement(element, Az3Value);
                    } else {
                        await changeElementValue(element, Az3Value);
                    }
                    await page.waitForTimeout(1000); // wait for 1 seconds
                }
                {
                    const targetPage = page;
                    await scrollIntoViewIfNeeded([
                        [
                            `#j_id0\\:j_id2\\:frm\\:j_id${formv}\\:5\\:arrayTilt`
                        ],
                        [
                            `xpath///*[@id="j_id0:j_id2:frm:j_id${formv}:5:arrayTilt"]`
                        ],
                        [
                            `pierce/#j_id0\\:j_id2\\:frm\\:j_id${formv}\\:5\\:arrayTilt`
                        ]
                    ], targetPage, timeout);
                    const element = await waitForSelectors([
                        [
                            `#j_id0\\:j_id2\\:frm\\:j_id${formv}\\:5\\:arrayTilt`
                        ],
                        [
                            `xpath///*[@id="j_id0:j_id2:frm:j_id${formv}:5:arrayTilt"]`
                        ],
                        [
                            `pierce/#j_id0\\:j_id2\\:frm\\:j_id${formv}\\:5\\:arrayTilt`
                        ]
                    ], targetPage, { timeout, visible: true });
                    await element.click({
                        offset: {
                            x: 255,
                            y: 53.015625,
                        },
                    });
                    await element.click({
                        clickCount: 2,
                        offset: {
                            x: 255,
                            y: 53.015625,
                        },
                    });
                    await page.waitForTimeout(1000); // wait for 1 seconds
                }
                {
                    const targetPage = page;
                    await scrollIntoViewIfNeeded([
                        [
                            `#j_id0\\:j_id2\\:frm\\:j_id${formv}\\:5\\:arrayTilt`
                        ],
                        [
                            `xpath///*[@id="j_id0:j_id2:frm:j_id${formv}:5:arrayTilt"]`
                        ],
                        [
                            `pierce/#j_id0\\:j_id2\\:frm\\:j_id${formv}\\:5\\:arrayTilt`
                        ]
                    ], targetPage, timeout);
                    const element = await waitForSelectors([
                        [
                            `#j_id0\\:j_id2\\:frm\\:j_id${formv}\\:5\\:arrayTilt`
                        ],
                        [
                            `xpath///*[@id="j_id0:j_id2:frm:j_id${formv}:5:arrayTilt"]`
                        ],
                        [
                            `pierce/#j_id0\\:j_id2\\:frm\\:j_id${formv}\\:5\\:arrayTilt`
                        ]
                    ], targetPage, { timeout, visible: true });
                    const inputType = await element.evaluate(el => el.type);
                    // Get value from Google Spreadsheet from chatgpt
                    const Tilt3Value = datarow.Tilt[5]; // adjust according to the actual column header in the spreadsheet
                    if (inputType === 'select-one') {
                        await changeSelectElement(element, Tilt3Value)
                    } else if ([
                        'textarea',
                        'text',
                        'url',
                        'tel',
                        'search',
                        'password',
                        'number',
                        'email'
                    ].includes(inputType)) {
                        await typeIntoElement(element, Tilt3Value);
                    } else {
                        await changeElementValue(element, Tilt3Value);
                    }
                    await page.waitForTimeout(1000); // wait for 1 seconds
                }
                {
                    const targetPage = page;
                    await scrollIntoViewIfNeeded([
                        [
                            `#j_id0\\:j_id2\\:frm\\:j_id${formv}\\:5\\:tsrf`
                        ],
                        [
                            `xpath///*[@id="j_id0:j_id2:frm:j_id${formv}:5:tsrf"]`
                        ],
                        [
                            `pierce/#j_id0\\:j_id2\\:frm\\:j_id${formv}\\:5\\:tsrf`
                        ]
                    ], targetPage, timeout);
                    const element = await waitForSelectors([
                        [
                            `#j_id0\\:j_id2\\:frm\\:j_id${formv}\\:5\\:tsrf`
                        ],
                        [
                            `xpath///*[@id="j_id0:j_id2:frm:j_id${formv}:5:tsrf"]`
                        ],
                        [
                            `pierce/#j_id0\\:j_id2\\:frm\\:j_id${formv}\\:5\\:tsrf`
                        ]
                    ], targetPage, { timeout, visible: true });
                    await element.click({
                        delay: 530.6999999880791,
                        offset: {
                            x: 1,
                            y: 20.515625,
                        },
                    });
                    await page.waitForTimeout(1000); // wait for 1 seconds
                }
                {
                    const targetPage = page;
                    await scrollIntoViewIfNeeded([
                        [
                            `#j_id0\\:j_id2\\:frm\\:j_id${formv}\\:5\\:tsrf`
                        ],
                        [
                            `xpath///*[@id="j_id0:j_id2:frm:j_id${formv}:5:tsrf"]`
                        ],
                        [
                            `pierce/#j_id0\\:j_id2\\:frm\\:j_id${formv}\\:5\\:tsrf`
                        ]
                    ], targetPage, timeout);
                    const element = await waitForSelectors([
                        [
                            `#j_id0\\:j_id2\\:frm\\:j_id${formv}\\:5\\:tsrf`
                        ],
                        [
                            `xpath///*[@id="j_id0:j_id2:frm:j_id${formv}:5:tsrf"]`
                        ],
                        [
                            `pierce/#j_id0\\:j_id2\\:frm\\:j_id${formv}\\:5\\:tsrf`
                        ]
                    ], targetPage, { timeout, visible: true });
                    const inputType = await element.evaluate(el => el.type);
                    // Get value from Google Spreadsheet from chatgpt
                    const Tsrf3Value = datarow.Tsrf[5]; // adjust according to the actual column header in the spreadsheet
                    if (inputType === 'select-one') {
                        await changeSelectElement(element, Tsrf3Value)
                    } else if ([
                        'textarea',
                        'text',
                        'url',
                        'tel',
                        'search',
                        'password',
                        'number',
                        'email'
                    ].includes(inputType)) {
                        await typeIntoElement(element, Tsrf3Value);
                    } else {
                        await changeElementValue(element, Tsrf3Value);
                    }
                    await page.waitForTimeout(1000); // wait for 1 seconds
                }
            }
            ////////////////////////////////////////////// ARRAY 7 STARTS BELOW/////////////////////////////////////////////////////////////////////////////////////
            if (datarow.Panels[6] !== "" && datarow.Panels[6] !== "0") {
				{
    const targetPage = page;
    await scrollIntoViewIfNeeded([
        [
            '#array-6-content > div > div:nth-of-type(2)'
        ],
        [
            'xpath///*[@id="array-6-content"]/div/div[2]'
        ],
        [
            'pierce/#array-6-content > div > div:nth-of-type(2)'
        ]
    ], targetPage, timeout);
    const element = await waitForSelectors([
        [
            '#array-6-content > div > div:nth-of-type(2)'
        ],
        [
            'xpath///*[@id="array-6-content"]/div/div[2]'
        ],
        [
            'pierce/#array-6-content > div > div:nth-of-type(2)'
        ]
    ], targetPage, { timeout, visible: true });
    await element.click({
      offset: {
        x: 727,
        y: 77.515625,
      },
    });
}

                {
                    const targetPage = page;
                    await scrollIntoViewIfNeeded([
                        [
                            'div.ModulePanel div.add-row > a'
                        ],
                        [
                            'xpath///*[@id="j_id0:j_id2:frm:moduleSection"]/div[7]/a'
                        ],
                        [
                            'pierce/div.ModulePanel div.add-row > a'
                        ],
                        [
                            'aria/ ADD MORE ARRAYS'
                        ],
                        [
                            'text/Add more arrays'
                        ]
                    ], targetPage, timeout);
                    const element = await waitForSelectors([
                        [
                            'div.ModulePanel div.add-row > a'
                        ],
                        [
                            'xpath///*[@id="j_id0:j_id2:frm:moduleSection"]/div[7]/a'
                        ],
                        [
                            'pierce/div.ModulePanel div.add-row > a'
                        ],
                        [
                            'aria/ ADD MORE ARRAYS'
                        ],
                        [
                            'text/Add more arrays'
                        ]
                    ], targetPage, { timeout, visible: true });
                    await element.click({
                        offset: {
                            x: 92,
                            y: 8.265625,
                        },
                    });
                    await page.waitForTimeout(2000); // wait for 1 seconds
                }
                {
                    const targetPage = page;
                    await scrollIntoViewIfNeeded([
                        [
                            `#j_id0\\:j_id2\\:frm\\:j_id${formv}\\:6\\:qty`
                        ],
                        [
                            `xpath///*[@id="j_id0:j_id2:frm:j_id${formv}:6:qty"]`
                        ],
                        [
                            `pierce/#j_id0\\:j_id2\\:frm\\:j_id${formv}\\:6\\:qty`
                        ]
                    ], targetPage, timeout);
                    const element = await waitForSelectors([
                        [
                            `#j_id0\\:j_id2\\:frm\\:j_id${formv}\\:6\\:qty`
                        ],
                        [
                            `xpath///*[@id="j_id0:j_id2:frm:j_id${formv}:6:qty"]`
                        ],
                        [
                            `pierce/#j_id0\\:j_id2\\:frm\\:j_id${formv}\\:6\\:qty`
                        ]
                    ], targetPage, { timeout, visible: true });
                    await element.click({
                        offset: {
                            x: 55,
                            y: 20.015625,
                        },
                    });
                    await page.waitForTimeout(1000); // wait for 1 seconds
                }
                {
                    const targetPage = page;
                    await scrollIntoViewIfNeeded([
                        [
                            `#j_id0\\:j_id2\\:frm\\:j_id${formv}\\:6\\:qty`
                        ],
                        [
                            `xpath///*[@id="j_id0:j_id2:frm:j_id${formv}:6:qty"]`
                        ],
                        [
                            `pierce/#j_id0\\:j_id2\\:frm\\:j_id${formv}\\:6\\:qty`
                        ]
                    ], targetPage, timeout);
                    const element = await waitForSelectors([
                        [
                            `#j_id0\\:j_id2\\:frm\\:j_id${formv}\\:6\\:qty`
                        ],
                        [
                            `xpath///*[@id="j_id0:j_id2:frm:j_id${formv}:6:qty"]`
                        ],
                        [
                            `pierce/#j_id0\\:j_id2\\:frm\\:j_id${formv}\\:6\\:qty`
                        ]
                    ], targetPage, { timeout, visible: true });
                    const inputType = await element.evaluate(el => el.type);
                    // Get value from Google Spreadsheet from chatgpt
                    const Panel3Value = datarow.Panels[6]; // adjust according to the actual column header in the spreadsheet
                    if (inputType === 'select-one') {
                        await changeSelectElement(element, Panel3Value)
                    } else if ([
                        'textarea',
                        'text',
                        'url',
                        'tel',
                        'search',
                        'password',
                        'number',
                        'email'
                    ].includes(inputType)) {
                        await typeIntoElement(element, Panel3Value);
                    } else {
                        await changeElementValue(element, Panel3Value);
                    }
                    await page.waitForTimeout(1000); // wait for 1 seconds
                }
                {
                    const targetPage = page;
                    await scrollIntoViewIfNeeded([
                        [
                            `#select2-j_id0j_id2frmj_id${formv}6module-container`
                        ],
                        [
                            `xpath///*[@id="select2-j_id0j_id2frmj_id${formv}6module-container"]`
                        ],
                        [
                            `pierce/#select2-j_id0j_id2frmj_id${formv}6module-container`
                        ]
                    ], targetPage, timeout);
                    const element = await waitForSelectors([
                        [
                            `#select2-j_id0j_id2frmj_id${formv}6module-container`
                        ],
                        [
                            `xpath///*[@id="select2-j_id0j_id2frmj_id${formv}6module-container"]`
                        ],
                        [
                            `pierce/#select2-j_id0j_id2frmj_id${formv}6module-container`
                        ]
                    ], targetPage, { timeout, visible: true });
                    await element.click({
                        offset: {
                            x: 95,
                            y: 15.015625,
                        },
                    });
                    await page.waitForTimeout(1000); // wait for 1 seconds
                }
                {
                    const targetPage = page;
                    await scrollIntoViewIfNeeded([
                        [
                            'body > span input'
                        ],
                        [
                            'xpath//html/body/span/span/span[1]/input'
                        ],
                        [
                            'pierce/body > span input'
                        ]
                    ], targetPage, timeout);
                    const element = await waitForSelectors([
                        [
                            'body > span input'
                        ],
                        [
                            'xpath//html/body/span/span/span[1]/input'
                        ],
                        [
                            'pierce/body > span input'
                        ]
                    ], targetPage, { timeout, visible: true });
                    const inputType = await element.evaluate(el => el.type);
                    // Get value from Google Spreadsheet from chatgpt
                    const Brand3Value = panelbrand; // adjust according to the actual column header in the spreadsheet
                    if (inputType === 'select-one') {
                        await changeSelectElement(element, Brand3Value)
                    } else if ([
                        'textarea',
                        'text',
                        'url',
                        'tel',
                        'search',
                        'password',
                        'number',
                        'email'
                    ].includes(inputType)) {
                        await typeIntoElement(element, Brand3Value);
                    } else {
                        await changeElementValue(element, Brand3Value);
                    }
                }
                {
                    const targetPage = page;
                    await targetPage.keyboard.down('Enter');
                }
                {
                    const targetPage = page;
                    await targetPage.keyboard.up('Enter');
                    await page.waitForTimeout(1000); // wait for 1 seconds
                }
                {
                    const targetPage = page;
                    await scrollIntoViewIfNeeded([
                        [
                            `#select2-j_id0j_id2frmj_id${formv}6modelName-container`
                        ],
                        [
                            `xpath///*[@id="select2-j_id0j_id2frmj_id${formv}6modelName-container"]`
                        ],
                        [
                            `pierce/#select2-j_id0j_id2frmj_id${formv}6modelName-container`
                        ]
                    ], targetPage, timeout);
                    const element = await waitForSelectors([
                        [
                            `#select2-j_id0j_id2frmj_id${formv}6modelName-container`
                        ],
                        [
                            `xpath///*[@id="select2-j_id0j_id2frmj_id${formv}6modelName-container"]`
                        ],
                        [
                            `pierce/#select2-j_id0j_id2frmj_id${formv}6modelName-container`
                        ]
                    ], targetPage, { timeout, visible: true });
                    await element.click({
                        offset: {
                            x: 90,
                            y: 17.015625,
                        },
                    });
                    await page.waitForTimeout(1000); // wait for 1 seconds
                }
                {
                    const targetPage = page;
                    await scrollIntoViewIfNeeded([
                        [
                            'body > span input'
                        ],
                        [
                            'xpath//html/body/span/span/span[1]/input'
                        ],
                        [
                            'pierce/body > span input'
                        ]
                    ], targetPage, timeout);
                    const element = await waitForSelectors([
                        [
                            'body > span input'
                        ],
                        [
                            'xpath//html/body/span/span/span[1]/input'
                        ],
                        [
                            'pierce/body > span input'
                        ]
                    ], targetPage, { timeout, visible: true });
                    const inputType = await element.evaluate(el => el.type);
                    // Get value from Google Spreadsheet from chatgpt
                    const Model3Value = Panel; // adjust according to the actual column header in the spreadsheet
                    if (inputType === 'select-one') {
                        await changeSelectElement(element, Model3Value)
                    } else if ([
                        'textarea',
                        'text',
                        'url',
                        'tel',
                        'search',
                        'password',
                        'number',
                        'email'
                    ].includes(inputType)) {
                        await typeIntoElement(element, Model3Value);
                    } else {
                        await changeElementValue(element, Model3Value);
                    }
                }
                {
                    const targetPage = page;
                    await targetPage.keyboard.down('Enter');
                }
                {
                    const targetPage = page;
                    await targetPage.keyboard.up('Enter');
                    await page.waitForTimeout(2000); // wait for 1 seconds
                }
                {
                    const targetPage = page;
                    await scrollIntoViewIfNeeded([
                        [
                            `#j_id0\\:j_id2\\:frm\\:j_id${formv}\\:6\\:arrayType`
                        ],
                        [
                            `xpath///*[@id="j_id0:j_id2:frm:j_id${formv}:6:arrayType"]`
                        ],
                        [
                            `pierce/#j_id0\\:j_id2\\:frm\\:j_id${formv}\\:6\\:arrayType`
                        ]
                    ], targetPage, timeout);
                    const element = await waitForSelectors([
                        [
                            `#j_id0\\:j_id2\\:frm\\:j_id${formv}\\:6\\:arrayType`
                        ],
                        [
                            `xpath///*[@id="j_id0:j_id2:frm:j_id${formv}:6:arrayType"]`
                        ],
                        [
                            `pierce/#j_id0\\:j_id2\\:frm\\:j_id${formv}\\:6\\:arrayType`
                        ]
                    ], targetPage, { timeout, visible: true });
                    await element.click({
                        offset: {
                            x: 44,
                            y: 14.015625,
                        },
                    });
                    await page.waitForTimeout(2000); // wait for 2 seconds
                }
                {
                    const targetPage = page;
                    await scrollIntoViewIfNeeded([
                        [
                            `#j_id0\\:j_id2\\:frm\\:j_id${formv}\\:6\\:arrayType`
                        ],
                        [
                            `xpath///*[@id="j_id0:j_id2:frm:j_id${formv}:6:arrayType"]`
                        ],
                        [
                            `pierce/#j_id0\\:j_id2\\:frm\\:j_id${formv}\\:6\\:arrayType`
                        ]
                    ], targetPage, timeout);
                    const element = await waitForSelectors([
                        [
                            `#j_id0\\:j_id2\\:frm\\:j_id${formv}\\:6\\:arrayType`
                        ],
                        [
                            `xpath///*[@id="j_id0:j_id2:frm:j_id${formv}:6:arrayType"]`
                        ],
                        [
                            `pierce/#j_id0\\:j_id2\\:frm\\:j_id${formv}\\:6\\:arrayType`
                        ]
                    ], targetPage, { timeout, visible: true });
                    const inputType = await element.evaluate(el => el.type);
                    // Get value from Google Spreadsheet from chatgpt
                    const AType3Value = Arraytype; // adjust according to the actual column header in the spreadsheet
                    if (inputType === 'select-one') {
                        await changeSelectElement(element, AType3Value)
                    } else if ([
                        'textarea',
                        'text',
                        'url',
                        'tel',
                        'search',
                        'password',
                        'number',
                        'email'
                    ].includes(inputType)) {
                        await typeIntoElement(element, AType3Value);
                    } else {
                        await changeElementValue(element, AType3Value);
                    }
                    await page.waitForTimeout(2000); // wait for 2 seconds
                }
                {
                    const targetPage = page;
                    await scrollIntoViewIfNeeded([
                        [
                            `#j_id0\\:j_id2\\:frm\\:j_id${formv}\\:6\\:arrayAzimuth`
                        ],
                        [
                            `xpath///*[@id="j_id0:j_id2:frm:j_id${formv}:6:arrayAzimuth"]`
                        ],
                        [
                            `pierce/#j_id0\\:j_id2\\:frm\\:j_id${formv}\\:6\\:arrayAzimuth`
                        ]
                    ], targetPage, timeout);
                    const element = await waitForSelectors([
                        [
                            `#j_id0\\:j_id2\\:frm\\:j_id${formv}\\:6\\:arrayAzimuth`
                        ],
                        [
                            `xpath///*[@id="j_id0:j_id2:frm:j_id${formv}:6:arrayAzimuth"]`
                        ],
                        [
                            `pierce/#j_id0\\:j_id2\\:frm\\:j_id${formv}\\:6\\:arrayAzimuth`
                        ]
                    ], targetPage, { timeout, visible: true });
                    await element.click({
                        offset: {
                            x: 37,
                            y: 22.515625,
                        },
                    });
                    await page.waitForTimeout(1000); // wait for 1 seconds
                }
                {
                    const targetPage = page;
                    await scrollIntoViewIfNeeded([
                        [
                            `#j_id0\\:j_id2\\:frm\\:j_id${formv}\\:6\\:arrayAzimuth`
                        ],
                        [
                            `xpath///*[@id="j_id0:j_id2:frm:j_id${formv}:6:arrayAzimuth"]`
                        ],
                        [
                            `pierce/#j_id0\\:j_id2\\:frm\\:j_id${formv}\\:6\\:arrayAzimuth`
                        ]
                    ], targetPage, timeout);
                    const element = await waitForSelectors([
                        [
                            `#j_id0\\:j_id2\\:frm\\:j_id${formv}\\:6\\:arrayAzimuth`
                        ],
                        [
                            `xpath///*[@id="j_id0:j_id2:frm:j_id${formv}:6:arrayAzimuth"]`
                        ],
                        [
                            `pierce/#j_id0\\:j_id2\\:frm\\:j_id${formv}\\:6\\:arrayAzimuth`
                        ]
                    ], targetPage, { timeout, visible: true });
                    const inputType = await element.evaluate(el => el.type);
                    // Get value from Google Spreadsheet from chatgpt
                    const Az3Value = datarow.Az[6]; // adjust according to the actual column header in the spreadsheet
                    if (inputType === 'select-one') {
                        await changeSelectElement(element, Az3Value)
                    } else if ([
                        'textarea',
                        'text',
                        'url',
                        'tel',
                        'search',
                        'password',
                        'number',
                        'email'
                    ].includes(inputType)) {
                        await typeIntoElement(element, Az3Value);
                    } else {
                        await changeElementValue(element, Az3Value);
                    }
                    await page.waitForTimeout(1000); // wait for 1 seconds
                }
                {
                    const targetPage = page;
                    await scrollIntoViewIfNeeded([
                        [
                            `#j_id0\\:j_id2\\:frm\\:j_id${formv}\\:6\\:arrayTilt`
                        ],
                        [
                            `xpath///*[@id="j_id0:j_id2:frm:j_id${formv}:6:arrayTilt"]`
                        ],
                        [
                            `pierce/#j_id0\\:j_id2\\:frm\\:j_id${formv}\\:6\\:arrayTilt`
                        ]
                    ], targetPage, timeout);
                    const element = await waitForSelectors([
                        [
                            `#j_id0\\:j_id2\\:frm\\:j_id${formv}\\:6\\:arrayTilt`
                        ],
                        [
                            `xpath///*[@id="j_id0:j_id2:frm:j_id${formv}:6:arrayTilt"]`
                        ],
                        [
                            `pierce/#j_id0\\:j_id2\\:frm\\:j_id${formv}\\:6\\:arrayTilt`
                        ]
                    ], targetPage, { timeout, visible: true });
                    await element.click({
                        offset: {
                            x: 255,
                            y: 53.015625,
                        },
                    });
                    await element.click({
                        clickCount: 2,
                        offset: {
                            x: 255,
                            y: 53.015625,
                        },
                    });
                    await page.waitForTimeout(1000); // wait for 1 seconds
                }
                {
                    const targetPage = page;
                    await scrollIntoViewIfNeeded([
                        [
                            `#j_id0\\:j_id2\\:frm\\:j_id${formv}\\:6\\:arrayTilt`
                        ],
                        [
                            `xpath///*[@id="j_id0:j_id2:frm:j_id${formv}:6:arrayTilt"]`
                        ],
                        [
                            `pierce/#j_id0\\:j_id2\\:frm\\:j_id${formv}\\:6\\:arrayTilt`
                        ]
                    ], targetPage, timeout);
                    const element = await waitForSelectors([
                        [
                            `#j_id0\\:j_id2\\:frm\\:j_id${formv}\\:6\\:arrayTilt`
                        ],
                        [
                            `xpath///*[@id="j_id0:j_id2:frm:j_id${formv}:6:arrayTilt"]`
                        ],
                        [
                            `pierce/#j_id0\\:j_id2\\:frm\\:j_id${formv}\\:6\\:arrayTilt`
                        ]
                    ], targetPage, { timeout, visible: true });
                    const inputType = await element.evaluate(el => el.type);
                    // Get value from Google Spreadsheet from chatgpt
                    const Tilt3Value = datarow.Tilt[6]; // adjust according to the actual column header in the spreadsheet
                    if (inputType === 'select-one') {
                        await changeSelectElement(element, Tilt3Value)
                    } else if ([
                        'textarea',
                        'text',
                        'url',
                        'tel',
                        'search',
                        'password',
                        'number',
                        'email'
                    ].includes(inputType)) {
                        await typeIntoElement(element, Tilt3Value);
                    } else {
                        await changeElementValue(element, Tilt3Value);
                    }
                    await page.waitForTimeout(1000); // wait for 1 seconds
                }
                {
                    const targetPage = page;
                    await scrollIntoViewIfNeeded([
                        [
                            `#j_id0\\:j_id2\\:frm\\:j_id${formv}\\:6\\:tsrf`
                        ],
                        [
                            `xpath///*[@id="j_id0:j_id2:frm:j_id${formv}:6:tsrf"]`
                        ],
                        [
                            `pierce/#j_id0\\:j_id2\\:frm\\:j_id${formv}\\:6\\:tsrf`
                        ]
                    ], targetPage, timeout);
                    const element = await waitForSelectors([
                        [
                            `#j_id0\\:j_id2\\:frm\\:j_id${formv}\\:6\\:tsrf`
                        ],
                        [
                            `xpath///*[@id="j_id0:j_id2:frm:j_id${formv}:6:tsrf"]`
                        ],
                        [
                            `pierce/#j_id0\\:j_id2\\:frm\\:j_id${formv}\\:6\\:tsrf`
                        ]
                    ], targetPage, { timeout, visible: true });
                    await element.click({
                        delay: 530.6999999880791,
                        offset: {
                            x: 1,
                            y: 20.515625,
                        },
                    });
                    await page.waitForTimeout(1000); // wait for 1 seconds
                }
                {
                    const targetPage = page;
                    await scrollIntoViewIfNeeded([
                        [
                            `#j_id0\\:j_id2\\:frm\\:j_id${formv}\\:6\\:tsrf`
                        ],
                        [
                            `xpath///*[@id="j_id0:j_id2:frm:j_id${formv}:6:tsrf"]`
                        ],
                        [
                            `pierce/#j_id0\\:j_id2\\:frm\\:j_id${formv}\\:6\\:tsrf`
                        ]
                    ], targetPage, timeout);
                    const element = await waitForSelectors([
                        [
                            `#j_id0\\:j_id2\\:frm\\:j_id${formv}\\:6\\:tsrf`
                        ],
                        [
                            `xpath///*[@id="j_id0:j_id2:frm:j_id${formv}:6:tsrf"]`
                        ],
                        [
                            `pierce/#j_id0\\:j_id2\\:frm\\:j_id${formv}\\:6\\:tsrf`
                        ]
                    ], targetPage, { timeout, visible: true });
                    const inputType = await element.evaluate(el => el.type);
                    // Get value from Google Spreadsheet from chatgpt
                    const Tsrf3Value = datarow.Tsrf[6]; // adjust according to the actual column header in the spreadsheet
                    if (inputType === 'select-one') {
                        await changeSelectElement(element, Tsrf3Value)
                    } else if ([
                        'textarea',
                        'text',
                        'url',
                        'tel',
                        'search',
                        'password',
                        'number',
                        'email'
                    ].includes(inputType)) {
                        await typeIntoElement(element, Tsrf3Value);
                    } else {
                        await changeElementValue(element, Tsrf3Value);
                    }
                    await page.waitForTimeout(1000); // wait for 1 seconds
                }
            }
        }
        ////////////////////////////////////inverter information starts below////////////////////////////////////////////////////////////////////////////////////
        {
            // Extract dynamic part from the id that begins with 'j_id0:j_id2:frm:j_id'
            // Extract dynamic part from the id that begins with 'j_id0:j_id2:frm:j_id'
            // Extract dynamic part from the id that begins with 'j_id0:j_id2:frm:j_id'
            const elementHandle = await page.$('[id^="select2-j_id0j_id2frmj_id"]');
            const id = await page.evaluate(element => element.id, elementHandle);
            let match = id.match(/select2-j_id0j_id2frmj_id(\d{3})/);
            let invselector = match ? match[1] : null;
            console.log(`Constructed selectors: #j_id0\\:j_id2\\:frm\\:j_id${invselector}\\:0\\:qty2`);
            {
                const targetPage = page;
                await scrollIntoViewIfNeeded([
                    [
                        `#j_id0\\:j_id2\\:frm\\:j_id${invselector}\\:0\\:qty2`
                    ],
                    [
                        `xpath///*[@id="j_id0:j_id2:frm:j_id${invselector}:0:qty2"]`
                    ],
                    [
                        `pierce/#j_id0\\:j_id2\\:frm\\:j_id${invselector}\\:0\\:qty2`
                    ]
                ], targetPage, timeout);
                const element = await waitForSelectors([
                    [
                        `#j_id0\\:j_id2\\:frm\\:j_id${invselector}\\:0\\:qty2`
                    ],
                    [
                        `xpath///*[@id="j_id0:j_id2:frm:j_id${invselector}:0:qty2"]`
                    ],
                    [
                        `pierce/#j_id0\\:j_id2\\:frm\\:j_id${invselector}\\:0\\:qty2`
                    ]
                ], targetPage, { timeout, visible: true });
                await element.click({
                    offset: {
                        x: 51,
                        y: 23.015625,
                    },
                });
                await page.waitForTimeout(1000); // wait for 1 seconds
            }
            {
                const targetPage = page;
                await scrollIntoViewIfNeeded([
                    [
                        `#j_id0\\:j_id2\\:frm\\:j_id${invselector}\\:0\\:qty2`
                    ],
                    [
                        `xpath///*[@id="j_id0:j_id2:frm:j_id${invselector}:0:qty2"]`
                    ],
                    [
                        `pierce/#j_id0\\:j_id2\\:frm\\:j_id${invselector}\\:0\\:qty2`
                    ]
                ], targetPage, timeout);
                const element = await waitForSelectors([
                    [
                        `#j_id0\\:j_id2\\:frm\\:j_id${invselector}\\:0\\:qty2`
                    ],
                    [
                        `xpath///*[@id="j_id0:j_id2:frm:j_id${invselector}:0:qty2"]`
                    ],
                    [
                        `pierce/#j_id0\\:j_id2\\:frm\\:j_id${invselector}\\:0\\:qty2`
                    ]
                ], targetPage, { timeout, visible: true });
                const inputType = await element.evaluate(el => el.type);
                // Get value from Google Spreadsheet from chatgpt
                const Panel4Value = Panel_count.toString(); // adjust according to the actual column header in the spreadsheet
                if (inputType === 'select-one') {
                    await changeSelectElement(element, Panel4Value)
                } else if ([
                    'textarea',
                    'text',
                    'url',
                    'tel',
                    'search',
                    'password',
                    'number',
                    'email'
                ].includes(inputType)) {
                    await typeIntoElement(element, Panel4Value);
                } else {
                    await changeElementValue(element, Panel4Value);
                }
                await page.waitForTimeout(2000); // wait for 1 seconds
            }
            {
                const targetPage = page;
                await scrollIntoViewIfNeeded([
                    [
                        `#select2-j_id0j_id2frmj_id${invselector}0module2-container`
                    ],
                    [
                        `xpath///*[@id="select2-j_id0j_id2frmj_id${invselector}0module2-container"]`
                    ],
                    [
                        `pierce/#select2-j_id0j_id2frmj_id${invselector}0module2-container`
                    ]
                ], targetPage, timeout);
                const element = await waitForSelectors([
                    [
                        `#select2-j_id0j_id2frmj_id${invselector}0module2-container`
                    ],
                    [
                        `xpath///*[@id="select2-j_id0j_id2frmj_id${invselector}0module2-container"]`
                    ],
                    [
                        `pierce/#select2-j_id0j_id2frmj_id${invselector}0module2-container`
                    ]
                ], targetPage, { timeout, visible: true });
                await element.click({
                    offset: {
                        x: 86,
                        y: 12.015625,
                    },
                });
                await page.waitForTimeout(1000); // wait for 1 seconds
            }
            {
                const targetPage = page;
                await scrollIntoViewIfNeeded([
                    [
                        'body > span input'
                    ],
                    [
                        'xpath//html/body/span/span/span[1]/input'
                    ],
                    [
                        'pierce/body > span input'
                    ]
                ], targetPage, timeout);
                const element = await waitForSelectors([
                    [
                        'body > span input'
                    ],
                    [
                        'xpath//html/body/span/span/span[1]/input'
                    ],
                    [
                        'pierce/body > span input'
                    ]
                ], targetPage, { timeout, visible: true });
                const inputType = await element.evaluate(el => el.type);
                if (inputType === 'select-one') {
                    await changeSelectElement(element, 'enphase')
                } else if ([
                    'textarea',
                    'text',
                    'url',
                    'tel',
                    'search',
                    'password',
                    'number',
                    'email'
                ].includes(inputType)) {
                    await typeIntoElement(element, 'enphase energy inc.');
                } else {
                    await changeElementValue(element, 'enphase energy inc.');
                }
            }
            {
                const targetPage = page;
                await targetPage.keyboard.down('Enter');
            }
            {
                const targetPage = page;
                await targetPage.keyboard.up('Enter');
            }
            {
                const targetPage = page;
                await scrollIntoViewIfNeeded([
                    [
                        `#j_id0\\:j_id2\\:frm\\:j_id${invselector}\\:0\\:model2`
                    ],
                    [
                        `xpath///*[@id="j_id0:j_id2:frm:j_id${invselector}:0:model2"]`
                    ],
                    [
                        `pierce/#j_id0\\:j_id2\\:frm\\:j_id${invselector}\\:0\\:model2`
                    ],
                    [
                        `aria/Inverter Model`
                    ]
                ], targetPage, timeout);
                const element = await waitForSelectors([
                    [
                        `#j_id0\\:j_id2\\:frm\\:j_id${invselector}\\:0\\:model2`
                    ],
                    [
                        `xpath///*[@id="j_id0:j_id2:frm:j_id${invselector}:0:model2"]`
                    ],
                    [
                        `pierce/#j_id0\\:j_id2\\:frm\\:j_id${invselector}\\:0\\:model2`
                    ],
                    [
                        `aria/Inverter Model`
                    ]
                ], targetPage, { timeout, visible: true });
                await element.click({
                    offset: {
                        x: 114,
                        y: 10.515625,
                    },
                });
                await page.waitForTimeout(1000); // wait for 1 seconds
            }
            {
                const targetPage = page;
                await scrollIntoViewIfNeeded([
                    [
                        `#j_id0\\:j_id2\\:frm\\:j_id${invselector}\\:0\\:model2`
                    ],
                    [
                        `xpath///*[@id="j_id0:j_id2:frm:j_id${invselector}:0:model2"]`
                    ],
                    [
                        `pierce/#j_id0\\:j_id2\\:frm\\:j_id${invselector}\\:0\\:model2`
                    ],
                    [
                        `aria/Inverter Model`
                    ]
                ], targetPage, timeout);
                const element = await waitForSelectors([
                    [
                        `#j_id0\\:j_id2\\:frm\\:j_id${invselector}\\:0\\:model2`
                    ],
                    [
                        `xpath///*[@id="j_id0:j_id2:frm:j_id${invselector}:0:model2"]`
                    ],
                    [
                        `pierce/#j_id0\\:j_id2\\:frm\\:j_id${invselector}\\:0\\:model2`
                    ],
                    [
                        `aria/Inverter Model`
                    ]
                ], targetPage, { timeout, visible: true });
                const inputType = await element.evaluate(el => el.type);
                if (inputType === 'select-one') {
                    await changeSelectElement(element, 'a1gt0000004na0aAAA')
                } else if ([
                    'textarea',
                    'text',
                    'url',
                    'tel',
                    'search',
                    'password',
                    'number',
                    'email'
                ].includes(inputType)) {
                    await typeIntoElement(element, 'a1gt0000004na0aAAA');
                } else {
                    await changeElementValue(element, 'a1gt0000004na0aAAA');
                }
                await page.waitForTimeout(1000); // wait for 1 seconds
            }
        }
        ////////////////Cost Distribution starts here////////////////////////////////////////////////////////////////////////////////////////////////////////////////
        {
            const targetPage = page;
            await scrollIntoViewIfNeeded([
                [
                    '#j_id0\\:j_id2\\:frm\\:labour'
                ],
                [
                    'xpath///*[@id="j_id0:j_id2:frm:labour"]'
                ],
                [
                    'pierce/#j_id0\\:j_id2\\:frm\\:labour'
                ],
                [
                    'aria/Labor Cost s($)'
                ]
            ], targetPage, timeout);
            const element = await waitForSelectors([
                [
                    '#j_id0\\:j_id2\\:frm\\:labour'
                ],
                [
                    'xpath///*[@id="j_id0:j_id2:frm:labour"]'
                ],
                [
                    'pierce/#j_id0\\:j_id2\\:frm\\:labour'
                ],
                [
                    'aria/Labor Cost s($)'
                ]
            ], targetPage, { timeout, visible: true });
            await element.click({
                offset: {
                    x: 135,
                    y: 12.015625,
                },
            });
            await page.waitForTimeout(1000); // wait for 1 seconds
        }
        {
            const targetPage = page;
            await scrollIntoViewIfNeeded([
                [
                    '#j_id0\\:j_id2\\:frm\\:labour'
                ],
                [
                    'xpath///*[@id="j_id0:j_id2:frm:labour"]'
                ],
                [
                    'pierce/#j_id0\\:j_id2\\:frm\\:labour'
                ],
                [
                    'aria/Labor Cost s($)'
                ]
            ], targetPage, timeout);
            const element = await waitForSelectors([
                [
                    '#j_id0\\:j_id2\\:frm\\:labour'
                ],
                [
                    'xpath///*[@id="j_id0:j_id2:frm:labour"]'
                ],
                [
                    'pierce/#j_id0\\:j_id2\\:frm\\:labour'
                ],
                [
                    'aria/Labor Cost s($)'
                ]
            ], targetPage, { timeout, visible: true });
            const inputType = await element.evaluate(el => el.type);
            // Get value from Google Spreadsheet from chatgpt
            const LaborValue = LaborValue1; // adjust according to the actual column header in the spreadsheet
            if (inputType === 'select-one') {
                await changeSelectElement(element, LaborValue)
            } else if ([
                'textarea',
                'text',
                'url',
                'tel',
                'search',
                'password',
                'number',
                'email'
            ].includes(inputType)) {
                await typeIntoElement(element, LaborValue);
            } else {
                await changeElementValue(element, LaborValue);
            }
            await page.waitForTimeout(1000); // wait for 1 seconds
        }
        {
            const targetPage = page;
            await scrollIntoViewIfNeeded([
                [
                    '#j_id0\\:j_id2\\:frm\\:balanceSystem'
                ],
                [
                    'xpath///*[@id="j_id0:j_id2:frm:balanceSystem"]'
                ],
                [
                    'pierce/#j_id0\\:j_id2\\:frm\\:balanceSystem'
                ],
                [
                    'aria/Balance of System Cost ($)'
                ]
            ], targetPage, timeout);
            const element = await waitForSelectors([
                [
                    '#j_id0\\:j_id2\\:frm\\:balanceSystem'
                ],
                [
                    'xpath///*[@id="j_id0:j_id2:frm:balanceSystem"]'
                ],
                [
                    'pierce/#j_id0\\:j_id2\\:frm\\:balanceSystem'
                ],
                [
                    'aria/Balance of System Cost ($)'
                ]
            ], targetPage, { timeout, visible: true });
            await element.click({
                offset: {
                    x: 115,
                    y: 19.015625,
                },
            });
            await page.waitForTimeout(1000); // wait for 1 seconds
        }
        {
            const targetPage = page;
            await scrollIntoViewIfNeeded([
                [
                    '#j_id0\\:j_id2\\:frm\\:balanceSystem'
                ],
                [
                    'xpath///*[@id="j_id0:j_id2:frm:balanceSystem"]'
                ],
                [
                    'pierce/#j_id0\\:j_id2\\:frm\\:balanceSystem'
                ],
                [
                    'aria/Balance of System Cost ($)'
                ]
            ], targetPage, timeout);
            const element = await waitForSelectors([
                [
                    '#j_id0\\:j_id2\\:frm\\:balanceSystem'
                ],
                [
                    'xpath///*[@id="j_id0:j_id2:frm:balanceSystem"]'
                ],
                [
                    'pierce/#j_id0\\:j_id2\\:frm\\:balanceSystem'
                ],
                [
                    'aria/Balance of System Cost ($)'
                ]
            ], targetPage, { timeout, visible: true });
            const inputType = await element.evaluate(el => el.type);
            // Get value from Google Spreadsheet from chatgpt
            const BOSValue = BOSValue1; // adjust according to the actual column header in the spreadsheet
            if (inputType === 'select-one') {
                await changeSelectElement(element, BOSValue)
            } else if ([
                'textarea',
                'text',
                'url',
                'tel',
                'search',
                'password',
                'number',
                'email'
            ].includes(inputType)) {
                await typeIntoElement(element, BOSValue);
            } else {
                await changeElementValue(element, BOSValue);
            }
            await page.waitForTimeout(1000); // wait for 1 seconds
        }
        {
            const targetPage = page;
            await scrollIntoViewIfNeeded([
                [
                    '#j_id0\\:j_id2\\:frm\\:permittingLoss'
                ],
                [
                    'xpath///*[@id="j_id0:j_id2:frm:permittingLoss"]'
                ],
                [
                    'pierce/#j_id0\\:j_id2\\:frm\\:permittingLoss'
                ],
                [
                    'aria/Permitting Fees ($)'
                ]
            ], targetPage, timeout);
            const element = await waitForSelectors([
                [
                    '#j_id0\\:j_id2\\:frm\\:permittingLoss'
                ],
                [
                    'xpath///*[@id="j_id0:j_id2:frm:permittingLoss"]'
                ],
                [
                    'pierce/#j_id0\\:j_id2\\:frm\\:permittingLoss'
                ],
                [
                    'aria/Permitting Fees ($)'
                ]
            ], targetPage, { timeout, visible: true });
            await element.click({
                offset: {
                    x: 102,
                    y: 20.515625,
                },
            });
            await page.waitForTimeout(1000); // wait for 1 seconds
        }
        {
            const targetPage = page;
            await scrollIntoViewIfNeeded([
                [
                    '#j_id0\\:j_id2\\:frm\\:permittingLoss'
                ],
                [
                    'xpath///*[@id="j_id0:j_id2:frm:permittingLoss"]'
                ],
                [
                    'pierce/#j_id0\\:j_id2\\:frm\\:permittingLoss'
                ],
                [
                    'aria/Permitting Fees ($)'
                ]
            ], targetPage, timeout);
            const element = await waitForSelectors([
                [
                    '#j_id0\\:j_id2\\:frm\\:permittingLoss'
                ],
                [
                    'xpath///*[@id="j_id0:j_id2:frm:permittingLoss"]'
                ],
                [
                    'pierce/#j_id0\\:j_id2\\:frm\\:permittingLoss'
                ],
                [
                    'aria/Permitting Fees ($)'
                ]
            ], targetPage, { timeout, visible: true });
            const inputType = await element.evaluate(el => el.type);
            // Get value from Google Spreadsheet from chatgpt
            const PermitCValue = PermitCValue1; // adjust according to the actual column header in the spreadsheet
            if (inputType === 'select-one') {
                await changeSelectElement(element, PermitCValue)
            } else if ([
                'textarea',
                'text',
                'url',
                'tel',
                'search',
                'password',
                'number',
                'email'
            ].includes(inputType)) {
                await typeIntoElement(element, PermitCValue);
            } else {
                await changeElementValue(element, PermitCValue);
            }
            await page.waitForTimeout(1000); // wait for 1 seconds
        }
        {
            const targetPage = page;
            await scrollIntoViewIfNeeded([
                [
                    '#j_id0\\:j_id2\\:frm\\:inspectionCosts'
                ],
                [
                    'xpath///*[@id="j_id0:j_id2:frm:inspectionCosts"]'
                ],
                [
                    'pierce/#j_id0\\:j_id2\\:frm\\:inspectionCosts'
                ],
                [
                    'aria/Inspection Cost ($)'
                ]
            ], targetPage, timeout);
            const element = await waitForSelectors([
                [
                    '#j_id0\\:j_id2\\:frm\\:inspectionCosts'
                ],
                [
                    'xpath///*[@id="j_id0:j_id2:frm:inspectionCosts"]'
                ],
                [
                    'pierce/#j_id0\\:j_id2\\:frm\\:inspectionCosts'
                ],
                [
                    'aria/Inspection Cost ($)'
                ]
            ], targetPage, { timeout, visible: true });
            await element.click({
                offset: {
                    x: 149,
                    y: 19.515625,
                },
            });
            await page.waitForTimeout(1000); // wait for 1 seconds
        }
        {
            const targetPage = page;
            await scrollIntoViewIfNeeded([
                [
                    '#j_id0\\:j_id2\\:frm\\:inspectionCosts'
                ],
                [
                    'xpath///*[@id="j_id0:j_id2:frm:inspectionCosts"]'
                ],
                [
                    'pierce/#j_id0\\:j_id2\\:frm\\:inspectionCosts'
                ],
                [
                    'aria/Inspection Cost ($)'
                ]
            ], targetPage, timeout);
            const element = await waitForSelectors([
                [
                    '#j_id0\\:j_id2\\:frm\\:inspectionCosts'
                ],
                [
                    'xpath///*[@id="j_id0:j_id2:frm:inspectionCosts"]'
                ],
                [
                    'pierce/#j_id0\\:j_id2\\:frm\\:inspectionCosts'
                ],
                [
                    'aria/Inspection Cost ($)'
                ]
            ], targetPage, { timeout, visible: true });
            const inputType = await element.evaluate(el => el.type);
            // Get value from Google Spreadsheet from chatgpt
            const InspectionCValue = InspectionCValue1; // adjust according to the actual column header in the spreadsheet
            if (inputType === 'select-one') {
                await changeSelectElement(element, InspectionCValue)
            } else if ([
                'textarea',
                'text',
                'url',
                'tel',
                'search',
                'password',
                'number',
                'email'
            ].includes(inputType)) {
                await typeIntoElement(element, InspectionCValue);
            } else {
                await changeElementValue(element, InspectionCValue);
            }
            await page.waitForTimeout(1000); // wait for 1 seconds
        }
        {
            const targetPage = page;
            await scrollIntoViewIfNeeded([
                [
                    '#j_id0\\:j_id2\\:frm\\:arrayCosts'
                ],
                [
                    'xpath///*[@id="j_id0:j_id2:frm:arrayCosts"]'
                ],
                [
                    'pierce/#j_id0\\:j_id2\\:frm\\:arrayCosts'
                ],
                [
                    'aria/Array Cost ($)'
                ]
            ], targetPage, timeout);
            const element = await waitForSelectors([
                [
                    '#j_id0\\:j_id2\\:frm\\:arrayCosts'
                ],
                [
                    'xpath///*[@id="j_id0:j_id2:frm:arrayCosts"]'
                ],
                [
                    'pierce/#j_id0\\:j_id2\\:frm\\:arrayCosts'
                ],
                [
                    'aria/Array Cost ($)'
                ]
            ], targetPage, { timeout, visible: true });
            await element.click({
                offset: {
                    x: 129,
                    y: 6.015625,
                },
            });
        }
        {
            const targetPage = page;
            await scrollIntoViewIfNeeded([
                [
                    '#j_id0\\:j_id2\\:frm\\:arrayCosts'
                ],
                [
                    'xpath///*[@id="j_id0:j_id2:frm:arrayCosts"]'
                ],
                [
                    'pierce/#j_id0\\:j_id2\\:frm\\:arrayCosts'
                ],
                [
                    'aria/Array Cost ($)'
                ]
            ], targetPage, timeout);
            const element = await waitForSelectors([
                [
                    '#j_id0\\:j_id2\\:frm\\:arrayCosts'
                ],
                [
                    'xpath///*[@id="j_id0:j_id2:frm:arrayCosts"]'
                ],
                [
                    'pierce/#j_id0\\:j_id2\\:frm\\:arrayCosts'
                ],
                [
                    'aria/Array Cost ($)'
                ]
            ], targetPage, { timeout, visible: true });
            const inputType = await element.evaluate(el => el.type);
            // Get value from Google Spreadsheet from chatgpt
            const ArrayCValue = ArrayCValue1; // adjust according to the actual column header in the spreadsheet
            if (inputType === 'select-one') {
                await changeSelectElement(element, ArrayCValue)
            } else if ([
                'textarea',
                'text',
                'url',
                'tel',
                'search',
                'password',
                'number',
                'email'
            ].includes(inputType)) {
                await typeIntoElement(element, ArrayCValue);
            } else {
                await changeElementValue(element, ArrayCValue);
            }
            await page.waitForTimeout(1000); // wait for 1 seconds
        }
        {
            const targetPage = page;
            await scrollIntoViewIfNeeded([
                [
                    '#j_id0\\:j_id2\\:frm\\:inverterCosts'
                ],
                [
                    'xpath///*[@id="j_id0:j_id2:frm:inverterCosts"]'
                ],
                [
                    'pierce/#j_id0\\:j_id2\\:frm\\:inverterCosts'
                ],
                [
                    'aria/Inverter Cost ($)'
                ]
            ], targetPage, timeout);
            const element = await waitForSelectors([
                [
                    '#j_id0\\:j_id2\\:frm\\:inverterCosts'
                ],
                [
                    'xpath///*[@id="j_id0:j_id2:frm:inverterCosts"]'
                ],
                [
                    'pierce/#j_id0\\:j_id2\\:frm\\:inverterCosts'
                ],
                [
                    'aria/Inverter Cost ($)'
                ]
            ], targetPage, { timeout, visible: true });
            await element.click({
                offset: {
                    x: 68,
                    y: 30.015625,
                },
            });
            await page.waitForTimeout(1000); // wait for 1 seconds
        }
        {
            const targetPage = page;
            await scrollIntoViewIfNeeded([
                [
                    '#j_id0\\:j_id2\\:frm\\:inverterCosts'
                ],
                [
                    'xpath///*[@id="j_id0:j_id2:frm:inverterCosts"]'
                ],
                [
                    'pierce/#j_id0\\:j_id2\\:frm\\:inverterCosts'
                ],
                [
                    'aria/Inverter Cost ($)'
                ]
            ], targetPage, timeout);
            const element = await waitForSelectors([
                [
                    '#j_id0\\:j_id2\\:frm\\:inverterCosts'
                ],
                [
                    'xpath///*[@id="j_id0:j_id2:frm:inverterCosts"]'
                ],
                [
                    'pierce/#j_id0\\:j_id2\\:frm\\:inverterCosts'
                ],
                [
                    'aria/Inverter Cost ($)'
                ]
            ], targetPage, { timeout, visible: true });
            const inputType = await element.evaluate(el => el.type);
            // Get value from Google Spreadsheet from chatgpt
            const InvCValue = InvCValue1; // adjust according to the actual column header in the spreadsheet
            if (inputType === 'select-one') {
                await changeSelectElement(element, InvCValue)
            } else if ([
                'textarea',
                'text',
                'url',
                'tel',
                'search',
                'password',
                'number',
                'email'
            ].includes(inputType)) {
                await typeIntoElement(element, InvCValue);
            } else {
                await changeElementValue(element, InvCValue);
            }
            await page.waitForTimeout(1000); // wait for 1 seconds
        }
        {
            const targetPage = page;
            await scrollIntoViewIfNeeded([
                [
                    '#j_id0\\:j_id2\\:frm\\:histAnnualEnergy'
                ],
                [
                    'xpath///*[@id="j_id0:j_id2:frm:histAnnualEnergy"]'
                ],
                [
                    'pierce/#j_id0\\:j_id2\\:frm\\:histAnnualEnergy'
                ],
                [
                    'aria/Historical Annual Energy Consumption (kWh/year) *'
                ]
            ], targetPage, timeout);
            const element = await waitForSelectors([
                [
                    '#j_id0\\:j_id2\\:frm\\:histAnnualEnergy'
                ],
                [
                    'xpath///*[@id="j_id0:j_id2:frm:histAnnualEnergy"]'
                ],
                [
                    'pierce/#j_id0\\:j_id2\\:frm\\:histAnnualEnergy'
                ],
                [
                    'aria/Historical Annual Energy Consumption (kWh/year) *'
                ]
            ], targetPage, { timeout, visible: true });
            const inputType = await element.evaluate(el => el.type);
            // Get value from Google Spreadsheet from chatgpt
            const AnnualEValue = AnnualEValue1; // adjust according to the actual column header in the spreadsheet
            if (inputType === 'select-one') {
                await changeSelectElement(element, AnnualEValue)
            } else if ([
                'textarea',
                'text',
                'url',
                'tel',
                'search',
                'password',
                'number',
                'email'
            ].includes(inputType)) {
                await typeIntoElement(element, AnnualEValue);
            } else {
                await changeElementValue(element, AnnualEValue);
            }
            await page.waitForTimeout(1000); // wait for 1 seconds
        }
        {
            const timeout = 10000;
            const targetPage = page;
            const promises = [];
            promises.push(targetPage.waitForNavigation());
            await scrollIntoViewIfNeeded([
                'footer a.btn-primary',
                'xpath///*[@id="j_id0:j_id2:frm"]/footer/div/a[3]',
                'pierce/footer a.btn-primary',
                'aria/CONTINUE'
            ], targetPage, timeout);
            const element = await waitForSelectors([
                'footer a.btn-primary',
                'xpath///*[@id="j_id0:j_id2:frm"]/footer/div/a[3]',
                'pierce/footer a.btn-primary',
                'aria/CONTINUE'
            ], targetPage, { timeout, visible: true });
            await element.click({
                offset: {
                    x: 59.34375,
                    y: 14.015625,
                },
            });
            await Promise.all(promises);
            //await page.waitForTimeout(1000); // wait for 1 seconds
            //await page.waitForNavigation({ waitUntil: 'networkidle0' });

        }
        {
            const targetPage = page;
            await scrollIntoViewIfNeeded([
                [
                    '#j_id0\\:j_id2\\:frm\\:pType'
                ],
                [
                    'xpath///*[@id="j_id0:j_id2:frm:pType"]'
                ],
                [
                    'pierce/#j_id0\\:j_id2\\:frm\\:pType'
                ],
                [
                    'aria/Purchase Type*'
                ]
            ], targetPage, timeout);
            const element = await waitForSelectors([
                [
                    '#j_id0\\:j_id2\\:frm\\:pType'
                ],
                [
                    'xpath///*[@id="j_id0:j_id2:frm:pType"]'
                ],
                [
                    'pierce/#j_id0\\:j_id2\\:frm\\:pType'
                ],
                [
                    'aria/Purchase Type*'
                ]
            ], targetPage, { timeout, visible: true });
            await element.click({
                offset: {
                    x: 67,
                    y: 23.015625,
                },
            });
        }
        {
            const targetPage = page;
            await scrollIntoViewIfNeeded([
                [
                    '#j_id0\\:j_id2\\:frm\\:pType'
                ],
                [
                    'xpath///*[@id="j_id0:j_id2:frm:pType"]'
                ],
                [
                    'pierce/#j_id0\\:j_id2\\:frm\\:pType'
                ],
                [
                    'aria/Purchase Type*'
                ]
            ], targetPage, timeout);
            const element = await waitForSelectors([
                [
                    '#j_id0\\:j_id2\\:frm\\:pType'
                ],
                [
                    'xpath///*[@id="j_id0:j_id2:frm:pType"]'
                ],
                [
                    'pierce/#j_id0\\:j_id2\\:frm\\:pType'
                ],
                [
                    'aria/Purchase Type*'
                ]
            ], targetPage, { timeout, visible: true });
            const inputType = await element.evaluate(el => el.type);
            // Get value from Google Spreadsheet from chatgpt
            const PurcharseTValue = PurchaseTValue1; // adjust according to the actual column header in the spreadsheet
            if (inputType === 'select-one') {
                await changeSelectElement(element, PurcharseTValue)
            } else if ([
                'textarea',
                'text',
                'url',
                'tel',
                'search',
                'password',
                'number',
                'email'
            ].includes(inputType)) {
                await typeIntoElement(element, PurcharseTValue);
            } else {
                await changeElementValue(element, PurcharseTValue);
            }
            await page.waitForTimeout(1000); // wait for 1 seconds
        }
        // Assuming 'GIGNY' is the column header for the checkbox column in your spreadsheet
        if (GIGNY === true) { // replace 'TRUE' with the value that represents a checked checkbox in your spreadsheet
            const targetPage = page;

            await scrollIntoViewIfNeeded([
                [
                    '#j_id0\\:j_id2\\:frm\\:GJGNYSectorValue'
                ],
                [
                    'xpath///*[@id="j_id0:j_id2:frm:GJGNYSectorValue"]'
                ],
                [
                    'pierce/#j_id0\\:j_id2\\:frm\\:GJGNYSectorValue'
                ],
                [
                    'aria/Will the Project use GJGNY Financing?   '
                ]
            ], targetPage, timeout);

            const element = await waitForSelectors([
                [
                    '#j_id0\\:j_id2\\:frm\\:GJGNYSectorValue'
                ],
                [
                    'xpath///*[@id="j_id0:j_id2:frm:GJGNYSectorValue"]'
                ],
                [
                    'pierce/#j_id0\\:j_id2\\:frm\\:GJGNYSectorValue'
                ],
                [
                    'aria/Will the Project use GJGNY Financing?   '
                ]
            ], targetPage, { timeout, visible: true });

            await element.click({
                offset: {
                    x: 1.015625,
                    y: 4.015625,
                },
            });
            await page.waitForTimeout(3000); // wait for 1 seconds
        }
        /////////////////////////////////CHECKBOX ABOVE///////////////////////////////////////////////////////////////////////////////////////////
        {
            //const elementHandleArray = await page.$x('//*[starts-with(@id, "j_id0:j_id2:j_id")]');
            //const elementHandle = elementHandleArray[0]; // If you are sure that the element exists
            //const id = await page.evaluate(element => element.id, elementHandle);
            //let match = id.match(/j_id0:j_id2:j_id(\d+)/);
            //let invselector2 = match ? match[1] : null;
            const targetPage = page;
            const promises = [];
            promises.push(targetPage.waitForNavigation());
           // console.log(`Constructed selectors: *[@id="j_id0:j_id2:j_id${invselector2}"]/a[2]`);
            await scrollIntoViewIfNeeded([
                'a.btn-primary',
               // `xpath///*[@id="j_id0:j_id2:j_id${invselector2}"]/a[2]`,
                'pierce/a.btn-primary',
                'aria/CONTINUE'
            ], targetPage, timeout);
            const element = await waitForSelectors([
                'a.btn-primary',
                //`xpath///*[@id="j_id0:j_id2:j_id${invselector2}"]/a[2]`,
                'pierce/a.btn-primary',
                'aria/CONTINUE'
            ], targetPage, { timeout, visible: true });
            await element.click({
                offset: {
                    x: 90.84375,
                    y: 16.015625,
                },
            });
            await Promise.all(promises);
            await page.waitForTimeout(5000); // wait for 1 seconds
        }
        {
            const targetPage = page;
            const buttonHandles = await targetPage.$$('input[type="submit"][id^="j_id0:j_id1:j_id"]');
            if (buttonHandles.length === 0) {
                throw new Error("No button found with id starting with 'j_id0:j_id1:j_id'");
            }
            // Get id of the first button
            const id = await targetPage.evaluate(button => button.id, buttonHandles[0]);
            const match = id.match(/j_id0:j_id1:j_id(\d+):j_id(\d+)/);
            const continueselector = match ? match[1] : null;
            const continueselector2 = match ? match[2] : null;
            console.log(`Constructed selectors:`, continueselector);
            console.log(`Constructed selectors:`, continueselector2);
          
                const promises = [];
                promises.push(targetPage.waitForNavigation());
                await scrollIntoViewIfNeeded([
                    [
                        `#j_id0\\:j_id1\\:j_id${continueselector}\\:j_id${continueselector2}`
                    ],
                    [
                        `xpath///*[@id="j_id0:j_id1:j_id${continueselector}:j_id${continueselector2}"]`
                    ],
                    [
                        `pierce/#j_id0\\:j_id1\\:j_id${continueselector}\\:j_id${continueselector2}`
                    ],
                    [
                        `aria/Continue`
                    ]
                ], targetPage, timeout);
                const element = await waitForSelectors([
                    [
                        `#j_id0\\:j_id1\\:j_id${continueselector}\\:j_id${continueselector2}`
                    ],
                    [
                        `xpath///*[@id="j_id0:j_id1:j_id${continueselector}:j_id${continueselector2}"]`
                    ],
                    [
                        `pierce/#j_id0\\:j_id1\\:j_id${continueselector}\\:j_id${continueselector2}`
                    ],
                    [
                        `aria/Continue`
                    ]
                ], targetPage, { timeout, visible: true });
                await element.click({
                    offset: {
                        x: 65.34375,
                        y: 14.015625,
                    },
                });
                    await Promise.all(promises);

                
                await page.waitForTimeout(1000); // wait for 1 seconds
            
        }
        //chatgpt start////////////////////////////////////////////////////////////////////////////////////////////////////////////////////
        const axios = require('axios');
        const fs = require('fs');
        const path = require('path');
        const url = require('url');  // import the url module
        const firstName = Firstname;
        const lastName = Lastname;
        const targetPage = page;
        const newPagePromise = new Promise(x => targetPage.browser().once('targetcreated', target => x(target.page())));
        await scrollIntoViewIfNeeded([
            'a.btn-default',
            'xpath///*[@id="j_id0:j_id2:j_id166:btnPanel"]/a[3]',
            'pierce/a.btn-default',
            'aria/PRINT',
            'text/Print'
        ], targetPage, timeout);
        const element = await waitForSelectors([
            'a.btn-default',
            'xpath///*[@id="j_id0:j_id2:j_id166:btnPanel"]/a[3]',
            'pierce/a.btn-default',
            'aria/PRINT',
            'text/Print'
        ], targetPage, { timeout, visible: true });
        await element.click({
            offset: {
                x: 45.609375,
                y: 24.515625,
            },
        });
        const newPage = await newPagePromise;
        await newPage.waitForNavigation({ waitUntil: 'networkidle0', timeout: 10000 }).catch(() => { });
        function determineNyserdaFolderName() {
            if (NYC_LI === 'LI' && Affordable === false && finance.includes("Nyserda")) {
                return "nyserdaloanApplicationFolderId";
            } else if (NYC_LI === 'LI' && Affordable === true && finance.includes("Nyserda")) {
                return "nyserdaloanwithaffordable";
            } else if (NYC_LI === 'LI' && Affordable === true && finance.includes("Nyserda")) {
                return "nyserdawithaffordable";
            } else if (Affordable === true && finance.includes("Nyserda")) {
                return "nyserdaloanwithaffordable";
            } else if (Affordable === true) {
                return "nyserdawithaffordable";
            } else if (NYC_LI === 'NYC' && !finance.includes("Nyserda")) {
                return "nyserdaincentiveapplication";
            }
            return null;
        }
        let nyserdafolderidfield = determineNyserdaFolderName();
        if (!nyserdafolderidfield) {
            throw new Error('Could not determine the appropriate nyserda folder id field');
        }
        // Assuming caspioAuthorization() is defined and returns an access token
        const nyserdaFolderIdURL = `https://c1abd578.caspio.com/rest/v2/tables/FolderID/records?q.select=${nyserdafolderidfield}&q.where=Cust_ID=${Cust_ID}`
        let folderData = await fetchCaspioData(accessToken, nyserdaFolderIdURL);
        // Assuming the nyserdaFolderId is the first record in the response data
        let nyserdaFolderId;
        if (folderData.Result && folderData.Result.length > 0) {
            nyserdaFolderId = folderData.Result[0][nyserdafolderidfield];
        }
        if (!nyserdaFolderId) {
            throw new Error('Could not find the target folder');
        }
        await processPDF(page, firstName, lastName, client, nyserdaFolderId).catch(console.error);
        // Close the new tab
        await newPage.close();
        await page.waitForTimeout(5000); // wait for 1 seconds
		
		/*
        if (finance.toLowerCase().includes("nyserda")) {
            const pro2 = spawn('node', ['./pro2.js', nyserdaFolderId, Cust_ID]);

            pro2.stdout.on('data', (data) => {
                console.log(`stdout: ${data}`);
            });

            pro2.stderr.on('data', (data) => {
                console.error(`stderr: ${data}`);
            });

            pro2.on('close', (code) => {
                console.log(`child process exited with code ${code}`);
            });
        }
		*/
		
        //chatgpt start////////////////////////////////////////////////////////////////////////////////////////////////////////////////////
        {
            const targetPage = page;
            const promises = [];
            promises.push(targetPage.waitForNavigation());
            await scrollIntoViewIfNeeded([
                [
                    '#continueBtn'
                ],
                [
                    'xpath///*[@id="continueBtn"]'
                ],
                [
                    'pierce/#continueBtn'
                ],
                [
                    'aria/CONTINUE'
                ]
            ], targetPage, timeout);
            const element = await waitForSelectors([
                [
                    '#continueBtn'
                ],
                [
                    'xpath///*[@id="continueBtn"]'
                ],
                [
                    'pierce/#continueBtn'
                ],
                [
                    'aria/CONTINUE'
                ]
            ], targetPage, { timeout, visible: true });
            await element.click({
                offset: {
                    x: 104.34375,
                    y: 26.515625,
                },
            });
            await Promise.all(promises);
            await page.waitForTimeout(5000); // wait for 1 seconds
        }
        {
            const targetPage = page;
            const promises = [];
            promises.push(targetPage.waitForNavigation());
            await scrollIntoViewIfNeeded([
                'a.btn-default',
                'xpath///*[@id="j_id0:j_id2:saveBtn1"]/a[3]',
                'pierce/a.btn-default',
                'aria/SAVE',
                'text/Save'
            ], targetPage, timeout);
            const element = await waitForSelectors([
                'a.btn-default',
                'xpath///*[@id="j_id0:j_id2:saveBtn1"]/a[3]',
                'pierce/a.btn-default',
                'aria/SAVE',
                'text/Save'
            ], targetPage, { timeout, visible: true });
            await element.click({
                offset: {
                    x: 58.578125,
                    y: 18.515625,
                },
            });
            await Promise.all(promises);
            await page.waitForTimeout(5000); // wait for 23 seconds
        }
        let dataupload = {
            "Incentive_Approved": 'Prepared & Saved',
            "Submission_Required_Items": submission_requiments

        };
        await SenddatatoCaspio(accessToken, sendurl, dataupload, method = 'PUT');
        await sleep(20000);
        /* {
            const targetPage = page;
            await scrollIntoViewIfNeeded([
                [
                    'a.left'
                ],
                [
                    'xpath///*[@id="j_id0:j_id2:saveBtn1"]/a[1]'
                ],
                [
                    'pierce/a.left'
                ],
                [
                    'aria/DISCARD'
                ],
                [
                    'text/Discard'
                ]
            ], targetPage, timeout);
            const element = await waitForSelectors([
                [
                    'a.left'
                ],
                [
                    'xpath///*[@id="j_id0:j_id2:saveBtn1"]/a[1]'
                ],
                [
                    'pierce/a.left'
                ],
                [
                    'aria/DISCARD'
                ],
                [
                    'text/Discard'
                ]
            ], targetPage, { timeout, visible: true });
            await element.click({
                offset: {
                    x: 86,
                    y: 24.515625,
                },
            });
        }
        {
            const targetPage = page;
            const promises = [];
            promises.push(targetPage.waitForNavigation());
            await scrollIntoViewIfNeeded([
                [
                    '#saveBtn2'
                ],
                [
                    'xpath///*[@id="saveBtn2"]'
                ],
                [
                    'pierce/#saveBtn2'
                ],
                [
                    'aria/YES'
                ],
                [
                    'text/Yes'
                ]
            ], targetPage, timeout);
            const element = await waitForSelectors([
                [
                    '#saveBtn2'
                ],
                [
                    'xpath///*[@id="saveBtn2"]'
                ],
                [
                    'pierce/#saveBtn2'
                ],
                [
                    'aria/YES'
                ],
                [
                    'text/Yes'
                ]
            ], targetPage, { timeout, visible: true });
            await element.click({
                offset: {
                    x: 46,
                    y: 16,
                },
            });
            await Promise.all(promises);
        }
        */
    }
    await browser.close();

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
      const isInViewport = await element.isIntersectingViewport({threshold: 0});
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
        return await element.isIntersectingViewport({threshold: 0});
      }, timeout);
    }

    async function waitForSelector(selector, frame, options)  {
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
})().catch(err => {
    console.error(err);
    process.exit(1);
});
