const fs = require('fs').promises;
const axios = require('axios');
const qs = require('querystring'); 
const path = require('path');
const process = require('process');
const { authenticate } = require('@google-cloud/local-auth');
const { google } = require('googleapis');
const papaparse = require('papaparse');
let parsedCSVData = null;
//const Cust_ID = 7124;
const Cust_ID = process.argv[2];
const caspioTokenUrl = 'https://c1abd578.caspio.com/oauth/token';
const clientID = process.env.clientID;
const clientSecret = process.env.clientSecret;
const GoggleApi = process.env.GoggleApi;
const keys =JSON.parse(process.env.SERVICE_KEY);
const PDFDocument1 = require('pdfkit');
const API_KEY = 'WrEuodSc5lzfQyfHfCCRmPH790rIz8ta7pkrtFnX';
const BASE_URL = 'https://developer.nrel.gov/api/pvwatts/v6.json';
// If modifying these scopes, delete token.json.
const SCOPES = ['https://www.googleapis.com/auth/gmail.readonly'];
const TOKEN_PATH = path.join(process.cwd(), 'token.json');
const CREDENTIALS_PATH = path.join(process.cwd(), 'cred.json');
const url = `https://c1abd578.caspio.com/rest/v2/tables/SolarProcess/records?q.where=CustID=${Cust_ID}`;
// Function to load saved credentials if they exist
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
async function loadSavedCredentialsIfExist() {
    try {
        const content = await fs.readFile(TOKEN_PATH);
        const credentials = JSON.parse(content);
        return google.auth.fromJSON(credentials);
    } catch (err) {
        console.log("Error reading saved credentials:", err);
        return null;
    }
}

// Function to save credentials
async function saveCredentials(client) {
    const content = await fs.readFile(CREDENTIALS_PATH);
    const keys = JSON.parse(content);
    const key = keys.installed || keys.web;
    const payload = JSON.stringify({
        type: 'authorized_user',
        client_id: key.client_id,
        client_secret: key.client_secret,
        refresh_token: client.credentials.refresh_token,
    });
    await fs.writeFile(TOKEN_PATH, payload);
}

// Function to authorize the application
async function authorize() {
    let client = await loadSavedCredentialsIfExist();
    if (client) {
        return client;
    }
    client = await authenticate({
        scopes: SCOPES,
        keyfilePath: CREDENTIALS_PATH,
    });
    if (client.credentials) {
        await saveCredentials(client);
    }
    return client;
}

async function handleSingleMessage(auth, messageId) {
  const gmail = google.gmail({ version: 'v1', auth });
  
  let dataRow = {};

  try {
    let response = await gmail.users.messages.get({
      userId: 'me',
      id: messageId,
      format: 'full',
    });

    let headers = response.data.payload.headers;
    let dateHeader = headers.find(header => header.name === "Date").value;
    
let subjectHeader = headers.find(header => header.name === "Subject").value;
// Extracting Customer_name from the subject line
let customerNameMatch = subjectHeader.match(/for (.+?) - PAM/);
let customerName = customerNameMatch ? customerNameMatch[1] : 'Not found';
	
    let parsedDate = new Date(dateHeader);
    let formattedDate = `${parsedDate.getMonth() + 1}/${parsedDate.getDate()}/${parsedDate.getFullYear()}`;
    
    let parts = response.data.payload.parts;
    let emailBody;
    
    for(let part of parts) {
      if (part.mimeType === 'text/plain') {
        emailBody = Buffer.from(part.body.data, 'base64').toString('utf-8');
        break;
      }
    }

    if(emailBody) {
      let pamMatch = emailBody.match(/job ID # for this application is (PAM-\d+-\d+)/);
      let dcKwMatch = emailBody.match(/(\d+.\d+) kW DC/);
      let acKwMatch = emailBody.match(/(\d+.\d+) kW AC/);
      let combinedKw = `${dcKwMatch ? dcKwMatch[1] : 'Not found'} DC/${acKwMatch ? acKwMatch[1] : 'Not found'} AC`;  
      dataRow = {
        date: formattedDate,
        pamId: pamMatch ? pamMatch[1] : "Not found",
        combinedKw: combinedKw,
		customerName: customerName
      };
    } else {
      console.log("Email body not found.");
    }

  } catch (error) {
    console.log('The API returned an error: ' + error);
  }

  return dataRow;
}

async function getLatestApprovalMessage(auth,account) {
  const gmail = google.gmail({ version: 'v1', auth });

  const query = `from:(pseg-li-paminterconnect@pseg.com) to:(teamdesign@patriotenergysolution.com) subject:("${account}") subject:(approval) newer_than:12m`;

  try {
    const res = await gmail.users.messages.list({
      userId: 'me',
      q: query,
      maxResults: 1 // Fetch only the latest message
    });

    if (res.data.messages) {
      console.log(`Found ${res.data.messages.length} message(s).`);
      
      // Return the ID of the latest message
      return res.data.messages[0].id;
    } else {
      console.log('No new messages.');
      return null;
    }

  } catch (error) {
    console.log('The API returned an error: ' + error);
    return null;
  }
}

async function caspiomain(CIDvalue, accessToken) {
    try {
        const masterTableUrl = `https://c1abd578.caspio.com/rest/v2/tables/MasterTable/records?q.select=Name%2C%20Address1%2C%20Phone1%2C%20email&q.where=CID=${CIDvalue}`;
        const solarProcessUrl = `https://c1abd578.caspio.com/rest/v2/tables/SolarProcess/records?q.select=PAM%2CPAM_approvedate%2CInterconnectionName2%2CInterconnectionName3%2CRevised_panel%2CRevised_system_size%2CNumber_Of_Panels%2CInterconnectionMeter1%2CBin%2CInterconnectionMeter2%2CNYC_LI%2CInterconnectionMeter3%2CInterconnectionAcc1%2CInterconnectionAcc2%2CInterconnectionAcc3%2CInterconnectionName1%2CCB&q.where=CustID=${CIDvalue}`;
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
        const Number_Of_Panels = combinedData.Number_Of_Panels;
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

(async () => {
  const auth = await authorize();
  const accessToken = await caspioAuthorization();
  let caspioData = await caspiomain(Cust_ID, accessToken);
  let account = caspioData.InterconnectionAcc1;
    console.log(`30% Complete: Account number: ${account} found for ${caspioData.FullName}`);
  const latestMessageId = await getLatestApprovalMessage(auth,account);
  let dataRow;
  if (latestMessageId) {
      dataRow = await handleSingleMessage(auth, latestMessageId);
      console.log(`65% Complete: Approval email found for Account number: ${account}`);
	let updatedData = {
        "PAM": dataRow.pamId,
		"PAM_approvedate": dataRow.date,
		"ACDCApprove": dataRow.combinedKw,
		"InterconnectionName1": dataRow.customerName
		
    };
      await fetchCaspioData(accessToken, url, updatedData, 'PUT');
      //console.log(`100% Complete: PAM & Approval data updated for : ${caspioData.FullName}`);
    //console.log(dataRow);
	
  } else {
    console.log("No latest message with 'Application Approval' found.");
    }
    console.log(`100% Complete: PAM & Approval data updated for : ${caspioData.FullName}`);
})();

