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
let Cust_ID = process.argv[2];
//let Cust_ID = 9122;

async function cheriodata(url) {
    try {
        // Set a timeout and user-agent to mimic a real browser request and possibly avoid blocks.
        const response = await axios.get(url, {
            timeout: 10000, // 10 seconds timeout
            headers: {
                'User-Agent': 'Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/92.0.4515.159 Safari/537.36'
            }
        });

        // Check if the status code indicates a successful response.
        if (response.status !== 200) {
            console.error(`Received status code ${response.status}.`);
            return null; // Explicitly return null for a bad status code
        }

        // Load the response data into Cheerio for parsing.
        const loadedCheerio = cheerio.load(response.data);

        console.log("Successfully loaded Cheerio object"); // Debugging line
        return loadedCheerio;
    } catch (err) {
        console.error(`Error occurred: ${err.message}`);
        return null; // Explicitly return null if an error occurs
    }
}

async function cheriodatawithloading(url) {
    try {
        const browser = await puppeteer.launch({ headless: true });
        const page = await browser.newPage();

        await page.goto(url, { waitUntil: 'networkidle2' }); // Wait for the page to load
        const content = await page.content(); // Get page content

        await browser.close();

        return cheerio.load(content); // Load the content into cheerio for further processing

    } catch (err) {
        console.error(`Error occurred: ${err}`);
    }
}
async function scrapingdata($) {

    const info = {
        BIN: 'NA',
        communityBoard: 'NA',
        buildingsOnLot: 'NA',
        landmarkStatus: 'NA',
        additionalBin: 'NA',
        tidalWetlandsMapCheck: 'NA',
        specialFloodHazardAreaCheck: 'NA',
        violationsOATH_ECB: 'NA',
        jobsFilings: 'NA',
        openOATH_ECB: 'NA',
        Garfound: false,
        buildingClassification: 'NA'
    }
    // Ensure $ is valid before proceeding
    if (!$) {
        console.error("Invalid cheerio object passed to scrapingdata function.");
        return info; // Return default info object
    }

     $('table').each((tableIndex, table) => {


        if (tableIndex === 2) {
            $(table).find('tr').each((_, row) => {
                $(row).find('td').each((_, cell) => {
                    const cellText = $(cell).text().trim();
                    //console.log(cellText);
                    if (/BIN#/.test(cellText)) {
                        info.BIN = cellText.replace(/BIN#/, '').trim();
                        isBinFound = true;
                        console.log(info.BIN);
                    }
                });
            });
        } else if (tableIndex === 3) {
            $(table).find('tr').each((_, row) => {
                $(row).find('td').each((i, cell) => {
                    const cellText = $(cell).text().trim();
                    //console.log(cellText);
                    if (cellText.startsWith('Community Board')) {
                        let communityBoardText = $(cell).next().text().trim() || 'N/A';
                        if (communityBoardText !== 'N/A') {
                            let parts = communityBoardText.split(':');
                            if (parts.length > 1) {
                                info.communityBoard = parts[1].trim();
                            } else {
                                info.communityBoard = communityBoardText; // Keeps original if no ':' found
                            }
                        } else {
                            info.communityBoard = 'N/A';
                        }
                    }
                });
            });
        } else if (tableIndex === 5) {
            $(table).find('tr').each((_, row) => {
                $(row).find('td').each((i, cell) => {
                    const cellText = $(cell).text().trim();
                    //console.log(cellText);
                    if (cellText.startsWith('Landmark Status')) {
                        info.landmarkStatus = $(cell).next().text().trim() || 'N/A';
                    } else if (cellText.startsWith('Additional BINs for Building')) {
                        info.additionalBin = $(cell).next().text().trim() || 'N/A';
                    } else if (cellText === 'Department of Finance Building Classification:') {
                        let nextCellText = $(cell).next().text().trim() || 'N/A';
                        info.buildingClassification = nextCellText || 'N/A';
                    }
                });
            });
        }
        else if (tableIndex === 7) {
            let cells = $(table).find('td');
            for (let i = 0; i < cells.length; i++) {
                let cellText = $(cells[i]).text().trim();
                //console.log(`Debug: Index ${i} | Content: '${cellText}'`); // Debug log
                if (cellText === 'Tidal Wetlands Map Check:' || cellText === 'Special Flood Hazard Area Check:') {
                    let nextCellText = $(cells[i + 1]).text().trim();
                    if (cellText === 'Tidal Wetlands Map Check:') {
                        info.tidalWetlandsMapCheck = nextCellText || 'N/A';
                    } else if (cellText === 'Special Flood Hazard Area Check:') {
                        info.specialFloodHazardAreaCheck = nextCellText || 'N/A';
                    }
                }

                // Added condition for 'Department of Finance Building Classification'
                if (cellText === 'Department of Finance Building Classification:') {
                    let nextCellText = $(cells[i + 1]).text().trim();
                    info.buildingClassification = nextCellText || 'N/A';
                }
            }
        } else if (tableIndex === 8) {
            let cells = $(table).find('td');
            for (let i = 0; i < cells.length; i++) {
                let cellText = $(cells[i]).text().trim();

                // Find job/filings number
                if (cellText.includes('Jobs/Filings')) {
                    let nextCellText = $(cells[i + 1]).text().trim();
                    info.jobsFilings = nextCellText || 'N/A';
                }

                // Extract open OATH/ECB violation numbers
                if (cellText.includes('open OATH/ECB')) {
                    let match = cellText.match(/\d+/);  // Regular expression to find the first number in the string
                    if (match) {
                        info.openOATH_ECB = match[0];  // Assign the first number found
                    }
                }
            }
        }
        // Check for 'GAR' in table index 3 and redefine URL
        if (tableIndex === 3) {
            let cells = $(table).find('td');
            for (let i = 0; i < cells.length; i++) {
                let cellText = $(cells[i]).text().trim();
                if (cellText.includes('GAR')) {
                    info.Garfound = true;
                }
            }
        }


    });
    return info;
}
async function receivedata(initialUrl) {
    let $ = await cheriodatawithloading(initialUrl);
    let data = await scrapingdata($);

    if (!data.BIN || data.BIN === 'NA') {
        data = await scrapingdata($);
    }

    let finalData = data;
    console.log(data);
    if (data.Garfound && data.additionalBin) {

        const redefineUrl = `https://a810-bisweb.nyc.gov/bisweb/PropertyProfileOverviewServlet?requestid=0&bin=${data.additionalBin}`;
        if (redefineUrl) {
            console.log('Redefining url please wait. . . . .');
            console.log('Redefined url:', redefineUrl);
        } else {
            console.log('No GAR found.Moving on.....,');
        }

        let attempts = 0;
        let htmlLength = 0;  // Initialize htmlLength

        while ((!$ || htmlLength < 300) && attempts < 3) {
            if ($ && htmlLength < 300) {
                console.log("Delaying 10 seconds, $ is less than 300 characters.");
                await new Promise(resolve => setTimeout(resolve, 20000)); // Wait for 20 seconds
            }

            $ = await cheriodatawithloading(redefineUrl);

            if ($ === null) {
                console.log("Failed to get data, skipping...");
                return;
            }

            htmlLength = $ ? $('html').html().length : 0;  // Update htmlLength
            console.log(htmlLength);
            attempts++;  // Increment the counter
        }

        if (attempts >= 3) {
            console.log("Stopping after 3 attempts.");
            return; // Stop the function if 3 attempts are made
        }

        for (let i = 0; i < 5; i++) { // Loop up to 5 times
            let success = false; // A flag to track whether scraping was successful

            const $$ = $;
            finalData = await scrapingdata($$);

            if (finalData.BIN && finalData.BIN !== 'NA' && finalData.BIN !== '') {
                success = true; // Mark the attempt as successful
                break; // Exit the loop if BIN value is acceptable
            }

            console.log(`Run ${i + 1} data:`, finalData);
            if (!success) {
                console.log(success);
            }
        }
    }

    return finalData;
}
async function fetchECBStatuses(BIN) {
    const url = `https://a810-bisweb.nyc.gov/bisweb/ECBWwopByLocationServlet?requestid=0&allbin=${BIN}`;

    const response = await axios.get(url);
    if (response.status !== 200) {
        console.error(`Received status code ${response.status}.`);
        return null;
    }
    await new Promise(resolve => setTimeout(resolve, 1000));  // 1000 ms = 1 sec delay
    const $ = cheerio.load(response.data);

    let ecbNumStatuses = [];

    $('center > table:nth-child(3)').find('tr').each((rowIndex, row) => {
        const cells = $(row).find('td');

        if (cells.length > 1) {
            const ecbNum = $(cells[1]).text().trim();
            if (ecbNum.match(/\d{8}[A-Z]/)) {
                let ecbNumStatus = ecbNum + "-" + $(cells[5]).text().trim();
                ecbNumStatuses.push(ecbNumStatus);
            }
        }
    });
    Violation_numbers = ecbNumStatuses;
    return ecbNumStatuses;
}

async function fetchJobs(BIN) {
    const jobsUrl = `https://a810-bisweb.nyc.gov/bisweb/JobsQueryByLocationServlet?requestid=0&allbin=${BIN}`;

    const jobsResponse = await axios.get(jobsUrl);
    if (jobsResponse.status !== 200) {
        console.error(`Received status code ${jobsResponse.status}.`);
        return null;
    }

    let jobsdata = cheerio.load(jobsResponse.data);
    let jobs = [];
    let jobNumber = '';

    jobsdata('body > center > table > tbody > tr').each((i, element) => {
        let tds = jobsdata(element).find('td');

        if (tds.length === 10) { // Assuming job number rows always have 10 tds
            jobNumber = jobsdata(tds[1]).text().trim();
        }

        // Check for floor details in every row
        let workFloorsText = jobsdata(tds[1]).text().trim();
        if (workFloorsText.startsWith('Work on Floor(s):')) {
            let floors = workFloorsText.replace('Work on Floor(s):', '').trim();
            if (jobNumber && floors) {
                jobs.push(`${jobNumber}-${floors}`);
                jobNumber = ''; // Clear jobNumber after it has been used
            }
        }
    });

    return jobs;
}

/*
async function scrapeWebsite(url, depth = 0) {
                        // Check if we've hit the recursion limit
                        if (depth >= 2) { // Assuming you want to limit recursion to one additional call
                            console.log('Recursion limit reached.');
                            return;
                        }

                        let additionalBIN;
                        
                        try {                                                                                  // third try- catch starts here                                                                                      
                            const response = await axios.get(url);
                            if (response.status !== 200) {
                                console.error(`Received status code ${response.status}.`);
                                return null;

                            }
                            await new Promise(resolve => setTimeout(resolve, 1000));  // 1000 ms = 1 sec delay
                            const $ = cheerio.load(response.data);
                          
                            const info = { 
                                BIN: '',
                                communityBoard: 'NA',
                                buildingsOnLot: 'NA',
                                landmarkStatus: 'NA',
                                environmentalRestrictions: 'NA',
                                additionalBINsForBuilding: 'NA',
                                tidalWetlandsMapCheck: 'NA',
                                freshwaterWetlandsMapCheck: 'NA',
                                coastalErosionHazardAreaMapCheck: 'NA',
                                specialFloodHazardAreaCheck: 'NA',
                                violationsOATH_ECB: 'NA',
                                jobsFilings: 'NA',
                                openOATH_ECB: 'NA',
                                foundGAR: false  // New field added
                            }
                            let additionalBINPromises = [];
                            $('table').each((tableIndex, table) => {
                                $(table).find('tr').each((rowIndex, row) => {
                                    $(row).find('td').each((cellIndex, cell) => {
                                        const cellText = $(cell).text().normalize("NFD").trim().replace(/[":,]/g, '').replace(/\s+/g, ' ');
                                        console.log(cellText);
                                        if (cellText.startsWith('BIN#')) {
                                            const parts = cellText.split(' ');
                                            const binPart = parts.find(part => part.includes('#'));
                                            if (binPart) {
                                                info.BIN = binPart.split('#')[1];
                                            }
                                        }
                                       
                                        if (cellText === 'Community Board') {
                                            let communityBoard = $(row).find('td').eq(cellIndex + 1).text().trim();
                                            communityBoard = communityBoard.replace(/^:/, '');
                                            info.communityBoard = communityBoard;
                                        }
                                        if (cellText.includes('GAR')) {
                                            info.foundGAR = true;  // Update 'foundGAR' if 'GAR' is found
                                        }
                                        if (cellText === 'Buildings on Lot') {
                                            info.buildingsOnLot = $(row).find('td').eq(cellIndex + 1).text().trim();
                                        }

                                        if (cellText === 'Landmark Status') {
                                            let landmarkText = $(row).find('td').eq(cellIndex + 1).text().trim();
                                            if (landmarkText.includes('LANDMARK')) {
                                                info.landmarkStatus = 'Yes';
                                            } else {
                                                info.landmarkStatus = 'No';
                                            }
                                        }

                                        if (cellText === 'Environmental Restrictions') {
                                            info.environmentalRestrictions = $(row).find('td').eq(cellIndex + 1).text().trim();
                                        }

                                        if (cellText === 'Additional BINs for Building') {
                                            info.additionalBINsForBuilding = $(row).find('td').eq(cellIndex + 1).text().trim();
                                        }

                                        if (cellText === 'Tidal Wetlands Map Check') {
                                            info.tidalWetlandsMapCheck = $(row).find('td').eq(cellIndex + 1).text().trim();
                                        }

                                        if (cellText === 'Freshwater Wetlands Map Check') {
                                            info.freshwaterWetlandsMapCheck = $(row).find('td').eq(cellIndex + 1).text().trim();
                                        }

                                        if (cellText === 'Coastal Erosion Hazard Area Map Check') {
                                            info.coastalErosionHazardAreaMapCheck = $(row).find('td').eq(cellIndex + 1).text().trim();
                                        }

                                        if (cellText === 'Special Flood Hazard Area Check') {
                                            info.specialFloodHazardAreaCheck = $(row).find('td').eq(cellIndex + 1).text().trim();
                                        }

                                        if (cellText === 'Violations-OATH/ECB') {
                                            info.violationsOATH_ECB = $(row).find('td').eq(cellIndex + 1).text().trim();
                                        }

                                        if (cellText === 'Jobs/Filings') {
                                            info.jobsFilings = $(row).find('td').eq(cellIndex + 1).text().trim();
                                        }

                                        if (cellText.includes('open OATH/ECB')) {
                                            const matches = cellText.match(/(\d+)\sopen OATH\/ECB/);
                                            if (matches && matches.length > 1) {
                                                info.openOATH_ECB = matches[1];
                                            }
                                        }

                                    });
                                });
                            });

                            if (info.foundGAR && info.additionalBINsForBuilding !== 'NA') {
                                additionalBIN = info.additionalBINsForBuilding;
                            }
                            console.log(`additional bin:`, additionalBIN);
                            
                            if (additionalBIN) {
                                const url2 = `https://a810-bisweb.nyc.gov/bisweb/PropertyProfileOverviewServlet?requestid=1&bin=${additionalBIN}`;
                                console.log(`Redefining URL to: ${url2}`);
                                const updatedInfo = await scrapeWebsite(url2, depth + 1);

                                if (updatedInfo) {

                                    return updatedInfo;
                                }
                            }

                            console.log(`BIN: ${info.BIN}`);
                            console.log(`communityBoard: ${info.communityBoard}`);
                            console.log(`buildingsOnLot: ${info.buildingsOnLot}`);
                            console.log(`landmarkStatus: ${info.landmarkStatus}`);
                            console.log(`environmentalRestrictions: ${info.environmentalRestrictions}`);
                            console.log(`additionalBINsForBuilding: ${info.additionalBINsForBuilding}`);
                            console.log(`tidalWetlandsMapCheck: ${info.tidalWetlandsMapCheck}`);
                            console.log(`freshwaterWetlandsMapCheck: ${info.freshwaterWetlandsMapCheck}`);
                            console.log(`specialFloodHazardAreaCheck: ${info.specialFloodHazardAreaCheck}`);
                            console.log(`jobsFiling: ${info.jobsFilings}`);
                            console.log(`open OATH/ECB: ${info.openOATH_ECB}`);
                            console.log(`GAR Found: ${info.foundGAR}`);


                            if (additionalBINPromises.length > 0) {
                                const additionalBINInfo = await Promise.all(additionalBINPromises);
                                info.additionalBINInfo = additionalBINInfo;
                            }
                            return info;

                        } catch (error) {
                            console.error(error);
                            return null;
                        }
                    }
*/
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
// New code starts............
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
//////////////////////....ENDS
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
async function deleteFileWithRetry(filePath, retryCount = 5) {
    for (let attempt = 0; attempt < retryCount; attempt++) {
        try {
            fs.unlinkSync(filePath);
            console.log(`Successfully deleted ${filePath}`);
            return;  // exit the function if deletion is successful
        } catch (error) {
            console.error(`Error deleting file on attempt ${attempt + 1}:`, error);
            await new Promise(resolve => setTimeout(resolve, 1000));  // wait for 1 second before next attempt
        }
    }
    console.error(`Failed to delete ${filePath} after ${retryCount} attempts`);
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
        const Name = combinedData.Name;
        const Phone = combinedData.Phone1;
        const email = combinedData.email;
        const Price = combinedData.Contract_Price;
        const NYC_LI = combinedData.NYC_LI;
        const Finance = combinedData.Finance;
        const Size = combinedData.System_Size_Sold;
        const Panel = combinedData.Panel;
        const Panel_count = combinedData.Number_Of_Panels;
        const permit_type = combinedData.Canopy_Used;
        const Affordable = combinedData.Affordable;
        const Solarinsure = combinedData.Solarinsure;

        return { CID, Name, Address, Phone, email, Price, NYC_LI, Finance, Size, Panel, Panel_count, permit_type, Affordable, Solarinsure };
    } catch (error) {
        console.error(`Error in main: ${error}`);
        if (error.response) {
            console.error('Error status:', error.response.status);
            console.error('Error data:', error.response.data);
        }
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
async function runProcess(NYC_LI, client, filePath, buildingPermitInitialFolderId, Oath_Ecb_Violation, Landmark, Flood_Zone) {
    if (NYC_LI === 'NYC') {
        console.log('Spawning the child process');

        const child = spawn('node', ['./Nycdocs.js']);

        child.stdout.on('data', (data) => {
            console.log(`stdout: ${data}`);
        });

        child.stderr.on('data', (data) => {
            console.error(`stderr: ${data}`);
        });

        child.on('error', (error) => {
            console.error(`error: ${error.message}`);
        });

        child.on('close', async (code) => {
            if (code !== 0) {
                console.log(`child process exited with code ${code}`);
            } else {
                console.log("Child process completed successfully. Checking for file...");
                if (fs.existsSync(filePath)) {
                    await uploadFile(client, filePath, buildingPermitInitialFolderId, Oath_Ecb_Violation, Landmark, Flood_Zone);
                } else {
                    console.log("File still does not exist after child process completion.");
                }
            }
        });
    } else if (NYC_LI !== 'NYC') {
        console.log('Skipping process because Project is not in NYC');
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
/*
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

    async function deleteFolder(client, folderId) {
        const drive = google.drive({ version: 'v3', auth: client });
        try {
            await drive.files.delete({
                'fileId': folderId
            });
            console.log('Folder deleted');
        } catch (error) {
            console.log('An error occurred: ', error);
        }
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

    const existingFolders = await checkIfFolderExists(client, `${customerName}-${custID}`, parentFolderId);
    if (existingFolders.length > 0) {
        console.log("Folder already exists");
        for (const folder of existingFolders) {
            await deleteFolder(client, folder.id);
        }
    }

    let mainFolderId = await createFolder(client, `${customerName}-${custID}`, parentFolderId);
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
        let subFolderId = await createFolder(client, folder, mainFolderId);
        folderIds[folder] = { id: subFolderId };

        if (subSubFolders[folder]) {
            for (const subSubFolder of subSubFolders[folder]) {
                let subSubFolderId = await createFolder(client, subSubFolder, subFolderId);
                folderIds[folder][subSubFolder] = subSubFolderId;
            }
        }
    }

    console.log(`Created folders for ${customerName}-${custID}`);
    return folderIds;
}
*/
/* 
async function createFoldersInDrive(client, customerName) {
    const drive = google.drive({ version: 'v3', auth: client });
    const parentFolderId = '1cg70dBWgSqTCwO94dmN4Gb9KbWA0v5bk';
    const folderName = `${customerName} NYC forms`;

    async function checkIfFolderExists(client, folderName, parentFolderId) {
        const response = await drive.files.list({
            q: `mimeType='application/vnd.google-apps.folder' and '${parentFolderId}' in parents and name='${folderName}' and trashed=false`,
            fields: 'files(id, name)',
        });
        return response.data.files;
    }

    async function deleteFolder(client, folderId) {
        const drive = google.drive({ version: 'v3', auth: client });
        try {
            await drive.files.delete({
                'fileId': folderId
            });
            console.log('Folder deleted');
        } catch (error) {
            console.log('An error occurred: ', error);
        }
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

    const existingFolders = await checkIfFolderExists(client, folderName, parentFolderId);
    if (existingFolders.length > 0) {
        console.log("Folder already exists");
        for (const folder of existingFolders) {
            await deleteFolder(client, folder.id);
        }
    }

    let mainFolderId = await createFolder(client, folderName, parentFolderId);
    console.log(`Created folder: ${folderName}`);
    return mainFolderId;
}
*/
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
function sleep(ms) {
  return new Promise(resolve => setTimeout(resolve, ms));
}

(async () => {        
    let family;
    let floors;
    let year;
    let CB;
    // PUPPETEER BEGINS ...........................................................
    try {

    //....................................................................
    // HAS TO BE UNDER A TRY BLOCK TO GET THE CELL VALUE
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
        const filePath = './filled.pdf';
		let browser;
        for (let i = 0; i < sheetdata; i++) {
            try {
                // Attempt to delete the file at the start of each iteration
                if (fs.existsSync(filePath)) {
                    await deleteFileWithRetry(filePath);
                }
                //let Cust_ID = sheetdata[i];
                let caspiodata = await caspiomain(Cust_ID, accessToken);
                console.log("Data: ", caspiodata);
                let cellValue = caspiodata.Address;
                const sendurl = `https://c1abd578.caspio.com/rest/v2/tables/SolarProcess/records?q.where=CustID=${Cust_ID}`;
                let NYC_LI = caspiodata.NYC_LI;
                //console.log('sendurl:', sendurl);
                await sleep(5000); // adjust the delay time as needed
                if (NYC_LI !== 'NYC') {
                    console.log("Can not perform analysis as its not NYC. Skipping process..........");
                }
                else { 
				      browser = await puppeteer.launch({
            headless: false,
            slowMo: 1 // delay in milliseconds
        });

        const context = await browser.createIncognitoBrowserContext();
        const page = await context.newPage();  // Now using context to open the page
        const timeout = 60000;
        page.setDefaultTimeout(timeout);
        await page.setViewport({
            width: 1263,
            height: 937
        });
        await page.goto('https://zola.planning.nyc.gov/about#9.72/40.7125/-73.733');
        if (page.url() !== 'https://zola.planning.nyc.gov/about#9.72/40.7125/-73.733') {
            console.log('Relaunching due to slow load...');
            await browser.close();
            const newBrowser = await puppeteer.launch({
                headless: false,
                slowMo: 1 // delay in milliseconds
            });
            const newContext = await newBrowser.createIncognitoBrowserContext();
            const newPage = await newContext.newPage();
			await newPage.setViewport({
            width: 1263,
            height: 937
        });
            await newPage.goto('https://zola.planning.nyc.gov/about#9.72/40.7125/-73.733', { waitUntil: 'networkidle0' });
            await newPage.waitForTimeout(5000);  // waits for 5000 milliseconds = 5 seconds
        } else {
            await page.waitForTimeout(5000);  // waits for 5000 milliseconds = 5 seconds
        }
                    //....................................................................
                    {
                        const targetPage = page;
                        await scrollIntoViewIfNeeded([['#ember18']], targetPage, timeout);
                        await retryClickElement(targetPage, '#ember18');
                        await page.waitForTimeout(1000); // wait for 1 seconds
                    }
                    {
                        const targetPage = page;
                        await scrollIntoViewIfNeeded([['#ember18']], targetPage, timeout);
                        const element = await waitForSelectors([['#ember18']], targetPage, { timeout, visible: true });
                        const inputType = await element.evaluate(el => el.type);
                        if (inputType === 'select-one') {
                            await changeSelectElement(element, cellValue)
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
                            await typeIntoElement(element, cellValue);
                        } else {
                            await changeElementValue(element, cellValue);
                        }
                        await page.waitForTimeout(1000); // wait for 1 seconds
                    }
                    {
                        const targetPage = page;
                        await scrollIntoViewIfNeeded([
                            [
                                'li.result'
                            ]
                        ], targetPage, timeout);
                        const element = await waitForSelectors([
                            [
                                'li.result'
                            ]
                        ], targetPage, { timeout, visible: true });
                        await element.click({
                            offset: {
                                x: 144,
                                y: 14.203125,
                            },
                        });
                    }
                    //// Click on <span> "30 ft" frontage clicK
                    {
                        const targetPage = page;
                        await scrollIntoViewIfNeeded([
                            [
                                '.lot-details > .data-grid:nth-child(4) > .datum'
                            ]
                        ], targetPage, timeout);
                        const element = await waitForSelectors([
                            [
                                '.lot-details > .data-grid:nth-child(4) > .datum'
                            ]
                        ], targetPage, { timeout, visible: true });


                        // Evaluate and log the text content of the element
                        const frontage = await element.evaluate(el => el.textContent);
                        await writeSpreadsheetData(client, frontage, 'M5');
                        console.log("Frontage: ", frontage);

                        await element.click({
                            delay: 2888.399999976158,
                            offset: {
                                x: 134.25,
                                y: 3.75,
                            },
                        });
                    }
                    // Click on <span> "100 ft" depth click
                    {
                        const targetPage = page;
                        await scrollIntoViewIfNeeded([
                            [
                                '.lot-details > .data-grid:nth-child(5) > .datum'
                            ]
                        ], targetPage, timeout);
                        const element = await waitForSelectors([
                            [
                                '.lot-details > .data-grid:nth-child(5) > .datum'
                            ]
                        ], targetPage, { timeout, visible: true });

                        // Evaluate and log the text content of the element
                        const depth = await element.evaluate(el => el.textContent);
                        await writeSpreadsheetData(client, depth, 'N5');
                        console.log("Depth:", depth);

                        await element.click({
                            delay: 3773.2999999523163,
                            offset: {
                                x: 122.25,
                                y: 2.75,
                            },
                        });
                    }

                    // YEAR BUILT------------------------1920 Example
                    {
                        const targetPage = page;
                        await scrollIntoViewIfNeeded([
                            [
                                '.lot-details > .data-grid:nth-child(6) > .datum'
                            ]
                        ], targetPage, timeout);
                        const element = await waitForSelectors([
                            [
                                '.lot-details > .data-grid:nth-child(6) > .datum'
                            ]
                        ], targetPage, { timeout, visible: true });

                        // Evaluate and log the text content of the element
                        year = await element.evaluate(el => el.textContent.trim());
                        await writeSpreadsheetData(client, year, 'O5');
                        console.log("Year Built:", year);

                        await element.click({
                            delay: 846.8999999761581,
                            offset: {
                                x: 154.25,
                                y: 17.5625,
                            },
                        });
                    }

                    // Click on <span> "2.75". NUMBER OF FLOORS
                    {
                        const targetPage = page;
                        await scrollIntoViewIfNeeded([
                            [
                                '.lot-details > .data-grid:nth-child(9) > .datum'
                            ]
                        ], targetPage, timeout);
                        const element = await waitForSelectors([
                            [
                                '.lot-details > .data-grid:nth-child(9) > .datum'
                            ]
                        ], targetPage, { timeout, visible: true });

                        // Evaluate and log the text content of the element
                        floors = await element.evaluate(el => el.textContent.trim());
                        await writeSpreadsheetData(client, floors, 'P5');
                        console.log("# of Floors:", floors);

                        await element.click({
                            delay: 709.1000000238419,
                            offset: {
                                x: 147.25,
                                y: 15.5625,
                            },
                        });
                    }

                    // Click on <span> "1" NUMBER OF UNITS

                    {
                        const targetPage = page;
                        await scrollIntoViewIfNeeded([
                            [
                                '.data-grid:nth-child(12) > .datum'
                            ]
                        ], targetPage, timeout);
                        const element = await waitForSelectors([
                            [
                                '.data-grid:nth-child(12) > .datum'
                            ]
                        ], targetPage, { timeout, visible: true });

                        // Evaluate and log the text content of the element
                        const units = await element.evaluate(el => el.textContent);
                        await writeSpreadsheetData(client, units, 'Q5');
                        console.log("# of Units:", units);

                        await element.click({
                            button: 'right',
                            offset: {
                                x: 53.84375,
                                y: 5.5625,
                            },
                        });
                    }

                    // For Acrislink......................................................................

                    {
                        const acrislinkElement = await page.$('.data-grid:nth-child(14) > .datum a');
                        const Acrislink = await page.evaluate(el => el.getAttribute('href'), acrislinkElement);
                        console.log("Acrislink:", Acrislink);
                    }

                    // Single or multiple family................

                    {
                        const targetPage = page;
                        await scrollIntoViewIfNeeded([
                            [
                                '.lot-details > .data-grid:nth-child(2) > .datum'
                            ]
                        ], targetPage, timeout);
                        const element = await waitForSelectors([
                            [
                                '.lot-details > .data-grid:nth-child(2) > .datum'
                            ]
                        ], targetPage, { timeout, visible: true });

                        // Evaluate and log the text content of the element
                        family = await element.evaluate(el => el.textContent.trim());
                        console.log("FamilyType:", family);

                        await element.click({
                            button: 'right',
                            offset: {
                                x: 53.84375,
                                y: 5.5625,
                            },
                        });
                    }
                    // For Zoning information............................
                    let Zone;
                    {
                        const zoningElement = await page.$('.lot-zoning-list'); // Please replace '.zoning-selector' with the actual selector
                        let zoning = await page.evaluate(el => el.textContent, zoningElement);
                        zoning = zoning.split(":")[1].trim(); // This will give you "RXX" part from "Zoning: RXX"
                        await writeSpreadsheetData(client, zoning, 'R5');
                        console.log("Zoning:", zoning);
                        Zone = zoning;
                    }
                    //.....................................................................................................

                    const targetPage = page;
                    await scrollIntoViewIfNeeded([
                        [
                            '.text-small:nth-child(3)'
                        ]
                    ], targetPage, timeout);
                    const element = await waitForSelectors([
                        [
                            '.text-small:nth-child(3)'
                        ]
                    ], targetPage, { timeout, visible: true });

                    const elementText = await element.evaluate(el => el.textContent);
                    const splitText = elementText.trim().split('|');
                    const boroughRaw = splitText[0].trim();
                    const boroughData = boroughRaw.split('(');
                    const borough = boroughData[0].trim();
                    const boroughNum = boroughData[1].replace(')', '').replace('Borough', '').trim();
                    const block = splitText[1].replace('Block', '').trim();
                    const lot = splitText[2].replace('Lot', '').trim();
                    //console.log("Borough: ", borough);
                    //console.log("Borough #: ", boroughNum);
                    //console.log("Block: ", block);
                    //console.log("Lot: ", lot);
                    // Constructing the new URL based on boroughNum, block and lot
                    const newUrl = `https://a810-bisweb.nyc.gov/bisweb/PropertyProfileOverviewServlet?boro=${boroughNum}&block=${block}&lot=${lot}&go3=+GO+&requestid=0`;
                    console.log("New URL: ", newUrl);  // Log the new URL
                    await element.click({
                        delay: 1277.1000000238419,
                        offset: {
                            x: 275.25,
                            y: 4.3125,
                        },
                    });
                    //..........................................................................................................................................
                    //..........................................................................................................................................
                    const finaldata = await receivedata(newUrl); 
                    console.log("Scraped data: ", finaldata);
                    //let Permit_Status;
                    let County;
                    //let Job;
                    //const cellMapping = ['W6', 'W7', 'W8', 'W9', 'W10', 'W11', 'W12', 'W13'];
                    //let ecbNumStatuses = [];
                    let BIN = finaldata.BIN;                   
                    let Flood_Zone = finaldata.specialFloodHazardAreaCheck;
                    let Oath_Ecb_Violation = finaldata.openOATH_ECB;
                    let Landmark = finaldata.landmarkStatus;
                    let CB = finaldata.communityBoard;
                    let Violation_numbers = finaldata.openOATH_ECB;
                    let Current_job_filings = finaldata.jobsFilings;
                    let borough1 = borough;
                    let yearbuilt = year;
                    switch (borough1) {
                        case "Queens":
                            County = "Queens";
                            break;
                        case "Brooklyn":
                            County = "Kings";
                            break;
                        case "Bronx":
                            County = "Bronx";
                            break;
                        case "Staten Island":
                            County = "Staten Island";
                            break;
                        default:
                            County = "Unknown";
                    }
                    let Asbestos = (yearbuilt >= 1987) ? "Not Required" : "Required"; 
                    //...................................................................................................................................
                    let updatedData = {
                            "Block": block,
                            "Lot": lot,
                            "Bin": BIN,
                            "Flood_Zone": Flood_Zone,
                            "Oath_Ecb_Violation": Oath_Ecb_Violation,
                            "Landmark": Landmark,
                            "Violation_numbers": Violation_numbers,
                            "Current_job_filings": Current_job_filings,
                            "Borough": borough,
                            "County": County,
                            //"Towns_of_Longisland": "NA",
                            "Asbestos": Asbestos,
                            "built_year": year,
                            "Stories": floors,
                            "Family": family,
                            "Zoning": Zone,
                            "CB": CB
                        };
                    
                    console.log('Data to be updated:', updatedData);
                    await fetchCaspioData(accessToken, sendurl, updatedData, 'PUT')
                    //....................................................................................................................................                                                                      

                
                }
                await sleep(10000);
                } catch (error) {
                    console.error(`Error in iteration ${i}: `, error);
            }
            }
        await browser.close();
        console.log('browser closed');
        //......................................................................................
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
        //.................................................


    } catch (error) {                                                               // 2nd try-catch ends here 
        console.error(error);
    }

    })().catch(err => {                                                                                                                  
        console.error(err);
        process.exit(1);
    });