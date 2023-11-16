//const { exec } = require('child_process');
const axios = require('axios');
const util = require('util'); // Importing util module
const fs = require('fs');
const https = require('https');
const qs = require('querystring'); // If you are using querystring to stringify request body
const unlinkAsync = util.promisify(fs.unlink);
const exec = util.promisify(require('child_process').exec);
const Cust_ID = process.argv[2];
//const Cust_ID = '7528';
const caspioFileURL1 = `https://c1abd578.caspio.com/rest/v2/tables/SolarProcess/records?q.select=shadereport&q.where=CustID=${Cust_ID}`;
const caspioTokenUrl = 'https://c1abd578.caspio.com/oauth/token';
const clientID = process.env.clientID;
const clientSecret = process.env.clientSecret;
const GoggleApi = process.env.GoggleApi;
const keys =JSON.parse(process.env.SERVICE_KEY);
const delay = (ms) => new Promise((resolve) => setTimeout(resolve, ms));
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
const downloadBills = async (accessToken, Cust_ID, caspioFileURL) => {
    const appKey = '68107000';
    // Extract utility bill property from URL.
    const match = caspioFileURL.match(/q\.select=(shadereport)/);;
    if (!match) {
        console.error("Failed to match utility bill property from URL.");
        return null;
    }
    const utilityBillProperty = match[1];
    // Fetch the file URL using the customer ID.
    const result = await axios.get(caspioFileURL, { headers: { 'Authorization': `Bearer ${accessToken}` } });
    if (!result.data || !result.data.Result || result.data.Result.length === 0) {
        console.warn('No file path found for given customer ID.');
        return null;
    }
    const filePath = result.data.Result[0][utilityBillProperty];
    if (!filePath) {
        console.warn(`${utilityBillProperty} not found or is undefined.`);
        return null;
    }
    const sanitizedFileName = filePath.substring(filePath.lastIndexOf('/') + 1).replace(/\s+/g, '_'); // replace spaces with underscores
    const fileExtension = sanitizedFileName.substring(sanitizedFileName.lastIndexOf('.') + 1);
    const localPath = `/tmp/shadereport.${fileExtension}`;
    //const caspioFileDownloadURL = `https://cdn.caspio.com/${appKey}${filePath}?ver=1`;
    return new Promise((resolve, reject) => {
        const file = fs.createWriteStream(localPath);
        // Encode the filepath to ensure spaces and other special characters are correctly formatted for a URL
        const encodedFilePath = encodeURIComponent(filePath);

        const requestOptions = {
            hostname: 'cdn.caspio.com',
            path: `/${appKey}${encodedFilePath}?ver=1`,
            headers: {
                'Authorization': `Bearer ${accessToken}`,
                'User-Agent': 'Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/92.0.4515.159 Safari/537.36',
                'Referer': 'https://c1abd578.caspio.com/'
            }
        };
        const request = https.get(requestOptions, (response) => {
            if (response.statusCode !== 200) {
                reject(new Error(`Failed to download. Status code: ${response.statusCode}`));
                response.resume(); // Consume response to free up memory.
                return;
            }

            response.pipe(file);

            file.on('finish', () => {
                file.close(() => {
                    console.log('File saved successfully');
                    resolve(localPath);
                });
            });
        });
        request.on('error', (error) => {
            fs.unlink(localPath, () => { }); // Delete the file if there's an error (optional).
            reject(error.message);
        });
        file.on('error', (error) => {
            fs.unlink(localPath, () => { }); // Delete the file on any error.
            reject(error.message);
        });
    });
}
const deleteFiles = async (mainFile, arrayOfFiles = []) => {
    try {
        // Delete the main file (like the PDF itself)
        fs.unlinkSync(mainFile);
        console.log(`Deleted: ${mainFile}`);

        // If there are other files in the array, delete them too
        for (let file of arrayOfFiles) {
            fs.unlinkSync(file);
            console.log(`Deleted: ${file}`);
        }
    } catch (error) {
        console.error(`Error deleting files: ${error.message}`);
    }
};
(async () => {
    const accessToken = await caspioAuthorization();
    console.log("Successfully connected to Caspio!");
    downloadBills(accessToken, Cust_ID, caspioFileURL1)
        .then(filepath => {
            if (filepath) {
                console.log(`File path is: ${filepath}`);
                // Return the filepath or perform other actions as needed
                return filepath;
            } else {
                console.error('No file path was returned from downloadshade.');
            }
        })
        .catch(error => console.error('Error downloading shade:', error));
})();