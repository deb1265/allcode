const axios = require('axios');
const util = require('util'); // Importing util module
const qs = require('querystring'); // If you are using querystring to stringify request body
const { DocumentProcessorServiceClient } = require('@google-cloud/documentai').v1;
const fs = require('fs').promises;
const caspioTokenUrl = 'https://c1abd578.caspio.com/oauth/token';
const clientID = process.env.clientID;
const clientSecret = process.env.clientSecret;
const GoggleApi = process.env.GoggleApi;
const keys =JSON.parse(process.env.SERVICE_KEY);
const Cust_ID = process.argv[2];
const imagePath = process.argv[3];
const PUTurl = `https://c1abd578.caspio.com/rest/v2/tables/SolarProcess/records?q.where=CustID=${Cust_ID}`;
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
async function fetchCaspioData(url, data = null, method = 'GET') {
    const accessToken = await caspioAuthorization();
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
function findText(obj) {
    if (obj && typeof obj === 'object') {
        for (const key in obj) {
            if (key === 'text' && typeof obj[key] === 'string') {
                return obj[key];
            }
            const found = findText(obj[key]);
            if (found) return found;
        }
    }
    return null;
}
async function detectTextInImage(imagePath) {
    if (!imagePath) {
        console.warn("Image path is null or undefined in detectTextInImage.");
        return null;
    }

    // Create a client
    const client = new DocumentProcessorServiceClient({
        keyFilename: './servicekey.json'
    });
    const processorEndpoint = 'projects/206322115501/locations/us/processors/6836530c68231a60';   //

    // Read the file into memory
    const imageFile = await fs.readFile(imagePath);

    // Convert the image data to a Buffer and base64 encode it
    const encodedImage = Buffer.from(imageFile).toString('base64');

    // Create the request
const request = {
    name: processorEndpoint,
    inlineDocument: {
        content: encodedImage,
        mimeType: 'application/pdf', // Make sure to provide the correct mime type
    },
};
// Process the document
    // Process the document
    const [response] = await client.processDocument(request);

    const textContent = findText(response);
    console.log(textContent);
    if (textContent) {
        //console.log(textContent); // Output the text content
    } else {
        console.warn("Text is null or undefined in extractValues.");
    }
    // Applying the extractValues function to the extracted text
    const resultArray = extractValues(textContent);

    return resultArray; // Returning the formatted result
    console.log(resultArray);
}
async function askGPT4(question, context) {
    const payload = {
        model: 'davinci-codex', // You can choose the model you want to use here
        prompt: `${context}\nQuestion: ${question}`,
        max_tokens: 50
    };

    const response = await fetch('https://api.openai.com/v1/engines/davinci-codex/completions', {
        method: 'POST',
        headers: {
            'Content-Type': 'application/json',
            'Authorization': `Bearer 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` // Replace with your API key
        },
        body: JSON.stringify(payload)
    });

    const data = await response.json();
    return data.choices ? data.choices[0].text : null;
}

function extractValues(text) {
    if (!text) {
        console.warn("Text is null or undefined in extractValues.");
        return null;
    }
    const summaryIndex = text.indexOf('Summary');
    const endIndex = text.indexOf('Weighted average by');
    const summaryContent = text.slice(summaryIndex, endIndex).trim();
    const lines = summaryContent.split('\n').slice(8); // Skip the header lines
    const columns = ['Array', 'Panel Count', 'Azimuth (deg.)', 'Pitch (deg.)', 'Annual TOF (%)', 'Annual Solar Access (%)', 'Annual TSRF (%)'];
    const rows = [];
    let currentRow = [];

    for (const line of lines) {
        const maybeNumber = parseInt(line.trim(), 10);

        // If the line starts with a number between 1 and 8 (inclusive), and there is a previous row, add it to rows
        if (!isNaN(maybeNumber) && maybeNumber >= 1 && maybeNumber <= 8) {
            if (currentRow.length >= 7) {
                rows.push(currentRow.slice(0, 7)); // Take the first 7 elements
            }
            currentRow = [line]; // Start a new row with the current line
        } else {
            // If it's not a row number, just add it to the current row
            currentRow.push(line);
        }
    }

    // Add the last row if it's complete
    if (currentRow.length === 7) {
        rows.push(currentRow);
    }

    return { columns, rows };
}
function isNumber(value) {
    return !isNaN(parseFloat(value)) && isFinite(value);
}
// Example usage
detectTextInImage(imagePath).then(result => {
    console.log(result);
    const context = result; // Assuming result contains the paragraph of text
    const question = "Can you extract last two lines from the text?";

    askGPT4(question, context).then(answer => {
        console.log(`Answer: ${answer}`);
    });
});