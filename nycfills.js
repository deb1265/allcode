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
const nycsformsfileId = '1SlkgPpX_xCwph-Nlf0ej-vCaMy8sZhv7';
const sharp = require('sharp');
const delay = (ms) => new Promise((resolve) => setTimeout(resolve, ms));
//const nycblanksdownloadedpath = path.resolve(__dirname, 'file.pdf'); // define nycblanksdownloadedpath outside the functions
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
async function uploadFile(client, filePath, folderId, Oath_Ecb_Violation, Landmark, Flood_Zone,Name,Lastname) {
    try {
        const drive = google.drive({ version: 'v3', auth: client });
        const ext = path.extname(filePath);
        const signedFormsFolderId = await createFolder(`${Name}NYC-forms`, folderId, drive);
        // If the file is a PDF, split it into parts and upload each part
        if (ext === '.pdf') {
            const pdfBytes = fs.readFileSync(filePath);
            const pdfDoc = await PDFDocument.load(pdfBytes);

const parts = {
    [`${Lastname}-PW3`]: [0, 1],
    [`${Lastname}-PTA4`]: [2, 5],
    [`${Lastname}-TPP1`]: [6, 7],
    [`${Lastname}-TR1`]: [8, 10],
    [`${Lastname}-TR8`]: [11, 12],
    [`${Lastname}-PW2`]: [20, 22],
    [`${Lastname}-EnergyReports`]: [26, 40]
};

if (Oath_Ecb_Violation !== "NA") {
    parts[`${Lastname}-L2_form`] = [13, 14];
}
if (Landmark === "Yes") {
    parts[`${Lastname}-Landmark_Application`] = [15, 17];
}
if (Flood_Zone !== "NA") {
    parts[`${Lastname}-Flood_zone_docs`] = [18, 19];
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

                fs.unlinkSync(newFilePath);
            }
            //fs.unlinkSync(filePath);
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
            
            //fs.unlinkSync(filePath);
        }
		return (signedFormsFolderId);
        console.log('File exists after deletion: ', fs.existsSync(filePath));
    } catch (error) {
        console.error('Error uploading file: ', error);
    }
}
async function fetchData(block, lot, boroughNum) {
    let marketValue;
    let actual_av;
    let actual_av_land;
    const browser = await puppeteer.launch({
        headless: false, slowMo: 10, // Uncomment to visualize test
    });
    const page = await browser.newPage();
    const timeout = 30000;
    page.setDefaultTimeout(timeout);

    {
        const targetPage = page;
        await targetPage.setViewport({
            width: 1263,
            height: 937
        })
    }
    {
        const targetPage = page;
        const promises = [];
        promises.push(targetPage.waitForNavigation());
        await targetPage.goto('https://www.nyc.gov/site/finance/pts-interim-message.page');
        await Promise.all(promises);
    }
    {
        const targetPage = page;
        const promises = [];
        promises.push(targetPage.waitForNavigation());
        await scrollIntoViewIfNeeded([
            [
                '#main div.span9 a'
            ],
            [
                'xpath///*[@id="1675780670296"]/div/div/div[2]/div/div/a'
            ],
            [
                'pierce/#main div.span9 a'
            ],
            [
                'aria/Continue to Property Tax System'
            ],
            [
                'text/Continue to Property'
            ]
        ], targetPage, timeout);
        const element = await waitForSelectors([
            [
                '#main div.span9 a'
            ],
            [
                'xpath///*[@id="1675780670296"]/div/div/div[2]/div/div/a'
            ],
            [
                'pierce/#main div.span9 a'
            ],
            [
                'aria/Continue to Property Tax System'
            ],
            [
                'text/Continue to Property'
            ]
        ], targetPage, { timeout, visible: true });
        await element.click({
            offset: {
                x: 125.453125,
                y: 12.40625,
            },
        });
        await Promise.all(promises);
    }
    {
        const targetPage = page;
        const promises = [];
        promises.push(targetPage.waitForNavigation());
        await scrollIntoViewIfNeeded([
            [
                '#pagewrapper a:nth-of-type(2)'
            ],
            [
                'xpath///*[@id="content"]/div/div[2]/a[2]'
            ],
            [
                'pierce/#pagewrapper a:nth-of-type(2)'
            ],
            [
                'aria/BBL Search'
            ],
            [
                'text/BBL Search'
            ]
        ], targetPage, timeout);
        const element = await waitForSelectors([
            [
                '#pagewrapper a:nth-of-type(2)'
            ],
            [
                'xpath///*[@id="content"]/div/div[2]/a[2]'
            ],
            [
                'pierce/#pagewrapper a:nth-of-type(2)'
            ],
            [
                'aria/BBL Search'
            ],
            [
                'text/BBL Search'
            ]
        ], targetPage, { timeout, visible: true });
        await element.click({
            offset: {
                x: 92.15625,
                y: 17,
            },
        });
        await Promise.all(promises);
    }
    {
        try {
            const targetPage = page;
            const promises = [];
            promises.push(targetPage.waitForNavigation());
            await scrollIntoViewIfNeeded([
                ['#btAgree'],
                ['xpath///*[@id="btAgree"]'],
                ['pierce/#btAgree'],
                ['aria/Agree'],
                ['text/Agree']
            ], targetPage, timeout);
            const element = await waitForSelectors([
                ['#btAgree'],
                ['xpath///*[@id="btAgree"]'],
                ['pierce/#btAgree'],
                ['aria/Agree'],
                ['text/Agree']
            ], targetPage, { timeout, visible: true });
            await element.click({
                offset: {
                    x: 26.15625,
                    y: 19,
                },
            });
            await Promise.all(promises);
        } catch (error) {
            console.log("Selector not found in the first block, moving to the next block. Error: ", error);
        }
    }
    {
        const targetPage = page;
        await scrollIntoViewIfNeeded([
            [
                '#inpParid'
            ],
            [
                'xpath///*[@id="inpParid"]'
            ],
            [
                'pierce/#inpParid'
            ],
            [
                'aria/Parid'
            ]
        ], targetPage, timeout);
        const element = await waitForSelectors([
            [
                '#inpParid'
            ],
            [
                'xpath///*[@id="inpParid"]'
            ],
            [
                'pierce/#inpParid'
            ],
            [
                'aria/Parid'
            ]
        ], targetPage, { timeout, visible: true });
        await element.click({
            offset: {
                x: 32.796875,
                y: 13.46875,
            },
        });
    }
    {
        const targetPage = page;
        await scrollIntoViewIfNeeded([
            [
                '#inpParid'
            ],
            [
                'xpath///*[@id="inpParid"]'
            ],
            [
                'pierce/#inpParid'
            ],
            [
                'aria/Parid'
            ]
        ], targetPage, timeout);
        const element = await waitForSelectors([
            [
                '#inpParid'
            ],
            [
                'xpath///*[@id="inpParid"]'
            ],
            [
                'pierce/#inpParid'
            ],
            [
                'aria/Parid'
            ]
        ], targetPage, { timeout, visible: true });
        const inputType = await element.evaluate(el => el.type);
        if (inputType === 'select-one') {
            await changeSelectElement(element, boroughNum)
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
            await typeIntoElement(element, boroughNum);
        } else {
            await changeElementValue(element, boroughNum);
        }
    }
    {
        const targetPage = page;
        await scrollIntoViewIfNeeded([
            [
                '#inpTag'
            ],
            [
                'xpath///*[@id="inpTag"]'
            ],
            [
                'pierce/#inpTag'
            ],
            [
                'aria/Block *'
            ]
        ], targetPage, timeout);
        const element = await waitForSelectors([
            [
                '#inpTag'
            ],
            [
                'xpath///*[@id="inpTag"]'
            ],
            [
                'pierce/#inpTag'
            ],
            [
                'aria/Block *'
            ]
        ], targetPage, { timeout, visible: true });
        await element.click({
            offset: {
                x: 36.796875,
                y: 12.46875,
            },
        });
    }
    {
        const targetPage = page;
        await scrollIntoViewIfNeeded([
            [
                '#inpTag'
            ],
            [
                'xpath///*[@id="inpTag"]'
            ],
            [
                'pierce/#inpTag'
            ],
            [
                'aria/Block *'
            ]
        ], targetPage, timeout);
        const element = await waitForSelectors([
            [
                '#inpTag'
            ],
            [
                'xpath///*[@id="inpTag"]'
            ],
            [
                'pierce/#inpTag'
            ],
            [
                'aria/Block *'
            ]
        ], targetPage, { timeout, visible: true });
        const inputType = await element.evaluate(el => el.type);
        if (inputType === 'select-one') {
            await changeSelectElement(element, block)
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
            await typeIntoElement(element, block);
        } else {
            await changeElementValue(element, block);
        }
    }
    {
        const targetPage = page;
        await scrollIntoViewIfNeeded([
            [
                '#inpStat'
            ],
            [
                'xpath///*[@id="inpStat"]'
            ],
            [
                'pierce/#inpStat'
            ],
            [
                'aria/Lot'
            ]
        ], targetPage, timeout);
        const element = await waitForSelectors([
            [
                '#inpStat'
            ],
            [
                'xpath///*[@id="inpStat"]'
            ],
            [
                'pierce/#inpStat'
            ],
            [
                'aria/Lot'
            ]
        ], targetPage, { timeout, visible: true });
        await element.click({
            offset: {
                x: 22.109375,
                y: 12.46875,
            },
        });
    }
    {
        const targetPage = page;
        await scrollIntoViewIfNeeded([
            [
                '#inpStat'
            ],
            [
                'xpath///*[@id="inpStat"]'
            ],
            [
                'pierce/#inpStat'
            ],
            [
                'aria/Lot'
            ]
        ], targetPage, timeout);
        const element = await waitForSelectors([
            [
                '#inpStat'
            ],
            [
                'xpath///*[@id="inpStat"]'
            ],
            [
                'pierce/#inpStat'
            ],
            [
                'aria/Lot'
            ]
        ], targetPage, { timeout, visible: true });
        const inputType = await element.evaluate(el => el.type);
        if (inputType === 'select-one') {
            await changeSelectElement(element, lot)
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
            await typeIntoElement(element, lot);
        } else {
            await changeElementValue(element, lot);
        }
    }
    {
        const targetPage = page;
        const promises = [];
        promises.push(targetPage.waitForNavigation());
        await scrollIntoViewIfNeeded([
            [
                '#btSearch'
            ],
            [
                'xpath///*[@id="btSearch"]'
            ],
            [
                'pierce/#btSearch'
            ]
        ], targetPage, timeout);
        const element = await waitForSelectors([
            [
                '#btSearch'
            ],
            [
                'xpath///*[@id="btSearch"]'
            ],
            [
                'pierce/#btSearch'
            ]
        ], targetPage, { timeout, visible: true });
        await element.click({
            offset: {
                x: 41.875,
                y: 14.203125,
            },
        });
        await Promise.all(promises);
    }
    {
        const targetPage = page;
        const promises = [];
        promises.push(targetPage.waitForNavigation());
        await scrollIntoViewIfNeeded([
            [
                'li:nth-of-type(12) span'
            ],
            [
                'xpath///*[@id="sidemenu"]/ul/li[12]/a/span'
            ],
            [
                'pierce/li:nth-of-type(12) span'
            ],
            [
                'aria/      2023-2024 Final',
                'aria/[role="generic"]'
            ],
            [
                'text/2023-2024 Final'
            ]
        ], targetPage, timeout);
        const element = await waitForSelectors([
            [
                'li:nth-of-type(12) span'
            ],
            [
                'xpath///*[@id="sidemenu"]/ul/li[12]/a/span'
            ],
            [
                'pierce/li:nth-of-type(12) span'
            ],
            [
                'aria/      2023-2024 Final',
                'aria/[role="generic"]'
            ],
            [
                'text/2023-2024 Final'
            ]
        ], targetPage, { timeout, visible: true });
        await element.click({
            offset: {
                x: 109,
                y: 4,
            },
        });
        await Promise.all(promises);
    }
    {
        const targetPage = page;
        await scrollIntoViewIfNeeded([
            ['tr:nth-of-type(2) > td:nth-of-type(4)'],
            ['xpath///*[@id="Assessment Information"]/tbody/tr[2]/td[4]'],
            ['pierce/tr:nth-of-type(2) > td:nth-of-type(4)']
        ], targetPage, timeout);
        const element = await waitForSelectors([
            ['tr:nth-of-type(2) > td:nth-of-type(4)'],
            ['xpath///*[@id="Assessment Information"]/tbody/tr[2]/td[4]'],
            ['pierce/tr:nth-of-type(2) > td:nth-of-type(4)']
        ], targetPage, { timeout, visible: true });
        marketValue = await element.evaluate(el => el.textContent);
        console.log(marketValue); // print value for debugging
    }
    {
        const targetPage = page;
        await scrollIntoViewIfNeeded([
            ['tr:nth-of-type(3) > td:nth-of-type(4)'],
            ['xpath///*[@id="Assessment Information"]/tbody/tr[3]/td[4]'],
            ['pierce/tr:nth-of-type(3) > td:nth-of-type(4)']
        ], targetPage, timeout);
        const element = await waitForSelectors([
            ['tr:nth-of-type(3) > td:nth-of-type(4)'],
            ['xpath///*[@id="Assessment Information"]/tbody/tr[3]/td[4]'],
            ['pierce/tr:nth-of-type(3) > td:nth-of-type(4)']
        ], targetPage, { timeout, visible: true });
        actual_av = await element.evaluate(el => el.textContent); // adjust variable name as needed
        console.log(actual_av); // print value for debugging
    }
    {
        const targetPage = page;
        await scrollIntoViewIfNeeded([
            ['#datalet_div_5 tr:nth-of-type(3) > td:nth-of-type(3)'],
            ['xpath///*[@id="Assessment Information"]/tbody/tr[3]/td[3]'],
            ['pierce/#datalet_div_5 tr:nth-of-type(3) > td:nth-of-type(3)']
        ], targetPage, timeout);
        const element = await waitForSelectors([
            ['#datalet_div_5 tr:nth-of-type(3) > td:nth-of-type(3)'],
            ['xpath///*[@id="Assessment Information"]/tbody/tr[3]/td[3]'],
            ['pierce/#datalet_div_5 tr:nth-of-type(3) > td:nth-of-type(3)']
        ], targetPage, { timeout, visible: true });
        actual_av_land = await element.evaluate(el => el.textContent);  // adjust variable name as needed
        console.log(actual_av_land); // print value for debugging
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

    return {
        marketValue,
        actual_av_land,
        actual_av
    };
}
async function caspiomain(CIDvalue, accessToken) {
    try {
        const masterTableUrl = `https://c1abd578.caspio.com/rest/v2/tables/MasterTable/records?q.select=Name%2C%20Address1%2C%20Phone1%2C%20email&q.where=CID=${CIDvalue}`;
        const solarProcessUrl = `https://c1abd578.caspio.com/rest/v2/tables/SolarProcess/records?q.select=Contract_Price%2CPanel%2CSystem_Size_Sold%2CBlock%2CLot%2CBin%2CFlood_Zone%2CNYC_LI%2COath_Ecb_Violation%2CFlood_Zone%2CLandmark%2CBorough%2CStories%2CCB&q.where=CustID=${CIDvalue}`;
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
        const Contract_Price = combinedData.Contract_Price;
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
        let Zip = addressParts[3].trim();
        // Get the zip
        // Assuming that the address format is always "House, City, State, Zip"
        // Zip would be in the 3rd index position after splitting
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
    const accessToken = await caspioAuthorization();
    console.log("Successfully connected to Google Sheets!");
    console.log("Successfully connected to Caspio!");
    //const sheetData = await fetchSpreadsheetData(client);
    let sheetData = 1;
    for (let i = 0; i < sheetData; i++) {
        try {
            let Cust_ID = process.argv[2];
			//let Cust_ID = 7964;
            let caspioData = await caspiomain(Cust_ID, accessToken);
            console.log("Data: ", caspioData);
			let FullName=caspioData.FullName;
			let Lastname = caspioData.Lastname;
			const blankformPath = `/tmp/NYCBLANK.pdf`;
			const filePath = `/tmp/${FullName}-filled.pdf`;
			await downloadFile(drive, nycsformsfileId, blankformPath);
            const blankpdfpath = await PDFDocument.load(fs.readFileSync(blankformPath));
            const pdfDoc = blankpdfpath;
            const form = pdfDoc.getForm();
            const totalCost = Math.round(parseFloat(caspioData.Contract_Price));
            const stories = parseFloat(caspioData.Stories);
            let panel = caspioData.Panel;

			
			const contractPath = path.resolve('/tmp', `${FullName}-contract.pdf`); // define nycblanksdownloadedpath outside the functions
			//if (fs.existsSync(filePath)) {
            // await deleteFileWithRetry(filePath);
            //}
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
            //const dateofreport = moment().add(-7, 'days').format('MM.DD.YYYY');
            const dateofreport = moment().subtract(7, 'days').format('MM/DD/YYYY');
            let systemsize = caspioData.Size.toString();
            let Landmark = caspioData.Landmark;
			let mktvalue;
			let actualav;
			let actualavland;
			let market_value_rounded;
			let structure_value;
			let Percent;
			let improvement_percent;
			const yesCheckbox = form.getCheckBox('Flood checkbox');
            const noCheckbox = form.getCheckBox('No flood checkbox');
            // Function to format numbers as currency
const formatCurrency = async (number) => {
  return new Intl.NumberFormat('en-US', { 
    style: 'currency', 
    currency: 'USD',
    minimumFractionDigits: 0,
    maximumFractionDigits: 0
  }).format(number);
}
            const formatPercent = async (number) => {
                return `${number}%`;
            }
			if (Flood_Zone === 'Yes') {
                yesCheckbox.check();
                noCheckbox.uncheck();
			const { marketValue, actual_av_land, actual_av } = await fetchData(block, lot, boroughNum);
            mktvalue = marketValue.replace(/,/g, '');
			console.log(mktvalue);
            actualav = actual_av.replace(/,/g, '');
			console.log(actualav);
            actualavland = actual_av_land.replace(/,/g, '');
			console.log(actualavland);
            market_value_rounded = Math.round(parseFloat(mktvalue));
			console.log(market_value_rounded);
            structure_value = Math.round(parseFloat(market_value_rounded) * ((parseFloat(actualav) - parseFloat(actualavland)) / parseFloat(actualav)));
            Percent = Math.round((parseFloat(actualav) - parseFloat(actualavland)) / parseFloat(actualav) * 100);
            improvement_percent = Math.round(parseFloat(totalCost) / parseFloat(structure_value) * 100);
			
            } else if (Flood_Zone === 'NA') {
                noCheckbox.check();
                yesCheckbox.uncheck();
            }
            const fields = {
                Labor: await formatCurrency(Math.round(0.20 * totalCost)),
                Materials: await formatCurrency(Math.round(0.30 * totalCost)),
                Equipment: await formatCurrency(Math.round(0.25 * totalCost)),
                DA_cost: await formatCurrency(Math.round(0.25 * totalCost)),
                DA_Labor: 0,   // Initialize dependent fields
                L_M_E_cost: 0, // Initialize dependent fields
                D_cost: 0,     // Initialize dependent fields
                A_cost: 0,     // Initialize dependent fields
                First: caspioData.Firstname,
                Last: Lastname,
                House: caspioData.House,
                Street: caspioData.Street,
                Borough: Borough,
                Block: block,
                Lot: lot,
                Bin: caspioData.Bin,
                Stories: caspioData.Stories,
                CB: caspioData.CB,
                lead_violation: caspioData.lead === '0' || caspioData.lead === '' || caspioData.lead === undefined ? "No lead violation exist" : "lead violation exist",
                Zip: caspioData.Zip,
                Kw: systemsize,
                Panel_Model: caspioData.Panel,
                Brand: panelbrand,
                Totalcost: await formatCurrency(totalCost),
                Market_Value: await formatCurrency(market_value_rounded),
                Actual_Av: await formatCurrency(actualav),
                Actual_Av_Land: await formatCurrency(actualavland),
                ans: await formatCurrency(structure_value),
                percent: await formatPercent(Percent),
                small_Percent: await formatPercent(improvement_percent),
                Date: dateToday,
                Date4_af_date: dateofreport,
                Address: caspioData.Address,
                Improvement_cost: await formatCurrency(Math.round(1.3 * totalCost)),
                Natural_gas_cost: await formatCurrency(Math.round(1500 + 13.00 * caspioData.Size)),
                Natural_gas_savings: await formatCurrency(Math.round(400 + 13.00 * caspioData.Size)),
                Electric_cost: await formatCurrency(Math.round(1350 * caspioData.Size * 0.30)),
                Electric_Savings: await formatCurrency(Math.round(1350 * caspioData.Size * 0.30 - 100)),
                Oil_Cost: await formatCurrency(Math.round(1600 + 15.00 * caspioData.Size)),
                Oil_savings: await formatCurrency(Math.round(1600 + 15.00 * caspioData.Size - 1500 - 10.00 * caspioData.Size)),
                Fullname: FullName,

            };
				
            // Handling panel address, city, state, zip
            //const panelBrand = process.env.panelbrand;
            const panelAddresses = {
                REC: { address: "1820 Gateway Drive Suite 170", city: "San Mateo", state: "CA", zip: "94404" },
                Solaria: { address: "45700 Northport Loop E", city: "Fremont", state: "CA", zip: "94538" },
                'S-energy': { address: "1170 N Gilbert St", city: "Anaheim", state: "CA", zip: "92801" },
                Certainteed: { address: "20 Moores Rd", city: "Malvern", state: "PA", zip: "19355" },
                Trina: { address: "1425 K Street, N.W., Suite 1000", city: "Washington", state: "DC", zip: "20005" }
            };
            const panelDetails = panelAddresses[panelbrand] || { address: "", city: "", state: "", zip: "" };
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
            fields.DA_Labor = await formatCurrency(Math.round(Math.round(0.25 * totalCost) + Math.round(0.20 * totalCost)));
            fields.L_M_E_cost = await formatCurrency(Math.round(Math.round(0.20 * totalCost) + Math.round(0.30 * totalCost) + Math.round(0.25 * totalCost)));
            fields.D_cost = await formatCurrency(Math.round(Math.round(0.25 * totalCost) * 0.60));
            fields.A_cost = await formatCurrency(Math.round(Math.round(0.25 * totalCost) - Math.round(Math.round(0.25 * totalCost) * 0.60)));
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
            fs.writeFileSync(`/tmp/${FullName}-filed.pdf`, pdfBytes);
            // Wait for some time to ensure the file is completely written
            await new Promise(resolve => setTimeout(resolve, 10000)); // 5 seconds delay
            let opts = {
                format: 'png',
                out_dir: path.dirname(contractPath),
                out_prefix: path.basename(contractPath, path.extname(contractPath)),
                page: 17
            }
			let pngPath;
			let filedPath;
            pdfPoppler.convert(contractPath, opts)
                .then(async res => {
                    pngPath = path.join(opts.out_dir, `${opts.out_prefix}-${opts.page}.${opts.format}`);

                    if (fs.existsSync(pngPath)) {
                        const inputPng = fs.readFileSync(pngPath);
                        const outputPng = await sharp(inputPng).extract({
                            left: 191,  // x-coordinate
                            top: 660,   // y-coordinate
                            width: 220, // width
                            height: 83 // height
                        }).toBuffer();
                        await new Promise(resolve => setTimeout(resolve, 5000)); // 5 seconds delay
                        filedPath = path.resolve(__dirname, `${FullName}-filed.pdf`); // define pdfPath outside the functions
                        const pdfPath = await PDFDocument.load(fs.readFileSync(filedPath));
                        const imagePage = await pdfPath.embedPng(outputPng);
                        const pagesData = [
                            { pageNum: 2, imageX: 1.2947 * 72, imageY: 3.8382 * 72, imageWidth: 1.1926 * 72, imageHeight: 0.314 * 72 },
                            { pageNum: 6, imageX: 5.0931 * 72, imageY: 1.404 * 72, imageWidth: 1.1926 * 72, imageHeight: 0.314 * 72 },
                            { pageNum: 8, imageX: 2.1453 * 72, imageY: 5.6732 * 72, imageWidth: 0.8449 * 72, imageHeight: 0.1909 * 72 },
                            { pageNum: 11, imageX: 1.4712 * 72, imageY: 8.2763 * 72, imageWidth: 1.0294 * 72, imageHeight: 0.2521 * 72 }, // changed position and dimensions
                            { pageNum: 18, imageX: 3.6861 * 72, imageY: 3.8435 * 72, imageWidth: 1.1926 * 72, imageHeight: 0.314 * 72 },
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
                        fs.writeFileSync(filePath, pdfBytes);
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
            let PermitfolderIDs;
			
				let METER = `ACC${i + 1}`;
                if (folderData.Result && folderData.Result.length > 0) {
                    PermitfolderIDs = folderData.Result[0][FolderID];
					
                }
                if (!PermitfolderIDs) {                             
                            PermitfolderIDs = '1cg70dBWgSqTCwO94dmN4Gb9KbWA0v5bk';
                            
                }
            await new Promise(resolve => setTimeout(resolve, 15000)); // 5 seconds delay
            const signedFormsFolderId= await uploadFile(client, filePath, PermitfolderIDs, Oath_Ecb_Violation, Landmark, Flood_Zone,FullName,Lastname);
			 console.log(signedFormsFolderId);

			let drivelink = `https://drive.google.com/drive/u/0/folders/${signedFormsFolderId}`;
			            let solarprocessPUT = {
                "nycinitialfolderlink": drivelink
            }
			let sendurl = `https://c1abd578.caspio.com/rest/v2/tables/SolarProcess/records?q.where=CustID=${Cust_ID}`;		
			await fetchCaspioData(accessToken, sendurl, solarprocessPUT, 'PUT'); 
            //await deleteFileWithRetry(filePath);
            await new Promise(resolve => setTimeout(resolve, 5000)); // 5 seconds delay
            fs.unlinkSync(filePath);
            fs.unlinkSync(pngPath);
            fs.unlinkSync(contractPath);
			fs.unlinkSync(filedPath);
    } catch (error) {
        console.error(error);
       }
    }
})().catch(err => {
    console.error(err);
    process.exit(1);
});