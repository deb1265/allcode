const express = require('express');
const bodyParser = require('body-parser');
const { exec } = require('child_process');
const path = require('path');
const cors = require('cors');
const app = express();

app.use(cors());
app.use(bodyParser.json());
app.use(express.static(path.join(__dirname, 'public')));

app.locals.progressUpdate = null;

app.get('/', (req, res) => {
    res.sendFile(path.join(__dirname, 'public', 'index.html'));
});

// Endpoint for running mainforclosing.js
app.post('/fetchData1', (req, res) => {
    const custID = req.body.custID;
    runScript('nycanalysis.js', custID,1, res);
});

// Endpoint for running contract.js
app.post('/fetchData2', (req, res) => {
    const custID = req.body.custID;
    runScript('creatingfolder.js', custID,2, res);
});

// Endpoint for running interconnect.js
app.post('/fetchData3', (req, res) => {
    const custID = req.body.custID;
    runScript('nycfills.js', custID,3, res);
});

// Endpoint for running interconnect.js
app.post('/fetchData4', (req, res) => {
    const custID = req.body.custID;
    runScript('changeorder.js', custID,4, res);
});

// Endpoint for running interconnect.js
app.post('/fetchData5', (req, res) => {
    const custID = req.body.custID;
    runScript('mainforclosing.js', custID,5, res);
});

// Endpoint for running interconnect.js
app.post('/fetchData6', (req, res) => {
    const custID = req.body.custID;
    runScript('mainforaccountfill.js', custID,6, res);
});

// Endpoint for running interconnect.js
app.post('/fetchData7', (req, res) => {
    const custID = req.body.custID;
    runScript('maininterconnect.js', custID,7, res);
});

// Endpoint for running interconnect.js
app.post('/fetchData8', (req, res) => {
    const custID = req.body.custID;
    runScript('PAM.js', custID,8, res);
});

// Endpoint for running interconnect.js
app.post('/fetchData9', (req, res) => {
    const custID = req.body.custID;
    runScript('PTOcloseout.js', custID,9, res);
});

// Endpoint for running interconnect.js
app.post('/fetchData10', (req, res) => {
    const custID = req.body.custID;
    runScript('mainforshade.js', custID,10, res);
});

// Endpoint for running interconnect.js
app.post('/fetchData11', (req, res) => {
    const custID = req.body.custID;
    runScript('nyserda.js', custID,11, res);
});

// Endpoint for running interconnect.js
app.post('/fetchData12', (req, res) => {
    const custID = req.body.custID;
    runScript('contract.js', custID,12, res);
});
// Function to handle the execution of scripts with progress updates
function runScript(scriptName, custID, sectionNumber, res) {
    const process = exec(`node ${scriptName} "${custID}"`);
    
    process.stdout.on('data', (data) => {
        sendProgressUpdate(sectionNumber, data);
    });

    process.on('close', (code) => {
        sendProgressUpdate(sectionNumber, `Completed with code ${code}`);
        res.json({ message: `Script ${scriptName} executed with CustID: ${custID}` });
    });

    process.on('error', (error) => {
        console.error(`Error: ${error}`);
        res.status(500).json({ error: `Error executing ${scriptName} for CustID: ${custID}` });
    });
}

// Function to send progress updates
function sendProgressUpdate(sectionNumber, message) {
    if (app.locals.progressUpdate) {
        app.locals.progressUpdate.write(`data: ${JSON.stringify({ sectionNumber, message })}\n\n`);
    }
}

// Endpoint for server-sent events (SSE)
app.get('/progress', (req, res) => {
    res.setHeader('Content-Type', 'text/event-stream');
    res.setHeader('Cache-Control', 'no-cache');
    res.setHeader('Connection', 'keep-alive');
    res.flushHeaders();

    const keepAlive = setInterval(() => {
        res.write(':keep-alive\n\n');
    }, 20000);

    app.locals.progressUpdate = res;

    req.on('close', () => {
        console.log('Client closed connection');
        clearInterval(keepAlive);
        app.locals.progressUpdate = null;
    });

    req.on('error', (error) => {
        console.error('SSE connection error:', error);
        clearInterval(keepAlive);
        app.locals.progressUpdate = null;
    });
});

const PORT = 3001;
app.listen(PORT, () => {
    console.log(`Server running on port ${PORT}`);
});
