/*
 * Copyright (c) Microsoft Corporation. All rights reserved.
 * Licensed under the MIT License.
 */
//https://docs.microsoft.com/en-us/azure/active-directory/develop/msal-client-application-configuration
//https://docs.microsoft.com/en-us/azure/active-directory/develop/authentication-national-cloud#azure-ad-authentication-endpoints
const express = require("express");

const msal = require('@azure/msal-node');
const { Client } = require("@microsoft/microsoft-graph-client");
const fetch = require("node-fetch");
const path = require('path');
var bodyParser = require('body-parser')

const SERVER_PORT = process.env.PORT || 3000;
const REDIRECT_URI = "http://localhost:3000/redirect";

// Before running the sample, you will need to replace the values in the config, 
// including the clientSecret
const config = {
    auth: {
        clientId: "8a792f49-ae0d-4b9b-92d2-614fcba43bea",
        authority: "https://login.microsoftonline.com/common",
        clientSecret: "PHx7Q~l~OordXC3.Yf6UY1tv.9vyRtTdoHuhz"
    },
    system: {
        loggerOptions: {
            loggerCallback(loglevel, message, containsPii) {
                console.log(message);
            },
            piiLoggingEnabled: false,
            logLevel: msal.LogLevel.Verbose,
        }
    }
};

// Create msal application object
const pca = new msal.ConfidentialClientApplication(config);

// Create Express App and Routes
const app = express();

function callMSGraph(uri, versions, endpoint, accessToken) {
    
    var headers = new fetch.Headers();
    const bearer = `Bearer ${accessToken}`;
    headers.append("Authorization", bearer);
  
    const options = {
        method: "GET",
        headers: headers
    };
  
    console.log('request made to Graph API at: ' + new Date().toString());
  
    fetch(uri+"/"+versions+"/"+endpoint, options)
        .then(response => response.json())
        .then(response => console.log(response, endpoint))
        .catch(error => console.log(error));
}

app.get('/', (req, res) => {
    const authCodeUrlParameters = {
        scopes: ["user.read"],
        redirectUri: REDIRECT_URI,
    };

    // get url to sign user in and consent to scopes needed for application
    pca.getAuthCodeUrl(authCodeUrlParameters).then((response) => {
        res.redirect(response);
    }).catch((error) => console.log(JSON.stringify(error)));
});

app.get('/redirect', (req, res) => {
    const tokenRequest = {
        code: req.query.code,
        scopes: ["user.read"],
        redirectUri: REDIRECT_URI,
    };

    pca.acquireTokenByCode(tokenRequest).then((response) => {
        console.log("\nResponse: \n:", response);
        res.sendStatus(200);
        callMSGraph("https://graph.microsoft.com", "v1.0", "me", response.accessToken)
    }).catch((error) => {
        console.log(error);
        res.status(500).send(error);
    });
});
var jsonParser = bodyParser.json()

app.post('/GraphCall',jsonParser, (req, res) => {
    console.log('Got body:', req.body);
    res.sendStatus(200);
});

// app.use(express.staticProvider(__dirname + '/build'));
app.use(express.static('build'))

app.get('/Home', function(req, res) {
    res.sendFile(path.join(__dirname, '/build/index.html'));
});

app.listen(SERVER_PORT, () => console.log(`Msal Node Auth Code Sample app listening on port ${SERVER_PORT}!`))
