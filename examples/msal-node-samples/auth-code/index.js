/*
 * Copyright (c) Microsoft Corporation. All rights reserved.
 * Licensed under the MIT License.
 */
const express = require("express");
const msal = require('@azure/msal-node');
const { ExchangeService, OAuthCredentials, WellKnownFolderName, ItemView, PropertySet, BasePropertySet, EmailMessageSchema, Uri, EmailMessage, FindItemsResults } = require("ews-javascript-api");

const SERVER_PORT = process.env.PORT || 3000;

/** @type {msal.AuthenticationResult} */
let tokens = {};

// Before running the sample, you will need to replace the values in the config, 
// including the clientSecret
const config = {
    auth: {
        clientId: process.env.CLIENT_ID,
        authority: "https://login.microsoftonline.com/common",
        clientSecret: process.env.CLIENT_SECRET
    },
    system: {
        loggerOptions: {
            loggerCallback(loglevel, message, containsPii) {
                console.log(message);
            },
            piiLoggingEnabled: false,
            logLevel: msal.LogLevel.Verbose,
        }
    }
};

// Create msal application object
const cca = new msal.ConfidentialClientApplication(config);

// Create Express App and Routes
const app = express();

// pretty print json
app.set('json spaces', 2)

app.get('/', async (req, res) => {

    if (tokens) {
        try {
            const svc = new ExchangeService();
            svc.Url = new Uri("https://outlook.office365.com/ews/exchange.asmx");
            svc.Credentials = new OAuthCredentials(tokens.accessToken);
            const view = new ItemView(20);
            view.PropertySet = new PropertySet(BasePropertySet.IdOnly, [EmailMessageSchema.Subject, EmailMessageSchema.DateTimeReceived, EmailMessageSchema.From, EmailMessageSchema.HasAttachments]);

            /** @type {FindItemsResults<EmailMessage>} */
            const result = await svc.FindItems(WellKnownFolderName.Inbox, view);
            // console.log(result);
            res.json(result.Items.map(({ Subject, From, DateTimeReceived, HasAttachments }) => ({ Subject, From: From.Address, DateTimeReceived: DateTimeReceived.toString(), HasAttachments })));
            res.end();
        } catch (error) {
            console.log(error);
            console.log("Error in using credential, will try to get new token")
        }
    }

    // "https://outlook.office.com/.default"
    const authCodeUrlParameters = {
        scopes: ["user.read", "EWS.AccessAsUser.All"],
        redirectUri: "http://localhost:3000/redirect",
    };

    // get url to sign user in and consent to scopes needed for application
    cca.getAuthCodeUrl(authCodeUrlParameters).then((response) => {
        res.redirect(response);
    }).catch((error) => console.log(JSON.stringify(error)));
});

app.get('/redirect', (req, res) => {
    const tokenRequest = {
        code: req.query.code,
        //  "user.read", "offline_access", "EWS.AccessAsUser.All"
        scopes: ["https://outlook.office.com/.default"],
        redirectUri: "http://localhost:3000/redirect",
    };

    cca.acquireTokenByCode(tokenRequest).then((response) => {
        console.log("\nResponse: \n:", response);
        tokens = { ...response };
        res.send(`<a href="/">Token Received, go to Home to see mailbox</a>`);
    }).catch((error) => {
        console.log(error);
        res.status(500).send(error);
    });
});


app.listen(SERVER_PORT, () => console.log(`Msal Node Auth Code Sample app listening on port ${SERVER_PORT}!`))
