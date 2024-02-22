const axios = require('axios');
const { ConfidentialClientApplication } = require('@azure/msal-node');

// App registration details from Azure Portal
const clientId = 'ddf32454-36a6-4a4b-a808-28325fd7a27d';
const clientSecret = 'bsx8Q~1WH2RhwmugwQVKBvC5BisEf5Se.w.pkbb5';
const tenantId = '986dc901-f440-4014-9902-1fd094d52323';

// Scopes required for Microsoft Graph API
const scopes = ['https://graph.microsoft.com/.default'];

// Microsoft Graph token endpoint
const tokenEndpoint = `https://login.microsoftonline.com/${tenantId}/oauth2/v2.0/token`;

// Create a confidential client application
const cca = new ConfidentialClientApplication({
    auth: {
        clientId,
        clientSecret,
        authority: `https://login.microsoftonline.com/${tenantId}`,
    },
});

// Acquire a token
cca.acquireTokenByClientCredential({ scopes })
    .then((response) => {
        // Token acquired successfully
        const accessToken = response.accessToken;
        console.log(`Access Token: ${accessToken}`);
    })
    .catch((error) => {
        // Handle error
        console.error(`Error acquiring token: ${error.message}`);
    });