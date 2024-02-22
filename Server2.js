const axios = require('axios');
const { ConfidentialClientApplication } = require('@azure/msal-node');

// App registration details
const clientId = 'ddf32454-36a6-4a4b-a808-28325fd7a27d';
const clientSecret = 'bsx8Q~1WH2RhwmugwQVKBvC5BisEf5Se.w.pkbb5';
const tenantId = '986dc901-f440-4014-9902-1fd094d52323';

// Scopes required for Microsoft Graph API
const scopes = ['https://graph.microsoft.com/.default'];

// Microsoft Graph token endpoint
const tokenEndpoint = `https://login.microsoftonline.com/${tenantId}/oauth2/v2.0/token`;

// Microsoft Graph API endpoint for creating an online meeting
const graphApiEndpoint = 'https://graph.microsoft.com/v1.0/me/onlineMeetings';

// Create a confidential client application
const cca = new ConfidentialClientApplication({
    auth: {
        clientId,
        clientSecret,
        authority: `https://login.microsoftonline.com/${tenantId}`,
    },
});

// Function to format the date and time
function formatDateTime(dateString) {
    return new Date(dateString).toISOString();
}

// Set the start and end time for the meeting (tomorrow 1:00 AM to 1:30 AM)
const startTime = formatDateTime('2024-02-13T01:00:00');  // Adjust the date accordingly
const endTime = formatDateTime('2024-02-13T01:30:00');    // Adjust the date accordingly

// Define the meeting payload
const meetingPayload = {
    startDateTime: startTime,
    endDateTime: endTime,
    subject: 'Morning Meeting',
};

// Acquire a token
cca.acquireTokenByClientCredential({ scopes })
    .then((response) => {
        // Token acquired successfully
        const accessToken = response.accessToken;
        console.log(`Access Token: ${accessToken}`);

        // Make a request to Microsoft Graph API to create the online meeting
        axios.post(graphApiEndpoint, meetingPayload, {
            headers: {
                Authorization: `Bearer ${accessToken}`,
                'Content-Type': 'application/json',
            },
        })
        .then((graphResponse) => {
            // Successfully created the online meeting
            console.log('Meeting created successfully. Join URL:', graphResponse.data.joinWebUrl);
        })
        .catch((graphError) => {
            // Handle error from Microsoft Graph API request
            console.error('Error creating meeting:', graphError.message);
        });
    })
    .catch((error) => {
        // Handle error acquiring token
        console.error(`Error acquiring token: ${error.message}`);
    });