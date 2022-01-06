/*
 * Copyright (c) Microsoft Corporation. All rights reserved.
 * Licensed under the MIT License.
 */

var msal = require('@azure/msal-node');
var graph = require('@microsoft/microsoft-graph-client');
require('isomorphic-fetch');

// Build MSAL ClientApplication Configuration object
const clientConfig = {
    auth: {
        "clientId": process.env.MicrosoftAppId,
        "authority": `https://login.microsoftonline.com/${process.env.MicrosoftAppTenantId}`,
        "clientSecret": process.env.MicrosoftAppPassword
    }
};

function getClientCredentialsToken(cca) {
    // With client credentials flows permissions need to be granted in the portal by a tenant administrator. 
    // The scope is always in the format "<resource>/.default"
    const clientCredentialRequest = {
        scopes: ["https://graph.microsoft.com/.default"],
        skipCache: true // (optional) this skips the cache and forces MSAL to get a new token from Azure AD
    };

    return cca
        .acquireTokenByClientCredential(clientCredentialRequest)
        .then((response) => {
            // Uncomment to see the successful response logged
            //console.log("Response: ", response);
            return response.accessToken;
        }).catch((error) => {
            // Uncomment to see the errors logges
            console.log(JSON.stringify(error));
            throw error;
        });
}

// Execute sample application with the configured MSAL PublicClientApplication
const getAccessToken = async () => {
    const confidentialClientApplication = new msal.ConfidentialClientApplication(clientConfig);
    return await getClientCredentialsToken(confidentialClientApplication)
        .then((token) => {
            return token;
        });
}

// Create a MS Graph client using MSAL access token
function getAuthenticatedClient(accessToken) {
    // Initialize Graph client
    const client = graph.Client.init({
        // Use the provided access token to authenticate requests
        authProvider: (done) => {
        done(null, accessToken);
        }
    });
    return client;
}

const getTeamsAppID = async (accessToken, AADUserID, externalAppID) => {
    const client = getAuthenticatedClient(accessToken);
    const response = await client
                        .api(`/users/${AADUserID}/teamwork/installedApps`)
                        .expand(`teamsApp`)
                        .filter(`teamsApp/externalId eq '${externalAppID}'`)
                        .get();
    if (response.value.length != 0) {
        return response.value[0].teamsApp;
    }
    else return [];
}

module.exports = { getAccessToken, getTeamsAppID };