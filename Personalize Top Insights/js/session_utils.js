// API Endpoint to get the JSON response of Embed URL, Embed Token and reportId
const layoutShowcaseReportEndpoint = "https://aka.ms/layoutReportEmbedConfig";

// Set minutes before the access token should get refreshed
const minutesToRefreshBeforeExpiration = 2;

// Store the amount of time left for refresh token
let tokenExpiration;

// This function will make the AJAX request to the endpoint and get the JSON response which it will set in the sessions
function populateEmbedConfigIntoCurrentSession(updateCurrentToken) {

    try {
        let configFromParentWindow = window.parent.showcases.personalizeTopInsights;
        if (configFromParentWindow) {
            let diffMs = new Date(configFromParentWindow.expiration) - new Date();
            let diffMins = Math.round(((diffMs % 86400000) % 3600000) / 60000);

            embedConfig = {
                EmbedUrl: configFromParentWindow.embedUrl,
                EmbedToken: {
                    Token: configFromParentWindow.token
                },
                Id: configFromParentWindow.id,
                MinutesToExpiration: diffMins,
            };

            handleNewEmbedConfig(embedConfig, updateCurrentToken);
        }

        return;
    } catch (error) {
        console.error(error);
    }

    // This returns the JSON response
    return $.getJSON(layoutShowcaseReportEndpoint, function (embedConfig) {
        handleNewEmbedConfig(embedConfig, updateCurrentToken);
    });
}

function handleNewEmbedConfig(embedConfig, updateCurrentToken) {

    // Set Embed Token, Embed URL and Report Id
    setConfig(embedConfig.EmbedToken.Token, embedConfig.EmbedUrl, embedConfig.Id);
    if (updateCurrentToken) {

        // Get the reference to the embedded element
        const reportContainer = $("#report-container").get(0);
        if (reportContainer) {
            const report = powerbi.get(reportContainer);
            report.setAccessToken(embedConfig.EmbedToken.Token);
        }
    }

    // Get the milliseconds after token will get refreshed
    tokenExpiration = embedConfig.MinutesToExpiration * 60 * 1000;

    // Set the tokenRefresh timer to count the seconds and request the JSON again when token expires
    setTokenExpirationListener();
}

// Check the remaining time and call the API again
function setTokenExpirationListener() {

    const safetyInterval = minutesToRefreshBeforeExpiration * 60 * 1000;

    // Time until token refresh in milliseconds
    const timeout = tokenExpiration - safetyInterval;
    if (timeout <= 0) {
        populateEmbedConfigIntoCurrentSession(true /* updateCurrentToken */);
    }
    else {
        console.log(`Report Embed Token will be updated in ${timeout} milliseconds.`);
        setTimeout(function () {
            populateEmbedConfigIntoCurrentSession(true /* updateCurrentToken */)
        }, timeout);
    }
}

// Add a listener to make sure token is updated after tab was inactive
document.addEventListener("visibilitychange", function () {
    // Check the access token when the tab is visible
    if (!document.hidden) {
        setTokenExpirationListener();
    }
});

// Get the data from the API and pass it to the front-end
function loadLayoutShowcaseReportIntoSession() {

    // Call the function for the first time to call the API, set the sessions values and return the response to the front-end
    return populateEmbedConfigIntoCurrentSession(false /* updateCurrentToken */);
}

// Set the embed config in global object
function setConfig(accessToken, embedUrl, reportId) {

    // Fill the global object
    reportConfig.accessToken = accessToken;
    reportConfig.embedUrl = embedUrl;
    reportConfig.reportId = reportId;
}