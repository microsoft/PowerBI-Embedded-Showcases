// API Endpoint to get the JSON response of Embed Url, Embed token and reportId
const reportUrl = "https://playgroundbe-bck-1.azurewebsites.net/Reports/ThemesReportV2";

// Set the report refresh token timer
const reportRefreshTokenTimer = 0;

// This function will make the AJAX request to the endpoint and get the JSON response which it will set in the sessions
function populateEmbedConfigIntoCurrentSession(url, updateCurrentToken) {

    try {
        let configFromParentWindow = window.parent.showcases.personalizeReportDesign;
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

            handleNewEmbedConfig(url, embedConfig, updateCurrentToken);
        }

        return;
    } catch (error) {
        console.error(error);
    }

    // This returns the JSON response
    return $.getJSON(url, function (embedConfig) {
        handleNewEmbedConfig(url, embedConfig, updateCurrentToken);
    });
}

function tokenExpirationRefreshListener(minutesToExpiration, url, entityType) {

    // Used to count the milliseconds after which API call will be made
    const updateAfterMilliSeconds = (minutesToExpiration - 2) * 60 * 1000;

    // Set the tokenRefresh timer to count the seconds and request the JSON again when token expires
    setTokenRefreshListener(updateAfterMilliSeconds, reportRefreshTokenTimer, url, entityType);
}

function handleNewEmbedConfig(tokenRefreshUrl, embedConfig, updateCurrentToken) {
    // Set the embedToken, embedUrl, reportId
    setConfig(embedConfig.EmbedToken.Token, embedConfig.EmbedUrl, embedConfig.Id);
    if (updateCurrentToken) {

        // Get the reference to the embedded element
        let reportContainer = $("#report-container")[0];
        let embedContainer = powerbi.get(reportContainer);

        if (embedContainer) {
            embedContainer.setAccessToken(embedConfig.EmbedToken.Token);
        }
    }
    tokenExpirationRefreshListener(embedConfig.MinutesToExpiration, tokenRefreshUrl, "Report");
}

// Checking the remaining time and calling the API again
function setTokenRefreshListener(updateAfterMilliSeconds, refreshTokenTimer, url, entityType) {
    if (refreshTokenTimer) {
        console.log("step current " + entityType + " Embed Token update threads.");
        clearTimeout(refreshTokenTimer);
    }
    console.log(entityType + " Embed Token will be updated in " + updateAfterMilliSeconds + " milliseconds.");

    // Making the call when token expires
    refreshTokenTimer = setTimeout(function () {
        if (url) {
            populateEmbedConfigIntoCurrentSession(url, true /* updateCurrentToken */); // It suggests that the token is expired so API request is made
        }
    }, updateAfterMilliSeconds);
}

// Get the data from the API and pass it to the front
function loadThemesShowcaseReportIntoSession() {

    // Call the function for the first time to call the API, set the sessions values and return the response to the front-end
    return populateEmbedConfigIntoCurrentSession(reportUrl, false /* updateCurrentToken */);
}

// Set the embed config in globals
function setConfig(accessToken, embedUrl, reportId) {

    // Fill the global object
    reportConfig.accessToken = accessToken;
    reportConfig.embedUrl = embedUrl;
    reportConfig.reportId = reportId;
}
