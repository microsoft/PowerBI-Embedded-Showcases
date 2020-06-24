// API Endpoint to get the JSON response of Embed Url, Embed token and reportId
const reportUrl = "https://powerbiplaygroundbe.azurewebsites.net/api/Reports/SampleReport";

// Set the report refresh token timer
let reportRefreshTokenTimer = 0;

// This function will make the AJAX request to the endpoint and get the JSON response which it will set in the sessions
function populateEmbedConfigIntoCurrentSession(url, updateCurrentToken) {

    // This returns the JSON response
    return $.getJSON(url, function(embedConfig) {

        // Set the embedToken, embedUrl, reportId
        setConfig(embedConfig.embedToken.token, embedConfig.embedUrl, embedConfig.id);
        if (updateCurrentToken) {

            // Get the reference to the embedded element
            let reportContainer = $("#report-container")[0];
            let embedContainer = powerbi.get(reportContainer);

            if (embedContainer) {
                embedContainer.setAccessToken(embedConfig.embedToken.token);
            }
        }
        tokenExpirationRefreshListener(embedConfig.minutesToExpiration, url, "Report");
    });
}

function tokenExpirationRefreshListener(minutesToExpiration, url, entityType) {

    // Used to count the milliseconds after which API call will be made
    const updateAfterMilliSeconds = (minutesToExpiration - 2) * 60 * 1000;

    // Set the tokenRefresh timer to count the seconds and request the JSON again when token expires
    setTokenRefreshListener(updateAfterMilliSeconds, reportRefreshTokenTimer, url, entityType);
}

// Checking the remaining time and calling the API again
function setTokenRefreshListener(updateAfterMilliSeconds, refreshTokenTimer, url, entityType) {
    if (refreshTokenTimer) {
        console.log("step current " + entityType + " Embed Token update threads.");
        clearTimeout(refreshTokenTimer);
    }
    console.log(entityType + " Embed Token will be updated in " + updateAfterMilliSeconds + " milliseconds.");

    // Making the call when token expires
    refreshTokenTimer = setTimeout(function() {
        if (url) {
            populateEmbedConfigIntoCurrentSession(url, true /* updateCurrentToken */ ); // It suggests that the token is expired so API request is made
        }
    }, updateAfterMilliSeconds);
}

// Get the data from the API and pass it to the front
function loadSampleReportIntoSession() {

    // Call the function for the first time to call the API, set the sessions values and return the response to the front-end
    return populateEmbedConfigIntoCurrentSession(reportUrl, false /* updateCurrentToken */ );
}

// Set the embed config in globals
function setConfig(accessToken, embedUrl, reportId) {

    // Fill the global object
    reportConfig.accessToken = accessToken;
    reportConfig.embedUrl = embedUrl;
    reportConfig.reportId = reportId;
}