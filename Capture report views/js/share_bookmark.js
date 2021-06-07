// ----------------------------------------------------------------------------
// Copyright (c) Microsoft Corporation.
// Licensed under the MIT license.
// ----------------------------------------------------------------------------

// Set props for accessibility insights
function setReportAccessibilityProps(report) {
    report.setComponentTitle("Playground showcase sample report");
    report.setComponentTabIndex(0);
}

$(document).ready(function () {

    // Bootstrap the bookmark container
    powerbi.bootstrap(bookmarkContainer, reportConfig);

    embedSharedBookmarkReport();
});

// Embed shared report with bookmark on load
async function embedSharedBookmarkReport() {

    // Load sample report properties into session
    await loadSampleReportIntoSession()

    // Get models. models contains enums that can be used
    const models = window["powerbi-client"].models;

    // Use View permissions
    let permissions = models.Permissions.View;

    // Get the bookmark name from url param
    let bookmarkName = getBookmarkNameFromURL();

    // Get the bookmark state from local storage
    // any type of database can be used
    let bookmarkState = localStorage.getItem(bookmarkName);

    // Embed configuration used to describe the what and how to embed
    // This object is used when calling powerbi.embed
    // This also includes settings and options such as filters
    // You can find more information at https://github.com/Microsoft/PowerBI-JavaScript/wiki/Embed-Configuration-Details
    let config = {
        type: "report",
        tokenType: models.TokenType.Embed,
        accessToken: reportConfig.accessToken,
        embedUrl: reportConfig.embedUrl,
        id: reportConfig.reportId,
        permissions: permissions,
        settings: {
            panes: {
                filters: {
                    visible: false
                },
                pageNavigation: {
                    visible: false
                }
            }
        },
        layoutType: models.LayoutType.Custom,
        customLayout: {
            displayOption: models.DisplayOption.FitToWidth
        },
        // Adding bookmark attribute will apply the bookmark on load
        bookmark: {
            state: bookmarkState
        }
    };

    // Embed the report and display it within the div container
    bookmarkShowcaseState.report = powerbi.embed(bookmarkContainer, config);

    // For accessibility insights
    setReportAccessibilityProps(bookmarkShowcaseState.report);

    bookmarkShowcaseState.report.on("loaded", function () {

        // Hide the loader and display the report
        overlay.addClass(INVISIBLE);
        $("#share-bookmark").addClass(VISIBLE);
        bookmarkShowcaseState.report.off("loaded");
    });
}

// Get the bookmark name from url "id" argument
function getBookmarkNameFromURL() {
    let url = (window.location != window.parent.location) ?
        document.referrer :
        document.location.href;

    const results = regex.exec(url);

    // It can take id parameter from the URL using ? or &
    // If ? or & is not in the URL, returns NULL
    if (!results) { return null };

    // Returns Empty String if id parameter's value is not specified
    if (!results[2]) { return "" };

    // Returns the ID of the Bookmark
    return decodeURIComponent(results[2]);
}