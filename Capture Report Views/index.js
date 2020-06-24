// Local report config object
let embedConfig = {
    accessToken: "",
    embedUrl: "",
    reportId: ""
};

// Set the report embed configurations from global object
function setEmbedConfig() {

    // Set the config - accessToken, embedUrl, reportId
    embedConfig.accessToken = reportConfig.accessToken;
    embedConfig.embedUrl = reportConfig.embedUrl;
    embedConfig.reportId = reportConfig.reportId;
}

// Make sure Document object is ready
$(document).ready(function() {

    // Initalize and cache global DOM object
    hiddenSuccess = $("#hidden-success");
    viewName = $("#viewname");
    tickBtn = $("#tick-btn");
    tickIcon = $("#tick-icon");
    copyLinkBtn = $("#copy-link-btn");
    bookmarksList = $("#bookmarks-list");
    copyBtn = $("#copy-btn");
    saveViewBtn = $("#save-view-btn");
    captureViewDiv = $("#capture-view-div");
    saveViewDiv = $("#save-view-div");
    copyLinkText = $("#copy-link-text");
    overlay = $("#overlay");

    // Embed the report in the report-container
    embedBookmarksReport();

    hiddenSuccess.addClass(hiddenClass);
    bookmarksList.hide();

    $("#display-btn").click(function() {
        bookmarksList.toggle("slide");
    });

    $("#close-btn").click(function() {
        bookmarksList.hide("slide");
    });

    copyLinkBtn.click(function() {
        modalButtonClicked(this);
        createLink();
    });

    copyBtn.click(function() {
        copyLink(this);
    });

    saveViewBtn.click(function() {
        modalButtonClicked(this);
    });

    $("#save-bookmark-btn").click(function() {
        onBookmarkCaptureClicked();
    });

    $("#modal-action").on("hidden.bs.modal", function() {

        // Events executed on BootStrap Modal close event
        $(this).find("input").val("").end();
        hiddenSuccess.removeClass(visibleClass);
        hiddenSuccess.addClass(hiddenClass);
        copyLinkBtn.removeClass(activeButtonClass);
        saveViewBtn.addClass(activeButtonClass);
        copyBtn.removeClass(blueBackgroundClass);
        copyBtn.addClass(copyBookmarkClass);
        captureViewDiv.hide();
        tickIcon.hide();
        tickBtn.show();
        viewName.removeClass(alertClass);
        viewName.addClass(validClass);
        saveViewDiv.show();
    });
});

// Embed the report and retrieve the existing report bookmarks
function embedBookmarksReport() {

    // Load sample report properties into session
    return loadSampleReportIntoSession().then(function() {

        // Get models. models contains enums that can be used
        const models = window["powerbi-client"].models;

        // Set the report config from the globals
        setEmbedConfig();

        // Use View permissions
        let permissions = models.Permissions.View;

        // Embed configuration used to describe the what and how to embed
        // This object is used when calling powerbi.embed
        // This also includes settings and options such as filters
        // You can find more information at https://github.com/Microsoft/PowerBI-JavaScript/wiki/Embed-Configuration-Details
        let config = {
            type: "report",
            tokenType: models.TokenType.Embed,
            accessToken: embedConfig.accessToken,
            embedUrl: embedConfig.embedUrl,
            id: embedConfig.reportId,
            permissions: permissions,
            settings: {
                panes: {
                    filters: {
                        expanded: false,
                        visible: true
                    },
                    pageNavigation: {
                        visible: false
                    },
                },
                layoutType: models.LayoutType.Custom,
                customLayout: {
                    displayOption: models.DisplayOption.FitToWidth
                }
            }
        };

        // Get a reference to the embedded report HTML element
        let embedContainer = $("#report-container")[0];

        // Embed the report and display it within the div container
        bookmarkShowcaseState.report = powerbi.embed(embedContainer, config);

        // Report.on will add an event handler for report loaded event.
        bookmarkShowcaseState.report.on("loaded", function() {

            // Get report"s existing bookmarks
            bookmarkShowcaseState.report.bookmarksManager.getBookmarks().then(function(bookmarks) {

                // Create bookmarks list from the existing report bookmarks
                createBookmarksList(bookmarks);
            });

            // Hide the loader
            overlay.hide();

            // Show the container
            $("#main-div").show();
        });
    });
}

// Embed shared report with bookmark on load
function embedSharedBookmark() {

    // Load sample report properties into session
    loadSampleReportIntoSession().then(function() {

        // Get models. models contains enums that can be used
        const models = window["powerbi-client"].models;

        // Set the report config from the globals
        setEmbedConfig();

        // Use View permissions
        let permissions = models.Permissions.View;

        // Get the bookmark name from url param
        let bookmarkName = GetBookmarkNameFromURL();

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
            accessToken: embedConfig.accessToken,
            embedUrl: embedConfig.embedUrl,
            id: embedConfig.reportId,
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

        // Get a reference to the embedded report HTML element
        let embedContainer = $("#bookmark-container")[0];

        // Embed the report and display it within the div container
        bookmarkShowcaseState.report = powerbi.embed(embedContainer, config);

        bookmarkShowcaseState.report.on("loaded", function() {

            // Hide the loader and display the report
            overlay.hide();
            $("#share-bookmark").show();
            bookmarkShowcaseState.report.off("loaded");
        });
    });
}

// Create a bookmarks list from the existing report bookmarks and update the HTML
function createBookmarksList(bookmarks) {

    // Reset next bookmark ID
    bookmarkShowcaseState.nextBookmarkId = 1;

    // Set bookmarks array to the report"s fetched bookmarks
    bookmarkShowcaseState.bookmarks = bookmarks;

    // Build the bookmarks list HTML code
    bookmarkShowcaseState.bookmarks.forEach(function(element) {
        bookmarksList.append(buildBookmarkElement(element));
    });

    // Set first bookmark active
    if (bookmarksList.length) {
        let firstBookmark = $("#" + bookmarkShowcaseState.bookmarks[0].name);

        // Apply first bookmark state
        onBookmarkClicked(firstBookmark[0]);
    }
}

// Build bookmark radio button HTML element
function buildBookmarkElement(bookmark) {
    let labelElement = document.createElement("label");
    labelElement.setAttribute("class", "showcase-radio-container");

    let inputElement = document.createElement("input");
    inputElement.setAttribute("type", "radio");
    inputElement.setAttribute("name", "bookmark");
    inputElement.setAttribute("id", bookmark.name);
    inputElement.setAttribute("onclick", "onBookmarkClicked(this);");
    labelElement.appendChild(inputElement);

    let spanElement = document.createElement("span");
    spanElement.setAttribute("class", "showcase-radio-checkmark");
    labelElement.appendChild(spanElement);

    let secondSpanElement = document.createElement("span");
    secondSpanElement.setAttribute("class", "radio-title");
    let radioTitleElement = document.createTextNode(bookmark.displayName);
    secondSpanElement.appendChild(radioTitleElement);
    labelElement.appendChild(secondSpanElement);

    return labelElement;
}

// Apply clicked bookmark state and set it as the active bookmark on the list
function onBookmarkClicked(element) {

    // Set the clicked bookmark as active
    setBookmarkActive($(element));

    // Apply respective color to the label of the bookmark
    applyColor(element.id);

    // Get bookmark Id from HTML
    const bookmarkId = $(element).attr("id");

    // Find the bookmark in the bookmarks array
    let currentBookmark = getBookmarkByID(bookmarkId);

    // Apply the bookmark state
    bookmarkShowcaseState.report.bookmarksManager.applyState(currentBookmark.state);
}

// Set the bookmark as the active bookmark on the list
function setBookmarkActive(bookmarkSelector) {

    // Set bookmark radio button to checked
    bookmarkSelector.attr("checked", true);
}

// Apply color to the selected checkbox
function applyColor(elementId) {
    let radioSelected = "input[type=radio]";

    // Looping through the radio buttons of the div
    bookmarksList.find(radioSelected).each(function() {
        if (this.id === elementId) {
            $(this.parentNode).removeClass(blackColorClass);
            $(this.parentNode).addClass(blueColorClass);
        } else {
            $(this.parentNode).removeClass(blueColorClass);
            $(this.parentNode).addClass(blackColorClass);
        }
    });
}

// Get the bookmark with bookmarkId name
function getBookmarkByID(bookmarkId) {
    return jQuery.grep(bookmarkShowcaseState.bookmarks, function(bookmark) { return bookmark.name === bookmarkId })[0];
}

// Capture new bookmark of the current state and update the bookmarks list
function onBookmarkCaptureClicked() {

    let capturedViewname = viewName.val();

    // Regex to identify any tags in the name
    let scriptRegex = new RegExp(/<(|\/|[^\/>][^>]+|\/[^>][^>]+)>/g);

    // Regex to identify any special characters except -
    let specialRegex = new RegExp(/[`!@#$%^&*()_+=\[\]{};':\"\\|,.<>\/?~]/);

    // Regex to check to always start with 'Bookmark'
    let bookmarkRegex = new RegExp(/^[B|b]ookmark/);

    if (scriptRegex.test(capturedViewname) || specialRegex.test(capturedViewname)) {
        viewName.removeClass(validClass);
        viewName.addClass(alertClass);
    } else {
        if (bookmarkRegex.test(capturedViewname)) {
            // Capture the report"s current state
            bookmarkShowcaseState.report.bookmarksManager.capture().then(function(capturedBookmark) {

                // Build bookmark element
                let bookmark = {
                    name: "bookmark_" + bookmarkShowcaseState.bookmarkCounter,
                    displayName: capturedViewname,
                    state: capturedBookmark.state
                }

                // Add the new bookmark to the HTML list
                bookmarksList.append(buildBookmarkElement(bookmark));
                bookmarksList.show();

                // Set the captured bookmark as active
                setBookmarkActive($("#bookmark_" + bookmarkShowcaseState.bookmarkCounter));

                // Apply the color when the new bookmark is created
                applyColor("bookmark_" + bookmarkShowcaseState.bookmarkCounter);

                // Add the bookmark to the bookmarks array and increase the bookmarks number counter
                bookmarkShowcaseState.bookmarks.push(bookmark);
                bookmarkShowcaseState.bookmarkCounter++;
                $("#modal-action").modal("hide");
            });
        } else {
            viewName.removeClass(validClass);
            viewName.addClass(alertClass);
        }
    }
}

// Called when the buttons on the Modal gets clicked
function modalButtonClicked(element) {

    // Events executed on BootStrap Modal close event
    $(this).find("input").val("").end();

    saveViewBtn.removeClass(activeButtonClass);
    copyLinkBtn.removeClass(activeButtonClass);

    if (element.id === "save-view-btn") {
        saveViewBtn.addClass(activeButtonClass);
        hiddenSuccess.removeClass(visibleClass);
        hiddenSuccess.addClass(hiddenClass);
        tickIcon.hide();
        tickBtn.show();
        captureViewDiv.hide();
        copyBtn.removeClass(blueBackgroundClass);
        copyBtn.addClass(copyBookmarkClass);
        viewName.removeClass(alertClass);
        viewName.addClass(validClass);
        saveViewDiv.show();
    } else if (element.id === "copy-link-btn") {
        copyLinkBtn.addClass(activeButtonClass);
        saveViewDiv.hide();
        captureViewDiv.show();
    }
}

function createLink() {

    // To get the URL of the parent page 
    let url = (window.location != window.parent.location) ?
        document.referrer :
        document.location.href;

    // Capture the report"s current state
    bookmarkShowcaseState.report.bookmarksManager.capture().then(function(capturedBookmark) {

        // Build bookmark element
        let bookmark = {
            name: "bookmark_" + bookmarkShowcaseState.bookmarkCounter,
            state: capturedBookmark.state
        }

        // Build the share bookmark URL
        let shareUrl = url.substring(0, url.lastIndexOf("/")) + "/shareBookmark.html" + "?id=" + bookmark.name;

        // Store bookmark state with name as a key on the local storage
        // any type of database can be used
        localStorage.setItem(bookmark.name, bookmark.state);

        copyLinkText.val(shareUrl);

        // Increase the bookmarks number counter
        bookmarkShowcaseState.bookmarkCounter++;
    });
}

function copyLink(element) {

    // Set the background color once the copy button is clicked to display SVG image
    $(element).removeClass(copyBookmarkClass);

    // Apply the color
    $(element).addClass(blueBackgroundClass);

    // Hide the Copy text
    tickBtn.hide();

    // Show the tick image
    tickIcon.show();

    // Select the Text Field
    copyLinkText.select();

    // Executing the copy command
    document.execCommand("copy");

    // De-select the text
    if (window.getSelection) { // All browsers, except IE <= 8
        window.getSelection().removeAllRanges();
    }
    hiddenSuccess.removeClass(hiddenClass);
    hiddenSuccess.addClass(visibleClass);
}

// Get the bookmark name from url "id" argument
function GetBookmarkNameFromURL() {
    let url = (window.location != window.parent.location) ?
        document.referrer :
        document.location.href;

    // Using Regex to get the id parameter from the URL
    let regex = new RegExp("[?&]id(=([^&#]*)|&|#|$)"),
        results = regex.exec(url);

    // It can take id parameter from the URL using ? or &
    // If ? or & is not in the URL, returns NULL
    if (!results) { return null };

    // Returns Empty String if id parameter's value is not specified
    if (!results[2]) { return "" };

    // Returns the ID of the Bookmark
    return decodeURIComponent(results[2]);
}