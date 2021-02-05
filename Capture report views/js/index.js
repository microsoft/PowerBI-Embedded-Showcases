// ----------------------------------------------------------------------------
// Copyright (c) Microsoft Corporation.
// Licensed under the MIT license.
// ----------------------------------------------------------------------------

// Make sure Document object is ready
$(document).ready(function () {

    // Bootstrap the report-container
    powerbi.bootstrap(reportContainer, reportConfig);

    // Embed the report in the report-container
    embedBookmarksReport();

    closeBtn.click(function () {
        listViewsBtn.focus();
        bookmarksList.removeClass(displayClass);
        bookmarksDropdown.removeClass(displayClass);

        // Set aria-expanded to false when the dropdown is closed by clicking on the Cross button
        const btn = document.getElementById("display-btn");
        btn.setAttribute("aria-expanded", false);

        // Scroll the page to the top
        document.body.scrollTop = document.documentElement.scrollTop = 0;
    });

    // When Pressed Tab on Close button, focus should move to the active bookmark label
    closeBtn.on("keydown", function (e) {
        if (e.key === "Tab" || e.keyCode === KEYCODE_TAB) {
            if (!e.shiftKey) /* Tab */ {
                const activeLabel = document.getElementsByClassName(activeBookmark);
                const activeCheckbox = $(activeLabel).find(checkbox);
                activeCheckbox.focus();
                e.preventDefault();
            }
            else /* Shift + Tab */ {

                // Move focus back to the Button, close the dropdown
                closeBtn.click();
                e.preventDefault();
            }
        }
    });

    copyLinkBtn.click(function () {
        modalButtonClicked(this);
        createLink();
    });

    copyBtn.click(function () {
        copyLink(this);
    });

    saveViewBtn.click(function () {
        modalButtonClicked(this);
    });

    viewName.on("focus", function () {
        viewName.removeClass(invalidField);
    });

    $("#save-bookmark-btn").click(function () {
        onBookmarkCaptureClicked();
    });

    // Stop the form submit event when pressed Enter
    $("form").submit(function () {
        return false;
    });

    closeBtn.on("focus", clearFocus);

    // Move the focus back to the button which triggered the dropdown
    bookmarksDropdown.on("hidden.bs.dropdown", function () {
        listViewsBtn.focus();
        clearFocus();
    });

    // When dropdown is open, focus on the close button
    bookmarksDropdown.on("shown.bs.dropdown", function () {
        closeBtn.focus();
        // Scroll the page to the top
        document.body.scrollTop = document.documentElement.scrollTop = 0;
    });

    // Apply focus on the close button when it is opened
    captureModal.on("shown.bs.modal", function () {
        closeModal.focus();
    });

    // To trap the focus inside the capture view modal while it is open
    captureModal.on("keydown", function (e) {
        let visibleDiv = saveViewDiv.is(":visible");

        if (visibleDiv) {
            lastActiveElement = captureModalElements.lastElement.saveView;
        }
        else {
            lastActiveElement = captureModalElements.lastElement.copyLink;
        }

        if (e.key === "Tab" || e.keyCode === KEYCODE_TAB) {
            if (e.shiftKey) /* shift + tab */ {
                // Compare the activeElement using id
                if ($(document.activeElement)[0].id === captureModalElements.firstElement[0].id) {
                    lastActiveElement.focus();
                    e.preventDefault();
                }
            } else /* tab */ {
                if ($(document.activeElement)[0].id === lastActiveElement[0].id) {
                    captureModalElements.firstElement.focus();
                    e.preventDefault();
                }
            }
        }
    });

    captureModal.on("hidden.bs.modal", function () {

        // Events executed on BootStrap Modal close event
        $(this).find("input").val("").end();
        copyLinkSuccessMsg.removeClass(visible).addClass(invisible);
        copyLinkBtn.removeClass(activeButtonClass);
        saveViewBtn.addClass(activeButtonClass);
        copyBtn.removeClass(selectedButtonClass).addClass(copyBookmarkClass);
        captureViewDiv.hide();
        tickIcon.hide();
        tickBtn.show();
        viewName.removeClass(invalidField);
        saveViewDiv.show();

        // Return Focus to the button which triggered the modal
        $("#capture-btn").focus();
    });
});

// To not to close the dropdown when clicked inside
$(document).on("click", ".allow-focus", function (element) {
    element.stopPropagation();
});

// Focus on the label elemene, whose checkbox has the focus
$(document).on("focus", "input:checkbox", function () {
    clearFocus();
    this.parentElement.classList.add(focused);
});

// Remove the focus from the label if mouse is used
$(document).on("click", "input:checkbox", function () {
    if (!checkBoxState) /* If Mouse is used */ {
        clearFocus();
    }

    checkBoxState = null;
});

// Remove focus from the labels
function clearFocus() {
    const labels = document.getElementsByClassName("showcase-checkbox-container");
    Array.from(labels).forEach(label => {
        label.classList.remove(focused);
    });
}

// Set props for accessibility insights
function setReportAccessibilityProps(report) {
    report.setComponentTitle("Playground showcase sample report");
    report.setComponentTabIndex(0);
}

// Embed the report and retrieve the existing report bookmarks
async function embedBookmarksReport() {

    // Load sample report properties into session
    await loadSampleReportIntoSession();

    // Get models. models contains enums that can be used
    const models = window["powerbi-client"].models;

    // Use View permissions
    let permissions = models.Permissions.View;

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

    // Embed the report and display it within the div container
    bookmarkShowcaseState.report = powerbi.embed(reportContainer, config);

    // For accessibility insights
    setReportAccessibilityProps(bookmarkShowcaseState.report);

    // Report.on will add an event handler for report loaded event.
    bookmarkShowcaseState.report.on("loaded", async function () {

        // Get report's existing bookmarks
        const bookmarks = await bookmarkShowcaseState.report.bookmarksManager.getBookmarks();

        // Create bookmarks list from the existing report bookmarks
        createBookmarksList(bookmarks);

        // Hide the loader
        overlay.addClass(invisible);

        // Show the container
        $("#main-div").addClass(visible);
    });
}

// Create a bookmarks list from the existing report bookmarks and update the HTML
function createBookmarksList(bookmarks) {

    // Reset next bookmark ID
    bookmarkShowcaseState.nextBookmarkId = 1;

    // Set bookmarks array to the report's fetched bookmarks
    bookmarkShowcaseState.bookmarks = bookmarks;

    // Build the bookmarks list HTML code
    bookmarkShowcaseState.bookmarks.forEach(function (element) {
        bookmarksList.append(buildBookmarkElement(element));
    });

    // Set first bookmark active
    if (bookmarksList.length) {
        let firstBookmark = $("#" + bookmarkShowcaseState.bookmarks[0].name);

        // Apply first bookmark state
        onBookmarkClicked(firstBookmark[0]);
    }
}

// Build bookmark checkboxes HTML element
function buildBookmarkElement(bookmark) {
    let labelElement = document.createElement("label");
    labelElement.setAttribute("class", "showcase-checkbox-container");
    labelElement.setAttribute("role", "menuitem");

    let inputElement = document.createElement("input");
    inputElement.setAttribute("type", "checkbox");
    inputElement.setAttribute("name", "bookmark");
    inputElement.setAttribute("id", bookmark.name);
    inputElement.setAttribute("onclick", "onBookmarkClicked(this);");
    labelElement.appendChild(inputElement);

    let spanElement = document.createElement("span");
    spanElement.setAttribute("class", "showcase-checkmark");
    labelElement.appendChild(spanElement);

    let secondSpanElement = document.createElement("span");
    secondSpanElement.setAttribute("class", "checkbox-title");
    let checkboxTitleElement = document.createTextNode(bookmark.displayName);
    secondSpanElement.appendChild(checkboxTitleElement);
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

    $("input:checkbox").prop("checked", false);

    // Set bookmark checkbox to checked
    bookmarkSelector.prop("checked", true);
}

// Activate selected checkbox
function applyColor(elementId) {
    let selectedCheckbox = "input[type=checkbox]";

    // Looping through the checkboxes of the div
    bookmarksList.find(selectedCheckbox).each(function () {
        if (this.id === elementId) {
            $(this.parentNode).removeClass(inactiveBookmark).addClass(activeBookmark);
        } else {
            $(this.parentNode).removeClass(activeBookmark).addClass(inactiveBookmark);
        }
    });
}

// Get the bookmark with bookmarkId name
function getBookmarkByID(bookmarkId) {
    return jQuery.grep(bookmarkShowcaseState.bookmarks, function (bookmark) { return bookmark.name === bookmarkId })[0];
}

// Capture new bookmark of the current state and update the bookmarks list
async function onBookmarkCaptureClicked() {

    let capturedViewname = viewName.val().trim();
    if (!capturedViewname) {
        viewName.addClass(invalidField);
    } else {
        viewName.removeClass(invalidField);

        // Capture the report's current state with personalized visuals
        const capturedBookmark = await bookmarkShowcaseState.report.bookmarksManager.capture({ personalizeVisuals: true });

        // Build bookmark element
        let bookmark = {
            name: "bookmark_" + bookmarkShowcaseState.bookmarkCounter,
            displayName: capturedViewname,
            state: capturedBookmark.state
        }

        // Add the new bookmark to the HTML list
        bookmarksList.append(buildBookmarkElement(bookmark));

        // Open the bookmarks list div and show the applied bookmark
        bookmarksList.addClass("show position");

        bookmarksDropdown.addClass(displayClass);

        // Set aria-expanded to false when the dropdown is closed by clicking on the Cross button
        const btn = document.getElementById("display-btn");
        btn.setAttribute("aria-expanded", true);

        // Set the captured bookmark as active
        const newBookmark = "bookmark_" + bookmarkShowcaseState.bookmarkCounter;
        setBookmarkActive($(newBookmark));

        // Apply the color when the new bookmark is created
        applyColor("bookmark_" + bookmarkShowcaseState.bookmarkCounter);

        // Add the bookmark to the bookmarks array and increase the bookmarks number counter
        bookmarkShowcaseState.bookmarks.push(bookmark);
        bookmarkShowcaseState.bookmarkCounter++;
        captureModal.modal("hide");
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
        copyLinkSuccessMsg.removeClass(visible).addClass(invisible);
        tickIcon.hide();
        tickBtn.show();
        captureViewDiv.hide();
        copyBtn.removeClass(selectedButtonClass).addClass(copyBookmarkClass);
        viewName.removeClass(invalidField);
        saveViewDiv.show();
    } else if (element.id === "copy-link-btn") {
        copyLinkBtn.addClass(activeButtonClass);
        saveViewDiv.hide();
        captureViewDiv.show();
    }
}

async function createLink() {

    // To get the URL of the parent page
    let url = (window.location != window.parent.location) ?
        document.referrer :
        document.location.href;

    // Capture the report's current state with personalized visuals
    const capturedBookmark = await bookmarkShowcaseState.report.bookmarksManager.capture({ personalizeVisuals: true });

    // Build bookmark element
    let bookmark = {
        name: "bookmark_" + bookmarkShowcaseState.bookmarkCounter,
        state: capturedBookmark.state
    }

    // Build the share bookmark URL
    let shareUrl = url.substring(0, url.lastIndexOf("/")) + "/share_bookmark.html" + "?id=" + bookmark.name;

    // Store bookmark state with name as a key on the local storage
    // any type of database can be used
    localStorage.setItem(bookmark.name, bookmark.state);

    copyLinkText.val(shareUrl);

    // Increase the bookmarks number counter
    bookmarkShowcaseState.bookmarkCounter++;
}

function copyLink(element) {

    // Set the background color once the copy button is clicked to display SVG image
    $(element).removeClass(copyBookmarkClass);

    // Apply the color
    $(element).addClass(selectedButtonClass);

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
    
    copyLinkSuccessMsg.removeClass(invisible).addClass(visible);
}
