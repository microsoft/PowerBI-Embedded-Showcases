// ----------------------------------------------------------------------------
// Copyright (c) Microsoft Corporation.
// Licensed under the MIT license.
// ----------------------------------------------------------------------------

// Perform events only after DOM is loaded
$(document).ready(function () {

    // Bootstrap the embed-container for the report embedding
    powerbi.bootstrap(embedContainer, {
        "type": "report"
    });

    // Embed the report in the report-container
    embedThemesReport();

    // Build Theme palette
    buildThemePalette();

    // Cache dynamic elements to toggle the theme
    themeSlider = $("#theme-slider");
    dataColorNameElements = $(".data-color-name");
    themeSwitchLabel = $(".theme-switch-label");
    horizontalSeparator = $(".dropdown-separator");
    sliderCheckbox = $(".slider");

    // Move the focus back to the button which triggered the dropdown
    dropdownDiv.on("hidden.bs.dropdown", function () {
        themeButton.focus();
    });

    // Focus on the theme slider once the dropdown opens
    dropdownDiv.on("shown.bs.dropdown", function () {
        themeSlider.focus();
    });

    // Get all the UI elements to toggle the dark theme
    allUIElements = [bodyElement, contentElement, themeContainer, themeSwitchLabel, horizontalSeparator, horizontalRule, sliderCheckbox, themeButton, themeBucket, dataColorNameElements];
});

// Close the dropdown when focus moves to Choose theme button from Toggle slider
$(document).on("keydown", "#theme-slider", function(e) {
    if (e.shiftKey && (e.key === "Tab" || e.keyCode === KEYCODE_TAB)) {
        dropdownDiv.removeClass("show");
        themesList.removeClass("show");
        $(".btn-theme")[0].setAttribute("aria-expanded", false);
    }
});

// Set properties for Accessibility insights
function setReportAccessibilityProps(report) {
    report.setComponentTitle("Playground showcase sample Theme report");
    report.setComponentTabIndex(0);
}

// To stop the page load on click event inside dropdown
$(document).on("click", ".allow-focus", function (element) {
    element.stopPropagation();
});

// Embed the report
async function embedThemesReport() {

    // Load sample report properties into session
    await loadThemesShowcaseReportIntoSession();

    // Get models. models contains enums that can be used
    const models = window["powerbi-client"].models;

    // Get embed application token from globals
    const accessToken = reportConfig.accessToken;

    // Get embed URL from globals
    const embedUrl = reportConfig.embedUrl;

    // Get report Id from globals
    const embedReportId = reportConfig.reportId;

    // Use View permissions
    const permissions = models.Permissions.View;

    // Embed configuration used to describe the what and how to embed
    // This object is used when calling powerbi.embed
    // This also includes settings and options such as filters
    // You can find more information at https://github.com/Microsoft/PowerBI-JavaScript/wiki/Embed-Configuration-Details
    let config = {
        type: "report",
        tokenType: models.TokenType.Embed,
        accessToken: accessToken,
        embedUrl: embedUrl,
        id: embedReportId,
        permissions: permissions,
        settings: {
            panes: {
                filters: {
                    expanded: false,
                    visible: false
                },
                pageNavigation: {
                    visible: false
                },
            },
            layoutType: models.LayoutType.Custom,
            customLayout: {
                displayOption: models.DisplayOption.FitToPage
            },
            background: models.BackgroundType.Transparent
        },
        // Adding theme attribute to the config, will apply the light theme and default data-colors on load
        theme: { themeJson: jsonDataColors[0] },
    };

    // Embed the report and display it within the div container
    themesShowcaseState.themesReport = powerbi.embed(embedContainer, config);

    // For accessibility insights
    setReportAccessibilityProps(themesShowcaseState.themesReport);

    // Report.on will add an event handler for report loaded event.
    themesShowcaseState.themesReport.on("loaded", function () {

        // Hide the loader
        overlay.hide();

        // Show the container
        $(".content").show();

        // Set the first data-color on the list as active
        themesList.find("#datacolor0").prop("checked", true);
    });
}

// Build the theme palette
function buildThemePalette() {

    // Create Theme switcher
    buildThemeSwitcher();

    // Create separator
    buildSeparator();

    // Building the data-colors list
    for (let i = 0; i < jsonDataColors.length; i++) {
        themesList.append(buildDataColorElement(i));
    }
}

// Build the theme switcher with the theme slider
function buildThemeSwitcher() {

    // Div wrapper element
    let divElement = document.createElement("div");
    divElement.setAttribute("class", "theme-element-container");
    divElement.setAttribute("role", "menuitem");

    let spanElement = document.createElement("span");
    spanElement.setAttribute("class", "theme-switch-label");
    spanElement.setAttribute("id", "dark-label-text");
    let textNodeElement = document.createTextNode("Dark mode");
    spanElement.appendChild(textNodeElement);
    divElement.appendChild(spanElement);

    // Build the checkbox slider
    let checkboxElement = document.createElement("label");
    checkboxElement.setAttribute("class", "switch");
    checkboxElement.setAttribute("aria-labelledby", "dark-label-text");

    let inputCheckboxElement = document.createElement("input");
    inputCheckboxElement.setAttribute("id", "theme-slider");
    inputCheckboxElement.setAttribute("type", "checkbox");
    inputCheckboxElement.setAttribute("onchange", "toggleTheme()");

    let sliderElement = document.createElement("span");
    sliderElement.setAttribute("class", "slider round");

    checkboxElement.appendChild(inputCheckboxElement);
    checkboxElement.appendChild(sliderElement);

    // Attach the checox slider to text label
    divElement.appendChild(checkboxElement);

    // Append the element to the dropdown
    themesList.append(divElement);
}

// Build horizontal separator between the theme switcher and data color elements
function buildSeparator() {

    // Build the separator between theme-switcher and data-colors
    let separatorElement = document.createElement("div");
    separatorElement.setAttribute("class", "dropdown-separator");
    separatorElement.setAttribute("role", "separator");
    themesList.append(separatorElement);
}

// Build data-colors list based on the JSON object
function buildDataColorElement(id) {

    // Div wrapper element for the data-color
    let divElement = document.createElement("div");
    divElement.setAttribute("class", "theme-element-container");
    divElement.setAttribute("role", "group");

    let inputElement = document.createElement("input");
    inputElement.setAttribute("role", "menuitemradio");
    inputElement.setAttribute("type", "radio");
    inputElement.setAttribute("name", "data-color");
    inputElement.setAttribute("id", "datacolor" + id);
    inputElement.setAttribute("aria-label", jsonDataColors[id].name + " color theme");
    inputElement.setAttribute("onclick", "onDataColorWrapperClicked(this);");
    divElement.appendChild(inputElement);

    // Text-element based on the JSON object
    let secondSpanElement = document.createElement("span");
    secondSpanElement.setAttribute("class", "data-color-name");
    secondSpanElement.setAttribute("onclick", "onDataColorWrapperClicked(this)");

    let radioTitleElement = document.createTextNode(jsonDataColors[id].name);
    secondSpanElement.appendChild(radioTitleElement);
    divElement.appendChild(secondSpanElement);

    // Div for displaying data-colors based on the JSON object
    let colorsDivElement = document.createElement("div");
    colorsDivElement.setAttribute("class", "theme-colors");
    colorsDivElement.setAttribute("onclick", "onDataColorWrapperClicked(this)");

    const dataColors = jsonDataColors[id].dataColors;
    for (let i = 0; i < dataColors.length; i++) {
        let dataColorElement = document.createElement("div");
        dataColorElement.setAttribute("class", "data-color");
        dataColorElement.setAttribute("style", "background:#" + dataColors[i].substr(1));
        colorsDivElement.appendChild(dataColorElement);
    }

    // Add the colors div to the label element
    divElement.appendChild(colorsDivElement);

    return divElement;
}

// Apply the selected data-color to the report from the wrapper element
function onDataColorWrapperClicked(element) {

    // Deselect the previously selected data-color
    $("input[name=data-color]:checked", "#theme-dropdown").prop("checked", false);

    // Set the respective data-color as active from the wrapper element
    $(element.parentElement.firstElementChild).prop("checked", true);

    // Apply the theme to the report based on the data-color and the background
    applyTheme();
}

// Apply the theme based on the mode and the data-color selected
async function applyTheme() {

    // Get active data-color id
    activeDataColorId = Number($("input[name=data-color]:checked", "#theme-dropdown")[0].getAttribute("id").slice(-1));

    // Get the theme state from the slider toggle
    let activeThemeSlider = $("#theme-slider").get(0);

    // Get the index of the theme based on the state of the slider
    // 1 => Dark theme
    // 0 => Light theme
    const themeId = (activeThemeSlider.checked) ? 1 : 0;

    // Create new theme object
    let newTheme = {};

    // Append the data-colors and the theme
    $.extend(newTheme, jsonDataColors[activeDataColorId], themes[themeId]);

    // Apply the theme to the report
    await themesShowcaseState.themesReport.applyTheme({ themeJson: newTheme });
}

// Apply theme to the report and toggle dark theme for the UI elements
async function toggleTheme() {

    // Apply the theme in the report
    await applyTheme();

    // Toggle the dark theme for all UI elements
    toggleDarkThemeOnElements();
}

// Toggle dark theme for the UI elements
function toggleDarkThemeOnElements() {

    // Toggle theme for all the UI elements
    allUIElements.forEach(element => {
        element.toggleClass("dark");
    });
}
