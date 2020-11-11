// Constants used for report configurations as key-value pair
const reportConfig = {
    accessToken: null,
    embedUrl: null,
    reportId: null,
}

// Maintain the state for the showcase
const themesShowcaseState = {
    themesArray: null,
    themesReport: null,
};

// Declare dynamic DOM objects
let dataColorNameElements;
let themeSwitchLabel;
let horizontalSeparator;
let sliderCheckbox;
let allUIElements;

// Cache global DOM elements
const overlay = $("#overlay");
const themesList = $("#theme-dropdown");
const contentElement = $(".content");
const themeContainer = $(".theme-container");
const horizontalRule = $(".horizontal-rule");
const themeButton = $(".btn-theme");
const themeBucket = $(".bucket-theme");
const embedContainer = $(".report-container").get(0);
