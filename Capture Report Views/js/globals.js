// To cache report config
let reportConfig = {
    accessToken: undefined,
    embedUrl: undefined,
    reportId: undefined,
    type: "report"
}

// To cache bookmark state
let bookmarkShowcaseState = {
    bookmarks: null,
    report: null,

    // Next bookmark ID counter
    bookmarkCounter: 1
};

// Define global DOM objects
let hiddenSuccess = undefined;
let viewName = undefined;
let tickBtn = undefined;
let tickIcon = undefined;
let copyLinkBtn = undefined;
let bookmarksList = undefined;
let copyBtn = undefined;
let saveViewBtn = undefined;
let captureViewDiv = undefined;
let saveViewDiv = undefined;
let copyLinkText = undefined;
let overlay = undefined;

// Cache the report containers
const bookmarkContainer = $("#bookmark-container").get(0);
const reportContainer = $("#report-container").get(0);

// Cache CSS classes
const blueBackgroundClass = "blue-background";
const copyBookmarkClass = "copy-bookmark";
const activeButtonClass = "btn-active";
const visibleClass = "visible";
const hiddenClass = "div-hidden";
const inactiveBookmark = "inactive-bookmark";
const activeBookmark = "active-bookmark";
const invalidField = "is-invalid";

// Using Regex to get the id parameter from the URL
const regex = new RegExp("[?&]id(=([^&#]*)|&|#|$)");
