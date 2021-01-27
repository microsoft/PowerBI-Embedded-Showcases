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
let listViewsBtn = undefined;
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

// Store keycode for TAB key
const KEYCODE_TAB = 9;

// Cache the DOM Elements
const bookmarksDropdown = $(".bookmarks-dropdown");
const captureModal = $("#modal-action");
const closeModal = $("#close-modal-btn");
const viewLinkBtn = $("#copy-btn");
const saveBtn = $("#save-bookmark-btn");
const closeBtn = $("#close-btn");

// Cache CSS classes
const selectedButtonClass = "selected-button";
const copyBookmarkClass = "copy-bookmark";
const activeButtonClass = "btn-active";
const visible = "visible";
const invisible = "invisible";
const inactiveBookmark = "inactive-bookmark";
const activeBookmark = "active-bookmark";
const invalidField = "is-invalid";

// Cache DOM elements to use for trapping the focus inside the modal
const captureModalElements = {
    firstElement: closeModal,
    lastElement:{
        saveView: saveBtn,
        copyLink: viewLinkBtn
    }
}

// Store the last active element 
let lastActiveElement = undefined;

// Using Regex to get the id parameter from the URL
const regex = new RegExp("[?&]id(=([^&#]*)|&|#|$)");

// First Bookmark id
const firstBookmarkId = "Bookmarkea8f1d8ea6e588f8334a";
