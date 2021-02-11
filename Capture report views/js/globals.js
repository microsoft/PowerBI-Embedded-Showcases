// ----------------------------------------------------------------------------
// Copyright (c) Microsoft Corporation.
// Licensed under the MIT license.
// ----------------------------------------------------------------------------

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

// Cache global DOM objects
const listViewsBtn = $("#display-btn");
const copyLinkSuccessMsg = $("#copy-link-success-msg");
const viewName = $("#viewname");
const tickBtn = $("#tick-btn");
const tickIcon = $("#tick-icon");
const bookmarksList = $("#bookmarks-list");
const copyBtn = $("#copy-btn");
const copyLinkText = $("#copy-link-text");
const copyLinkBtn = $("#copy-link-btn");
const saveViewBtn = $("#save-view-btn");
const captureViewDiv = $("#capture-view-div");
const saveViewDiv = $("#save-view-div");
const overlay = $("#overlay");
const bookmarksDropdown = $(".bookmarks-dropdown");
const captureModal = $("#modal-action");
const closeModal = $("#close-modal-btn");
const viewLinkBtn = $("#copy-btn");
const saveBtn = $("#save-bookmark-btn");
const closeBtn = $("#close-btn");

// Store keycode for TAB key
const KEYCODE_TAB = 9;

// Enum for Keys
const Keys = {
    TAB : "Tab"
}

// Freezing the contents of enum object
Object.freeze(Keys);

// Cache CSS classes
const SELECTED_BUTTON = "selected-button";
const COPY_BOOKMARK = "copy-bookmark";
const ACTIVE_BUTTON = "btn-active";
const VISIBLE = "visible";
const INVISIBLE = "invisible";
const INACTIVE_BOOKMARK = "inactive-bookmark";
const ACTIVE_BOOKMARK = "active-bookmark";
const INVALID_FIELD = "is-invalid";
const FOCUSED = "focused";
const DISPLAY = "show";
const CHECKBOX = "input[type=checkbox]";

// Store IDs of the elements
const SAVE_VIEW_BUTTON_ID = "save-view-btn";
const COPY_LINK_BUTTON_ID =  "copy-link-btn";

// Cache the report containers
const bookmarkContainer = $("#bookmark-container").get(0);
const reportContainer = $("#report-container").get(0);

// Store the state for the checkbox focus
let checkBoxState = null;

// Cache DOM elements to use for trapping the focus inside the modal
const captureModalElements = {
    firstElement: closeModal,
    lastElement: {
        saveView: saveBtn,
        copyLink: viewLinkBtn
    }
}

// Store the last active element 
let lastActiveElement;

// Using Regex to get the id parameter from the URL
const regex = new RegExp("[?&]id(=([^&#]*)|&|#|$)");
