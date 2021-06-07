// ----------------------------------------------------------------------------
// Copyright (c) Microsoft Corporation.
// Licensed under the MIT license.
// ----------------------------------------------------------------------------

// For the decision of the layout
const COLUMNS = {
    ONE: 1,
    TWO: 2,
    THREE: 3
};

// Freezing the contents of COLUMNS object
Object.freeze(COLUMNS);

// For the decision of two custom layout with spanning
const SPAN_TYPE = {
    NONE: 0,
    ROWSPAN: 1,
    COLSPAN: 2
};

// Freezing the contents of SPAN_TYPE object
Object.freeze(SPAN_TYPE);

// To give consistent margin to each visual in the custom showcase
const LAYOUT_SHOWCASE = {
    MARGIN: 16,
    VISUAL_ASPECT_RATIO: 9 / 16,
};

// Constants used for report configurations as key-value pair
let reportConfig = {
    accessToken: null,
    embedUrl: null,
    reportId: null,
}

// Maintain the state for the showcase
let layoutShowcaseState = {
    columns: COLUMNS.TWO,
    span: SPAN_TYPE.NONE,
    layoutVisuals: null,
    layoutReport: null,
    layoutPageName: null
};

// Get models. models contain enums that can be used
const models = window["powerbi-client"].models;

// Cache DOM elements
const visualsDropdown = $("#visuals-list");
const visualsDiv = $(".dropdown");
const layoutsDiv = $(".layouts");
const layoutsDropdown = $("#layouts-list");
const layoutButtons = $(".btn-util");
const chooseVisualsBtn = $("#choose-visuals-btn");
const chooseLayoutBtn = $("#choose-layouts-btn");

// Store keycode for TAB key
const KEYCODE_TAB = 9;

const Keys = {
    TAB : "Tab"
}

// Freezing the contents of enum object
Object.freeze(Keys);

// Store id for the first visual
let firstVisualId;

// Store id for first button
const firstButtonId = $(".btn-util")[0].id;

// Cache the report containers
const reportContainer = $("#report-container").get(0);