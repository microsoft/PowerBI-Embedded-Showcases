// ----------------------------------------------------------------------------
// Copyright (c) Microsoft Corporation.
// Licensed under the MIT license.
// ----------------------------------------------------------------------------

// Constants used for report configurations as key-value pair
const reportConfig = {
    accessToken: null,
    embedUrl: null,
    reportId: null
}

// To cache report
const reportShowcaseState = {
    report: null,
    data: null,
    allChecked: false,
    tooltipNextPressed: false
}

// Initialize and cache global DOM objects
const embedContainer = $("#report-container").get(0);
const overlay = $("#overlay");
const distributionDialog = $("#distribution-dialog");
const dialogMask = $("#dialog-mask");
const sendDialog = $("#send-dialog");
const successDialog = $("#success-dialog");
const closeBtn1 = $("#close1");
const closeBtn2 = $("#close2");
const sendCouponBtn = $("#send-coupon");
const sendDiscountBtn = $("#send-discount");
const sendMessageBtn = $("#send-message");
const successCross = $("#success-cross");

// Check if dialog box is closed
let isDialogClosed  = true;

// Key codes
const KEYCODE_TAB = 9;
const KEYCODE_ESCAPE = 27;

// Enum for keys
const Keys = {
    TAB : "Tab",
    ESCAPE: "Escape"
}

// Freezing the contents of enum object
Object.freeze(Keys);

// Table visual GUID
const TABLE_VISUAL_GUID = "1149606f2a101953b4ba";
let tableVisual;

// Icon for the custom extension
const base64Icon = "data:image/png;base64,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"

const distributionDialogElements = {
    firstElement: closeBtn1,
    lastElement: sendDiscountBtn
}

const sendDialogElements = {
    firstElement: closeBtn2,
    lastElement: sendMessageBtn
}

const successDialogElements = {
    firstElement: successCross,
    lastElement: successCross
}