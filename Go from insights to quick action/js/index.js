// ----------------------------------------------------------------------------
// Copyright (c) Microsoft Corporation.
// Licensed under the MIT license.
// ----------------------------------------------------------------------------

// Make sure Document object is ready
$(document).ready(function () {

    // Bootstrap the embed-container for the report embedding
    powerbi.bootstrap(embedContainer, {
        "type": "report"
    });

    // Initially hide the dialog mask and hide all the dialogs boxes
    distributionDialog.hide();
    dialogMask.hide();
    sendDialog.hide();
    successDialog.hide();

    // Embed the report in the report-container
    embedReport();

    closeBtn1.on("click", function () {
        onCloseClicked();
    });

    closeBtn2.on("click", function () {
        onCloseClicked();
    });

    successCross.on("click", function () {
        onCloseClicked();
    });

    sendDiscountBtn.on("click", function () {
        onSendClicked("discount");
    });

    sendCouponBtn.on("click", function () {
        onSendClicked("coupon");
    });

    sendMessageBtn.on("click", function () {
        onSendDialogSendClicked();
        setTimeout(() => {
            if (isDialogClosed === false) {
                onCloseClicked();
            }
        }, 3000);
    });

    // Select the contents of text input when they receive focus
    $(".input-content").focus(function () { $(this).select(); });

    // To trap the focus inside the success dialog and close it on Escape press
    successDialog.on("keydown", event => handleKeyEvents(event, successDialogElements));

    // To trap the focus inside the distribution dialog and close dialog on Escape key press
    distributionDialog.on("keydown", event => handleKeyEvents(event, distributionDialogElements));

    // To trap the focus inside the send dialog and close dialog on Escape key press
    sendDialog.on("keydown", event => handleKeyEvents(event, sendDialogElements));
});

function handleKeyEvents(event, elements) {
    if (event.keyCode === KEYCODE_ESCAPE || event.key === Keys.ESCAPE) {
        onCloseClicked();
        return;
    }
    if (event.key === Keys.TAB || event.keyCode === KEYCODE_TAB) {

        // Shift + Tab
        if (event.shiftKey) {
            // Compare the activeElement using id
            if ($(document.activeElement)[0].id === elements.firstElement[0].id) {
                elements.lastElement.focus();
                event.preventDefault();
            }
        } 
        // Tab
        else {
            if ($(document.activeElement)[0].id === elements.lastElement[0].id) {
                elements.firstElement.focus();
                event.preventDefault();
            }
        }
    }
}

// Set props for accessibility insights
function setReportAccessibilityProps(report) {
    report.setComponentTitle("Insight to Action report");
    report.setComponentTabIndex(0);
}

// Embed the report
async function embedReport() {

    // Load sample report properties into session
    await loadReportIntoSession();

    // Get models. models contains enums that can be used
    const models = window["powerbi-client"].models;

    // Use View permissions
    const permissions = models.Permissions.View;

    // Get embed application token from globals
    const accessToken = reportConfig.accessToken;

    // Get embed URL from globals
    const embedUrl = reportConfig.embedUrl;

    // Get report Id from globals
    const embedReportId = reportConfig.reportId;

    // Embed configuration used to describe the what and how to embed
    // This object is used when calling powerbi.embed
    // This also includes settings and options such as filters
    // You can find more information at https://github.com/Microsoft/PowerBI-JavaScript/wiki/Embed-Configuration-Details
    const config = {
        type: "report",
        tokenType: models.TokenType.Embed,
        accessToken: accessToken,
        embedUrl: embedUrl,
        id: embedReportId,
        permissions: permissions,
        settings: {
            panes: {
                filters: {
                    visible: false
                },
                pageNavigation: {
                    visible: false
                },
            },
            layoutType: models.LayoutType.Custom,
            customLayout: {
                displayOption: models.DisplayOption.FitToWidth
            },

            // Adding the extension command to the options menu
            extensions: [
                {
                    command: {
                        name: "campaign",
                        title: "Start campaign",
                        icon: base64Icon,
                        selector: {
                            $schema: "http://powerbi.com/product/schema#visualSelector",
                            visualName: TABLE_VISUAL_GUID
                        },
                        extend: {
                            visualOptionsMenu: {
                                title: "Start campaign",
                                menuLocation: models.MenuLocation.Top,
                            }
                        }
                    }
                },
            ],
        }
    };

    // Embed the report and display it within the div container
    reportShowcaseState.report = powerbi.embed(embedContainer, config);

    // For accessibility insights
    setReportAccessibilityProps(reportShowcaseState.report);

    // Report.on will add an event handler for report loaded event.
    reportShowcaseState.report.on("loaded", async function () {

        const pages = await reportShowcaseState.report.getPages();

        // Retrieve active page
        const activePage = pages.filter(function (page) {
            return page.isActive
        })[0];

        // Get page's visuals
        const visuals = await activePage.getVisuals();

        // Retrieve the desired visual
        tableVisual = visuals.filter(function (visual) {
            return visual.name === TABLE_VISUAL_GUID;
        })[0];

        // Exports visual data
        tableVisual.exportData(models.ExportDataType.Underlying).then(handleExportData);

        // Hide the loader
        overlay.hide();

        // Show the container
        $("#main-div").show();
    });

    // Adding onClick listener for the button in the report
    reportShowcaseState.report.on("buttonClicked", async function () {

        // Populate data according to the current filters on the table visual
        const result = await tableVisual.exportData(models.ExportDataType.Underlying);
        handleExportData(result);
        onStartCampaignClicked();
    });

    // Adding onClick listener for the custom menu in the table visual in the report
    reportShowcaseState.report.on("commandTriggered", async function (event) {
        if (event.detail.command === "campaign") {

            // Populate data according to the current filters on the table visual
            const result = await tableVisual.exportData(models.ExportDataType.Underlying);
            handleExportData(result);
            onStartCampaignClicked();
        }
    });
}

// Open the send coupon/discount dialog
function onSendClicked(name) {
    const headerText = document.createTextNode("Send " + name + " to distribution list");
    $("#send-dialog .text-dialog-header").empty();
    $("#send-dialog .text-dialog-header").append(headerText);
    $("#send-dialog .title").val("Special offer just for you")

    const promotionToSend = name === "coupon" ? "30$ coupon" : "10% discount";
    const bodyText = "Hi <customer name>, get your " + promotionToSend + " today!";
    $("#send-dialog textarea").val(bodyText);

    distributionDialog.hide();
    successDialog.hide();
    dialogMask.show();
    sendDialog.show();
    closeBtn2.focus();
}

// Handles the export data API result
function handleExportData(result) {

    // Parse the received data from csv to 2d array
    const resultData = parseData(result.data);

    // Filter the unwanted columns
    reportShowcaseState.data = filterTable(["Latest Purchase Category", "Total spend ($)", "Days since last purchase"], resultData);

    // Create a table from the 2d array
    const table = createTable(reportShowcaseState.data)

    // Clear the div
    $("#dialog-table").empty();

    // Add the table to the dialog
    $("#dialog-table").append(table);
}

// Open Campaign list dialog
function onStartCampaignClicked() {
    $(".checkbox-element").prop("checked", true);
    body.addClass(HIDE_OVERFLOW);
    successDialog.hide();
    sendDialog.hide();
    dialogMask.show();
    distributionDialog.show();
    closeBtn1.focus();
}

// Open success dialog
function onSendDialogSendClicked() {
    distributionDialog.hide();
    sendDialog.hide();
    dialogMask.show();
    successDialog.show();
    successCross.focus();
    isDialogClosed = false;
}

// Closes the dialogs
function onCloseClicked() {
    body.removeClass(HIDE_OVERFLOW);
    dialogMask.hide();
    successDialog.hide();
    sendDialog.hide();
    distributionDialog.hide();
    isDialogClosed = true;
}

// Parse the data from the API
function parseData(data) {
    const result = [];
    data.split("\n").forEach(function (row) {
        if (!row) {
            return;
        }
        const rowArray = [];
        row.split(",").forEach(function (cell) {
            rowArray.push(cell);
        });

        result.push(rowArray);
    });
    return result;
}

// Filter the table's data - removing the 'filterValues' columns
function filterTable(filterValues, table) {
    for (let i = 0; i < filterValues.length; i++) {
        valueIndex = table[0].indexOf(
            table[0].filter(function (value) { return value === filterValues[i] })[0]
        );

        for (let j = 0; j < table.length; j++) {
            table[j].splice(valueIndex, 1);
        }
    }
    return table;
}

// Build the HTML table from the data
function createTable(tableData) {
    let table = document.createElement("table");
    let tableBody = document.createElement("tbody");

    // Building table headers, table rows and table columns
    tableData.forEach(function (rowData, rowIndex) {
        let row = document.createElement("tr");
        row.setAttribute("class", "table-row");

        if (rowIndex !== 0) {
            let cell = document.createElement("td");
            cell.setAttribute("class", "cell-checkbox");

            let tableCheckbox = document.createElement("label");
            tableCheckbox.setAttribute("class", "table-checkbox");
            tableCheckbox.setAttribute("aria-label", "Include " + rowData[0]);

            let checkboxElement = document.createElement("input");
            checkboxElement.setAttribute("type", "checkbox");
            checkboxElement.setAttribute("class", "checkbox-element");
            checkboxElement.setAttribute("name", "table-row-checkbox");
            checkboxElement.setAttribute("id", "row" + rowIndex);
            checkboxElement.checked = true;

            let spanElement = document.createElement("span");
            spanElement.setAttribute("class", "checkbox-circle");

            let spanElementChild = document.createElement("span");
            spanElementChild.setAttribute("class", "checkbox-checkmark");
            tableCheckbox.append(checkboxElement);
            tableCheckbox.append(spanElement);
            tableCheckbox.append(spanElementChild);
            cell.append(tableCheckbox);
            row.append(cell);
        }

        rowData.forEach(function (cellData, columnIndex) {
            let cell;
            if (rowIndex !== 0) {
                cell = document.createElement("td");
                if (columnIndex === 0) {
                    cell.setAttribute("class", "name-cell");
                } else if (columnIndex === 1) {
                    cell.setAttribute("class", "region-cell");
                } else if (columnIndex === 2) {
                    cell.setAttribute("class", "mail-cell");
                } else if (columnIndex === 3) {
                    cell.setAttribute("class", "phone-cell");
                }
            } else {
                cell = document.createElement("th");
                cell.setAttribute("class", "table-headers");
                if (columnIndex === 0) {
                    cell.setAttribute("id", "name");
                } else if (columnIndex === 1) {
                    cell.setAttribute("id", "region");
                } else if (columnIndex === 2) {
                    cell.setAttribute("id", "mail");
                } else if (columnIndex === 3) {
                    cell.setAttribute("id", "phone");
                }
            }
            cell.append(document.createTextNode(cellData));
            row.append(cell);
        });
        tableBody.append(row);
    });

    table.append(tableBody);
    return table;
}