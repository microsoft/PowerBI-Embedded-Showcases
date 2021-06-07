// ----------------------------------------------------------------------------
// Copyright (c) Microsoft Corporation.
// Licensed under the MIT license.
// ----------------------------------------------------------------------------

// To stop the page load on click event
$(document).on("click", ".allow-focus", function (element) {
    element.stopPropagation();
});

// On page resize, visuals should get rearranged according to the page
$(document).ready(function () {

    // Bootstrap the bookmark embed container
    powerbi.bootstrap(reportContainer, { type: "report" });

    // Embed the report by calling Endpoint
    embedCustomLayoutReport();

    $("#btn-one-col").click(function () {
        onModifyLayoutClicked(0, 1, this);
    });

    $("#btn-two-col-colspan").click(function () {
        onModifyLayoutClicked(1, 2, this);
    });

    $("#btn-two-col-rowspan").click(function () {
        onModifyLayoutClicked(2, 2, this);
    });

    $("#btn-two-cols").click(function () {
        onModifyLayoutClicked(0, 2, this);
    });

    $("#btn-three-cols").click(function () {
        onModifyLayoutClicked(0, 3, this);
    });

    // Focus on the first visual selection when the dropdown opens
    visualsDiv.on("shown.bs.dropdown", function () {
        $("input[type=checkbox]")[0].focus();
    });

    layoutsDiv.on("shown.bs.dropdown", function () {

        // Focus on the current selected layout
        const activeLayout = $(".active-columns-btn");
        activeLayout.focus();
    });

    // Move the focus back to the button which triggered the dropdown
    visualsDiv.on("hidden.bs.dropdown", function () {
        chooseVisualsBtn.focus();
    });

    // Move the focus back to the button which triggered the dropdown
    layoutsDiv.on("hidden.bs.dropdown", function () {
        chooseLayoutBtn.focus();
    });

    // Close the layouts dropdown when focus moves from first layout-option to button
    layoutButtons.on("keydown", function (e) {

        // Shift + Tab
        if (e.shiftKey && (e.key === Keys.TAB || e.keyCode === KEYCODE_TAB)) {
            if (document.activeElement.id === firstButtonId) {

                // Close the layouts dropdown
                layoutsDropdown.removeClass("show");
                layoutsDiv.removeClass("show");
                document.getElementById("choose-layouts-btn").setAttribute("aria-expanded", false);
            }
        }
    });

    window.addEventListener("resize", renderVisuals);
});

// Close the visuals dropdown when focus moves from first checkbox to button
$(document).on("keydown", "input:checkbox", function (e) {

    // Shift + Tab
    if (e.shiftKey && (e.key === Keys.TAB || e.keyCode === KEYCODE_TAB)) {
        if (this.id === firstVisualId) {

            // Close the visuals dropdown
            visualsDropdown.removeClass("show");
            visualsDiv.removeClass("show");
            document.getElementById("choose-visuals-btn").setAttribute("aria-expanded", false);
        }
    }
});

// Show tooltip only when ellipsis is active
$(document).on("mouseenter", ".text-truncate", function () {
    const element = $(this);

    if (this.offsetWidth < this.scrollWidth && !element.prop("title")) {
        element.prop("title", element.text());
    }
});

// Embed the report and retrieve all report visuals
async function embedCustomLayoutReport() {

    // Default columns value is two columns
    layoutShowcaseState.columns = COLUMNS.TWO;
    layoutShowcaseState.span = SPAN_TYPE.NONE;

    // Load custom layout report properties into session
    await loadLayoutShowcaseReportIntoSession();

    // Get embed application token from globals
    let accessToken = reportConfig.accessToken;

    // Get embed URL from globals
    let embedUrl = reportConfig.embedUrl;

    // Get report Id from globals
    let embedReportId = reportConfig.reportId;

    // Use View permissions
    let permissions = models.Permissions.View;

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
                    visible: false
                },
                pageNavigation: {
                    visible: false
                }
            }
        }
    };

    // Embed Power BI report when Access token and Embed URL are available
    layoutShowcaseState.layoutReport = powerbi.load(reportContainer, config);

    // For accessibility insights
    layoutShowcaseState.layoutReport.setComponentTitle("Playground showcase custom layouts report");
    layoutShowcaseState.layoutReport.setComponentTabIndex(0);

    // Clear any other loaded handler events
    layoutShowcaseState.layoutReport.off("loaded");

    // Triggers when a report schema is successfully loaded
    layoutShowcaseState.layoutReport.on("loaded", async function () {
        const pages = await layoutShowcaseState.layoutReport.getPages();

        // Retrieve first page.
        let activePage = jQuery.grep(pages, function (page) { return page.isActive; })[0];

        // Set layoutPageName to active page name
        layoutShowcaseState.layoutPageName = activePage.name;

        // Get the visuals of the active page
        const visuals = await activePage.getVisuals();

        let reportVisuals = visuals.map(function (visual) {
            return {
                name: visual.name,
                title: visual.title,
                checked: true
            };
        });

        await createVisualsArray(reportVisuals);

        // Implement phase embedding to first load the report, arrange the visuals and then render
        layoutShowcaseState.layoutReport.render();

        // Phase-embedding
        // Hide the loader
        $("#overlay").hide();
        $("#main-div").children().show();
        console.log("Report render successfully");
    });

    // Clear any other loaded handler events
    layoutShowcaseState.layoutReport.off("rendered");

    // Triggers when a report is successfully embedded in UI
    layoutShowcaseState.layoutReport.on("rendered", function () {
        layoutShowcaseState.layoutReport.off("rendered");
        console.log("The personalize top insights report rendered successfully");

        // Protection against cross-origin failure
        try {
            if (window.parent.playground && window.parent.playground.logShowcaseDoneRendering) {
                window.parent.playground.logShowcaseDoneRendering("PersonalizeTopInsights");
            }
        } catch { }
    });

    // Clear any other error handler events
    layoutShowcaseState.layoutReport.off("error");

    // Handle embed errors
    layoutShowcaseState.layoutReport.on("error", function (event) {
        let errorMsg = event.detail;
        console.error(errorMsg);
    });
}

// Create visuals array from the report visuals and update the HTML
async function createVisualsArray(reportVisuals) {

    // Remove all visuals without titles (i.e cards)
    layoutShowcaseState.layoutVisuals = reportVisuals.filter(function (visual) {
        return visual.title !== undefined;
    });

    // Clear visualDropdown div
    visualsDropdown.empty();

    // Build checkbox html list and insert the html code to visualDropdown div
    layoutShowcaseState.layoutVisuals.forEach(function (element) {
        visualsDropdown.append(buildVisualElement(element));
    });

    // Store the id of the first visual in state
    firstVisualId = $("input:checkbox")[0].id;

    // Render all visuals
    await renderVisuals();
}

// Build visual checkbox HTML element
function buildVisualElement(visual) {
    let labelElement = document.createElement("label");
    labelElement.setAttribute("class", "checkbox-container checked");
    labelElement.setAttribute("for", "visual_" + visual.name);
    labelElement.setAttribute("role", "menuitem");

    let inputElement = document.createElement("input");
    inputElement.setAttribute("type", "checkbox");
    inputElement.setAttribute("id", "visual_" + visual.name);
    inputElement.setAttribute("value", visual.name);
    inputElement.setAttribute("onclick", "onCheckboxClicked(this);");
    inputElement.setAttribute("checked", "true");
    labelElement.append(inputElement);

    let spanElement = document.createElement("span");
    spanElement.setAttribute("class", "checkbox-checkmark");
    labelElement.append(spanElement);

    let secondSpanElement = document.createElement("span");
    secondSpanElement.setAttribute("class", "checkbox-title text-truncate");
    let checkboxTitleElement = document.createTextNode(visual.title);
    secondSpanElement.append(checkboxTitleElement);
    labelElement.append(secondSpanElement);

    return labelElement;
}

// Returns true if current browser is Firefox
function isBrowserFirefox() {
    // Refer https://developer.mozilla.org/en-US/docs/Web/HTTP/Browser_detection_using_the_user_agent#Browser_Name
    return navigator.userAgent.includes("Firefox");
}

// Render all visuals with current configuration
async function renderVisuals() {

    // render only if report and visuals initialized
    if (!layoutShowcaseState.layoutReport || !layoutShowcaseState.layoutVisuals) {
        return;
    }

    // Get report-container width and height
    let reportContainer = $("#report-container");

    let reportWidth = reportContainer.width();
    let reportHeight = 0;

    // Adjust the report width in Firefox to circumvent the horizontal scrollbar issue
    if (isBrowserFirefox()) {
        // Adjust custom layout width for scrollbar
        reportWidth -= 8;
    }

    // Filter the visuals list to display only the checked visuals
    let checkedVisuals = layoutShowcaseState.layoutVisuals.filter(function (visual) { return visual.checked; });

    // Calculating the combined width of the all visuals in a row
    let visualsTotalWidth = reportWidth - (LAYOUT_SHOWCASE.MARGIN * (layoutShowcaseState.columns + 1));

    // Get all the available width for visuals total width, get the space from right margin of the report
    visualsTotalWidth += LAYOUT_SHOWCASE.MARGIN / 2;

    // Calculate the width of a single visual, according to the number of columns
    // For one and three columns visuals width will be a third of visuals total width
    let visualWidth = visualsTotalWidth / layoutShowcaseState.columns;

    // Building visualsLayout object
    // You can find more information at https://github.com/Microsoft/PowerBI-JavaScript/wiki/Custom-Layout
    let visualsLayout = {};

    // Visuals starting point
    let x = LAYOUT_SHOWCASE.MARGIN;
    let y = LAYOUT_SHOWCASE.MARGIN;

    // Calculate visualHeight with margins
    let visualHeight = visualWidth * LAYOUT_SHOWCASE.VISUAL_ASPECT_RATIO;

    // Section means a single unit that will be repeating as pattern to form the layout
    // These 2 variables are used for the 2 custom layouts with spanning
    let rowsPerSection = 2;
    let visualsPerSection = 3;

    // Calculate the number of rows
    let rows = 0;

    if (layoutShowcaseState.span === SPAN_TYPE.COLSPAN) {
        rows = rowsPerSection * Math.floor(checkedVisuals.length / visualsPerSection);
        if (checkedVisuals.length % visualsPerSection) {
            rows += 1;
        }
        reportHeight = Math.max(reportHeight, (rows * visualHeight) + (rows + 1) * LAYOUT_SHOWCASE.MARGIN);

        checkedVisuals.forEach(function (element, idx) {
            visualsLayout[element.name] = {
                x: x,
                y: y,
                width: (idx % visualsPerSection === visualsPerSection - 1) ? visualWidth * 2 + LAYOUT_SHOWCASE.MARGIN : visualWidth,
                height: visualHeight,
                displayState: {

                    // Change the selected visuals display mode to visible
                    mode: models.VisualContainerDisplayMode.Visible
                }
            };

            // Calculating (x,y) position for the next visual
            x += LAYOUT_SHOWCASE.MARGIN + ((idx % visualsPerSection === visualsPerSection - 1) ? visualWidth * 2 : visualWidth);

            // Reset x
            if (x + visualWidth > reportWidth) {
                x = LAYOUT_SHOWCASE.MARGIN;
                y += visualHeight + LAYOUT_SHOWCASE.MARGIN;
            }
        });

    } else if (layoutShowcaseState.span === SPAN_TYPE.ROWSPAN) {
        rows = rowsPerSection * Math.floor(checkedVisuals.length / visualsPerSection);
        if (checkedVisuals.length % visualsPerSection) {
            rows += 2;
        }
        reportHeight = Math.max(reportHeight, (rows * visualHeight) + (rows + 1) * LAYOUT_SHOWCASE.MARGIN);

        checkedVisuals.forEach(function (element, idx) {
            visualsLayout[element.name] = {
                x: x,
                y: y,
                width: visualWidth,
                height: !(idx % visualsPerSection) ? visualHeight * 2 + LAYOUT_SHOWCASE.MARGIN : visualHeight,
                displayState: {

                    // Change the selected visuals display mode to visible
                    mode: models.VisualContainerDisplayMode.Visible
                }
            };

            // Calculating (x,y) position for the next visual
            x += visualWidth + LAYOUT_SHOWCASE.MARGIN;

            // Reset x
            if (x + visualWidth > reportWidth) {
                x = ((idx + 1) % visualsPerSection === 0) ? LAYOUT_SHOWCASE.MARGIN : (2 * LAYOUT_SHOWCASE.MARGIN) + visualWidth;
                y += (idx % visualsPerSection === 0) ? visualHeight * 2 : visualHeight + LAYOUT_SHOWCASE.MARGIN;
            }
        });

    } else if (layoutShowcaseState.span === SPAN_TYPE.NONE) {
        if (layoutShowcaseState.columns === COLUMNS.One) {
            visualHeight /= 2;
        }

        rows = Math.ceil(checkedVisuals.length / layoutShowcaseState.columns);
        reportHeight = Math.max(reportHeight, (rows * visualHeight) + (rows + 1) * LAYOUT_SHOWCASE.MARGIN);

        checkedVisuals.forEach(function (element) {
            visualsLayout[element.name] = {
                x: x,
                y: y,
                width: visualWidth,
                height: visualHeight,
                displayState: {

                    // Change the selected visuals display mode to visible
                    mode: models.VisualContainerDisplayMode.Visible
                }
            };

            // Calculating (x,y) position for the next visual
            x += visualWidth + LAYOUT_SHOWCASE.MARGIN;

            // Reset x
            if (x + visualWidth > reportWidth) {
                x = LAYOUT_SHOWCASE.MARGIN;
                y += visualHeight + LAYOUT_SHOWCASE.MARGIN;
            }
        });
    }

    // Building visualsLayout object
    // You can find more information at https://github.com/Microsoft/PowerBI-JavaScript/wiki/Custom-Layout
    // Building pagesLayout object
    let pagesLayout = {};
    pagesLayout[layoutShowcaseState.layoutPageName] = {
        defaultLayout: {
            displayState: {

                // Default display mode for visuals is hidden
                mode: models.VisualContainerDisplayMode.Hidden
            }
        },
        visualsLayout: visualsLayout
    };

    // Building settings object
    let settings = {

        // Change page background to transparent
        background: models.BackgroundType.Transparent,
        layoutType: models.LayoutType.Custom,
        customLayout: {
            pageSize: {
                type: models.PageSizeType.Custom,
                width: reportWidth,
                height: reportHeight
            },
            displayOption: models.DisplayOption.FitToPage,
            pagesLayout: pagesLayout
        }
    };

    // If reportWidth  or reportHeight is changed, change display option to actual size to add scroll bar
    if (reportWidth !== reportContainer.width() || reportHeight !== reportContainer.height()) {

        // Reset the height of the report-container to avoid the scroll-bar
        resetContainerHeight(reportHeight);

        settings.customLayout.displayOption = models.DisplayOption.ActualSize;
    }

    // Call updateSettings with the new settings object
    await layoutShowcaseState.layoutReport.updateSettings(settings);
}

// Reset the report-container based on the visuals inside it
function resetContainerHeight(newHeight) {
    const reportContainer = $("#report-container");
    reportContainer.height(newHeight);
}

// Update the visuals list with the change and re-render all visuals
function onCheckboxClicked(checkbox) {
    let visual = jQuery.grep(layoutShowcaseState.layoutVisuals, function (visual) { return visual.name === checkbox.value; })[0];
    visual.checked = $(checkbox).is(":checked");
    renderVisuals();
};

// Update columns number and re-render the visuals
function onModifyLayoutClicked(spanType, column, clickedElement) {

    // Selecting the layout option as per the selection
    if (spanType === SPAN_TYPE.ROWSPAN) {
        layoutShowcaseState.columns = column;
        layoutShowcaseState.span = SPAN_TYPE.ROWSPAN;
    } else if (spanType === SPAN_TYPE.COLSPAN) {
        layoutShowcaseState.columns = column;
        layoutShowcaseState.span = SPAN_TYPE.COLSPAN;
    } else {
        layoutShowcaseState.columns = column;
        layoutShowcaseState.span = SPAN_TYPE.NONE;
    }
    setLayoutButtonActive(clickedElement);
    renderVisuals();
}

// Set clicked columns button active
function setLayoutButtonActive(clickedElement) {

    // CSS classes
    const activeBtnClass = "active-columns-btn";
    const layoutButton = "btn-layout";
    const buttons = document.getElementsByClassName("btn-util");

    // DOM objects
    const activeBtnClassElements = $("." + activeBtnClass);

    // Add the White background to the previous active layout button
    activeBtnClassElements.addClass(layoutButton);

    // Remove the selection from the previous active layout button
    activeBtnClassElements.removeClass(activeBtnClass);

    // Remove the white background to the currently selected button
    $(clickedElement).removeClass(layoutButton);

    // Add the active class to the current selected layout
    $(clickedElement).addClass(activeBtnClass);

    // Reset the aria-checked property
    for (btn of buttons) {
        btn.setAttribute("aria-checked", false);
    }

    // Apply the aria-checked property to the selected layout button
    clickedElement.setAttribute("aria-checked", true);
}