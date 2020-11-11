// To stop the page load on click event
$(document).on("click", ".allow-focus", function(element) {
    element.stopPropagation();
});

// On page resize, visuals should get rearranged according to the page
$(document).ready(function() {

    // Embed the report by calling Endpoint
    embedCustomLayoutReport();

    let layoutsDiv = $("#layouts-div");
    layoutsDiv.hide();

    $("#visuals-layout-btn").click(function() {
        layoutsDiv.toggle();
    });

    $("#visuals-click-btn").click(function() {
        layoutsDiv.hide();
    });

    $(document).click(function() {
        layoutsDiv.hide();
    })

    $("#btn-one-col").click(function() {
        onModifyLayoutClicked(0, 1, this);
    });

    $("#btn-two-col-colspan").click(function() {
        onModifyLayoutClicked(1, 2, this);
    });

    $("#btn-two-col-rowspan").click(function() {
        onModifyLayoutClicked(2, 2, this);
    });

    $("#btn-two-cols").click(function() {
        onModifyLayoutClicked(0, 2, this);
    });

    $("#btn-three-cols").click(function() {
        onModifyLayoutClicked(0, 3, this);
    });

    window.addEventListener("resize", renderVisuals);
});

// Embed the report and retrieve all report visuals
function embedCustomLayoutReport() {

    // Defualt columns value is two columns
    layoutShowcaseState.columns = ColumnsNumber.Two;
    LayoutShowcaseConsts.span = SpanType.None;

    // Load custom layout report properties into session    
    loadLayoutShowcaseReportIntoSession().then(function() {

        // Get models. models contains enums that can be used
        const models = window["powerbi-client"].models;

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

        let reportContainer = $("#report-container").get(0);

        // Embed Power BI report when Access token and Embed URL are available
        layoutShowcaseState.layoutReport = powerbi.load(reportContainer, config);

        // For accessibility insights
        layoutShowcaseState.layoutReport.setComponentTitle('Playground showcase custom layouts report');
        layoutShowcaseState.layoutReport.setComponentTabIndex(0);

        // Clear any other loaded handler events
        layoutShowcaseState.layoutReport.off("loaded");

        // Triggers when a report schema is successfully loaded
        layoutShowcaseState.layoutReport.on("loaded", function() {
            layoutShowcaseState.layoutReport.getPages()
                .then(function(pages) {

                    // Retrieve first page.
                    let activePage = jQuery.grep(pages, function(page) { return page.isActive; })[0];

                    // Set layoutPageName to active page name
                    layoutShowcaseState.layoutPageName = activePage.name;

                    // Get the visuals of the active page
                    activePage.getVisuals()
                        .then(function(visuals) {
                            let reportVisuals = visuals.map(function(visual) {
                                return {
                                    name: visual.name,
                                    title: visual.title,
                                    checked: true
                                };
                            });
                            createVisualsArray(reportVisuals);
                        });
                })

            // Implement phase embedding to first load the report, arrange the visuals and then render
            .then(function() {
                layoutShowcaseState.layoutReport.render();
            });
        });

        // Clear any other rendered handler events
        layoutShowcaseState.layoutReport.off("rendered");

        // Triggers when a report is successfully embedded in UI
        layoutShowcaseState.layoutReport.on("rendered", function() {

            // Phase-embedding
            // Hide the loader
            $("#overlay").hide();
            $('#main-div').children().show();
            console.log("Report render successful");
        });

        // Clear any other error handler events
        layoutShowcaseState.layoutReport.off("error");

        // Handle embed errors
        layoutShowcaseState.layoutReport.on("error", function(event) {
            let errorMsg = event.detail;
            console.error(errorMsg);
        });
    });
}

// Create visuals array from the report visuals and update the HTML
function createVisualsArray(reportVisuals) {

    // Remove all visuals without titles (i.e cards)
    layoutShowcaseState.layoutVisuals = reportVisuals.filter(function(visual) {
        return visual.title !== undefined;
    });

    // Clear visualDropdown div
    let visualsDropdown = $("#visuals-list");
    visualsDropdown.empty();

    // Build checkbox html list and insert the html code to visualDropdown div
    layoutShowcaseState.layoutVisuals.forEach(function(element) {
        visualsDropdown.append(buildVisualElement(element));
    });

    // Render all visuals
    renderVisuals();
}

// Build visual checkbox HTML element
function buildVisualElement(visual) {
    let labelElement = document.createElement("label");
    labelElement.setAttribute("class", "checkbox-container checked");
    labelElement.setAttribute("for", "visual_" + visual.name);

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
    secondSpanElement.setAttribute("class", "checkbox-title");
    let checkboxTitleElement = document.createTextNode(visual.title);
    secondSpanElement.append(checkboxTitleElement);
    labelElement.append(secondSpanElement);

    return labelElement;
}

// Returns true if current browser is Firefox
function isBrowserFirefox() {
    // Refer https://developer.mozilla.org/en-US/docs/Web/HTTP/Browser_detection_using_the_user_agent#Browser_Name
    return navigator.userAgent.includes('Firefox');
}

// Render all visuals with current configuration
function renderVisuals() {

    // render only if report and visuals initialized
    if (!layoutShowcaseState.layoutReport || !layoutShowcaseState.layoutVisuals) {
        return;
    }

    // Get models. models contains enums that can be used
    const models = window["powerbi-client"].models;

    // Get report-container width and height
    let reportContainer = $("#report-container");

    let reportWidth = reportContainer.width();
    let reportHeight = reportContainer.height();

    // Adjust the report width in Firefox to circumvent the horizontal scrollbar issue
    if (isBrowserFirefox()) {
        // Adjust custom layout width for scrollbar
        reportWidth -= 8;
    }

    // Filter the visuals list to display only the checked visuals
    let checkedVisuals = layoutShowcaseState.layoutVisuals.filter(function(visual) { return visual.checked; });

    // Calculating the combined width of the all visuals in a row
    let visualsTotalWidth = reportWidth - (LayoutShowcaseConsts.margin * (layoutShowcaseState.columns + 1));

    // Calculate the width of a single visual, according to the number of columns
    // For one and three columns visuals width will be a third of visuals total width
    let visualWidth = visualsTotalWidth / layoutShowcaseState.columns;

    // Building visualsLayout object
    // You can find more information at https://github.com/Microsoft/PowerBI-JavaScript/wiki/Custom-Layout
    let visualsLayout = {};

    // Visuals starting point
    let x = LayoutShowcaseConsts.margin;
    let y = LayoutShowcaseConsts.margin;

    // Calculate visualHeight with margins
    let visualHeight = visualWidth * LayoutShowcaseConsts.visualAspectRatio;

    // Section means a single unit that will be repeating as pattern to form the layout
    // These 2 variables are used for the 2 custom layouts with spanning
    let rowsPerSection = 2;
    let visualsPerSection = 3;

    // Calculate the number of rows
    let rows = 0;

    if (layoutShowcaseState.span === SpanType.ColSpan) {
        rows = rowsPerSection * Math.floor(checkedVisuals.length / visualsPerSection);
        if (checkedVisuals.length % visualsPerSection) {
            rows += 1;
        }
        reportHeight = Math.max(reportHeight, (rows * visualHeight) + (rows + 1) * LayoutShowcaseConsts.margin);

        checkedVisuals.forEach(function(element, idx) {
            visualsLayout[element.name] = {
                x: x,
                y: y,
                width: (idx % visualsPerSection === visualsPerSection - 1) ? visualWidth * 2 : visualWidth,
                height: visualHeight,
                displayState: {

                    // Change the selected visuals display mode to visible
                    mode: models.VisualContainerDisplayMode.Visible
                }
            };

            // Calculating (x,y) position for the next visual
            x += LayoutShowcaseConsts.margin + ((idx % visualsPerSection === visualsPerSection - 1) ? visualWidth * 2 : visualWidth);

            // Reset x
            if (x + visualWidth > reportWidth) {
                x = LayoutShowcaseConsts.margin;
                y += visualHeight + LayoutShowcaseConsts.margin;
            }
        });

    } else if (layoutShowcaseState.span === SpanType.RowSpan) {
        rows = rowsPerSection * Math.floor(checkedVisuals.length / visualsPerSection);
        if (checkedVisuals.length % visualsPerSection) {
            rows += 2;
        }
        reportHeight = Math.max(reportHeight, (rows * visualHeight) + (rows + 1) * LayoutShowcaseConsts.margin);

        checkedVisuals.forEach(function(element, idx) {
            visualsLayout[element.name] = {
                x: x,
                y: y,
                width: visualWidth,
                height: !(idx % visualsPerSection) ? visualHeight * 2 : visualHeight,
                displayState: {

                    // Change the selected visuals display mode to visible
                    mode: models.VisualContainerDisplayMode.Visible
                }
            };

            // Calculating (x,y) position for the next visual
            x += visualWidth + LayoutShowcaseConsts.margin;

            // Reset x
            if (x + visualWidth > reportWidth) {
                x = ((idx + 1) % visualsPerSection === 0) ? LayoutShowcaseConsts.margin : (2 * LayoutShowcaseConsts.margin) + visualWidth;
                y += (idx % visualsPerSection === 0) ? visualHeight * 2 : visualHeight + LayoutShowcaseConsts.margin;
            }
        });

    } else if (layoutShowcaseState.span === SpanType.None) {
        if (layoutShowcaseState.columns === ColumnsNumber.One) {
            visualHeight /= 2;
        }

        rows = Math.ceil(checkedVisuals.length / layoutShowcaseState.columns);
        reportHeight = Math.max(reportHeight, (rows * visualHeight) + (rows + 1) * LayoutShowcaseConsts.margin);

        checkedVisuals.forEach(function(element) {
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
            x += visualWidth + LayoutShowcaseConsts.margin;

            // Reset x
            if (x + visualWidth > reportWidth) {
                x = LayoutShowcaseConsts.margin;
                y += visualHeight + LayoutShowcaseConsts.margin;
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
                width: reportWidth - 10,
                height: reportHeight - 20
            },
            displayOption: models.DisplayOption.FitToPage,
            pagesLayout: pagesLayout
        }
    };

    // If reportWidth  or reportHeight is changed, change display option to actual size to add scroll bar
    if (reportWidth !== reportContainer.width() || reportHeight !== reportContainer.height()) {
        settings.customLayout.displayOption = models.DisplayOption.ActualSize;
    }

    // Call updateSettings with the new settings object
    layoutShowcaseState.layoutReport.updateSettings(settings);
}

// Update the visuals list with the change and rerender all visuals
function onCheckboxClicked(checkbox) {
    let visual = jQuery.grep(layoutShowcaseState.layoutVisuals, function(visual) { return visual.name === checkbox.value; })[0];
    visual.checked = $(checkbox).is(":checked");
    renderVisuals();
};

// Update columns number and rerender the visuals
function onModifyLayoutClicked(spanType, column, clickedElement) {

    // Selecting the layout option as per the selection
    if (spanType === SpanType.RowSpan) {
        layoutShowcaseState.columns = column;
        layoutShowcaseState.span = SpanType.RowSpan;
    } else if (spanType === SpanType.ColSpan) {
        layoutShowcaseState.columns = column;
        layoutShowcaseState.span = SpanType.ColSpan;
    } else {
        layoutShowcaseState.columns = column;
        layoutShowcaseState.span = SpanType.None;
    }
    setLayoutButtonActive(clickedElement);
    renderVisuals();
}

// Set clicked columns button active
function setLayoutButtonActive(clickedElement) {

    // CSS classes
    const activeBtnClass = "active-columns-btn";
    const layoutButton = "btn-layout";

    // DOM objects
    let activeBtnClassElements = $("." + activeBtnClass);

    // Add the White background to the previous active layout button
    activeBtnClassElements.addClass(layoutButton);

    // Remove the selection from the previous active layout button
    activeBtnClassElements.removeClass(activeBtnClass);

    // Remove the white background to the currently selected button
    $(clickedElement).removeClass(layoutButton);

    // Add the active class to the current selected layout
    $(clickedElement).addClass(activeBtnClass);
}