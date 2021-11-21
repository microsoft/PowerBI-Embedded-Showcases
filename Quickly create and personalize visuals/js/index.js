// ----------------------------------------------------------------------------
// Copyright (c) Microsoft Corporation.
// Licensed under the MIT license.
// ----------------------------------------------------------------------------

// On ready event, bootstrap the embed containers and embed the report
$(document).ready(async function () {

    // Bootstrap the embed-container for report embedding
    powerbi.bootstrap(reportContainer, {
        "type": "report"
    });

    // Bootstrap the visual display container for visual authoring
    powerbi.bootstrap(visualDisplayArea, {
        "type": "report"
    });

    // Hide the authoring area and display the edit screen
    visualAuthoringArea.hide();

    // Load quick visual creator report properties into session
    await loadQuickVisualCreatorReportConfigIntoSession();

    // Embed the report by calling Endpoint
    await embedBaseReport();

    // Embed the report for visual-authoring
    await embedVisualAuthoringReport();

    // Initialize the custom dropdowns
    initializeDropdowns();

    // Select the contents of the visual title when it is focused
    visualTitleText.focus(function () { $(this).select(); });

    // Focus the close button when modal opens
    visualCreatorModal.on("shown.bs.modal", function () {
        closeModalButton.focus();
    });

    // Pressing Tab should move focus to Close button
    createVisualButton.on("keydown", function (event) {

        // Tab
        if (!event.shiftKey && (event.keyCode === KEYCODE_TAB || event.key === Keys.TAB)) {
            closeModalButton.focus();
            event.preventDefault();
        }
    });

    // If Create button is enabled, moves focus to it
    closeModalButton.on("keydown", function (event) {

        // Shift + Tab
        if (event.shiftKey && (event.keyCode === KEYCODE_TAB || event.key === Keys.TAB)) {
            if (!createVisualButton.is(":disabled")) {
                createVisualButton.focus();
                event.preventDefault();
            }
            else {
                event.preventDefault();
            }
        }
    });

    // When the close button is clicked
    closeModalButton.click(function () {

        // Empty the custom visual title variable
        customVisualTitle = "";

        // Empty the state for the edited visual
        selectedVisual.visual = null;

        // Close the modal
        visualCreatorModal.modal("hide");

        // Reset visual generator
        resetVisualGenerator();

        // Clean up the modal
        resetModal();
    });

    // On click of 'Create', append the visual to the base report
    createVisualButton.click(async function () {

        // Hide the modal
        visualCreatorModal.modal("hide");

        // Append the visual to the base report
        appendVisualToReport();

        // Empty the state for the edited visual
        selectedVisual.visual = null;

        // Clean up the modal
        resetModal();
    });

    titleToggle.change(function () {

        if (this.checked) {

            // Reset the visual title if no custom title is set
            if (!customVisualTitle) {
                visualCreatorShowcaseState.newVisual.resetProperty(propertyToSelector("titleText"));
            }
            customTitleWrapper.removeClass(TOGGLE_WRAPPERS_DISABLED);
            visualTitleText.prop("disabled", false);
            hideDisabledEraserAndAligns();
        }
        else {
            customTitleWrapper.addClass(TOGGLE_WRAPPERS_DISABLED);
            visualTitleText.prop("disabled", true);
            showDisabledEraserAndAligns();
        }
    });

    // Close all the open select dropdowns if clicked inside the modal
    visualCreatorModal.click(function () {
        const selectItems = $(".select-items");
        selectItems.each(function () {
            $(this).addClass(HIDE);
        })
    });

    // Focus on the alignment blocks using keyboard
    alignmentBlocks.on("keydown", function (event) {
        if ((event.keyCode === KEYCODE_ENTER || event.key === Keys.ENTER || event.keyCode === KEYCODE_SPACE || event.key === Keys.SPACE)) {

            // Split the alignment from the id and call the function
            const alignmentId = this.id.split("-")[1];
            onAlignClicked(alignmentId);
        }
    })

    // Focus on the eraser tool blocks using keyboard
    enabledEraseTool.on("keydown", function (event) {
        if ((event.keyCode === KEYCODE_ENTER || event.key === Keys.ENTER || event.keyCode === KEYCODE_SPACE || event.key === Keys.SPACE)) {
            onEraseToolClicked();
        }
    });

    // Attach the rearrangeInCustomLayout() to the resize event
    $(window).on("resize", rearrangeInCustomLayout);
});

// Add event listener on document for keyboard
$(document).keydown(function (event) {

    // Close the modal on Escape key
    if (event.keyCode === KEYCODE_ESCAPE || event.key === Keys.ESCAPE) {

        // Hide the modal
        visualCreatorModal.modal("hide");

        // Reset visual generator
        resetVisualGenerator();

        // Clean up the modal
        resetModal();
    }
});

function hideDisabledEraserAndAligns() {
    enabledEraseTool.prop("disabled", false);
    alignmentBlocks.prop("disabled", false);
    disabledEraseTool.hide();
    enabledEraseTool.show();
    enabledAligns.show();
    disabledAligns.hide();
}

// Show the disabled items on modal close
function showDisabledEraserAndAligns() {
    disabledEraseTool.prop("disabled", true);
    alignmentBlocks.prop("disabled", true);
    disabledEraseTool.show();
    enabledEraseTool.hide();
    disabledAligns.show();
    enabledAligns.hide();
}

// Reset the modal and perform clean-up activities
function resetModal() {

    // Hide all the select-box when the modal is closed
    $(".select-items").addClass(HIDE);

    // Show the Edit icon container and hide the authoring container
    editArea.show();
    visualAuthoringArea.hide();

    resetVisualCreatorOptions();
}

function resetVisualCreatorOptions() {

    // Disable the create button
    createVisualButton.prop("disabled", true);

    // Uncheck the visual properties checkboxes
    visualPropertiesCheckboxes.prop("checked", false);

    // Disable the visual properties checkboxes
    visualPropertiesCheckboxes.prop("disabled", true);

    // Disable the visual title textbox
    visualTitleText.prop("disabled", true);

    // Enable all the toggle wrappers
    toggleWrappers.removeClass(TOGGLE_WRAPPERS_DISABLED);

    togglePropertiesSliders.addClass(DISABLED_SLIDERS);

    // Remove disabled class from custom title wrapper
    customTitleWrapper.removeClass(TOGGLE_WRAPPERS_DISABLED);
}

// Set accessibility insights for the report
function setAccessibilityInsights() {
    baseReportState.report.setComponentTitle("Playground showcase quick visual creator");
    baseReportState.report.setComponentTabIndex(0);
}

function setAccessibilityInsightsForAuthoringReport() {
    visualCreatorShowcaseState.report.setComponentTitle("Visual authoring report");
}

// Embed the report and retrieve all report visuals
async function embedBaseReport() {

    // Embed configuration used to describe the what and how to embed
    // This object is used when calling powerbi.embed
    // This also includes settings and options such as filters
    // You can find more information at https://github.com/Microsoft/PowerBI-JavaScript/wiki/Embed-Configuration-Details
    let config = {
        type: "report",
        tokenType: models.TokenType.Embed,
        accessToken: reportConfig.accessToken,
        embedUrl: reportConfig.embedUrl,
        id: reportConfig.reportId,
        // Use the view permissions
        permissions: models.Permissions.View,
        settings: {
            panes: {
                filters: {
                    visible: false
                },
                pageNavigation: {
                    visible: false
                }
            },
            extensions: [
                {
                    command: {
                        name: "changeVisual",
                        title: "Change visual",
                        extend: {
                            visualOptionsMenu: {
                                title: "Change visual",
                                menuLocation: models.MenuLocation.Top,
                            }
                        }
                    }
                }
            ]
        },
        theme: { themeJson: theme }
    };

    // Embed Power BI report when Access token and Embed URL are available
    baseReportState.report = powerbi.load(reportContainer, config);

    // For accessibility insights
    setAccessibilityInsights();

    // Clear any other loaded handler events
    baseReportState.report.off("loaded");

    // Triggers when a report schema is successfully loaded
    baseReportState.report.on("loaded", async function () {
        const pages = await baseReportState.report.getPages();

        // Get the visuals from the first page
        pages[0].setActive();
        baseReportState.page = pages[0];

        // Get the visuals of the active page
        baseReportState.visuals = await baseReportState.page.getVisuals();

        // Rearrange the visuals in 3x3 custom layout
        await rearrangeInCustomLayout();

        // Implement phase embedding to first load the report, arrange the visuals and call the render
        baseReportState.report.render();

        // Implement Phase-embedding
        // Hide the loader
        overlay.hide();
        $(".content").children().show();
    });

    // Clear any other rendered handler events
    baseReportState.report.off("rendered");

    // Triggers when a report is successfully embedded in UI
    baseReportState.report.on("rendered", function () {

        // Update available visual types on UI
        updateAvailableVisualTypes();

        // Enable choosing visual type
        generatorType.removeClass(DISABLED);
        generatorType.removeClass(TYPES_DISABLED);

        console.log("Report render successfully");
    });

    // Listen the commandTriggered event
    baseReportState.report.on("commandTriggered", function (event) {

        // Open the modal and set the fields, properties and title for the visual
        openModalAndFillState(event.detail);
    });

    baseReportState.report.on("buttonClicked", function () {

        // Show the modal to create the visual
        openModalAndFillState();
    });

    // Clear any other error handler events
    baseReportState.report.off("error");

    // Handle embed errors
    baseReportState.report.on("error", function (event) {
        console.error(event.detail);
    });
}

async function embedVisualAuthoringReport() {

    let config = {
        type: "report",
        tokenType: models.TokenType.Embed,
        accessToken: reportConfig.accessToken,
        embedUrl: reportConfig.embedUrl,
        id: reportConfig.reportId,
        // Use the view permissions
        permissions: models.Permissions.View,
        settings: {
            panes: {
                filters: {
                    visible: false
                },
                pageNavigation: {
                    visible: false
                }
            },
            background: models.BackgroundType.Transparent
        }
    };

    // Embed Power BI report when Access token and Embed URL are available
    visualCreatorShowcaseState.report = powerbi.embed(visualDisplayArea, config);

    // For accessibility insights
    setAccessibilityInsightsForAuthoringReport();

    // Clear any other loaded handler events
    visualCreatorShowcaseState.report.off("loaded");

    // Triggers when a report schema is successfully loaded
    visualCreatorShowcaseState.report.on("loaded", async function () {

        // Set the tabindex to -1 to remove the authoring iFrame from keyboard navigation
        authoringiFrame = $(visualDisplayArea).find("iframe");
        authoringiFrame.prop("tabindex", -1);

        const pages = await visualCreatorShowcaseState.report.getPages();

        // pages[1] is an empty page, on which the visual would be created
        pages[1].setActive().catch(error => console.log(error));
        visualCreatorShowcaseState.page = pages[1];
    });

    // Clear any other rendered handler events
    visualCreatorShowcaseState.report.off("rendered");

    // Triggers when a report is successfully embedded in UI
    visualCreatorShowcaseState.report.on("rendered", function () {
        visualCreatorShowcaseState.report.off("rendered");
        console.log("Visual authoring report render successfully");

        // Protection against cross-origin failure
        try {
            if (window.parent.playground && window.parent.playground.logShowcaseDoneRendering) {
                window.parent.playground.logShowcaseDoneRendering('QuickCreate');
            }
        } catch { }
    });

    // Clear any other error handler events
    visualCreatorShowcaseState.report.off("error");

    // Handle embed errors
    visualCreatorShowcaseState.report.on("error", function (event) {
        console.error(event.detail);
    });
}

// Render all visuals with 3x3 custom layout
async function rearrangeInCustomLayout() {

    // render only if report and visuals initialized
    if (!baseReportState.report || !baseReportState.visuals) {
        return;
    }

    // Get report-container width and height
    let reportContainer = $(".report-container");

    let reportWidth = reportContainer.width();
    let reportHeight = 0;

    let visuals = baseReportState.visuals;

    // Calculating the combined width of the all visuals in a row
    let visualsTotalWidth = reportWidth - (VISUAL_CREATOR_SHOWCASE.MARGIN * (VISUAL_CREATOR_SHOWCASE.COLUMNS + 1));

    // Calculate the width of a single visual, according to the number of columns
    // For one and three columns visuals width will be a third of visuals total width
    let visualWidth = visualsTotalWidth / VISUAL_CREATOR_SHOWCASE.COLUMNS;

    // Building visualsLayout object
    // You can find more information at https://github.com/Microsoft/PowerBI-JavaScript/wiki/Custom-Layout
    let visualsLayout = {};

    // Visuals starting point
    let x = VISUAL_CREATOR_SHOWCASE.MARGIN;
    let y = VISUAL_CREATOR_SHOWCASE.MARGIN;

    // Calculate visualHeight with margins
    let visualHeight = visualWidth * VISUAL_CREATOR_SHOWCASE.VISUAL_ASPECT_RATIO;

    // Calculate the number of rows
    let rows = 0;

    // Do not count the overlapping visuals in generating the rows and final report height
    rows = Math.ceil((visuals.length - 2) / VISUAL_CREATOR_SHOWCASE.COLUMNS);
    reportHeight = Math.max(reportHeight, (rows * visualHeight) + (rows + 1) * VISUAL_CREATOR_SHOWCASE.MARGIN);

    visuals.forEach((visual) => {
        // Hide the main and overlapping visuals, if new visual is being created
        if (visualCreationInProgress) {
            // Hide the visual
            if (visual.name === MAIN_VISUAL_GUID || visual.name === imageVisual.name || visual.name === actionButtonVisual.name) {
                visualsLayout[visual.name] = {
                    displayState: {
                        mode: models.VisualContainerDisplayMode.Hidden
                    }
                }
                return;
            }
        }
        if (visual.name === MAIN_VISUAL_GUID) {
            // Store the position of the mainVisual
            mainVisualState = {
                x: x,
                y: y,
                width: visualWidth,
                height: visualHeight,
                displayState: {

                    // Change the selected visuals display mode to visible
                    mode: models.VisualContainerDisplayMode.Visible
                }
            }
        }

        // If the visual is image, which is to be overlapped in the main visual, position it accordingly
        if (visual.name === imageVisual.name && mainVisualState.x) {
            visualsLayout[imageVisual.name] = {
                x: mainVisualState.x + mainVisualState.width * imageVisual.ratio.xPositionRatioWithMainVisual - 6,
                y: mainVisualState.y + mainVisualState.height * imageVisual.ratio.yPositionRatioWithMainVisual,
                // Set minimum width and height for image visual in smaller screens
                width: Math.max(mainVisualState.width * imageVisual.ratio.widthRatioWithMainVisual - 4.5, 36),
                height: Math.max(mainVisualState.height * imageVisual.ratio.heightRatioWithMainVisual - 4.5, 36),
                displayState: {

                    // Change the selected visuals display mode to visible
                    mode: models.VisualContainerDisplayMode.Visible
                }
            };
            imageVisual.yPos = visualsLayout[imageVisual.name].y;
            imageVisual.height = visualsLayout[imageVisual.name].height;
        }

        // If the visual to be placed is the action button, which is to be overlapped in the main visual, position it accordingly
        else if (visual.name === actionButtonVisual.name && mainVisualState.x) {
            visualsLayout[actionButtonVisual.name] = {
                // To center align the action-button
                x: mainVisualState.x + (mainVisualState.width - actionButtonVisual.width) / 2,
                y: imageVisual.height + imageVisual.yPos + DISTANCE,
                // Set the width constant for the button because of non-scalable button value
                width: actionButtonVisual.width,
                // Set the height constant for the button because of non-scalable button value
                height: actionButtonVisual.height,
                displayState: {

                    // Change the selected visuals display mode to visible
                    mode: models.VisualContainerDisplayMode.Visible
                }
            };
        }

        // For remaining visuals, position them and update the x and y coordinates
        else {
            if (visual.name === MAIN_VISUAL_GUID) {
                visualWidth += 4.5;
                visualHeight += 4.5;
            }
            visualsLayout[visual.name] = {
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
            x += visualWidth + VISUAL_CREATOR_SHOWCASE.MARGIN;

            // Reset x
            if (x + visualWidth > reportWidth) {
                x = VISUAL_CREATOR_SHOWCASE.MARGIN;
                y += visualHeight + VISUAL_CREATOR_SHOWCASE.MARGIN;
            }
        }
    });

    // Building visualsLayout object
    // You can find more information at https://github.com/Microsoft/PowerBI-JavaScript/wiki/Custom-Layout
    // Building pagesLayout object
    let pagesLayout = {};
    pagesLayout[baseReportState.page.name] = {
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
            displayOption: models.DisplayOption.FitToWidth,
            pagesLayout: pagesLayout
        },
        // Hide default commands from the context menu for the visuals on the base report
        commands: [
            {
                exportData: {
                    displayOption: models.CommandDisplayOption.Hidden,
                },
                drill: {
                    displayOption: models.CommandDisplayOption.Hidden,
                },
                spotlight: {
                    displayOption: models.CommandDisplayOption.Hidden,
                },
                sort: {
                    displayOption: models.CommandDisplayOption.Hidden,
                },
                seeData: {
                    displayOption: models.CommandDisplayOption.Hidden,
                }
            }
        ],
    };

    // If reportWidth  or reportHeight is changed, change display option to actual size to add scroll bar
    if (reportWidth !== reportContainer.width() || reportHeight !== reportContainer.height()) {

        // Reset the height of the report-container to avoid the scroll-bar
        resetContainerHeight(reportHeight + VISUAL_CREATOR_SHOWCASE.MARGIN);

        settings.customLayout.displayOption = models.DisplayOption.FitToWidth;
    }

    // Call updateSettings with the new settings object
    await baseReportState.report.updateSettings(settings);
}

// Reset the report-container based on the visuals inside it
function resetContainerHeight(newHeight) {
    const reportContainer = $(".report-container");
    reportContainer.height(newHeight);
}

// Initialize the visual-types dropdown and attach a click-listener to the dropdown
function initializeDropdowns() {

    // Look for any elements with the class "styled-select"
    const styledSelects = document.getElementsByClassName("styled-select");
    const length = styledSelects.length;
    for (let i = 0; i < length; i++) {
        const selectedElement = styledSelects[i].getElementsByTagName("select")[0];

        // For each element, create a new div that will act as the selected item
        const dropdownElement = document.createElement("div");
        dropdownElement.setAttribute("class", "select-selected");
        dropdownElement.setAttribute("id", "selected-value-" + i);
        dropdownElement.setAttribute("tabindex", 0);
        dropdownElement.innerHTML = selectedElement.options[selectedElement.selectedIndex].innerHTML;
        styledSelects[i].appendChild(dropdownElement);

        // For each element, create a new div that will contain the option list
        const dropdownItem = document.createElement("div");
        dropdownItem.setAttribute("class", "select-items select-hide");
        dropdownItem.setAttribute("role", "listbox");

        // Create 3 options of data-field for selected visual-type
        const selectedElementLength = selectedElement.length;
        for (let j = 0; j < selectedElementLength; j++) {

            // For each option in the original select element,
            // create a new div that will act as an option item
            const optionItem = document.createElement("div");
            optionItem.setAttribute("role", "option");
            optionItem.setAttribute("tabindex", 0);

            // Create first default option for visual selection dropdown
            if (i === 0 && j === 0) {
                optionItem.innerHTML = VISUAL_TYPE_HEADER;
            }
            else {
                optionItem.innerHTML = selectedElement.options[j].innerHTML;
            }

            // Add new click event listener
            optionItem.addEventListener("click", function () {

                // When an item is clicked, update the original select box, and the selected item
                updateAuthoringVisual(this);
            });

            // Add new keydown event listener
            optionItem.addEventListener("keydown", function (event) {

                // Handle all the keyboard events
                handleKeyEventsForDropdownItems(event, this);
            });

            dropdownItem.appendChild(optionItem);
        }

        styledSelects[i].appendChild(dropdownItem);

        // Add new click event listener for the select box
        dropdownElement.addEventListener("click", function (event) {
            // When the select box is clicked, close any other select boxes,
            // and open/close the current select box
            event.stopPropagation();
            closeAllSelect(this);
            this.nextSibling.classList.toggle(HIDE);
        });

        // Open the dropdowns using Enter OR Space press
        dropdownElement.addEventListener("keydown", function (event) {
            // When the Enter OR Space is pressed on the button, close any other select boxes,
            // and open/close the current select box
            handleKeyEventsForDropdowns(event, this);
        });
    }
}

// Handle key events for dropdown elements
function handleKeyEventsForDropdowns(event, dropdownElement) {
    if (event.keyCode === KEYCODE_ENTER || event.key === Keys.ENTER || event.keyCode === KEYCODE_SPACE || event.key === Keys.SPACE) {
        if (dropdownElement.id === TYPE_DROPDOWN_ID || !generatorFields.hasClass(FIELDS_DISABLED)) {
            event.stopPropagation();
            closeAllSelect(dropdownElement);
            dropdownElement.nextSibling.classList.toggle(HIDE);

            // Focus on the first option when dropdown opens
            if (!dropdownElement.nextSibling.classList.contains(HIDE)) {
                dropdownElement.nextSibling.firstChild.focus();
            }
        }
    }

    // If dropdown is open, the focus should not move using Keyboard to other dropdowns
    if ((event.keyCode === KEYCODE_TAB || event.key === Keys.TAB)) {

        // If data-fields dropdowns is disabled, then move the focus to Close button
        if (dropdownElement.id === TYPE_DROPDOWN_ID && dropdownElement.nextElementSibling.classList.contains(HIDE) && generatorFields.hasClass(FIELDS_DISABLED)) {
            closeModalButton.focus();
            event.preventDefault();
        }

        // Shift + Tab
        if (event.shiftKey) {
            if (!dropdownElement.nextSibling.classList.contains(HIDE)) {
                dropdownElement.focus();
                event.preventDefault();
            }
        }
        // Tab
        else {
            if (dropdownElement.id === "selected-value-3" && createVisualButton.is(":disabled")) {
                closeModalButton.focus();
                event.preventDefault();
            }
        }
    }
}

// All the key-events for the dropdown items
function handleKeyEventsForDropdownItems(event, optionItem) {

    // Focus trap for the dropdowns
    if (event.keyCode === KEYCODE_TAB || event.key === Keys.TAB) {

        // Shift + Tab
        if (event.shiftKey) {
            if (document.activeElement.innerHTML === optionItem.parentNode.firstChild.innerHTML) {
                optionItem.parentElement.previousSibling.focus();
                event.preventDefault();
            }
        }
        // Tab
        else {
            // To handle the case of hidden dropdown items
            // If the focus is on the second last item of the dropdown with the last item as hidden and Tab is pressed, Move focus to the top
            let flag = false;
            let flag2 = false;
            const sibilings = optionItem.parentElement.children;
            const length = sibilings.length;
            if (sibilings[length - 1].style.display === "none") {
                flag = true;
            }

            // If the focus is on the first item of the dropdown and last two items are hidden, Move focus to the top
            for (let i = 0; i < length; i++) {
                if (i < 2) {
                    if (sibilings[i].style.display === "none" && sibilings[i + 1].style.display === "none") {
                        flag2 = true;
                    }
                    else {
                        flag2 = false;
                    }
                }
            }

            if (document.activeElement.innerHTML === (flag2 ? (optionItem.parentNode.firstChild.innerHTML) : (flag) ? optionItem.parentNode.lastChild.previousSibling.innerHTML : optionItem.parentNode.lastChild.innerHTML)) {
                optionItem.parentElement.previousSibling.focus();
                event.preventDefault();
            }

            // Reset
            flag = false;
            flag2 = false;
        }
    }

    // Using the keyboard, update the original select box, and the selected item by pressing Space OR Enter
    if ((event.keyCode === KEYCODE_ENTER || event.key === Keys.ENTER || event.keyCode === KEYCODE_SPACE || event.key === Keys.SPACE)) {
        updateAuthoringVisual(optionItem);
    }
}

// Update the visual based on the selections in the dropdowns
async function updateAuthoringVisual(element) {
    let selects, previousSibling;
    selects = element.parentNode.parentNode.getElementsByTagName("select")[0];
    const selectsLength = selects.length;
    previousSibling = element.parentNode.previousSibling;
    for (let i = 0; i < selectsLength; i++) {
        if (selects.options[i].innerHTML === element.innerHTML) {
            selects.selectedIndex = i;
            previousSibling.innerHTML = element.innerHTML;
            let childElements = element.parentNode.getElementsByClassName(SAME_AS_SELECTED);
            let childElementsLength = childElements.length;
            for (let k = 0; k < childElementsLength; k++) {
                childElements[k].removeAttribute("class");
            }

            // Do not add the class if default option is selected for the data-fields
            if (getFirstWord(element.innerHTML) !== "select") {
                element.setAttribute("class", SAME_AS_SELECTED);
            }

            break;
        }
    }

    // Focus on the div when the dropdown is closed
    document.getElementById(element.parentElement.previousSibling.id).focus();

    previousSibling.click();

    // Change the visual type or update the data role field, according to the dropdown id
    if (selects.id === "visual-type") {
        await changeVisualType(previousSibling.innerHTML);

        // If default option is selected, show the edit area and hide the authoring container
        if (previousSibling.innerHTML === VISUAL_TYPE_HEADER) {
            editArea.show();
            visualAuthoringArea.hide();
        }
        else {
            editArea.hide();
            visualAuthoringArea.show();
        }
    }
    else {
        await updateDataRoleField(selects.parentNode.parentNode.children[0].id, previousSibling.innerHTML);
    }
}

// Close all select boxes in the document, except the current select box
function closeAllSelect(element) {
    const arrNo = [];
    const selected = $(".select-selected");
    const selectItems = $(".select-items");
    for (let i = 0; i < selected.length; i++) {
        if (element === selected[i]) {
            arrNo.push(i);
        }
    }

    for (let i = 0; i < selectItems.length; i++) {
        if (arrNo.indexOf(i)) {
            selectItems[i].classList.add(HIDE);
        }
    }
}

// Get the visual data from it's display name (e.x. Area Chart)
function getVisualFromDisplayName(visualTypeDisplayName) {
    return visualTypeToDataRoles.filter((function (e) { return e.displayName === visualTypeDisplayName }))[0];
}

// Get the visual data from it's name (e.x. areaChart)
function getVisualFromName(name) {
    return visualTypeToDataRoles.filter((function (e) { return e.name === name }))[0];
}

// Change the visual type
async function changeVisualType(visualTypeDisplayName) {

    // If default option is selected, delete the already created visual and reset the showcase state and modal options
    if (visualTypeDisplayName === VISUAL_TYPE_HEADER) {
        resetVisualGenerator();
        visualCreatorShowcaseState.newVisual = null;
        resetModal();
        return;
    }

    // Get the visual-type and data-roles from it's display name
    const visual = getVisualFromDisplayName(visualTypeDisplayName);
    const visualType = visual.name;
    const dataRoles = visual.dataRoles;

    // Do not change OR reset the modal when the same option is selected while changing the visual
    if (visualCreatorShowcaseState.visualType === visualType) {
        return;
    }

    // Remove all data-fields from the state if visual is being edited
    if (selectedVisual.visual) {
        resetGeneratorDataRoles();
    }

    // Retrieve the visual's capabilities
    const capabilities = await baseReportState.report.getVisualCapabilities(visualType);

    // Validate data roles existence on the given visual type
    if (!validateDataRoles(capabilities, dataRoles)) {
        resetVisualGenerator();
        handleInvalidDataRoles();
        return;
    }

    // Enable the data fields section
    generatorFields.removeClass(DISABLED);
    generatorFields.removeClass(FIELDS_DISABLED);

    // Disable the properties section
    generatorProperties.addClass(DISABLED);

    // Disable the toggle sliders
    togglePropertiesSliders.addClass(DISABLED_SLIDERS);

    // Reset all the properties
    resetGeneratorProperties();

    // Disable the Create button
    createVisualButton.prop("disabled", true);

    // Reset the data fields count
    visualCreatorShowcaseState.dataFieldsCount = 0;

    // If the visual doesn't exist, create new visual, otherwise, delete the old visual and create new visual
    if (!visualCreatorShowcaseState.newVisual) {
        await visualCreatorShowcaseState.page.createVisual(visualType, getVisualLayout());
        updateVisualType(visualType, dataRoles);
    }
    else if (visualType !== visualCreatorShowcaseState.visualType) {
        await visualCreatorShowcaseState.page.deleteVisual(visualCreatorShowcaseState.newVisual.name);
        await visualCreatorShowcaseState.page.createVisual(visualType, getVisualLayout());
        updateVisualType(visualType, dataRoles);
    }
}

// Update showcase after visual type change
function updateVisualType(visualTypeName, dataRoles) {

    // Hide the visual headers for the visual inside the modal
    visualCreatorShowcaseState.report.updateSettings(visualHeaderReportSetting);
    updateCurrentVisualState(visualTypeName);
    resetGeneratorDataRoles();
    updateAvailableDataRoles(dataRoles);

    // Update the dropdown options to hide the selected items
    updateDropdownOptions();

    // Uncheck all the properties checkbox
    visualPropertiesCheckboxes.prop("checked", false);

    // Enable all the properties checkbox
    visualPropertiesCheckboxes.prop("disabled", true);

    // Enable the visual title textbox
    visualTitleText.prop("disabled", true);

    // Disable the properties div
    generatorProperties.addClass(PROPERTIES_DISABLED);

    // Show the disabled items on visual type is changed
    showDisabledEraserAndAligns();

    // Focus on the visual type dropdown
    $("#selected-value-0").focus();
}

// Update the visual state
async function updateCurrentVisualState(visualTypeName) {

    const visuals = await visualCreatorShowcaseState.page.getVisuals();

    // Update visual and visual type
    visualCreatorShowcaseState.newVisual = visuals[0];
    visualCreatorShowcaseState.visualType = visualTypeName;

    // Enable the pie chart legend (disabled by default)
    if (visualTypeName === "pieChart") {
        visualCreatorShowcaseState.newVisual.setProperty(propertyToSelector("legend"), { schema: schemas.property, value: true });
    }

    // Format the title to be more accessible
    visualCreatorShowcaseState.newVisual.setProperty(propertyToSelector("titleSize"), { schema: schemas.property, value: 25 });
    visualCreatorShowcaseState.newVisual.setProperty(propertyToSelector("titleColor"), { schema: schemas.property, value: "#000" });

    // Disable unavailable properties for specific visual types
    toggleWrappers.removeClass(TOGGLE_WRAPPERS_DISABLED);

    for (let i = 0; i < showcasePropertiesLength; i++) {
        if (visualTypeProperties[visualTypeName].indexOf(showcaseProperties[i]) < 0) {

            // Uncheck the unavailable properties for the created visual
            $("#" + showcaseProperties[i] + "-toggle").prop("checked", false);

            // Disable the pointer events for the properties
            $("#" + showcaseProperties[i] + ".toggle-wrapper").addClass(TOGGLE_WRAPPERS_DISABLED);
        }
    }
}

// Update the labels for the dropdowns
function updateAvailableDataRoles(dataRoles) {
    const dataRolesNamesElements = document.querySelectorAll(".inline-select-text");

    // Get the select wrappers to change the title to the data-roles
    const selectWrappers = $(".select-selected").slice(1);
    const length = dataRoles.length;
    for (let i = 0; i < length; i++) {
        dataRolesNamesElements[i].innerHTML = dataRoles[i];
        dataRolesNamesElements[i].id = dataRoles[i];

        selectWrappers[i].innerHTML = "Select " + dataRoles[i];
        let dataFields = dataRolesToFields.filter(function (e) { return e.dataRole === dataRoles[i] })[0].Fields;
        updateAvailableDataFields(dataRolesNamesElements[i].parentElement, dataFields);
    }
}

// Update the data fields on the dropdown menus
function updateAvailableDataFields(dataRoleElement, dataFields) {
    const fieldDivElements = dataRoleElement.querySelector(".select-items").children;
    const fieldOptionElements = dataRoleElement.querySelectorAll("option");

    const defaultOptionText = dataRoleElement.firstElementChild.innerHTML;
    fieldDivElements[0].innerHTML = "Select " + defaultOptionText;
    fieldOptionElements[0].innerHTML = "Select " + defaultOptionText;

    const length = dataFields.length;
    for (let i = 0; i < length; i++) {
        fieldDivElements[i + 1].innerHTML = dataFields[i];
        fieldOptionElements[i + 1].innerHTML = dataFields[i];
    }
}

// Update visual types on UI
function updateAvailableVisualTypes() {
    const typesDivElements = $(".select-items")[0].children;
    const typesOptionElements = $("#visual-type")[0].children;

    const visualTypeToDataRolesLength = visualTypeToDataRoles.length;
    for (let i = 0; i < visualTypeToDataRolesLength; i++) {
        typesDivElements[i + 1].innerHTML = visualTypeToDataRoles[i].displayName;
        typesOptionElements[i + 1].innerHTML = visualTypeToDataRoles[i].displayName;
    }
}

// Returns the first word of the text in lowercase
function getFirstWord(text) {
    return text.split(" ")[0].toLowerCase();
}

// If data-role is being reset, then remove it from the visual and return
async function checkForResetDataRole(dataRole, field) {

    // If option with "select" word is selected, remove the data-role from visual and return
    if (getFirstWord(field) === "select") {

        // Get the visual capabilities
        const capabilities = await visualCreatorShowcaseState.newVisual.getCapabilities();

        // Get the data role name
        const dataRoleName = capabilities.dataRoles.filter(function (dr) { return dr.displayName === dataRole })[0].name;

        // Check if the data role already has a field
        if (visualCreatorShowcaseState.dataRoles[dataRoleName]) {

            // Remove the existing data-field from the visual
            await visualCreatorShowcaseState.newVisual.removeDataField(dataRoleName, 0);
            visualCreatorShowcaseState.dataRoles[dataRoleName] = null;
            visualCreatorShowcaseState.dataFieldsCount--;

            // If dataroles count becomes one, then disable the UI
            if (visualCreatorShowcaseState.dataFieldsCount === 1) {
                generatorProperties.addClass(DISABLED);
                generatorProperties.addClass(PROPERTIES_DISABLED);
                resetVisualCreatorOptions();
                showDisabledEraserAndAligns();
            }
        }
        return true;
    }
    return false;
}

// Update data roles field on the visual
async function updateDataRoleField(dataRole, field) {

    // Check if data-role is getting reset for the visual
    const isResetDataRole = await checkForResetDataRole(dataRole, field);
    if (isResetDataRole) {
        return;
    }

    // Check if the requested field is not the same as the selected field
    if (field !== visualCreatorShowcaseState.dataRoles[dataRole]) {

        // Get the visual capabilities
        const capabilities = await visualCreatorShowcaseState.newVisual.getCapabilities();

        // Get the data role name
        const dataRoleName = capabilities.dataRoles.filter(function (dr) { return dr.displayName === dataRole })[0].name;

        // Remove whitespace from field
        const dataFieldKey = field.replace(/\s+/g, "");

        // Check if the data role already has a field
        if (visualCreatorShowcaseState.dataRoles[dataRoleName]) {

            // If the data role has a field, remove it
            await visualCreatorShowcaseState.newVisual.removeDataField(dataRoleName, 0);
            visualCreatorShowcaseState.dataFieldsCount--;

            // If there are no more data fields, recreate the visual before adding the data field
            if (visualCreatorShowcaseState.dataFieldsCount === 0) {
                await visualCreatorShowcaseState.newVisual.addDataField(dataRoleName, dataFieldsTargets[dataFieldKey]);
                visualCreatorShowcaseState.dataRoles[dataRoleName] = dataFieldKey;
                visualCreatorShowcaseState.dataFieldsCount++;

                // Update the dropdown options to hide the selected items
                updateDropdownOptions();
            } else {
                visualCreatorShowcaseState.dataFieldsCount++;
                visualCreatorShowcaseState.dataRoles[dataRoleName] = dataFieldKey;
                visualCreatorShowcaseState.newVisual.addDataField(dataRoleName, dataFieldsTargets[dataFieldKey]);

                // Update the dropdown options to hide the selected items
                updateDropdownOptions();
            }

        } else {

            // Add a new field
            visualCreatorShowcaseState.dataRoles[dataRoleName] = dataFieldKey;
            await visualCreatorShowcaseState.newVisual.addDataField(dataRoleName, dataFieldsTargets[dataFieldKey]);

            // Update the dropdown options to hide the selected items
            updateDropdownOptions();
            visualCreatorShowcaseState.dataFieldsCount++;

            // Show the visual if there are 2 or more data fields
            if (visualCreatorShowcaseState.dataFieldsCount > 1) {
                generatorProperties.removeClass(DISABLED);
                generatorProperties.removeClass(PROPERTIES_DISABLED);
                createVisualButton.prop("disabled", false);

                // If data-roles are three, do not repeat
                if (visualCreatorShowcaseState.dataFieldsCount !== 3) {
                    visualPropertiesCheckboxes.prop("checked", true);
                    visualPropertiesCheckboxes.prop("disabled", false);
                    customTitleWrapper.removeClass(TOGGLE_WRAPPERS_DISABLED);
                    updateAvailableProperties(visualCreatorShowcaseState.visualType);
                    visualTitleText.prop("disabled", false);

                    // Set title property active in visual creation
                    setTitlePropActive();

                    // Show the enabled items to change, align or clear the title
                    hideDisabledEraserAndAligns();
                }
            }
        }
    }
}

function setTitlePropActive() {
    const titleProp = $("#" + "title" + "-toggle");
    const relatedToggle = titleProp.next();
    relatedToggle.removeClass(DISABLED_SLIDERS);
}

function toggleSliders(element, action) {
    const property = $("#" + element + "-toggle");
    const relatedToggle = property.next();
    if (action === "disable") {
        relatedToggle.addClass(DISABLED_SLIDERS);
    }
    else {
        relatedToggle.removeClass(DISABLED_SLIDERS);
    }
}

// Update available properties as per visual type
function updateAvailableProperties(visualType) {
    for (let i = 0; i < showcasePropertiesLength; i++) {
        if (visualTypeProperties[visualType].indexOf(showcaseProperties[i]) < 0) {
            $("#" + showcaseProperties[i] + "-toggle").prop("checked", false);
            $("#" + showcaseProperties[i] + "-toggle").prop("disabled", true);
            toggleSliders(showcaseProperties[i], "disable");
        }
        else {
            $("#" + showcaseProperties[i] + "-toggle").prop("checked", true);
            toggleSliders(showcaseProperties[i], "enable");
        }
    }
}

// Update the dropdown options to hide the selected items
function updateDropdownOptions() {
    $(".select-items div").show();

    // Hide the selected dropdown option from the other data-fields dropdowns
    const selected = $(".select-selected");
    selected.each(function () {
        const selectedValue = $(this).text();
        $(".select-items div:contains(" + selectedValue + ")").hide();

        // Do not hide the selected option in the respective dropdown options list
        const dropdownElements = document.getElementById(this.id).nextElementSibling.children;
        for (element of dropdownElements) {
            if (element.innerHTML === selectedValue) {
                $(element).show();
            }
        }
    });
}

// Return the visual layout
function getVisualLayout() {

    // Width, height, X, Y positions are returned for the new visual to be created in the popup div
    return {
        width: 1240,
        height: 680,
        x: (0.1 * visualAuthoringArea.width()) / 2,
        y: (0.2 * visualAuthoringArea.height()) / 2,
        displayState: {
            // Change the selected visuals display mode to visible
            mode: models.VisualContainerDisplayMode.Visible
        }
    };
}

// Toggle a property value
function toggleProperty(propertyName) {
    if (!visualCreatorShowcaseState.newVisual)
        return;

    const newValue = $("#" + propertyName + "-toggle")[0].checked;

    visualCreatorShowcaseState.properties[propertyName] = newValue;

    // Set the property on the visual
    visualCreatorShowcaseState.newVisual.setProperty(propertyToSelector(propertyName), { schema: schemas.property, value: newValue });
}

// Update the title alignment
function onAlignClicked(direction) {
    if (!visualCreatorShowcaseState.newVisual) {
        return;
    }

    alignmentBlocks.removeClass(SELECTED);
    $("#align-" + direction).addClass(SELECTED);
    visualCreatorShowcaseState.properties["titleAlign"] = direction;

    // Set the property on the visual
    visualCreatorShowcaseState.newVisual.setProperty(propertyToSelector("titleAlign"), { schema: schemas.property, value: direction });
}

// Convert property name to selector
function propertyToSelector(propertyName) {
    switch (propertyName) {
        case "title":
            return { objectName: "title", propertyName: "visible" };
        case "xAxis":
            return { objectName: "categoryAxis", propertyName: "visible" };
        case "yAxis":
            return { objectName: "valueAxis", propertyName: "visible" };
        case "legend":
            return { objectName: "legend", propertyName: "visible" };
        case "titleText":
            return { objectName: "title", propertyName: "titleText" };
        case "titleAlign":
            return { objectName: "title", propertyName: "alignment" };
        case "titleSize":
            return { objectName: "title", propertyName: "textSize" };
        case "titleColor":
            return { objectName: "title", propertyName: "fontColor" };
    }
}

// Handles erase tool click
function onEraseToolClicked() {
    if (!visualCreatorShowcaseState.newVisual)
        return;

    visualTitleText.val("");
    customVisualTitle = "";

    // Reset the title text to auto generated
    visualCreatorShowcaseState.newVisual.resetProperty(propertyToSelector("titleText"));
}

// Update the title's text
function updateTitleText() {
    if (!visualCreatorShowcaseState.newVisual)
        return;

    const visualTitle = visualTitleText.val();
    customVisualTitle = visualTitle;

    // If the title is blank, reset the title to auto generated
    if (visualTitle === "") {
        customVisualTitle = "";
        onEraseToolClicked();
        return;
    }

    visualCreatorShowcaseState.newVisual.setProperty(propertyToSelector("titleText"), { schema: schemas.property, value: visualTitle });
}

// Reset the data roles section
function resetGeneratorDataRoles() {
    if (!visualCreatorShowcaseState.newVisual)
        return;

    visualCreatorShowcaseState.dataRoles = {
        Legend: null,
        Values: null,
        Axis: null,
        Tooltips: null,
    };

    visualCreatorShowcaseState.dataFieldsCount = 0;

    // All dropdowns except of visual type selection
    const nodesToReset = $(".select-selected").slice(1);
    for (let i = 0; i < nodesToReset.length; i++) {
        nodesToReset[i].innerHTML = "Select an option";
    }

    $(".field ~ .select-items").children().show();
    $(".field ~ .select-items").children().removeClass("same-as-selected");
}

// Reset the current visual, call it when the Modal is clicked
function resetGeneratorVisual() {
    if (!visualCreatorShowcaseState.newVisual)
        return;

    visualCreatorShowcaseState.page.deleteVisual(visualCreatorShowcaseState.newVisual.name);
    visualCreatorShowcaseState.newVisual = null;
    visualCreatorShowcaseState.visualType = null;
    $(".select-selected")[0].innerHTML = VISUAL_TYPE_HEADER;

    // Remove sameAsSelected class
    const visualTypeOption = $("#visual-type ~ .select-items > .same-as-selected")[0];
    if (visualTypeOption) {
        visualTypeOption.removeAttribute("class");
    }
}

// Reset the properties section
function resetGeneratorProperties() {
    if (!visualCreatorShowcaseState.newVisual)
        return;

    visualCreatorShowcaseState.properties = {
        legend: true,
        xAxis: true,
        yAxis: true,
        title: true,
        titleText: null,
        titleAlign: null
    };

    alignmentBlocks.removeClass(SELECTED);
    alignLeft.addClass(SELECTED);
    visualTitleText.val("");
}

// Reset the visual generator (data roles, properties and visual)
function resetVisualGenerator() {

    if (!visualCreatorShowcaseState.newVisual)
        return;

    generatorFields.addClass(DISABLED);
    generatorProperties.addClass(DISABLED);
    generatorFields.addClass(FIELDS_DISABLED);
    generatorProperties.addClass(PROPERTIES_DISABLED);

    // Reset data-roles, properties, modal
    resetGeneratorDataRoles();
    resetGeneratorProperties();
    resetGeneratorVisual();
    showDisabledEraserAndAligns();
}

// Validate the existence of each dataRole on the visual's capabilities
function validateDataRoles(capabilities, dataRolesDisplayNames) {
    const length = dataRolesDisplayNames.length;
    for (let i = 0; i < length; i++) {

        // Filter the corrsponding dataRole in the visual's capabilities dataRoles
        if (capabilities.dataRoles.filter(function (dr) { return dr.displayName === dataRolesDisplayNames[i] }).length === 0) {
            return false;
        }
    }
    return true;
}

function handleInvalidDataRoles() {
    // Display error message that particular data-role can not be attached
    console.error("Applied data-roles cannot be assigned to the created visual.");
}

// Create a visual and append that to the base report
// If the visual is selected for editing, this function will edit the visual, otherwise it will create a new visual
async function appendVisualToReport() {
    const newVisual = visualCreatorShowcaseState.newVisual;
    if (!newVisual) {
        resetVisualGenerator();
        return;
    }

    if (!selectedVisual.visual) {

        // Visual creation is started
        visualCreationInProgress = true;

        // Hide the main and overlapped visuals
        rearrangeInCustomLayout();

        // mainVisualState is the position of the current custom visual, where the new visual would be created
        const visualResponse = await baseReportState.page.createVisual(visualCreatorShowcaseState.visualType, mainVisualState);
        const visual = visualResponse.visual;
        const visualType = visual.type;

        // Format the title to be more accessible
        visual.setProperty(propertyToSelector("titleSize"), { schema: schemas.property, value: 13 });
        visual.setProperty(propertyToSelector("titleColor"), { schema: schemas.property, value: "#000" });

        // Enable the legend property for Pie chart
        if (visualCreatorShowcaseState.visualType === "pieChart") {
            visual.setProperty(propertyToSelector("legend"), { schema: schemas.property, value: true });
        }

        // Add properties to the created visual
        Object.entries(visualCreatorShowcaseState.properties).forEach(property => {
            let [propertyName, propertyValue] = property;
            if (propertyName === "titleText") {
                if (customVisualTitle !== "") {

                    // Apply the custom title if available
                    propertyValue = customVisualTitle;
                }
            }

            // Check the validity of the given property for the given visual-type and apply it to the visual
            applyValidPropertiesToTheVisual(visual, visualType, propertyName, propertyValue);
        });

        // Disable the legend for the column and bar charts
        if (visualCreatorShowcaseState.visualType === "columnChart" || visualCreatorShowcaseState.visualType === "barChart") {
            visual.setProperty(propertyToSelector("legend"), { schema: schemas.property, value: false });
        }

        // Remove the data-roles which are null
        Object.keys(visualCreatorShowcaseState.dataRoles).forEach(key => (!visualCreatorShowcaseState.dataRoles[key]) && delete visualCreatorShowcaseState.dataRoles[key]);

        // Add data-fields to the created visual
        Object.entries(visualCreatorShowcaseState.dataRoles).forEach(dataField => {
            const [dataRole, field] = dataField;
            visual.addDataField(dataRole, dataFieldsTargets[field]);
        });

        customVisualTitle = "";

        // Append the created visual to the visuals state of the report
        baseReportState.visuals.push(visual);

        // Visual creation is completed
        visualCreationInProgress = false;

        // Shift the custom visual at last and rearrange all the visuals
        await shiftCustomVisualAtEndAndRearrange();
    }
    else {
        if (visualTitleText.val() !== "") {
            customVisualTitle = visualTitleText.val();
        }

        const oldVisualType = selectedVisual.visual.type;
        const oldVisual = selectedVisual.visual;
        if (oldVisualType !== visualCreatorShowcaseState.visualType) {

            // If visual-type is changed, remove all active data-fields on the visual
            await removeAllActiveDataRoles(oldVisual, oldVisualType);

            // Change the visual type
            await oldVisual.changeType(visualCreatorShowcaseState.visualType);
        }

        // Format the title to be more accessible
        oldVisual.setProperty(propertyToSelector("titleSize"), { schema: schemas.property, value: 13 });
        oldVisual.setProperty(propertyToSelector("titleColor"), { schema: schemas.property, value: "#000" });

        // Enable the legend property for Pie chart
        if (visualCreatorShowcaseState.visualType === "pieChart") {
            oldVisual.setProperty(propertyToSelector("legend"), { schema: schemas.property, value: true });
        }

        // Add properties to the created visual
        Object.entries(visualCreatorShowcaseState.properties).forEach(property => {
            let [propertyName, propertyValue] = property;
            if (propertyName === "titleText" && customVisualTitle !== "") {
                // Apply the custom title
                propertyValue = customVisualTitle;
            }
            if (propertyName === "titleText" && customVisualTitle === "") {

                // Reset the title if custom title is not given
                oldVisual.resetProperty(propertyToSelector("titleText"));
            }
            else {
                // Only apply valid properties to the visual
                applyValidPropertiesToTheVisual(oldVisual, visualCreatorShowcaseState.visualType, propertyName, propertyValue);
            }
        });

        // Disable the legend for the column and bar charts
        if (visualCreatorShowcaseState.visualType === "columnChart" || visualCreatorShowcaseState.visualType === "barChart") {
            oldVisual.setProperty(propertyToSelector("legend"), { schema: schemas.property, value: false });
        }

        // Get related datarole names with the current visual-type
        const dataRoleNames = getVisualFromName(visualCreatorShowcaseState.visualType).dataRoleNames;

        // Add data-fields to the created visual
        Object.entries(visualCreatorShowcaseState.dataRoles).forEach(async function (dataField) {
            const [dataRole, field] = dataField;
            if (dataRoleNames.indexOf(dataRole) < 0) {
                return;
            }

            if (field) {
                // Get data-fields from the data-role
                const dataFieldProp = await oldVisual.getDataFields(dataRole);

                // Check if any data-role is associated with the data-field, If yes then first remove then add new one
                if (dataFieldProp.length === 0) {
                    oldVisual.addDataField(dataRole, dataFieldsTargets[field]);
                }
                else {
                    await oldVisual.removeDataField(dataRole, 0);
                    oldVisual.addDataField(dataRole, dataFieldsTargets[field]);
                }
            }
            else /* If field is null then remove the datarole */ {
                await oldVisual.removeDataField(dataRole, 0);
            }
        });

        customVisualTitle = "";

        // Use the visual from the state to update it's properties
        selectedVisual.visual = null;
    }

    // Reset the dropdowns, authoring-div and modal
    resetVisualGenerator();
}

// Remove all active dataroles from the visual if visual-type is changed
async function removeAllActiveDataRoles(oldVisual, oldVisualType) {
    const dataRoleNames = getVisualFromName(oldVisualType).dataRoleNames;
    dataRoleNames.forEach(async function (dataRole) {
        const dataFieldProp = await oldVisual.getDataFields(dataRole);
        if (dataFieldProp.length === 1) {

            // If data field exists for the data-role, remove it
            await oldVisual.removeDataField(dataRole, 0);
        }
    });
}

// Push the custom visual - (actionButton, shape, image) to last and Rearrange
async function shiftCustomVisualAtEndAndRearrange() {

    // Shift baseShape, Image, actionButton to the end
    const mainVisualIndex = baseReportState.visuals.findIndex(
        (visual) => visual.type === "basicShape"
    );

    if (mainVisualIndex !== -1) {
        baseReportState.visuals.push(baseReportState.visuals.splice(mainVisualIndex, 1)[0]);
    }

    const imageVisualIndex = baseReportState.visuals.findIndex(
        (visual) => visual.type === "image"
    );

    if (imageVisualIndex !== -1) {
        baseReportState.visuals.push(baseReportState.visuals.splice(imageVisualIndex, 1)[0]);
    }

    const actionButtonVisualIndex = baseReportState.visuals.findIndex(
        (visual) => visual.type === "actionButton"
    );

    if (actionButtonVisualIndex !== -1) {
        baseReportState.visuals.push(baseReportState.visuals.splice(actionButtonVisualIndex, 1)[0]);
    }

    // Rearrange the visuals
    await rearrangeInCustomLayout();
}

// This function opens the modal and fill the dropdowns with the data-roles, properties and title of the visual
async function openModalAndFillState(visualData) {

    if (!visualData) {
        // If visualData is not preset, just show the modal
        visualCreatorModal.modal("show");
        return;
    }

    // Fill the state object from the visual response
    await fillStateFromTheVisualData(visualData);
}

// Fill the state object from the visual response
async function fillStateFromTheVisualData(visualData) {

    // Pass the visual to get the IVisual response
    const visualResponse = await getIVisualResponse(visualData.visual);
    selectedVisual.visual = visualResponse;

    const visualType = visualResponse.type;

    // Get the visual data-roles and data-role names from it's name
    const visualResult = getVisualFromName(visualType);
    const dataRoles = visualResult.dataRoles;
    const dataRoleNames = visualResult.dataRoleNames;

    dataRoleNames.forEach(async function (dataRole) {

        // Get data-roles from the visual
        const dataField = await visualResponse.getDataFields(dataRole);

        if (dataField[0] !== undefined) {
            let columnValue = "";

            // Get data-field key
            if (dataField[0].hasOwnProperty("column")) {
                columnValue = dataField[0].column;
            }
            else if (dataField[0].hasOwnProperty("measure")) {
                columnValue = dataField[0].measure;
            }

            // Get Key from Value
            const dataFieldKey = Object.keys(dataFieldsMappings).find(key => {
                return dataFieldsMappings[key] === columnValue
            });

            // Set the data-roles in the state
            visualCreatorShowcaseState.dataRoles[dataRole] = dataFieldKey;
            visualCreatorShowcaseState.dataFieldsCount++;
        }
    });

    // Fetch properties from the visual, which properties need to be reset
    Object.entries(visualCreatorShowcaseState.properties).forEach(async function (visualProperty) {

        // Get the property name
        const propertyName = visualProperty[0];
        if (visualTypeProperties[visualType].indexOf(propertyName) < 0 && titleProperties.indexOf(propertyName) < 0) {
            return;
        }
        const property = await visualResponse.getProperty(propertyToSelector(propertyName));
        if (property.schema === schemas.default) {
            if (propertyName === "xAxis" || propertyName === "yAxis") {
                visualCreatorShowcaseState.properties[propertyName] = true;
            }
            else if (propertyName === "legend") {
                visualCreatorShowcaseState.properties[propertyName] = false;
            }
        }
        else {
            visualCreatorShowcaseState.properties[propertyName] = property.value;
        }
    });

    // Based on the state object, create a visual inside the modal
    await createVisualInsideTheModalInEditMode(visualType, dataRoles);
}

// Check the validity of the given property and apply it to the visual
function applyValidPropertiesToTheVisual(visual, newVisualType, propertyName, propertyValue) {
    if (visualTypeProperties[newVisualType].indexOf(propertyName) < 0 && titleProperties.indexOf(propertyName) < 0) {
        return;
    }

    if ((propertyName === "titleText" || propertyName === "titleAlign") && !propertyValue) {
        visual.resetProperty(propertyToSelector(propertyName));
        return;
    }

    visual.setProperty(propertyToSelector(propertyName), { schema: schemas.property, value: propertyValue })
        .catch(console.error);
}

// Based on the state object, create a visual inside the modal
async function createVisualInsideTheModalInEditMode(visualType, dataRoles) {

    // Create visual inside the modal
    const newVisual = await visualCreatorShowcaseState.page.createVisual(visualType, getVisualLayout());

    // Update state
    visualCreatorShowcaseState.newVisual = newVisual.visual;
    visualCreatorShowcaseState.visualType = newVisual.visual.type;
    const visual = newVisual.visual;
    const newVisualType = visual.type;

    // Format the title to be more accessible
    visual.setProperty(propertyToSelector("titleSize"), { schema: schemas.property, value: 25 });
    visual.setProperty(propertyToSelector("titleColor"), { schema: schemas.property, value: "#000" });

    // Enable the legend property for Pie chart
    if (visualCreatorShowcaseState.visualType === "pieChart") {
        visual.setProperty(propertyToSelector("legend"), { schema: schemas.property, value: true });
    }

    // Add properties to the created visual
    Object.entries(visualCreatorShowcaseState.properties).forEach(property => {
        let [propertyName, propertyValue] = property;
        if (propertyName === "titleText") {

            // If the custom title is given, add that title to the visual
            if (customVisualTitle !== "") {
                propertyValue = customVisualTitle;
            }
        }

        // Check the validity of the given property for the given visual-type and apply it to the visual
        applyValidPropertiesToTheVisual(visual, newVisualType, propertyName, propertyValue);
    });

    // Disable the legend for the column and bar charts
    if (visualCreatorShowcaseState.visualType === "columnChart" || visualCreatorShowcaseState.visualType === "barChart") {
        visual.setProperty(propertyToSelector("legend"), { schema: schemas.property, value: false });
    }

    // Remove the data-roles which are empty from the state
    Object.keys(visualCreatorShowcaseState.dataRoles).forEach(key => (!visualCreatorShowcaseState.dataRoles[key]) && delete visualCreatorShowcaseState.dataRoles[key]);

    // Add data-fields to the created visual
    Object.entries(visualCreatorShowcaseState.dataRoles).forEach(dataField => {
        const [dataRole, field] = dataField;
        visual.addDataField(dataRole, dataFieldsTargets[field]);
    });

    // Update data-roles for the given visual type in the UI
    updateAvailableDataRoles(dataRoles);

    // Set the title property active
    setTitlePropActive();

    visualPropertiesCheckboxes.prop("disabled", false);

    // Populate properties as per state inside the modal
    populateProperties(visualCreatorShowcaseState);

    if (titleToggle.is(":checked")) {
        visualTitleText.prop("disabled", false);
        customTitleWrapper.removeClass(TOGGLE_WRAPPERS_DISABLED);
    }
    else {
        visualTitleText.prop("disabled", true);
        customTitleWrapper.addClass(TOGGLE_WRAPPERS_DISABLED);
    }

    // Remove disabled class from data-roles and properties
    generatorFields.removeClass(FIELDS_DISABLED);
    generatorFields.removeClass(DISABLED);
    generatorProperties.removeClass(PROPERTIES_DISABLED);
    generatorProperties.removeClass(DISABLED);

    // Enable Create visual button
    createVisualButton.prop("disabled", false);

    editArea.hide();
    visualAuthoringArea.show();
    visualCreatorShowcaseState.report.updateSettings(visualHeaderReportSetting);

    // Show the modal
    visualCreatorModal.modal("show");
}

// Populate the visual data inside the modal
function populateProperties(visualCreatorShowcaseState) {

    // Get the visual-type, data-roles and data-role names from it's name
    const visual = getVisualFromName(visualCreatorShowcaseState.visualType);
    const visualDisplayName = visual.displayName;
    const dataRoles = visual.dataRoles;
    const dataRoleNames = visual.dataRoleNames;

    // Set the type of the visual in visual-type dropdown
    $("#selected-value-0").text(visualDisplayName);
    const visualSelectItems = $(".select-items").get(0).children;
    Array.from(visualSelectItems).forEach(visualSelectItem => {
        if (visualSelectItem.innerHTML === visualDisplayName) {
            visualSelectItem.classList.add(SAME_AS_SELECTED);
        }
    });

    // Set the data-roles for the visual
    Object.entries(visualCreatorShowcaseState.dataRoles).forEach(dataField => {
        const [dataRole, field] = dataField;
        const index = dataRoleNames.indexOf(dataRole);
        const dataRoleField = dataFieldsMappings[field];
        if (index !== -1) {
            const value = dataRoles[index];
            selectDataRoles(value, dataRoleField);
        }
    });

    // Update the dropdown options to hide the selected items
    updateDropdownOptions();

    // Set the properties for the visual
    setVisualProperties();
}

// Set the visual properties as per the state object
function setVisualProperties() {
    for (let i = 0; i < showcasePropertiesLength; i++) {
        if (visualTypeProperties[visualCreatorShowcaseState.visualType].indexOf(showcaseProperties[i]) < 0) {

            // Uncheck the inapplicable properties for the created visual
            $("#" + showcaseProperties[i] + "-toggle").prop("checked", false);
            $("#" + showcaseProperties[i] + "-toggle").prop("disabled", true);

            // Disable the pointer events for the properties
            $("#" + showcaseProperties[i] + ".toggle-wrapper").addClass(TOGGLE_WRAPPERS_DISABLED);
            toggleSliders(showcaseProperties[i], "disable");
        }
        else {
            $("#" + showcaseProperties[i] + "-toggle").prop("checked", true);
            toggleSliders(showcaseProperties[i], "enable");
        }
    }

    Object.entries(visualCreatorShowcaseState.properties).forEach(property => {
        let [propertyName, propertyValue] = property;

        if (propertyName === "titleAlign") {
            if (propertyValue === "center" || propertyValue === "right") {
                alignmentBlocks.removeClass(SELECTED);
                $("#align-" + propertyValue).addClass(SELECTED);
            }
        }

        if (propertyName === "titleText") {
            if (typeof propertyValue !== "object" && propertyValue) {
                visualTitleText.val(propertyValue);
            }
        }
        $("#" + propertyName + "-toggle").prop("checked", propertyValue);
    });

    const titleCheck = visualCreatorShowcaseState.properties["title"];
    if (titleCheck) {
        hideDisabledEraserAndAligns();
    }
    else {
        showDisabledEraserAndAligns();
        customVisualTitle = "";
        visualTitleText.val("");
        visualTitleText.prop("disabled", true);
    }

    // Disable the toggle property for Barchart and Columnchart
    if (visualCreatorShowcaseState.visualType === "barChart" || visualCreatorShowcaseState.visualType === "columnChart") {
        legendToggle.prop("checked", false);
        legendToggle.prop("disabled", true);
    }

    // Disable the Axis properties for Pie chart
    if (visualCreatorShowcaseState.visualType === "pieChart") {
        xAxisToggle.prop("checked", false);
        xAxisToggle.prop("disabled", true);
        yAxisToggle.prop("checked", false);
        yAxisToggle.prop("disabled", true);
    }
}

// Populate the data-roles in the modal
function selectDataRoles(dataRoleName, dataRoleValue) {
    const selectSpanWrappers = $(".select-wrapper span");
    const length = selectSpanWrappers.length;

    for (let i = 0; i < length; i++) {
        if (selectSpanWrappers[i].innerHTML === dataRoleName) {
            $("#selected-value-" + (i + 1)).text(dataRoleValue);
            const visualSelectItems = $(".select-items").get(i + 1).children;
            Array.from(visualSelectItems).forEach(visualSelectItem => {
                if (visualSelectItem.innerHTML === dataRoleValue) {
                    visualSelectItem.classList.add(SAME_AS_SELECTED);
                }
            });
        }
    }
}

// Return IVisual response from visual-name
async function getIVisualResponse(visual) {
    const pageVisuals = await baseReportState.page.getVisuals();

    // Get the visual response from the visual of the page by passing the visual name
    const visualResponse = pageVisuals.filter((function (pageVisual) {
        return pageVisual.name === visual.name;
    }))[0];

    return visualResponse;
}