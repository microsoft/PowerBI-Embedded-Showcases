// To stop the page load on click event
$(document).on("click", ".allow-focus", function (element) {
    element.stopPropagation();
});

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

    $(".slider").addClass(disabledSliders);

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

        // If title toggle is checked then show enabled erase-tool
        if (this.checked) {
            customTitleWrapper.removeClass(toggleWrappersDisabledClass);
            disabledEraseTool.hide();
            enabledEraseTool.show();
        }
        else {
            customTitleWrapper.addClass(toggleWrappersDisabledClass);
            disabledEraseTool.show();
            enabledEraseTool.hide();
        }
    });

    // Close all the open select dropdowns if clicked inside the modal
    visualCreatorModal.click(function () {
        const selectItems = $(".select-items");

        selectItems.each(function () {
            $(this).addClass(selectHideClass);
        })
    });

    // Disable the Create button on first load
    createVisualButton.prop("disabled", true);

    // Attach the rearrangeInCustomLayout() to the resize event
    $(window).on("resize", rearrangeInCustomLayout);
});

// Add event listener on document for key-board
$(document).keydown(function (event) {

    // Close the modal on Escape key
    if (event.keyCode == 27) {

        // Hide the modal
        visualCreatorModal.modal("hide");

        // Reset visual generator
        resetVisualGenerator();

        // Clean up the modal
        resetModal();
    }
});

// Reset the modal and perform clean-up activities
function resetModal() {

    // Hide all the select-box when the modal is closed
    $(".select-items").addClass(selectHideClass);

    // Show the Edit icon-div and hide the authoring container-div
    editArea.show();
    visualAuthoringArea.hide();

    // Disable the create button
    createVisualButton.prop("disabled", true);

    // Uncheck the visual properties checkbox
    visualPropertiesCheckboxes.prop("checked", false);

    // Enable all the toggle wrappers
    toggleWrappers.removeClass(toggleWrappersDisabledClass);

    $(".slider").addClass(disabledSliders);

    // Enable the toggle sliders for properties
    legendToggle.prop("disabled", false);
    xAxisToggle.prop("disabled", false);
    yAxisToggle.prop("disabled", false);
}

// Set accessibility insights for the report
function setAccessibilityInsights() {
    baseReportState.report.setComponentTitle("Playground showcase quick visual creator");
    baseReportState.report.setComponentTabIndex(0);
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
        rearrangeInCustomLayout();

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
        generatorType.removeClass(disabledClass);
        generatorType.removeClass(generatorTypeDisabledClass);

        console.log("Report render successful");
    });

    // Listen the commandTriggered event
    baseReportState.report.on("commandTriggered", function (event) {

        // Open the modal and set the fields, properties and title for the visual
        openModal(event.detail);
    });

    baseReportState.report.on("buttonClicked", function () {
        // Show the modal
        openModal();
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

    // Clear any other loaded handler events
    visualCreatorShowcaseState.report.off("loaded");

    // Triggers when a report schema is successfully loaded
    visualCreatorShowcaseState.report.on("loaded", async function () {

        const pages = await visualCreatorShowcaseState.report.getPages();

        // pages[1] is an empty page, on which the visual would be created
        pages[1].setActive().catch(error => console.log(error));
        visualCreatorShowcaseState.page = pages[1];
    });

    // Clear any other rendered handler events
    visualCreatorShowcaseState.report.off("rendered");

    // Triggers when a report is successfully embedded in UI
    visualCreatorShowcaseState.report.on("rendered", function () {
        console.log("Visual authoring report render successful");
    });

    // Clear any other error handler events
    visualCreatorShowcaseState.report.off("error");

    // Handle embed errors
    visualCreatorShowcaseState.report.on("error", function (event) {
        console.error(event.detail);
    });
}

// Render all visuals with 3x3 custom layout
function rearrangeInCustomLayout() {

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
    let visualsTotalWidth = reportWidth - (visualCreatorShowcaseConstants.margin * (visualCreatorShowcaseConstants.columns + 1));

    // Calculate the width of a single visual, according to the number of columns
    // For one and three columns visuals width will be a third of visuals total width
    let visualWidth = visualsTotalWidth / visualCreatorShowcaseConstants.columns;

    // Building visualsLayout object
    // You can find more information at https://github.com/Microsoft/PowerBI-JavaScript/wiki/Custom-Layout
    let visualsLayout = {};

    // Visuals starting point
    let x = visualCreatorShowcaseConstants.margin;
    let y = visualCreatorShowcaseConstants.margin;

    // Calculate visualHeight with margins
    let visualHeight = visualWidth * visualCreatorShowcaseConstants.visualAspectRatio;

    // Calculate the number of rows
    let rows = 0;

    // Do not count the overlapping visuals in generating the rows and final report height
    rows = Math.ceil((visuals.length - 2) / visualCreatorShowcaseConstants.columns);
    reportHeight = Math.max(reportHeight, (rows * visualHeight) + (rows + 1) * visualCreatorShowcaseConstants.margin);

    visuals.forEach((visual) => {
        // Hide the main and overlapping visuals, if new visual is being created
        if (visualCreationInProgress) {
            // Hide the visual
            if (visual.name === mainVisualGuid || visual.name === imageVisual.name || visual.name === actionButtonVisual.name) {
                visualsLayout[visual.name] = {
                    displayState: {
                        mode: models.VisualContainerDisplayMode.Hidden
                    }
                }
                return;
            }
        }
        if (visual.name === mainVisualGuid) {
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
                x: mainVisualState.x + mainVisualState.width * actionButtonVisual.ratio.xPositionRatioWithMainVisual,
                y: imageVisual.height + imageVisual.yPos + 18,
                width: mainVisualState.width * actionButtonVisual.ratio.widthRatioWithMainVisual - 4.5,
                // Set minimum height for action button visual in smaller screens
                height: Math.max(mainVisualState.height * actionButtonVisual.ratio.heightRatioWithMainVisual - 4.5, actionButtonVisual.height),
                displayState: {

                    // Change the selected visuals display mode to visible
                    mode: models.VisualContainerDisplayMode.Visible
                }
            };
        }

        // For remaining visuals, position them and update the x and y coordinates
        else {
            if (visual.name === mainVisualGuid) {
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
            x += visualWidth + visualCreatorShowcaseConstants.margin;

            // Reset x
            if (x + visualWidth > reportWidth) {
                x = visualCreatorShowcaseConstants.margin;
                y += visualHeight + visualCreatorShowcaseConstants.margin;
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
        resetContainerHeight(reportHeight + visualCreatorShowcaseConstants.margin);

        settings.customLayout.displayOption = models.DisplayOption.FitToWidth;
    }

    // Call updateSettings with the new settings object
    baseReportState.report.updateSettings(settings);
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
    for (let i = 0; i < styledSelects.length; i++) {
        const selectedElement = styledSelects[i].getElementsByTagName("select")[0];

        // For each element, create a new div that will act as the selected item
        const dropdownElement = document.createElement("div");
        dropdownElement.setAttribute("class", "select-selected");
        dropdownElement.setAttribute("id", "selected-value-" + i);
        dropdownElement.innerHTML = selectedElement.options[selectedElement.selectedIndex].innerHTML;
        styledSelects[i].appendChild(dropdownElement);

        // For each element, create a new div that will contain the option list
        const dropdownItem = document.createElement("div");
        dropdownItem.setAttribute("class", "select-items select-hide");
        for (let j = 1; j < selectedElement.length; j++) {

            // For each option in the original select element,
            // create a new div that will act as an option item
            const optionItem = document.createElement("div");
            optionItem.innerHTML = selectedElement.options[j].innerHTML;

            // Adding new click event listener
            optionItem.addEventListener("click", function () {

                // When an item is clicked, update the original select box, and the selected item
                let selects, previousSibling;
                selects = this.parentNode.parentNode.getElementsByTagName("select")[0];
                previousSibling = this.parentNode.previousSibling;
                for (let i = 0; i < selects.length; i++) {
                    if (selects.options[i].innerHTML === this.innerHTML) {
                        selects.selectedIndex = i;
                        previousSibling.innerHTML = this.innerHTML;
                        let childElements = this.parentNode.getElementsByClassName(sameAsSelectedClass);
                        for (let k = 0; k < childElements.length; k++) {
                            childElements[k].removeAttribute("class");
                        }

                        this.setAttribute("class", sameAsSelectedClass);
                        break;
                    }
                }

                previousSibling.click();

                // Changing the visual type or updating the data role field, according to the dropdown id
                if (selects.id === "visual-type") {
                    changeVisualType(previousSibling.innerHTML);
                    editArea.hide();
                    visualAuthoringArea.show();
                } else {
                    updateDataRoleField(selects.parentNode.parentNode.children[0].id, previousSibling.innerHTML);
                }
            });

            dropdownItem.appendChild(optionItem);
        }

        styledSelects[i].appendChild(dropdownItem);

        // Adding new click event listener for the select box
        dropdownElement.addEventListener("click", function (event) {
            // When the select box is clicked, close any other select boxes,
            // and open/close the current select box
            event.stopPropagation();
            closeAllSelect(this);
            this.nextSibling.classList.toggle(selectHideClass);
            this.classList.toggle("select-arrow-active");
        });
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
        } else {
            selected[i].classList.remove("select-arrow-active");
        }
    }

    for (let i = 0; i < selectItems.length; i++) {
        if (arrNo.indexOf(i)) {
            selectItems[i].classList.add(selectHideClass);
        }
    }
}

// Changing the visual type
async function changeVisualType(visualTypeDisplayName) {

    // Remove all data-fields from the state if visual is being edited
    if (selectedVisual.visual) {
        resetGeneratorDataRoles();
    }

    // Get the visual type from the display name
    const visualTypeData = visualTypeToDataRoles.filter((function (e) { return e.displayName === visualTypeDisplayName }))[0];
    const visualTypeName = visualTypeData.name;

    // Retrieve the visual's capabilities
    const capabilities = await baseReportState.report.getVisualCapabilities(visualTypeName);

    // Validating data roles existence on the given visual type
    if (!validateDataRoles(capabilities, visualTypeData.dataRoles)) {
        resetVisualGenerator();
        handleInvalidDataRoles();
        return;
    }

    // Enable the fields section
    generatorFields.removeClass(disabledClass);
    generatorFields.removeClass(generatorFieldsDisabledClass);

    // Disable the properties section
    generatorProperties.addClass(disabledClass);

    $(".slider").addClass(disabledSliders);

    // Disable the Create button
    createVisualButton.prop("disabled", true);

    // Reset all the properties
    resetGeneratorProperties();

    // Reset the data fields count
    visualCreatorShowcaseState.dataFieldsCount = 0;

    // If the visual doesn't exist, create new visual, otherwise, delete the old visual and create new visual
    if (!visualCreatorShowcaseState.newVisual) {
        await visualCreatorShowcaseState.page.createVisual(visualTypeName, getVisualLayout());
        updateVisualType(visualTypeName, visualTypeData.dataRoles);
    }
    else if (visualTypeName !== visualCreatorShowcaseState.visualType) {
        await visualCreatorShowcaseState.page.deleteVisual(visualCreatorShowcaseState.newVisual.name);
        await visualCreatorShowcaseState.page.createVisual(visualTypeName, getVisualLayout());
        updateVisualType(visualTypeName, visualTypeData.dataRoles);
    }
}

// Update showcase after visual type change
function updateVisualType(visualTypeName, dataRoles) {
    visualCreatorShowcaseState.report.updateSettings(visualHeaderReportSetting);
    updateCurrentVisualState(visualTypeName);
    resetGeneratorDataRoles();
    updateAvailableDataRoles(dataRoles);
    updateDropdownsVisibility();

    // Uncheck all the properties checkbox
    visualPropertiesCheckboxes.prop("checked", false);

    // Disable the properties div
    generatorProperties.addClass(generatorPropertiesDisabledClass);

    // Show the disabled items on visual type is changed
    disabledEraseTool.show();
    enabledEraseTool.hide();
    disabledAligns.show();
    enabledAligns.hide();
}

// Update the visual state
async function updateCurrentVisualState(visualTypeName) {

    const visuals = await visualCreatorShowcaseState.page.getVisuals();

    // Update visual and visual type
    visualCreatorShowcaseState.newVisual = visuals[0];
    visualCreatorShowcaseState.visualType = visualTypeName;

    // Enabling the pie chart legend (disabled by default)
    if (visualTypeName === "pieChart") {
        visualCreatorShowcaseState.newVisual.setProperty(propertyToSelector("legend"), { schema: schemas.property, value: true });
    }

    // Formatting the title to be more accessible
    visualCreatorShowcaseState.newVisual.setProperty(propertyToSelector("titleSize"), { schema: schemas.property, value: 25 });
    visualCreatorShowcaseState.newVisual.setProperty(propertyToSelector("titleColor"), { schema: schemas.property, value: "#000" });

    // Disabling unavailable properties for specific visual types
    toggleWrappers.removeClass(toggleWrappersDisabledClass);
    for (let i = 0; i < showcaseProperties.length; i++) {
        if (visualTypeProperties[visualTypeName].indexOf(showcaseProperties[i]) < 0) {

            // Uncheck the unavailable properties for the created visual
            $("#" + showcaseProperties[i] + "-toggle").prop("checked", false);

            // Disable the pointer events for the properties
            $("#" + showcaseProperties[i] + ".toggle-wrapper").addClass(toggleWrappersDisabledClass);
        }
    }
}

// Update the data roles and the data roles fields, on the dropdown menus
function updateAvailableDataRoles(dataRoles) {
    const dataRolesNamesElements = document.querySelectorAll(".inline-select-text");

    // Get the select wrappers to change the title to the data-roles
    const selectWrappers = $(".select-selected").slice(1);

    for (let i = 0; i < dataRoles.length; i++) {
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
    for (let i = 0; i < dataFields.length; i++) {
        fieldDivElements[i].innerHTML = dataFields[i];
        fieldOptionElements[i + 1].innerHTML = dataFields[i];
    }
}

// Update visual types on UI
function updateAvailableVisualTypes() {
    const typesDivElements = $(".select-items")[0].children;
    const typesOptionElements = $("#visual-type")[0].children;
    for (let i = 0; i < visualTypeToDataRoles.length; i++) {
        typesDivElements[i].innerHTML = visualTypeToDataRoles[i].displayName;
        typesOptionElements[i + 1].innerHTML = visualTypeToDataRoles[i].displayName;
    }
}

// Update data roles field on the visual
async function updateDataRoleField(dataRole, field) {

    // Check if the requested field is not the same as the selected field
    if (field !== visualCreatorShowcaseState.dataRoles[dataRole]) {

        // Getting the visual capabilities
        const capabilities = await visualCreatorShowcaseState.newVisual.getCapabilities();

        // Getting the data role name
        const dataRoleName = capabilities.dataRoles.filter(function (dr) { return dr.displayName === dataRole })[0].name;

        // Remove whitespace from field
        const dataFieldKey = field.replace(/\s+/g, "");

        // Check if the data role already has a field
        if (visualCreatorShowcaseState.dataRoles[dataRoleName]) {

            // If the data role has a field, remove it
            await visualCreatorShowcaseState.newVisual.removeDataField(dataRoleName, 0);
            visualCreatorShowcaseState.dataFieldsCount--;

            // If there are no more data fields, recreating the visual before adding the data field
            if (visualCreatorShowcaseState.dataFieldsCount === 0) {
                await visualCreatorShowcaseState.page.createVisual(visualCreatorShowcaseState.visualType, getVisualLayout());

                const visuals = await visualCreatorShowcaseState.page.getVisuals();
                visualCreatorShowcaseState.newVisual = visuals[0];
                visualCreatorShowcaseState.dataFieldsCount++;
                visualCreatorShowcaseState.dataRoles[dataRoleName] = dataFieldKey;
                await visualCreatorShowcaseState.newVisual.addDataField(dataRoleName, dataFieldsTargets[dataFieldKey]);
                updateDropdownsVisibility();
            } else {
                visualCreatorShowcaseState.dataFieldsCount++;
                visualCreatorShowcaseState.dataRoles[dataRoleName] = dataFieldKey;
                visualCreatorShowcaseState.newVisual.addDataField(dataRoleName, dataFieldsTargets[dataFieldKey]);
                updateDropdownsVisibility()
            }

        } else {

            // Adding a new field
            visualCreatorShowcaseState.dataRoles[dataRoleName] = dataFieldKey;
            await visualCreatorShowcaseState.newVisual.addDataField(dataRoleName, dataFieldsTargets[dataFieldKey]);
            updateDropdownsVisibility();
            visualCreatorShowcaseState.dataFieldsCount++;

            // Showing the visual if there are 2 or more data fields
            if (visualCreatorShowcaseState.dataFieldsCount > 1) {
                generatorProperties.removeClass(disabledClass);
                generatorProperties.removeClass(generatorPropertiesDisabledClass);
                customTitleWrapper.removeClass(toggleWrappersDisabledClass);
                createVisualButton.prop("disabled", false);
                visualPropertiesCheckboxes.prop("checked", true);
                visualPropertiesCheckboxes.prop("disabled", false);
                updateAvailableProperties(visualCreatorShowcaseState.visualType);

                // Make title property active
                makeTitlePropActive();
                // Show the enabled items to change, align or clear the title
                disabledEraseTool.hide();
                enabledEraseTool.show();
                disabledAligns.hide();
                enabledAligns.show();
            }
        }
    }
}

function makeTitlePropActive() {
    const titleProp = $("#" + "title" + "-toggle").prop("checked", true);
    const relatedToggle = titleProp.next();
    relatedToggle.removeClass(disabledSliders);
}

// Update available properties as per visual type
function updateAvailableProperties(visualType) {
    for (let i = 0; i < showcaseProperties.length; i++) {
        if (visualTypeProperties[visualType].indexOf(showcaseProperties[i]) < 0) {
            $("#" + showcaseProperties[i] + "-toggle").prop("checked", false);
            const property = $("#" + showcaseProperties[i] + "-toggle");
            const relatedToggle = property.next();
            relatedToggle.addClass(disabledSliders);
        }
        else {
            $("#" + showcaseProperties[i] + "-toggle").prop("checked", true);
            const property = $("#" + showcaseProperties[i] + "-toggle");
            const relatedToggle = property.next();
            relatedToggle.removeClass(disabledSliders);
        }
    }
}

// Update the visibility of the dropdowns
function updateDropdownsVisibility() {
    $(".select-items div").show();

    // Hide the option which is selected in the dropdowns above
    const selected = $(".select-selected");
    selected.each(function () {
        const selectedValue = $(this).text();
        $(".select-items div:contains(" + selectedValue + ")").hide();
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

    // TODO : Temporary fix for bar chart as xAxis effect is coming as yAxis and vice-versa
    if (visualCreatorShowcaseState.visualType === "barChart") {
        if (propertyName === "xAxis") {
            propertyName = "yAxis";
        }
        else if (propertyName === "yAxis") {
            propertyName = "xAxis";
        }
    }

    visualCreatorShowcaseState.properties[propertyName] = newValue;

    // Setting the property on the visual
    visualCreatorShowcaseState.newVisual.setProperty(propertyToSelector(propertyName), { schema: schemas.property, value: newValue });
}

// Update the title alignment
function onAlignClicked(direction) {
    if (!visualCreatorShowcaseState.newVisual) {
        return;
    }

    alignmentBlocks.removeClass(selectedClass);
    $("#align-" + direction).addClass(selectedClass);
    visualCreatorShowcaseState.properties["titleAlign"] = direction;

    // Setting the property on the visual
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
    // Resetting the title text to auto generated
    visualCreatorShowcaseState.newVisual.resetProperty(propertyToSelector("titleText"));
}

// Update the title's text
function updateTitleText() {
    if (!visualCreatorShowcaseState.newVisual)
        return;

    const visualTitle = visualTitleText.val();
    customVisualTitle = visualTitle;

    // If the title is blank, reseting the title to auto generated
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
    $(".select-selected")[0].innerHTML = "Select visual type";
    $("#visual-type ~ .select-items > .same-as-selected").show();
    $("#visual-type ~ .select-items > .same-as-selected")[0].removeAttribute("class");
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

    alignmentBlocks.removeClass(selectedClass);
    alignLeft.addClass(selectedClass);
    visualTitleText.val("");
}

// Reset the visual generator (data roles, properties and visual)
function resetVisualGenerator() {

    if (!visualCreatorShowcaseState.newVisual) {
        resetPropertiesWrapper();
        return;
    }

    generatorFields.addClass(disabledClass);
    generatorProperties.addClass(disabledClass);
    generatorFields.addClass(generatorFieldsDisabledClass);
    generatorProperties.addClass(generatorPropertiesDisabledClass);

    // Reset data-roles, properties, modal
    resetGeneratorDataRoles();
    resetGeneratorProperties();
    resetGeneratorVisual();
    resetPropertiesWrapper();
}

// Validate the existence of each dataRole on the visual's capabilities
function validateDataRoles(capabilities, dataRolesDisplayNames) {
    for (let i = 0; i < dataRolesDisplayNames.length; i++) {

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

// Show the disabled items on modal close
function resetPropertiesWrapper() {
    disabledEraseTool.show();
    enabledEraseTool.hide();
    disabledAligns.show();
    enabledAligns.hide();
}

// Create a visual and append that to the base report
// If the visual is selected for editing, this function will edit the visual, otherwise will create a visual
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

        // Formatting the title to be more accessible
        visual.setProperty(propertyToSelector("titleSize"), { schema: schemas.property, value: 13 });
        visual.setProperty(propertyToSelector("titleColor"), { schema: schemas.property, value: "#000" });

        // Enabling the legend property for Pie chart
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
            if (visualCreatorShowcaseState.visualType === "pieChart" && (propertyName === "xAxis" || propertyName === "yAxis")) {
                return;
            }
            if ((visualCreatorShowcaseState.visualType === "columnChart" || visualCreatorShowcaseState.visualType === "barChart") && (propertyName === "legend")) {
                return;
            }
            visual.setProperty(propertyToSelector(propertyName), { schema: schemas.property, value: propertyValue });
        });

        // Disabling the legend for the column and bar charts
        if (visualCreatorShowcaseState.visualType === "columnChart" || visualCreatorShowcaseState.visualType === "barChart") {
            visual.setProperty(propertyToSelector("legend"), { schema: schemas.property, value: false });
        }

        // Remove the data-roles which are null
        Object.keys(visualCreatorShowcaseState.dataRoles).forEach((key) => (visualCreatorShowcaseState.dataRoles[key] === null) && delete visualCreatorShowcaseState.dataRoles[key]);

        // Add data-fields to the created visual
        Object.entries(visualCreatorShowcaseState.dataRoles).forEach(dataField => {
            const [dataRole, field] = dataField;
            visual.addDataField(dataRole, dataFieldsTargets[field]);
        });

        customVisualTitle = "";

        // Append the created visual to the visuals state of the report
        baseReportState.visuals.push(visual);

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

        // Visual creation is completed
        visualCreationInProgress = false;
        rearrangeInCustomLayout();
    }
    else {
        if (visualTitleText.val() !== "") {
            customVisualTitle = visualTitleText.val();
        }

        const oldVisualType = selectedVisual.visual.type;
        const oldVisual = selectedVisual.visual;
        if (oldVisualType !== visualCreatorShowcaseState.visualType) {
            await oldVisual.changeType(visualCreatorShowcaseState.visualType);
        }

        // Formatting the title to be more accessible
        oldVisual.setProperty(propertyToSelector("titleSize"), { schema: schemas.property, value: 13 });
        oldVisual.setProperty(propertyToSelector("titleColor"), { schema: schemas.property, value: "#000" });

        // Enabling the legend property for Pie chart
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
                if (visualCreatorShowcaseState.visualType === "pieChart" && (propertyName === "xAxis" || propertyName === "yAxis")) {
                    return;
                }
                if ((visualCreatorShowcaseState.visualType === "columnChart" || visualCreatorShowcaseState.visualType === "barChart") && (propertyName === "legend")) {
                    return;
                }
                oldVisual.setProperty(propertyToSelector(propertyName), { schema: schemas.property, value: propertyValue });
            }
        });

        // Disabling the legend for the column and bar charts
        if (visualCreatorShowcaseState.visualType === "columnChart" || visualCreatorShowcaseState.visualType === "barChart") {
            oldVisual.setProperty(propertyToSelector("legend"), { schema: schemas.property, value: false });
        }

        // Remove the data-roles which are null
        Object.keys(visualCreatorShowcaseState.dataRoles).forEach((key) => (visualCreatorShowcaseState.dataRoles[key] === null) && delete visualCreatorShowcaseState.dataRoles[key]);

        // Add data-fields to the created visual
        Object.entries(visualCreatorShowcaseState.dataRoles).forEach(async function (dataField) {
            const [dataRole, field] = dataField;

            // Get data-fields from the data-role
            const dataFieldProp = await oldVisual.getDataFields(dataRole);

            if (dataFieldProp.length === 0) {
                oldVisual.addDataField(dataRole, dataFieldsTargets[field]);
            }
            else {
                await oldVisual.removeDataField(dataRole, 0);
                oldVisual.addDataField(dataRole, dataFieldsTargets[field]);
            }
        });
        customVisualTitle = "";
        // Use the visual from the state to update it's properties
        selectedVisual.visual = null;
    }

    // Reset the dropdowns, authoring-div and modal
    resetVisualGenerator();
}

// This function opens the modal and fill the dropdowns with the data-roles, properties and title of the visual
async function openModal(visualData) {

    if (!visualData) {
        // If visualData is not preset, just show the modal
        visualCreatorModal.modal("show");
        return;
    }

    // Pass the visual to get the IVisual response
    const visualResponse = await getIVisualResponse(visualData.visual);
    selectedVisual.visual = visualResponse;

    const visualType = visualResponse.type;
    const visualDataRole = visualTypeToDataRoles.filter((function (e) { return e.name === visualType }))[0];

    visualDataRole.dataRoleNames.forEach(async function (dataRole) {

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

    // Add properties to the created visual
    Object.entries(visualCreatorShowcaseState.properties).forEach(async function (visualProperty) {
        const [propertyName, propertyValue] = visualProperty;
        const property = await visualResponse.getProperty(propertyToSelector(propertyName));
        visualCreatorShowcaseState.properties[propertyName] = property.value;
    });

    // This will create visual inside the modal
    const newVisual = await visualCreatorShowcaseState.page.createVisual(visualType, getVisualLayout());

    // Update state
    visualCreatorShowcaseState.newVisual = newVisual.visual;
    visualCreatorShowcaseState.visualType = newVisual.visual.type;
    const visual = newVisual.visual;

    // Formatting the title to be more accessible
    visual.setProperty(propertyToSelector("titleSize"), { schema: schemas.property, value: 25 });
    visual.setProperty(propertyToSelector("titleColor"), { schema: schemas.property, value: "#000" });

    // Enabling the legend property for Pie chart
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
        if (visualCreatorShowcaseState.visualType === "pieChart" && (propertyName === "xAxis" || propertyName === "yAxis")) {
            return;
        }
        if ((visualCreatorShowcaseState.visualType === "columnChart" || visualCreatorShowcaseState.visualType === "barChart") && (propertyName === "legend")) {
            return;
        }
        visual.setProperty(propertyToSelector(propertyName), { schema: schemas.property, value: propertyValue });
    });

    // Disabling the legend for the column and bar charts
    if (visualCreatorShowcaseState.visualType === "columnChart" || visualCreatorShowcaseState.visualType === "barChart") {
        visual.setProperty(propertyToSelector("legend"), { schema: schemas.property, value: false });
    }

    // Remove the data-roles which are empty from the state
    Object.keys(visualCreatorShowcaseState.dataRoles).forEach((key) => (visualCreatorShowcaseState.dataRoles[key] === null) && delete visualCreatorShowcaseState.dataRoles[key]);

    // Add data-fields to the created visual
    Object.entries(visualCreatorShowcaseState.dataRoles).forEach(dataField => {
        const [dataRole, field] = dataField;
        visual.addDataField(dataRole, dataFieldsTargets[field]);
    });

    // Update data-roles for the given visual type in the UI
    updateAvailableDataRoles(visualDataRole.dataRoles);

    // Make the title property active
    makeTitlePropActive();

    // Populate properties as per state inside the modal
    populateProperties(visualCreatorShowcaseState);

    // Remove disabled class from data-roles and properties
    generatorFields.removeClass(generatorFieldsDisabledClass);
    generatorFields.removeClass(disabledClass);
    generatorProperties.removeClass(generatorPropertiesDisabledClass);
    generatorProperties.removeClass(disabledClass);

    // Hide the disabled-erase-tool and alignments
    disabledEraseTool.hide();
    enabledEraseTool.show();
    disabledAligns.hide();
    enabledAligns.show();

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
    const visualTypeData = visualTypeToDataRoles.filter((function (e) { return e.name === visualCreatorShowcaseState.visualType }))[0];
    const visualTypeName = visualTypeData.displayName;

    // Set the type of the visual in visual-type dropdown
    $("#selected-value-0").text(visualTypeName);
    const visualSelectItems = $(".select-items").get(0).children;
    Array.from(visualSelectItems).forEach(visualSelectItem => {
        if (visualSelectItem.innerHTML === visualTypeName) {
            visualSelectItem.classList.add(sameAsSelectedClass);
        }
    });

    // Set the data-roles for the visual
    Object.entries(visualCreatorShowcaseState.dataRoles).forEach(dataField => {
        const [dataRole, field] = dataField;
        const index = visualTypeData.dataRoleNames.indexOf(dataRole);
        const dataRoleField = dataFieldsMappings[field];
        if (index !== -1) {
            const value = visualTypeData.dataRoles[index];
            selectDataRoles(value, dataRoleField);
        }
    });

    // Set the properties for the visual
    for (let i = 0; i < showcaseProperties.length; i++) {
        if (visualTypeProperties[visualCreatorShowcaseState.visualType].indexOf(showcaseProperties[i]) < 0) {

            // Uncheck the inapplicable properties for the created visual
            $("#" + showcaseProperties[i] + "-toggle").prop("checked", false);

            // Disable the pointer events for the properties
            $("#" + showcaseProperties[i] + ".toggle-wrapper").addClass(toggleWrappersDisabledClass);
            const property = $("#" + showcaseProperties[i] + "-toggle");
            const relatedToggle = property.next();
            relatedToggle.addClass(disabledSliders);

        }
        else {
            $("#" + showcaseProperties[i] + "-toggle").prop("checked", true);
            const property = $("#" + showcaseProperties[i] + "-toggle");
            const relatedToggle = property.next();
            relatedToggle.removeClass(disabledSliders);
        }
    }

    Object.entries(visualCreatorShowcaseState.properties).forEach(property => {
        let [propertyName, propertyValue] = property;
        if (visualCreatorShowcaseState.visualType === "barChart") {
            if (propertyName === "xAxis") {
                propertyName = "yAxis";
            }
            else if (propertyName === "yAxis") {
                propertyName = "xAxis";
            }
        }

        if (propertyName === "titleAlign") {
            if (propertyValue === "center" || propertyValue === "right") {
                alignmentBlocks.removeClass(selectedClass);
                $("#align-" + propertyValue).addClass(selectedClass);
            }
        }

        if (propertyName === "titleText") {
            if (typeof propertyValue !== "object" && propertyValue !== null) {
                visualTitleText.val(propertyValue);
            }
        }
        $("#" + propertyName + "-toggle").prop("checked", propertyValue);
    });

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

    for (let i = 0; i < selectSpanWrappers.length; i++) {
        if (selectSpanWrappers[i].innerHTML === dataRoleName) {
            $("#selected-value-" + (i + 1)).text(dataRoleValue);
            const visualSelectItems = $(".select-items").get(i + 1).children;
            Array.from(visualSelectItems).forEach(visualSelectItem => {
                if (visualSelectItem.innerHTML === dataRoleValue) {
                    visualSelectItem.classList.add(sameAsSelectedClass);
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