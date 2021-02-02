const visualCreatorShowcaseState = {
    report: null,
    page: null, // The page from where the 3x3 visuals will be displayed
    newVisual: null, // New visual to be created on the page for the base-report
    visualType: null,
    dataRoles: {
        Legend: null,
        Values: null,
        Axis: null,
        Tooltips: null,
    },
    dataFieldsCount: 0,
    properties: {
        legend: true,
        xAxis: true,
        yAxis: true,
        title: true,
        titleText: null,
        titleAlign: null
    },
}

const selectedVisual = {
    visual: null,
}

const baseReportState = {
    report: null,
    visuals: null,
    page: null
}

const visualCreatorShowcaseConstants = {
    columns: 3,
    margin: 16,
    visualAspectRatio: 9 / 16
}

// Constants used for report configurations as key-value pair
const reportConfig = {
    accessToken: null,
    embedUrl: null,
    reportId: null,
}

// Visual overlapping
const mainVisualGuid = "a6d74a71de4135e00a59";

const imageVisual = {
    name: "2270e4eea9242400a0cd",
    yPos: undefined,
    height: undefined,
    ratio: {
        widthRatioWithMainVisual: 36 / 426,
        heightRatioWithMainVisual: 36 / 252,
        xPositionRatioWithMainVisual: 195 / 426,
        yPositionRatioWithMainVisual: 90 / 252
    }
}

const actionButtonVisual = {
    name: "946862f32d49b6573406",
    height: 32,
    ratio: {
        widthRatioWithMainVisual: 151 / 426,
        heightRatioWithMainVisual: 32 / 252,
        xPositionRatioWithMainVisual: 135 / 426,
        yPositionRatioWithMainVisual: 144 / 252
    }
}

// Cache DOM Elements
const overlay = $("#overlay");
const visualDisplayArea = $("#visual-authoring-container").get(0);
const editArea = $("#edit-area");
const visualAuthoringArea = $("#visual-authoring-container");
const visualTypeDropdown = $("#selected-value-0");
const createVisualButton = $("#create-visual-btn");
const generatorType = $("#generator-type");
const generatorFields = $("#generator-fields");
const generatorProperties = $("#generator-properties");
const disabledEraseTool = $("#erase-tool-disabled");
const enabledEraseTool = $("#erase-tool-enabled");
const disabledAligns = $("#aligns-disabled");
const enabledAligns = $("#aligns-enabled");
const visualCreatorModal = $("#visual-creator");
const visualTitleText = $("#visual-title");
const legendToggle = $("#legend-toggle");
const xAxisToggle = $("#xAxis-toggle");
const yAxisToggle = $("#yAxis-toggle");
const titleToggle = $("#title-toggle");
const alignRight = $("#align-right");
const closeModalButton = $("#close-modal");
const alignLeft = $("#align-left");
const reportContainer = $(".report-container").get(0);
const customTitleWrapper = $(".custom-title-wrapper");
const alignmentBlocks = $(".alignment-block");
const visualPropertiesCheckboxes = $(".property-checkbox");
const toggleWrappers = $(".toggle-wrapper");

// Cache showcasePropertiesLength
const showcasePropertiesLength = showcaseProperties.length;

// Get models. models contain enums that can be used
const models = window["powerbi-client"].models;

// CSS Classes
const disabledClass = "generator-disabled";
const selectHideClass = "select-hide";
const generatorTypeDisabledClass = "generator-type-disabled";
const generatorFieldsDisabledClass = "generator-fields-disabled";
const generatorPropertiesDisabledClass = "generator-properties-disabled";
const selectedClass = "selected";
const sameAsSelectedClass = "same-as-selected";
const toggleWrappersDisabledClass = "disabled";
const disabledSliders = "disabled-sliders";
const visualTypeId = "selected-value-0";

// Key codes
const KEYCODE_TAB = 9;
const KEYCODE_ENTER = 13;
const KEYCODE_ESCAPE = 27;
const KEYCODE_SPACE = 32;

// Store the position of the main visual [basicShape]
let mainVisualState;

// Get the reference for the iframe inside the modal to remove it from the tab-order
let authoringiFrame;

// Custom title for the visual
let customVisualTitle = "";

// To store the state of the visual creation
let visualCreationInProgress = false;

// To apply setting to the new visual created in the Modal
const visualHeaderReportSetting = {
    visualSettings: {
        visualHeaders: [
            {
                settings: {
                    visible: false
                }
            }
        ]
    }
}

// Headers
const selectVisualTypeHeader = "Select visual type";