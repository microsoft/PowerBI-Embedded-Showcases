const visualCreatorShowcaseState = {
    report: null,
    page: null, // The page from where the 3x3 visuals will be displayed
    newVisual: null, // New visual to be created on the page for the base-report
    visualType: null,
    dataRoles: {
        Legend: null,
        Values: null,
        Value: null,
        Axis: null,
        Tooltips: null,
        "Y Axis": null,
        Category: null,
        Breakdown: null,
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
    visual: null
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

// Cache DOM Elements
const overlay = $("#overlay");
const reportContainer = $(".report-container").get(0);
const visualDisplayArea = $("#visual-authoring-container").get(0);
const editArea = $("#edit-area");
const visualAuthoringArea = $("#visual-authoring-container");
const createVisualButton = $("#create-visual-btn");
const generatorType = $("#generator-type");
const generatorFields = $("#generator-fields");
const generatorProperties = $("#generator-properties");
const customTitleWrapper = $(".custom-title-wrapper");
const disabledEraseTool = $("#erase-tool-disabled");
const enabledEraseTool = $("#erase-tool-enabled");
const disabledAligns = $("#aligns-disabled");
const enabledAligns = $("#aligns-enabled");
const visualCreatorModal = $("#visual-creator");
const visualTitleText = $("#visual-title");
const alignmentBlocks = $(".alignment-block");
const visualPropertiesCheckboxes = $(".property-checkbox");
const toggleWrappers = $(".toggle-wrapper");
const legendToggle = $("#legend-toggle");
const xAxisToggle = $("#xAxis-toggle");
const yAxisToggle = $("#yAxis-toggle");
const titleToggle = $("#title-toggle");
const closeModalButton = $("#close-modal");
const alignLeft = $("#align-left");
const contosoLogo = $("#image-btn");

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

// Custom title for the visual
let customVisualTitle = "";