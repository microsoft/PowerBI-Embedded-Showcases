// Define the available data roles for the visual types
const visualTypeToDataRoles = [
    { name: "columnChart", displayName: "Column chart", dataRoles: ["Axis", "Values", "Tooltips"], dataRoleNames: ["Category", "Y", "Tooltips"] },
    { name: "areaChart", displayName: "Area chart", dataRoles: ["Axis", "Legend", "Values"], dataRoleNames: ["Category", "Series", "Y"] },
    { name: "barChart", displayName: "Bar chart", dataRoles: ["Axis", "Values", "Tooltips"], dataRoleNames: ["Category", "Y", "Tooltips"] },
    { name: "pieChart", displayName: "Pie chart", dataRoles: ["Legend", "Values", "Tooltips"], dataRoleNames: ["Category", "Y", "Tooltips"] },
    { name: "lineChart", displayName: "Line chart", dataRoles: ["Axis", "Legend", "Values"], dataRoleNames: ["Category", "Series", "Y"] },
];

// Define the available fields for each data role
const dataRolesToFields = [
    { dataRole: "Axis", Fields: ["Industry", "Salesperson", "Lead Rating"] },
    { dataRole: "Values", Fields: ["Actual Revenue", "Estimated Revenue", "Number of Opportunities"] },
    { dataRole: "Legend", Fields: ["Industry", "Salesperson", "Oppportunity Status"] },
    { dataRole: "Tooltips", Fields: ["Actual Close Date", "Estimated Revenue", "Actual Revenue"] },
];

// Define schemas for visuals API
const schemas = {
    column: "http://powerbi.com/product/schema#column",
    measure: "http://powerbi.com/product/schema#measure",
    property: "http://powerbi.com/product/schema#property",
};

// Define mapping from fields to target table and column/measure
const dataFieldsTargets = {
    ActualRevenue: { column: "Actual Revenue", table: "QVC Report", schema: schemas.column },
    NumberofOpportunities: { measure: "Number of Opportunities", table: "QVC Report", schema: schemas.measure },
    Salesperson: { column: "Salesperson", table: "QVC Report", schema: schemas.column },
    EstimatedRevenue: { column: "Estimated Revenue", table: "QVC Report", schema: schemas.column },
    OppportunityStatus: { column: "Oppportunity Status", table: "QVC Report", schema: schemas.column },
    Industry: { column: "Industry", table: "QVC Report", schema: schemas.column },
    LeadRating: { column: "Lead Rating", table: "QVC Report", schema: schemas.column },
    ActualCloseDate: { column: "Actual Close Date", table: "QVC Report", schema: schemas.column },
};

const dataFieldsMappings = {
    ActualRevenue: "Actual Revenue",
    NumberofOpportunities: "Number of Opportunities",
    Salesperson: "Salesperson",
    EstimatedRevenue: "Estimated Revenue",
    OppportunityStatus: "Oppportunity Status",
    Industry: "Industry",
    LeadRating: "Lead Rating",
    ActualCloseDate: "Actual Close Date"
}

// Define the available properties
const showcaseProperties = ["legend", "xAxis", "yAxis"];

const visualTypeProperties = {
    columnChart: ["xAxis", "yAxis"],
    areaChart: ["legend", "xAxis", "yAxis"],
    barChart: ["xAxis", "yAxis"],
    pieChart: ["legend"],
    lineChart: ["legend", "xAxis", "yAxis"]
};