// ----------------------------------------------------------------------------
// Copyright (c) Microsoft Corporation.
// Licensed under the MIT license.
// ----------------------------------------------------------------------------

// Define the available data roles for the visual types
const visualTypeToDataRoles = [
    { name: "columnChart", displayName: "Column chart", dataRoleNames: ["Category", "Y", "Tooltips"] },
    { name: "areaChart", displayName: "Area chart", dataRoleNames: ["Category", "Series", "Y"] },
    { name: "barChart", displayName: "Bar chart", dataRoleNames: ["Category", "Y", "Tooltips"] },
    { name: "pieChart", displayName: "Pie chart", dataRoleNames: ["Category", "Y", "Tooltips"] },
    { name: "lineChart", displayName: "Line chart", dataRoleNames: ["Category", "Series", "Y"] },
];

// Define the available fields for each data role
const dataRolesToFields = [
   { dataRole: "Axis", dataRoleName:"Category", Fields: ["Industry", "Opportunity Status", "Lead Rating", "Salesperson"] },
    { dataRole: "Values",dataRoleName:"Y", Fields: ["Actual Revenue", "Estimated Revenue", "Number of Opportunities", "Salesperson"] },
    { dataRole: "Legend",dataRoleName:"Series", Fields: ["Industry", "Lead Rating", "Opportunity Status", "Salesperson"] },
    { dataRole: "Tooltips",dataRoleName:"Tooltips", Fields: ["Industry", "Actual Close Date", "Actual Revenue", "Estimated Revenue"] },
];

// Define schemas for visuals API
const schemas = {
    column: "http://powerbi.com/product/schema#column",
    measure: "http://powerbi.com/product/schema#measure",
    property: "http://powerbi.com/product/schema#property",
    default: "http://powerbi.com/product/schema#default",
};

// Define mapping from fields to target table and column/measure
const dataFieldsTargets = {
    ActualRevenue: { column: "Actual Revenue", table: "QVC Report", schema: schemas.column },
    NumberofOpportunities: { measure: "Number of Opportunities", table: "QVC Report", schema: schemas.measure },
    Salesperson: { column: "Salesperson", table: "QVC Report", schema: schemas.column },
    EstimatedRevenue: { column: "Estimated Revenue", table: "QVC Report", schema: schemas.column },
    OpportunityStatus: { column: "Opportunity Status", table: "QVC Report", schema: schemas.column },
    Industry: { column: "Industry", table: "QVC Report", schema: schemas.column },
    LeadRating: { column: "Lead Rating", table: "QVC Report", schema: schemas.column },
    Salesperson: { column: "Salesperson", table: "QVC Report", schema: schemas.column },
    ActualCloseDate: { column: "Actual Close Date", table: "QVC Report", schema: schemas.column },
};

const dataFieldsMappings = {
    ActualRevenue: "Actual Revenue",
    NumberofOpportunities: "Number of Opportunities",
    Salesperson: "Salesperson",
    EstimatedRevenue: "Estimated Revenue",
    OpportunityStatus: "Opportunity Status",
    Industry: "Industry",
    LeadRating: "Lead Rating",
    Salesperson: "Salesperson",
    ActualCloseDate: "Actual Close Date"
}

// Define the available properties
const showcaseProperties = ["legend", "xAxis", "yAxis"];

// Define title related properties
const titleProperties = ["title", "titleText", "titleAlign"];

const visualTypeProperties = {
    columnChart: ["xAxis", "yAxis"],
    areaChart: ["legend", "xAxis", "yAxis"],
    barChart: ["xAxis", "yAxis"],
    pieChart: ["legend"],
    lineChart: ["legend", "xAxis", "yAxis"]
};