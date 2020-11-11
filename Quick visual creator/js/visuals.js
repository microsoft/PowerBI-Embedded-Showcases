// Define the available data roles for the visual types
const visualTypeToDataRoles = [
    { name: "columnChart", displayName: "Column chart", dataRoles: ["Axis", "Values", "Tooltips"], dataRoleNames: ["Category", "Y", "Tooltips"] },
    { name: "areaChart", displayName: "Area chart", dataRoles: ["Axis", "Legend", "Values"], dataRoleNames: ["Category", "Series", "Y"] },
    { name: "barChart", displayName: "Bar chart", dataRoles: ["Axis", "Values", "Tooltips"], dataRoleNames: ["Category", "Y", "Tooltips"] },
    { name: "pieChart", displayName: "Pie chart", dataRoles: ["Legend", "Values", "Tooltips"], dataRoleNames: ["Category", "Y", "Tooltips"] },
];

// Define the available fields for each data role
const dataRolesToFields = [
    { dataRole: "Legend", Fields: ["State", "Region", "Manufacturer"] },
    { dataRole: "Values", Fields: ["Total Units", "Total Category Volume", "Total Compete Volume"] },
    { dataRole: "Axis", Fields: ["State", "Region", "Manufacturer"] },
    { dataRole: "Value", Fields: ["Total Units", "Total Category Volume", "Total Compete Volume"] },
    { dataRole: "Y Axis", Fields: ["Total Units", "Total Category Volume", "Total Compete Volume"] },
    { dataRole: "Tooltips", Fields: ["Total Units", "Total Category Volume", "Total Compete Volume"] },
    { dataRole: "Category", Fields: ["State", "Region", "Date"] },
    { dataRole: "Breakdown", Fields: ["State", "Region", "Manufacturer"] },
];

// Define schemas for visuals API
const schemas = {
    column: "http://powerbi.com/product/schema#column",
    measure: "http://powerbi.com/product/schema#measure",
    property: "http://powerbi.com/product/schema#property",
};

// Define mapping from fields to target table and column/measure
const dataFieldsTargets = {
    State: { column: "State", table: "Geo", schema: schemas.column },
    Region: { column: "Region", table: "Geo", schema: schemas.column },
    District: { column: "District", table: "Geo", schema: schemas.column },
    Manufacturer: { column: "Manufacturer", table: "Manufacturer", schema: schemas.column },
    TotalUnits: { measure: "Total Units", table: "SalesFact", schema: schemas.measure },
    TotalCategoryVolume: { measure: "Total Category Volume", table: "SalesFact", schema: schemas.measure },
    TotalCompeteVolume: { measure: "Total Compete Volume", table: "SalesFact", schema: schemas.measure },
    Date: { measure: "Date", table: "Date", schema: schemas.measure },
};

const dataFieldsMappings = {
    State: "State",
    Region: "Region", 
    District: "District", 
    Manufacturer: "Manufacturer", 
    TotalUnits: "Total Units", 
    Date: "Date",
    TotalCategoryVolume: "Total Category Volume",
    TotalCompeteVolume: "Total Compete Volume"
}

// Define the available properties
const showcaseProperties = ["legend", "xAxis", "yAxis"];

const visualTypeProperties = {
    columnChart: ["xAxis", "yAxis"],
    areaChart: ["legend", "xAxis", "yAxis"],
    barChart: ["xAxis", "yAxis"],
    pieChart: ["legend"],
};