// For the decision of the layout
const ColumnsNumber = {
    One: 1,
    Two: 2,
    Three: 3
};

// For the decision of two custom layout with spanning
const SpanType = {
    None: 0,
    RowSpan: 1,
    ColSpan: 2
};

// To give consistent margin to each visual in the custom showcase
const LayoutShowcaseConsts = {
    margin: 16,
    visualAspectRatio: 9 / 16,
};

// Constants used for report configurations as key-value pair
let reportConfig = {
    accessToken: null,
    embedUrl: null,
    reportId: null,
}

// Maintain the state for the showcase
let layoutShowcaseState = {
    columns: ColumnsNumber.Two,
    span: SpanType.None,
    layoutVisuals: null,
    layoutReport: null,
    layoutPageName: null
};

// Cache the report containers
const reportContainer = $("#report-container").get(0);