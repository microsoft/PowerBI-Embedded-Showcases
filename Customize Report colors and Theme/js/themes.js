// For report themes documentation please check https://docs.microsoft.com/en-us/power-bi/desktop-report-themes
const jsonDataColors = [{
    "name": "Default",
    "dataColors": ["#1A81FB", "#142091", "#E16338", "#5F076E", "#DA3F9D", "#6945B8", "#D3AA22", "#CF404A"],
    "foreground": "#252423",
    "background": "#FFFFFF",
    "tableAccent": "#B73A3A"
},
{
    "name": "Divergent",
    "dataColors": ["#B73A3A", "#EC5656", "#F28A90", "#F8BCBD", "#99E472", "#23C26F", "#0AAC00", "#026645"],
    "foreground": "#252423",
    "background": "#F4F4F4",
    "tableAccent": "#B73A3A"
},
{
    "name": "Executive",
    "dataColors": ["#3257A8", "#37A794", "#8B3D88", "#DD6B7F", "#6B91C9", "#F5C869", "#77C4A8", "#DEA6CF"],
    "background": "#FFFFFF",
    "foreground": "#9C5252",
    "tableAccent": "#6076B4"
},
{
    "name": "Tidal",
    "dataColors": ["#094782", "#0B72D7", "#098BF5", "#54B5FB", "#71C0A7", "#57B956", "#478F48", "#326633"],
    "tableAccent": "#094782",
    "visualStyles": {
        "*": {
            "*": {
                "background": [{ "show": true, "transparency": 3 }],
                "visualHeader": [{
                    "foreground": { "solid": { "color": "#094782" } },
                    "transparency": 3
                }]
            }
        },
        "group": { "*": { "background": [{ "show": false }] } },
        "basicShape": { "*": { "background": [{ "show": false }] } },
        "image": { "*": { "background": [{ "show": false }] } },
        "page": {
            "*": {
                "background": [{ "transparency": 100 }],
            }
        }
    }
}
];

const themes = [{
    "background": "#FFFFFF",
},
{
    "background": "#252423",
    "foreground": "#FFFFFF",
    "tableAccent": "#FFFFFF",
    "textClasses": {
		"title": {
			"color": "#FFF",
			"fontFace": "Segoe UI Bold"
        },
	},
    "visualStyles": {
        "*": {
            "*": {
                "*": [{
                    "fontFamily": "Segoe UI",
                    "color": { "solid": { "color": "#252423" } },
                    "labelColor": { "solid": { "color": "#FFFFFF" } },
                    "secLabelColor": { "solid": { "color": "#FFFFFF" } },
                    "titleColor": { "solid": { "color": "#FFFFFF" } },
                }],
                "labels": [{
                    "color": { "solid": { "color": "#FFFFFF" } }
                }],
                "categoryLabels": [{
                    "color": { "solid": { "color": "#FFFFFF" } }
                }]
            }
        }
    }
}
];
