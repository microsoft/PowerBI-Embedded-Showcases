const theme = {
    "name": "visualsTheme",
    "dataColors": [
        "#118dff",
        "#12239e",
        "#e66c37",
        "#6B007B",
        "#E044A7",
        "#744EC2",
        "#D9B300",
        "#D64550",
        "#000000"
    ],
    "visualStyles": {
        "*": {
            "*": {
                "dropShadow": [
                    {
                        "color": {
                            "solid": {
                                "color": "#000000"
                            }
                        },
                        "show": true,
                        "position": "Outer",
                        "preset": "Custom",
                        "shadowSpread": 1,
                        "shadowBlur": 3,
                        "angle": 45,
                        "shadowDistance": 1,
                        "transparency": 87
                    }
                ]
            },
        },
        "image": {
            "*": {
                "dropShadow": [
                    {
                        "show": false,
                    }
                ]
            },
        },
        "actionButton": {
            "*": {
                "dropShadow": [
                    {
                        "show": false,
                    }
                ]
            },
        }
    }
}