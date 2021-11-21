// ----------------------------------------------------------------------------
// Copyright (c) Microsoft Corporation.
// Licensed under the MIT license.
// ----------------------------------------------------------------------------

window.addEventListener('error', function(event) {
    // Protection against cross-origin failure
    try {
        if (window.parent.playground && window.parent.playground.logShowcaseError) {
            window.parent.playground.logShowcaseError("PersonalizeTopInsights", event);
        }
    } catch { }
});