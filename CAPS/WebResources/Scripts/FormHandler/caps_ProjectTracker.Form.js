"use strict";

/* INCLUDE CAPS.Common.js */

var CAPS = CAPS || {};
CAPS.ProjectTracker = CAPS.ProjectTracker || {};

CAPS.ProjectTracker.onLoad = function (executionContext) {
    var formContext = executionContext.getFormContext();

    //Show Summary Report
    CAPS.ProjectTracker.showSummaryReport(formContext);
}

CAPS.ProjectTracker.showSummaryReport = function (formContext) {
    debugger;
    //Get iframe 
    var iframeObject = formContext.getControl("IFRAME_SummaryReport");

    if (iframeObject !== null) {
        var strURL = "/crmreports/viewer/viewer.aspx"
            + "?id=ed1744bc-98a6-ea11-a813-000d3af42496"
            + "&action=run"
            + "&context=records"
            + "&recordstype=10078"
            + "&records=" + formContext.data.entity.getId()
            + "&helpID=Monthly%20Project%20Summary.rdl";


        //Set URL of iframe
        iframeObject.setSrc(strURL);
    }
}

