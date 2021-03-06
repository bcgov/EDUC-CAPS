"use strict";

var CAPS = CAPS || {};
CAPS.ProgressReport = CAPS.ProgressReport || {};

/*
Main function for Project Tracker.  This function calls all other form functions.
*/
CAPS.ProgressReport.onLoad = function (executionContext) {
    var formContext = executionContext.getFormContext();

    formContext.getControl("sgd_FutureCashFlow").addOnLoad(CAPS.ProgressReport.UpdateTotalFutureCashFlow);
    //embed report
    CAPS.ProgressReport.ShowMonthlyReport(formContext);
};

/*
Sums up the total projected provincial forecast and projected actuals.
*/
CAPS.ProgressReport.UpdateTotalFutureCashFlow = function (executionContext) {
    debugger;

    var formContext = executionContext.getFormContext();
    var id = formContext.data.entity.getId().replace("{", "").replace("}", "");
    Xrm.WebApi.retrieveMultipleRecords("caps_cashflowprojection", "?$select=caps_provincialamount,caps_agencyamount&$filter=caps_ProgressReport/caps_progressreportid eq " + id).then(
        function success(result) {
            var totalProvincial = 0;
            var totalAgency = 0;
            for (var i = 0; i < result.entities.length; i++) {
                totalProvincial += result.entities[i].caps_provincialamount;
                totalAgency += result.entities[i].caps_agencyamount;
            }

            // perform operations on record retrieval
            formContext.getAttribute('caps_totalfutureprovincialprojection').setValue(totalProvincial);
            formContext.getAttribute('caps_totalfutureagencyprojection').setValue(totalAgency);
        },
        function (error) {
            console.log(error.message);
            // handle error conditions
        }
    );
};

/**
 * Function to embed the SSRS Monthly Progress Report on the form.
 * @param {any} formContext the form context
 */
CAPS.ProgressReport.ShowMonthlyReport = function (formContext) {
    debugger;
    var requestUrl = "/api/data/v9.1/EntityDefinitions?$filter=LogicalName eq 'caps_progressreport'&$select=ObjectTypeCode";

    var globalContext = Xrm.Utility.getGlobalContext();

    var req = new XMLHttpRequest();
    req.open("GET", globalContext.getClientUrl() + requestUrl, true);
    req.setRequestHeader("OData-MaxVersion", "4.0");
    req.setRequestHeader("OData-Version", "4.0");
    req.setRequestHeader("Accept", "application/json");
    req.setRequestHeader("Content-Type", "application/json; charset=utf-8");
    req.onreadystatechange = function () {
        if (this.readyState === 4) {
            req.onreadystatechange = null;
            if (this.status === 200) {
                debugger;
                var result = JSON.parse(this.response);
                var objectTypeCode = result.value[0].ObjectTypeCode;
                //use retrieved objectTypeCode
                //Get iframe 
                var iframeObject = formContext.getControl("IFRAME_SummaryReport");

                if (iframeObject !== null) {
                    var strURL = "/crmreports/viewer/viewer.aspx"
                        + "?id=ef515e3b-3c2f-eb11-a813-000d3af43595"
                        + "&action=run"
                        + "&context=records"
                        + "&recordstype=" + objectTypeCode
                        + "&records=" + formContext.data.entity.getId()
                        + "&helpID=SD%20Project%20Progress%20Report.rdl";

                    //Set URL of iframe
                    iframeObject.setSrc(strURL);
                }

            } else {
                var errorText = this.responseText;
                //handle error here
                //display error
                Xrm.Navigation.openErrorDialog({ message: errorText });
            }
        }
    };
    req.send();
};