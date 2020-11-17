"use strict";

var CAPS = CAPS || {};
CAPS.ProgressReport = CAPS.ProgressReport || {};

/*
Main function for Project Tracker.  This function calls all other form functions.
*/
CAPS.ProgressReport.onLoad = function (executionContext) {
    var formContext = executionContext.getFormContext();

    formContext.getControl("sgd_FutureCashFlow").addOnLoad(CAPS.ProgressReport.UpdateTotalFutureCashFlow);
};

/*
Sums up the total projected provincial forecast and projected actuals.
*/
CAPS.ProgressReport.UpdateTotalFutureCashFlow = function (executionContext) {

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