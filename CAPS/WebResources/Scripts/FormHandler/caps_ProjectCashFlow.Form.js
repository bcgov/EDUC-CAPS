"use strict";

var CAPS = CAPS || {};
CAPS.ProjectCashFlow = CAPS.ProjectCashFlow || {};

/**
 * Main function for ProjectCashFlow.  This function calls all other form functions and registers onChange and onLoad events.
 * @param {any} executionContext the form's execution context
 */
CAPS.ProjectCashFlow.onLoad = function(executionContext) {
    var formContext = executionContext.getFormContext();

    //Check if fiscal Year is set, if so check if it's the current year or current year plus 1
    var fiscalYear = formContext.getAttribute("caps_fiscalyear").getValue();

    if (fiscalYear === null || fiscalYear[0] === null) return;

    var currentYear = null;

    Xrm.WebApi.retrieveMultipleRecords("edu_year", "?$select=edu_yearid,edu_startyear,statuscode&$filter=statecode eq 0 and edu_type eq 757500000&$orderby=edu_startyear").then(
        function success(result) {
            //find current year record and current year +1

            for (var x = 0; x < result.entities.length; x++) {
                if (result.entities[x].statuscode === 1) {
                    //this is the current year
                    currentYear = result.entities[x].edu_yearid;
                    break;
                }
            }

            CAPS.ProjectCashFlow.showHideValuesBasedOnYear(formContext, currentYear);

        },
        function (error) {

            console.log(error.message);
            // handle error conditions
        });
}

/**
 * Display function for project cash flow, this method shows Quarterly fields for provincial (forecast) for current year.  Shows actual draws by quarter and remaining draws for current year only.
 * Shows agency by quarter for current year only, otherwise shows only total.
 * @param {any} formContext the form's context
 * @param {any} currentYear the current fiscal year
 */
CAPS.ProjectCashFlow.showHideValuesBasedOnYear = function(formContext, currentYear) {
    debugger;
    //Check if fiscal Year is set, if so check if it's the current year or current year plus 1
    var fiscalYear = formContext.getAttribute("caps_fiscalyear").getValue();

    var fiscalYearValue = fiscalYear[0].id.toLowerCase().replace("{", "").replace("}", "");

    var stateCode = formContext.getAttribute("statecode").getValue();
    
    if (fiscalYearValue === currentYear) {

            formContext.getControl("caps_totalactualdraws").setVisible(true);

            formContext.getControl("caps_remainingdraws").setVisible(true);
            //Show Actual Draws Chart
            formContext.ui.tabs.get("tab_general").sections.get("sec_actual_draws").setVisible(true);
    }
    else if (stateCode === 1) {

        formContext.getControl("caps_totalactualdraws").setVisible(true);

        formContext.getControl("caps_remainingdraws").setVisible(true);

    }
    else {

        formContext.getControl("caps_totalactualdraws").setDisabled(false);
        formContext.getControl("caps_totalprovincial").setDisabled(false);
        formContext.getControl("caps_totalagency").setDisabled(false);
    }


}

/**
 * Calculate the yearly totals for Provincial (Forecast) and Agency
 * @param {any} executionContext the form's context
 */
CAPS.ProjectCashFlow.calculateYearly = function(executionContext) {
    var formContext = executionContext.getFormContext();

    var totalProvincial = formContext.getAttribute("caps_q1provincial").getValue() + formContext.getAttribute("caps_q2provincial").getValue() + formContext.getAttribute("caps_q3provincial").getValue() + formContext.getAttribute("caps_q4provincial").getValue();

    var totalAgency = formContext.getAttribute("caps_q1agency").getValue() + formContext.getAttribute("caps_q2agency").getValue() + formContext.getAttribute("caps_q3agency").getValue() + formContext.getAttribute("caps_q4agency").getValue();

    formContext.getAttribute("caps_totalprovincial").setValue(totalProvincial);
    formContext.getAttribute("caps_totalagency").setValue(totalAgency);
}
