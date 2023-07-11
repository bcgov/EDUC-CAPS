"use strict";

var CAPS = CAPS || {};
CAPS.CallForSubmission = CAPS.CallForSubmission || {
    START_YEAR: null,
    END_YEAR: null
};

/**
 * Main function for Call for Submission.  This function calls all other form functions and registers onChange and onLoad events.
 * @param {any} executionContext the form's execution context
 */
CAPS.CallForSubmission.onLoad = async function (executionContext) {

    var formContext = executionContext.getFormContext();

    //Get current Year
    var currentYearRecord = await Xrm.WebApi.retrieveMultipleRecords("edu_year", "?$select=edu_yearid,edu_startyear,statuscode&$filter=edu_type eq 757500000 and statuscode eq 1");

    if (currentYearRecord.entities.length === 1) {
        //Good to go
        CAPS.CallForSubmission.START_YEAR = currentYearRecord.entities[0].edu_startyear;
        CAPS.CallForSubmission.END_YEAR = CAPS.CallForSubmission.START_YEAR + 2;

        CAPS.CallForSubmission.setFiscalYearFilter(executionContext);
    }
    // CAPS-1977 Hide PO Portfolio Ranking Tab from SD User View
    CAPS.CallForSubmission.ShowPOPortfolioRankingTab(executionContext);
}

/**
 * Function to add a filter to capital plan year.
 * @param {any} executionContext the form's execution context
 */
CAPS.CallForSubmission.setFiscalYearFilter = function (executionContext) {
    var formContext = executionContext.getFormContext();
    formContext.getControl("caps_capitalplanyear").addPreSearch(CAPS.CallForSubmission.filterFiscalYear);
}

/**
 * Filtering function for Fiscal Year.  This function limits the results to current year, current year + 1 and current year + 2.
 * @param {any} executionContext the form's execution context
 */
CAPS.CallForSubmission.filterFiscalYear = function (executionContext) {

    var formContext = executionContext.getFormContext();
    //Then add 2 to the starting year value and call addCustomFilter
    var fetchXML = "<filter type=\"and\">" +
        "<condition attribute=\"edu_startyear\" operator=\"ge\" value=\"" + CAPS.CallForSubmission.START_YEAR + "\" />" +
        "<condition attribute=\"edu_startyear\" operator=\"le\" value=\"" + CAPS.CallForSubmission.END_YEAR + "\" />" +
        "<condition attribute=\"statecode\" operator=\"eq\" value=\"0\" />" +
        "</filter>";

    formContext.getControl("caps_capitalplanyear").addCustomFilter(fetchXML, "edu_year");
}

/**
 * CAPS-1977 Hide PO Portfolio Ranking Tab from SD User View
 * Function to determine when the PO Portfolio Ranking Tab should be shown.
 * @param {any} executionContext the form's execution context
 */
CAPS.CallForSubmission.ShowPOPortfolioRankingTab = function (executionContext) {
    var formContext = executionContext.getFormContext();

    var userRoles = Xrm.Utility.getGlobalContext().userSettings.roles;
    var showTab = false;

    userRoles.forEach(function hasFinancialDirectorRole(item, index) {
        if (item.name === "CAPS CMB User") {
            showTab = true;
        }
    });

    // Get Tab
    var theTab = formContext.ui.tabs.get("tab_3");
    if (theTab != null) {
        theTab.setVisible(showTab);
    }
}
