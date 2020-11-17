"use strict";

/* INCLUDE CAPS.Common.js */

var CAPS = CAPS || {};
CAPS.ProjectTracker = CAPS.ProjectTracker || {};

/**
 * Main function for Project Tracker.  This function calls all other form functions.
 * @param {any} executionContext the form's execution context.
 */
CAPS.ProjectTracker.onLoad = function (executionContext) {
    var formContext = executionContext.getFormContext();

    //Show Summary Report
    CAPS.ProjectTracker.showSummaryReport(formContext);

    //Format the form by showing and hiding the relevant sections and fields
    CAPS.ProjectTracker.showHideCategoryRelevantSections(formContext);

    CAPS.ProjectTracker.ShowHideProgressReport(executionContext);
}

/**
 * Function to embed the SSRS Monthly Summary Report on the form.
 * @param {any} formContext the form context
 */
CAPS.ProjectTracker.showSummaryReport = function (formContext) {    
    debugger;
    var requestUrl = "/api/data/v9.1/EntityDefinitions?$filter=LogicalName eq 'caps_projecttracker'&$select=ObjectTypeCode";

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
                        + "?id=ed1744bc-98a6-ea11-a813-000d3af42496"
                        + "&action=run"
                        + "&context=records"
                        + "&recordstype="+objectTypeCode
                        + "&records=" + formContext.data.entity.getId()
                        + "&helpID=Monthly%20Project%20Summary.rdl";

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


/**
 * Shows and hides appropriate sections and fields for Major/Minor and AFG Projects.
 * @param {any} formContext the form context
 */
CAPS.ProjectTracker.showHideCategoryRelevantSections = function (formContext) {
    debugger;
    var submissionCategoryCode = formContext.getAttribute("caps_submissioncategorycode").getValue();

    if (submissionCategoryCode === 'AFG') {
        //hide procurement, hide other funding, hide both date sections, hide agency section
        formContext.getControl("caps_procurementmethod").setVisible(false);
        if (formContext.getControl("caps_totalagencyprojected") !== null) {
            formContext.getControl("caps_totalagencyprojected").setVisible(false);
        }
        if (formContext.getControl("caps_variancefromagencybudgeted") !== null) {
            formContext.getControl("caps_variancefromagencybudgeted").setVisible(false);
        }

        formContext.ui.tabs.get("tab_general").sections.get("sec_other_funding").setVisible(false);
        formContext.ui.tabs.get("tab_general").sections.get("sec_agreement_dates").setVisible(false);
        formContext.ui.tabs.get("tab_general").sections.get("sec_ministry_dates").setVisible(false);
        formContext.ui.tabs.get("tab_general").sections.get("sec_budget_breakdown").setVisible(false);
        formContext.ui.tabs.get("tab_general").sections.get("sec_provincial_budget").setVisible(false);

        if (formContext.ui.tabs.get("tab_general").sections.get("sec_cps") !== null) {
            formContext.ui.tabs.get("tab_general").sections.get("sec_cps").setVisible(false);
        }

        formContext.ui.tabs.get("tab_general").sections.get("sec_afg_budget").setVisible(true);

    }
    else if (submissionCategoryCode === 'BUS' || submissionCategoryCode === 'SEP' || submissionCategoryCode === 'PEP' || submissionCategoryCode === 'CNCP') {
        //hide other funding, provincial funding > hide all but total provincial field
        formContext.ui.tabs.get("tab_general").sections.get("sec_other_funding").setVisible(false);
        formContext.ui.tabs.get("tab_general").sections.get("sec_budget_breakdown").setVisible(false);


        formContext.getControl("caps_approvedreserve").setVisible(false);
        formContext.getControl("caps_totalapproved").setVisible(false);
        formContext.getControl("caps_unapprovedreserve").setVisible(false);
        formContext.getControl("caps_totalprovincialbudget").setVisible(false);
        
        if (formContext.getControl("caps_totalagencyprojected") !== null) {
            formContext.getControl("caps_totalagencyprojected").setVisible(false);
        }
        if (formContext.getControl("caps_variancefromagencybudgeted") !== null) {
            formContext.getControl("caps_variancefromagencybudgeted").setVisible(false);
        }

    }

}

/**
Shows the progress report for major projects, otherwise hides the tab.
*/
CAPS.ProjectTracker.ShowHideProgressReport = function (executionContext) {
    var formContext = executionContext.getFormContext();

    if (formContext.getAttribute("caps_showprogressreports").getValue()) {
        formContext.ui.tabs.get("tab_progressreports").setVisible(true);
    }
    else {
        formContext.ui.tabs.get("tab_progressreports").setVisible(false);
    }
};

