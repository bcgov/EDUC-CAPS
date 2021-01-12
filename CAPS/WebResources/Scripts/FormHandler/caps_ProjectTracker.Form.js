"use strict";

var CAPS = CAPS || {};
CAPS.ProjectTracker = CAPS.ProjectTracker || {};

const MAX_BUDGET_NOTIFICATION = "Max_Budget_Notification";
const PROVINCIAL_BUDGET_NOTIFICATION = "Provincial_Budget_Notification";
const STALE_DATES_NOTIFICATION = "Stale_Dates_Notification";

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

    CAPS.ProjectTracker.ShowBudgetMissmatch(executionContext);

    CAPS.ProjectTracker.ShowStaleDates(executionContext);

    formContext.getAttribute("caps_maxpotential_fundingsource").addOnChange(CAPS.ProjectTracker.ShowBudgetMissmatch);
    formContext.getAttribute("caps_maxpotentialprojectbudget").addOnChange(CAPS.ProjectTracker.ShowBudgetMissmatch);
    formContext.getAttribute("caps_provincial").addOnChange(CAPS.ProjectTracker.ShowBudgetMissmatch);
    formContext.getAttribute("caps_totalprovincialbudget").addOnChange(CAPS.ProjectTracker.ShowBudgetMissmatch);
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
        formContext.ui.tabs.get("tab_general").sections.get("section_designcapacity").setVisible(false);

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

        formContext.ui.tabs.get("tab_general").sections.get("section_designcapacity").setVisible(false);

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

/**
Shows a warning if the Max Potential Budget and Max Potential don't match or if the Provincial Funding Source or Total Provincial Budget don't match.
*/
CAPS.ProjectTracker.ShowBudgetMissmatch = function (executionContext) {
    var formContext = executionContext.getFormContext();

    if (formContext.getAttribute("caps_maxpotential_fundingsource").getValue() !== formContext.getAttribute("caps_maxpotentialprojectbudget").getValue()) {
        formContext.ui.setFormNotification('Max Potential Project Budget and Max Potential Funding Source don\'t match.', 'WARNING', MAX_BUDGET_NOTIFICATION);
    }
    else {
        formContext.ui.clearFormNotification(MAX_BUDGET_NOTIFICATION);
    }

    if (formContext.getAttribute("caps_provincial").getValue() !== formContext.getAttribute("caps_totalprovincialbudget").getValue()) {
        formContext.ui.setFormNotification('Provincial Funding Source and Total Provincial Budget don\'t match.', 'WARNING', PROVINCIAL_BUDGET_NOTIFICATION);
    }
    else {
        formContext.ui.clearFormNotification(PROVINCIAL_BUDGET_NOTIFICATION);
    }
};

/**
Shows a warning if any Milestones dates are in the past but are not flagged as complete.
**/
CAPS.ProjectTracker.ShowStaleDates = function (executionContext) {
    var formContext = executionContext.getFormContext();

    var recordId = formContext.data.entity.getId();

    var fetchXML = "<fetch version=\"1.0\" output-format=\"xml-platform\" mapping=\"logical\" distinct=\"false\">"+
              "<entity name=\"caps_projectmilestone\">"+
                "<attribute name=\"caps_projectmilestoneid\" />"+
                "<attribute name=\"caps_name\" />"+
                "<attribute name=\"createdon\" />"+
                "<order attribute=\"caps_name\" descending=\"false\" />"+
                "<filter type=\"and\">"+
                  "<condition attribute=\"caps_complete\" operator=\"eq\" value=\"0\" />"+
                  "<condition attribute=\"caps_expectedactualdate\" operator=\"olderthan-x-days\" value=\"1\" />" +
                  "<condition attribute=\"caps_projecttracker\" operator=\"eq\"  value=\""+recordId+"\" />"+
                "</filter>"+
              "</entity>"+
            "</fetch>";


    Xrm.WebApi.retrieveMultipleRecords("caps_projectmilestone", "?fetchXml=" + fetchXML).then(
                  function success(result) {
                      if (result.entities.length > 0) {
                          formContext.ui.setFormNotification('One or more project milestone dates requires updating.', 'WARNING', STALE_DATES_NOTIFICATION);
                    
                      }
                      else {
                          formContext.ui.clearFormNotification(STALE_DATES_NOTIFICATION);
                      }
                  },
                  function (error) {
                      Xrm.Navigation.openErrorDialog({ message: error.message });
                  }
                  );

};

