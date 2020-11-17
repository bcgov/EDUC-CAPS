"use strict";

var CAPS = CAPS || {};
CAPS.Submission = CAPS.Submission || {};

const SUBMISSION_STAUS = {
    DRAFT: 1,
    SUBMITTED: 2,
    CANCELLED: 100000001,
    RESULTS_RELEASED: 200870000,
};

/**
 * Main function for Capital Plan form, this function calls all other form functions.
 * @param {any} executionContext execution context
 */
CAPS.Submission.onLoad = function (executionContext) {
    debugger;
    var formContext = executionContext.getFormContext();

    var callForSubmissionType = formContext.getAttribute("caps_callforsubmissiontype").getValue();
    var selectedForm = formContext.ui.formSelector.getCurrentItem().getLabel(); //Ministry Capital Plan
    var status = formContext.getAttribute("statuscode").getValue();
    
    if (status === SUBMISSION_STAUS.DRAFT) {
        //hide report tab
        formContext.ui.tabs.get("tab_capitalplan").setVisible(false);

        if (callForSubmissionType === 100000002) {
            //AFG
            formContext.ui.tabs.get("tab_afg").setVisible(true);
            formContext.ui.tabs.get("tab_afg").sections.get("sec_afg_projects").setVisible(true);
            formContext.ui.tabs.get("tab_general").setVisible(false);
            formContext.ui.tabs.get("tab_general").sections.get("sec_major_minor_projects").setVisible(false);

            formContext.getControl("sgd_AFGProjects").addOnLoad(CAPS.Submission.UpdateTotalAllocated);
        }
        else {
            formContext.ui.tabs.get("tab_afg").setVisible(false);
            formContext.ui.tabs.get("tab_afg").sections.get("sec_afg_projects").setVisible(false);
            formContext.ui.tabs.get("tab_general").setVisible(true);
            formContext.ui.tabs.get("tab_general").sections.get("sec_major_minor_projects").setVisible(true);
        }
    }
    else if (status === SUBMISSION_STAUS.RESULTS_RELEASED || status === SUBMISSION_STAUS.CANCELLED) {
        //show report and hide all else
        formContext.ui.tabs.get("tab_capitalplan").setVisible(true);
        CAPS.Submission.embedCapitalPlanReport(executionContext);
        formContext.ui.tabs.get("tab_general").setVisible(false);
        formContext.ui.tabs.get("tab_afg").setVisible(false);
    }
    else if (status === SUBMISSION_STAUS.SUBMITTED) {
        if (selectedForm === 'SD Capital Plan') {
            //show report and hide all else
            formContext.ui.tabs.get("tab_capitalplan").setVisible(true);
            CAPS.Submission.embedCapitalPlanReport(executionContext);
            formContext.ui.tabs.get("tab_general").setVisible(false);
            formContext.ui.tabs.get("tab_afg").setVisible(false);
        }
        else {
            //hide report tab
            formContext.ui.tabs.get("tab_capitalplan").setVisible(false);

            if (callForSubmissionType === 100000002) {
                //AFG
                formContext.ui.tabs.get("tab_afg").setVisible(true);
                formContext.ui.tabs.get("tab_afg").sections.get("sec_afg_projects").setVisible(true);
                formContext.ui.tabs.get("tab_general").setVisible(false);
                formContext.ui.tabs.get("tab_general").sections.get("sec_major_minor_projects").setVisible(false);

                formContext.getControl("sgd_AFGProjects").addOnLoad(CAPS.Submission.UpdateTotalAllocated);
            }
            else {
                formContext.ui.tabs.get("tab_afg").setVisible(false);
                formContext.ui.tabs.get("tab_afg").sections.get("sec_afg_projects").setVisible(false);
                formContext.ui.tabs.get("tab_general").setVisible(true);
                formContext.ui.tabs.get("tab_general").sections.get("sec_major_minor_projects").setVisible(true);
            }
        }
    }
};

/**
 * Function that updates the iFrame and embeds the SSRS Capital Plan Report.
 * @param {any} executionContext execution context
 */
CAPS.Submission.embedCapitalPlanReport = function (executionContext) {
    var formContext = executionContext.getFormContext();
    debugger;
    var requestUrl = "/api/data/v9.1/EntityDefinitions?$filter=LogicalName eq 'caps_submission'&$select=ObjectTypeCode";

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
                var iframeObject = formContext.getControl("IFRAME_CapitalPlanReport");

                if (iframeObject !== null) {
                    var strURL = "/crmreports/viewer/viewer.aspx"
                        + "?id=a3365e88-c014-eb11-a813-000d3af43595"
                        + "&action=run"
                        + "&context=records"
                        + "&recordstype=" + objectTypeCode
                        + "&records=" + formContext.data.entity.getId()
                        + "&helpID=CapitalPlanSubmissionReport.rdl";

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
 * Triggered when AFG PCF is updated, this function updates the AFG total and the variance.
 * @param {any} executionContext execution context
 */
CAPS.Submission.UpdateTotalAllocated = function (executionContext) {
    debugger;
    var formContext = executionContext.getFormContext();
    var id = formContext.data.entity.getId().replace("{", "").replace("}", "");
    Xrm.WebApi.retrieveMultipleRecords("caps_project", "?$select=caps_totalprojectcost&$filter=caps_Submission/caps_submissionid eq " + id).then(
        function success(result) {
            var totalAllocated = 0;
            for (var i = 0; i < result.entities.length; i++) {
                totalAllocated += result.entities[i].caps_totalprojectcost;
            }

            // perform operations on record retrieval
            formContext.getAttribute('caps_totalestimatedcost').setValue(totalAllocated);
            //calculate variance
            var totalCost = formContext.getAttribute('caps_totalallocationtodistrict').getValue();
            var variance = null;
            if (totalCost !== null && totalAllocated !== null) {
                variance = totalCost - totalAllocated;
                formContext.getAttribute('caps_variance').setValue(variance);
            }
        },
        function (error) {
            console.log(error.message);
            // handle error conditions
        }
    );

};