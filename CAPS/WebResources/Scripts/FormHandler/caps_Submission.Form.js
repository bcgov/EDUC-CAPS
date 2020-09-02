"use strict";

var CAPS = CAPS || {};
CAPS.Submission = CAPS.Submission || {};

/**
 * Main function for Capital Plan form, this function calls all other form functions.
 * @param {any} executionContext execution context
 */
CAPS.Submission.onLoad = function (executionContext) {
    var formContext = executionContext.getFormContext();

    //get call for submission type
    if (formContext.getAttribute("caps_callforsubmissiontype").getValue() === 100000002) {
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