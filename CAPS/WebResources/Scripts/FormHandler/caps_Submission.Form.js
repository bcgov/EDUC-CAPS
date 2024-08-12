"use strict";

var CAPS = CAPS || {};
CAPS.Submission = CAPS.Submission || {};

const SUBMISSION_STAUS = {
    DRAFT: 1,
    SUBMITTED: 2,
    CANCELLED: 100000001,
    RESULTS_RELEASED: 200870000,
    ACCEPTED: 200870001,
    COMPLETE: 200870002
};

/**
 * Main function for Capital Plan form, this function calls all other form functions.
 * @param {any} executionContext execution context
 */
CAPS.Submission.onLoad = function (executionContext) {
   
    var formContext = executionContext.getFormContext();

    var callForSubmissionType = formContext.getAttribute("caps_callforsubmissiontype").getValue();
    var selectedForm = formContext.ui.formSelector.getCurrentItem().getLabel(); //Ministry Capital Plan
    var status = formContext.getAttribute("statuscode").getValue();

    formContext.getAttribute("caps_boardresolution").addOnChange(CAPS.Submission.FlagBoardResolutionAsUploaded);

    CAPS.Submission.checkSupportingDocuments(executionContext);

    if (status === SUBMISSION_STAUS.DRAFT || status === SUBMISSION_STAUS.ACCEPTED) {
        //hide report tab
        formContext.ui.tabs.get("tab_capitalplan").setVisible(false);

        if (callForSubmissionType === 100000002) {
            //AFG
            formContext.ui.tabs.get("tab_afg").setVisible(true);
            formContext.ui.tabs.get("tab_CCAFG").setVisible(false);
            formContext.ui.tabs.get("tab_afg").sections.get("sec_afg_projects").setVisible(true);
            formContext.ui.tabs.get("tab_general").setVisible(false);
            formContext.ui.tabs.get("tab_general").sections.get("sec_major_projects").setVisible(false);
            formContext.ui.tabs.get("tab_general").sections.get("sec_minor_projects").setVisible(false);
            formContext.getControl("sgd_AFGProjects").addOnLoad(CAPS.Submission.UpdateTotalAllocated);
            
        }
        //CC-AFG
        else if (callForSubmissionType === 385610001) {
            formContext.ui.tabs.get("tab_CCAFG").setVisible(true);
            formContext.ui.tabs.get("tab_afg").setVisible(false);
            formContext.ui.tabs.get("tab_CCAFG").sections.get("sec_CCAFG_projects").setVisible(true);
            formContext.ui.tabs.get("tab_general").setVisible(false);
            formContext.ui.tabs.get("tab_general").sections.get("sec_major_projects").setVisible(false);
            formContext.ui.tabs.get("tab_general").sections.get("sec_minor_projects").setVisible(false);
            formContext.getControl("sgd_CCAFGProjects").addOnLoad(CAPS.Submission.UpdateTotalAllocated);

        }
        else {
            formContext.ui.tabs.get("tab_afg").setVisible(false);
            formContext.ui.tabs.get("tab_CCAFG").setVisible(false);
            formContext.ui.tabs.get("tab_afg").sections.get("sec_afg_projects").setVisible(false);
            formContext.ui.tabs.get("tab_general").setVisible(true);
            if (callForSubmissionType === 100000000) {
                //Major
                formContext.ui.tabs.get("tab_general").sections.get("sec_major_projects").setVisible(true);
                formContext.ui.tabs.get("tab_general").sections.get("sec_minor_projects").setVisible(false);
            }
            else {
                //Minor
                formContext.ui.tabs.get("tab_general").sections.get("sec_major_projects").setVisible(false);
                formContext.ui.tabs.get("tab_general").sections.get("sec_minor_projects").setVisible(true);
            }

        }
    }
    else if (status === SUBMISSION_STAUS.RESULTS_RELEASED || status === SUBMISSION_STAUS.CANCELLED || status === SUBMISSION_STAUS.COMPLETE) {
        //show report and hide all else
        formContext.ui.tabs.get("tab_capitalplan").setVisible(true);
        CAPS.Submission.embedCapitalPlanReport(executionContext);

        formContext.ui.tabs.get("tab_general").setVisible(false);
        formContext.ui.tabs.get("tab_afg").setVisible(false);
        formContext.ui.tabs.get("tab_CCAFG").setVisible(false);

        if (callForSubmissionType === 100000002) {
            formContext.ui.tabs.get("tab_capitalplan").sections.get("tab_capitalplan_section_boardresolution").setVisible(false);
        }
        if (status === SUBMISSION_STAUS.CANCELLED) {
            //show reason for cancellation
            formContext.getControl("caps_reasonforcancellation").setVisible(true);
        }
    }
    else if (status === SUBMISSION_STAUS.SUBMITTED) {
        if (selectedForm === 'SD Capital Plan') {
            //show report and hide all else
            formContext.ui.tabs.get("tab_capitalplan").setVisible(true);
            CAPS.Submission.embedCapitalPlanReport(executionContext);

            formContext.ui.tabs.get("tab_general").setVisible(false);
            formContext.ui.tabs.get("tab_afg").setVisible(false);
            formContext.ui.tabs.get("tab_CCAFG").setVisible(false);

            if (callForSubmissionType === 100000002) {
                formContext.ui.tabs.get("tab_capitalplan").sections.get("tab_capitalplan_section_boardresolution").setVisible(false);
            }
        }
        else {
            //hide report tab
            formContext.ui.tabs.get("tab_capitalplan").setVisible(false);

            if (callForSubmissionType === 100000002) {
                //AFG
                formContext.ui.tabs.get("tab_afg").setVisible(true);
                formContext.ui.tabs.get("tab_CCAFG").setVisible(false);
                formContext.ui.tabs.get("tab_afg").sections.get("sec_afg_projects").setVisible(true);
                formContext.ui.tabs.get("tab_general").setVisible(false);
                formContext.ui.tabs.get("tab_general").sections.get("sec_major_projects").setVisible(false);
                formContext.ui.tabs.get("tab_general").sections.get("sec_minor_projects").setVisible(false);

                formContext.getControl("sgd_AFGProjects").addOnLoad(CAPS.Submission.UpdateTotalAllocated);
            }
            //CC-AFG
            else if (callForSubmissionType === 385610001) {
                formContext.ui.tabs.get("tab_afg").setVisible(false);
                formContext.ui.tabs.get("tab_CCAFG").setVisible(true);
                formContext.ui.tabs.get("tab_general").setVisible(false);
            }
            else {
                formContext.ui.tabs.get("tab_afg").setVisible(false);
                formContext.ui.tabs.get("tab_CCAFG").setVisible(false);
                formContext.ui.tabs.get("tab_afg").sections.get("sec_afg_projects").setVisible(false);
                formContext.ui.tabs.get("tab_general").setVisible(true);
                if (callForSubmissionType === 100000000) {
                    //Major
                    formContext.ui.tabs.get("tab_general").sections.get("sec_major_projects").setVisible(true);
                    formContext.ui.tabs.get("tab_general").sections.get("sec_minor_projects").setVisible(false);
                }
                else {
                    //Minor
                    formContext.ui.tabs.get("tab_general").sections.get("sec_major_projects").setVisible(false);
                    formContext.ui.tabs.get("tab_general").sections.get("sec_minor_projects").setVisible(true);
                }
            }
        }
    }
};

//onLoad validate the related capital plan supporting documents files//
CAPS.Submission.checkSupportingDocuments = function (executionContext) {
    
    var formContext = executionContext.getFormContext();
    var addSupportingDocuments = CAPS.Submission.GetLookup("caps_supportingdocuments", formContext);
    var submissionType = formContext.getAttribute("caps_submissiontype").getValue();
    //No Expenditure Plans
    if (submissionType !== "200870001") {
        if (addSupportingDocuments !== undefined && formContext.getAttribute("statuscode").getValue() === 200870000) {
            formContext.ui.clearFormNotification("MissingSupportingDocumentsError");
            var addSupportingDocumentsID = CAPS.Submission.RemoveCurlyBraces(addSupportingDocuments.id);
            var options = "?$select=caps_capitalplanresponseletter,caps_capitalbylaw,caps_childcaredeclaration,caps_name";
            Xrm.WebApi.retrieveRecord("caps_capitalplansupportingdocuments", addSupportingDocumentsID, options).then(
                function success(result) {
                    
                    var supportingDocumentName = result.caps_name;
                    var capitalPlanResponseLetter = result.caps_capitalplanresponseletter;
                    var capitalBylaw = result.caps_capitalbylaw;
                    var childCareDeclaration = result.caps_childcaredeclaration;
                    //var capitalPlanSupportingDocumentName = result.caps_name;
                    if (capitalPlanResponseLetter === null || capitalBylaw === null || childCareDeclaration === null) {
                        formContext.ui.setFormNotification("At least one of the supporting documents is missing. Please click the supporting documents record \"" + supportingDocumentName + "\" to upload necessary files in the Submission tab below.", "WARNING", "MissingSupportingDocumentsBanner");
                    }
                    else {
                        formContext.ui.clearFormNotification("MissingSupportingDocumentsBanner");
                    }

                },
                function (error) {
                    console.log(error.message);
                }
            );
        }
        else if (formContext.getAttribute("statuscode").getValue() !== 200870000) { // Do not show the notification if submission status is not Results Released
            formContext.ui.clearFormNotification("MissingSupportingDocumentsBanner");
        }
        else if (addSupportingDocuments === undefined && formContext.getAttribute("statuscode").getValue() === 200870000) {
            formContext.ui.setFormNotification("No related capital plan supporting documents record found.", "ERROR", "MissingSupportingDocumentsError");
        }
    }
};

/**
 * Function that updates the iFrame and embeds the SSRS Capital Plan Report.
 * @param {any} executionContext execution context
 */
CAPS.Submission.embedCapitalPlanReport = function (executionContext) {
    var formContext = executionContext.getFormContext();
   
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
            formContext.data.entity.save();
        },
        function (error) {
            console.log(error.message);
            // handle error conditions
        }
    );

};

/**
Set caps_boardofresolutionattached field to true when file added and false when removed.
*/
CAPS.Submission.FlagBoardResolutionAsUploaded = function (executionContext) {
    var formContext = executionContext.getFormContext();

    var bylaw = formContext.getAttribute("caps_boardresolution").getValue();

    if (bylaw !== null) {
        formContext.getAttribute("caps_boardofresolutionattached").setValue(true);
    }
    else {
        formContext.getAttribute("caps_boardofresolutionattached").setValue(false);
    }
}

CAPS.Submission.GetLookup = function (fieldName, formContext) {
    var lookupFieldObject = formContext.data.entity.attributes.get(fieldName);
    if (lookupFieldObject !== null && lookupFieldObject.getValue() !== null && lookupFieldObject.getValue()[0] !== null) {
        var entityId = lookupFieldObject.getValue()[0].id;
        var entityName = lookupFieldObject.getValue()[0].entityType;
        var entityLabel = lookupFieldObject.getValue()[0].name;
        var obj = {
            id: entityId,
            type: entityName,
            value: entityLabel
        };
        return obj;
    }
}
CAPS.Submission.RemoveCurlyBraces = function (str) {
    return str.replace(/[{}]/g, "");
}
