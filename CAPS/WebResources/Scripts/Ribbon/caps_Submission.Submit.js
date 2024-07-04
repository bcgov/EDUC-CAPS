"use strict";

var CAPS = CAPS || {};
CAPS.Submission = CAPS.Submission || {};

/***
Function to determine when the Complete button should be shown.
*/
CAPS.Submission.ShowComplete = function (primaryControl) {

    var formContext = primaryControl;

    //If not in status of accepted
    if (formContext.getAttribute("statuscode").getValue() !== 200870001) {
        return false;
    }
    //200,870,001
    var userRoles = Xrm.Utility.getGlobalContext().userSettings.roles;

    var showValidation = false;

    userRoles.forEach(function hasFinancialDirectorRole(item, index) {
        if (item.name === "CAPS School District Approver - Add On") {
            showValidation = true;
        }
    });

    return showValidation;
};

/*  For AFG Only.  This function validates the submission and completes it if it's valid.
*/
CAPS.Submission.Complete = async function (primaryControl) {
    
    var formContext = primaryControl;

    //Declare variables
    var resultsArray = [];

    Xrm.Utility.showProgressIndicator("Validating Submission...");

    var resultsMessage = "Validation Errors\r\n";

    let checkProjectRequestResults = await CAPS.Submission.CheckProjectRequests(formContext);
    let checkAFGVariance = await CAPS.Submission.CheckAFGTotal(formContext);

    resultsArray.push(checkProjectRequestResults);
    resultsArray.push(checkAFGVariance);

    var allValidationPassed = true;
    //loop through results
    resultsArray.forEach(function (item, index) {
        if (typeof item !== 'undefined' && item !== null) {
            if (!item.validationResult) {
                allValidationPassed = false;
                resultsMessage += item.validationMessage + "\r\n";
            }
        }
    });

    //Show Results
    formContext.ui.tabs.forEach(function (tab, i) {
        //loop through sections
        tab.sections.forEach(function (section, j) {
            section.controls.forEach(function (control, k) {

                if (control.getAttribute().getName() === "caps_validationresults") {
                    control.getAttribute().setValue(resultsMessage);
                    control.setVisible(true);
                }
            });
        });
    });

    //sgd_AFGProjects
    formContext.getControl("sgd_AFGProjects").refresh();

    if (allValidationPassed) {
        let confirmStrings = { text: "This is your year end expenditure plan. By completing your expenditure plan, you are reporting your AFG actual spending for the year. Press OK to proceed with completing your expenditure plan or Cancel to exit.", title: "Submission Confirmation" };
        let confirmOptions = { height: 200, width: 450 };
        Xrm.Navigation.openConfirmDialog(confirmStrings, confirmOptions).then(
            function (success) {
                if (success.confirmed) {
                    formContext.getAttribute("statecode").setValue(1);
                    formContext.getAttribute("statuscode").setValue(200870002);
                    formContext.data.entity.save();
                }
            });
    }
    else {
        var alertStrings = { confirmButtonLabel: "Ok", text: "The validation of this submission failed.  Please see the Validation results section below for details.", title: "Validation Result" };
        var alertOptions = { height: 120, width: 260 };
        Xrm.Navigation.openAlertDialog(alertStrings, alertOptions);
    }

    //close the indicator
    Xrm.Utility.closeProgressIndicator();

};

/**
 * Function to determine when the Submit (and Validate) button should be shown.
 * @param {any} primaryControl primary control
 * @returns {boolean} true if should be shown, otherwise false.
 */
CAPS.Submission.ShowValidate = function (primaryControl) {

    var formContext = primaryControl;

    //If not in status of draft
    if (formContext.getAttribute("statuscode").getValue() !== 1) {
        return false;
    }

    var userRoles = Xrm.Utility.getGlobalContext().userSettings.roles;

    var showValidation = false;

    userRoles.forEach(function hasFinancialDirectorRole(item, index) {
        if (item.name === "CAPS School District Approver - Add On") {
            showValidation = true;
        }
    });

    return showValidation;
};

/**
 * Called from Capital Plan (caps_submission) form.  This function validates the capital plan, displays any errors and submits the plan if there are no validation errors.
 * @param {any} primaryControl primary control
 */
CAPS.Submission.Validate = async function (primaryControl) {
    
    var formContext = primaryControl;

    //Declare variables
    var resultsArray = [];

    Xrm.Utility.showProgressIndicator("Validating Submission...");

    var resultsMessage = "Validation Errors\r\n";


    let checkProjectRequestResults = await CAPS.Submission.CheckProjectRequests(formContext);
    let checkProjectCountResults = await CAPS.Submission.CheckNothingToSubmit(formContext);
    let checkAFGVariance = await CAPS.Submission.CheckAFGTotal(formContext);
    let checkAFGPreviousSubmission = await CAPS.Submission.CheckAFGProjects(formContext);

    // resultsArray.push(checkEnrolmentResult);
    resultsArray.push(checkProjectRequestResults);
    resultsArray.push(checkAFGVariance);
    resultsArray.push(checkProjectCountResults);
    resultsArray.push(checkAFGPreviousSubmission);

    var allValidationPassed = true;
    //loop through results
    resultsArray.forEach(function (item, index) {
        if (typeof item !== 'undefined' && item !== null) {
            if (!item.validationResult) {
                allValidationPassed = false;
                resultsMessage += item.validationMessage + "\r\n";
            }
        }
    });

    //Check if board resolution attached
    let checkBoardResolutionResults = await CAPS.Submission.CheckBoardResolution(formContext);

    let checkEnrolmentResult = await CAPS.Submission.CheckEnrolmentProjections(formContext);

    //Show Results
    formContext.ui.tabs.forEach(function (tab, i) {
        //loop through sections
        tab.sections.forEach(function (section, j) {
            section.controls.forEach(function (control, k) {

                if (control.getAttribute().getName() === "caps_validationresults") {
                    if (allValidationPassed) {
                        control.getAttribute().setValue(null);
                        control.setVisible(false);
                    }
                    else {
                        control.getAttribute().setValue(resultsMessage);
                        control.setVisible(true);
                    }
                }
            });
        });
    });

    //refresh the project subgrids
    formContext.getControl("sgd_Major_Projects").refresh();
    formContext.getControl("sgd_Minor_Projects").refresh();

    //sgd_AFGProjects
    formContext.getControl("sgd_AFGProjects").refresh();


    if (allValidationPassed) {
        
        //all good so change status or display confirmation
        if ((checkBoardResolutionResults != null && !checkBoardResolutionResults.validationResult)
            || (checkEnrolmentResult != null && !checkEnrolmentResult.validationResult)) {
            let warningText = "WARNING: MISSING INFORMATION";
            if (!checkBoardResolutionResults.validationResult) {
                warningText = warningText + "\n\r⠀\n\rNo Board Resolution was attached. This will need to be submitted in order for the Ministry to consider approving projects on the capital plan submission.";
            }
            if (!checkEnrolmentResult.validationResult) {
                warningText = warningText + "\n\r⠀\n\rEnrolment Projections are not complete. These will need to be submitted in order for the Ministry to review and process the capital plan submission.";
            }

            warningText = warningText + "\n\r⠀\n\rClick OK to submit or Cancel to exit.";
            let confirmStrings = { text: warningText, title: "Submission Confirmation" };
            let confirmOptions = { height: 350, width: 550 };
            Xrm.Navigation.openConfirmDialog(confirmStrings, confirmOptions).then(
                function (success) {
                    
                    if (success.confirmed) {
                        formContext.getAttribute("statecode").setValue(1);
                        formContext.getAttribute("statuscode").setValue(2);
                        formContext.data.entity.save();
                    }

                });
        }
        else {
            let confirmStrings = { text: "This submission is valid and ready for submission.  Click OK to Submit or Cancel to exit.", title: "Submission Confirmation" };
            let confirmOptions = { height: 200, width: 450 };
            Xrm.Navigation.openConfirmDialog(confirmStrings, confirmOptions).then(
                function (success) {
                    
                    if (success.confirmed) {
                        formContext.getAttribute("statecode").setValue(1);
                        formContext.getAttribute("statuscode").setValue(2);
                        formContext.data.entity.save();
                    }

                });
        }
    }
    else {
        var alertStrings = { confirmButtonLabel: "Ok", text: "The validation of this submission failed.  Please see the Validation results section below for details.", title: "Validation Result" };
        var alertOptions = { height: 120, width: 260 };
        Xrm.Navigation.openAlertDialog(alertStrings, alertOptions);
    }

    //close the indicator
    Xrm.Utility.closeProgressIndicator();


}

/**
 * This is the validation object all validation methods will return
 * @param {any} validationResult boolean return True if validation passes, or false if it doesn't
 * @param {any} validationMessage string detailed message about validation
 */
CAPS.Submission.ValidationResult = function (validationResult, validationMessage) {
    this.validationResult = validationResult;
    this.validationMessage = validationMessage;
}

/**
 * Called from CAPS.Submission.Validate, this function checks all enrolment projections for the selected school district.
 * @param {any} formContext formContext
 * @returns {CAPS.Submission.ValidationResult} object that contains the validation results
 */
CAPS.Submission.CheckEnrolmentProjections = function (formContext) {

    //Get flag to identify if enrolment validation is required
    var doEnrolmentValidation = formContext.getAttribute("caps_validateenrolment").getValue();

    if (!doEnrolmentValidation) return new CAPS.Submission.ValidationResult(true, "Enrolment Validation: Not Required");

    //Get the School District
    var schoolDistrict = formContext.getAttribute("caps_schooldistrict").getValue();
    var schoolDistrictId = schoolDistrict[0].id;


    var fetchXML = "<fetch version=\"1.0\" output-format=\"xml-platform\" mapping=\"logical\" distinct=\"false\">" +
        "<entity name=\"caps_facility\">" +
        "<attribute name=\"caps_facilityid\" />" +
        "<attribute name=\"caps_name\" />" +
        "<attribute name=\"createdon\" />" +
        "<order attribute=\"caps_name\" descending=\"false\" />" +
        "<filter type=\"and\">" +
        "<condition attribute=\"caps_outstandingenrolmentprojection\" operator=\"ne\" value=\"1\" />" +
        "<condition attribute=\"statecode\" operator=\"eq\" value=\"0\" />" +
        "<condition attribute=\"caps_schooldistrict\" operator=\"eq\" value=\"" + schoolDistrictId + "\" /> " +
        "</filter>" +
        "<link-entity name=\"caps_facilitytype\" from=\"caps_facilitytypeid\" to=\"caps_currentfacilitytype\" link-type=\"inner\" alias=\"ab\">" +
        "<filter type=\"and\">" +
        "<condition attribute=\"caps_schooltype\" operator=\"not-null\" />" +
        "</filter>" +
        "</link-entity>" +
        "</entity>" +
        "</fetch>";

    //Get all Facilities on the school district
    return Xrm.WebApi.retrieveMultipleRecords("caps_facility", "?fetchXml=" + fetchXML).then(
        function success(result) {
            if (result.entities.length > 0) {
                //Some enrolment projections aren't finished
                return new CAPS.Submission.ValidationResult(false, "Enrolment Validation: There are missing enrolment projections");
            }
            return new CAPS.Submission.ValidationResult(true, "Enrolment Validation: Success");
        },
        function (error) {
            alert(error.message);
        }
    );
}

/**
 * Called from CAPS.Submission.Validate, this function checks all BEP projects to ensure they are only for eligible facilities
 * @param {any} formContext the form Context
 * @returns  {CAPS.Submission.ValidationResult} object that contains the validation results
 */
CAPS.Submission.CheckBEPProjects = function (formContext) {
    //Get record ID
    var submissionId = formContext.data.entity.getId();

    //Check if any projects are BEP Projects, if they are check that their facility is tagged as BEP Eligible
    var fetchXML = "<fetch version=\"1.0\" output-format=\"xml - platform\" mapping=\"logical\" distinct=\"false\">" +
        "<entity name = \"caps_project\" >" +
        "<attribute name=\"caps_projectid\" />" +
        "<attribute name=\"caps_projectcode\" />" +
        "<order attribute=\"caps_projectcode\" descending=\"false\" />" +
        "<filter type=\"and\">" +
        "<condition attribute=\"caps_submission\" operator=\"eq\" value=\"" + submissionId + "\" />" +
        "</filter>" +

        "<link-entity name=\"caps_facility\" from=\"caps_facilityid\" to=\"caps_facility\" link-type=\"inner\" alias=\"ab\">" +
        "<filter type=\"and\">" +
        "<condition attribute=\"caps_bepeligible\" operator=\"ne\" value=\"1\" />" +
        "</filter>" +
        "</link-entity>" +

        "<link-entity name=\"caps_submissioncategory\" from=\"caps_submissioncategoryid\" to=\"caps_submissioncategory\" link-type=\"inner\" alias=\"ad\">" +
        "<filter type=\"and\">" +
        "<condition attribute=\"caps_categorycode\" operator=\"eq\" value=\"BEP\" />" +
        "</filter>" +
        "</link-entity>" +

        "</entity>" +
        "</fetch >";

    //Get all BEP Projects who's facilities aren't eligible
    return Xrm.WebApi.retrieveMultipleRecords("caps_project", "?fetchXml=" + fetchXML).then(
        function success(result) {

            if (result.entities.length > 0) {
                //Some bad projects
                return new CAPS.Submission.ValidationResult(false, "BEP Validation: One or more BEP projects is for a facility that is no longer eligible.");
            }
            return new CAPS.Submission.ValidationResult(true, "BEP Validation: Success");
        },
        function (error) {
            alert(error.message);
        }
    );
}

/**
 * Called from CAPS.Submission.Validate, this function checks all BUS projects to ensure they are only for eligible buses
 * @param {any} formContext the form Context
 * @returns  {CAPS.Submission.ValidationResult} object that contains the validation results
 */
CAPS.Submission.CheckBUSProjects = function (formContext) {
    //Get record ID
    var submissionId = formContext.data.entity.getId();

    //Check if any projects are BEP Projects, if they are check that their facility is tagged as BEP Eligible
    var fetchXML = "<fetch version=\"1.0\" output-format=\"xml - platform\" mapping=\"logical\" distinct=\"false\">" +
        "<entity name = \"caps_project\" >" +
        "<attribute name=\"caps_projectid\" />" +
        "<attribute name=\"caps_projectcode\" />" +
        "<order attribute=\"caps_projectcode\" descending=\"false\" />" +
        "<filter type=\"and\">" +
        "<condition attribute=\"caps_submission\" operator=\"eq\" value=\"" + submissionId + "\" />" +
        "</filter>" +

        "<link-entity name=\"caps_bus\" from=\"caps_busid\" to=\"caps_bus\" link-type=\"inner\" alias=\"ab\">" +
        "<filter type=\"and\">" +
        "<condition attribute=\"caps_nonreplaceable\" operator=\"eq\" value=\"0\" />" +
        "</filter>" +
        "</link-entity>" +

        "<link-entity name=\"caps_submissioncategory\" from=\"caps_submissioncategoryid\" to=\"caps_submissioncategory\" link-type=\"inner\" alias=\"ad\">" +
        "<filter type=\"and\">" +
        "<condition attribute=\"caps_categorycode\" operator=\"eq\" value=\"BUS\" />" +
        "</filter>" +
        "</link-entity>" +

        "</entity>" +
        "</fetch >";

    //Get all BEP Projects who's facilities aren't eligible
    return Xrm.WebApi.retrieveMultipleRecords("caps_project", "?fetchXml=" + fetchXML).then(
        function success(result) {

            if (result.entities.length > 0) {
                //Some bad projects
                return new CAPS.Submission.ValidationResult(false, "Bus Validation: One or more BUS projects is a replacement for a bus that is not eligible for replacement.");
            }
            return new CAPS.Submission.ValidationResult(true, "Bus Validation: Success");
        },
        function (error) {
            alert(error.message);
        }
    );
}

/**
 * Called from CAPS.Submission.Validate, this function calls an action that validates all project requests in the submission.
 * @param {any} formContext form context
 * @returns  {CAPS.Submission.ValidationResult} object that contains the validation results
 */
CAPS.Submission.CheckProjectRequests = function (formContext) {
    //Call action on all Project Requests to check they are valid
    //Get record ID
    var recordId = formContext.data.entity.getId().replace("{", "").replace("}", "");
    //call action
    var req = {};
    var target = { entityType: "caps_submission", id: recordId };
    req.entity = target;

    req.getMetadata = function () {
        return {
            boundParameter: "entity",
            operationType: 0,
            operationName: "caps_ValidateCapitalPlan",
            parameterTypes: {
                "entity": {
                    "typeName": "mscrm.caps_submission",
                    "structuralProperty": 5
                }
            }
        };
    };

    return Xrm.WebApi.online.execute(req).then(
        function (result) {

            if (result.ok) {
                return result.json().then(
                    function (response) {
                        //return response
                        var responseMessage = "Project Request Validation Succeeded.";
                        if (!response.ValidationStatus) {
                            responseMessage = "Project Request Validation Failed.  Select the project requests with a Validation Status of Invalid for further details.";
                        }
                        return new CAPS.Submission.ValidationResult(response.ValidationStatus, responseMessage);
                    });
            }
            else {
                return new CAPS.Submission.ValidationResult(false, "Project Request Validation Failed.  Select the project requests with a Validation Status of Invalid for further details.");
            }
        },
        function (e) {

            return new CAPS.Submission.ValidationResult(false, "Project Request Validation Failed. Details:" + e.message);
        }
    );
}

/**
 * Called from CAPS.Submission.Validate, this function checks that a board resolution has been attached if needed.
  * @param {any} formContext form context
 * @returns  {CAPS.Submission.ValidationResult} object that contains the validation results
 */
CAPS.Submission.CheckBoardResolution = function (formContext) {
    
    //check if Major or Minor
    var callForSubmissionType = formContext.getAttribute("caps_callforsubmissiontype").getValue();
    var submissionBoardResolutionAttached = formContext.getAttribute("caps_boardofresolutionattached").getValue();

    if (callForSubmissionType === 100000000 || callForSubmissionType === 100000001) {
        var callForSubmission = CAPS.Submission.GetLookup("caps_callforsubmission", formContext);
        var options = "?$select=caps_boardresolutionrequired"
        return Xrm.WebApi.retrieveRecord("caps_callforsubmission", callForSubmission.id, options).then(
            function success(result) {
                
                var boardResolutionRequired = result.caps_boardresolutionrequired;
                if (boardResolutionRequired === true && submissionBoardResolutionAttached === false) {
                    return new CAPS.Submission.ValidationResult(false, "Board Resolution Validation Failed.");
                }
                else if ((boardResolutionRequired === true && submissionBoardResolutionAttached === true) || (boardResolutionRequired === false)) {
                    return new CAPS.Submission.ValidationResult(true, "Board Resolution Validation Succeeded.");
                }
            },
            function (error) {
                console.log(error.message);
                // handle error conditions
            }

        );
        //if (formContext.getAttribute("caps_boardresolution").getValue()) {
        //    return new CAPS.Submission.ValidationResult(true, "Board Resolution Validation Succeeded.");
        //}
        //else {
        //    return new CAPS.Submission.ValidationResult(false, "Board Resolution Validation Failed.");
        //}
    }
}

/**
 * Called from CAPS.Submission.Validate, this function checks that the AFG variance is 0.
  * @param {any} formContext form context
 * @returns  {CAPS.Submission.ValidationResult} object that contains the validation results
 */
CAPS.Submission.CheckAFGTotal = function (formContext) {
    if (formContext.getAttribute("caps_callforsubmissiontype").getValue() === 100000002) {
        //Validate AFG Total
        if (formContext.getAttribute("caps_variance").getValue() !== 0) {
            return new CAPS.Submission.ValidationResult(false, "AFG Allocation Validation Failed.  The variance should be 0.");
        }
    }
}

/***
Called from CAPS.Submission.Validate, this function checks that Nothing to Submit is selected if there are no related project requests
and not selected if there are.
*/
CAPS.Submission.CheckNothingToSubmit = function (formContext) {
    
    //Get flag to identify if enrolment validation is required
    var nothingToSubmitAttribute = formContext.getAttribute("caps_noprojectstosubmit");
    var nothingToSubmit = false;
    if (nothingToSubmitAttribute != null) {
        nothingToSubmit = nothingToSubmitAttribute.getValue();
    }


    //Get record ID
    var submissionId = formContext.data.entity.getId();

    var fetchXML = "<fetch version=\"1.0\" output-format=\"xml - platform\" mapping=\"logical\" distinct=\"false\">" +
        "<entity name = \"caps_project\" >" +
        "<attribute name=\"caps_projectid\" />" +
        "<attribute name=\"caps_projectcode\" />" +
        "<order attribute=\"caps_projectcode\" descending=\"false\" />" +
        "<filter type=\"and\">" +
        "<condition attribute=\"caps_submission\" operator=\"eq\" value=\"" + submissionId + "\" />" +
        "</filter>" +
        "</entity>" +
        "</fetch >";

    //Get all Facilities on the school district
    return Xrm.WebApi.retrieveMultipleRecords("caps_project", "?fetchXml=" + fetchXML).then(
        function success(result) {

            if (result.entities.length > 0 && nothingToSubmit) {

                return new CAPS.Submission.ValidationResult(false, "Project Count Validation: There are project requests on this capital plan.  Please set \"I confirm I am submitting an empty plan\" to No or remove the project requests.");
            }
            else if (result.entities.length == 0 && !nothingToSubmit) {
                return new CAPS.Submission.ValidationResult(false, "Project Count Validation: There are no project requests on this capital plan.  Please set \"I confirm I am submitting an empty plan\" to Yes or add a project request.");
            }
            return new CAPS.Submission.ValidationResult(true, "Project Count Validation: Success");
        },
        function (error) {
            alert(error.message);
        }
    );
}

/*
Checking that there isn't already an approved AFG Expenditure plan
*/
CAPS.Submission.CheckAFGProjects = function (formContext) {
    //Get record ID
    var submissionId = formContext.data.entity.getId();

    //Get Submission type
    var submissionType = formContext.getAttribute("caps_submissiontype").getValue();

    //get School District
    var schoolDistrict = formContext.getAttribute("caps_schooldistrict").getValue();
    var schoolDistrictId = schoolDistrict[0].id;

    if (submissionType !== 200870001) return;


    var fetchXML = "<fetch version=\"1.0\" output-format=\"xml - platform\" mapping=\"logical\" distinct=\"false\">" +
        "<entity name = \"caps_submission\" >" +
        "<attribute name=\"caps_submissionid\" />" +
        "<attribute name=\"caps_name\" />" +
        "<order attribute=\"caps_name\" descending=\"false\" />" +
        "<filter type=\"and\">" +
        "<condition attribute=\"caps_submissiontype\" operator=\"eq\" value=\"200870001\" />" +
        "<condition attribute=\"statuscode\" operator=\"eq\" value=\"200870001\" />" +
        "<condition attribute=\"caps_schooldistrict\" operator=\"eq\" value=\"" + schoolDistrictId + "\" />" +
        "</filter>" +
        "</entity>" +
        "</fetch >";


    return Xrm.WebApi.retrieveMultipleRecords("caps_submission", "?fetchXml=" + fetchXML).then(
        function success(result) {

            if (result.entities.length > 0) {
                //Some bad projects
                return new CAPS.Submission.ValidationResult(false, "AFG Validation: You must complete your previous AFG expenditure plan before submitting a new one.");
            }
            return new CAPS.Submission.ValidationResult(true, "AFG Validation: Success");
        },
        function (error) {
            alert(error.message);
        }
    );
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
