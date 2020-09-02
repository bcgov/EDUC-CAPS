﻿"use strict";

var CAPS = CAPS || {};
CAPS.Submission = CAPS.Submission || {};

/**
 * Called from Capital Plan (caps_submission) form.  This function validates the capital plan, displays any errors and submits the plan if there are no validation errors.
 * @param {any} primaryControl primary control
 */
CAPS.Submission.Validate = async function (primaryControl) {
    debugger;
    var formContext = primaryControl;

    //Declare variables
    var resultsArray = [];

    Xrm.Utility.showProgressIndicator("Validating Capital Plan...");

    var resultsMessage = "Validation Errors\r\n";

    let checkEnrolmentResult = await CAPS.Submission.CheckEnrolmentProjections(formContext);

    //let checkBEPResult = await CAPS.Submission.CheckBEPProjects(formContext);

    //let checkBUSResult = await CAPS.Submission.CheckBUSProjects(formContext);

    let checkProjectRequestResults = await CAPS.Submission.CheckProjectRequests(formContext);

    let checkBoardResolutionResults = await CAPS.Submission.CheckBoardResolution(formContext);

    let checkAFGVariance = await CAPS.Submission.CheckAFGTotal(formContext);

    resultsArray.push(checkEnrolmentResult);
    //resultsArray.push(checkBEPResult);
    //resultsArray.push(checkBUSResult);
    resultsArray.push(checkProjectRequestResults);
    resultsArray.push(checkBoardResolutionResults);
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
    //formContext.getControl("caps_validationresults").setVisible(true);
    //formContext.getAttribute("caps_validationresults").setValue(resultsMessage);


    formContext.ui.tabs.forEach(function (tab, i) {
        //loop through sections
        //if (arrTabNames.includes(tab.getName())) {
            //loop through sections
            tab.sections.forEach(function (section, j) {
                section.controls.forEach(function (control, k) {

                    if (control.getAttribute().getName() === "caps_validationresults") {
                            control.getAttribute().setValue(resultsMessage);
                            control.setVisible(true);
                    }
                });
            });
        //}
    });


    if (allValidationPassed) {
        //all good so change status or display confirmation
        var confirmStrings = { text: "This capital plan is valid and ready for submission.  Click OK to Submit or Cancel to exit.", title: "Submission Confirmation" };
        var confirmOptions = { height: 200, width: 450 };
        Xrm.Navigation.openConfirmDialog(confirmStrings, confirmOptions).then(
            function (success) {
                if (success.confirmed) {
                    formContext.getAttribute("statuscode").setValue(100000000);
                    formContext.data.entity.save();
                }

            });
    }
    else {
        var alertStrings = { confirmButtonLabel: "Ok", text: "The validation of this capital plan failed.  Please see the Validation results section below for details.", title: "Validation Result" };
        var alertOptions = { height: 120, width: 260 };
        Xrm.Navigation.openAlertDialog(alertStrings, alertOptions);
    }

    //close the indicator
    Xrm.Utility.closeProgressIndicator();

    ////Get flag to identify if enrolment validation is required
    //var doEnrolmentValidation = formContext.getAttribute("caps_validateenrolment").getValue();

    //if (doEnrolmentValidation) {
    //    //Do Enrolment Validation
    //    let checkEnrolmentResult = CAPS.Submission.CheckEnrolmentProjections(formContext);
    //    enrolmentResults = await checkEnrolmentResult;

    //    //Show Results
    //    formContext.getControl("caps_validationresults").setVisible(true);

    //    resultsMessage += "Enrolment Validation: " + enrolmentResults.validationMessage + "\r\n";
    //}

    //formContext.getAttribute("caps_validationresults").setValue(resultsMessage);

    //if (!doEnrolmentValidation || enrolmentResults.validationResult) {
    //    //all good so change status or display confirmation
    //    var confirmStrings = { text: "This capital plan is valid and ready for submission.  Click OK to Submit or Cancel to exit.", title: "Submission Confirmation" };
    //    var confirmOptions = { height: 200, width: 450 };
    //    Xrm.Navigation.openConfirmDialog(confirmStrings, confirmOptions).then(
    //        function (success) {
    //            if (success.confirmed) {
    //                formContext.getAttribute("statuscode").setValue(100000000);
    //                formContext.data.entity.save();
    //            }

    //        });
    //}
    //else {
    //    var alertStrings = { confirmButtonLabel: "Ok", text: "The validation of this capital plan failed.  Please see the Validation results section below for details.", title: "Validation Result" };
    //    var alertOptions = { height: 120, width: 260 };
    //    Xrm.Navigation.openAlertDialog(alertStrings, alertOptions);
    //}
    
    ////close the indicator
    //Xrm.Utility.closeProgressIndicator();
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
    debugger;
    //Get flag to identify if enrolment validation is required
    var doEnrolmentValidation = formContext.getAttribute("caps_validateenrolment").getValue();

    if (!doEnrolmentValidation) return;

    //Get the School District
    var schoolDistrict = formContext.getAttribute("caps_schooldistrict").getValue();
    var schoolDistrictId = schoolDistrict[0].id;

    var fetchXML = "<fetch version=\"1.0\" output-format=\"xml - platform\" mapping=\"logical\" distinct=\"false\">"+
         "<entity name=\"caps_enrolmentprojections_sd\" >" +
            "<attribute name=\"caps_enrolmentprojections_sdid\" />" +
            "<attribute name=\"caps_name\" />" +
            "<attribute name=\"createdon\" />" +
            "<order attribute=\"caps_name\" descending=\"false\" />" +
            "<filter type=\"and\">" +
        "<condition attribute=\"caps_isvalid\" operator=\"eq\" value=\"0\" />" +
        "<condition attribute=\"statecode\" operator=\"eq\"  value=\"0\" />" +
            "</filter>" +
            "<link-entity name=\"caps_facility\" from=\"caps_facilityid\" to=\"caps_facility\" link-type=\"inner\" alias=\"ab\">"+
                "<filter type=\"and\">"+
        "<condition attribute=\"caps_schooldistrict\" operator=\"eq\" value=\"" + schoolDistrictId + "\" /> " +
                "</filter>"+
            "</link-entity>" +
  "</entity>"+
        "</fetch>";

    //Get all Facilities on the school district
    return Xrm.WebApi.retrieveMultipleRecords("caps_enrolmentprojections_sd", "?fetchXml=" + fetchXML).then(
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
        "<condition attribute=\"caps_nonreplaceable\" operator=\"eq\" value=\"1\" />" +
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
            debugger;
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
            debugger;
            return new CAPS.Submission.ValidationResult(false, "Project Request Validation Failed. Details:"+e.message);
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

    if (callForSubmissionType === 100000000 || callForSubmissionType === 100000001) {
        if (formContext.getAttribute("caps_boardofresolutionattached").getValue()) {
            return new CAPS.Submission.ValidationResult(true, "Board Resolution Validation Succeeded.");
        }
        else {
            return new CAPS.Submission.ValidationResult(false, "Board Resolution Validation Failed.");
        }
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

