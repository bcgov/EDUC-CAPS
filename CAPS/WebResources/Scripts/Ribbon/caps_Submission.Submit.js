"use strict";

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

    var resultsMessage = "Validation Results\r\n";

    let checkEnrolmentResult = await CAPS.Submission.CheckEnrolmentProjections(formContext);

    let checkBEPResult = await CAPS.Submission.CheckBEPProjects(formContext);

    let checkBUSResult = await CAPS.Submission.CheckBUSProjects(formContext);

    resultsArray.push(checkEnrolmentResult);
    resultsArray.push(checkBEPResult);
    resultsArray.push(checkBUSResult);

    var allValidationPassed = true;
    //loop through results
    resultsArray.forEach(function (item, index) {
        if (typeof item !== 'undefined' && item !== null) {
            if (!item.validationResult) {
                allValidationPassed = false;
            }

            resultsMessage += item.validationMessage + "\r\n";
        }
    });

    //Show Results
    formContext.getControl("caps_validationresults").setVisible(true);
    formContext.getAttribute("caps_validationresults").setValue(resultsMessage);

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
            "</filter>" +
            "<link-entity name=\"caps_facility\" from=\"caps_facilityid\" to=\"caps_facility\" link-type=\"inner\" alias=\"ab\">"+
                "<filter type=\"and\">"+
                    "<condition attribute=\"caps_schooldistrict\" operator=\"eq\" value=\""+schoolDistrictId+"\" /> "+
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
