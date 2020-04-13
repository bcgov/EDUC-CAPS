"use strict";

var CAPS = CAPS || {};
CAPS.Submission = CAPS.Submission || {};

/**
 * Called from Capital Plan (caps_submission) form.  This function validates the capital plan, displays any errors and submits the plan if there are no validation errors.
 * @param {any} primaryControl primary control
 */
CAPS.Submission.Validate = async function (primaryControl) {
    var formContext = primaryControl;

    //Declare variables
    var enrolmentResults = {};

    Xrm.Utility.showProgressIndicator("Validating Capital Plan...");

    //Get flag to identify if enrolment validation is required
    var doEnrolmentValidation = formContext.getAttribute("caps_validateenrolment").getValue();

    if (doEnrolmentValidation) {
        //Do Enrolment Validation
        let checkEnrolmentResult = CAPS.Submission.CheckEnrolmentProjections(formContext);
        enrolmentResults = await checkEnrolmentResult;

        //Show Results
        formContext.getControl("caps_validationresults").setVisible(true);

        var resultsMessage = "Validation Results\r\n";
        resultsMessage += "Enrolment Validation: " + enrolmentResults.validationMessage + "\r\n";

        formContext.getAttribute("caps_validationresults").setValue(resultsMessage);
    }

    if (!doEnrolmentValidation || enrolmentResults.validationResult) {
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
                return new CAPS.Submission.ValidationResult(false, "There are missing enrolment projections");
            }
            return new CAPS.Submission.ValidationResult(true, "Success");
        },
        function (error) {
            alert(error.message);
        }
    );
}