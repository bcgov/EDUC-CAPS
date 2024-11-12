"use strict";

var CAPS = CAPS || {};
CAPS.ChildCareFacility = CAPS.ChildCareFacility || {};

CAPS.ChildCareFacility.form_onload = function (executionContext) {
    var formContext = executionContext.getFormContext();
    CAPS.ChildCareFacility.updateAddressFields(formContext);
    CAPS.ChildCareFacility.UpdateActualProjectedEnrolment(executionContext);

    formContext.getAttribute("caps_sameaddressasschoolfacility").addOnChange(function () {
        CAPS.ChildCareFacility.updateAddressFields(formContext, true);
    });

    formContext.getAttribute("caps_facility").addOnChange(function () {
        CAPS.ChildCareFacility.updateAddressFields(formContext);
    });
};

// Update the address fields based on the caps_sameaddressasschoolfacility toggle and the selected school facility
// The 'forceClear' parameter ensures we clear the address fields when the toggle is manually changed to false
CAPS.ChildCareFacility.updateAddressFields = function (formContext, forceClear = false) {
    var isSameAddress = formContext.getAttribute("caps_sameaddressasschoolfacility").getValue();
    var facility = formContext.getAttribute("caps_facility").getValue();

    var mailingAddress = formContext.getAttribute("caps_mailingaddress").getValue();
    var streetAddress = formContext.getAttribute("caps_streetaddress").getValue();
    var postalCode = formContext.getAttribute("caps_postalcode").getValue();

    // Check if a valid facility is selected
    if (facility && facility.length > 0 && facility[0].id) {

        formContext.getControl("caps_sameaddressasschoolfacility").setVisible(true);
        formContext.getControl("caps_mailingaddress").setVisible(true);
        formContext.getControl("caps_streetaddress").setVisible(true);
        formContext.getControl("caps_postalcode").setVisible(true);

        if (isSameAddress) {
            var facilityId = facility[0].id.replace("{", "").replace("}", "");

            Xrm.WebApi.retrieveRecord("caps_facility", facilityId, "?$select=caps_poboxaddress,caps_streetaddress,caps_postalcode").then(
                function (result) {
                    var poBoxAddress = result.caps_poboxaddress;
                    var streetAddressFromFacility = result.caps_streetaddress;
                    var postalCodefromFacility = result.caps_postalcode;

                    // Set the facility address values to the form's mailing and street address fields
                    formContext.getAttribute("caps_mailingaddress").setValue(poBoxAddress);
                    formContext.getAttribute("caps_streetaddress").setValue(streetAddressFromFacility);
                    formContext.getAttribute("caps_postalcode").setValue(postalCodefromFacility);

                    // Disable the fields and make them optional
                    formContext.getControl("caps_mailingaddress").setDisabled(true);
                    formContext.getControl("caps_streetaddress").setDisabled(true);
                    formContext.getControl("caps_postalcode").setDisabled(true);

                    formContext.getAttribute("caps_mailingaddress").setRequiredLevel("none");
                    formContext.getAttribute("caps_streetaddress").setRequiredLevel("none");
                    formContext.getAttribute("caps_postalcode").setRequiredLevel("none");
                },
                function (error) {
                    Xrm.Utility.alertDialog("Error retrieving facility address: " + error.message);
                }
            );
        } else {
            // If "Same address as school facility" is unchecked, clear fields if needed
            if (forceClear || (!mailingAddress && !streetAddress && !postalCode)) {
                formContext.getAttribute("caps_mailingaddress").setValue(null);
                formContext.getAttribute("caps_streetaddress").setValue(null);
                formContext.getAttribute("caps_postalcode").setValue(null);
            }

            // Enable the fields for manual entry and make street address required
            formContext.getControl("caps_mailingaddress").setDisabled(false);
            formContext.getControl("caps_streetaddress").setDisabled(false);
            formContext.getControl("caps_postalcode").setDisabled(false);

            formContext.getAttribute("caps_streetaddress").setRequiredLevel("required");
            formContext.getAttribute("caps_postalcode").setRequiredLevel("required");
        }

    } else {
        // Facility is null or empty, hide and clear all relevant fields
        formContext.getControl("caps_sameaddressasschoolfacility").setVisible(false);
        formContext.getControl("caps_mailingaddress").setVisible(false);
        formContext.getControl("caps_streetaddress").setVisible(false);
        formContext.getControl("caps_postalcode").setVisible(false);

        formContext.getAttribute("caps_mailingaddress").setValue(null);
        formContext.getAttribute("caps_streetaddress").setValue(null);
        formContext.getAttribute("caps_postalcode").setValue(null);
    }
};

CAPS.ChildCareFacility.UpdateActualProjectedEnrolment = function (executionContext) {

    var formContext = executionContext.getFormContext();
    var currentCCActualEnrolment = CAPS.ChildCareFacility.GetLookup("caps_currentchildcareactualenrolment", formContext);
    if (currentCCActualEnrolment !== undefined) {
        var options = "?$select=caps_capacityunder36months,caps_capacity30monthstoschoolage,caps_capacitypreschool,caps_capacitymultiage,caps_capacityschoolage,caps_capacitysasg";
        Xrm.WebApi.retrieveRecord("caps_childcareactualenrolment", CAPS.ChildCareFacility.RemoveCurlyBraces(currentCCActualEnrolment.id), options).then(
            function success(result) {
                var capacityUnder36Months = result.caps_capacityunder36months;
                var capacity30MonthsToSchoolAge = result.caps_capacity30monthstoschoolage;
                var capacityPreSchool = result.caps_capacitypreschool;
                var capacityMultiAge = result.caps_capacitymultiage;
                var capacitySchoolAge = result.caps_capacityschoolage;
                var capacitySASG = result.caps_capacitysasg;

                formContext.getAttribute("caps_under36months_currentenrolment").setValue(capacityUnder36Months);
                formContext.getAttribute("caps_30monthstoschoolage_currentenrolment").setValue(capacity30MonthsToSchoolAge);
                formContext.getAttribute("caps_preschool_currentenrolment").setValue(capacityPreSchool);
                formContext.getAttribute("caps_multiage_currentenrolment").setValue(capacityMultiAge);
                formContext.getAttribute("caps_schoolage_currentenrolment").setValue(capacitySchoolAge);
                formContext.getAttribute("caps_sasg_currentenrolment").setValue(capacitySASG);

            },
            function (error) {
                console.log(error.message);
            }
        );
    }
    else if (currentCCActualEnrolment === undefined) {
        var recordId = formContext.data.entity.getId().replace("{", "").replace("}", "");
        Xrm.WebApi.retrieveMultipleRecords("edu_year", "?$filter=edu_type eq 757500001 and statuscode eq 1").then(
            function retrieveCurrentYearSuccess(currentYearRecord) {
                if (currentYearRecord == null || currentYearRecord.entities.length == 0) {
                    // Not Found for some reason
                    return;
                }
                // Get Current Enrolment Projection
                // Statuscdoe = 714430001 - Current
                // Projection Year matching that obtained from previous query
                // Facility matching that of current record
                var childCareEnrolmentProjectionOptions = "?$filter=statuscode eq 714430001 and _caps_childcarefacility_value eq '" + recordId + "' and _caps_schoolyear_value eq '" + currentYearRecord.entities[0].edu_yearid + "'";
                Xrm.WebApi.retrieveMultipleRecords("caps_childcareenrolmentprojection", childCareEnrolmentProjectionOptions).then(
                    function success(result) {
                        if (result == null || result.entities.length == 0) {
                            return; // Nothing to process
                        }
                        var hasChanges = false;
                        var enrolmentProjUnder36Months = result.entities[0].caps_under36months;
                        var enrolmentProj30MonthsToSchoolAge = result.entities[0].caps_monthstoschoolage;
                        var enrolmentProjePreSchool = result.entities[0].caps_preschool;
                        var enrolmentProjeMultiAge = result.entities[0].caps_multiage;
                        var enrolmentProjSchoolAge = result.entities[0].caps_schoolage;
                        var enrolmentProjSASG = result.entities[0].caps_schoolageonschoolgrounds;
                        if (enrolmentProjUnder36Months != null) {

                            formContext.getAttribute("caps_under36months_currentenrolment").setValue(enrolmentProjUnder36Months);
                            hasChanges = true;
                        }
                        if (enrolmentProj30MonthsToSchoolAge !== null) {
                            formContext.getAttribute("caps_30monthstoschoolage_currentenrolment").setValue(enrolmentProj30MonthsToSchoolAge);
                            hasChanges = true;
                        }
                        if (enrolmentProjePreSchool !== null) {
                            formContext.getAttribute("caps_preschool_currentenrolment").setValue(enrolmentProjePreSchool);
                            hasChanges = true;
                        }
                        if (enrolmentProjeMultiAge !== null) {
                            formContext.getAttribute("caps_multiage_currentenrolment").setValue(enrolmentProjeMultiAge);
                            hasChanges = true;
                        }
                        if (enrolmentProjSchoolAge !== null) {
                            formContext.getAttribute("caps_schoolage_currentenrolment").setValue(enrolmentProjSchoolAge);
                            hasChanges = true;
                        }
                        if (enrolmentProjSASG !== null) {
                            formContext.getAttribute("caps_sasg_currentenrolment").setValue(enrolmentProjSASG);
                            hasChanges = true;
                        }

                        if (hasChanges) {
                            formContext.data.entity.save(); // Save the record if there were changes made.
                        }

                        // Should only have 1 record
                        return;

                    });
            });
    }
    
};

CAPS.ChildCareFacility.GetLookup = function (fieldName, formContext) {
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
};

CAPS.ChildCareFacility.RemoveCurlyBraces = function (str) {
    return str.replace(/[{}]/g, "");
}
