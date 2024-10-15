"use strict";

var CAPS = CAPS || {};
CAPS.ChildCareFacility = CAPS.ChildCareFacility || {};

CAPS.ChildCareFacility.form_onload = function (executionContext) {
    var formContext = executionContext.getFormContext();
    CAPS.ChildCareFacility.updateAddressFields(formContext);

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