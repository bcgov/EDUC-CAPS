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

    // If "Same address as school facility" is checked and a facility is selected, populate the address fields from the facility
    if (isSameAddress && facility && facility[0] && facility[0].id) {
        var facilityId = facility[0].id.replace("{", "").replace("}", "");

        Xrm.WebApi.retrieveRecord("caps_facility", facilityId, "?$select=caps_poboxaddress,caps_streetaddress").then(
            function (result) {
                var poBoxAddress = result.caps_poboxaddress;
                var streetAddressFromFacility = result.caps_streetaddress;

                // Set the facility address values to the form's mailing and street address fields
                formContext.getAttribute("caps_mailingaddress").setValue(poBoxAddress);
                formContext.getAttribute("caps_streetaddress").setValue(streetAddressFromFacility);

                // Disable the fields and make them optional (not required)
                formContext.getControl("caps_mailingaddress").setDisabled(true);
                formContext.getControl("caps_streetaddress").setDisabled(true);

                formContext.getAttribute("caps_mailingaddress").setRequiredLevel("none");
                formContext.getAttribute("caps_streetaddress").setRequiredLevel("none");
            },
            function (error) {
                console.log("Error retrieving facility address: " + error.message);
            }
        );
    }
    // If "Same address as school facility" is unchecked, handle address clearing and enabling for manual input
    else if (!isSameAddress) {
        // If forceClear is true (i.e., when the user manually changes the toggle), clear the address fields
        if (forceClear || (!mailingAddress && !streetAddress)) {
            formContext.getAttribute("caps_mailingaddress").setValue(null);
            formContext.getAttribute("caps_streetaddress").setValue(null);
        }

        // Enable the fields for manual entry and make street address required
        formContext.getControl("caps_mailingaddress").setDisabled(false);
        formContext.getControl("caps_streetaddress").setDisabled(false);

        formContext.getAttribute("caps_streetaddress").setRequiredLevel("required");
    }
};