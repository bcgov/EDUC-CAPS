"use strict";

var CAPS = CAPS || {};
CAPS.ChildCareFacility = CAPS.ChildCareFacility || {};

CAPS.ChildCareFacility.form_onload = function (executionContext) {
    var formContext = executionContext.getFormContext();
    CAPS.ChildCareFacility.updateAddressFields(formContext);
    formContext.getAttribute("caps_sameaddressasschoolfacility").addOnChange(function () {
        CAPS.ChildCareFacility.updateAddressFields(formContext);
    });
    formContext.getAttribute("caps_facility").addOnChange(function () {
        CAPS.ChildCareFacility.updateAddressFields(formContext);
    });
};

// Update the address fields based on the caps_sameaddressasschoolfacility toggle and the selected school facility
CAPS.ChildCareFacility.updateAddressFields = function (formContext) {
    var isSameAddress = formContext.getAttribute("caps_sameaddressasschoolfacility").getValue();
    var facility = formContext.getAttribute("caps_facility").getValue();

    if (isSameAddress && facility && facility[0] && facility[0].id) {
        var facilityId = facility[0].id.replace("{", "").replace("}", "");

        Xrm.WebApi.retrieveRecord("caps_facility", facilityId, "?$select=caps_poboxaddress,caps_streetaddress").then(
            function (result) {
                var poBoxAddress = result.caps_poboxaddress;
                var streetAddress = result.caps_streetaddress;

                formContext.getAttribute("caps_mailingaddress").setValue(poBoxAddress);
                formContext.getAttribute("caps_streetaddress").setValue(streetAddress);

                formContext.getControl("caps_mailingaddress").setDisabled(true);
                formContext.getControl("caps_streetaddress").setDisabled(true);

                formContext.getAttribute("caps_mailingaddress").setRequiredLevel("none");
                formContext.getAttribute("caps_streetaddress").setRequiredLevel("none");
            },
            function (error) {
                console.log("Error retrieving facility address: " + error.message);
            }
        );
    } else if (!isSameAddress) {
        formContext.getAttribute("caps_mailingaddress").setValue(null);
        formContext.getAttribute("caps_streetaddress").setValue(null);

        formContext.getControl("caps_mailingaddress").setDisabled(false);
        formContext.getControl("caps_streetaddress").setDisabled(false);

        formContext.getAttribute("caps_streetaddress").setRequiredLevel("required");
    }
};