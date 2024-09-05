"use strict";

var CAPS = CAPS || {};
CAPS.ChildCareActualEnrolment = CAPS.ChildCareActualEnrolment || {};

CAPS.ChildCareActualEnrolment.form_onload = function (executionContext) {
    var formContext = executionContext.getFormContext();

    if (formContext) {
        CAPS.ChildCareActualEnrolment.updateLicensedCapacityFromFacility(formContext);
        CAPS.ChildCareActualEnrolment.updateCapacityTotal(formContext);
        CAPS.ChildCareActualEnrolment.checkOperatingAtFullCapacity(formContext);

        formContext.getAttribute("caps_childcarefacility").addOnChange(function () {
            CAPS.ChildCareActualEnrolment.updateLicensedCapacityFromFacility(formContext);
        });

        // Add onchange handlers for the 6 capacity fields
        var capacityFields = [
            "caps_capacityunder36months",
            "caps_capacitymultiage",
            "caps_capacity30monthstoschoolage",
            "caps_capacityschoolage",
            "caps_capacitypreschool",
            "caps_capacitysasg"
        ];

        capacityFields.forEach(function (field) {
            formContext.getAttribute(field)?.addOnChange(function () {
                CAPS.ChildCareActualEnrolment.updateCapacityTotal(formContext);
                CAPS.ChildCareActualEnrolment.checkOperatingAtFullCapacity(formContext);
            });
        });

        formContext.getAttribute("caps_reasonforunderutilization").addOnChange(function () {
            CAPS.ChildCareActualEnrolment.updateUtilizationFields(formContext);
        });

        CAPS.ChildCareActualEnrolment.updateUtilizationFields(formContext);
    }
};

// Update caps_licensedcapacitytotal based on the selected childcare facility
CAPS.ChildCareActualEnrolment.updateLicensedCapacityFromFacility = function (formContext) {
    var childcareFacility = formContext.getAttribute("caps_childcarefacility")?.getValue();

    if (childcareFacility && childcareFacility[0] && childcareFacility[0].id) {
        var facilityId = childcareFacility[0].id.replace("{", "").replace("}", "");

        // Retrieve caps_licensedcapacitytotal from the caps_childcare table (child care facility)
        Xrm.WebApi.retrieveRecord("caps_childcare", facilityId, "?$select=caps_licensedcapacitytotal").then(
            function (result) {
                var facilityLicensedCapacity = result.caps_licensedcapacitytotal;

                if (facilityLicensedCapacity !== null) {
                    // Set the value of caps_licensedcapacitytotal on the current form
                    formContext.getAttribute("caps_licensedcapacitytotal").setValue(facilityLicensedCapacity);
                    CAPS.ChildCareActualEnrolment.checkOperatingAtFullCapacity(formContext);
                }
            },
            function (error) {
                console.log("Error retrieving facility licensed capacity: " + error.message);
            }
        );
    } else {
        // Clear the caps_licensedcapacitytotal field if no childcare facility is selected
        formContext.getAttribute("caps_licensedcapacitytotal").setValue(null);
        formContext.getAttribute("caps_operatingatfulllicensecapacity").setValue(false);
    }
};

// Update the caps_capacitytotal based on the 6 capacity fields
CAPS.ChildCareActualEnrolment.updateCapacityTotal = function (formContext) {
    var capacityFields = [
        "caps_capacityunder36months",
        "caps_capacitymultiage",
        "caps_capacity30monthstoschoolage",
        "caps_capacityschoolage",
        "caps_capacitypreschool",
        "caps_capacitysasg"
    ];

    // Sum up the values of the 6 capacity fields
    var totalCapacity = capacityFields.reduce(function (total, field) {
        var fieldValue = formContext.getAttribute(field)?.getValue();
        return total + (fieldValue !== null ? fieldValue : 0);
    }, 0);

    // Set the total value to caps_capacitytotal
    formContext.getAttribute("caps_capacitytotal").setValue(totalCapacity);
};

// Check if the form is operating at full license capacity by comparing caps_licensedcapacitytotal and caps_capacitytotal
CAPS.ChildCareActualEnrolment.checkOperatingAtFullCapacity = function (formContext) {
    var licensedCapacityTotal = formContext.getAttribute("caps_licensedcapacitytotal")?.getValue();
    var capacityTotal = formContext.getAttribute("caps_capacitytotal")?.getValue();

    if (licensedCapacityTotal !== null && capacityTotal !== null) {
        // Operating at full capacity if capacity total is greater than or equal to licensed capacity
        var isOperatingAtFullCapacity = capacityTotal >= licensedCapacityTotal;
        formContext.getAttribute("caps_operatingatfulllicensecapacity").setValue(isOperatingAtFullCapacity);

        // Update the visibility and requirements for additional fields
        CAPS.ChildCareActualEnrolment.updateUtilizationFields(formContext);
    } else {
        formContext.getAttribute("caps_operatingatfulllicensecapacity").setValue(false);
        CAPS.ChildCareActualEnrolment.updateUtilizationFields(formContext);
    }
};

// Update the visibility and requirement of caps_reasonforunderutilization and caps_ifothercommentonunderutilizationreasons
CAPS.ChildCareActualEnrolment.updateUtilizationFields = function (formContext) {
    var isOperatingAtFullCapacity = formContext.getAttribute("caps_operatingatfulllicensecapacity")?.getValue();
    var reasonForUnderutilization = formContext.getAttribute("caps_reasonforunderutilization")?.getValue();
    var isOtherSelected = reasonForUnderutilization && reasonForUnderutilization.includes(746660002); // "Other"

    if (!isOperatingAtFullCapacity) { // Show fields when not operating at full capacity
        formContext.getControl("caps_reasonforunderutilization").setVisible(true);
        formContext.getAttribute("caps_reasonforunderutilization").setRequiredLevel("required");

        // Show and require the comment field if "Other" is selected
        if (isOtherSelected) {
            formContext.getControl("caps_ifothercommentonunderutilizationreasons").setVisible(true);
            formContext.getAttribute("caps_ifothercommentonunderutilizationreasons").setRequiredLevel("required");
        } else {
            formContext.getControl("caps_ifothercommentonunderutilizationreasons").setVisible(false);
            formContext.getAttribute("caps_ifothercommentonunderutilizationreasons").setValue(null);
            formContext.getAttribute("caps_ifothercommentonunderutilizationreasons").setRequiredLevel("none");
        }
    } else {
        formContext.getControl("caps_reasonforunderutilization").setVisible(false);
        formContext.getAttribute("caps_reasonforunderutilization").setValue(null);
        formContext.getAttribute("caps_reasonforunderutilization").setRequiredLevel("none");

        formContext.getControl("caps_ifothercommentonunderutilizationreasons").setVisible(false);
        formContext.getAttribute("caps_ifothercommentonunderutilizationreasons").setValue(null);
        formContext.getAttribute("caps_ifothercommentonunderutilizationreasons").setRequiredLevel("none");
    }
};
