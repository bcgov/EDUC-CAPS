"use strict";

var CAPS = CAPS || {};
CAPS.ChildCareActualEnrolment = CAPS.ChildCareActualEnrolment || {};

CAPS.ChildCareActualEnrolment.form_onload = function (executionContext) {
    var formContext = executionContext.getFormContext();
    CAPS.ChildCareActualEnrolment.updateUtilizationFields(formContext);

    formContext.getAttribute("caps_operatingatfulllicensecapacity").addOnChange(function () {
        CAPS.ChildCareActualEnrolment.updateUtilizationFields(formContext);
    });

    formContext.getAttribute("caps_reasonforunderutilization").addOnChange(function () {
        CAPS.ChildCareActualEnrolment.updateUtilizationFields(formContext);
    });
};

// Update the visibility and requirement of caps_reasonforunderutilization and caps_ifothercommentonunderutilizationreasons
CAPS.ChildCareActualEnrolment.updateUtilizationFields = function (formContext) {
    var isOperatingAtFullCapacity = formContext.getAttribute("caps_operatingatfulllicensecapacity").getValue();
    var capacityTotal = formContext.getAttribute("caps_capacitytotal").getValue(); // Check capacity total
    var reasonForUnderutilization = formContext.getAttribute("caps_reasonforunderutilization").getValue();
    var isOtherSelected = reasonForUnderutilization && reasonForUnderutilization.includes(746660002); // 746660002 = "other"

    // Only apply logic if caps_capacitytotal is not null
    if (capacityTotal !== null) {
        if (!isOperatingAtFullCapacity) { // Show fields when caps_operatingatfulllicensecapacity is false
            formContext.getControl("caps_reasonforunderutilization").setVisible(true);
            formContext.getAttribute("caps_reasonforunderutilization").setRequiredLevel("required");

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
    }
};