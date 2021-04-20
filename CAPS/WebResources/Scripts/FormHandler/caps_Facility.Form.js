"use strict";

var CAPS = CAPS || {};
CAPS.Facility = CAPS.Facility || {};

const STATUS_MISSMATCH_NOTIFICATION = "Status_Missmatch_Notification";

/***
Main function for Facility.  This function calls all other functions.
**/
CAPS.Facility.onLoad = function (executionContext) {
    var formContext = executionContext.getFormContext();

    CAPS.Facility.ShowStatusMissmatch(executionContext);
    formContext.getAttribute("statecode").addOnChange(CAPS.Facility.ShowStatusMissmatch);
    formContext.getAttribute("caps_closedate").addOnChange(CAPS.Facility.ShowStatusMissmatch);
    formContext.getAttribute("caps_closuredescription").addOnChange(CAPS.Facility.ShowStatusMissmatch);

    CAPS.Facility.ShowHideTabs(executionContext);
    formContext.getAttribute("caps_currentfacilitytype").addOnChange(CAPS.Facility.ShowHideTabs);
    

    //Design Capacity Checks
    CAPS.Facility.ValidateKindergartenDesignCapacity(executionContext);
    formContext.getAttribute("caps_designcapacitykindergarten").addOnChange(CAPS.Facility.ValidateKindergartenDesignCapacity);

    CAPS.Facility.ValidateElementaryDesignCapacity(executionContext);
    formContext.getAttribute("caps_designcapacityelementary").addOnChange(CAPS.Facility.ValidateElementaryDesignCapacity);

    CAPS.Facility.ValidateSecondaryDesignCapacity(executionContext);
    formContext.getAttribute("caps_designcapacitysecondary").addOnChange(CAPS.Facility.ValidateSecondaryDesignCapacity);
}

/***
This function validates that the kindergarten design capacity is a multiple of the specified design capacity and shows a warning if it isn't.
***/
CAPS.Facility.ValidateKindergartenDesignCapacity = function (executionContext) {
    debugger;
    var formContext = executionContext.getFormContext();
    var designCapacity = formContext.getAttribute("caps_designcapacitykindergarten").getValue();
    var validateDesignCapacityKRequest = new CAPS.Facility.ValidateDesignCapacityRequest("Kindergarten", designCapacity);

    Xrm.WebApi.online.execute(validateDesignCapacityKRequest).then(
    function (result) {
        if (result.ok) {

            return result.json().then(
                function (response) {
                    debugger;
                    if (!response.IsValid) {
                        formContext.ui.setFormNotification(response.ValidationMessage, 'INFO', 'KINDERGARTEN DESIGN WARNING');
                    }
                    else {
                        formContext.ui.clearFormNotification('KINDERGARTEN DESIGN WARNING');
                    }
                });
        }
    },
    function (error) {
        console.log(error.message);
        // handle error conditions
    }
);
}

/***
This function validates that the elementary design capacity is a multiple of the specified design capacity and shows a warning if it isn't.
**/
CAPS.Facility.ValidateElementaryDesignCapacity = function (executionContext) {
    debugger;
    var formContext = executionContext.getFormContext();
    var designCapacity = formContext.getAttribute("caps_designcapacityelementary").getValue();
    var validateDesignCapacityKRequest = new CAPS.Facility.ValidateDesignCapacityRequest("Elementary", designCapacity);

    Xrm.WebApi.online.execute(validateDesignCapacityKRequest).then(
    function (result) {
        if (result.ok) {

            return result.json().then(
                function (response) {
                    debugger;
                    if (!response.IsValid) {
                        formContext.ui.setFormNotification(response.ValidationMessage, 'INFO', 'ELEMENTARY DESIGN WARNING');
                    }
                    else {
                        formContext.ui.clearFormNotification('ELEMENTARY DESIGN WARNING');
                    }
                });
        }
    },
    function (error) {
        console.log(error.message);
        // handle error conditions
    }
);
};

/***
This function validates that the secondary design capacity is a multiple of the specified design capacity and shows a warning if it isn't.
**/
CAPS.Facility.ValidateSecondaryDesignCapacity = function (executionContext) {
    debugger;
    var formContext = executionContext.getFormContext();
    var designCapacity = formContext.getAttribute("caps_designcapacitysecondary").getValue();
    var validateDesignCapacityKRequest = new CAPS.Facility.ValidateDesignCapacityRequest("Secondary", designCapacity);

    Xrm.WebApi.online.execute(validateDesignCapacityKRequest).then(
    function (result) {
        if (result.ok) {

            return result.json().then(
                function (response) {
                    debugger;
                    if (!response.IsValid) {
                        formContext.ui.setFormNotification(response.ValidationMessage, 'INFO', 'SECONDARY DESIGN WARNING');
                    }
                    else {
                        formContext.ui.clearFormNotification('SECONDARY DESIGN WARNING');
                    }
                });
        }
    },
    function (error) {
        console.log(error.message);
        // handle error conditions
    }
);
};

CAPS.Facility.ValidateDesignCapacityRequest = function (capacityType, capacityCount) {
    this.Type = capacityType;
    this.Count = capacityCount;
};

CAPS.Facility.ValidateDesignCapacityRequest.prototype.getMetadata = function () {
    return {
        boundParameter: null,
        parameterTypes: {
            "Type": {
                "typeName": "Edm.String",
                "structuralProperty": 1
            },
            "Count": {
                "typeName": "Edm.Int32",
                "structuralProperty": 1
            }
        },
        operationType: 0, // This is a function. Use '0' for actions and '2' for CRUD
        operationName: "caps_ValidateDesignCapacity"
    };
};

/**
This function shows a warning if the facility is set to closed but the closed date and reason haven't been filled in or if they have been filled in and the facility is still marked as open.
**/
CAPS.Facility.ShowStatusMissmatch = function (executionContext) {
    var formContext = executionContext.getFormContext();

    //caps_closedate
    //caps_closuredescription
    //statecode
    if (formContext.getAttribute("statecode").getValue() == 0 
        && (formContext.getAttribute("caps_closedate").getValue() != null || formContext.getAttribute("caps_closuredescription").getValue() != null)) {
        formContext.ui.setFormNotification('You have Close Date and/or Close Description set on an open facility. Either close this facility record or clear those closure fields.', 'INFO', STATUS_MISSMATCH_NOTIFICATION);
    }
    else {
        formContext.ui.clearFormNotification(STATUS_MISSMATCH_NOTIFICATION);
    }
}

/**
This function hides school specific tabs/sections for non-school facilities.
**/
CAPS.Facility.ShowHideTabs = function (executionContext) {
    var formContext = executionContext.getFormContext();
    if (formContext.getAttribute("caps_isschool").getValue()) {
        formContext.ui.tabs.get("tab_CapacityandUtilization").setVisible(true);
        formContext.ui.tabs.get("tab_EnrolmentProjections").setVisible(true);
        formContext.ui.tabs.get("tab_PastEnrolmentProjections").setVisible(true);
    }
    else {
        formContext.ui.tabs.get("tab_CapacityandUtilization").setVisible(false);
        formContext.ui.tabs.get("tab_EnrolmentProjections").setVisible(false);
        formContext.ui.tabs.get("tab_PastEnrolmentProjections").setVisible(false);
    }
    //tab_CapacityandUtilization
    //tab_EnrolmentProjections
    //tab_PastEnrolmentProjections
    //formContext.ui.tabs.get("tab_general").sections.get("section_designcapacity").setVisible(false);
}