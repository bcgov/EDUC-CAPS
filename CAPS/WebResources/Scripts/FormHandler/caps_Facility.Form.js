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
}

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
        formContext.ui.setFormNotification('You have Close Date and/or Close Description set on an open facility. Either close this facility record or clear those closure fields.', 'WARNING', STATUS_MISSMATCH_NOTIFICATION);
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