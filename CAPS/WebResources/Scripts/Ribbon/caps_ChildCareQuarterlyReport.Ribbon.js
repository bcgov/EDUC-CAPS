"use strict";

var CAPS = CAPS || {};
CAPS.ChildCareQuarterlyReport = CAPS.ChildCareQuarterlyReport || {
   
};

CAPS.ChildCareQuarterlyReport.ShowHideActivateDeactivate = function () {
    
    var userRoles = Xrm.Utility.getGlobalContext().userSettings.roles;
    var showButton = false;
    userRoles.forEach(function hasSuperUserRole(item, index) {
        if (item.name === "CAPS CMB Super User - Add On") {
            showButton = true;
        }
    });

    return showButton;
};

CAPS.ChildCareQuarterlyReport.ShowSubmitButton = function (primaryControl) {
    var formContext = primaryControl;
    var userRoles = Xrm.Utility.getGlobalContext().userSettings.roles;
    var statusReason = formContext.getAttribute("statuscode").getValue();
    var showButton = false;
    userRoles.forEach(function hasSuperUserRole(item, index) {
        //User has school district role and status reason is draft
        if (item.name === "CAPS School District User" && statusReason === 1) {
            showButton = true;
        }
    });

    return showButton;
};

CAPS.ChildCareQuarterlyReport.Submit = function (primaryControl) {
    var formContext = primaryControl;
    //Set value to Submit
    formContext.getAttribute("statuscode").setValue(2);
    formContext.getAttribute("statecode").setValue(1);
    formContext.data.entity.save();
}

CAPS.ChildCareQuarterlyReport.ShowUnSubmitButton = function (primaryControl) {
    var formContext = primaryControl;
    var userRoles = Xrm.Utility.getGlobalContext().userSettings.roles;
    var statusReason = formContext.getAttribute("statuscode").getValue();
    var showButton = false;
    userRoles.forEach(function hasSuperUserRole(item, index) {
        //User has Ministry user role and status reason is Submitted
        if (item.name === "CAPS CMB User" && statusReason === 2) {
            showButton = true;
        }
    });

    return showButton;
};

CAPS.ChildCareQuarterlyReport.UnSubmit = function (primaryControl) {
    var formContext = primaryControl;
    //Set value to Draft
    formContext.getAttribute("statuscode").setValue(1);
    formContext.getAttribute("statecode").setValue(0);
    formContext.data.entity.save();
}