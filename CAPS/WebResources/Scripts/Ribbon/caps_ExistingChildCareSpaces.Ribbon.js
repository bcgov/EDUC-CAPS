"use strict";

var CAPS = CAPS || {};
CAPS.ExistingChildCareSpaces = CAPS.ExistingChildCareSpaces || {};

CAPS.ExistingChildCareSpaces.ShowExistingChildCareSpacesButton = function (primaryControl) {
    
    var formContext = primaryControl;
    var projectRequestStatus = formContext.getAttribute("statuscode").getValue();
    //Draft
    if (projectRequestStatus === 1) return true;

    var globalContext = Xrm.Utility.getGlobalContext();
    var userRoles = globalContext.userSettings.roles;
    if (userRoles === null) return false;
    var hasRole = false;
   
    userRoles.forEach(function (item) {
        debugger;
        if (item.name.toLowerCase() === "caps school district user") {
            hasRole = true;
        }
       
    });

    return !hasRole;
    
}

