"use strict";

var CAPS = CAPS || {};
CAPS.ChildCareEnrolmentProjectionRibbon = CAPS.ChildCareEnrolmentProjectionRibbon || {
    
};

/*
Function to check if the current user has CAPS SD User Role.
*/
CAPS.ChildCareEnrolmentProjectionRibbon.IsSDUser = function () {
    
    var userRoles = Xrm.Utility.getGlobalContext().userSettings.roles;
    var showButton = true;
    userRoles.forEach(function hasSDRole(item, index) {
        if (item.name === "CAPS School District User") {
            showButton = false;
        }
    });
    return showButton;
    
}

