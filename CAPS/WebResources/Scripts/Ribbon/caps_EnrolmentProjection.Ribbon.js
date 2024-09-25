"use strict";

var CAPS = CAPS || {};
CAPS.EnrolmentProjectionRibbon = CAPS.EnrolmentProjectionRibbon || {

};

/*
Function to check if the current user has CAPS SD User Role.
*/
CAPS.EnrolmentProjectionRibbon.IsSDUser = function () {
   
    var userRoles = Xrm.Utility.getGlobalContext().userSettings.roles;
    var showButton = true;
    userRoles.forEach(function hasSDRole(item, index) {
        if (item.name === "CAPS School District User") {
            showButton = false;
        }
    });

    return showButton;
}