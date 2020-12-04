"use strict";

var CAPS = CAPS || {};
CAPS.ProjectTracker = CAPS.ProjectTracker || {};

/**
 * Function called by Submit ribbon button, this changes the status to submitted
 * @param {any} primaryControl primary control
 */
CAPS.ProjectTracker.Closure = function (primaryControl) {
    debugger;
    var formContext = primaryControl;

    var recordId = formContext.data.entity.getId().replace("{", "").replace("}", "");

    var data =
    {
        "caps_Project@odata.bind": "/caps_projecttrackers(" + recordId + ")"
    }

    // create account record
    Xrm.WebApi.createRecord("caps_projectclosure", data).then(
        function success(result) {
            debugger;
            //console.log("Account created with ID: " + result.id);
            Xrm.Navigation.navigateTo({ pageType: "entityrecord", entityName: "caps_projectclosure", formType: 2, entityId: result.id}, { target: 2, position: 1, width: { value: 95, unit: "%" } });
            // perform operations on record creation
        },
        function (error) {
            console.log(error.message);
            // handle error conditions
        }
    );

    
}

/**
 * Function to determine if the Submit button should be displayed.
 * @param {any} primaryControl primary control
 * @returns {bool} true if shown, otherwise false
 */
CAPS.ProjectTracker.ShowClosure = function (primaryControl) {
    //check that record is draft & user's roles
    var formContext = primaryControl;

    if (formContext.getAttribute("statecode").getValue() !== 0) {
        return false;
    }

    var userRoles = Xrm.Utility.getGlobalContext().userSettings.roles;

    var showButton = false;

    userRoles.forEach(function hasAppropriateRole(item, index) {
        if (item.name === "CAPS School District User") {
            showButton = true;
        }
    });

    return showButton;
}
