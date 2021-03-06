"use strict";

var CAPS = CAPS || {};
CAPS.ProjectTracker = CAPS.ProjectTracker || {};

/**
 * Function called by Close Project ribbon button, this creates a project closure record if one hasn't been created and opens it or the existing record in a modal.
 * @param {any} primaryControl primary control
 */
CAPS.ProjectTracker.Closure = function (primaryControl) {
    debugger;
    var formContext = primaryControl;

    var projectClosureField = formContext.getAttribute("caps_projectclosure").getValue();

    //Check if projectclosure existis
    if (projectClosureField != null) {
        var projectClosureId = projectClosureField[0].id.replace("{", "").replace("}", "");

        Xrm.Navigation.navigateTo({ pageType: "entityrecord", entityName: "caps_projectclosure", entityId:projectClosureId, formType: 2 }, { target: 2, position: 1, width: { value: 95, unit: "%" } });
    }
    else {
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
                Xrm.Navigation.navigateTo({ pageType: "entityrecord", entityName: "caps_projectclosure", formType: 2, entityId: result.id }, { target: 2, position: 1, width: { value: 95, unit: "%" } });
                formContext.ui.tabs.get("tab_projectclosure").setVisible(true);
                // perform operations on record creation
            },
            function (error) {
                console.log(error.message);
                // handle error conditions
            }
        );
    }

    
}

/**
 * Function to determine if the Closure Project button should be displayed.
 * @param {any} primaryControl primary control
 * @returns {bool} true if shown, otherwise false
 */
CAPS.ProjectTracker.ShowClosure = function (primaryControl) {
    //check that record is draft & user's roles
    var formContext = primaryControl;

    if (formContext.getAttribute("statecode").getValue() !== 0 || formContext.getAttribute("caps_showprogressreports").getValue() !== true) {
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

/***
 * Function called by Complete ribbon button, this changes the status to complete
 * @param {any} primaryControl primary control
*/
CAPS.ProjectTracker.Complete = function (primaryControl) {
    debugger;
    var formContext = primaryControl;

    var confirmStrings = { text: "Do you wish to complete this project? Click OK to continue.  ", title: "Confirm Completion" };
    var confirmOptions = { height: 200, width: 450 };
    Xrm.Navigation.openConfirmDialog(confirmStrings, confirmOptions).then(
        function (success) {
            if (success.confirmed) {
                formContext.getAttribute("statecode").setValue(1);
                formContext.getAttribute("statuscode").setValue(200870007);
                formContext.data.entity.save();
            }
        },
        function (error) {
            Xrm.Navigation.openErrorDialog({ message: error });
        }

    );

}

/**
 * Function to determine if the Complete button should be displayed.
 * @param {any} primaryControl primary control
 * @returns {bool} true if shown, otherwise false
 */
CAPS.ProjectTracker.ShowComplete = function (primaryControl) {
    //check that record is draft & user's roles
    var formContext = primaryControl;

    if (formContext.getAttribute("statecode").getValue() !== 0) {
        return false;
    }

    var userRoles = Xrm.Utility.getGlobalContext().userSettings.roles;

    var showButton = false;

    userRoles.forEach(function hasAppropriateRole(item, index) {
        if (item.name === "CAPS Ministry User") {
            showButton = true;
        }
    });

    return showButton;
}

/***
 * Function called by Cancel ribbon button, this changes the status to cancelled.
 * @param {any} primaryControl primary control
*/
CAPS.ProjectTracker.Cancel = function (primaryControl) {
    debugger;
    var formContext = primaryControl;

    var confirmStrings = { text: "Do you wish to cancel this project? Click OK to continue.  ", title: "Confirm Cancel" };
    var confirmOptions = { height: 200, width: 450 };
    Xrm.Navigation.openConfirmDialog(confirmStrings, confirmOptions).then(
        function (success) {
            if (success.confirmed) {
                formContext.getAttribute("statecode").setValue(1);
                formContext.getAttribute("statuscode").setValue(200870008);
                formContext.data.entity.save();
            }
        },
        function (error) {
            Xrm.Navigation.openErrorDialog({ message: error });
        }

    );

}

/**
 * Function to determine if the Cancel button should be displayed.
 * @param {any} primaryControl primary control
 * @returns {bool} true if shown, otherwise false
 */
CAPS.ProjectTracker.ShowCancel = function (primaryControl) {
    //check that record is draft & user's roles
    var formContext = primaryControl;

    if (formContext.getAttribute("statecode").getValue() !== 0) {
        return false;
    }

    var userRoles = Xrm.Utility.getGlobalContext().userSettings.roles;

    var showButton = false;

    userRoles.forEach(function hasAppropriateRole(item, index) {
        if (item.name === "CAPS Ministry User") {
            showButton = true;
        }
    });

    return showButton;
}


