"use strict";

var CAPS = CAPS || {};
CAPS.ProjectClosure = CAPS.ProjectClosure || {};

const PROJECT_STATE = {
    DRAFT: 1,
    SUBMIT: 2,
    COMPLETE: 200870000
};

/**
 * Function called by Submit ribbon button, this changes the status to submitted
 * @param {any} primaryControl primary control
 */
CAPS.ProjectClosure.Submit = function (primaryControl) {
    var formContext = primaryControl;

    var confirmStrings = { text: "Do you wish to submit this project closeout? Click OK to continue.  ", title: "Confirm Submission" };
    var confirmOptions = { height: 200, width: 450 };
    Xrm.Navigation.openConfirmDialog(confirmStrings, confirmOptions).then(
        function (success) {
            if (success.confirmed) {

                if (formContext.getAttribute("statuscode").getValue() === PROJECT_STATE.DRAFT) {

                    formContext.getAttribute("statecode").setValue(1);
                    formContext.getAttribute("statuscode").setValue(PROJECT_STATE.SUBMIT);
                    formContext.data.entity.save();
                }

            }
        },
        function (error) {
            Xrm.Navigation.openErrorDialog({ message: error });
        }

    );
}

/**
 * Function to determine if the Submit button should be displayed.
 * @param {any} primaryControl primary control
 * @returns {bool} true if shown, otherwise false
 */
CAPS.ProjectClosure.ShowSubmit = function (primaryControl) {
    //check that record is draft & user's roles
    var formContext = primaryControl;

    if (formContext.getAttribute("statuscode").getValue() !== PROJECT_STATE.DRAFT) {
        return false;
    }

    var userRoles = Xrm.Utility.getGlobalContext().userSettings.roles;

    var showButton = false;

    userRoles.forEach(function hasAppropriateRole(item, index) {
        if (item.name === "CAPS School District Approver - Add On") {
            showButton = true;
        }
    });

    return showButton;
}

/**
 * Function called by Submit ribbon button, this changes the status to submitted
 * @param {any} primaryControl primary control
 */
CAPS.ProjectClosure.Unsubmit = function (primaryControl) {
    var formContext = primaryControl;

    var confirmStrings = { text: "Do you wish to unsubmit this project closure? Click OK to continue.  ", title: "Confirm Unsubmission" };
    var confirmOptions = { height: 200, width: 450 };
    Xrm.Navigation.openConfirmDialog(confirmStrings, confirmOptions).then(
        function (success) {
            if (success.confirmed) {

                if (formContext.getAttribute("statuscode").getValue() === PROJECT_STATE.SUBMIT) {

                    formContext.getAttribute("statecode").setValue(0);
                    formContext.getAttribute("statuscode").setValue(PROJECT_STATE.DRAFT);
                    formContext.data.entity.save();
                }

            }
        },
        function (error) {
            Xrm.Navigation.openErrorDialog({ message: error });
        }

    );
}

/**
 * Function to determine if the Submit button should be displayed.
 * @param {any} primaryControl primary control
 * @returns {bool} true if shown, otherwise false
 */
CAPS.ProjectClosure.ShowUnsubmit = function (primaryControl) {
    debugger;
    //check that record is draft & user's roles
    var formContext = primaryControl;

    if (formContext.getAttribute("statuscode").getValue() !== PROJECT_STATE.SUBMIT) {
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

/**
 * Function called by Submit ribbon button, this changes the status to submitted
 * @param {any} primaryControl primary control
 */
CAPS.ProjectClosure.Complete = function (primaryControl) {
    var formContext = primaryControl;

    var confirmStrings = { text: "Do you wish to complete this project closure? Click OK to continue.  ", title: "Confirm Completion" };
    var confirmOptions = { height: 200, width: 450 };
    Xrm.Navigation.openConfirmDialog(confirmStrings, confirmOptions).then(
        function (success) {
            if (success.confirmed) {

                if (formContext.getAttribute("statuscode").getValue() === PROJECT_STATE.SUBMIT) {

                    formContext.getAttribute("statecode").setValue(1);
                    formContext.getAttribute("statuscode").setValue(PROJECT_STATE.COMPLETE);
                    formContext.data.entity.save();
                }

            }
        },
        function (error) {
            Xrm.Navigation.openErrorDialog({ message: error });
        }

    );
}

/**
 * Function to determine if the Submit button should be displayed.
 * @param {any} primaryControl primary control
 * @returns {bool} true if shown, otherwise false
 */
CAPS.ProjectClosure.ShowComplete = function (primaryControl) {
    debugger;
    //check that record is draft & user's roles
    var formContext = primaryControl;

    if (formContext.getAttribute("statuscode").getValue() !== PROJECT_STATE.SUBMIT) {
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