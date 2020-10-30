"use strict";

var CAPS = CAPS || {};
CAPS.Submission = CAPS.Submission || {};

/**
 * Function to determine when the Unsubmit button should be shown.
 * @param {any} primaryControl primary control
 * @returns {boolean} true if should be shown, otherwise false.
 */
CAPS.Submission.ShowUnsubmit = function (primaryControl) {
    var formContext = primaryControl;

    //TODO: check current state of capital plan, if submitted (2)
    if (formContext.getAttribute("statuscode").getValue() !== 2) {
        return false;
    }

    var userRoles = Xrm.Utility.getGlobalContext().userSettings.roles;

    var showButton = false;

    userRoles.forEach(function hasFinancialDirectorRole(item, index) {
        if (item.name === "CAPS Ministry User") {
            showButton = true;
        }
    });

    return showButton;
}

/**
 * Function to determine when the Cancel button should be shown.
 * @param {any} primaryControl primary control
 * @returns {boolean} true if should be shown, otherwise false.
 */
CAPS.Submission.ShowCancel = function (primaryControl) {
    var formContext = primaryControl;

    //TODO: Check current state of capital plan
    if (!(formContext.getAttribute("statuscode").getValue() == 2
        || formContext.getAttribute("statuscode").getValue() == 1)) {
        return false;
    }
    var userRoles = Xrm.Utility.getGlobalContext().userSettings.roles;

    var showButton = false;

    userRoles.forEach(function hasFinancialDirectorRole(item, index) {
        if (item.name === "CAPS Financial Director Ministry User - Add On") {
            showButton = true;
        }
    });

    return showButton;
}

/**
 * Function to unsubmit the project request record.
 * @param {any} primaryControl primary control
 */
CAPS.Submission.Unsubmit = function (primaryControl) {
    var formContext = primaryControl;
    //Change status to DRAFT
    let confirmStrings = { text: "This will unsubmit the captial plan.  Click OK to continue or Cancel to exit.", title: "Unsubmit Confirmation" };
    let confirmOptions = { height: 200, width: 450 };
    Xrm.Navigation.openConfirmDialog(confirmStrings, confirmOptions).then(
        function (success) {
            if (success.confirmed) {
                debugger;
                formContext.getAttribute("statecode").setValue(0);
                formContext.getAttribute("statuscode").setValue(1);
                formContext.data.entity.save();
            }

        });
}

/**
 * Function to cancel the project request record.
 * @param {any} primaryControl primary control
 */
CAPS.Submission.Cancel = function (primaryControl) {
    var formContext = primaryControl;
    //Change status to DRAFT
    let confirmStrings = { text: "This will cancel the captial plan.  Click OK to continue or Cancel to exit.", title: "Cancel Confirmation" };
    let confirmOptions = { height: 200, width: 450 };
    Xrm.Navigation.openConfirmDialog(confirmStrings, confirmOptions).then(
        function (success) {
            if (success.confirmed) {
                debugger;
                formContext.getAttribute("statuscode").setValue(100000001);
                formContext.data.entity.save();
            }

        });
}