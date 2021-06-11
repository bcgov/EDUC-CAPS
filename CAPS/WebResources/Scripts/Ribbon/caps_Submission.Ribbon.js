"use strict";

var CAPS = CAPS || {};
CAPS.Submission = CAPS.Submission || {
    GLOBAL_FORM_CONTEXT: null
};

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
        if (item.name === "CAPS CMB User") {
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
    if (!(formContext.getAttribute("statuscode").getValue() == 1
        || formContext.getAttribute("statuscode").getValue() == 2)) {
        return false;
    }
    var userRoles = Xrm.Utility.getGlobalContext().userSettings.roles;

    var showButton = false;

    userRoles.forEach(function hasFinancialDirectorRole(item, index) {
        if (item.name === "CAPS CMB Finance Unit - Add On" || item.name === "CAPS CMB Super User - Add On") {
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
    CAPS.Submission.GLOBAL_FORM_CONTEXT = formContext;

    var globalContext = Xrm.Utility.getGlobalContext();
    var clientUrl = globalContext.getClientUrl();

    var webResource = '/caps_/Apps/ReasonForCancellation.html';

    Alert.showWebResource(webResource, 500, 275, "Reason for Cancellation", [
        new Alert.Button("Confirm", CAPS.Submission.ReasonForCancellationResult, true, true),
        new Alert.Button("Discard")
    ], clientUrl, true, null);
}

/*
Called on confirmation of Cancellation.  This function updates the record with the reason for cancellation and changes the status to cancelled.
*/
CAPS.Submission.ReasonForCancellationResult = function () {
    var formContext = CAPS.Submission.GLOBAL_FORM_CONTEXT;

    var validationResult = Alert.getIFrameWindow().validate();

    if (validationResult) {
        //update hidden field to trigger flow
        formContext.getControl("caps_reasonforcancellation").setDisabled(false);
        formContext.getAttribute("caps_reasonforcancellation").setValue(validationResult);
        formContext.getAttribute("statuscode").setValue(100000001);
        formContext.getAttribute("statecode").setValue(1);
        formContext.data.entity.save();

        //Close Popup
        Alert.hide();
    }

}

/**
 * Function to determine when the Cancel button should be shown.
 * @param {any} primaryControl primary control
 * @returns {boolean} true if should be shown, otherwise false.
 */
CAPS.Submission.ShowBulkCancel = function (primaryContext) {
    var formContext = primaryContext;
    var userRoles = Xrm.Utility.getGlobalContext().userSettings.roles;

    var showButton = false;

    //Hide if not AFG
    if (formContext.getAttribute("caps_callforsubmissiontype").getValue() != 100000002)
    {
        return false;

    }

    userRoles.forEach(function hasFinancialDirectorRole(item, index) {
        if (item.name === "CAPS CMB Finance Unit - Add On" || item.name === "CAPS CMB Super User - Add On") {
            showButton = true;
        }
    });

    return showButton;
}

/**
 * Function to cancel the project request record.
 * @param {any} primaryControl primary control
 */
CAPS.Submission.BulkCancel = function (selectedControlIds, selectedControl) {
    //call action
    var promises = [];
    selectedControlIds.forEach((record) => {
        let req = {};
        req.getMetadata = function () {
            return {
                boundParameter: "entity",
                operationType: 0,
                operationName: "caps_CancelSubmission",
                parameterTypes: {
                    "entity": {
                        "typeName": "mscrm.caps_submission",
                        "structuralProperty": 5
                    }
                }
            }
        };
        req.entity = { entityType: "caps_submission", id: record };
        promises.push(Xrm.WebApi.online.execute(req));
    });

    Promise.all(promises).then(
    function (results) {
        selectedControl.refresh();
    }
    , function (error) {
        console.log(error);
    }
);
}

/*
Function to check if the current user has CAPS CMB User Role.
*/
CAPS.Submission.ShowActivate = function (primaryContext) {
    var formContext = primaryContext;
    //Hide if not AFG
    if (formContext.getAttribute("caps_callforsubmissiontype").getValue() != 100000002) {
        return false;

    }

    var userRoles = Xrm.Utility.getGlobalContext().userSettings.roles;

    var showButton = false;

    userRoles.forEach(function hasFinancialDirectorRole(item, index) {
        if (item.name === "CAPS CMB Super User - Add On") {
            showButton = true;
        }
    });

    return showButton;
}