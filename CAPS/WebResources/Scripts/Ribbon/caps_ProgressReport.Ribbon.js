"use strict";

var CAPS = CAPS || {};
CAPS.ProgressReport = CAPS.ProgressReport || {};

const PROJECT_STATE = {
    DRAFT: 1,
    SUBMITTED: 2
};

/**
 * Function to determine when the Unsubmit button should be shown.
 * @param {any} primaryControl primary control
 * @returns {boolean} true if should be shown, otherwise false.
 */
CAPS.ProgressReport.ShowUnsubmit = function (primaryControl) {
    var formContext = primaryControl;

    //Check current state of capital plan, if submitted
    if (formContext.getAttribute("statuscode").getValue() !== PROJECT_STATE.SUBMITTED) {
        return false;
    }

    // Check caps_UseNewForm flag
    var useNewFormAttribute = formContext.getAttribute("caps_usenewform");
    if (useNewFormAttribute != null) {
        var useNewForm = useNewFormAttribute.getValue();
        if (useNewForm != true) {
            return false; 
        }
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
CAPS.ProgressReport.ShowSubmit = function (primaryControl) {
    var formContext = primaryControl;

    //TODO: Check current state of capital plan
    if (formContext.getAttribute("statuscode").getValue() != PROJECT_STATE.DRAFT) {
        return false;
    }
    var userRoles = Xrm.Utility.getGlobalContext().userSettings.roles;

    var showButton = false;

    userRoles.forEach(function hasFinancialDirectorRole(item, index) {
        if (item.name === "CAPS School District User") {
            showButton = true;
        }
    });

    return showButton;
}

/**
 * Function to unsubmit the progress report record.
 * @param {any} primaryControl primary control
 */
CAPS.ProgressReport.Unsubmit = function (primaryControl) {
    var formContext = primaryControl;
    //Change status to DRAFT
    let confirmStrings = { text: "This will unsubmit the progress report.  Click OK to continue or Cancel to exit.", title: "Unsubmit Confirmation" };
    let confirmOptions = { height: 200, width: 450 };
    Xrm.Navigation.openConfirmDialog(confirmStrings, confirmOptions).then(
        function (success) {
            if (success.confirmed) {
                debugger;
                formContext.getAttribute("statecode").setValue(0);
                formContext.getAttribute("statuscode").setValue(PROJECT_STATE.DRAFT);
                formContext.data.entity.save();
            }

        });
}

/**
 * Function to submit the progress report record.
 * @param {any} primaryControl primary control
 */
CAPS.ProgressReport.Submit = function (primaryControl) {
    var formContext = primaryControl;
    // Call Validation Method at Form JavaScript.  
    //CAPS.ProgressReport.ValidateStatusComment(null, formContext);
    // Check for Legacy Form
    var useNewFormAttribute = formContext.getAttribute("caps_usenewform");
    if (useNewFormAttribute == null) {
        return; // Ignore if the attribute cannot be found.
    }
    var useNewForm = useNewFormAttribute.getValue();
    if (useNewForm) {

        CAPS.ProgressReport.UpdateProjectBudgetValues(null, formContext);
        CAPS.ProgressReport.ValidateOccupancyDate(null, formContext, true);
    }
    else {
        CAPS.ProgressReport.ConfirmSubmit(formContext, null);
    }
    /* Move this segment to its own method so it can be callback from Validation method in the form.
    //Change status to SUBMITTED
    let confirmStrings = { text: "This will submit the progress report.  Click OK to continue or Cancel to exit.", title: "Submit Confirmation" };
    let confirmOptions = { height: 200, width: 450 };
    Xrm.Navigation.openConfirmDialog(confirmStrings, confirmOptions).then(
        function (success) {
            if (success.confirmed) {
                formContext.getAttribute("statecode").setValue(1);
                formContext.getAttribute("statuscode").setValue(PROJECT_STATE.SUBMITTED);
                formContext.data.entity.save();
            }

        });
    */
};

CAPS.ProgressReport.ConfirmSubmit = function (formContext, additionalMessage) {

    var baseMessage = "This will submit the progress report.  Click OK to continue or Cancel to exit.";
    var message = baseMessage;
    if (additionalMessage != null && additionalMessage != "") {
        message = additionalMessage + "\n\r⠀\n\r" + baseMessage;
    }

    //Change status to SUBMITTED
    let confirmStrings = { text: message, title: "Submit Confirmation" };
    let confirmOptions = { height: 200, width: 450 };
    Xrm.Navigation.openConfirmDialog(confirmStrings, confirmOptions).then(
        function (success) {
            if (success.confirmed) {
                formContext.getAttribute("statecode").setValue(1);
                formContext.getAttribute("statuscode").setValue(PROJECT_STATE.SUBMITTED);
                formContext.data.entity.save();
            }

        });
}

