"use strict";

var CAPS = CAPS || {};
CAPS.Batch = CAPS.Batch || {};

//Ribbon ShowHide logic
CAPS.Batch.ShowSubmit = function (primaryControl) {
    var formContext = primaryControl;
    var showButton = false;

    var submissionDeadline = formContext.getAttribute("caps_submissiondeadline");
    if (submissionDeadline) {
        var submissionDeadlineDate = submissionDeadline.getValue();

        if (submissionDeadlineDate !== null) {
            const now = new Date();
            const today = new Date(now.getFullYear(), now.getMonth(), now.getDate());
            const deadlineDate = new Date(submissionDeadlineDate);
            const deadline = new Date(deadlineDate.getFullYear(), deadlineDate.getMonth(), deadlineDate.getDate());

            showButton = today <= deadline;
        }
    }

    // Check user roles synchronously
    var userRoles = Xrm.Utility.getGlobalContext().userSettings.roles;
    var canSubmitBasedOnRole = false;
    userRoles.forEach(function (role) {
        if (role.name === "CAPS CMB Finance Unit - Add On") {
            canSubmitBasedOnRole = true;
        }
    });

    // Return true only if both conditions are met
    return showButton && canSubmitBasedOnRole;
};

// Ribbon ShowHide logic
CAPS.Batch.ShowCancel = function (primaryControl) {
    var formContext = primaryControl;
    var showButton = false;

    var statecode = formContext.getAttribute("statecode");
    if (statecode) {
        var statecodevalue = statecode.getValue();
        if (statecodevalue === 0) {  
            showButton = true;
        }
    }

    // Check user roles synchronously
    var userRoles = Xrm.Utility.getGlobalContext().userSettings.roles;
    var canSubmitBasedOnRole = false;
    userRoles.forEach(function (role) {
        if (role.name === "CAPS CMB Finance Unit - Add On") {
            canSubmitBasedOnRole = true;
        }
    });

    // Return true only if both conditions are met
    return showButton && canSubmitBasedOnRole;
};

// Function to check if the deadline is past and manage form notifications. This Function is used on Batch form.
CAPS.Batch.isDeadlinePast = function (executionContext) {
    var formContext = executionContext.getFormContext();
    var submissionDeadline = formContext.getAttribute("caps_submissiondeadline");
    var statecode = formContext.getAttribute("statecode").getValue();

    if (submissionDeadline) {
        var submissionDeadlineDate = submissionDeadline.getValue();

        if (submissionDeadlineDate !== null) {
            const now = new Date();
            const today = new Date(now.getFullYear(), now.getMonth(), now.getDate());

            const deadlineDate = new Date(submissionDeadlineDate);
            const deadline = new Date(deadlineDate.getFullYear(), deadlineDate.getMonth(), deadlineDate.getDate());

            // Compare the two dates
            if (today > deadline && statecode == 0) {
                formContext.ui.setFormNotification("The submission deadline is past. Extend the deadline or create a new batch.", "ERROR", "deadlinePast");
                return true;
            }
            formContext.ui.clearFormNotification("deadlinePast");
            return false;
        }
    }
}
