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

    batchNotification(formContext);

    formContext.getAttribute("caps_numberofreadytosubmit").addOnChange(function () {
        batchNotification(formContext);
    });

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
};

/*
function batchNotification(formContext) {
    var eligibleToSign = formContext.getAttribute("caps_numberofreadytosubmit");
    var statecode = formContext.getAttribute("statecode").getValue();

    if (statecode !== 0) {
        return false;
    }

    if (!eligibleToSign || !eligibleToSign.getValue() || eligibleToSign.getValue() === 0 && statecode === 0) {
        formContext.ui.setFormNotification("At least one draw request needs to be in Ready to Submit status to sign the batch", "WARNING", "ineligibleNotification");
    } else {
        formContext.ui.clearFormNotification("ineligibleNotification");
    }
};
*/

function batchNotification(formContext) {
    var statecode = formContext.getAttribute("statecode").getValue();

    if (statecode !== 0) {
        return false;
    }

    var batchId = formContext.data.entity.getId().replace('{', '').replace('}', '');

    var fetchXml = "<fetch top='1'>" +
        "<entity name='caps_drawrequest'>" +
        "<filter>" +
        "<condition attribute='caps_batch' operator='eq' value='" + batchId + "' />" +
        "<condition attribute='statuscode' operator='eq' value='2' />" +
        "</filter>" +
        "</entity>" +
        "</fetch>";

    Xrm.WebApi.retrieveMultipleRecords("caps_drawrequest", "?fetchXml=" + encodeURIComponent(fetchXml)).then(
        function success(result) {
            if (result.entities.length > 0) {
                formContext.ui.clearFormNotification("ineligibleNotification");
            } else {
                formContext.ui.setFormNotification("At least one draw request needs to be in Ready to Submit status to sign the batch", "WARNING", "ineligibleNotification");
            }
        },
        function error(error) {
            console.error("Error fetching data: ", error.message);
        }
    );
};

/*
CAPS.Batch.monitorSubgridChanges = function () {

    var checkSubgrid_draft = window.setInterval(function () {
        var subgridDraft = Xrm.Page.getControl("checkSubgrid_draft");
        if (subgridDraft != null) {
            subgridDraft.addOnLoad(function () {
                CAPS.Batch.isDeadlinePast({ getFormContext: function () { return Xrm.Page; } });
            });
            window.clearInterval(checkSubgrid_draft);
        }
    }, 1000);
};

CAPS.Batch.monitorSubgridChanges ();
*/

CAPS.Batch.monitorSubgridChanges = function () {
    var checkSubgrids = window.setInterval(function () {
        var subgridDraft = Xrm.Page.getControl("Subgrid_draft");
        var subgridNonDraft = Xrm.Page.getControl("Subgrid_nondraft");

        function SubgridLoad(subgrid) {
            if (subgrid) {
                subgrid.addOnLoad(function () {
                    var formContext = { getFormContext: function () { return Xrm.Page; } };
                    CAPS.Batch.isDeadlinePast(formContext);
                    CAPS.Batch.ShowSubmit(formContext.getFormContext());
                    CAPS.Batch.ShowCancel(formContext.getFormContext());
                });
            }
        }

        if (subgridDraft || subgridNonDraft) {
            SubgridLoad(subgridDraft);
            SubgridLoad(subgridNonDraft);
            window.clearInterval(checkSubgrids);
        }
    }, 1000);
};

CAPS.Batch.monitorSubgridChanges();


