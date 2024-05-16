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
    var statuscode = formContext.getAttribute("statuscode");
    if (statecode && statuscode) {
        var statecodevalue = statecode.getValue();
        var statuscodevalue = statuscode.getValue();
        if (statecodevalue === 0 || statuscodevalue === 200870004 || statuscodevalue === 2) {  
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

// Function to check if the deadline is past and manage form notifications.
// This function is used on the Batch form.
CAPS.Batch.isDeadlinePast = function (executionContext) {
    var formContext = executionContext.getFormContext();
    var submissionDeadline = formContext.getAttribute("caps_submissiondeadline");
    var drawDate = formContext.getAttribute("caps_drawdate");
    var statecode = formContext.getAttribute("statecode");

    var today = new Date();
    today.setHours(0, 0, 0, 0);

    var isValid = true;

    // Function to validate dates
    function validateDates() {
        var isValidLocal = true;
        var drawDateValue = drawDate.getValue() ? new Date(drawDate.getValue()) : null;
        var deadlineDateValue = submissionDeadline.getValue() ? new Date(submissionDeadline.getValue()) : null;

        if (drawDateValue && drawDateValue < today && statecode.getValue() === 0) {
            formContext.getControl("caps_drawdate").setNotification("Draw Date must be today or in the future.", "drawdate-error");
            isValidLocal = false;
        } else {
            formContext.getControl("caps_drawdate").clearNotification("drawdate-error");
        }

        if (deadlineDateValue && today > deadlineDateValue && statecode.getValue() === 0) {
            formContext.getControl("caps_submissiondeadline").setNotification("The Submission Deadline must be today or in the future.", "submissiondeadline-error");
            isValidLocal = false;
        } else {
            formContext.getControl("caps_submissiondeadline").clearNotification("submissiondeadline-error");
        }

        return isValidLocal;
    }

    // Initial validation call
    isValid = validateDates();

    // Set onChange events
    if (drawDate) {
        drawDate.addOnChange(validateDates);
    }
    if (submissionDeadline) {
        submissionDeadline.addOnChange(validateDates);
    }

    batchNotification(formContext);

    formContext.getAttribute("caps_numberofreadytosubmit").addOnChange(function () {
        batchNotification(formContext);
    });

    return isValid;
};

//Set fiscal year automatically based on submission deadline
CAPS.Batch.DrawDate_OnChange = function (executionContext) {
    var formContext = executionContext.getFormContext();
    var drawDate = formContext.getAttribute("caps_drawdate").getValue();
    var statecode = formContext.getAttribute("statecode").getValue();

    if (!drawDate) {
        formContext.getAttribute("caps_fiscalyear").setValue(null);
        return;
    }

    if (statecode === 0) {
        var drawDateValue = drawDate.toISOString();

        var fetchXml = "<fetch top='1'>" +
            "<entity name='edu_year'>" +
            "<attribute name='edu_yearid' />" +
            "<attribute name='edu_name' />" +
            "<filter type='and'>" +
            "<condition attribute='edu_startdate' operator='le' value='" + drawDateValue + "' />" +
            "<condition attribute='edu_enddate' operator='ge' value='" + drawDateValue + "' />" +
            "<condition attribute='statecode' operator='eq' value='0' />" +
            "<condition attribute='edu_type' operator='eq' value='757500000' />" +
            "</filter>" +
            "</entity>" +
            "</fetch>";

        var encodedFetchXml = encodeURIComponent(fetchXml);

        Xrm.WebApi.retrieveMultipleRecords("edu_year", "?fetchXml=" + encodedFetchXml).then(
            function success(result) {
                if (result.entities.length > 0) {
                    var firstRecord = result.entities[0];

                    formContext.getAttribute("caps_fiscalyear").setValue([
                        {
                            id: firstRecord.edu_yearid,
                            name: firstRecord.edu_name,
                            entityType: "edu_year"
                        }
                    ]);
                } else {
                    formContext.getAttribute("caps_fiscalyear").setValue(null);
                }
            },
            function (error) {
                console.error("Error retrieving year records: " + error.message);
            }
        );
    }
}

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

CAPS.Batch.monitorSubgridChanges = function (formContext) {
    var checkSubgrids = window.setInterval(async function () { 
        var subgridDraft = formContext.getControl("Subgrid_draft");
        var subgridNonDraft = formContext.getControl("Subgrid_nondraft");

        async function SubgridLoad(subgrid) {
            if (subgrid) {
                return new Promise(resolve => {
                    subgrid.addOnLoad(async function () {
                        await CAPS.Batch.isDeadlinePast(formContext);
                        await CAPS.Batch.ShowSubmit(formContext);
                        await CAPS.Batch.ShowCancel(formContext);
                        await batchNotification(formContext);
                        resolve(); 
                    });
                });
            }
            return Promise.resolve(); 
        }

        let loadPromises = [];
        if (subgridDraft) {
            loadPromises.push(SubgridLoad(subgridDraft));
        }
        if (subgridNonDraft) {
            loadPromises.push(SubgridLoad(subgridNonDraft));
        }

        Promise.all(loadPromises).then(() => {
            saveBatch(formContext); 
            window.clearInterval(checkSubgrids); 
        });
    }, 1000);

    formContext.ui.refreshRibbon();
};

function saveBatch(formContext) {
    formContext.data.save().then(function () {
        console.log("Batch record saved successfully.");
    }).catch(function (error) {
        console.error("Failed to save the batch record:", error);
    });
}
