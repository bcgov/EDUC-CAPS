﻿//handles draw requests subgrid buttons on Batch form. Cancel, Ready to Submit, Set Back to Draft, Remove from Batch.

"use strict";

var CAPS = CAPS || {};
CAPS.DrawRequest = CAPS.DrawRequest || {};

CAPS.DrawRequest.SubgridShowReadytoSubmitButton = async function (selectedControl, selectedRecordIds) {

    // Process each selected record asynchronously to check conditions
    var allPromises = selectedRecordIds.map(async (recordId) => {
        try {
            // Retrieve specific fields from the 'caps_drawrequest' entity for each selected record
            var entityReference = await Xrm.WebApi.retrieveRecord("caps_drawrequest", recordId, "?$select=caps_drawdate,caps_processdate,statecode,caps_remainingdrawrequestbalance");
            console.log("Entity reference retrieved:", entityReference);

            var stateCode = entityReference.statecode;
            var balance = entityReference.caps_remainingdrawrequestbalance;
            // If the state code is not 0 or the balance is negative, log and return false

            if (stateCode !== 0 || balance < 0) {
                console.log("State code is not 0 or balance is below 0 for record ID:", recordId);
                return false;
            }

            var userSettings = Xrm.Utility.getGlobalContext().userSettings;
            var currentUserId = userSettings.userId.replace('{', '').replace('}', '');
            var teamName = "CMB Expense Signing Authority";

            var fetchXml = "<fetch>" +
                "<entity name='team'>" +
                "<attribute name='name' />" +
                "<link-entity name='teammembership' from='teamid' to='teamid' intersect='true'>" +
                "<link-entity name='systemuser' from='systemuserid' to='systemuserid' alias='user'>" +
                "<filter type='and'>" +
                "<condition attribute='systemuserid' operator='eq' value='" + currentUserId + "' />" +
                "</filter>" +
                "</link-entity>" +
                "</link-entity>" +
                "<filter type='and'>" +
                "<condition attribute='name' operator='eq' value='" + teamName + "' />" +
                "</filter>" +
                "</entity>" +
                "</fetch>";

            var encodedFetchXml = encodeURIComponent(fetchXml);
            var result = await Xrm.WebApi.retrieveMultipleRecords("team", "?fetchXml=" + encodedFetchXml);

            if (result.entities.length === 0) {
                console.log("User not in team for record ID:", recordId);
                return false;
            }

            return validateDates(entityReference);
        } catch (error) {
            console.error("Error in processing record ID:", recordId, error);
            return false;
        }
    });

    // Wait for all promises to resolve, and check if all returned true
    var results = await Promise.all(allPromises);
    return results.every(result => result === true);
};

function validateDates(entityReference) {
    var today = new Date();
    today.setHours(0, 0, 0, 0);

    var drawDateInput = entityReference.caps_drawdate;
    var processDateInput = entityReference.caps_processdate;

    var drawDate = new Date(drawDateInput + 'T00:00:00');
    var processDate = new Date(processDateInput + 'T00:00:00');

    return (drawDate >= today) && (processDate >= today);
};

// Function to trigger the action "Draw Request: Set Status to Ready to Submit"
CAPS.DrawRequest.SubgridSubmit = function (selectedControl, selectedRecordIds) {
    let confirmStrings = { text: "This will make the draw request ready to submit. Click OK to continue or Cancel to exit.", title: "Submit Confirmation" };
    let confirmOptions = { height: 200, width: 450 };

    Xrm.Navigation.openConfirmDialog(confirmStrings, confirmOptions).then(function (response) {
        if (response.confirmed) {
            let promises = selectedRecordIds.map(function (recordId) {
                return submitDrawRequest(recordId);
            });

            Promise.all(promises).then(() => {
                forceSaveBatchRecord();  
            });
        }
    });
};
        function submitDrawRequest(recordId) {
            var req = {
                entity: {
                    entityType: "caps_drawrequest",
                    id: recordId
                },
                getMetadata: function () {
                    return {
                        boundParameter: "entity",
                        operationType: 0,
                        operationName: "caps_DrawRequestSetStatustoReadytoSubmit",
                        parameterTypes: {
                            "entity": {
                                "typeName": "mscrm.caps_drawrequest",
                                "structuralProperty": 5
                            }
                        }
                    };
                }
            };

         return Xrm.WebApi.online.execute(req).then(
        function (result) {
            Xrm.Utility.closeProgressIndicator();
            if (result.ok) {
                console.log("Draw Request successfully submitted for ID:", recordId);
                refreshSubgrid('Subgrid_draft');
                refreshSubgrid('Subgrid_nondraft');
            } else {
                var alertStrings = { confirmButtonLabel: "OK", text: "There was an error submitting the draw request. Please contact IT.", title: "Error" };
                var alertOptions = { height: 350, width: 450 };
                Xrm.Navigation.openAlertDialog(alertStrings, alertOptions);
            }
        },
        function (error) {
            Xrm.Utility.closeProgressIndicator();
            var alertStrings = { confirmButtonLabel: "OK", text: "Submitting the draw request ran into an error. Details: " + error.message, title: "Error" };
            var alertOptions = { height: 350, width: 450 };
            Xrm.Navigation.openAlertDialog(alertStrings, alertOptions);
            console.error("Error when submitting Draw Request:", error);
        }
    );
};

// Show Remove from Batch and Cancel on Subgrids
CAPS.DrawRequest.SubgridShowCancelandRemoveButton = async function (selectedControl, selectedRecordIds) {
    var userSettings = Xrm.Utility.getGlobalContext().userSettings;
    var userRoles = userSettings.roles;
    var currentUserId = userSettings.userId.replace('{', '').replace('}', '');
    var teamName = "CMB Expense Signing Authority";
    var currentGridName = selectedControl._controlName;
    var canCancelRemoveBasedOnRole = false;
    var draftDrawRequest = "Subgrid_draft";
    var nonDraftDrawRequest = "Subgrid_nondraft";

    var parentStateCode = Xrm.Page.getAttribute("statecode").getValue(); 

    if (currentGridName === draftDrawRequest && parentStateCode === 0) {
        userRoles.forEach(function (role) {
            if (role.name === "CAPS CMB User") {
                canCancelRemoveBasedOnRole = true;
            }
        });
    } else if (currentGridName === nonDraftDrawRequest && parentStateCode === 0) {
        var fetchXml = '<fetch>' +
            '<entity name="team">' +
            '<attribute name="name" />' +
            '<link-entity name="teammembership" from="teamid" to="teamid" intersect="true">' +
            '<link-entity name="systemuser" from="systemuserid" to="systemuserid" alias="user">' +
            '<filter type="and">' +
            '<condition attribute="systemuserid" operator="eq" value="' + currentUserId + '" />' +
            '</filter>' +
            '</link-entity>' +
            '</link-entity>' +
            '<filter type="and">' +
            '<condition attribute="name" operator="eq" value="' + teamName + '" />' +
            '</filter>' +
            '</entity>' +
            '</fetch>';

        var encodedFetchXml = encodeURIComponent(fetchXml);
        var result = await Xrm.WebApi.retrieveMultipleRecords("team", "?fetchXml=" + encodedFetchXml);
        var isExpenseAuthorityTeamMember = result.entities.length > 0;

        userRoles.forEach(function (role) {
            if (role.name === "CAPS CMB Finance Unit - Add On" || isExpenseAuthorityTeamMember) {
                canCancelRemoveBasedOnRole = true;
            }
        });
    } else if (currentGridName === draftDrawRequest && parentStateCode !== 0) {
        userRoles.forEach(function (role) {
            if (role.name === "CAPS CMB Super User - Add On") {
                canCancelRemoveBasedOnRole = true;
            }
        });
    } else if (currentGridName === nonDraftDrawRequest && parentStateCode !== 0) {
        userRoles.forEach(function (role) {
            if (role.name === "CAPS CMB Super User - Add On") {
                canCancelRemoveBasedOnRole = true;
            }
        });
    }

    // Return true only if the condition is met
    return canCancelRemoveBasedOnRole;
};

// Show Set Back to Draft on non draft request subgrid
CAPS.DrawRequest.SubgridShowSetBacktoDraftButton = async function (selectedControl, selectedRecordIds) {
    var userSettings = Xrm.Utility.getGlobalContext().userSettings;
    var userRoles = userSettings.roles;
    var currentUserId = userSettings.userId.replace('{', '').replace('}', '');
    var teamName = "CMB Expense Signing Authority";
    var currentGridName = selectedControl._controlName;
    var canSetBacktoDraft = false;
    var draftDrawRequest = "Subgrid_draft";
    var nonDraftDrawRequest = "Subgrid_nondraft";

    var parentStateCode = Xrm.Page.getAttribute("statecode").getValue();

    if (currentGridName === nonDraftDrawRequest && parentStateCode === 0) {
        var fetchXml = '<fetch>' +
            '<entity name="team">' +
            '<attribute name="name" />' +
            '<link-entity name="teammembership" from="teamid" to="teamid" intersect="true">' +
            '<link-entity name="systemuser" from="systemuserid" to="systemuserid" alias="user">' +
            '<filter type="and">' +
            '<condition attribute="systemuserid" operator="eq" value="' + currentUserId + '" />' +
            '</filter>' +
            '</link-entity>' +
            '</link-entity>' +
            '<filter type="and">' +
            '<condition attribute="name" operator="eq" value="' + teamName + '" />' +
            '</filter>' +
            '</entity>' +
            '</fetch>';

        var encodedFetchXml = encodeURIComponent(fetchXml);
        var result = await Xrm.WebApi.retrieveMultipleRecords("team", "?fetchXml=" + encodedFetchXml);
        var isExpenseAuthorityTeamMember = result.entities.length > 0;

        userRoles.forEach(function (role) {
            if (role.name === "CAPS CMB Finance Unit - Add On" || isExpenseAuthorityTeamMember) {
                canSetBacktoDraft = true;
            }
        });
    } else if (currentGridName === nonDraftDrawRequest && parentStateCode !== 0) {
        userRoles.forEach(function (role) {
            if (role.name === "CAPS CMB Super User - Add On") {
                canSetBacktoDraft = true;
            }
        });
    }

    // Return true only if the condition is met
    return canSetBacktoDraft;
};

//remove from batch
CAPS.DrawRequest.SubgridRemovefromBatch = function (selectedControl, selectedRecordIds) {
    let confirmStrings = { text: "This will remove the draw request from this batch. Click OK to continue or Cancel to exit.", title: "Submit Confirmation" };
    let confirmOptions = { height: 200, width: 450 };

    Xrm.Navigation.openConfirmDialog(confirmStrings, confirmOptions).then(function (response) {
        if (response.confirmed) {
            let removalPromises = selectedRecordIds.map(recordId => removeDrawRequest(recordId));
            Promise.all(removalPromises).then(() => {
                forceSaveBatchRecord();
            });
        }
    });
};

        function removeDrawRequest(recordId) {
            var req = {
                entity: {
                    entityType: "caps_drawrequest",
                    id: recordId
                },
                "DrawRequestID": recordId,
                getMetadata: function () {
                    return {
                        boundParameter: "entity",
                        operationType: 0,
                        operationName: "caps_DrawRequestClearBatch",
                        parameterTypes: {
                            "entity": {
                                "typeName": "mscrm.caps_drawrequest",
                                "structuralProperty": 5
                            },
                            "DrawRequestID": {
                                "typeName": "Edm.Guid",
                                "structuralProperty": 1
                            }
                        }
                    };
                }
            };

    return Xrm.WebApi.online.execute(req).then(
        function (result) {
            Xrm.Utility.closeProgressIndicator();
            if (result.ok) {
                console.log("Draw Request successfully removed for ID:", recordId);
                refreshSubgrid('Subgrid_draft');
                refreshSubgrid('Subgrid_nondraft');
            } else {
                var alertStrings = { confirmButtonLabel: "OK", text: "There was an error removing the draw request. Please contact IT.", title: "Error" };
                var alertOptions = { height: 350, width: 450 };
                Xrm.Navigation.openAlertDialog(alertStrings, alertOptions);
            }
        },
        function (error) {
            Xrm.Utility.closeProgressIndicator();
            var alertStrings = { confirmButtonLabel: "OK", text: "Removing the draw request ran into an error. Details: " + error.message, title: "Error" };
            var alertOptions = { height: 350, width: 450 };
            Xrm.Navigation.openAlertDialog(alertStrings, alertOptions);
            console.error("Error when removing Draw Request:", error);
        }
    );
};


// Function to trigger the action "Draw Request: Set Status to Draft"
CAPS.DrawRequest.SubgridSetBacktoDraft = function (selectedControl, selectedRecordIds) {
    let confirmStrings = { text: "This will put the draw request(s) back to draft. Click OK to continue or Cancel to exit.", title: "Submit Confirmation" };
    let confirmOptions = { height: 200, width: 450 };

    Xrm.Navigation.openConfirmDialog(confirmStrings, confirmOptions).then(function (response) {
        if (response.confirmed) {
            let draftPromises = selectedRecordIds.map(recordId => setBackToDraft(recordId));
            Promise.all(draftPromises).then(() => {
                forceSaveBatchRecord();
            });
        }
    });
};

        function setBackToDraft(recordId) {
            var req = {
                entity: {
                    entityType: "caps_drawrequest",
                    id: recordId
                },
                getMetadata: function () {
                    return {
                        boundParameter: "entity",
                        operationType: 0,
                        operationName: "caps_DrawRequestSetStatustoDraft",
                        parameterTypes: {
                            "entity": {
                                "typeName": "mscrm.caps_drawrequest",
                                "structuralProperty": 5
                            }
                        }
                    };
                }
            };

    return Xrm.WebApi.online.execute(req).then(
        function (result) {
            Xrm.Utility.closeProgressIndicator();
            if (result.ok) {
                console.log("Draw Request successfully set back to draft for ID:", recordId);
                refreshSubgrid('Subgrid_draft');
                refreshSubgrid('Subgrid_nondraft');
            } else {
                var alertStrings = { confirmButtonLabel: "OK", text: "There was an error setting the draw request back to draft. Please contact IT.", title: "Error" };
                var alertOptions = { height: 350, width: 450 };
                Xrm.Navigation.openAlertDialog(alertStrings, alertOptions);
            }
        },
        function (error) {
            Xrm.Utility.closeProgressIndicator();
            var alertStrings = { confirmButtonLabel: "OK", text: "Setting the draw request back to draft ran into an error. Details: " + error.message, title: "Error" };
            var alertOptions = { height: 350, width: 450 };
            Xrm.Navigation.openAlertDialog(alertStrings, alertOptions);
            console.error("Error when setting Draw Request back to draft:", error);
        }
    );
};

/*
// Function to trigger the action "Draw Request: Cancel Request"
CAPS.DrawRequest.SubgridCancelRequest = function (selectedControl, selectedRecordIds) {
    let confirmStrings = { text: "This will cancel the draw request(s). Click OK to continue or Cancel to exit.", title: "Submit Confirmation" };
    let confirmOptions = { height: 200, width: 450 };

    Xrm.Navigation.openConfirmDialog(confirmStrings, confirmOptions).then(function (response) {
        if (response.confirmed) {
            let cancelPromises = selectedRecordIds.map(recordId => cancelRequest(recordId));
            Promise.all(cancelPromises).then(() => {
                forceSaveBatchRecord();
            });
        }
    });
};

    function cancelRequest(recordId) {
        var req = {
            entity: {
                entityType: "caps_drawrequest",
                id: recordId
            },
            getMetadata: function () {
                return {
                    boundParameter: "entity",
                    operationType: 0,
                    operationName: "caps_DrawRequestCancelDrawRequest",
                    parameterTypes: {
                        "entity": {
                            "typeName": "mscrm.caps_drawrequest",
                            "structuralProperty": 5
                        }
                    }
                };
            }
        };

    return Xrm.WebApi.online.execute(req).then(
        function (result) {
            Xrm.Utility.closeProgressIndicator();
            if (result.ok) {
                console.log("Draw Request successfully cancelled for ID:", recordId);
                refreshSubgrid('Subgrid_draft');
                refreshSubgrid('Subgrid_nondraft');
            } else {
                var alertStrings = { confirmButtonLabel: "OK", text: "There was an error cancelling the draw request. Please contact IT.", title: "Error" };
                var alertOptions = { height: 350, width: 450 };
                Xrm.Navigation.openAlertDialog(alertStrings, alertOptions);
            }
        },
        function (error) {
            Xrm.Utility.closeProgressIndicator();
            var alertStrings = { confirmButtonLabel: "OK", text: "Cancelling the draw request ran into an error. Details: " + error.message, title: "Error" };
            var alertOptions = { height: 350, width: 450 };
            Xrm.Navigation.openAlertDialog(alertStrings, alertOptions);
            console.error("Error when cancelling Draw Request:", error);
        }
    );
};
*/

function forceSaveBatchRecord() {
    var formContext = Xrm.Page;
    formContext.data.save().then(
        function () {
            console.log("Batch record forcefully saved after updating draw requests.");
        },
        function (error) {
            console.error("Failed to force save the batch record:", error);
        }
    );
}

function refreshSubgrid(subgridName) {
    var subgridControl = Xrm.Page.getControl(subgridName);
    if (subgridControl) {
        subgridControl.refresh();
    } else {
        console.error("Subgrid control not found: " + subgridName);
    }
}