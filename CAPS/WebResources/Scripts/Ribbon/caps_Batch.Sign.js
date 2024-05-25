"use strict";

var CAPS = CAPS || {};
CAPS.Batch = CAPS.Batch || {};

// Ribbon ShowHide logic
CAPS.Batch.ShowSign = async function (primaryControl) {
    var formContext = primaryControl;
    var batchId = getBatchId(formContext);

    if (!batchId) {
        console.error("Batch ID could not be retrieved.");
        return false; 
    }

    // User settings and current user ID retrieval
    var userSettings = Xrm.Utility.getGlobalContext().userSettings;
    var currentUserId = userSettings.userId.replace('{', '').replace('}', '');

    var teamName = "CMB Expense Signing Authority";

    // Fetch XML to check if the current user is part of the specified team
    var teamFetchXml = '<fetch>' +
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
    teamFetchXml = encodeURIComponent(teamFetchXml);

    try {
        var teamResult = await Xrm.WebApi.retrieveMultipleRecords("team", "?fetchXml=" + teamFetchXml);
        var isTeamMember = teamResult.entities.length > 0;

        if (!isTeamMember) {
            console.log("User is not a team member.");
            return false;
        }

        // Fetch XML to check for at least one draw request with statuscode = 2
        var drawRequestFetchXml = '<fetch aggregate="true">' +
            '<entity name="caps_drawrequest">' +
            '<attribute name="caps_drawrequestid" alias="count" aggregate="count" />' +
            '<filter type="and">' +
            '<condition attribute="statuscode" operator="eq" value="2" />' +
            '<condition attribute="caps_batch" operator="eq" value="' + batchId + '" />' +
            '</filter>' +
            '</entity>' +
            '</fetch>';
        drawRequestFetchXml = encodeURIComponent(drawRequestFetchXml);

        var drawRequestResult = await Xrm.WebApi.retrieveMultipleRecords("caps_drawrequest", "?fetchXml=" + drawRequestFetchXml);
        var hasEligibleDrawRequests = drawRequestResult.entities.length > 0 && parseInt(drawRequestResult.entities[0].count) > 0;

        return hasEligibleDrawRequests;
    } catch (error) {
        console.error("Error during ShowSign execution:", error);
        return false; // Return false in case of any error during the process
    }
};

function getBatchId(formContext) {
    if (formContext && formContext.data && formContext.data.entity) {
        return formContext.data.entity.getId().replace('{', '').replace('}', '');
    } else {
        console.error("Invalid formContext provided");
        return null;
    }
};

// Function to trigger the action "Batch: Mark as Pending Submission"
CAPS.Batch.Sign = function (primaryControl) {
    var formContext = primaryControl;
    let confirmStrings = { text: "This action will mark the batch to Pending Submission. Any ineligible draw requests will be cleared from the batch. Click OK to continue or Cancel to exit.", title: "Submit Confirmation" };
    let confirmOptions = { height: 200, width: 450 };

    Xrm.Navigation.openConfirmDialog(confirmStrings, confirmOptions).then(function (response) {
        if (response.confirmed) {
            var recordId = formContext.data.entity.getId().replace("{", "").replace("}", "");
            var req = {
                entity: {
                    entityType: "caps_batch",
                    id: recordId
                },
                getMetadata: function () {
                    return {
                        boundParameter: "entity",
                        operationType: 0,
                        operationName: "caps_BatchMarkasPendingSubmission",
                        parameterTypes: {
                            "entity": {
                                "typeName": "mscrm.caps_batch",
                                "structuralProperty": 5
                            }
                        }
                    };
                }
            };

            Xrm.Utility.showProgressIndicator("Completing Batch...");

            Xrm.WebApi.online.execute(req).then(
                function (result) {
                    Xrm.Utility.closeProgressIndicator();
                    if (result.ok) {
                        formContext.data.refresh();
                    } else {
                        var alertStrings = { confirmButtonLabel: "OK", text: "There was an error signing the batch. Please contact IT.", title: "Error" };
                        var alertOptions = { height: 350, width: 450 };
                        Xrm.Navigation.openAlertDialog(alertStrings, alertOptions);
                    }
                },
                function (error) {
                    Xrm.Utility.closeProgressIndicator();
                    var alertStrings = { confirmButtonLabel: "OK", text: "Signing the batch ran into an error. Details: " + error.message, title: "Error" };
                    var alertOptions = { height: 350, width: 450 };
                    Xrm.Navigation.openAlertDialog(alertStrings, alertOptions);
                }
            );
        }
    });
};