"use strict";

var CAPS = CAPS || {};
CAPS.Batch = CAPS.Batch || {};

//Ribbon ShowHide logic
CAPS.Batch.ShowComplete = function (primaryControl) {
    var formContext = primaryControl;
    var showButton = false;
    var statuscode = formContext.getAttribute("statuscode");

    if (statuscode) {
        var statuscodeValue = statuscode.getValue();

        if (statuscodeValue !== null && statuscodeValue === 2) {
            showButton = true;
        }
    }

    var userRoles = Xrm.Utility.getGlobalContext().userSettings.roles;
    var canSubmitBasedOnRole = false;
    userRoles.forEach(function (role) {
        if (role.name === "CAPS CMB Finance Unit - Add On") {
            canSubmitBasedOnRole = true;
        }
    });

    return showButton && canSubmitBasedOnRole;
};

// Function to trigger the action "Batch: Mark as Completed"
CAPS.Batch.Complete = function (primaryControl) {
    var formContext = primaryControl;
    let confirmStrings = { text: "This action will mark the batch and related submitted draw requests, and actual draws to Completed. Click OK to continue or Cancel to exit.", title: "Submit Confirmation" };
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
                        operationName: "caps_BatchMarkasCompleted",
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
                        var alertStrings = { confirmButtonLabel: "OK", text: "There was an error completing the batch. Please contact IT.", title: "Error" };
                        var alertOptions = { height: 350, width: 450 };
                        Xrm.Navigation.openAlertDialog(alertStrings, alertOptions);
                    }
                },
                function (error) {
                    Xrm.Utility.closeProgressIndicator();
                    var alertStrings = { confirmButtonLabel: "OK", text: "Completing the batch ran into an error. Details: " + error.message, title: "Error" };
                    var alertOptions = { height: 350, width: 450 };
                    Xrm.Navigation.openAlertDialog(alertStrings, alertOptions);
                }
            );
        }
    });
};