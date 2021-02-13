"use strict";

var CAPS = CAPS || {};
CAPS.PRFSOption = CAPS.PRFSOption || {};

/**
 * This function triggers the Calculate Schedule B action
 * @param {any} primaryControl primary control
 */
CAPS.PRFSOption.CalculateScheduleB = function (primaryControl) {
    var formContext = primaryControl;
    debugger;

    //If dirty, then save and call again
    if (formContext.data.entity.getIsDirty() || formContext.ui.getFormType() === 1) {
        formContext.data.save({ saveMode: 1 }).then(function success(result) { CAPS.PRFSOption.CalculateScheduleB(primaryControl); });
    }
    else {
        var recordId = formContext.data.entity.getId().replace("{", "").replace("}", "");
        //call action
        var req = {};
        var target = { entityType: "caps_prfsalternativeoption", id: recordId };
        req.entity = target;

        req.getMetadata = function () {
            return {
                boundParameter: "entity",
                operationType: 0,
                operationName: "caps_TriggerScheduleBCalculation69523368a6d2ea11a813000d3af42496",
                parameterTypes: {
                    "entity": {
                        "typeName": "mscrm.caps_prfsalternativeoption",
                        "structuralProperty": 5
                    }
                }
            }
        };

        Xrm.WebApi.online.execute(req).then(
                        function (result) {
                            if (result.ok) {
                                return result.json().then(
                                    function (response) {
                                        debugger;
                                        //get error message
                                        if (response.ErrorMessage == null) {
                                            var alertStrings = { confirmButtonLabel: "OK", text: "The preliminary budget calculation has completed successfully.", title: "Preliminary Budget Result" };
                                            var alertOptions = { height: 120, width: 260 };
                                            Xrm.Navigation.openAlertDialog(alertStrings, alertOptions).then(
                                                function success(result) {
                                                    console.log("Alert dialog closed");
                                                    formContext.data.refresh();
                                                },
                                                function (error) {
                                                    console.log(error.message);
                                                }
                                            );
                                        }
                                        else {
                                            var alertStrings = { confirmButtonLabel: "OK", text: "The preliminary budget calculation ran into a problem. Details: " + response.ErrorMessage, title: "Preliminary Budget Result" };
                                            var alertOptions = { height: 120, width: 260 };
                                            Xrm.Navigation.openAlertDialog(alertStrings, alertOptions).then(
                                                function success(result) {
                                                    console.log("Alert dialog closed");
                                                },
                                                function (error) {
                                                    console.log(error.message);
                                                }
                                            );
                                        }
                                    });
                            }

                        },
            function (e) {

                var alertStrings = { confirmButtonLabel: "OK", text: "Schedule B failed. Details: " + e.message, title: "Schedule B Result" };
                var alertOptions = { height: 120, width: 260 };
                Xrm.Navigation.openAlertDialog(alertStrings, alertOptions).then(
                    function success(result) {
                        console.log("Alert dialog closed");
                    },
                    function (error) {
                        console.log(error.message);
                    }
                );
            }
        );
    }
}

/**
 * This function determines if the Calculate Schedule B button should be displayed
 * @param {any} primaryControl primary control
 * @return {bool} true if should be displayed, otherwise false
 */
CAPS.PRFSOption.ShowCalculateScheduleB = function (primaryControl) {
    debugger;
    var formContext = primaryControl;

    if (!(formContext.getAttribute("statuscode").getValue() == 1)) {
        return false;
    }

    return formContext.getAttribute("caps_requiresscheduleb").getValue();
}

/**
 * This function deactivates the record
 * @param {any} primaryControl primary control
 */
CAPS.PRFSOption.Deactivate = function (primaryControl) {
    var formContext = primaryControl;
    //Change status to DRAFT
    let confirmStrings = { text: "The selected PRFS Alternative Option record's status will be changed to Cancelled and cannot be undone.  Click OK to continue or Cancel to exit.", title: "Deactivate Confirmation" };
    let confirmOptions = { height: 200, width: 450 };
    Xrm.Navigation.openConfirmDialog(confirmStrings, confirmOptions).then(
        function (success) {
            if (success.confirmed) {
                debugger;
                formContext.getAttribute("statecode").setValue(1);
                formContext.getAttribute("statuscode").setValue(2);
                formContext.data.entity.save();
            }

        });
}

/**
 * This function determines if the Deactivate button should be displayed
 * @param {any} primaryControl primary control
 * @return {bool} true if should be displayed, otherwise false
 */
CAPS.PRFSOption.ShowDeactivate = function (primaryControl) {
    var formContext = primaryControl;

    //TODO: Check current state of capital plan
    if (!(formContext.getAttribute("statuscode").getValue() == 1)) {
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