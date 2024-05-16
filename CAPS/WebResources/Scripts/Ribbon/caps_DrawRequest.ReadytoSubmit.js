"use strict";

var CAPS = CAPS || {};
CAPS.DrawRequest = CAPS.DrawRequest || {};

// Ribbon ShowHide logic based on state code and user role
CAPS.DrawRequest.ShowReadytoSubmitButton = async function (primaryControl) {
    var formContext = primaryControl;
    if (formContext.getAttribute("statecode").getValue() !== 0 || formContext.getAttribute("caps_remainingdrawrequestbalance").getValue() < 0) {
        return false;
    }

    var userSettings = Xrm.Utility.getGlobalContext().userSettings;
    var currentUserId = userSettings.userId.replace('{', '').replace('}', '');
    var teamName = "CMB Expense Signing Authority";

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

    fetchXml = encodeURIComponent(fetchXml);

    //Logic below Await is in an event queue, and the above async code waits for a promise (the outcome of the team fetch query)
    var result = await Xrm.WebApi.retrieveMultipleRecords("team", "?fetchXml=" + fetchXml);
    // Check if the user is a member of the required team
    var canSubmit = result.entities.length > 0;

    // Date validation
    function validateDates() {
        var today = new Date();
        today.setHours(0, 0, 0, 0);

        var drawDate = formContext.getAttribute("caps_drawdate");
        var processDate = formContext.getAttribute("caps_processdate");

        //check dates and ensure the valid values exists before the function can be executed
        var drawDateValue = drawDate && drawDate.getValue() ? new Date(drawDate.getValue()) : null;
        var processDateValue = processDate && processDate.getValue() ? new Date(processDate.getValue()) : null;

        // Check if both dates are either today or in the future
        return (drawDateValue && drawDateValue >= today) && (processDateValue && processDateValue >= today);
    }

    return canSubmit && validateDates();
};

// Ribbon ShowHide logic based on state code and user role
CAPS.DrawRequest.ShowCancelButton = async function (primaryControl) {
    var formContext = primaryControl;
    if (formContext.getAttribute("statecode").getValue() !== 0 || formContext.getAttribute("caps_remainingdrawrequestbalance").getValue() < 0) {
        return false;
    }

    var userSettings = Xrm.Utility.getGlobalContext().userSettings;
    var currentUserId = userSettings.userId.replace('{', '').replace('}', '');
    var teamName = "CMB Expense Signing Authority";

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

    fetchXml = encodeURIComponent(fetchXml);

    var result = await Xrm.WebApi.retrieveMultipleRecords("team", "?fetchXml=" + fetchXml);
    var canCancel = result.entities.length > 0;

    return canCancel
};

// Function to trigger the action "Draw Request: Set Status to Ready to Submit"
CAPS.DrawRequest.Submit = function (primaryControl) {
    var formContext = primaryControl;
    let confirmStrings = { text: "This will make the draw request ready to submit. Click OK to continue or Cancel to exit.", title: "Submit Confirmation" };
    let confirmOptions = { height: 200, width: 450 };

    // Save the form
    Xrm.Page.data.save();

    Xrm.Navigation.openConfirmDialog(confirmStrings, confirmOptions).then(function (response) {
        if (response.confirmed) {
            var recordId = formContext.data.entity.getId().replace("{", "").replace("}", "");
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

            Xrm.Utility.showProgressIndicator("Submitting Draw Request...");

            Xrm.WebApi.online.execute(req).then(
                function (result) {
                    Xrm.Utility.closeProgressIndicator();
                    if (result.ok) {
                        formContext.data.refresh();
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
                }
            );
        }
    });
};

// Function to trigger the action "Draw Request: Cancel Draw Request"
CAPS.DrawRequest.Cancel = function (primaryControl) {
    var formContext = primaryControl;
    let confirmStrings = { text: "This will cancel the draw request(s). Click OK to continue or Cancel to exit.", title: "Submit Confirmation" };
    let confirmOptions = { height: 200, width: 450 };

    // Save the form
    Xrm.Page.data.save();

    Xrm.Navigation.openConfirmDialog(confirmStrings, confirmOptions).then(function (response) {
        if (response.confirmed) {
            var recordId = formContext.data.entity.getId().replace("{", "").replace("}", "");
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

            Xrm.Utility.showProgressIndicator("Cancelling Draw Request...");

            Xrm.WebApi.online.execute(req).then(
                function (result) {
                    Xrm.Utility.closeProgressIndicator();
                    if (result.ok) {
                        formContext.data.refresh();
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
                }
            );
        }
    });
};


/*
// Subgrid Ribbon ShowHide logic based on state code and user role
CAPS.DrawRequest.SubgridShowReadytoSubmitButton = async function (selectedItemIds) {

    if (!selectedItemIds || selectedItemIds.length === 0) {
        return false; 
    }

    var recordId = selectedItemIds[0]; 
    try {
        var entityReference = await Xrm.WebApi.retrieveRecord("entityLogicalName", recordId);

        var stateCode = entityReference.statecode;
        var balance = entityReference.caps_remainingdrawrequestbalance;

        if (stateCode !== 0 || balance < 0) {
            return false; 
        }

    var userSettings = Xrm.Utility.getGlobalContext().userSettings;
    var currentUserId = userSettings.userId.replace('{', '').replace('}', '');
    var teamName = "CMB Expense Signing Authority";

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
    var canSubmit = result.entities.length > 0;

    // Date validation function
    function validateDates() {
        var today = new Date();
        today.setHours(0, 0, 0, 0);

        var drawDate = entityReference.attributes.getByTitle("caps_drawdate").getValue();
        var processDate = entityReference.attributes.getByTitle("caps_processdate").getValue();

        drawDate = drawDate ? new Date(drawDate) : null;
        processDate = processDate ? new Date(processDate) : null;

        return (drawDate && drawDate >= today) && (processDate && processDate >= today);
    }

    return canSubmit && validateDates();
};

// Function to trigger the action "Draw Request: Set Status to Ready to Submit"
CAPS.DrawRequest.SubgridSubmit = function (selectedControl) {
    var formContext = selectedControl;
    let confirmStrings = { text: "This will make the draw request ready to submit. Click OK to continue or Cancel to exit.", title: "Submit Confirmation" };
    let confirmOptions = { height: 200, width: 450 };

    // Save the form
    Xrm.Page.data.save();

    Xrm.Navigation.openConfirmDialog(confirmStrings, confirmOptions).then(function (response) {
        if (response.confirmed) {
            var recordId = formContext.data.entity.getId().replace("{", "").replace("}", "");
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

            Xrm.Utility.showProgressIndicator("Submitting Draw Request...");

            Xrm.WebApi.online.execute(req).then(
                function (result) {
                    Xrm.Utility.closeProgressIndicator();
                    if (result.ok) {
                        formContext.data.refresh();
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
                }
            );
        }
    });
};
*/