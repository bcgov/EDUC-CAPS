"use strict";

var CAPS = CAPS || {};
CAPS.ChildCareProgressReport = CAPS.ChildCareProgressReport || {};

// Ribbon ShowHide logic based on status code (Report Status) and user role
CAPS.ChildCareProgressReport.ShowSubmitButton = async function (primaryControl) {
    var formContext = primaryControl;

    // Check if the status code is either 1 (Draft) or 714430003 (Unsubmitted)
    var statusCode = formContext.getAttribute("statuscode").getValue();
    if (statusCode !== 1 && statusCode !== 714430003) {
        return false;
    }

    var userSettings = Xrm.Utility.getGlobalContext().userSettings;
    var currentUserId = userSettings.userId.replace('{', '').replace('}', '');

    // FetchXml to check roles directly assigned to the user
    var fetchXmlDirect = '<fetch>' +
        '<entity name="role">' +
        '<attribute name="name" />' +
        '<filter type="or">' +
        '<condition attribute="name" operator="eq" value="CAPS CMB Super User - Add On" />' +
        '<condition attribute="name" operator="eq" value="CAPS School District User" />' +
        '</filter>' +
        '<link-entity name="systemuserroles" from="roleid" to="roleid" intersect="true">' +
        '<link-entity name="systemuser" from="systemuserid" to="systemuserid" alias="user">' +
        '<filter type="and">' +
        '<condition attribute="systemuserid" operator="eq" value="' + currentUserId + '" />' +
        '</filter>' +
        '</link-entity>' +
        '</link-entity>' +
        '</entity>' +
        '</fetch>';

    // FetchXml to check roles assigned to the user through teams
    var fetchXmlTeam = '<fetch>' +
        '<entity name="systemuser">' +
        '<attribute name="fullname" />' +
        '<filter type="and">' +
        '<condition attribute="systemuserid" operator="eq" value="' + currentUserId + '" />' +
        '</filter>' +
        '<link-entity name="teammembership" from="systemuserid" to="systemuserid" alias="tm">' +
        '<link-entity name="team" from="teamid" to="teamid" alias="t">' +
        '<link-entity name="teamroles" from="teamid" to="teamid" alias="tr">' +
        '<link-entity name="role" from="roleid" to="roleid" alias="r">' +
        '<filter type="or">' +
        '<condition attribute="name" operator="eq" value="CAPS CMB Super User - Add On" />' +
        '<condition attribute="name" operator="eq" value="CAPS School District User" />' +
        '</filter>' +
        '</link-entity>' +
        '</link-entity>' +
        '</link-entity>' +
        '</link-entity>' +
        '</entity>' +
        '</fetch>';

    fetchXmlDirect = encodeURIComponent(fetchXmlDirect);
    fetchXmlTeam = encodeURIComponent(fetchXmlTeam);

    // Execute both queries and combine results
    var directResult = await Xrm.WebApi.retrieveMultipleRecords("role", "?fetchXml=" + fetchXmlDirect);
    var teamResult = await Xrm.WebApi.retrieveMultipleRecords("systemuser", "?fetchXml=" + fetchXmlTeam);

    // If the user has a required role directly or through a team, show the button
    if (directResult.entities.length > 0 || teamResult.entities.length > 0) {
        return true;
    }
    return false;
};

// Ribbon ShowHide logic based on status code (Report Status) and user role
CAPS.ChildCareProgressReport.ShowUnsubmitButton = async function (primaryControl) {
    var formContext = primaryControl;

    // Check if the status code is 714430001 (Submitted)
    if (formContext.getAttribute("statuscode").getValue() !== 714430001) {
        return false;
    }

    var userSettings = Xrm.Utility.getGlobalContext().userSettings;
    var currentUserId = userSettings.userId.replace('{', '').replace('}', '');

    // FetchXml to check roles directly assigned to the user
    var fetchXmlDirect = '<fetch>' +
        '<entity name="role">' +
        '<attribute name="name" />' +
        '<filter type="or">' +
        '<condition attribute="name" operator="eq" value="CAPS CMB User" />' +
        '</filter>' +
        '<link-entity name="systemuserroles" from="roleid" to="roleid" intersect="true">' +
        '<link-entity name="systemuser" from="systemuserid" to="systemuserid" alias="user">' +
        '<filter type="and">' +
        '<condition attribute="systemuserid" operator="eq" value="' + currentUserId + '" />' +
        '</filter>' +
        '</link-entity>' +
        '</link-entity>' +
        '</entity>' +
        '</fetch>';

    // FetchXml to check roles assigned to the user through teams
    var fetchXmlTeam = '<fetch>' +
        '<entity name="systemuser">' +
        '<attribute name="fullname" />' +
        '<filter type="and">' +
        '<condition attribute="systemuserid" operator="eq" value="' + currentUserId + '" />' +
        '</filter>' +
        '<link-entity name="teammembership" from="systemuserid" to="systemuserid" alias="tm">' +
        '<link-entity name="team" from="teamid" to="teamid" alias="t">' +
        '<link-entity name="teamroles" from="teamid" to="teamid" alias="tr">' +
        '<link-entity name="role" from="roleid" to="roleid" alias="r">' +
        '<filter type="or">' +
        '<condition attribute="name" operator="eq" value="CAPS CMB User" />' +
        '</filter>' +
        '</link-entity>' +
        '</link-entity>' +
        '</link-entity>' +
        '</link-entity>' +
        '</entity>' +
        '</fetch>';

    fetchXmlDirect = encodeURIComponent(fetchXmlDirect);
    fetchXmlTeam = encodeURIComponent(fetchXmlTeam);

    // Execute both queries and combine results
    var directResult = await Xrm.WebApi.retrieveMultipleRecords("role", "?fetchXml=" + fetchXmlDirect);
    var teamResult = await Xrm.WebApi.retrieveMultipleRecords("systemuser", "?fetchXml=" + fetchXmlTeam);

    // If the user has a required role directly or through a team, show the button
    if (directResult.entities.length > 0 || teamResult.entities.length > 0) {
        return true;
    }
    return false;
};

// Ribbon ShowHide logic based on status code (Report Status) and user role
CAPS.ChildCareProgressReport.ShowCancelButton = async function (primaryControl) {
    var formContext = primaryControl;

    // Check if the status code is either 1 (Draft) or 714430003 (Unsubmitted)
    var statusCode = formContext.getAttribute("statuscode").getValue();
    if (statusCode !== 1 && statusCode !== 714430003) {
        return false;
    }

    var userSettings = Xrm.Utility.getGlobalContext().userSettings;
    var currentUserId = userSettings.userId.replace('{', '').replace('}', '');

    // FetchXml to check roles directly assigned to the user
    var fetchXmlDirect = '<fetch>' +
        '<entity name="role">' +
        '<attribute name="name" />' +
        '<filter type="or">' +
        '<condition attribute="name" operator="eq" value="CAPS CMB Super User - Add On" />' +
        '</filter>' +
        '<link-entity name="systemuserroles" from="roleid" to="roleid" intersect="true">' +
        '<link-entity name="systemuser" from="systemuserid" to="systemuserid" alias="user">' +
        '<filter type="and">' +
        '<condition attribute="systemuserid" operator="eq" value="' + currentUserId + '" />' +
        '</filter>' +
        '</link-entity>' +
        '</link-entity>' +
        '</entity>' +
        '</fetch>';

    // FetchXml to check roles assigned to the user through teams
    var fetchXmlTeam = '<fetch>' +
        '<entity name="systemuser">' +
        '<attribute name="fullname" />' +
        '<filter type="and">' +
        '<condition attribute="systemuserid" operator="eq" value="' + currentUserId + '" />' +
        '</filter>' +
        '<link-entity name="teammembership" from="systemuserid" to="systemuserid" alias="tm">' +
        '<link-entity name="team" from="teamid" to="teamid" alias="t">' +
        '<link-entity name="teamroles" from="teamid" to="teamid" alias="tr">' +
        '<link-entity name="role" from="roleid" to="roleid" alias="r">' +
        '<filter type="or">' +
        '<condition attribute="name" operator="eq" value="CAPS CMB Super User - Add On" />' +
        '</filter>' +
        '</link-entity>' +
        '</link-entity>' +
        '</link-entity>' +
        '</link-entity>' +
        '</entity>' +
        '</fetch>';

    fetchXmlDirect = encodeURIComponent(fetchXmlDirect);
    fetchXmlTeam = encodeURIComponent(fetchXmlTeam);

    // Execute both queries and combine results
    var directResult = await Xrm.WebApi.retrieveMultipleRecords("role", "?fetchXml=" + fetchXmlDirect);
    var teamResult = await Xrm.WebApi.retrieveMultipleRecords("systemuser", "?fetchXml=" + fetchXmlTeam);

    // If the user has a required role directly or through a team, show the button
    if (directResult.entities.length > 0 || teamResult.entities.length > 0) {
        return true;
    }
    return false;
};

// Function to trigger the action "Unsubmit Report"
CAPS.ChildCareProgressReport.UnsubmitReport = function (primaryControl) {
    var formContext = primaryControl;
    let confirmStrings = { text: "This will unsubmit the report and trigger the system to overwrite with the latest capacity numbers and license info in a moment. Click OK to continue or Cancel to exit.", title: "Unsubmit Confirmation" };
    let confirmOptions = { height: 200, width: 450 };

    // Save the form
    Xrm.Page.data.save();

    Xrm.Navigation.openConfirmDialog(confirmStrings, confirmOptions).then(function (response) {
        if (response.confirmed) {
            var recordId = formContext.data.entity.getId().replace("{", "").replace("}", "");
            var req = {
                entity: {
                    entityType: "caps_childcareprogressreport",
                    id: recordId
                },
                getMetadata: function () {
                    return {
                        boundParameter: "entity",
                        operationType: 0,
                        operationName: "caps_ChildCareProgressReportUnsubmitReport",
                        parameterTypes: {
                            "entity": {
                                "typeName": "mscrm.caps_childcareprogressreport",
                                "structuralProperty": 5
                            }
                        }
                    };
                }
            };

            Xrm.Utility.showProgressIndicator("Unsubmitting Report...");

            Xrm.WebApi.online.execute(req).then(
                function (result) {
                    Xrm.Utility.closeProgressIndicator();
                    if (result.ok) {
                        formContext.data.refresh();
                    } else {
                        var alertStrings = { confirmButtonLabel: "OK", text: "There was an error unsubmitting the report. Please contact IT.", title: "Error" };
                        var alertOptions = { height: 350, width: 450 };
                        Xrm.Navigation.openAlertDialog(alertStrings, alertOptions);
                    }
                },
                function (error) {
                    Xrm.Utility.closeProgressIndicator();
                    var alertStrings = { confirmButtonLabel: "OK", text: "Unsubmitting the report ran into an error. Details: " + error.message, title: "Error" };
                    var alertOptions = { height: 350, width: 450 };
                    Xrm.Navigation.openAlertDialog(alertStrings, alertOptions);
                }
            );
        }
    });
};

// Function to trigger the action "Submit Report"
CAPS.ChildCareProgressReport.SubmitReport = function (primaryControl) {
    var formContext = primaryControl;
    let confirmStrings = { text: "This will submit the report and you won't be able to modify the record. Click OK to continue or Cancel to exit.", title: "Submit Confirmation" };
    let confirmOptions = { height: 200, width: 450 };

    // Save the form
    Xrm.Page.data.save();

    Xrm.Navigation.openConfirmDialog(confirmStrings, confirmOptions).then(function (response) {
        if (response.confirmed) {
            var recordId = formContext.data.entity.getId().replace("{", "").replace("}", "");
            var req = {
                entity: {
                    entityType: "caps_childcareprogressreport",
                    id: recordId
                },
                getMetadata: function () {
                    return {
                        boundParameter: "entity",
                        operationType: 0,
                        operationName: "caps_ChildCareProgressReportSubmitReport",
                        parameterTypes: {
                            "entity": {
                                "typeName": "mscrm.caps_childcareprogressreport",
                                "structuralProperty": 5
                            }
                        }
                    };
                }
            };

            Xrm.Utility.showProgressIndicator("Submitting Report...");

            Xrm.WebApi.online.execute(req).then(
                function (result) {
                    Xrm.Utility.closeProgressIndicator();
                    if (result.ok) {
                        formContext.data.refresh();
                    } else {
                        var alertStrings = { confirmButtonLabel: "OK", text: "There was an error submitting the report. Please contact IT.", title: "Error" };
                        var alertOptions = { height: 350, width: 450 };
                        Xrm.Navigation.openAlertDialog(alertStrings, alertOptions);
                    }
                },
                function (error) {
                    Xrm.Utility.closeProgressIndicator();
                    var alertStrings = { confirmButtonLabel: "OK", text: "Submitting the report ran into an error. Details: " + error.message, title: "Error" };
                    var alertOptions = { height: 350, width: 450 };
                    Xrm.Navigation.openAlertDialog(alertStrings, alertOptions);
                }
            );
        }
    });
};

// Function to trigger the action "Cancel Report"
CAPS.ChildCareProgressReport.CancelReport = function (primaryControl) {
    var formContext = primaryControl;
    let confirmStrings = { text: "This will cancel the report and you will have to ask the MOE super user if you want to start a new report. Click OK to continue or Cancel to exit.", title: "Cancel Confirmation" };
    let confirmOptions = { height: 200, width: 450 };

    // Save the form
    Xrm.Page.data.save();

    Xrm.Navigation.openConfirmDialog(confirmStrings, confirmOptions).then(function (response) {
        if (response.confirmed) {
            var recordId = formContext.data.entity.getId().replace("{", "").replace("}", "");
            var req = {
                entity: {
                    entityType: "caps_childcareprogressreport",
                    id: recordId
                },
                getMetadata: function () {
                    return {
                        boundParameter: "entity",
                        operationType: 0,
                        operationName: "caps_ChildCareProgressReportCancelReport",
                        parameterTypes: {
                            "entity": {
                                "typeName": "mscrm.caps_childcareprogressreport",
                                "structuralProperty": 5
                            }
                        }
                    };
                }
            };

            Xrm.Utility.showProgressIndicator("Cancelling Report...");

            Xrm.WebApi.online.execute(req).then(
                function (result) {
                    Xrm.Utility.closeProgressIndicator();
                    if (result.ok) {
                        formContext.data.refresh();
                    } else {
                        var alertStrings = { confirmButtonLabel: "OK", text: "There was an error cancelling the report. Please contact IT.", title: "Error" };
                        var alertOptions = { height: 350, width: 450 };
                        Xrm.Navigation.openAlertDialog(alertStrings, alertOptions);
                    }
                },
                function (error) {
                    Xrm.Utility.closeProgressIndicator();
                    var alertStrings = { confirmButtonLabel: "OK", text: "Cancelling the report ran into an error. Details: " + error.message, title: "Error" };
                    var alertOptions = { height: 350, width: 450 };
                    Xrm.Navigation.openAlertDialog(alertStrings, alertOptions);
                }
            );
        }
    });
};