"use strict";

var CAPS = CAPS || {};
CAPS.ChildCareActualEnrolment = CAPS.ChildCareActualEnrolment || {};

// ShowSubmit logic for "CAPS School District User" 
CAPS.ChildCareActualEnrolment.ShowSubmit = async function (primaryControl) {
    var formContext = primaryControl;
    var userSettings = Xrm.Utility.getGlobalContext().userSettings;
    var currentUserId = userSettings.userId.replace('{', '').replace('}', '');

    // Check if the statuscode is either 1 (Unsubmitted) 
    var statusCode = formContext.getAttribute("statuscode").getValue();
    if (statusCode !== 1) {
        return false;
    }

    // FetchXml to check roles directly assigned to the user
    var fetchXmlDirect = '<fetch>' +
        '<entity name="role">' +
        '<attribute name="name" />' +
        '<filter type="and">' +
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
        '<condition attribute="name" operator="eq" value="CAPS School District User" />' +
        '</filter>' +
        '</link-entity>' +
        '</link-entity>' +
        '</link-entity>' +
        '</link-entity>' +
        '</entity>' +
        '</fetch>';

    try {
        // Execute both queries
        var directResult = await Xrm.WebApi.retrieveMultipleRecords("role", "?fetchXml=" + fetchXmlDirect);
        var teamResult = await Xrm.WebApi.retrieveMultipleRecords("systemuser", "?fetchXml=" + fetchXmlTeam);

        if (directResult.entities.length > 0 || teamResult.entities.length > 0) {
            return true;
        }

        console.log("User does not have the required roles.");
        return false;
    } catch (error) {
        console.error("Error retrieving roles: " + error.message);
        return false;
    }
};

// ShowUnsubmit logic for "CAPS CMB User"
CAPS.ChildCareActualEnrolment.ShowUnsubmit = async function (primaryControl) {
    var formContext = primaryControl;
    var userSettings = Xrm.Utility.getGlobalContext().userSettings;
    var currentUserId = userSettings.userId.replace('{', '').replace('}', '');

    // Check if the statuscode is "Submitted" (746660001)
    var statusCode = formContext.getAttribute("statuscode").getValue();
    if (statusCode !== 746660001) {
        return false;
    }

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

    // Encode the FetchXml queries
    fetchXmlDirect = encodeURIComponent(fetchXmlDirect);
    fetchXmlTeam = encodeURIComponent(fetchXmlTeam);

    // Execute both queries
    var directResult = await Xrm.WebApi.retrieveMultipleRecords("role", "?fetchXml=" + fetchXmlDirect);
    var teamResult = await Xrm.WebApi.retrieveMultipleRecords("systemuser", "?fetchXml=" + fetchXmlTeam);

    // If the user has the required role directly or through a team, show the unsubmit button
    if (directResult.entities.length > 0 || teamResult.entities.length > 0) {
        return true;
    }
    return false;
};

// Function to trigger the action "Submit Child Care Actual Enrolment"
CAPS.ChildCareActualEnrolment.SubmitReport = function (primaryControl) {
    var formContext = primaryControl;
    let confirmStrings = { text: "This will submit the actual enrolment and you won't be able to modify the record. Click OK to continue or Cancel to exit.", title: "Submit Confirmation" };
    let confirmOptions = { height: 200, width: 450 };

    // Save the form
    Xrm.Page.data.save();

    Xrm.Navigation.openConfirmDialog(confirmStrings, confirmOptions).then(function (response) {
        if (response.confirmed) {
            var recordId = formContext.data.entity.getId().replace("{", "").replace("}", "");
            var req = {
                entity: {
                    entityType: "caps_childcareactualenrolment",
                    id: recordId
                },
                getMetadata: function () {
                    return {
                        boundParameter: "entity",
                        operationType: 0,
                        operationName: "caps_CCActualEnrolmentSubmit",
                        parameterTypes: {
                            "entity": {
                                "typeName": "mscrm.caps_childcareactualenrolment",
                                "structuralProperty": 5
                            }
                        }
                    };
                }
            };

            Xrm.Utility.showProgressIndicator("Submitting Enrolment...");

            Xrm.WebApi.online.execute(req).then(
                function (result) {
                    Xrm.Utility.closeProgressIndicator();
                    if (result.ok) {
                        formContext.data.refresh();
                    } else {
                        var alertStrings = { confirmButtonLabel: "OK", text: "There was an error submitting the enrolment. Please contact IT.", title: "Error" };
                        var alertOptions = { height: 350, width: 450 };
                        Xrm.Navigation.openAlertDialog(alertStrings, alertOptions);
                    }
                },
                function (error) {
                    Xrm.Utility.closeProgressIndicator();
                    var alertStrings = { confirmButtonLabel: "OK", text: "Submitting the enrolment ran into an error. Details: " + error.message, title: "Error" };
                    var alertOptions = { height: 350, width: 450 };
                    Xrm.Navigation.openAlertDialog(alertStrings, alertOptions);
                }
            );
        }
    });
};

// Function to trigger the action "Unsubmit Child Care Actual Enrolment"
CAPS.ChildCareActualEnrolment.UnsubmitReport = function (primaryControl) {
    var formContext = primaryControl;
    let confirmStrings = { text: "This will unsubmit the actual enrolment and become editable. Click OK to continue or Cancel to exit.", title: "Unsubmit Confirmation" };
    let confirmOptions = { height: 200, width: 450 };

    // Save the form
    Xrm.Page.data.save();

    Xrm.Navigation.openConfirmDialog(confirmStrings, confirmOptions).then(function (response) {
        if (response.confirmed) {
            var recordId = formContext.data.entity.getId().replace("{", "").replace("}", "");
            var req = {
                entity: {
                    entityType: "caps_childcareactualenrolment",
                    id: recordId
                },
                getMetadata: function () {
                    return {
                        boundParameter: "entity",
                        operationType: 0,
                        operationName: "caps_CCActualEnrolmentUnsubmit",
                        parameterTypes: {
                            "entity": {
                                "typeName": "mscrm.caps_childcareactualenrolment",
                                "structuralProperty": 5
                            }
                        }
                    };
                }
            };

            Xrm.Utility.showProgressIndicator("Unsubmitting Enrolment...");

            Xrm.WebApi.online.execute(req).then(
                function (result) {
                    Xrm.Utility.closeProgressIndicator();
                    if (result.ok) {
                        formContext.data.refresh();
                    } else {
                        var alertStrings = { confirmButtonLabel: "OK", text: "There was an error unsubmitting the enrolment. Please contact IT.", title: "Error" };
                        var alertOptions = { height: 350, width: 450 };
                        Xrm.Navigation.openAlertDialog(alertStrings, alertOptions);
                    }
                },
                function (error) {
                    Xrm.Utility.closeProgressIndicator();
                    var alertStrings = { confirmButtonLabel: "OK", text: "Unsubmitting the enrolment ran into an error. Details: " + error.message, title: "Error" };
                    var alertOptions = { height: 350, width: 450 };
                    Xrm.Navigation.openAlertDialog(alertStrings, alertOptions);
                }
            );
        }
    });

};