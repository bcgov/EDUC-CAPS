"use strict";

var CAPS = CAPS || {};
CAPS.Project = CAPS.Project ||
    {
    GLOBAL_FORM_CONTEXT: null,
    PROJECT_ARRAY: null,
    SHOW_CANCEL_ASYNC_COMPLETED: false,
    SHOW_CANCEL_BUTTON: false,
    VIEW_SELECTED_CONTROL: null
    };

const PROJECT_STATE = {
    DRAFT: 1,
    SUBMITTED: 100000009,
    PLANNED: 100000004,
    SUPPORTED: 200870002,
    APPROVED: 100000005,
    COMPLETE: 100000009,
    ACCEPTED: 200870000
};


/**
 * Called from Project Homepage (main list view).  Takes in a list of records.  This function
 * calls ShowSubmissionWindow to show the submission selection popup.
 * @param {any} selectedControlIds selected Control IDs
 * @param {any} selectedControl selected Control, used for refreshing view
 */
CAPS.Project.AddListToSubmission = function (selectedControlIds, selectedControl) {
    CAPS.Project.VIEW_SELECTED_CONTROL = selectedControl;
    //Get all "Draft" projects and confirm that the selected list only contains Draft ones.
    Xrm.WebApi.retrieveMultipleRecords("caps_project", "?$select=caps_projectid&$filter=statuscode eq 1 and caps_Submission eq null").then(
        function success(result) {

            var unqualifiedRecordFound = false;

            var draftProjects = [];

            for (var i = 0; i < result.entities.length; i++) {
                draftProjects.push(result.entities[i]["caps_projectid"]);
            }

            selectedControlIds.forEach((record) => {
                if (draftProjects.indexOf(record) === -1) {
                    unqualifiedRecordFound = true;
                }
            });

            if (unqualifiedRecordFound) {
                var alertStrings = { confirmButtonLabel: "OK", text: "One or more projects can't be added to the submission.  This is because they are either already in a submission or they are not in a draft/published state.", title: "Error" };
                var alertOptions = { height: 120, width: 260 };
                Xrm.Navigation.openAlertDialog(alertStrings, alertOptions);
            }
            else {
                CAPS.Project.ShowSubmissionWindow(selectedControlIds);
            }
        },
        function (error) {
            console.log(error.message);
            // handle error conditions
            Xrm.Navigation.openErrorDialog({ message: error.message });
            return false;
        }
    );
}

/**
 * Called from Project Form.  Takes in the form context as primary Control.  This function
 * calls ShowSubmissionWindow to show the submission selection popup.
 * @param {any} primaryControl primary Control
 */
CAPS.Project.AddToSubmission = function (primaryControl) {
    var formContext = primaryControl;
    CAPS.Project.GLOBAL_FORM_CONTEXT = formContext;

    //If dirty, then save and call again
    if (formContext.data.entity.getIsDirty() || formContext.ui.getFormType() === 1) {
        formContext.data.save({ saveMode: 1 }).then(function success(result) { CAPS.Project.AddToSubmission(primaryControl); });
    }
    else {
        var selectedControlIds = [formContext.data.entity.getId()];
        CAPS.Project.ShowSubmissionWindow(selectedControlIds);
    }
}

/**
 * Function to determine if the Add to Capital Plan button should be displayed.
 * @param {any} primaryControl primary control
 * @return {boolean} true if should be shown, otherwise false
 */
CAPS.Project.ShowAddToSubmission = function (primaryControl) {
    var formContext = primaryControl;

    if (formContext.getAttribute("statecode").getValue() === 0
        && formContext.getAttribute("caps_submission").getValue() === null) {
        return true;
    }
    return false;
}

/**
 * Function which determines if the cancel button should be shown.
 * If using the MyCaps app, the button is only shown on draft records
 * If using the Caps app, the button is only shown for planned, supported or approved projects
 * @param {any} primaryControl record's primary Control 
 * @returns {any} true/false
 */
CAPS.Project.ShowCancelButton = function (primaryControl) {
    var formContext = primaryControl;

    if (CAPS.Project.SHOW_CANCEL_ASYNC_COMPLETED) {
        return CAPS.Project.SHOW_CANCEL_BUTTON;
    }

    var globalContext = Xrm.Utility.getGlobalContext();
    globalContext.getCurrentAppName().then(
        function success(result) {
            CAPS.Project.SHOW_CANCEL_ASYNC_COMPLETED = true;
            
            var recordStatus = formContext.getAttribute("statuscode").getValue();
            if (result === "MyCAPS") {
                if (recordStatus === PROJECT_STATE.DRAFT || recordStatus === PROJECT_STATE.ACCEPTED) {
                    CAPS.Project.SHOW_CANCEL_BUTTON = true;
                }
                
            }


            if (CAPS.Project.SHOW_CANCEL_BUTTON) {
                formContext.ui.refreshRibbon();
            }
        }
        , function (error) {
            CAPS.Project.SHOW_CANCEL_ASYNC_COMPLETED = true;
            CAPS.Project.SHOW_CANCEL_BUTTON = false;
            Xrm.Navigation.openAlertDialog({ text: error.message });
        });
}


/**
 * Called from Project Form, this function shows the cancel tab as well as the reason for cancellation field and makes it mandatory.
 * @param {any} primaryControl primary control
 */
CAPS.Project.CancelProject = function (primaryControl) {
    var formContext = primaryControl;

    var confirmStrings = { text: "Do you wish to cancel this project request? Click OK to continue or Cancel to keep the project request. \n\r⠀\n\rPlease note, AFG Project Requests will be permanently deleted. ", title: "Confirm Project Request Cancellation" };
    var confirmOptions = { height: 200, width: 450 };
    Xrm.Navigation.openConfirmDialog(confirmStrings, confirmOptions).then(
        function (success) {
            if (success.confirmed) {
                //show cancel tab 
                var currentStatus = formContext.getAttribute("statuscode").getValue().toString();

                if (formContext.getAttribute("statuscode").getValue() === PROJECT_STATE.DRAFT) {
                    //update pre-cancellation status
                    formContext.getAttribute("caps_precancellationstatus").setValue(currentStatus);

                    formContext.getAttribute("statecode").setValue(1);
                    formContext.getAttribute("statuscode").setValue(100000010);
                    formContext.data.entity.save();
                }
                else {
                    //show cancelation tab
                    formContext.ui.tabs.get("tab_cancel").setVisible(true);

                    //update pre-cancellation status
                    formContext.getAttribute("caps_precancellationstatus").setValue(currentStatus);

                    //show reason for cancellation and make mandatory
                    formContext.getAttribute("caps_reasonforcancellation").setRequiredLevel("required");
                    formContext.getControl("caps_reasonforcancellation").setVisible(true);
                    formContext.getControl("caps_reasonforcancellation").setFocus();

                    //change status
                    formContext.getAttribute("statecode").setValue(1);
                    formContext.getAttribute("statuscode").setValue(100000010);

                    //Prevent auto-save
                    CAPS.Project.PREVENT_AUTO_SAVE = true;
                }
            }
        },
        function (error) {
            Xrm.Navigation.openErrorDialog({ message: error });
        }

    );
}

/**
 * Function which determines if the uncancel button should be shown.
 * If using the MyCaps app, the button is only shown on records that were draft records
 * If using the Caps app, the button is only shown on records that were planned, supported or approved projects
 * @param {any} primaryControl record's primary Control
 * @returns {any} true/false
 */
CAPS.Project.ShowUncancelButton = function (primaryControl) {
    var formContext = primaryControl;

    if (CAPS.Project.SHOW_CANCEL_ASYNC_COMPLETED) {
        return CAPS.Project.SHOW_CANCEL_BUTTON;
    }

    var globalContext = Xrm.Utility.getGlobalContext();
    globalContext.getCurrentAppName().then(
        function success(result) {
            CAPS.Project.SHOW_CANCEL_ASYNC_COMPLETED = true;

            var recordStatus = parseInt(formContext.getAttribute("caps_precancellationstatus").getValue());

            if (result === "MyCAPS") {
                if (recordStatus === PROJECT_STATE.DRAFT) {
                    CAPS.Project.SHOW_CANCEL_BUTTON = true;
                }

            }
            else if (result === "CAPS") {
                if (recordStatus === PROJECT_STATE.PLANNED || recordStatus === PROJECT_STATE.SUPPORTED || recordStatus === PROJECT_STATE.APPROVED) {
                    CAPS.Project.SHOW_CANCEL_BUTTON = true;
                }
            }

            if (CAPS.Project.SHOW_CANCEL_BUTTON) {
                formContext.ui.refreshRibbon();
            }
        }
        , function (error) {
            CAPS.Project.SHOW_CANCEL_ASYNC_COMPLETED = true;
            CAPS.Project.SHOW_CANCEL_BUTTON = false;
            Xrm.Navigation.openAlertDialog({ text: error.message });
        });
}

/**
 * Called from Project Form, this function allows the user to uncancel a project and put it back into it's previous status.
 * @param {any} primaryControl the record's primary control
 */
CAPS.Project.UncancelProject = function (primaryControl) {
    var formContext = primaryControl;

    var confirmStrings = { text: "Are you sure you want to activate the selected Project?", title: "Confirm Project Activation" };
    var confirmOptions = { height: 200, width: 450 };
    Xrm.Navigation.openConfirmDialog(confirmStrings, confirmOptions).then(
        function (success) {
            if (success.confirmed) {
                //show cancel tab 
                var preCancellationStatus = PROJECT_STATE.DRAFT;
                if (formContext.getAttribute("caps_precancellationstatus").getValue() !== null) {
                    preCancellationStatus = parseInt(formContext.getAttribute("caps_precancellationstatus").getValue());
                }

                var data =
                {
                    "statecode": 0,
                    "statuscode": preCancellationStatus,
                    "caps_precancellationstatus": null,
                    "caps_reasonforcancellation": null
                };

                var recordId = formContext.data.entity.getId().replace("{", "").replace("}", "");

                // update the record
                Xrm.WebApi.updateRecord("caps_project", recordId, data).then(
                    function success(result) {
                        // perform operations on record update
                        formContext.data.refresh();
                    },
                    function (error) {
                        // handle error conditions
                        Xrm.Navigation.openErrorDialog({ message: error.message });
                    }
                );

            }
        },
        function (error) {
            Xrm.Navigation.openErrorDialog({ message: error });
        }
    );
}

/**
 * Opens a modal window with a submission drop down
 * @param {any} selectedControlIds selected Control IDs
 */
CAPS.Project.ShowSubmissionWindow = function (selectedControlIds) {

    CAPS.Project.PROJECT_ARRAY = selectedControlIds;
    var globalContext = Xrm.Utility.getGlobalContext();
    var clientUrl = globalContext.getClientUrl();
    
    var webResource = '/caps_/Apps/OpenSubmissionList.htm';

    Alert.showWebResource(webResource, 500, 230, "Add to Submission", [
        new Alert.Button("Add", CAPS.Project.SubmissionResult, true, true),
        new Alert.Button("Cancel")
    ], clientUrl, true, null);

}

/*
 * Called on Add button of the modal.  This function updates the project(s) records with the selected submission.
 * ** */
CAPS.Project.SubmissionResult = function () {

    var validationResult = Alert.getIFrameWindow().validate();
    var data =
    {
        "caps_Submission@odata.bind": "/caps_submissions(" + validationResult + ")"
    };
    var promises = [];

    //Update records
    CAPS.Project.PROJECT_ARRAY.forEach(function (item, index) {
        // update the record
        promises.push(Xrm.WebApi.updateRecord("caps_project", item, data));
    });

    Promise.all(promises).then(
        function (results) {
            debugger;
            //Close Popup
            Alert.hide();
            //refesh view
            if (CAPS.Project.VIEW_SELECTED_CONTROL !== undefined && CAPS.Project.VIEW_SELECTED_CONTROL !== null) {
                CAPS.Project.VIEW_SELECTED_CONTROL.refresh();
            }
            if (CAPS.Project.GLOBAL_FORM_CONTEXT !== undefined && CAPS.Project.GLOBAL_FORM_CONTEXT !== null) {
                CAPS.Project.GLOBAL_FORM_CONTEXT.data.refresh();
            }
        }
        , function (error) {
            //Close Popup
            Alert.hide();
            Xrm.Navigation.openErrorDialog({ message: error.message });
            //refresh view
            if (CAPS.Project.VIEW_SELECTED_CONTROL !== null) {
                CAPS.Project.VIEW_SELECTED_CONTROL.refresh();
            }
        }
    );
}

/**
 * This function triggers the Calculate Schedule B action.
 * @param {any} primaryControl primary control
 */
CAPS.Project.CalculateScheduleB = function (primaryControl) {

    var formContext = primaryControl;

    //If dirty, then save and call again
    if (formContext.data.entity.getIsDirty() || formContext.ui.getFormType() === 1) {
        formContext.data.save({ saveMode: 1 }).then(function success(result) { CAPS.Project.CalculateScheduleB(primaryControl); });
    }
    else {
        Xrm.Utility.showProgressIndicator("Calculating Cost...");

        var recordId = formContext.data.entity.getId().replace("{", "").replace("}", "");
        //call action
        var req = {};
        var target = { entityType: "caps_project", id: recordId };
        req.entity = target;

        req.getMetadata = function () {
            return {
                boundParameter: "entity",
                operationType: 0,
                operationName: "caps_TriggerScheduleBCalculation",
                parameterTypes: {
                    "entity": {
                        "typeName": "mscrm.caps_project",
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
                            //close the indicator
                            Xrm.Utility.closeProgressIndicator();
                            
                            //get error message
                            if (response.ScheduleBErrorMessage == null) {
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
                                var alertStrings = { confirmButtonLabel: "OK", text: "The preliminary budget calculation ran into a problem. Details: " + response.ScheduleBErrorMessage, title: "Preliminary Budget Result" };
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
                //close the indicator
                Xrm.Utility.closeProgressIndicator();
                var alertStrings = { confirmButtonLabel: "OK", text:  e.message, title: "Preliminary Budget Result" };
                var alertOptions = { height: 350, width: 450 };
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
CAPS.Project.ShowCalculateScheduleB = function (primaryControl) {
    var formContext = primaryControl;

    return (formContext.getAttribute("caps_requiresscheduleb").getValue() && formContext.getAttribute("statuscode").getValue() === PROJECT_STATE.DRAFT);
}

/**
 * Function to validate the project request record.
 * @param {any} primaryControl primary control
 */
CAPS.Project.Validate = function (primaryControl) {
    var formContext = primaryControl;

    //If dirty, then save and call again
    if (formContext.data.entity.getIsDirty() || formContext.ui.getFormType() === 1) {
        formContext.data.save({ saveMode: 1 }).then(function success(result) { CAPS.Project.Validate(primaryControl); });
    }
    else {
        var recordId = formContext.data.entity.getId().replace("{", "").replace("}", "");
        //call action
        var req = {};
        var target = { entityType: "caps_project", id: recordId };
        req.entity = target;

        req.getMetadata = function () {
            return {
                boundParameter: "entity",
                operationType: 0,
                operationName: "caps_ValidateProjectRequest",
                parameterTypes: {
                    "entity": {
                        "typeName": "mscrm.caps_project",
                        "structuralProperty": 5
                    }
                }
            }
        };

        Xrm.WebApi.online.execute(req).then(
            function (response) {
                var alertStrings = { confirmButtonLabel: "OK", text: "Validation complete.", title: "Validation" };
                var alertOptions = { height: 120, width: 260 };
                Xrm.Navigation.openAlertDialog(alertStrings, alertOptions).then(
                    function success(result) {
                        console.log("Alert dialog closed");
                        formContext.data.refresh();
                        formContext.getControl("caps_validationstatus").setFocus();
                    },
                    function (error) {
                        console.log(error.message);
                    }
                );
            },
            function (e) {

                var alertStrings = { confirmButtonLabel: "OK", text: "Validation failed. Details: " + e.message, title: "Validation" };
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
 * Function to determine when the Validate button should be shown.
 * @param {any} primaryControl primary control
 * @returns {boolean} true if should be shown, otherwise false.
 */
CAPS.Project.ShowValidate = function (primaryControl) {
    var formContext = primaryControl;
    
    //validate for all draft projects except AFG
    if (formContext.getAttribute("statuscode").getValue() === PROJECT_STATE.DRAFT) {
        return true;
    }
    else {
        return false;
    }
}

/*
Function to check if the current user has CAPS CMB User Role.
*/
CAPS.Project.IsMinistryUser = function () {
    var userRoles = Xrm.Utility.getGlobalContext().userSettings.roles;

    var showButton = false;

    userRoles.forEach(function hasFinancialDirectorRole(item, index) {
        if (item.name === "CAPS CMB User") {
            showButton = true;
        }
    });

    return showButton;
}

/*
Function to mark all selected project requests as Supported.  This is done via an action "caps_SetProjectRequesttoSupported" which checks if the project is submitted and updates it only if it is.
*/
CAPS.Project.MarkAsSupported = function (selectedControlIds, selectedControl) {
    
    //call action
    var promises = [];

    selectedControlIds.forEach((record) => {
        let req = {};
        req.getMetadata = function () {
            return {
                boundParameter: "entity",
                operationType: 0,
                operationName: "caps_SetProjectRequesttoSupported",
                parameterTypes: {
                    "entity": {
                        "typeName": "mscrm.caps_project",
                        "structuralProperty": 5
                    }
                }
            }
        };
        req.entity = { entityType: "caps_project", id: record };
        promises.push(Xrm.WebApi.online.execute(req));
    });

    Promise.all(promises).then(
    function (results) {
        selectedControl.refresh();
    }
    , function (error) {
        console.log(error);
    }
    );
}

/*
Function to mark all selected project requests as UnSupported.  This is done via an action "caps_SetProjectRequesttoUnsupported" which checks if the project is submitted and updates it only if it is.
*/
CAPS.Project.MarkAsUnsupported = function (selectedControlIds, selectedControl) {
    //call action
    var promises = [];
    selectedControlIds.forEach((record) => {
        let req = {};
        req.getMetadata = function () {
            return {
                boundParameter: "entity",
                operationType: 0,
                operationName: "caps_SetProjectRequesttoUnsupported",
                parameterTypes: {
                    "entity": {
                        "typeName": "mscrm.caps_project",
                        "structuralProperty": 5
                    }
                }
            }
        };
        req.entity = { entityType: "caps_project", id: record };
        promises.push(Xrm.WebApi.online.execute(req));
    });

    Promise.all(promises).then(
    function (results) {
        selectedControl.refresh();
    }
    , function (error) {
        console.log(error);
    }
);
}

/*
Function to check if the current user has CAPS School District User Role.
*/
CAPS.Project.IsSDUser = function () {
    var userRoles = Xrm.Utility.getGlobalContext().userSettings.roles;

    var showButton = false;

    userRoles.forEach(function hasFinancialDirectorRole(item, index) {
        if (item.name === "CAPS School District User") {
            showButton = true;
        }
    });

    return showButton;
}

/*
Function to mark all selected project requests as Published.  This is done via an action "caps_PublishProjectRequest" which checks if the project is published and updates it if it isn't.
*/
CAPS.Project.MarkAsPublished = function (selectedControlIds, selectedControl) {

    //call action
    var promises = [];

    selectedControlIds.forEach((record) => {
        let req = {};
        req.getMetadata = function () {
            return {
                boundParameter: "entity",
                operationType: 0,
                operationName: "caps_PublishProjectRequest",
                parameterTypes: {
                    "entity": {
                        "typeName": "mscrm.caps_project",
                        "structuralProperty": 5
                    }
                }
            }
        };
        req.entity = { entityType: "caps_project", id: record };
        promises.push(Xrm.WebApi.online.execute(req));
    });

    Promise.all(promises).then(
    function (results) {
        selectedControl.refresh();
    }
    , function (error) {
        console.log(error);
    }
    );
}

/*
Function to mark all selected project requests as Unpublished.  This is done via an action "caps_UnpublishProjectRequest" which checks if the project is unpublished and updates it if it's not.
*/
CAPS.Project.MarkAsUnpublished = function (selectedControlIds, selectedControl) {
    //call action
    var promises = [];
    selectedControlIds.forEach((record) => {
        let req = {};
        req.getMetadata = function () {
            return {
                boundParameter: "entity",
                operationType: 0,
                operationName: "caps_UnpublishProjectRequest",
                parameterTypes: {
                    "entity": {
                        "typeName": "mscrm.caps_project",
                        "structuralProperty": 5
                    }
                }
            }
        };
        req.entity = { entityType: "caps_project", id: record };
        promises.push(Xrm.WebApi.online.execute(req));
    });

    Promise.all(promises).then(
    function (results) {
        selectedControl.refresh();
    }
    , function (error) {
        console.log(error);
    }
);
}

/*
Function to check if the current user has CAPS CMB User Role.
*/
CAPS.Project.IsMinistrySuperUser = function () {
    debugger;
    var userRoles = Xrm.Utility.getGlobalContext().userSettings.roles;

    var showButton = false;

    userRoles.forEach(function hasFinancialDirectorRole(item, index) {
        if (item.name === "CAPS CMB Super User - Add On") {
            showButton = true;
        }
    });

    return showButton;
}





