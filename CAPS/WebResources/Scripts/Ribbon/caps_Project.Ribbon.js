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
    SUPPORTED: 100000003,
    APPROVED: 100000005,
    COMPLETE: 100000009
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
    Xrm.WebApi.retrieveMultipleRecords("caps_project", "?$select=caps_projectid&$filter=statuscode eq 1").then(
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
                var alertStrings = { confirmButtonLabel: "OK", text: "One or more projects can't be added to the capital plan.", title: "Error" };
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
        formContext.data.save().then(CAPS.Project.AddToSubmission(primaryControl));
    }

    var selectedControlIds = [formContext.data.entity.getId()];
    CAPS.Project.ShowSubmissionWindow(selectedControlIds);

    //DB: This is the new way of opening a modal but it's not fully implmented yet.
    //var pageInput = {
    //    pageType: "webresource",
    //    webresourceName: webResource
    //};
    //var navigationOptions = {
    //    target: 2,
    //    width: 400,
    //    height: 300,
    //    position: 1
    //};
    //Xrm.Navigation.navigateTo(pageInput, navigationOptions).then(
    //    function success() {
    //        // Handle dialog closed
    //    },
    //    function error() {
    //        // Handle errors
    //    }
    //);
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
 * Called from Project Form, this function shows the cancel tab as well as the reason for cancellation field and makes it mandatory.
 * @param {any} primaryControl primary control
 */
CAPS.Project.CancelProject = function (primaryControl) {
    var formContext = primaryControl;

    var confirmStrings = { text: "Do you wish to deactivate this project? Click OK to continue or Cancel to keep the project.  ", title: "Confirm Project Cancellation" };
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

    Alert.showWebResource(webResource, 500, 230, "Add to Capital Plan", [
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
            //Close Popup
            Alert.hide();
            //refesh view
            if (CAPS.Project.VIEW_SELECTED_CONTROL !== null) {
                CAPS.Project.VIEW_SELECTED_CONTROL.refresh();
            }
            if (CAPS.Project.GLOBAL_FORM_CONTEXT !== null) {
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



