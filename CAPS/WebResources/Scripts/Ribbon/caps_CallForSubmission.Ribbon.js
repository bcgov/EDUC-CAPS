"use strict";

var CAPS = CAPS || {};
CAPS.CallForSubmission = CAPS.CallForSubmission || {};

const PROJECT_STATE = {
    DRAFT: 1,
    PUBLISHED: 2,
    RESULTS_RELEASED: 200870000,
    CANCELLED: 100000001
};

/**
 * Function called by Release Results ribbon button, this changes the status to trigger results being released
 * @param {any} primaryControl primary control
 */
CAPS.CallForSubmission.ReleaseResults = function (primaryControl) {
    var formContext = primaryControl;

    var confirmStrings = { text: "Do you wish to release results? Click OK to continue.  ", title: "Confirm Releasing Results" };
    var confirmOptions = { height: 200, width: 450 };
    Xrm.Navigation.openConfirmDialog(confirmStrings, confirmOptions).then(
        function (success) {
            if (success.confirmed) {
                //show cancel tab 
                var currentStatus = formContext.getAttribute("statuscode").getValue().toString();

                if (formContext.getAttribute("statuscode").getValue() === PROJECT_STATE.PUBLISHED) {

                    formContext.getAttribute("statecode").setValue(1);
                    formContext.getAttribute("statuscode").setValue(PROJECT_STATE.RESULTS_RELEASED);
                    formContext.data.entity.save();
                }

            }
        },
        function (error) {
            Xrm.Navigation.openErrorDialog({ message: error });
        }

    );
}

/**
 * Function to determine if the results released button should be displayed.
 * @param {any} primaryControl primary control
 * @returns {bool} true if shown, otherwise false
 */
CAPS.CallForSubmission.ShowReleaseResults = function (primaryControl) {
    debugger;
    var formContext = primaryControl;

    if (formContext.getAttribute("statuscode").getValue() !== PROJECT_STATE.PUBLISHED) {
        return false;
    }

    var userRoles = Xrm.Utility.getGlobalContext().userSettings.roles;

    var showReleaseResults = false;

    userRoles.forEach(function hasFinancialDirectorRole(item, index) {
        if (item.name === "CAPS Financial Director Ministry User - Add On") {
            showReleaseResults =  true;
        }
    });

    return showReleaseResults;
}

/**
 * Function called by Publish ribbon button, this changes the status to trigger the capital plan creation
 * @param {any} primaryControl primary control
 */
CAPS.CallForSubmission.Publish = function (primaryControl) {
    var formContext = primaryControl;
    debugger;

    //If dirty, then save and call again
    if (formContext.data.entity.getIsDirty() || formContext.ui.getFormType() === 1) {
        formContext.data.save({ saveMode: 1 }).then(function success(result) { CAPS.CallForSubmission.Publish(primaryControl); });
    }
    else {

        var confirmStrings = { text: "Do you wish to publish? Click OK to continue.  ", title: "Confirm Publish Call For Submission" };
        var confirmOptions = { height: 200, width: 450 };
        Xrm.Navigation.openConfirmDialog(confirmStrings, confirmOptions).then(
            function (success) {
                if (success.confirmed) {

                    var recordId = formContext.data.entity.getId().replace("{", "").replace("}", "");
                    //call action
                    var req = {};
                    var target = { entityType: "caps_callforsubmission", id: recordId };
                    req.entity = target;

                    req.getMetadata = function () {
                        return {
                            boundParameter: "entity",
                            operationType: 0,
                            operationName: "caps_FlipPublishCallforSubmissiontoYES",
                            parameterTypes: {
                                "entity": {
                                    "typeName": "mscrm.caps_callforsubmission",
                                    "structuralProperty": 5
                                }
                            }
                        };
                    };

                    Xrm.WebApi.online.execute(req).then(
                        function (result) {

                            if (result.ok) {
                                return result.json().then(
                                    function (response) {
                                        //return response
                                        if (response.isValid) {
                                            let alertStrings = { confirmButtonLabel: "OK", text: "Publish complete.", title: "Call For Submission" };
                                            let alertOptions = { height: 120, width: 260 };
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
                                            let alertStrings = { confirmButtonLabel: "OK", text: "Publish failed. Details: " + response.ValidationMessage, title: "Call For Submission" };
                                            let alertOptions = { height: 120, width: 260 };
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


                        }
                    );


                }
            });

    }
}

/**
 * Function to determine if the publish button should be displayed.
 * @param {any} primaryControl primary control
 * @returns {bool} true if shown, otherwise false
 */
CAPS.CallForSubmission.ShowPublish = function (primaryControl) {
    //check that record is draft & user's roles
    var formContext = primaryControl;

    if (formContext.ui.getFormType() === 1) {
        return false;
    }
    if (formContext.getAttribute("statuscode").getValue() !== PROJECT_STATE.DRAFT) {
        return false;
    }

    var userRoles = Xrm.Utility.getGlobalContext().userSettings.roles;

    var showPublish = false;

    userRoles.forEach(function hasFinancialDirectorRole(item, index) {
        if (item.name === "CAPS Financial Director Ministry User - Add On" || item.name === "CAPS Ministry Super User - Add On") {
            showPublish = true;
        }
    });

    return showPublish;
}

/**
 * Function called by Edit ribbon button, this changes the status to allow editing of the call for submission
 * @param {any} primaryControl primary control
 */
CAPS.CallForSubmission.Edit = function (primaryControl) {
    var formContext = primaryControl;

    var confirmStrings = { text: "Do you wish to edit the call for submission? Click OK to continue.  ", title: "Confirm Edit" };
    var confirmOptions = { height: 200, width: 450 };
    Xrm.Navigation.openConfirmDialog(confirmStrings, confirmOptions).then(
        function (success) {
            if (success.confirmed) {
                //show edit button
                if (formContext.getAttribute("statuscode").getValue() === PROJECT_STATE.PUBLISHED) {

                    formContext.getAttribute("statecode").setValue(0);
                    formContext.getAttribute("statuscode").setValue(PROJECT_STATE.DRAFT);
                    formContext.data.entity.save();
                }

            }
        },
        function (error) {
            Xrm.Navigation.openErrorDialog({ message: error });
        }

    );
}

/**
 * Function to determine if the edit button should be displayed.
 * @param {any} primaryControl primary control
 * @returns {bool} true if shown, otherwise false
 */
CAPS.CallForSubmission.ShowEdit = function (primaryControl) {
    //check that record is draft & user's roles
    var formContext = primaryControl;

    if (formContext.getAttribute("statuscode").getValue() !== PROJECT_STATE.PUBLISHED) {
        return false;
    }

    var userRoles = Xrm.Utility.getGlobalContext().userSettings.roles;

    var showPublish = false;

    userRoles.forEach(function hasFinancialDirectorRole(item, index) {
        if (item.name === "CAPS Financial Director Ministry User - Add On" || item.name === "CAPS Ministry Super User - Add On") {
            showPublish = true;
        }
    });

    return showPublish;
}

/**
 * Function called by Cancel ribbon button, this changes the status to cancel the call for submission
 * @param {any} primaryControl primary control
 */
CAPS.CallForSubmission.Cancel = function (primaryControl) {
    var formContext = primaryControl;

    var confirmStrings = { text: "Do you wish to cancel the call for submission? Click OK to continue.  ", title: "Confirm Cancel" };
    var confirmOptions = { height: 200, width: 450 };
    Xrm.Navigation.openConfirmDialog(confirmStrings, confirmOptions).then(
        function (success) {
            if (success.confirmed) {

                if (formContext.getAttribute("statuscode").getValue() === PROJECT_STATE.DRAFT) {

                    formContext.getAttribute("statecode").setValue(1);
                    formContext.getAttribute("statuscode").setValue(PROJECT_STATE.CANCELLED);
                    formContext.data.entity.save();
                }

            }
        },
        function (error) {
            Xrm.Navigation.openErrorDialog({ message: error });
        }

    );
}

/**
 * Function to determine if the cancel button should be displayed.
 * @param {any} primaryControl primary control
 * @returns {bool} true if shown, otherwise false
 */
CAPS.CallForSubmission.ShowCancel = function (primaryControl) {
    //check that record is draft & user's roles
    var formContext = primaryControl;

    if (formContext.getAttribute("statuscode").getValue() !== PROJECT_STATE.DRAFT) {
        return false;
    }

    var userRoles = Xrm.Utility.getGlobalContext().userSettings.roles;

    var showPublish = false;

    userRoles.forEach(function hasFinancialDirectorRole(item, index) {
        if (item.name === "CAPS Financial Director Ministry User - Add On" || item.name === "CAPS Ministry Super User - Add On") {
            showPublish = true;
        }
    });

    return showPublish;
}

