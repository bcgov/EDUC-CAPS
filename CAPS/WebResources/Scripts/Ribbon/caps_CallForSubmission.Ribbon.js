"use strict";

var CAPS = CAPS || {};
CAPS.CallForSubmission = CAPS.CallForSubmission || {};

const PROJECT_STATE = {
    DRAFT: 1,
    PUBLISHED: 2,
    RESULTS_RELEASED: 200870000,
    ACCEPTED: 200870001,
    CANCELLED: 100000001
};

/**
 * Function called by Release Results ribbon button, this changes the status to trigger results being released
 * @param {any} primaryControl primary control
 */
CAPS.CallForSubmission.ReleaseResults = function (primaryControl) {
    debugger;
    var formContext = primaryControl;
    var submissionId = formContext.data.entity.getId();
    var boardResolutionRequired = formContext.getAttribute("caps_boardresolutionrequired").getValue();
    var fetchXML = "";
    if (!boardResolutionRequired) {
        fetchXML = "<fetch version=\"1.0\" output-format=\"xml-platform\" mapping=\"logical\" distinct=\"false\">" +
            "<entity name=\"caps_submission\">" +
            "<attribute name=\"caps_submissionid\" />" +
            "<attribute name=\"caps_name\" />" +
            "<attribute name=\"createdon\" />" +
            "<order attribute=\"caps_name\" descending=\"false\" />" +
            "<filter type=\"and\">" +
            "<condition attribute=\"caps_callforsubmission\" operator=\"eq\" value=\"" + submissionId + "\" />" +
            "<condition attribute=\"statuscode\" value=\"100000001\" operator=\"ne\"/>" +
            "<filter type=\"or\">" +
            "<condition attribute=\"statuscode\" value=\"1\" operator=\"eq\"/>" +
            "</filter>" +
            "<filter type=\"and\">" +
            "<condition attribute=\"caps_callforsubmissiontype\" operator=\"in\">" +
            "<value>100000000</value>" +
            "<value>100000001</value>" +
            "</condition>" +
            "</filter>" +
            "</filter>" +
            "<link-entity name=\"caps_callforsubmission\" from=\"caps_callforsubmissionid\" to=\"caps_callforsubmission\" link-type=\"inner\" alias=\"aa\">" +
           // "<filter type=\"or\">" +
           // "<condition attribute=\"caps_boardresolutionrequired\" operator=\"eq\" value=\"1\" />" +
           // "</filter>" +
            "</link-entity>" +
            "</entity>" +
            "</fetch>";
    }
    else {
        fetchXML = "<fetch version=\"1.0\" output-format=\"xml-platform\" mapping=\"logical\" distinct=\"false\">" +
            "<entity name=\"caps_submission\">" +
            "<attribute name=\"caps_submissionid\" />" +
            "<attribute name=\"caps_name\" />" +
            "<attribute name=\"createdon\" />" +
            "<order attribute=\"caps_name\" descending=\"false\" />" +
            "<filter type=\"and\">" +
            "<condition attribute=\"caps_callforsubmission\" operator=\"eq\" value=\"" + submissionId + "\" />" +
            "<condition attribute=\"statuscode\" value=\"100000001\" operator=\"ne\"/>" +
            "<filter type=\"or\">" +
            "<condition attribute=\"statuscode\" value=\"1\" operator=\"eq\"/>" +
            "<condition attribute=\"caps_boardofresolutionattached\" value=\"0\" operator=\"eq\"/>" +
            "</filter>" +
            "<filter type=\"and\">" +
            "<condition attribute=\"caps_callforsubmissiontype\" operator=\"in\">" +
            "<value>100000000</value>" +
            "<value>100000001</value>" +
            "</condition>" +
            "</filter>" +
            "</filter>" +
            "<link-entity name=\"caps_callforsubmission\" from=\"caps_callforsubmissionid\" to=\"caps_callforsubmission\" link-type=\"inner\" alias=\"aa\">" +
            "<filter type=\"or\">" +
            "<condition attribute=\"caps_boardresolutionrequired\" operator=\"eq\" value=\"1\" />" +
            "</filter>" +
            "</link-entity>" +
            "</entity>" +
            "</fetch>";
    }

    //Get all Capital Plans without board resolutions
    Xrm.WebApi.retrieveMultipleRecords("caps_submission", "?fetchXml=" + fetchXML).then(
        function success(result) {

            var errorText = "Unable to release results as one or more Submission is in a draft state and/or missing a board resolution.";
            if (!boardResolutionRequired) {
                errorText = "Unable to release results as one or more Submission is in a draft state.";
            }
            
            /*var errorText = "Unable to release results as one or more Submission is in a draft state."*/

            if (formContext.getAttribute("caps_callforsubmissiontype").getValue() == 100000002) {
                errorText = "Unable to release results as one or more Submission is in a draft state and/or a project request is flagged for review.";
            }

            if (result.entities.length > 0) {
                //Some bad projects
                let alertStrings = { confirmButtonLabel: "OK", text: errorText, title: "Call For Submission" };
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

            else {
                //now if this is AFG, check if there are any projects requests that are flagged for review.
                if (formContext.getAttribute("caps_callforsubmissiontype").getValue() == 100000002) {
                    var fetchXML2 = "<fetch version=\"1.0\" output-format=\"xml-platform\" mapping=\"logical\" distinct=\"true\">" +
                        "<entity name=\"caps_submission\">" +
                        "<attribute name=\"caps_submissionid\" />" +
                        "<attribute name=\"caps_name\" />" +
                        "<attribute name=\"createdon\" />" +
                        "<attribute name=\"caps_submissiontype\" />" +
                        "<order attribute=\"caps_name\" descending=\"false\" />" +
                        "<filter type=\"and\">" +
                        "<condition attribute=\"caps_callforsubmission\" operator=\"eq\" value=\"" + submissionId + "\" />" +
                        "</filter>" +
                        "<link-entity name=\"caps_project\" from=\"caps_submission\" to=\"caps_submissionid\" link-type=\"inner\" alias=\"ad\">" +
                        "<filter type=\"and\">" +
                        "<condition attribute=\"caps_flaggedforreview\" operator=\"eq\" value=\"1\" />" +
                        "</filter>" +
                        "</link-entity>" +
                        "</entity>" +
                        "</fetch>";

                    Xrm.WebApi.retrieveMultipleRecords("caps_submission", "?fetchXml=" + fetchXML2).then(
                        function success(result) {
                            if (result.entities.length > 0) {
                                //Some bad projects
                                let alertStrings = { confirmButtonLabel: "OK", text: errorText, title: "Call For Submission" };
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
                            else {
                                CAPS.CallForSubmission.ConfirmResultsReleased(formContext);
                            }
                        },
                        function (error) {
                            Xrm.Navigation.openErrorDialog({ message: error });
                        });
                }
                else {
                    //check that all Project requests are marked as supported, not supported or planned
                    var fetchXML3 = "<fetch version=\"1.0\" output-format=\"xml-platform\" mapping=\"logical\" distinct=\"true\">" +
                        "<entity name=\"caps_submission\">" +
                        "<attribute name=\"caps_submissionid\" />" +
                        "<attribute name=\"caps_name\" />" +
                        "<attribute name=\"createdon\" />" +
                        "<attribute name=\"caps_submissiontype\" />" +
                        "<order attribute=\"caps_name\" descending=\"false\" />" +
                        "<filter type=\"and\">" +
                        "<condition attribute=\"caps_callforsubmission\" operator=\"eq\" value=\"" + submissionId + "\" />" +
                        "<condition attribute=\"statuscode\" value=\"100000001\" operator=\"ne\"/>" +
                        "</filter>" +
                        "<link-entity name=\"caps_project\" from=\"caps_submission\" to=\"caps_submissionid\" link-type=\"inner\" alias=\"ad\">" +
                        "<filter type=\"and\">" +
                        "<condition attribute=\"caps_ministryassessmentstatus\" operator=\"not-in\">" +
                        /*  "<value>200870000</value>" +*/
                        "<value>100000001</value>" +
                        "<value>100000000</value>" +
                        "</condition>" +
                        "</filter>" +
                        "</link-entity>" +
                        "</entity>" +
                        "</fetch>";

                    Xrm.WebApi.retrieveMultipleRecords("caps_submission", "?fetchXml=" + fetchXML3).then(
                        function success(result) {
                            if (result.entities.length > 0) {
                                //Some bad projects
                                /*  errorText = "Unable to release results as one or more Submissions contains a project request not marked as supported, not supported or planned.";*/
                                errorText = "Unable to release results as one or more Submissions contains a project request not marked as supported or not supported.";
                                let alertStrings = { confirmButtonLabel: "OK", text: errorText, title: "Call For Submission" };
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
                            else {
                                CAPS.CallForSubmission.ConfirmResultsReleased(formContext);

                            }
                        },
                        function (error) {
                            Xrm.Navigation.openErrorDialog({ message: error });
                        });
                }
            }

        },
        function (error) {
            alert(error.message);
        }
    );

}


/**
Function called to display the confirmation message when user clicks Release Results.
*/
CAPS.CallForSubmission.ConfirmResultsReleased = function (formContext) {

    var confirmStrings = { text: "Do you wish to release results? Click OK to continue.  ", title: "Confirm Releasing Results" };
    var confirmOptions = { height: 200, width: 450 };
    Xrm.Navigation.openConfirmDialog(confirmStrings, confirmOptions).then(
        function (success) {
            if (success.confirmed) {
                //show cancel tab 
                var currentStatus = formContext.getAttribute("statuscode").getValue().toString();

                if (formContext.getAttribute("statuscode").getValue() === PROJECT_STATE.PUBLISHED) {

                    formContext.getAttribute("statecode").setValue(1);
                    if (formContext.getAttribute("caps_callforsubmissiontype").getValue() == 100000002) {
                        formContext.getAttribute("statuscode").setValue(PROJECT_STATE.ACCEPTED);
                    }
                    else {
                        formContext.getAttribute("statuscode").setValue(PROJECT_STATE.RESULTS_RELEASED);
                    }
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
        if (item.name === "CAPS CMB Release Submission Results - Add On" || item.name === "CAPS CMB Super User - Add On") {
            showReleaseResults = true;
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
        if (item.name === "CAPS CMB Finance Unit - Add On" || item.name === "CAPS CMB Super User - Add On") {
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
        if (item.name === "CAPS CMB Finance Unit - Add On" || item.name === "CAPS CMB Super User - Add On") {
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
        if (item.name === "CAPS CMB Finance Unit - Add On" || item.name === "CAPS CMB Super User - Add On") {
            showPublish = true;
        }
    });

    return showPublish;
}