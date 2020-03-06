"use strict";

var CAPS = CAPS || {};
CAPS.Project = CAPS.Project ||
    {
    GLOBAL_FORM_CONTEXT: null,
    PROJECT_ARRAY: null
    };


/**
 * Called from Project Homepage (main list view).  Takes in a list of records.  This function
 * calls ShowSubmissionWindow to show the submission selection popup.
 * @param {any} selectedControlIds
 */
CAPS.Project.AddListToSubmission = function (selectedControlIds) {
    
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
                var alertStrings = { confirmButtonLabel: "OK", text: "One or more projects can't be added to the submisson.", title: "Error" };
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
            return false;
        }
    );
}

/**
 * Called from Project Form.  Takes in the form context as primary Control.  This function
 * calls ShowSubmissionWindow to show the submission selection popup.
 * @param {any} primaryControl
 */
CAPS.Project.AddToSubmission = function (primaryControl) {
    var formContext = primaryControl;

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
 * Opens a modal window with a submission drop down
 * @param {any} selectedControlIds
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
            //Close Popup
            Alert.hide();
        }
        , function (error) { }
    );
}



