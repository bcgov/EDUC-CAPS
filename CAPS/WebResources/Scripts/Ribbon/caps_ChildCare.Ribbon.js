"use strict";

var CAPS = CAPS || {};
CAPS.ChildCareFacilityRibbon = CAPS.ChildCareFacilityRibbon || {
    SHOW_LOCK_ASYNC_COMPLETED: false,
    SHOW_LOCK_BUTTON: false,
    SHOW_UNLOCK_ASYNC_COMPLETED: false,
    SHOW_UNLOCK_BUTTON: false,
};

/*Function to determine if lock enrolment projections button should be displayed.*/
CAPS.ChildCareFacilityRibbon.ShowLockEnrolmentProjections = function (primaryControl) {
    debugger;
    var formContext = primaryControl;
    //get child care facility Id
    if (formContext.data.entity.getId() !== '') {
        var id = formContext.data.entity.getId().replace("{", "").replace("}", "");
        var useFutureForUtilization = formContext.getAttribute("caps_usefutureforutilization").getValue();

        if (useFutureForUtilization) return false;

        //Check Users role and for draft enrolment projections

        var userRoles = Xrm.Utility.getGlobalContext().userSettings.roles;

        var showButton = false;

        userRoles.forEach(function hasFinancialDirectorRole(item, index) {

            if (item.name === "CAPS School District User" || item.name === "CAPS CMB Super User - Add On") {
                showButton = true;
            }
        });

        if (!showButton) return false;

        if (CAPS.ChildCareFacilityRibbon.SHOW_LOCK_ASYNC_COMPLETED) {
            return CAPS.ChildCareFacilityRibbon.SHOW_LOCK_BUTTON;
        }

        Xrm.WebApi.retrieveMultipleRecords("caps_childcareenrolmentprojection", "?$select=caps_childcareenrolmentprojectionid&$filter=caps_ChildCareFacility/caps_childcareid eq " + id + " and statuscode eq 1").then(
            function success(result) {

                CAPS.ChildCareFacilityRibbon.SHOW_LOCK_ASYNC_COMPLETED = true;
                if (result.entities.length > 0) {
                    CAPS.ChildCareFacilityRibbon.SHOW_LOCK_BUTTON = true;
                }

                if (CAPS.ChildCareFacilityRibbon.SHOW_LOCK_BUTTON) {
                    formContext.ui.refreshRibbon();
                }
            }
            , function (error) {
                CAPS.ChildCareFacilityRibbon.SHOW_LOCK_ASYNC_COMPLETED = true;
                CAPS.ChildCareFacilityRibbon.SHOW_LOCK_BUTTON = false;
                Xrm.Navigation.openAlertDialog({ text: error.message });
            }
        );
    }
}

/*Function to determine if unlock enrolment projections button should be displayed.*/
CAPS.ChildCareFacilityRibbon.ShowUnlockEnrolmentProjections = function (primaryControl) {
    var formContext = primaryControl;

    if (formContext.data.entity.getId() != '') {
        //get facility Id
        var id = formContext.data.entity.getId().replace("{", "").replace("}", "");

        //Get future for utilization
        var useFutureForUtilization = formContext.getAttribute("caps_usefutureforutilization").getValue();

        if (!useFutureForUtilization) return false;

        //Check Users role and for draft enrolment projections

        var userRoles = Xrm.Utility.getGlobalContext().userSettings.roles;

        var showButton = false;

        userRoles.forEach(function hasFinancialDirectorRole(item, index) {
            if (item.name === "CAPS School District User" || item.name === "CAPS CMB Super User - Add On") {
                showButton = true;
            }
        });

        if (!showButton) return false;

        if (CAPS.ChildCareFacilityRibbon.SHOW_UNLOCK_ASYNC_COMPLETED) {
            return CAPS.ChildCareFacilityRibbon.SHOW_UNLOCK_BUTTON;
        }

        Xrm.WebApi.retrieveMultipleRecords("caps_childcareenrolmentprojection", "?$select=caps_childcareenrolmentprojectionid&$filter=caps_ChildCareFacility/caps_childcareid eq " + id + " and statuscode eq 2").then(
            function success(result) {

                CAPS.ChildCareFacilityRibbon.SHOW_UNLOCK_ASYNC_COMPLETED = true;
                if (result.entities.length > 0) {
                    CAPS.ChildCareFacilityRibbon.SHOW_UNLOCK_BUTTON = true;
                }

                if (CAPS.ChildCareFacilityRibbon.SHOW_UNLOCK_BUTTON) {
                    formContext.ui.refreshRibbon();
                }
            }
            , function (error) {
                CAPS.ChildCareFacilityRibbon.SHOW_UNLOCK_ASYNC_COMPLETED = true;
                CAPS.ChildCareFacilityRibbon.SHOW_UNLOCK_BUTTON = false;
                Xrm.Navigation.openAlertDialog({ text: error.message });
            }
        );
    }
}


/*Function to confirm using enrolment projections for capacity*/
CAPS.ChildCareFacilityRibbon.LockEnrolmentProjections = function (primaryControl) {
    
    var formContext = primaryControl;

    //If dirty, then save and call again
    if (formContext.data.entity.getIsDirty() || formContext.ui.getFormType() === 1) {
        formContext.data.save({ saveMode: 1 }).then(function success(result) { CAPS.ChildCareFacilityRibbon.LockEnrolmentProjections(primaryControl); });
    }
    else {
        //check enrolment complete flag is Completed (1)
        if (formContext.getAttribute("caps_enrolmentprojection").getValue()) {
            //add are you sure question
            let confirmStrings = { text: "This will lock your future child care enrolment projections, allowing them to be used in utilization reporting.  Click OK to continue or Cancel to exit.", title: "Lock Child Care Enrolment Projections" };
            let confirmOptions = { height: 200, width: 450 };
            Xrm.Navigation.openConfirmDialog(confirmStrings, confirmOptions).then(
                function (success) {
                   
                    if (success.confirmed) {
                        CAPS.ChildCareFacilityRibbon.UpdateCapacityReporting(primaryControl);
                    }
                });
        }
        else {
            var alertStrings = { confirmButtonLabel: "OK", text: "Child Care Enrolment Projections for this child care facility are not marked as completed.", title: "Lock Child Care Enrolment Projections" };
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
    }

}

/*Function to confirm unlocking enrolment projections for editing*/
CAPS.ChildCareFacilityRibbon.UnlockEnrolmentProjections = function (primaryControl) {
    debugger;
    var formContext = primaryControl;

    //If dirty, then save and call again
    if (formContext.data.entity.getIsDirty() || formContext.ui.getFormType() === 1) {
        formContext.data.save({ saveMode: 1 }).then(function success(result) { CAPS.ChildCareFacilityRibbon.LockEnrolmentProjections(primaryControl); });
    }
    else {
        //check enrolment complete flag is Completed (1)
        if (formContext.getAttribute("caps_usefutureforutilization").getValue()) {
            //add are you sure question
            let confirmStrings = { text: "This will unlock your future child care enrolment projections, allowing them to be edited.  Click OK to continue or Cancel to exit.", title: "Unlock Child Care Enrolment Projections" };
            let confirmOptions = { height: 200, width: 450 };
            Xrm.Navigation.openConfirmDialog(confirmStrings, confirmOptions).then(
                function (success) {
                    if (success.confirmed) {
                        CAPS.ChildCareFacilityRibbon.ReverseCapacityReporting(primaryControl);
                    }
                });
        }
        else {
            var alertStrings = { confirmButtonLabel: "OK", text: "Unable to unlock Child Care Enrolment Projections for this facility.", title: "Unlock Child Care Enrolment Projections" };
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
    }

}

/*Function to set all enrolment projections to submitted and toggle Use Future for Utilization to update capacity reporting*/
CAPS.ChildCareFacilityRibbon.UpdateCapacityReporting = function (primaryControl) {
    
    var formContext = primaryControl;
    if (formContext.data.entity.getId() !== '') {
        //get facility Id
        var id = formContext.data.entity.getId().replace("{", "").replace("}", "");

        var promises = [];
        //update to statuscode == 200,870,001 (submitted) and statecode eq 1 (inactive)

        var data =
        {
            "caps_markedassubmitted": true
        };

        Xrm.Utility.showProgressIndicator("Updating Child Care Enrolment Projections, please wait...");

        //update all future enrolment projection records
        Xrm.WebApi.retrieveMultipleRecords("caps_childcareenrolmentprojection", "?$select=caps_childcareenrolmentprojectionid&$filter=caps_ChildCareFacility/caps_childcareid eq " + id + " and statuscode eq 1").then(
            function success(result) {

                for (var i = 0; i < result.entities.length; i++) {

                    var recordId = result.entities[i].caps_childcareenrolmentprojectionid;

                    //Update the record
                    promises.push(Xrm.WebApi.updateRecord("caps_childcareenrolmentprojection", recordId, data));
                }
                Promise.all(promises).then(
                    function (results) {

                        //mark as locked
                        formContext.data.refresh();
                        formContext.getAttribute("caps_usefutureforutilization").setValue(true);
                        formContext.data.entity.save();
                        Xrm.Utility.closeProgressIndicator();
                    }
                    , function (error) {

                        Xrm.Utility.closeProgressIndicator();
                        Xrm.Navigation.openErrorDialog({ message: error.message });
                    }
                );

            },
            function (error) {
                Xrm.Utility.closeProgressIndicator();
                console.log(error.message);
                // handle error conditions
            }
        );
        //
        //UseFutureForUtilization
    }
}

/*Function to set all enrolment projections to draft and toggle Use Future for Utilization to update capacity reporting*/
CAPS.ChildCareFacilityRibbon.ReverseCapacityReporting = function (primaryControl) {
    var formContext = primaryControl;
    //get facility Id
    if (formContext.data.entity.getId() !== '') {
        var id = formContext.data.entity.getId().replace("{", "").replace("}", "");

        var promises = [];
        //update to statuscode == 200,870,001 (submitted) and statecode eq 1 (inactive)

        var data =
        {
            "caps_markedassubmitted": false
        };

        Xrm.Utility.showProgressIndicator("Updating Child Care Enrolment Projections, please wait...");

        //update all future enrolment projection records
        Xrm.WebApi.retrieveMultipleRecords("caps_childcareenrolmentprojection", "?$select=caps_childcareenrolmentprojectionid&$filter=caps_ChildCareFacility/caps_childcareid eq " + id + " and statuscode eq 2").then(
            function success(result) {
                for (var i = 0; i < result.entities.length; i++) {
                    var recordId = result.entities[i].caps_childcareenrolmentprojectionid;

                    //Update the record
                    promises.push(Xrm.WebApi.updateRecord("caps_childcareenrolmentprojection", recordId, data));
                }
                Promise.all(promises).then(
                    function (results) {
                        //mark as locked
                        formContext.data.refresh();
                        formContext.getAttribute("caps_usefutureforutilization").setValue(false);
                        formContext.data.entity.save();
                        Xrm.Utility.closeProgressIndicator();
                    }
                    , function (error) {
                        Xrm.Utility.closeProgressIndicator();
                        Xrm.Navigation.openErrorDialog({ message: error.message });
                    }
                );

            },
            function (error) {
                Xrm.Utility.closeProgressIndicator();
                console.log(error.message);
                // handle error conditions
            }
        );
        //
        //UseFutureForUtilization
    }
}

/*
Function to check if the current user has CAPS CMB User Role.
*/
CAPS.ChildCareFacilityRibbon.IsMinistryUser = function () {
    var userRoles = Xrm.Utility.getGlobalContext().userSettings.roles;

    var showButton = false;

    userRoles.forEach(function hasFinancialDirectorRole(item, index) {
        if (item.name === "CAPS CMB User") {
            showButton = true;
        }
    });

    return showButton;
}