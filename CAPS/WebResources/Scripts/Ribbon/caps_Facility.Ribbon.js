"use strict";

var CAPS = CAPS || {};
CAPS.Facility = CAPS.Facility || {
    SHOW_LOCK_ASYNC_COMPLETED: false,
    SHOW_LOCK_BUTTON: false,
    SHOW_UNLOCK_ASYNC_COMPLETED: false,
    SHOW_UNLOCK_BUTTON: false,
};

/*Function to determine if lock enrolment projections button should be displayed.*/
CAPS.Facility.ShowLockEnrolmentProjections = function (primaryControl) {
    var formContext = primaryControl;
    //get facility Id
    var id = formContext.data.entity.getId().replace("{", "").replace("}", "");
    var useFutureForUtilization = formContext.getAttribute("caps_usefutureforutilization").getValue();

    if (useFutureForUtilization) return false;

    //Check Users role and for draft enrolment projections
    debugger;
    var userRoles = Xrm.Utility.getGlobalContext().userSettings.roles;

    var showButton = false;

    userRoles.forEach(function hasFinancialDirectorRole(item, index) {
        if (item.name === "CAPS School District User") {
            showButton = true;
        }
    });

    if (!showButton) return false;

    if (CAPS.Facility.SHOW_LOCK_ASYNC_COMPLETED) {
        return CAPS.Facility.SHOW_LOCK_BUTTON;
    }

    Xrm.WebApi.retrieveMultipleRecords("caps_enrolmentprojections_sd", "?$select=caps_enrolmentprojections_sdid&$filter=caps_Facility/caps_facilityid eq " + id + " and statuscode eq 1").then(
        function success(result) {
            debugger;
            CAPS.Facility.SHOW_LOCK_ASYNC_COMPLETED = true;
            if (result.entities.length > 0) {
                CAPS.Facility.SHOW_LOCK_BUTTON = true;
            }

            if (CAPS.Facility.SHOW_LOCK_BUTTON) {
                formContext.ui.refreshRibbon();
            }
        }
        , function (error) {
            CAPS.Facility.SHOW_LOCK_ASYNC_COMPLETED = true;
            CAPS.Facility.SHOW_LOCK_BUTTON = false;
            Xrm.Navigation.openAlertDialog({ text: error.message });
        }
    );
}

/*Function to confirm using enrolment projections for capacity*/
CAPS.Facility.LockEnrolmentProjections = function (primaryControl) {
    var formContext = primaryControl;

    //If dirty, then save and call again
    if (formContext.data.entity.getIsDirty() || formContext.ui.getFormType() === 1) {
        formContext.data.save({ saveMode: 1 }).then(function success(result) { CAPS.Facility.LockEnrolmentProjections(primaryControl); });
    }
    else {
        //check enrolment complete flag is Completed (1)
        if (formContext.getAttribute("caps_outstandingenrolmentprojection").getValue()) {
            //add are you sure question
            let confirmStrings = { text: "This will lock your future enrolment projections, allowing them to be used in utilization reporting.  Click OK to continue or Cancel to exit.", title: "Lock Enrolment Projections" };
            let confirmOptions = { height: 200, width: 450 };
            Xrm.Navigation.openConfirmDialog(confirmStrings, confirmOptions).then(
                function (success) {
                    if (success.confirmed) {
                        CAPS.Facility.UpdateCapacityReporting(primaryControl);
                    }
                });
        }
        else {
            var alertStrings = { confirmButtonLabel: "OK", text: "Enrolment Projections for this facility are not marked as completed.", title: "Lock Enrolment Projections" };
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

/*Function to determine if unlock enrolment projections button should be displayed.*/
CAPS.Facility.ShowUnlockEnrolmentProjections = function (primaryControl) {
    var formContext = primaryControl;
    //get facility Id
    var id = formContext.data.entity.getId().replace("{", "").replace("}", "");

    //Get future for utilization
    var useFutureForUtilization = formContext.getAttribute("caps_usefutureforutilization").getValue();

    if (!useFutureForUtilization) return false;

    //Check Users role and for draft enrolment projections
    debugger;
    var userRoles = Xrm.Utility.getGlobalContext().userSettings.roles;

    var showButton = false;

    userRoles.forEach(function hasFinancialDirectorRole(item, index) {
        if (item.name === "CAPS School District User") {
            showButton = true;
        }
    });

    if (!showButton) return false;

    if (CAPS.Facility.SHOW_UNLOCK_ASYNC_COMPLETED) {
        return CAPS.Facility.SHOW_UNLOCK_BUTTON;
    }

    Xrm.WebApi.retrieveMultipleRecords("caps_enrolmentprojections_sd", "?$select=caps_enrolmentprojections_sdid&$filter=caps_Facility/caps_facilityid eq " + id + " and statuscode eq 200870001").then(
        function success(result) {
            debugger;
            CAPS.Facility.SHOW_UNLOCK_ASYNC_COMPLETED = true;
            if (result.entities.length > 0) {
                CAPS.Facility.SHOW_UNLOCK_BUTTON = true;
            }

            if (CAPS.Facility.SHOW_UNLOCK_BUTTON) {
                formContext.ui.refreshRibbon();
            }
        }
        , function (error) {
            CAPS.Facility.SHOW_UNLOCK_ASYNC_COMPLETED = true;
            CAPS.Facility.SHOW_UNLOCK_BUTTON = false;
            Xrm.Navigation.openAlertDialog({ text: error.message });
        }
    );
}

/*Function to confirm unlocking enrolment projections for editing*/
CAPS.Facility.UnlockEnrolmentProjections = function (primaryControl) {
    var formContext = primaryControl;

    //If dirty, then save and call again
    if (formContext.data.entity.getIsDirty() || formContext.ui.getFormType() === 1) {
        formContext.data.save({ saveMode: 1 }).then(function success(result) { CAPS.Facility.LockEnrolmentProjections(primaryControl); });
    }
    else {
        //check enrolment complete flag is Completed (1)
        if (formContext.getAttribute("caps_usefutureforutilization").getValue()) {
            //add are you sure question
            let confirmStrings = { text: "This will unlock your future enrolment projections, allowing them to be edited.  Click OK to continue or Cancel to exit.", title: "Unlock Enrolment Projections" };
            let confirmOptions = { height: 200, width: 450 };
            Xrm.Navigation.openConfirmDialog(confirmStrings, confirmOptions).then(
                function (success) {
                    if (success.confirmed) {
                        CAPS.Facility.ReverseCapacityReporting(primaryControl);
                    }
                });
        }
        else {
            var alertStrings = { confirmButtonLabel: "OK", text: "Unable to unlock Enrolment Projections for this facility.", title: "Unlock Enrolment Projections" };
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
CAPS.Facility.UpdateCapacityReporting = function (primaryControl) {
    var formContext = primaryControl;
    //get facility Id
    var id = formContext.data.entity.getId().replace("{", "").replace("}", "");

    var promises = [];
    //update to statuscode == 200,870,001 (submitted) and statecode eq 1 (inactive)

    var data =
        {
            "caps_markassubmitted": true
        };

    Xrm.Utility.showProgressIndicator("Updating Enrolment Projections, please wait...");

    //update all future enrolment projection records
    Xrm.WebApi.retrieveMultipleRecords("caps_enrolmentprojections_sd", "?$select=caps_enrolmentprojections_sdid&$filter=caps_Facility/caps_facilityid eq " + id + " and statuscode eq 1").then(
        function success(result) {
            for (var i = 0; i < result.entities.length; i++) {
                var recordId = result.entities[i].caps_enrolmentprojections_sdid;

                //Update the record
                promises.push(Xrm.WebApi.updateRecord("caps_enrolmentprojections_sd", recordId, data));
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

/*Function to set all enrolment projections to draft and toggle Use Future for Utilization to update capacity reporting*/
CAPS.Facility.ReverseCapacityReporting = function (primaryControl) {
    var formContext = primaryControl;
    //get facility Id
    var id = formContext.data.entity.getId().replace("{", "").replace("}", "");

    var promises = [];
    //update to statuscode == 200,870,001 (submitted) and statecode eq 1 (inactive)

    var data =
        {
            "caps_markassubmitted": false
    };

    Xrm.Utility.showProgressIndicator("Updating Enrolment Projections, please wait...");

    //update all future enrolment projection records
    Xrm.WebApi.retrieveMultipleRecords("caps_enrolmentprojections_sd", "?$select=caps_enrolmentprojections_sdid&$filter=caps_Facility/caps_facilityid eq " + id + " and statuscode eq 200870001").then(
        function success(result) {
            for (var i = 0; i < result.entities.length; i++) {
                var recordId = result.entities[i].caps_enrolmentprojections_sdid;

                //Update the record
                promises.push(Xrm.WebApi.updateRecord("caps_enrolmentprojections_sd", recordId, data));
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

/*
Function to check if the current user has CAPS CMB User Role.
*/
CAPS.Facility.IsMinistryUser = function () {
    var userRoles = Xrm.Utility.getGlobalContext().userSettings.roles;

    var showButton = false;

    userRoles.forEach(function hasFinancialDirectorRole(item, index) {
        if (item.name === "CAPS CMB User") {
            showButton = true;
        }
    });

    return showButton;
}