"use strict";

var CAPS = CAPS || {};
CAPS.Facility = CAPS.Facility || {
    SHOW_LOCK_ASYNC_COMPLETED: false,
    SHOW_LOCK_BUTTON: false,
};

CAPS.Facility.ShowLockEnrolmentProjections = function (primaryControl) {
    var formContext = primaryControl;
    //get facility Id
    var id = formContext.data.entity.getId().replace("{", "").replace("}", "");
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
            let confirmStrings = { text: "This will lock your future enrolment projections, allowing them to be used in utilization reporting.  This action cannot be undone.  Click OK to continue or Cancel to exit.", title: "Lock Enrolment Projections" };
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

CAPS.Facility.UpdateCapacityReporting = function (primaryControl) {
    var formContext = primaryControl;
    //get facility Id
    var id = formContext.data.entity.getId().replace("{", "").replace("}", "");

    var promises = [];
    //update to statuscode == 200,870,001 (submitted) and statecode eq 1 (inactive)
    /*
                "statecode": 1,
            "statuscode": 200870001
    */
    var data =
        {
            "caps_markassubmitted": true
        };

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
                    
                }
                , function (error) {
                    //Close Popup
                    Alert.hide();
                    Xrm.Navigation.openErrorDialog({ message: error.message });
                }
            );

        },
        function (error) {
            console.log(error.message);
            // handle error conditions
        }
    );
        //
        //UseFutureForUtilization
}