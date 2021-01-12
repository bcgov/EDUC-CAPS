"use strict";

var CAPS = CAPS || {};
CAPS.COA = CAPS.COA || {};


/**
 * Function to determine when the Close button should be shown.
 * @param {any} primaryControl primary control
 * @returns {boolean} true if should be shown, otherwise false.
 */
CAPS.COA.ShowClose = function (primaryControl) {
    debugger;
    var formContext = primaryControl;

    //TODO: check current state of coa, if approved (200,870,001)
    if (formContext.getAttribute("statuscode").getValue() !== 200870001) {
        return false;
    }

    var userRoles = Xrm.Utility.getGlobalContext().userSettings.roles;

    var showButton = false;

    userRoles.forEach(function hasFinancialDirectorRole(item, index) {
        if (item.name === "CAPS Financial Director Ministry User - Add On") {
            showButton = true;
        }
    });

    return showButton;
}

/**
 * Function to close the coa record.
 * @param {any} primaryControl primary control
 */
CAPS.COA.Close = function (primaryControl) {
    var formContext = primaryControl;
    //Change status to DRAFT
    let confirmStrings = { text: "This will close the COA in CAPS, it will not close the COA in the Ministry of Finance's system.  This action cannot be undone.  Click OK to continue or Cancel to exit.", title: "Close Confirmation" };
    let confirmOptions = { height: 200, width: 450 };
    Xrm.Navigation.openConfirmDialog(confirmStrings, confirmOptions).then(
        function (success) {
            if (success.confirmed) {
                debugger;
                formContext.getAttribute("statecode").setValue(1);
                formContext.getAttribute("statuscode").setValue(200870004);
                formContext.data.entity.save();
            }

        });
}
