"use strict";

var CAPS = CAPS || {};
CAPS.DrawRequest = CAPS.DrawRequest || {};

//Ribbon ShowHide logic
CAPS.DrawRequest.ShowReadytoSubmitButton = function (primaryControl) {
    var formContext = primaryControl;

    if (formContext.getAttribute("statecode").getValue() !== 0) {
        return false;
    }

    var showButton = false;
    var remainingBalance = formContext.getAttribute("caps_remainingdrawrequestbalance");

    if (remainingBalance) {
    var remainingBalanceValue = remainingBalance.getValue();
        if (remainingBalanceValue >= 0) {
            showButton = true;
        }
        else {
            showButton = false;
        }
    };

return showButton;

}

/*
// Function to check if the deadline is past and manage form notifications. This Function is used on Batch form.
CAPS.DrawRequest.remainingBalanceExceeded = function (executionContext) {
    var formContext = executionContext.getFormContext();
    var remainingBalance = formContext.getAttribute("caps_remainingdrawrequestbalance");
    var statecode = formContext.getAttribute("statecode").getValue();

    if (remainingBalance) {
        var remainingBalanceValue = remainingBalance.getValue();

        if (remainingBalanceValue !== null) {

            if (remainingBalanceValue < 0 && statecode == 0) {
                formContext.ui.setFormNotification("The requested amount exceeds the remaining project funds. Reduce the amount to be able to update to ready to submit.", "WARNING", "exceededBalance");
                return true;
            }
            formContext.ui.clearFormNotification("exceededBalance");
            return false;
        }
    }
}*/
