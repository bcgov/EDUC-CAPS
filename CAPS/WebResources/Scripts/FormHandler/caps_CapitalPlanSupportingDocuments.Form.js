"use strict";

var CAPS = CAPS || {};
CAPS.CPSD = CAPS.CPSD || {};

/***
Main function for CPSD.  This function calls all other functions.
**/
CAPS.CPSD.onLoad = function (executionContext) {
    var formContext = executionContext.getFormContext();

    formContext.getAttribute("caps_capitalbylaw").addOnChange(CAPS.CPSD.FlagBylawAsUploaded);
}

/***
Called onLoad and onChange of Bylaw field.  This function sets a Bylaw Uploaded? field to yes or no.
**/
CAPS.CPSD.FlagBylawAsUploaded = function (executionContext) {
    var formContext = executionContext.getFormContext();

    var bylaw = formContext.getAttribute("caps_capitalbylaw").getValue();

    if (bylaw !== null) {
        formContext.getAttribute("caps_bylawuploaded").setValue(true);
        //formContext.data.save({ saveMode: 1 });
    }
    else {
        formContext.getAttribute("caps_bylawuploaded").setValue(false);
        //formContext.data.save({ saveMode: 1 });
    }
}

