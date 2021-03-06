"use strict";

var CAPS = CAPS || {};
CAPS.ProjectClosure = CAPS.ProjectClosure || {};

const MISSING_FLOOR_PLAN = "Missing_Floor_Plan";
const MISSING_DESIGN_AID_SHEET = "Missing_Design_Aid_Sheet";

/** 
 * Main function for Capital Plan form, this function calls all other form functions.
 * @param {any} executionContext execution context
*/
CAPS.ProjectClosure.onLoad = function (executionContext) {
    debugger;
    // Set variables
    var formContext = executionContext.getFormContext();

    formContext.getAttribute("caps_floorplan").addOnChange(CAPS.ProjectClosure.CheckFloorPlan);
    CAPS.ProjectClosure.CheckFloorPlan(executionContext);

    formContext.getAttribute("caps_designaidsheet").addOnChange(CAPS.ProjectClosure.CheckDesignAidSheet);
    CAPS.ProjectClosure.CheckDesignAidSheet(executionContext);
}

/** 
 * Function to check if a floor plan has been provided and show a notification if it hasn't.
 * @param {any} executionContext execution context
*/
CAPS.ProjectClosure.CheckFloorPlan = function (executionContext) {
    var formContext = executionContext.getFormContext();
    if (!formContext.getAttribute("caps_floorplan").getValue()) {
        formContext.ui.setFormNotification('You are missing a floor plan. Please attach one below before completing this project closure.', 'INFO', MISSING_FLOOR_PLAN);
    }
    else {
        formContext.ui.clearFormNotification(MISSING_FLOOR_PLAN);
    }
}

/** 
 * Function to check if a design aid sheet has been provided and show a notification if it hasn't.
 * @param {any} executionContext execution context
*/
CAPS.ProjectClosure.CheckDesignAidSheet = function (executionContext) {
    var formContext = executionContext.getFormContext();
    if (!formContext.getAttribute("caps_designaidsheet").getValue()) {
        formContext.ui.setFormNotification('You are missing a design aid sheet. Please attach one below before completing this project closure.', 'INFO', MISSING_DESIGN_AID_SHEET);
    }
    else {
        formContext.ui.clearFormNotification(MISSING_DESIGN_AID_SHEET);
    }
}
