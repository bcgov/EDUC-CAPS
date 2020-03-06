"use strict";

/* INCLUDE CAPS.Common.js */

var CAPS = CAPS || {};
CAPS.Project = CAPS.Project || { GLOBAL_FORM_CONTEXT: null};



const FORM_STATE = {
    UNDEFINED: 0,
    CREATE: 1,
    UPDATE: 2,
    READ_ONLY: 3,
    DISABLED: 4,
    BULK_EDIT: 6
};

const TIMELINE_TAB = "tab_timeline";
const GENERAL_TAB = "General";
//const CAPITAL_EXPENDITURE_TAB = "tab_Capital_Expenditure";
const COST_MISSMATCH_NOTIFICATION = "Cost_Missmatch_Notification";
const NO_FACILITY_NOTIFICATION = "No_Facility_Notification";

/**
 * Main function for Project.  This function calls all other form functions and registers onChange and onLoad events
 * @param {any} executionContext
 */
CAPS.Project.onLoad = function (executionContext) {
    
    var formContext = executionContext.getFormContext();
    CAPS.Project.GLOBAL_FORM_CONTEXT = formContext;

    CAPS.Project.ShowHideRelevantTabs(formContext);
    
    //Check if Expenditure Validation Required
    if (formContext.getAttribute("caps_submissioncategoryrequirecostallocation").getValue() === true) {
        
        formContext.getAttribute("caps_totalprojectcost").addOnChange(CAPS.Project.ValidateExpenditureDistribution);

        //caps_sumestimatedyearlyexpenditures caps_totalestimatedprojectcost
        formContext.getAttribute("caps_sumestimatedyearlyexpenditures").addOnChange(CAPS.Project.ValidateExpenditureDistribution);

        CAPS.Project.ValidateExpenditureDistribution(executionContext);
    }

    //Only call for SEP and CNCP!
    //TODO: Have added a flag to Project Submission, if we are keeping then add calculated field to Project and check here
    CAPS.Project.addFacilitiesEventListener(0);

    //Set School District based on User
    CAPS.Project.DefaultLookupIfSingle(formContext, "caps_schooldistrict", "caps_schooldistrict", "caps_schooldistrictid", "caps_name");

}

/**
 * This function shows either the General tab for new Projects or the relevant tab from the related Submission Category for existing projects.
 * It also calls a function to turn off any field validation for any tab not shown.
 * @param {any} formContext
 */
CAPS.Project.ShowHideRelevantTabs = function (formContext) {
    //check form state
    var formState = formContext.ui.getFormType();

    if (formState === FORM_STATE.CREATE) {
        // turn off all mandatory fields
        var createTabsToDisregard = [GENERAL_TAB, TIMELINE_TAB];
        CAPS.Project.RemoveRequirement(formContext, createTabsToDisregard);
    }

    if (formState === FORM_STATE.UPDATE || formState === FORM_STATE.DISABLED || formState === FORM_STATE.READ_ONLY) {
        //Hide General Tab
        formContext.ui.tabs.get(GENERAL_TAB).setVisible(false);

        //Show only appropriate tab
        var submissionCategoryTabName = formContext.getAttribute("caps_submissioncategorytabname").getValue();

        //Remove all mandatory fields and show relevant tab
        var tabsToDisregard = [submissionCategoryTabName, TIMELINE_TAB];
        CAPS.Project.RemoveRequirement(formContext, tabsToDisregard);

        formContext.ui.tabs.get(submissionCategoryTabName).setVisible(true);

        //if capital expense needs allocating, show the tab
        //if (formContext.getAttribute("caps_submissioncategoryrequirecostallocation").getValue() === true) {
        //    formContext.ui.tabs.get(CAPITAL_EXPENDITURE_TAB).setVisible(true);
        //}
    }
}

/**
 * This function turns off all field requirements for any field except those in the tabsToDisregard array
 * @param {any} formContext
 * @param {any} tabsToDisregard - array of tab names to disregard
 */
CAPS.Project.RemoveRequirement = function(formContext, tabsToDisregard){
    //loop through tabs
    formContext.ui.tabs.forEach(function (tab, i) {
        //loop through sections
        if (!tabsToDisregard.includes(tab.getName())) {
            tab.sections.forEach(function (section, j) {
                section.controls.forEach(function (control, k) {
                    control.getAttribute().setRequiredLevel("none");
                });
            });
        }
    });
}

/**
 * This function compares the total project cost to the sum of the estimated yearly expenditures and shows an error if they don't match.
 * This function is only called if the related Submission Category field Require Cost Allocation is set to Yes.
 * @param {any} executionContext
 */
CAPS.Project.ValidateExpenditureDistribution = function (executionContext) {
    //Only validate if Submission Category requires 10 year plan
    //If numbers don't match, show formContext.getControl(arg).setNotification();
    var formContext = executionContext.getFormContext();

    var totalProjectCost = formContext.getAttribute("caps_totalprojectcost").getValue();
    var sumOfEstimatedExpenditures = formContext.getAttribute("caps_sumestimatedyearlyexpenditures").getValue();

    if (totalProjectCost !== sumOfEstimatedExpenditures) {
        formContext.getControl("caps_totalprojectcost").setNotification('Total Project Cost Not Fully Allocated', COST_MISSMATCH_NOTIFICATION);
    }
    else {
        formContext.getControl("caps_totalprojectcost").clearNotification(COST_MISSMATCH_NOTIFICATION);
    }
}

/**
 * This function waits for the Facilities subgrid to load and adds an event listener to the grid for validating that at least one facility was added.
 * @param {any} loopCount
 */
CAPS.Project.addFacilitiesEventListener = function(loopCount) {
    var gridContext = CAPS.Project.GLOBAL_FORM_CONTEXT.getControl("Facilities - SEP");

    if (loopCount < 5) {
        if (gridContext === null) {
            setTimeout(function () { FACILITIES_EVENT_HANDLER_LOOP_COUNTER++; CAPS.Project.addFacilitiesEventListener(loopCount++); }, 500);
        }

        gridContext.addOnLoad(CAPS.Project.ValidateAtLeaseOneFacility);
    }
}

/**
 * This function validates that at least one facility has been added to the project
 * @param {any} executionContext
 */
CAPS.Project.ValidateAtLeaseOneFacility = function (executionContext) {
    var gridContext = executionContext.getFormContext(); 

    var filteredRecordCount = gridContext.getGrid().getTotalRecordCount();

    if (filteredRecordCount < 1) {
        CAPS.Project.GLOBAL_FORM_CONTEXT.ui.setFormNotification('You must add at least one facility to this project.', 'INFO', NO_FACILITY_NOTIFICATION);
    }
    else {
        CAPS.Project.GLOBAL_FORM_CONTEXT.ui.clearFormNotification(NO_FACILITY_NOTIFICATION);
    }

}


