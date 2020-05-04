"use strict";

/* INCLUDE CAPS.Common.js */

var CAPS = CAPS || {};
CAPS.ProjectHistory = CAPS.ProjectHistory || {
    GLOBAL_FORM_CONTEXT: null,
    PREVENT_AUTO_SAVE: false
};

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

/**
 * Main function for Project History.  This function calls all other form functions.
 * @param {any} executionContext form's execution context
 */
CAPS.ProjectHistory.onLoad = function (executionContext) {
    // Set variables
    var formContext = executionContext.getFormContext();
    CAPS.ProjectHistory.GLOBAL_FORM_CONTEXT = formContext;
    var formState = formContext.ui.getFormType();

    //Show/Hide Tabs
    CAPS.ProjectHistory.ShowHideRelevantTabs(formContext);
}

/**
 * This function shows either the General tab for new Projects or the relevant tab from the related Submission Category for existing projects.
 * It also calls a function to turn off any field validation for any tab not shown.
 * @param {any} formContext form context
 */
CAPS.ProjectHistory.ShowHideRelevantTabs = function (formContext) {
    //check form state
    var formState = formContext.ui.getFormType();

    if (formState === FORM_STATE.CREATE) {
        // turn off all mandatory fields
        var createTabsToDisregard = [GENERAL_TAB, TIMELINE_TAB];
        CAPS.ProjectHistory.RemoveRequirement(formContext, createTabsToDisregard);
    }

    if (formState === FORM_STATE.UPDATE || formState === FORM_STATE.DISABLED || formState === FORM_STATE.READ_ONLY) {
        //Hide General Tab
        formContext.ui.tabs.get(GENERAL_TAB).setVisible(false);

        //Show only appropriate tab
        var submissionCategoryTabNames = formContext.getAttribute("caps_submissioncategorytabname").getValue();
        var arrTabNames = submissionCategoryTabNames.split(", ");

        //Remove all mandatory fields and show relevant tab(s)
        var tabsToDisregard = [TIMELINE_TAB];
        arrTabNames.forEach(function (tabName) {
            tabsToDisregard.push(tabName);
        });


        CAPS.ProjectHistory.RemoveRequirement(formContext, tabsToDisregard);

        arrTabNames.forEach(function (tabName) {
            formContext.ui.tabs.get(tabName).setVisible(true);
        });

    }
}

/**
 * This function turns off all field requirements for any field except those in the tabsToDisregard array.
 * @param {any} formContext form context
 * @param {any} tabsToDisregard - array of tab names to disregard
 */
CAPS.ProjectHistory.RemoveRequirement = function (formContext, tabsToDisregard) {

    //Get array of all fields on tabs to disregard
    var fieldsToShow = [];

    formContext.ui.tabs.forEach(function (tab, i) {
        //loop through sections
        if (tabsToDisregard.includes(tab.getName())) {
            tab.sections.forEach(function (section, j) {
                section.controls.forEach(function (control, k) {
                    //add to array
                    fieldsToShow.push(control.getAttribute().getName());
                });
            });
        }
    });

    //loop through tabs
    formContext.ui.tabs.forEach(function (tab, i) {
        //loop through sections
        if (!tabsToDisregard.includes(tab.getName())) {
            tab.sections.forEach(function (section, j) {
                section.controls.forEach(function (control, k) {
                    //if the field isn't on a shown tab, then remove required flag
                    if (!fieldsToShow.includes(control.getAttribute().getName())) {
                        control.getAttribute().setRequiredLevel("none");
                    }
                });
            });
        }
    });
}