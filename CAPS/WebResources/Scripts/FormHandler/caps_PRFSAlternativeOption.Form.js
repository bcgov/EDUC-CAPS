"use strict";

var CAPS = CAPS || {};
CAPS.PRFSOption = CAPS.PRFSOption || {};

const FORM_STATE = {
    UNDEFINED: 0,
    CREATE: 1,
    UPDATE: 2,
    READ_ONLY: 3,
    DISABLED: 4,
    BULK_EDIT: 6
};

/***
Main Function for PRFS Alternative Option.  This function calls all other form functions.
**/
CAPS.PRFSOption.onLoad = function (executionContext) {
    debugger;
    // Set variables
    var formContext = executionContext.getFormContext();
    var formState = formContext.ui.getFormType();

    //On Create
    if (formState === FORM_STATE.CREATE) {
        //Set School District based on User
        CAPS.PRFSOption.DefaultSchoolDistrict(executionContext);
    }

    CAPS.PRFSOption.ToggleScheduleBFields(executionContext);
    formContext.getAttribute("caps_projecttype").addOnChange(CAPS.PRFSOption.ToggleScheduleBFields);

    formContext.getAttribute("caps_submissioncategory").addOnChange(CAPS.PRFSOption.ClearProjectType);

    formContext.getAttribute("caps_requiresscheduleb").addOnChange(CAPS.PRFSOption.ToggleScheduleBTab);
    CAPS.PRFSOption.ToggleScheduleBTab(executionContext);
}

/**
 * Sets the projects School District to the user's business unit's school district if it's set
 * @param {any} formContext the form's form context
 */
CAPS.PRFSOption.DefaultSchoolDistrict = function (executionContext) {
    var formContext = executionContext.getFormContext();
    //get Current User ID
    var userSettings = Xrm.Utility.getGlobalContext().userSettings;

    var userId = userSettings.userId;

    //Get BU from User record
    Xrm.WebApi.retrieveRecord("systemuser", userId, "?$select=_businessunitid_value").then(
        function success(result) {

            var businessUnit = result["_businessunitid_value"];

            //Now get Business Unit's School District if it exists
            Xrm.WebApi.retrieveRecord("businessunit", businessUnit, "?$select=_caps_schooldistrict_value").then(
                function success(resultBU) {

                    var sdID = resultBU["_caps_schooldistrict_value"];
                    if (sdID !== null) {
                        var sdName = resultBU["_caps_schooldistrict_value@OData.Community.Display.V1.FormattedValue"];
                        var sdType = resultBU["_caps_schooldistrict_value@Microsoft.Dynamics.CRM.lookuplogicalname"];

                        formContext.getAttribute("caps_hostschooldistrict").setValue([{ id: sdID, name: sdName, entityType: sdType }]);
                        formContext.getAttribute("caps_schooldistrict").setValue([{ id: sdID, name: sdName, entityType: sdType }]);
                        CAPS.PRFSOption.ToggleScheduleBFields(executionContext);
                    }
                },
                function (error) {
                    console.log(error.message);
                    // handle error conditions
                }
            );
        },
        function (error) {
            console.log(error.message);
            // handle error conditions
        }
    );
}

/***
Show's and hides the appropriate schedule B fields on PRFS based on Project Type.
*/
CAPS.PRFSOption.ToggleScheduleBFields = function (executionContext) {
    var formContext = executionContext.getFormContext();
    //Get Schedule B Type field on Project Type
    var projectType = formContext.getAttribute("caps_projecttype").getValue();

    var hostSD = formContext.getAttribute("caps_schooldistrict").getValue();

    if (hostSD !== null && !hostSD[0].name.includes("SD93")) {
        formContext.getControl("caps_hostschooldistrict").setVisible(false);
    }

    if (projectType !== null) {
        Xrm.WebApi.retrieveRecord("caps_projecttype", projectType[0].id, "?$select=caps_budgetcalculationtype").then(
        function success(result) {
            debugger;
            var calcType = result.caps_budgetcalculationtype;
            //New = 200,870,000
            //Replacement = 200,870,001
            //Partial Seismic = 200,870,004
            //Seismic = 200,870,005
            if (calcType === 200870000 || calcType === 200870001) {
                //show NLC
                formContext.getControl("caps_includenlc").setVisible(true);
            }
            else {
                //hide and set NLS to false
                formContext.getControl("caps_includenlc").setVisible(false);
                formContext.getAttribute("caps_includenlc").setValue(null);
            }
            if (calcType === 200870004 || calcType === 200870005) {
                //formContext.getControl("caps_seismicupgradespaceallocation").setVisible(true);
                //formContext.getControl("caps_seismicprojectidentificationreportfees").setVisible(true);
                formContext.getControl("caps_constructioncostsspir").setVisible(true);
                //formContext.getControl("caps_constructioncostsspiradjustments").setVisible(true);
                //formContext.getControl("caps_constructioncostsnonstructuralseismicup").setVisible(true);
            }
            else {
                //formContext.getControl("caps_seismicupgradespaceallocation").setVisible(false);
                //formContext.getControl("caps_seismicprojectidentificationreportfees").setVisible(false);
                formContext.getControl("caps_constructioncostsspir").setVisible(false);
                //formContext.getControl("caps_constructioncostsspiradjustments").setVisible(false);
                //formContext.getControl("caps_constructioncostsnonstructuralseismicup").setVisible(false);

                //formContext.getAttribute("caps_seismicupgradespaceallocation").setValue(null);
                //formContext.getAttribute("caps_seismicprojectidentificationreportfees").setValue(null);
                formContext.getAttribute("caps_constructioncostsspir").setValue(null);
                //formContext.getAttribute("caps_constructioncostsspiradjustments").setValue(null);
                //formContext.getAttribute("caps_constructioncostsnonstructuralseismicup").setValue(null);
            }
        },
        function (error) {
            console.log(error.message);
            // handle error conditions
        });
    }
}

/***
Shows/Hides the Schedule B tab if based on caps_requiresscheduleb field value.
*/
CAPS.PRFSOption.ToggleScheduleBTab = function (executionContext) {
    var formContext = executionContext.getFormContext();

    if (formContext.getAttribute("caps_requiresscheduleb").getValue()) {
        formContext.ui.tabs.get("tab_budget").setVisible(true);
    }
    else {
        formContext.ui.tabs.get("tab_budget").setVisible(false);
    }
}

/****
Clears the Project Type field when Submission Category is changed.
*/
CAPS.PRFSOption.ClearProjectType = function (executionContext) {
    var formContext = executionContext.getFormContext();

    formContext.getAttribute("caps_projecttype").setValue(null);
}