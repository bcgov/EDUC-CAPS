"use strict";

/* INCLUDE CAPS.Common.js */

var CAPS = CAPS || {};
CAPS.Project = CAPS.Project || {
    GLOBAL_FORM_CONTEXT: null,
    PREVENT_AUTO_SAVE: false,
    RECORD_JUST_CREATED: false,
    FACILITY_GRID_CONTROL: null
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
//const CAPITAL_EXPENDITURE_TAB = "tab_Capital_Expenditure";
const COST_MISSMATCH_NOTIFICATION = "Cost_Missmatch_Notification";
const NO_FACILITY_NOTIFICATION = "No_Facility_Notification";

/**
 * Main function for Project.  This function calls all other form functions and registers onChange and onLoad events.
 * @param {any} executionContext the form execution context
 */
CAPS.Project.onLoad = function (executionContext) {
    debugger;
    // Set variables
    var formContext = executionContext.getFormContext();
    CAPS.Project.GLOBAL_FORM_CONTEXT = formContext;
    var formState = formContext.ui.getFormType();

    //if record was just created, a full reload is needed
    if (CAPS.Project.RECORD_JUST_CREATED) {
        CAPS.Project.RECORD_JUST_CREATED = false;

        var entityFormOptions = {};
        entityFormOptions["entityName"] = "caps_project";
        entityFormOptions["entityId"] = formContext.data.entity.getId();

        // Open the form.
        Xrm.Navigation.openForm(entityFormOptions).then(
            function (success) {
                console.log(success);
            },
            function (error) {
                console.log(error);
            });
    }

    //Show/Hide Tabs
    CAPS.Project.ShowHideRelevantTabs(formContext);

    
    //On Create
    if (formState === FORM_STATE.CREATE) {
        CAPS.Project.RECORD_JUST_CREATED = true;
        //Set School District based on User
        CAPS.Project.DefaultSchoolDistrict(formContext);

        //add onchange events for create
        formContext.getAttribute("caps_submissioncategory").addOnChange(CAPS.Project.SetProjectTypeValue);
    }
    
    //Check if Expenditure Validation Required
    if (formContext.getAttribute("caps_submissioncategoryrequirecostallocation").getValue() === true) {
        
        formContext.getAttribute("caps_totalprojectcost").addOnChange(CAPS.Project.ValidateExpenditureDistribution);

        //caps_sumestimatedyearlyexpenditures caps_totalestimatedprojectcost
        formContext.getAttribute("caps_totalallocated").addOnChange(CAPS.Project.ValidateExpenditureDistribution);

        CAPS.Project.ValidateExpenditureDistribution(executionContext);
    }

    //Only call for SEP and CNCP!
    //TODO: Have added a flag to Project Submission, if we are keeping then add calculated field to Project and check here
    if (formContext.getAttribute("caps_submissioncategoryallowmultiplefacilities").getValue() === true) {
        CAPS.Project.SetMultipleFacility(executionContext);

        //Add onChange event to caps_multiplefacility
        formContext.getAttribute("caps_multiplefacilities").addOnChange(CAPS.Project.SetMultipleFacility);

        CAPS.Project.addFacilitiesEventListener(0);
    }

    //Check if AFG Project
    //Get Submission Category
    var submissionCategoryCode = formContext.getAttribute("caps_submissioncategorycode").getValue();

    if (submissionCategoryCode === "AFG") {
        //add on-change function to existing facility? caps_existingfacility
        formContext.getAttribute("caps_existingfacility").addOnChange(CAPS.Project.ToggleAFGFacility);
    }    
}
/**
 * Function to toggle showing and hiding of facility and facility site on AFG.
 * @param {any} executionContext the execution context
 */
CAPS.Project.ToggleAFGFacility = function (executionContext) {
    var formContext = executionContext.getFormContext();

    var submissionCategoryTabNames = formContext.getAttribute("caps_submissioncategorytabname").getValue();
    var arrTabNames = submissionCategoryTabNames.split(", ");

    var showExistingFacility = (formContext.getAttribute("caps_existingfacility").getValue() === true) ? true : false;

    //loop through tabs
    formContext.ui.tabs.forEach(function (tab, i) {
        //loop through sections
        if (arrTabNames.includes(tab.getName())) {
            //loop through sections
            tab.sections.forEach(function (section, j) {
                section.controls.forEach(function (control, k) {

                    if (control.getAttribute().getName() === "caps_facility") {
                        if (showExistingFacility) {
                            control.getAttribute().setRequiredLevel("required");
                            control.setVisible(true);
                        }
                        else {
                            control.getAttribute().setRequiredLevel("none");
                            control.setVisible(false);
                            control.getAttribute().setValue(null);
                        }
                    }

                    if (control.getAttribute().getName() === "caps_facilityview") {
                        if (showExistingFacility) {
                            control.getAttribute().setRequiredLevel("none");
                            control.setVisible(false);
                            control.getAttribute().setValue(null);
                        }
                        else {
                            control.getAttribute().setRequiredLevel("required");
                            control.setVisible(true);
                        }
                    }

                });
            });
        }
    });


}

/**
 * Function to toggle showing and hiding of facility lookup and subgrid depending on if the multiple facility? field is set to Yes or No.
 * @param {any} executionContext the form execution context
 */
CAPS.Project.SetMultipleFacility = function (executionContext) {
    var formContext = executionContext.getFormContext();

    var submissionCategoryTabNames = formContext.getAttribute("caps_submissioncategorytabname").getValue();
    var arrTabNames = submissionCategoryTabNames.split(", ");

    var showMultipleFacilities = (formContext.getAttribute("caps_multiplefacilities").getValue() === true) ? true : false;

    //loop through tabs
    formContext.ui.tabs.forEach(function (tab, i) {
        //loop through sections
        if (arrTabNames.includes(tab.getName())) {
            //loop through sections
            tab.sections.forEach(function (section, j) {
                section.controls.forEach(function (control, k) {


                    //add to array
                    //console.log(control.getControlType());
                    if (control.getControlType() === "subgrid") {
                        if (control.getEntityName() === "caps_facility") {
                            CAPS.Project.FACILITY_GRID_CONTROL = control.getName();
                            control.setVisible(showMultipleFacilities);
                        }
                    }

                    if (control.getAttribute().getName() === "caps_facility") {
                        control.setVisible(!showMultipleFacilities);

                        if (showMultipleFacilities) {
                            control.getAttribute().setRequiredLevel("none");
                            control.getAttribute().setValue(null);
                        }
                        else {
                            control.getAttribute().setRequiredLevel("required");
                        }
                    }

                });
            });
        }
    });
   
}

/**
 * Set's the Project Type if the lookup list only contains one value.
 * @param {any} executionContext execution context
 */
CAPS.Project.SetProjectTypeValue = function (executionContext) {

    var formContext = executionContext.getFormContext();

    //Get Submission Category
    var submissionCategory = formContext.getAttribute("caps_submissioncategory").getValue();

    if (submissionCategory !== null && submissionCategory[0] !== null) {
        var submissionCategoryID = submissionCategory[0].id;

        //Filtering fetch XML for Project Type
        var filterFetchXml = "<link-entity name=\"caps_submissioncategory_caps_projecttype\" from=\"caps_projecttypeid\" to=\"caps_projecttypeid\" visible=\"false\" intersect=\"true\">" +
            "<link-entity name=\"caps_submissioncategory\" from=\"caps_submissioncategoryid\" to=\"caps_submissioncategoryid\" alias=\"ab\">" +
            "<filter type=\"and\">" +
            "<condition attribute=\"caps_submissioncategoryid\" operator=\"eq\" value=\"" + submissionCategoryID + "\" />" +
            "</filter>" +
            "</link-entity>" +
            "</link-entity>";

        //Call to set default value if only one value exists
        CAPS.Common.DefaultLookupIfSingle(formContext, "caps_projecttype", "caps_projecttype", "caps_projecttypeid", "caps_type", filterFetchXml);
    }
}

/**
 * Prevents autosave if the global prevent autosave flag is set
 * @param {any} executionContext execution context
 */
CAPS.Project.onSave = function (executionContext) {
    var eventArgs = executionContext.getEventArgs();

    if (CAPS.Project.PREVENT_AUTO_SAVE) {

        //auto-save = 70
        if (eventArgs.getSaveMode() === 70) {
            eventArgs.preventDefault();
        }
    }
}

/**
 * Sets the projects School District to the user's business unit's school district if it's set
 * @param {any} formContext the form's form context
 */
CAPS.Project.DefaultSchoolDistrict = function (formContext) {
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
                        formContext.getAttribute("caps_schooldistrict").setValue([{ id: sdID, name: sdName, entityType: sdType }]);
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

/**
 * This function shows either the General tab for new Projects or the relevant tab from the related Submission Category for existing projects.
 * It also calls a function to turn off any field validation for any tab not shown.
 * @param {any} formContext form context
 */
CAPS.Project.ShowHideRelevantTabs = function (formContext) {
    debugger;
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
        var submissionCategoryTabNames = formContext.getAttribute("caps_submissioncategorytabname").getValue();
        var arrTabNames = submissionCategoryTabNames.split(", ");

        //Remove all mandatory fields and show relevant tab(s)
        var tabsToDisregard = [TIMELINE_TAB];
        arrTabNames.forEach(function (tabName) {
            tabsToDisregard.push(tabName);
        });
        

        CAPS.Project.RemoveRequirement(formContext, tabsToDisregard);

        arrTabNames.forEach(function (tabName) {
            formContext.ui.tabs.get(tabName).setVisible(true);
        });

        

        //if capital expense needs allocating, show the tab
        //if (formContext.getAttribute("caps_submissioncategoryrequirecostallocation").getValue() === true) {
        //    formContext.ui.tabs.get(CAPITAL_EXPENDITURE_TAB).setVisible(true);
        //}
    }
}

/**
 * This function turns off all field requirements for any field except those in the tabsToDisregard array
 * @param {any} formContext form context
 * @param {any} tabsToDisregard - array of tab names to disregard
 */
CAPS.Project.RemoveRequirement = function (formContext, tabsToDisregard) {
    debugger;
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

/**
 * This function compares the total project cost to the sum of the estimated yearly expenditures and shows an error if they don't match.
 * This function is only called if the related Submission Category field Require Cost Allocation is set to Yes.
 * @param {any} executionContext Execution Context
 */
CAPS.Project.ValidateExpenditureDistribution = function (executionContext) {
    //Only validate if Submission Category requires 10 year plan
    //If numbers don't match, show formContext.getControl(arg).setNotification();
    var formContext = executionContext.getFormContext();

    var totalProjectCost = formContext.getAttribute("caps_totalprojectcost").getValue();
    var sumOfEstimatedExpenditures = formContext.getAttribute("caps_totalallocated").getValue();

    if (totalProjectCost !== null && totalProjectCost !== sumOfEstimatedExpenditures) {
        formContext.ui.setFormNotification('Total Project Cost Not Fully Allocated', 'WARNING', COST_MISSMATCH_NOTIFICATION);
        //formContext.getControl("caps_totalprojectcost").setNotification('Total Project Cost Not Fully Allocated', COST_MISSMATCH_NOTIFICATION);
    }
    else {
        formContext.ui.clearFormNotification(COST_MISSMATCH_NOTIFICATION);
        //formContext.getControl("caps_totalprojectcost").clearNotification(COST_MISSMATCH_NOTIFICATION);
    }
}

/**
 * This function waits for the Facilities subgrid to load and adds an event listener to the grid for validating that at least one facility was added.
 * @param {any} loopCount count of loops
 */
CAPS.Project.addFacilitiesEventListener = function (loopCount) {

    var gridContext = CAPS.Project.GLOBAL_FORM_CONTEXT.getControl(CAPS.Project.FACILITY_GRID_CONTROL);

    if (loopCount < 5) {
        if (gridContext === null) {
            setTimeout(function () { FACILITIES_EVENT_HANDLER_LOOP_COUNTER++; CAPS.Project.addFacilitiesEventListener(loopCount++); }, 500);
        }

        gridContext.addOnLoad(CAPS.Project.ValidateAtLeastOneFacility);
    }
}

/**
 * This function validates that at least one facility has been added to the project
 * @param {any} executionContext Execution Context
 */
CAPS.Project.ValidateAtLeastOneFacility = function (executionContext) {
    debugger;
    var formContext = executionContext.getFormContext();
    if (formContext.getAttribute("caps_multiplefacilities").getValue() === true) {
        //var gridContext = executionContext.getFormContext();

        var filteredRecordCount = formContext.getControl(CAPS.Project.FACILITY_GRID_CONTROL).getGrid().getTotalRecordCount();

        if (filteredRecordCount < 1) {
            formContext.ui.setFormNotification('You must add at least one facility to this project.', 'INFO', NO_FACILITY_NOTIFICATION);
        }
        else {
            formContext.ui.clearFormNotification(NO_FACILITY_NOTIFICATION);
        }
    }

}


