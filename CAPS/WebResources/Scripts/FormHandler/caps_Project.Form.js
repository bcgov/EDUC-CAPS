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
const MINISTRY_REVIEW_TAB = "tab_ministry_review";
const COST_MISSMATCH_NOTIFICATION = "Cost_Missmatch_Notification";
const NO_FACILITY_NOTIFICATION = "No_Facility_Notification";
const PRFS_INCOMPLETE_NOTIFICATION = "PRFS_Incomplete_Notification";
const PRFS_NOSCHOOLS_NOTIFICATION = "PRFS_No_Surrounding_Schools_Notification";
const PRFS_NOALTERNATIVES_NOTIFICATION = "PRFS_No_Alternatives_Notification";
const CURRENT_FCI_NOTIFICATION = "Current_FCI_Notification";
const FUTURE_FCI_NOTIFICATION = "Future_FCI_Notification";
const HEALTH_AUTHORITY_NOTIFICATION = "Health_Authority_Notification";

/**
 * Main function for Project.  This function calls all other form functions and registers onChange and onLoad events.
 * @param {any} executionContext the form execution context
 */
CAPS.Project.onLoad = function (executionContext) {

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

    //Get Submission Category
    var submissionCategoryCode = formContext.getAttribute("caps_submissioncategorycode").getValue();

    //Show/Hide Tabs
    CAPS.Project.ShowHideRelevantTabs(formContext);




    //On Create
    if (formState === FORM_STATE.CREATE) {

        CAPS.Project.RECORD_JUST_CREATED = true;

        //remove Lease from list
        formContext.getControl("caps_submissioncategory").addPreSearch(CAPS.Project.HideLeaseSubmissionCategory);

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

        //sgd_EstimatedExpenditures
        formContext.getControl("sgd_EstimatedExpenditures").addOnLoad(CAPS.Project.UpdateTotalAllocated);

        //add on-change events for prfs if not BEP
        if (submissionCategoryCode != "BEP") {
            CAPS.Project.SetupPRFSValidation(executionContext);

        }
    }

    if (submissionCategoryCode === "CC_CONVERSION_MINOR" || submissionCategoryCode === "CC_UPGRADE_MINOR") {

        CAPS.Project.SetupPRFSValidation(executionContext);
    }

    if (submissionCategoryCode === "LEASE") {
        CAPS.Project.SetupPRFSValidation(executionContext);
    }

    //Only call for SEP and CNCP!
    //TODO: Have added a flag to Project Submission, if we are keeping then add calculated field to Project and check here
    if (formContext.getAttribute("caps_submissioncategoryallowmultiplefacilities").getValue() === true) {
        CAPS.Project.SetMultipleFacility(executionContext);

        //Add onChange event to caps_multiplefacility
        formContext.getAttribute("caps_multiplefacilities").addOnChange(CAPS.Project.SetMultipleFacility);

        CAPS.Project.addFacilitiesEventListener(0);
    }

    //Check if AFG or Demolition Project to setup existing facility toggle
    if (submissionCategoryCode === "AFG" || submissionCategoryCode === "DEMOLITION") {
        CAPS.Project.ToggleFacility(executionContext);
        //add on-change function to existing facility? caps_existingfacility
        formContext.getAttribute("caps_existingfacility").addOnChange(CAPS.Project.ToggleFacility);
    }

    //Hide Ministry Review Status of Planned if not allowed
    if (formContext.getAttribute("caps_submissioncategoryallowplannedstatus").getValue() !== true) {
        //remove planned (2008700000)
        formContext.getControl("caps_ministryassessmentstatus").removeOption(200870000);
    }

    //Hide Funding Awarded if not Minor
    if (submissionCategoryCode == "ADDITION" || submissionCategoryCode == "DEMOLITION" || submissionCategoryCode == "NEW_SCHOOL" ||
        submissionCategoryCode == "REPLACEMENT_RENOVATION" || submissionCategoryCode == "SEISMIC" ||
        submissionCategoryCode == "SITE_ACQUISITION" || submissionCategoryCode == "BEP" ||
        submissionCategoryCode === "CC_UPGRADE" || submissionCategoryCode === "CC_CONVERSION" ||
        submissionCategoryCode === "CC_MAJOR_NEW_SPACES_INTEGRATED" || submissionCategoryCode === "Major_CC_New_Spaces") {
        formContext.getControl("caps_fundingawarded").setVisible(false);
    }

    //Adding Schedule B Toggles & general setup of schedule B    
    CAPS.Project.SetupScheduleB(executionContext);

    //if PEP Replacement, show Playground age
    if (submissionCategoryCode === "PEP") {
        //add on-change to project type
        formContext.getAttribute("caps_projecttype").addOnChange(CAPS.Project.TogglePEPReplacement);
        CAPS.Project.TogglePEPReplacement(executionContext);
    }

    //if BUS Replacement, show bus to be replaced
    if (submissionCategoryCode === "BUS") {
        //add on-change to project type
        formContext.getAttribute("caps_projecttype").addOnChange(CAPS.Project.ToggleBUSReplacement);
        CAPS.Project.ToggleBUSReplacement(executionContext);

        formContext.getAttribute("caps_issuetype").addOnChange(CAPS.Project.ToggleBusIssueType);
        CAPS.Project.ToggleBusIssueType(executionContext);
    }

    //if site aquisition or lease hide procurement analysis
    if (submissionCategoryCode === "SITE_ACQUISITION" || submissionCategoryCode === "LEASE") {
        formContext.ui.tabs.get("tab_PRFS").sections.get("PRFS_section_PROCUREMENT_ANALYSIS").setVisible(false);
    }

    //Filter project group for all project types
    CAPS.Project.preFilterProjectGroupLookup(executionContext);


    if (submissionCategoryCode === "CC_CONVERSION_MINOR" || submissionCategoryCode === "CC_CONVERSION" ||
        submissionCategoryCode === "CC_MAJOR_NEW_SPACES_INTEGRATED" || submissionCategoryCode === "Major_CC_New_Spaces" ||
        submissionCategoryCode === "CC_UPGRADE" || submissionCategoryCode === "CC_UPGRADE_MINOR") {

        formContext.getAttribute("caps_existingchildcarefacility").addOnChange(CAPS.Project.ToggleExistingChildCareFacility);
        formContext.getAttribute("caps_existingfacility").addOnChange(CAPS.Project.ToggleExistingChildCareFacility);
        CAPS.Project.ToggleExistingChildCareFacility(executionContext);
        formContext.getAttribute("caps_programtailoredtomeetneedsofpopulationgroup").addOnChange(CAPS.Project.ShowHidePopulationGroup);
        CAPS.Project.ShowHidePopulationGroup(executionContext);
        formContext.getAttribute("caps_areyourelocatingexistingchildcarespaces").addOnChange(CAPS.Project.ShowHideSubgrid);
        formContext.getAttribute("caps_areyourelocatingexistingchildcarespaces").addOnChange(CAPS.Project.ShowConfirmationRelocatingCC);
        CAPS.Project.ShowHideSubgrid(executionContext);
        
    }

    formContext.getAttribute("caps_areyourelocatingexistingchildcarespaces").addOnChange(CAPS.Project.HideTotalChildCareSpace);
    CAPS.Project.HideTotalChildCareSpace(executionContext);


    if (submissionCategoryCode === "Major_CC_New_Spaces" || submissionCategoryCode === "CC_MAJOR_NEW_SPACES_INTEGRATED" ||
        submissionCategoryCode === "CC_CONVERSION_MINOR" || submissionCategoryCode === "CC_CONVERSION" ||
        submissionCategoryCode === "CC_UPGRADE" || submissionCategoryCode === "CC_UPGRADE_MINOR") {

        formContext.getAttribute("caps_childcare").addOnChange(CAPS.Project.PopulateSchoolFacility);
        CAPS.Project.PopulateSchoolFacility(executionContext);

        formContext.getAttribute("caps_childcare").addOnChange(CAPS.Project.WipeSchoolFacilityWhenCCWiped);
    }

    if (submissionCategoryCode === "Major_CC_New_Spaces" || submissionCategoryCode === "CC_MAJOR_NEW_SPACES_INTEGRATED" ||
        submissionCategoryCode === "CC_CONVERSION" || submissionCategoryCode === "CC_UPGRADE") {

        formContext.getAttribute("caps_projectrequireuniquesitedevelopment").addOnChange(CAPS.Project.ShowHideFieldsCCPBTab);
        CAPS.Project.ShowHideFieldsCCPBTab(executionContext);
    }

    if (submissionCategoryCode === "Major_CC_New_Spaces" || submissionCategoryCode === "CC_MAJOR_NEW_SPACES_INTEGRATED" ||
        submissionCategoryCode === "CC_CONVERSION" || submissionCategoryCode === "CC_CONVERSION_MINOR") {
        formContext.getAttribute("caps_existingchildcarefacility").addOnChange(CAPS.Project.WipeSchoolFacility);
        formContext.getAttribute("caps_childcare").addOnChange(CAPS.Project.WipeSchoolFacility);
    }

    //Hide Funding Requested for all CC Majors
    CAPS.Project.HideFundingRequested(executionContext);


    //Hide Health Authority on CC Majors when Existing Child Care Facility is Yes
    formContext.getAttribute("caps_existingchildcarefacility").addOnChange(CAPS.Project.ShowHideHealthAuthority);
    CAPS.Project.ShowHideHealthAuthority(executionContext);

    formContext.getAttribute("caps_childcare").addOnChange(CAPS.Project.ShowNotificationHealthAuthority);
    CAPS.Project.ShowNotificationHealthAuthority(executionContext);

    //setup LRFP field validation and display
    CAPS.Project.ToggleLRFP(executionContext);
    formContext.getAttribute("caps_longrangefacilityplan").addOnChange(CAPS.Project.ToggleLRFP);

    //Validate FCI
    formContext.getAttribute("caps_currentfci").addOnChange(CAPS.Project.ValidateCurrentFCI);
    CAPS.Project.ValidateCurrentFCI(executionContext);

    formContext.getAttribute("caps_futurefacilityconditionindex").addOnChange(CAPS.Project.ValidateFutureFCI);
    CAPS.Project.ValidateFutureFCI(executionContext);

    //Design Capacity Checks
    CAPS.Project.ValidateKindergartenDesignCapacity(executionContext);
    formContext.getAttribute("caps_changeindesigncapacitykindergarten").addOnChange(CAPS.Project.ValidateKindergartenDesignCapacity);

    CAPS.Project.ValidateElementaryDesignCapacity(executionContext);
    formContext.getAttribute("caps_changeindesigncapacityelementary").addOnChange(CAPS.Project.ValidateElementaryDesignCapacity);

    CAPS.Project.ValidateSecondaryDesignCapacity(executionContext);
    formContext.getAttribute("caps_changeindesigncapacitysecondary").addOnChange(CAPS.Project.ValidateSecondaryDesignCapacity);

    //Remove Previously Planned as an option from Ministry Review Status field
    formContext.getControl("caps_ministryassessmentstatus").removeOption(200870001);

}


CAPS.Project.SetupPRFSValidation = function (executionContext) {

    var formContext = executionContext.getFormContext();
    var submissionCategoryCode = formContext.getAttribute("caps_submissioncategorycode").getValue();

    //majors
    formContext.getAttribute("caps_projectrationale").addOnChange(CAPS.Project.ValidatePRFS);
    formContext.getAttribute("caps_scopeofwork").addOnChange(CAPS.Project.ValidatePRFS);
    formContext.getAttribute("caps_tempaccommodationandbusingplan").addOnChange(CAPS.Project.ValidatePRFS);
    formContext.getAttribute("caps_municipalrequirements").addOnChange(CAPS.Project.ValidatePRFS);
    formContext.getAttribute("caps_accessibility").addOnChange(CAPS.Project.ValidatePRFS);
    formContext.getAttribute("caps_justification").addOnChange(CAPS.Project.ValidatePRFS);

    //minors
    formContext.getAttribute("caps_indoorfloorplans").addOnChange(CAPS.Project.ValidatePRFS);
    formContext.getAttribute("caps_outdoorplans").addOnChange(CAPS.Project.ValidatePRFS);
    formContext.getAttribute("caps_projectbudget").addOnChange(CAPS.Project.ValidatePRFS);

    //demolition
    formContext.getAttribute("caps_demolitioncompletedinonefiscalyear").addOnChange(CAPS.Project.ValidatePRFS);
    formContext.getAttribute("caps_hazmatenvassesscomplete").addOnChange(CAPS.Project.ValidatePRFS);
    formContext.getAttribute("caps_datebuildingportionbecameunoccupied").addOnChange(CAPS.Project.ValidatePRFS);
    formContext.getAttribute("caps_hasschoolbeenpermanentlyclosed").addOnChange(CAPS.Project.ValidatePRFS);
    formContext.getAttribute("caps_estimatedmarketvalueofpropertywbuilding").addOnChange(CAPS.Project.ValidatePRFS);
    formContext.getAttribute("caps_estimatedmarketvalueoflandwobuilding").addOnChange(CAPS.Project.ValidatePRFS);
    formContext.getAttribute("caps_summaryofhazardousmaterialsenvassessment").addOnChange(CAPS.Project.ValidatePRFS);
    formContext.getAttribute("caps_challengestheexistingbuildingpresents").addOnChange(CAPS.Project.ValidatePRFS);
    formContext.getAttribute("caps_benefitsofafullorpartialdemolition").addOnChange(CAPS.Project.ValidatePRFS);
    formContext.getAttribute("caps_potentialplanforvacantsite").addOnChange(CAPS.Project.ValidatePRFS);

    formContext.getAttribute("caps_totalallocated").addOnChange(CAPS.Project.ValidatePRFS);
    formContext.getAttribute("caps_submission").addOnChange(CAPS.Project.ValidatePRFS);
    CAPS.Project.ValidatePRFS(executionContext);

    if (submissionCategoryCode != "DEMOLITION") {
        formContext.getControl("Subgrid_PRFS_Alt_Options").addOnLoad(CAPS.Project.ValidatePRFSAlternativeOptions);
        formContext.getControl("sgd_surroundingschools").addOnLoad(CAPS.Project.ValidatePRFSSurroundingSchools);
        formContext.getAttribute("caps_submission").addOnChange(CAPS.Project.ValidatePRFSAlternativeOptions);
        formContext.getAttribute("caps_submission").addOnChange(CAPS.Project.ValidatePRFSSurroundingSchools);


        CAPS.Project.ValidatePRFSAlternativeOptions(executionContext);
        CAPS.Project.ValidatePRFSSurroundingSchools(executionContext);
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

/*
This function hides the submission category lease for all school districts except SD93
*/
CAPS.Project.HideLeaseSubmissionCategory = function () {
    var customerCategoryFilter = "<filter type='and'><condition attribute='caps_categorycode' operator='ne' value='LEASE' /></filter>";
    CAPS.Project.GLOBAL_FORM_CONTEXT.getControl("caps_submissioncategory").addCustomFilter(customerCategoryFilter, "caps_submissioncategory");
}

/**
 * This function is called on change of the Estimated Expenditure PCF sub grid.  It calculates and updates the total allocated and total variance for the project cost.
 * @param {any} executionContext forms execution context.
 */
CAPS.Project.UpdateTotalAllocated = function (executionContext) {

    var formContext = executionContext.getFormContext();
    var id = formContext.data.entity.getId().replace("{", "").replace("}", "");
    Xrm.WebApi.retrieveMultipleRecords("caps_estimatedyearlycapitalexpenditure", "?$select=caps_yearlyexpenditure&$filter=caps_Project/caps_projectid eq " + id + " and statecode eq 0").then(
        function success(result) {
            var totalAllocated = 0;
            for (var i = 0; i < result.entities.length; i++) {
                totalAllocated += result.entities[i].caps_yearlyexpenditure;
            }

            // perform operations on record retrieval
            if (formContext.getAttribute('caps_totalallocated').getValue() != totalAllocated && executionContext.getFormContext() != FORM_STATE.READ_ONLY) {
                formContext.getAttribute('caps_totalallocated').setValue(totalAllocated);
            }
            //calculate variance
            var totalCost = formContext.getAttribute('caps_totalprojectcost').getValue();
            var variance = null;
            if (totalCost !== null && totalAllocated !== null) {
                variance = totalCost - totalAllocated;
                if (formContext.getAttribute('caps_totalprojectcostvariance').getValue() != variance && executionContext.getFormContext() != FORM_STATE.READ_ONLY) {
                    formContext.getAttribute('caps_totalprojectcostvariance').setValue(variance);
                }
            }

            CAPS.Project.ValidateExpenditureDistribution(executionContext);
            CAPS.Project.ValidatePRFS(executionContext);
        },
        function (error) {
            console.log(error.message);
            // handle error conditions
        }
    );

}

/**
 * This function is used to toggle showing and hiding of facility and facility site on onchange of Existing Facility? field.
 * @param {any} executionContext the execution context
 */
CAPS.Project.ToggleFacility = function (executionContext) {
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
                            if (control.getAttribute().getValue() != null && executionContext.getFormContext() != FORM_STATE.READ_ONLY) {
                                control.getAttribute().setValue(null);
                            }
                        }
                    }

                    if (control.getAttribute().getName() === "caps_otherfacility") {
                        if (showExistingFacility) {
                            control.getAttribute().setRequiredLevel("none");
                            control.setVisible(false);
                            if (control.getAttribute().getValue() != null && executionContext.getFormContext() != FORM_STATE.READ_ONLY) {
                                control.getAttribute().setValue(null);
                            }
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

    if (!showMultipleFacilities) {
        formContext.ui.clearFormNotification(NO_FACILITY_NOTIFICATION);
    }

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

                            if (control.getAttribute().getValue() != null && executionContext.getFormContext() != FORM_STATE.READ_ONLY) {
                                control.getAttribute().setValue(null);
                            }
                        }
                        else {
                            control.getAttribute().setRequiredLevel("required");
                        }
                    }

                    if (control.getAttribute().getName() === "caps_phasedprojectgroup") {
                        if (showMultipleFacilities) {
                            control.setVisible(false);

                            if (control.getAttribute().getValue() != null && executionContext.getFormContext() != FORM_STATE.READ_ONLY) {
                                control.getAttribute().setValue(null);
                            }
                        }
                        else {
                            control.setVisible(true);
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
        var filterFetchXml = "<filter type=\"and\">" +
            "<condition attribute=\"caps_submissioncategory\" operator=\"eq\" value=\"" + submissionCategoryID + "\" />" +
            "</filter>";

        //Call to set default value if only one value exists
        CAPS.Common.DefaultLookupIfSingle(formContext, "caps_projecttype", "caps_projecttype", "caps_projecttypeid", "caps_type", filterFetchXml);
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



                        //if SD93
                        if (sdName.includes("SD93")) {
                            formContext.getControl("caps_submissioncategory").removePreSearch(CAPS.Project.HideLeaseSubmissionCategory);
                        }
                        else {
                            formContext.getAttribute("caps_hostschooldistrict").setValue([{ id: sdID, name: sdName, entityType: sdType }]);
                        }
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

        //Get fields that should or should not be mandatory
        var mandatoryFields = formContext.getAttribute("caps_submissioncategorymandatoryfields").getValue();
        var optionalFields = formContext.getAttribute("caps_submissioncategoryoptionalfields").getValue();

        CAPS.Project.RemoveRequirement(formContext, tabsToDisregard, mandatoryFields, optionalFields);

        arrTabNames.forEach(function (tabName) {
            formContext.ui.tabs.get(tabName).setVisible(true);
        });

        //Show Ministry Review Tab if this is the CAPS app
        var globalContext = Xrm.Utility.getGlobalContext();
        globalContext.getCurrentAppName().then(
            function success(result) {
                if (result === "CAPS") {
                    formContext.ui.tabs.get(MINISTRY_REVIEW_TAB).setVisible(true);
                }
            }
        );

        //set focus to first tab in list
        formContext.ui.tabs.get(arrTabNames[0]).setDisplayState("expanded");

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
 * @param {string} mandatoryFields - string of mandatory fields
 * @param {string} optionalFields - string of optional fields
 * */
CAPS.Project.RemoveRequirement = function (formContext, tabsToDisregard, mandatoryFields, optionalFields) {


    var mandatoryFieldArray = (mandatoryFields !== null && mandatoryFields !== undefined) ? mandatoryFields.split(",") : [];
    var optionalFieldArray = (optionalFields !== null && optionalFields !== undefined) ? optionalFields.split(",") : [];

    //Get array of all fields on tabs to disregard
    var fieldsToShow = [];

    formContext.ui.tabs.forEach(function (tab, i) {
        //loop through sections
        if (tabsToDisregard.includes(tab.getName())) {
            tab.sections.forEach(function (section, j) {
                if (section.name !== "general_sec_hidden") {
                    section.controls.forEach(function (control, k) {
                        //add to array
                        fieldsToShow.push(control.getAttribute().getName());
                    });
                }
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

    //loop through one last time setting mandatory and not mandatory
    formContext.ui.tabs.forEach(function (tab, i) {
        //loop through sections
        //if (!tabsToDisregard.includes(tab.getName())) {
        tab.sections.forEach(function (section, j) {
            section.controls.forEach(function (control, k) {
                //if the field is in the mandatory list or optional list then setup appropriately
                if (mandatoryFieldArray.includes(control.getAttribute().getName())) {
                    control.getAttribute().setRequiredLevel("required");
                }
                if (optionalFieldArray.includes(control.getAttribute().getName())) {
                    control.getAttribute().setRequiredLevel("none");
                }
            });
        });
        //}
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
        formContext.ui.setFormNotification('Total Project Cost not fully allocated  on the Cashflow tab.', 'INFO', COST_MISSMATCH_NOTIFICATION);
        //formContext.getControl("caps_totalprojectcost").setNotification('Total Project Cost Not Fully Allocated', COST_MISSMATCH_NOTIFICATION);
    }
    else {
        formContext.ui.clearFormNotification(COST_MISSMATCH_NOTIFICATION);
        //formContext.getControl("caps_totalprojectcost").clearNotification(COST_MISSMATCH_NOTIFICATION);
    }
}

/**
 * This function checks if the 4 main PRFS fields are filled in for major projects if there is cash flow in the first 3 years.
 * If the fields aren't filled out, a warning is shown to the user.
 * @param {any} executionContext execution context
 */
CAPS.Project.ValidatePRFS = function (executionContext) {

    var formContext = executionContext.getFormContext();
    var prfsFieldsComplete = true;

    if (formContext.getAttribute("caps_submissioncategorycode").getValue() !== "BEP" && formContext.getAttribute("caps_submissioncategorycode").getValue() !== "DEMOLITION" &&
        formContext.getAttribute("caps_submissioncategorycode").getValue() !== "CC_CONVERSION" && formContext.getAttribute("caps_submissioncategorycode").getValue() !== "CC_CONVERSION_MINOR" &&
        formContext.getAttribute("caps_submissioncategorycode").getValue() !== "Major_CC_New_Spaces" && formContext.getAttribute("caps_submissioncategorycode").getValue() !== "CC_MAJOR_NEW_SPACES_INTEGRATED" &&
        formContext.getAttribute("caps_submissioncategorycode").getValue() !== "CC_UPGRADE" && formContext.getAttribute("caps_submissioncategorycode").getValue() !== "CC_UPGRADE_MINOR") {
        //Check the 4 fields
        var projectRationale = formContext.getAttribute("caps_projectrationale").getValue();
        var scopeOfWork = formContext.getAttribute("caps_scopeofwork").getValue();
        var tempAccomodation = formContext.getAttribute("caps_tempaccommodationandbusingplan").getValue();
        var municipalRequirements = formContext.getAttribute("caps_municipalrequirements").getValue();

        if (projectRationale !== null && scopeOfWork !== null && tempAccomodation !== null && municipalRequirements !== null) {
            prfsFieldsComplete = true;
            formContext.ui.clearFormNotification(PRFS_INCOMPLETE_NOTIFICATION);
        }
        else {
            prfsFieldsComplete = false;
        }
    }
    else if (formContext.getAttribute("caps_submissioncategorycode").getValue() == "DEMOLITION") {

        var demo1 = formContext.getAttribute("caps_demolitioncompletedinonefiscalyear").getValue();
        var demo2 = formContext.getAttribute("caps_hazmatenvassesscomplete").getValue();
        var demo3 = formContext.getAttribute("caps_datebuildingportionbecameunoccupied").getValue();
        var demo4 = formContext.getAttribute("caps_hasschoolbeenpermanentlyclosed").getValue();
        var demo5 = formContext.getAttribute("caps_estimatedmarketvalueofpropertywbuilding").getValue();
        var demo6 = formContext.getAttribute("caps_estimatedmarketvalueoflandwobuilding").getValue();
        var demo7 = formContext.getAttribute("caps_summaryofhazardousmaterialsenvassessment").getValue();
        var demo8 = formContext.getAttribute("caps_challengestheexistingbuildingpresents").getValue();
        var demo9 = formContext.getAttribute("caps_benefitsofafullorpartialdemolition").getValue();
        var demo10 = formContext.getAttribute("caps_potentialplanforvacantsite").getValue();

        if (demo1 !== null && demo2 !== null && demo3 !== null && demo4 !== null && demo5 !== null && demo6 !== null && demo7 !== null && demo8 !== null && demo9 !== null && demo10 !== null) {
            prfsFieldsComplete = true;
            formContext.ui.clearFormNotification(PRFS_INCOMPLETE_NOTIFICATION);
        }
        else {
            prfsFieldsComplete = false;
        }

    }

    else if (formContext.getAttribute("caps_submissioncategorycode").getValue() === "CC_CONVERSION" || formContext.getAttribute("caps_submissioncategorycode").getValue() === "Major_CC_New_Spaces" ||
        formContext.getAttribute("caps_submissioncategorycode").getValue() === "CC_MAJOR_NEW_SPACES_INTEGRATED" || formContext.getAttribute("caps_submissioncategorycode").getValue() === "CC_UPGRADE") {
        var projectRationale = formContext.getAttribute("caps_projectrationale").getValue();
        var scopeOfWork = formContext.getAttribute("caps_scopeofwork").getValue();
        var accessibility = formContext.getAttribute("caps_accessibility").getValue();
        var justification = formContext.getAttribute("caps_justification").getValue();

        if (projectRationale !== null && scopeOfWork !== null && accessibility !== null && justification !== null) {
            prfsFieldsComplete = true;
            formContext.ui.clearFormNotification(PRFS_INCOMPLETE_NOTIFICATION);
        }
        else {
            prfsFieldsComplete = false;
        }
    }

    else if (formContext.getAttribute("caps_submissioncategorycode").getValue() === "CC_CONVERSION_MINOR" || formContext.getAttribute("caps_submissioncategorycode").getValue() === "CC_UPGRADE_MINOR") {

        var projectRationale = formContext.getAttribute("caps_projectrationale").getValue();
        var scopeOfWork = formContext.getAttribute("caps_scopeofwork").getValue();
        var accessibility = formContext.getAttribute("caps_accessibility").getValue();
        var indoorFloorPlans = formContext.getAttribute("caps_indoorfloorplans").getValue();
        var outdoorPlans = formContext.getAttribute("caps_outdoorplans").getValue();
        var projectBudget = formContext.getAttribute("caps_projectbudget").getValue();

        if (projectRationale !== null && scopeOfWork !== null &&
            accessibility !== null && indoorFloorPlans !== null &&
            outdoorPlans !== null && projectBudget !== null) {
            prfsFieldsComplete = true;
            formContext.ui.clearFormNotification(PRFS_INCOMPLETE_NOTIFICATION);
        }
        else {
            prfsFieldsComplete = false;
        }
    }


    if (!prfsFieldsComplete) {

        //Call action to validate 
        var recordId = formContext.data.entity.getId().replace("{", "").replace("}", "");
        //call action
        var req = {};
        var target = { entityType: "caps_project", id: recordId };
        req.entity = target;

        req.getMetadata = function () {
            return {
                boundParameter: "entity",
                operationType: 0,
                operationName: "caps_ValidatePRFS",
                parameterTypes: {
                    "entity": {
                        "typeName": "mscrm.caps_project",
                        "structuralProperty": 5
                    }
                }
            };
        };

        Xrm.WebApi.online.execute(req).then(
            function (result) {

                if (result.ok) {
                    return result.json().then(
                        function (response) {

                            if (formContext.getAttribute("caps_submissioncategorycode").getValue() === "CC_CONVERSION" || formContext.getAttribute("caps_submissioncategorycode").getValue() === "CC_CONVERSION_MINOR" ||
                                formContext.getAttribute("caps_submissioncategorycode").getValue() === "Major_CC_New_Spaces" || formContext.getAttribute("caps_submissioncategorycode").getValue() === "CC_MAJOR_NEW_SPACES_INTEGRATED" ||
                                formContext.getAttribute("caps_submissioncategorycode").getValue() === "CC_UPGRADE" || formContext.getAttribute("caps_submissioncategorycode").getValue() === "CC_UPGRADE_MINOR") {

                                formContext.ui.setFormNotification('CC-PRFS is not complete.', 'INFO', PRFS_INCOMPLETE_NOTIFICATION);
                            }
                            else if (response.hasCashflow) {

                                formContext.ui.setFormNotification('PRFS is not complete.', 'INFO', PRFS_INCOMPLETE_NOTIFICATION);

                            }
                            else {
                                formContext.ui.clearFormNotification(PRFS_INCOMPLETE_NOTIFICATION);
                            }
                        });
                }
            },
            function (e) {
                formContext.ui.clearFormNotification(PRFS_INCOMPLETE_NOTIFICATION);
            }
        );
    }

}



/**
This function checks if there is at least one PRFS alternative option specified if there is cashflow in the first 3 years.
*/
CAPS.Project.ValidatePRFSAlternativeOptions = function (executionContext) {
    var formContext = executionContext.getFormContext();
    if (formContext.getAttribute("caps_submissioncategorycode").getValue() !== "CC_CONVERSION" && formContext.getAttribute("caps_submissioncategorycode").getValue() !== "CC_CONVERSION_MINOR" &&
        formContext.getAttribute("caps_submissioncategorycode").getValue() !== "Major_CC_New_Spaces" && formContext.getAttribute("caps_submissioncategorycode").getValue() !== "CC_MAJOR_NEW_SPACES_INTEGRATED" &&
        formContext.getAttribute("caps_submissioncategorycode").getValue() !== "CC_UPGRADE" && formContext.getAttribute("caps_submissioncategorycode").getValue() !== "CC_UPGRADE_MINOR") {

        var recordId = formContext.data.entity.getId().replace("{", "").replace("}", "");
        //call action
        var req = {};
        var target = { entityType: "caps_project", id: recordId };
        req.entity = target;

        req.getMetadata = function () {
            return {
                boundParameter: "entity",
                operationType: 0,
                operationName: "caps_ValidatePRFSAlternativeOptions",
                parameterTypes: {
                    "entity": {
                        "typeName": "mscrm.caps_project",
                        "structuralProperty": 5
                    }
                }
            };
        };

        Xrm.WebApi.online.execute(req).then(
            function (result) {

                if (result.ok) {
                    return result.json().then(
                        function (response) {

                            if (response.displayError) {
                                formContext.ui.setFormNotification('No PRFS Alternative Options have been provided on the PRFS tab.', 'INFO', PRFS_NOALTERNATIVES_NOTIFICATION);
                            }
                            else {
                                formContext.ui.clearFormNotification(PRFS_NOALTERNATIVES_NOTIFICATION);
                            }
                        });
                }
            },
            function (e) {
                formContext.ui.clearFormNotification(PRFS_NOALTERNATIVES_NOTIFICATION);
            }
        );
    }
}

/**
This function checks if there is at least one surrounding school specified if there is cashflow in the first 3 years.
*/
CAPS.Project.ValidatePRFSSurroundingSchools = function (executionContext) {
    var formContext = executionContext.getFormContext();
    if (formContext.getAttribute("caps_submissioncategorycode").getValue() !== "CC_CONVERSION" && formContext.getAttribute("caps_submissioncategorycode").getValue() !== "CC_CONVERSION_MINOR" &&
        formContext.getAttribute("caps_submissioncategorycode").getValue() !== "Major_CC_New_Spaces" && formContext.getAttribute("caps_submissioncategorycode").getValue() !== "CC_MAJOR_NEW_SPACES_INTEGRATED" &&
        formContext.getAttribute("caps_submissioncategorycode").getValue() !== "CC_UPGRADE" && formContext.getAttribute("caps_submissioncategorycode").getValue() !== "CC_UPGRADE_MINOR") {
        var recordId = formContext.data.entity.getId().replace("{", "").replace("}", "");
        //call action
        var req = {};
        var target = { entityType: "caps_project", id: recordId };
        req.entity = target;

        req.getMetadata = function () {
            return {
                boundParameter: "entity",
                operationType: 0,
                operationName: "caps_ValidatePRFSSurroundingSchools",
                parameterTypes: {
                    "entity": {
                        "typeName": "mscrm.caps_project",
                        "structuralProperty": 5
                    }
                }
            };
        };

        Xrm.WebApi.online.execute(req).then(
            function (result) {

                if (result.ok) {
                    return result.json().then(
                        function (response) {

                            if (response.displayError) {
                                formContext.ui.setFormNotification('No PRFS Surrounding Schools have been provided on the PRFS tab.', 'INFO', PRFS_NOSCHOOLS_NOTIFICATION);
                            }
                            else {
                                formContext.ui.clearFormNotification(PRFS_NOSCHOOLS_NOTIFICATION);
                            }
                        });
                }
            },
            function (e) {
                formContext.ui.clearFormNotification(PRFS_NOSCHOOLS_NOTIFICATION);
            }
        );
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

/**
 * This function set's the visibility and requirement level for a related supplemental cost item
 * @param {any} executionContext Execution Context
 * @param {any} toggleField logical name of toggle field (yes/no)
 * @param {any} displayField logical name of display field
 */
CAPS.Project.ToggleScheduleBSupplementalField = function (executionContext, toggleField, displayField) {
    var formContext = executionContext.getFormContext();

    if (formContext.getAttribute(toggleField) !== null && formContext.getAttribute(toggleField).getValue() === true) {
        //show display Field and make mandatory
        if (formContext.getAttribute(displayField) !== null) {
            formContext.getAttribute(displayField).setRequiredLevel("required");
            formContext.getControl(displayField).setVisible(true);
        }
    }
    else {
        if (formContext.getAttribute(displayField) !== null) {
            formContext.getAttribute(displayField).setRequiredLevel("none");
            if (formContext.getAttribute(displayField).getValue() != null && executionContext.getFormContext() != FORM_STATE.READ_ONLY) {
                formContext.getAttribute(displayField).setValue(null);
            }
            formContext.getControl(displayField).setVisible(false);
        }
    }
}

/**
 * This function set's the visibility and requirement level for other description when other cost is entered.
 * @param {any} executionContext Execution Context
 */
CAPS.Project.ToggleOtherSupplementalCostField = function (executionContext) {
    var formContext = executionContext.getFormContext();

    if (formContext.getAttribute("caps_othercost").getValue() !== null && formContext.getAttribute("caps_othercost").getValue() !== 0) {
        formContext.getAttribute("caps_othercostdescription").setRequiredLevel("required");
    }
    else {
        formContext.getAttribute("caps_othercostdescription").setRequiredLevel("none");
    }
}

/**
 * This function get's the related facility's school type and pre-set the value
 * @param {any} executionContext Execution Context
 */
CAPS.Project.GetSchoolType = function (executionContext) {

    var formContext = executionContext.getFormContext();

    var facility = formContext.getAttribute("caps_facility").getValue();

    if (facility !== null && facility[0] !== null) {

        Xrm.WebApi.retrieveRecord("caps_facility", facility[0].id, "?$select=_caps_currentfacilitytype_value").then(
            function success(result) {

                if (result !== null && result._caps_currentfacilitytype_value !== null) {
                    Xrm.WebApi.retrieveRecord("caps_facilitytype", result._caps_currentfacilitytype_value, "?$select=_caps_schooltype_value").then(
                        function success(result) {

                            var schoolType = new Array();
                            schoolType[0] = new Object();
                            schoolType[0].id = result._caps_schooltype_value;
                            schoolType[0].name = result["_caps_schooltype_value@OData.Community.Display.V1.FormattedValue"];
                            schoolType[0].entityType = result["_caps_schooltype_value@Microsoft.Dynamics.CRM.lookuplogicalname"];

                            formContext.getAttribute("caps_schooltype").setValue(schoolType);
                        }
                    );
                }

            },
            function (error) {
                console.log(error.message);
                // handle error conditions
            }
        );
    }

}

/**
 * This function locks the total project cost field if schedule b should be used to determine the project request cost.
 * @param {any} executionContext execution context
 */
CAPS.Project.ToggleRequiresScheduleB = function (executionContext) {
    var formContext = executionContext.getFormContext();
    var submissionCategoryCode = formContext.getAttribute("caps_submissioncategorycode").getValue();
    var attr = formContext.getAttribute("caps_totalprojectcost");

    if (submissionCategoryCode === "LEASE") {
        attr.controls.forEach(function (control) {
            control.setDisabled(true);
        });
    }
    else {
        if (formContext.getAttribute("caps_requiresscheduleb").getValue() === true) {
            //Lock Total Project Cost
            formContext.ui.tabs.get("tab_scheduleb").setVisible(true);
            attr.controls.forEach(function (control) {
                control.setDisabled(true);
            });
        }
        else {
            formContext.ui.tabs.get("tab_scheduleb").setVisible(false);
            //un-Lock Total Project Cost
            attr.controls.forEach(function (control) {
                control.setDisabled(false);
            });
        }
    }
}

/**
 * This function displays the correct schedule b fields based on the schedule b type
 * @param {any} executionContext execution context
 */
CAPS.Project.ToggleScheduleBFields = function (executionContext) {
    var formContext = executionContext.getFormContext();
    //Get Schedule B Type field on Project Type
    var projectType = formContext.getAttribute("caps_projecttype").getValue();
    var submissionCategoryCode = formContext.getAttribute("caps_submissioncategorycode").getValue();
    var hostSD = formContext.getAttribute("caps_schooldistrict").getValue();
    var includeNLC = formContext.getAttribute("caps_includenlc");

    if (hostSD !== null && !hostSD[0].name.includes("SD93")) {
        formContext.getControl("caps_hostschooldistrict").setVisible(false);
    }
    if (submissionCategoryCode === "LEASE") {
        includeNLC.controls.forEach(function (control) {
            control.setVisible(true);
        });

    }
    else {
        if (projectType !== null) {
            Xrm.WebApi.retrieveRecord("caps_projecttype", projectType[0].id, "?$select=caps_budgetcalculationtype").then(
                function success(result) {

                    var calcType = result.caps_budgetcalculationtype;

                    //New = 200,870,000
                    //Replacement = 200,870,001
                    //Partial Seismic = 200,870,004
                    //Seismic = 200,870,005
                    if (calcType === 200870000 || calcType === 200870001) {
                        //show NLC
                        includeNLC.controls.forEach(function (control) {
                            control.setVisible(true);
                        });
                        //formContext.getControl("caps_includenlc").setVisible(true);
                    }
                    else {
                        //hide and set NLS to false
                        //formContext.getControl("caps_includenlc").setVisible(false);
                        includeNLC.controls.forEach(function (control) {
                            control.setVisible(false);
                        });

                        if (formContext.getAttribute("caps_includenlc").getValue() != null && executionContext.getFormContext() != FORM_STATE.READ_ONLY) {
                            formContext.getAttribute("caps_includenlc").setValue(null);
                        }
                    }
                    if (calcType === 200870004 || calcType === 200870005) {
                        formContext.getControl("caps_constructioncostsspir").setVisible(true);
                    }
                    else {
                        formContext.getControl("caps_constructioncostsspir").setVisible(false);

                        if (formContext.getAttribute("caps_constructioncostsspir").getValue() != null && executionContext.getFormContext() != FORM_STATE.READ_ONLY) {
                            formContext.getAttribute("caps_constructioncostsspir").setValue(null);
                        }
                    }

                    if (calcType === 200870005) {
                        //blank values
                        if (formContext.getAttribute("caps_projectincludesdemolition").getValue() != false && executionContext.getFormContext() != FORM_STATE.READ_ONLY) {
                            formContext.getAttribute("caps_projectincludesdemolition").setValue(false);
                        }
                        if (formContext.getAttribute("caps_demolitioncost").getValue() != null && executionContext.getFormContext() != FORM_STATE.READ_ONLY) {
                            formContext.getAttribute("caps_demolitioncost").setValue(null);
                        }
                        if (formContext.getAttribute("caps_projectincludesabnormaltopography").getValue() != false && executionContext.getFormContext() != FORM_STATE.READ_ONLY) {
                            formContext.getAttribute("caps_projectincludesabnormaltopography").setValue(false);
                        }
                        if (formContext.getAttribute("caps_abnormaltopographycost").getValue() != null && executionContext.getFormContext() != FORM_STATE.READ_ONLY) {
                            formContext.getAttribute("caps_abnormaltopographycost").setValue(null);
                        }

                        //hide
                        formContext.getControl("caps_projectincludesdemolition").setVisible(false);
                        formContext.getControl("caps_demolitioncost").setVisible(false);
                        formContext.getControl("caps_projectincludesabnormaltopography").setVisible(false);
                        formContext.getControl("caps_abnormaltopographycost").setVisible(false);

                        //make optional
                        formContext.getAttribute("caps_demolitioncost").setRequiredLevel("none");
                        formContext.getAttribute("caps_abnormaltopographycost").setRequiredLevel("none");

                    }
                    else {
                        //show
                        formContext.getControl("caps_projectincludesdemolition").setVisible(true);
                        formContext.getControl("caps_demolitioncost").setVisible(formContext.getAttribute("caps_projectincludesdemolition").getValue());
                        formContext.getControl("caps_projectincludesabnormaltopography").setVisible(true);
                        formContext.getControl("caps_abnormaltopographycost").setVisible(formContext.getAttribute("caps_projectincludesabnormaltopography").getValue());
                    }
                },
                function (error) {
                    console.log(error.message);
                    // handle error conditions
                });
        }
    }
}

/**
 * This function calls all schedule b related functions
 * @param {any} executionContext Execution Context
 */
CAPS.Project.SetupScheduleB = function (executionContext) {
    var formContext = executionContext.getFormContext();

    CAPS.Project.ToggleScheduleBFields(executionContext);
    formContext.getAttribute("caps_projecttype").addOnChange(CAPS.Project.ToggleScheduleBFields);

    CAPS.Project.ToggleScheduleBSupplementalField(executionContext, "caps_projectincludesdemolition", "caps_demolitioncost");
    CAPS.Project.ToggleScheduleBSupplementalField(executionContext, "caps_projectincludesabnormaltopography", "caps_abnormaltopographycost");
    CAPS.Project.ToggleScheduleBSupplementalField(executionContext, "caps_projectincludestemporaryaccommodation", "caps_temporaryaccommodationcost");
    CAPS.Project.ToggleOtherSupplementalCostField(executionContext);
    CAPS.Project.ToggleRequiresScheduleB(executionContext);
    //add on change for project type to get and set school type

    //Add on change events
    formContext.getAttribute("caps_projectincludesdemolition").addOnChange(function () { CAPS.Project.ToggleScheduleBSupplementalField(executionContext, "caps_projectincludesdemolition", "caps_demolitioncost"); });
    formContext.getAttribute("caps_projectincludesabnormaltopography").addOnChange(function () { CAPS.Project.ToggleScheduleBSupplementalField(executionContext, "caps_projectincludesabnormaltopography", "caps_abnormaltopographycost"); });
    formContext.getAttribute("caps_projectincludestemporaryaccommodation").addOnChange(function () { CAPS.Project.ToggleScheduleBSupplementalField(executionContext, "caps_projectincludestemporaryaccommodation", "caps_temporaryaccommodationcost"); });
    formContext.getAttribute("caps_othercost").addOnChange(CAPS.Project.ToggleOtherSupplementalCostField);
    formContext.getAttribute("caps_facility").addOnChange(CAPS.Project.GetSchoolType);
    formContext.getAttribute("caps_requiresscheduleb").addOnChange(CAPS.Project.ToggleRequiresScheduleB);

}

/**
 * This function shows age of existing playground field if this is a PEP replacement
 * @param {any} executionContext execution context
 */
CAPS.Project.TogglePEPReplacement = function (executionContext) {

    var formContext = executionContext.getFormContext();
    var projectType = formContext.getAttribute("caps_projecttype").getValue();

    if (projectType !== null) {

        //get project type record
        Xrm.WebApi.retrieveRecord("caps_projecttype", projectType[0].id, "?$select=caps_projecttypecode").then(
            function success(record) {
                if (record.caps_projecttypecode == "PEP_REPLACEMENT") {
                    formContext.getControl("caps_ageofexistingplayground").setVisible(true);
                }
                else {
                    formContext.getControl("caps_ageofexistingplayground").setVisible(false);
                }
            },
            function (error) {
                console.log(error.message);
                // handle error conditions
            }
        );
    }

}

/**
 * This function shows the bus lookup field and makes it mandatory on a Bus Replacement project.
 * @param {any} executionContext execution context
 */
CAPS.Project.ToggleBUSReplacement = function (executionContext) {

    var formContext = executionContext.getFormContext();
    var projectType = formContext.getAttribute("caps_projecttype").getValue();

    if (projectType !== null) {

        Xrm.WebApi.retrieveRecord("caps_projecttype", projectType[0].id, "?$select=caps_projecttypecode").then(
            function success(record) {
                if (record.caps_projecttypecode == "BUS_REPLACEMENT") {
                    //Show Bus to be replaced and make mandatory
                    formContext.getControl("caps_bus").setVisible(true);
                    formContext.getAttribute("caps_bus").setRequiredLevel("required");
                }
                else {
                    //hide Bus to be replaced and make optional and clear
                    formContext.getControl("caps_bus").setVisible(false);
                    formContext.getAttribute("caps_bus").setRequiredLevel("none");
                    if (formContext.getAttribute("caps_bus").getValue() != null && executionContext.getFormContext() != FORM_STATE.READ_ONLY) {
                        formContext.getAttribute("caps_bus").setValue(null);
                    }
                }
            },
            function (error) {
                console.log(error.message);
                // handle error conditions
            }
        );


    }

}

/**
 * This function shows/hides the Date LRFP will be in place when an LRFP record is selected/removed.
 * @param {any} executionContext execution context
 */
CAPS.Project.ToggleLRFP = function (executionContext) {
    var formContext = executionContext.getFormContext();


    if (formContext.getAttribute("caps_longrangefacilityplan").getValue()) {
        formContext.getControl("caps_datelrfpwillbeinplace").setVisible(false);
        if (formContext.getAttribute("caps_datelrfpwillbeinplace").getValue() != null && executionContext.getFormContext() != FORM_STATE.READ_ONLY) {
            formContext.getAttribute("caps_datelrfpwillbeinplace").setValue(null);
        }

    }
    else {
        formContext.getControl("caps_datelrfpwillbeinplace").setVisible(true);
    }

}

/**
 * This function makes the Bus New Route Information file field manadatory or if the issue type is New Route.  It also makes inspection report mandatory if issue type is Mechanical or Safety
 * @param {any} executionContext execution context
 */
CAPS.Project.ToggleBusIssueType = function (executionContext) {
    var formContext = executionContext.getFormContext();

    //New Route = 100000003
    if (formContext.getAttribute("caps_issuetype").getValue() === 100000003) {
        formContext.getAttribute("caps_bus_newrouteinformation").setRequiredLevel("required");
    }
    else {
        formContext.getAttribute("caps_bus_newrouteinformation").setRequiredLevel("none");
    }

    //Mechanical = 100000000; Safety = 100000001
    if (formContext.getAttribute("caps_issuetype").getValue() === 100000000 || formContext.getAttribute("caps_issuetype").getValue() === 100000001) {
        formContext.getAttribute("caps_bus_inspectionreport").setRequiredLevel("required");
    }
    else {
        formContext.getAttribute("caps_bus_inspectionreport").setRequiredLevel("none");
    }
}

/***
This function validates that the kindergarten design capacity is a multiple of the specified design capacity and shows a warning if it isn't.
***/
CAPS.Project.ValidateKindergartenDesignCapacity = function (executionContext) {

    var formContext = executionContext.getFormContext();
    var designCapacity = formContext.getAttribute("caps_changeindesigncapacitykindergarten").getValue();
    var validateDesignCapacityKRequest = new CAPS.Project.ValidateDesignCapacityRequest("Kindergarten", designCapacity);

    Xrm.WebApi.online.execute(validateDesignCapacityKRequest).then(
        function (result) {
            if (result.ok) {

                return result.json().then(
                    function (response) {

                        if (!response.IsValid) {
                            formContext.ui.setFormNotification(response.ValidationMessage, 'INFO', 'KINDERGARTEN DESIGN WARNING');
                        }
                        else {
                            formContext.ui.clearFormNotification('KINDERGARTEN DESIGN WARNING');
                        }
                    });
            }
        },
        function (error) {
            console.log(error.message);
            // handle error conditions
        }
    );
}

/***
This function validates that the elementary design capacity is a multiple of the specified design capacity and shows a warning if it isn't.
**/
CAPS.Project.ValidateElementaryDesignCapacity = function (executionContext) {

    var formContext = executionContext.getFormContext();
    var designCapacity = formContext.getAttribute("caps_changeindesigncapacityelementary").getValue();
    var validateDesignCapacityKRequest = new CAPS.Project.ValidateDesignCapacityRequest("Elementary", designCapacity);

    Xrm.WebApi.online.execute(validateDesignCapacityKRequest).then(
        function (result) {
            if (result.ok) {

                return result.json().then(
                    function (response) {

                        if (!response.IsValid) {
                            formContext.ui.setFormNotification(response.ValidationMessage, 'INFO', 'ELEMENTARY DESIGN WARNING');
                        }
                        else {
                            formContext.ui.clearFormNotification('ELEMENTARY DESIGN WARNING');
                        }
                    });
            }
        },
        function (error) {
            console.log(error.message);
            // handle error conditions
        }
    );
};

/***
This function validates that the secondary design capacity is a multiple of the specified design capacity and shows a warning if it isn't.
**/
CAPS.Project.ValidateSecondaryDesignCapacity = function (executionContext) {

    var formContext = executionContext.getFormContext();
    var designCapacity = formContext.getAttribute("caps_changeindesigncapacitysecondary").getValue();
    var validateDesignCapacityKRequest = new CAPS.Project.ValidateDesignCapacityRequest("Secondary", designCapacity);

    Xrm.WebApi.online.execute(validateDesignCapacityKRequest).then(
        function (result) {
            if (result.ok) {

                return result.json().then(
                    function (response) {

                        if (!response.IsValid) {
                            formContext.ui.setFormNotification(response.ValidationMessage, 'INFO', 'SECONDARY DESIGN WARNING');
                        }
                        else {
                            formContext.ui.clearFormNotification('SECONDARY DESIGN WARNING');
                        }
                    });
            }
        },
        function (error) {
            console.log(error.message);
            // handle error conditions
        }
    );
};

CAPS.Project.ValidateDesignCapacityRequest = function (capacityType, capacityCount) {
    this.Type = capacityType;
    this.Count = capacityCount;
};

CAPS.Project.ValidateDesignCapacityRequest.prototype.getMetadata = function () {
    return {
        boundParameter: null,
        parameterTypes: {
            "Type": {
                "typeName": "Edm.String",
                "structuralProperty": 1
            },
            "Count": {
                "typeName": "Edm.Int32",
                "structuralProperty": 1
            }
        },
        operationType: 0, // This is a function. Use '0' for actions and '2' for CRUD
        operationName: "caps_ValidateDesignCapacity"
    };
};

/**
This function validates the current FCI and displays a notification if it's greater than 1.
*/
CAPS.Project.ValidateCurrentFCI = function (executionContext) {

    var formContext = executionContext.getFormContext();

    var fci = formContext.getAttribute("caps_currentfci");

    if (fci != null) {
        if (fci.getValue() > 1) {
            formContext.ui.setFormNotification('Current FCI is usually equal to or less than 1, please make sure it is correct.', 'INFO', CURRENT_FCI_NOTIFICATION);
        }
        else {
            formContext.ui.clearFormNotification(CURRENT_FCI_NOTIFICATION);
        }
    }
}

/**
This function validates the future FCI and displays a notification if it's greater than 1.
*/
CAPS.Project.ValidateFutureFCI = function (executionContext) {

    var formContext = executionContext.getFormContext();

    var fci = formContext.getAttribute("caps_futurefacilityconditionindex");

    if (fci != null) {
        if (fci.getValue() > 1) {
            formContext.ui.setFormNotification('Future FCI is usually equal to or less than 1, please make sure it is correct.', 'INFO', FUTURE_FCI_NOTIFICATION);
        }
        else {
            formContext.ui.clearFormNotification(FUTURE_FCI_NOTIFICATION);
        }
    }
}

CAPS.Project.ToggleExistingChildCareFacility = function (executionContext) {

    var formContext = executionContext.getFormContext();
    var existingChildCareFacility = formContext.getAttribute("caps_existingchildcarefacility").getValue();
    var existingFacility = formContext.getAttribute("caps_existingfacility").getValue();
    var submissionCategoryCode = formContext.getAttribute("caps_submissioncategorycode").getValue();
    if (submissionCategoryCode === "CC_UPGRADE" || submissionCategoryCode === "CC_UPGRADE_MINOR") {
        formContext.getAttribute("caps_facility").controls.forEach(control => control.setDisabled(true));
    }
    else if (submissionCategoryCode === "CC_CONVERSION_MINOR" || submissionCategoryCode === "CC_CONVERSION") {
        if (existingChildCareFacility === true) {
            formContext.getAttribute("caps_facility").controls.forEach(control => control.setVisible(true));
            formContext.getAttribute("caps_facility").setRequiredLevel("required");
            formContext.getAttribute("caps_facility").controls.forEach(control => control.setDisabled(true));
            formContext.getAttribute("caps_childcare").controls.forEach(control => control.setVisible(true)); //Show Child Care Facility
            formContext.getAttribute("caps_childcare").setRequiredLevel("required");
            formContext.getAttribute("caps_proposedchildcarefacility").controls.forEach(control => control.setVisible(false)); // Hide Proposed Child Care Facility
            formContext.getAttribute("caps_proposedchildcarefacility").setValue(null); //Wipe Proposed Child Care Facility

        }
        else {
            formContext.getAttribute("caps_facility").controls.forEach(control => control.setVisible(true));
            formContext.getAttribute("caps_facility").controls.forEach(control => control.setDisabled(false));
            formContext.getAttribute("caps_facility").setRequiredLevel("required");
            formContext.getAttribute("caps_childcare").controls.forEach(control => control.setVisible(false)); //Show Child Care Facility
            formContext.getAttribute("caps_childcare").setValue(null); //Wipe Child Care Facility
            formContext.getAttribute("caps_childcare").setRequiredLevel("none");
            formContext.getAttribute("caps_proposedchildcarefacility").controls.forEach(control => control.setVisible(true)); // Show Proposed Child Care Facility

        }
    }
    else {

        if (existingChildCareFacility === true) {
            formContext.getAttribute("caps_childcare").controls.forEach(control => control.setVisible(true)); //Show Child Care Facility
            formContext.getAttribute("caps_childcare").setRequiredLevel("required");
            formContext.getAttribute("caps_proposedchildcarefacility").controls.forEach(control => control.setVisible(false)); // Hide Proposed Child Care Facility
            formContext.getAttribute("caps_proposedchildcarefacility").setValue(null); //Wipe Proposed Child Care Facility
            formContext.getAttribute("caps_proposedfacility").controls.forEach(control => control.setVisible(false)); // Hide Proposed Child Care Facility
            formContext.getAttribute("caps_proposedfacility").setValue(null); //Wipe Proposed Child Care Facility
            formContext.getAttribute("caps_facility").controls.forEach(control => control.setVisible(true)); //Hide School Facility
            formContext.getAttribute("caps_facility").controls.forEach(control => control.setDisabled(true));
            formContext.getAttribute("caps_existingfacility").controls.forEach(control => control.setVisible(false)); //Hide Existing Facility

        }
        else {
            formContext.getAttribute("caps_childcare").controls.forEach(control => control.setVisible(false)); //Hide Child Care Facility
            formContext.getAttribute("caps_childcare").setValue(null); //Wipe Child Care Facility
            formContext.getAttribute("caps_childcare").setRequiredLevel("none");
            formContext.getAttribute("caps_proposedchildcarefacility").controls.forEach(control => control.setVisible(true)); // Show Proposed Child Care Facility
            formContext.getAttribute("caps_existingfacility").controls.forEach(control => control.setVisible(true)); //Show Existing Facility
            if (existingFacility === true) {
                formContext.getAttribute("caps_facility").controls.forEach(control => control.setVisible(true)); //Show School Facility
                formContext.getAttribute("caps_facility").controls.forEach(control => control.setDisabled(false));
                formContext.getAttribute("caps_proposedfacility").controls.forEach(control => control.setVisible(false)); //Hide Proposed School Facility
                formContext.getAttribute("caps_proposedfacility").setValue(null);
            }
            else {
                formContext.getAttribute("caps_facility").controls.forEach(control => control.setVisible(false)); //Hide School Facility
                formContext.getAttribute("caps_facility").setValue(null);
                formContext.getAttribute("caps_proposedfacility").controls.forEach(control => control.setVisible(true)); //Show Proposed School Facility
            }



        }
    }

}

CAPS.Project.WipeSchoolFacility = function (executionContext) {
    var formContext = executionContext.getFormContext();
    formContext.getAttribute("caps_facility").setValue(null);
}

CAPS.Project.ShowHideSubgrid = function (executionContext) {

    var formContext = executionContext.getFormContext();
    var relocatingExistingCCSpaces = formContext.getAttribute("caps_areyourelocatingexistingchildcarespaces").getValue();

    if (relocatingExistingCCSpaces === false) {
        formContext.getControl("Existing_Childcare_spaces").setVisible(false);
        formContext.ui.tabs.get("tab_ChildCare").sections.get("tab_ChildCare_section_11").setVisible(false);

    }
    else {
        formContext.getControl("Existing_Childcare_spaces").setVisible(true);
        formContext.ui.tabs.get("tab_ChildCare").sections.get("tab_ChildCare_section_11").setVisible(true);
    }

}

CAPS.Project.ShowConfirmationRelocatingCC = function (executionContext) {

    var formContext = executionContext.getFormContext();
    var relocatingExistingCCSpaces = formContext.getAttribute("caps_areyourelocatingexistingchildcarespaces").getValue();

    if (relocatingExistingCCSpaces === false) {
        var confirmStrings = { text: "Setting this to No will deactivate all the related relocation records. You will need to re-create them if you set it back to Yes. Are you sure you want to continue?", title: "Relocating Existing Child Care Spaces" };
        var confirmOptions = { height: 200, width: 450 };
        Xrm.Navigation.openConfirmDialog(confirmStrings, confirmOptions).then(
            function (success) {
                if (success.confirmed) {

                    formContext.getAttribute("caps_runflow").setValue(true);
                    formContext.data.save();
                }

                else {
                    formContext.getControl("Existing_Childcare_spaces").setVisible(true);
                    formContext.ui.tabs.get("tab_ChildCare").sections.get("tab_ChildCare_section_11").setVisible(true);
                    formContext.getAttribute("caps_areyourelocatingexistingchildcarespaces").setValue(true);
                    formContext.data.save();
                }

            });
    }

}


CAPS.Project.ShowHideFieldsCCPBTab = function (executionContext) {
    var formContext = executionContext.getFormContext();
    var projectRequiresUniqueSiteDev = formContext.getAttribute("caps_projectrequireuniquesitedevelopment").getValue();
    if (projectRequiresUniqueSiteDev === false) {
        formContext.getAttribute("caps_ifyespleasedescribeuniquesiterequirement").controls.forEach(control => control.setVisible(false));
    }
    else {
        formContext.getAttribute("caps_ifyespleasedescribeuniquesiterequirement").controls.forEach(control => control.setVisible(true));
    }
}

CAPS.Project.PopulateSchoolFacility = function (executionContext) {

    var formContext = executionContext.getFormContext();
    var childCareFacility = CAPS.Project.GetLookup("caps_childcare", formContext);
    //if (childCareFacility === undefined) {
    //    formContext.getAttribute("caps_facility").setValue(null);
    //}
    if (childCareFacility !== undefined) {
        var childCareFacilityId = CAPS.Project.RemoveCurlyBraces(childCareFacility.id);
        var options = "?$select=_caps_facility_value";
        Xrm.WebApi.retrieveRecord("caps_childcare", childCareFacilityId, options).then(
            function success(result) {

                //var facilityName = result.caps_name;
                var facilityId = result._caps_facility_value;
                var facilityName = result["_caps_facility_value@OData.Community.Display.V1.FormattedValue"];

                var schoolFacility = new Array();
                schoolFacility[0] = new Object();
                schoolFacility[0].id = facilityId;
                schoolFacility[0].name = facilityName;
                schoolFacility[0].entityType = "caps_facility";
                formContext.getAttribute("caps_facility").setValue(schoolFacility);
            },
            function (error) {

                console.log(error.message);
            }
        )
    }
}

CAPS.Project.WipeSchoolFacilityWhenCCWiped = function (executionContext) {
    var formContext = executionContext.getFormContext();
    var childCareFacility = CAPS.Project.GetLookup("caps_childcare", formContext);
    if (childCareFacility === undefined) {
        formContext.getAttribute("caps_facility").setValue(null);
    }
}

CAPS.Project.HideFundingRequested = function (executionContext) {
    var formContext = executionContext.getFormContext();
    var submissionCategoryCode = formContext.getAttribute("caps_submissioncategorycode").getValue();
    if (submissionCategoryCode === "CC_CONVERSION" || submissionCategoryCode === "CC_UPGRADE" ||
        submissionCategoryCode === "Major_CC_New_Spaces" || submissionCategoryCode === "CC_MAJOR_NEW_SPACES_INTEGRATED") {
        formContext.getAttribute("caps_fundingrequested").controls.forEach(control => control.setVisible(false));
    }
}

CAPS.Project.ShowHideHealthAuthority = function (executionContext) {
    var formContext = executionContext.getFormContext();
    var submissionCategoryCode = formContext.getAttribute("caps_submissioncategorycode").getValue();
    var existingChildCareFacility = formContext.getAttribute("caps_existingchildcarefacility").getValue();

    if (submissionCategoryCode === "Major_CC_New_Spaces" || submissionCategoryCode === "CC_MAJOR_NEW_SPACES_INTEGRATED" ||
        submissionCategoryCode === "CC_CONVERSION" || submissionCategoryCode === "CC_UPGRADE" || submissionCategoryCode === "CC_CONVERSION_MINOR") {
        if (existingChildCareFacility) {
            formContext.getAttribute("caps_healthauthority").controls.forEach(control => control.setVisible(false));
            formContext.getAttribute("caps_healthauthority").setRequiredLevel("none");
            var healthAuthority = CAPS.Project.GetLookup("caps_healthauthority", formContext);
            if (healthAuthority !== undefined) {
                formContext.getAttribute("caps_healthauthority").setValue(null);
            }

        }
        else {
            formContext.getAttribute("caps_healthauthority").controls.forEach(control => control.setVisible(true));
            formContext.getAttribute("caps_healthauthority").setRequiredLevel("required");
        }
    }
}

CAPS.Project.ShowNotificationHealthAuthority = function (executionContext) {

    var formContext = executionContext.getFormContext();
    var childCareFacility = CAPS.Project.GetLookup("caps_childcare", formContext);
    var existingChildCareFacility = formContext.getAttribute("caps_existingchildcarefacility").getValue();
    var submissionCategoryCode = formContext.getAttribute("caps_submissioncategorycode").getValue();

    if (submissionCategoryCode === "Major_CC_New_Spaces" || submissionCategoryCode === "CC_MAJOR_NEW_SPACES_INTEGRATED" ||
        submissionCategoryCode === "CC_CONVERSION" || submissionCategoryCode === "CC_UPGRADE") {

        if (childCareFacility !== undefined && existingChildCareFacility === true) {
            Xrm.WebApi.retrieveRecord("caps_childcare", CAPS.Project.RemoveCurlyBraces(childCareFacility.id), "?$select=_caps_healthauthority_value").then(
                function success(result) {

                    var healthAuthoriy = result._caps_healthauthority_value;
                    if (healthAuthoriy === null) {
                        formContext.ui.setFormNotification('Health Authority on Child Care Facility should be filled in.', 'INFO', HEALTH_AUTHORITY_NOTIFICATION);
                    }
                    else {
                        formContext.ui.clearFormNotification(HEALTH_AUTHORITY_NOTIFICATION);
                    }
                },
                function (error) {

                    console.log(error.message);
                    // handle error conditions
                }
            );
        }
        else {
            formContext.ui.clearFormNotification(HEALTH_AUTHORITY_NOTIFICATION);
        }
    }
}

CAPS.Project.ShowHidePopulationGroup = function (executionContext) {
    var formContext = executionContext.getFormContext();
    var programTailoredToPopulation = formContext.getAttribute("caps_programtailoredtomeetneedsofpopulationgroup").getValue();
    if (programTailoredToPopulation) {
        formContext.getControl("caps_ifyeswhichofthefollowingapply").setVisible(true);
    }
    else {
        formContext.getControl("caps_ifyeswhichofthefollowingapply").setVisible(false);
        formContext.getAttribute("caps_ifyeswhichofthefollowingapply").setValue(null);
    }
}

CAPS.Project.HideTotalChildCareSpace = function (executionContext) {
    var formContext = executionContext.getFormContext();
    var areYouRelocatingExistingCCSpace = formContext.getAttribute("caps_areyourelocatingexistingchildcarespaces").getValue();
    var submissionCategoryCode = formContext.getAttribute("caps_submissioncategorycode").getValue();

    if (areYouRelocatingExistingCCSpace === false && submissionCategoryCode === "Major_CC_New_Spaces") {
        formContext.ui.tabs.get("tab_MAJOR_CC_NEW_SPACE").sections.get("tab_MAJOR_CC_NEW_SPACE_section_4").setVisible(false);
        formContext.ui.tabs.get("tab_MAJOR_CC_NEW_SPACE").sections.get("tab_MAJOR_CC_NEW_SPACE_SECTION_RELOCATED").setVisible(false);
    }
    else if (areYouRelocatingExistingCCSpace === true && submissionCategoryCode === "Major_CC_New_Spaces") {
        formContext.ui.tabs.get("tab_MAJOR_CC_NEW_SPACE").sections.get("tab_MAJOR_CC_NEW_SPACE_section_4").setVisible(true);
        formContext.ui.tabs.get("tab_MAJOR_CC_NEW_SPACE").sections.get("tab_MAJOR_CC_NEW_SPACE_SECTION_RELOCATED").setVisible(true);
    }
    if (areYouRelocatingExistingCCSpace === false && submissionCategoryCode === "CC_MAJOR_NEW_SPACES_INTEGRATED") {
        formContext.ui.tabs.get("tab_MAJOR_CC_NEW_SPACE_INTEGRATED").sections.get("tab_MAJOR_CC_NEW_SPACE_INTEGRATED_section_4").setVisible(false);
        formContext.ui.tabs.get("tab_MAJOR_CC_NEW_SPACE_INTEGRATED").sections.get("tab_MAJOR_CC_NEW_SPACE_INTEGRATED_SECTION_RELOCATED").setVisible(false);
    }
    else if (areYouRelocatingExistingCCSpace === true && submissionCategoryCode === "CC_MAJOR_NEW_SPACES_INTEGRATED") {
        formContext.ui.tabs.get("tab_MAJOR_CC_NEW_SPACE_INTEGRATED").sections.get("tab_MAJOR_CC_NEW_SPACE_INTEGRATED_section_4").setVisible(true);
        formContext.ui.tabs.get("tab_MAJOR_CC_NEW_SPACE_INTEGRATED").sections.get("tab_MAJOR_CC_NEW_SPACE_INTEGRATED_SECTION_RELOCATED").setVisible(true);
    }
    if (areYouRelocatingExistingCCSpace === false && submissionCategoryCode === "CC_CONVERSION") {
        formContext.ui.tabs.get("tab_MAJOR_CC_CONVERSION").sections.get("tab_MAJOR_CC_CONVERSION_section_4").setVisible(false);
        formContext.ui.tabs.get("tab_MAJOR_CC_CONVERSION").sections.get("tab_MAJOR_CC_CONVERSION_SECTION_RELOCATED").setVisible(false);
    }
    else if (areYouRelocatingExistingCCSpace === true && submissionCategoryCode === "CC_CONVERSION") {
        formContext.ui.tabs.get("tab_MAJOR_CC_CONVERSION").sections.get("tab_MAJOR_CC_CONVERSION_section_4").setVisible(true);
        formContext.ui.tabs.get("tab_MAJOR_CC_CONVERSION").sections.get("tab_MAJOR_CC_CONVERSION_SECTION_RELOCATED").setVisible(true);
    }
    if (areYouRelocatingExistingCCSpace === false && submissionCategoryCode === "CC_CONVERSION_MINOR") {
        formContext.ui.tabs.get("tab_minor_cc_conversion").sections.get("tab_minor_cc_conversion_section_4").setVisible(false);
        formContext.ui.tabs.get("tab_minor_cc_conversion").sections.get("tab_MINOR_CC_CONVERSION_SECTION_RELOCATED").setVisible(false);
    }
    else if (areYouRelocatingExistingCCSpace === true && submissionCategoryCode === "CC_CONVERSION_MINOR") {
        formContext.ui.tabs.get("tab_minor_cc_conversion").sections.get("tab_minor_cc_conversion_section_4").setVisible(true);
        formContext.ui.tabs.get("tab_minor_cc_conversion").sections.get("tab_MINOR_CC_CONVERSION_SECTION_RELOCATED").setVisible(true);
    }
    if (areYouRelocatingExistingCCSpace === false && submissionCategoryCode === "CC_UPGRADE_MINOR") {
        formContext.ui.tabs.get("tab_cc_minor_upgrade").sections.get("tab_cc_minor_upgrade_section_4").setVisible(false);
        formContext.ui.tabs.get("tab_cc_minor_upgrade").sections.get("tab_CC_MINOR_UPGRADE_SECTION_RELOCATED").setVisible(false);
    }
    else if (areYouRelocatingExistingCCSpace === true && submissionCategoryCode === "CC_UPGRADE_MINOR") {
        formContext.ui.tabs.get("tab_cc_minor_upgrade").sections.get("tab_cc_minor_upgrade_section_4").setVisible(true);
        formContext.ui.tabs.get("tab_cc_minor_upgrade").sections.get("tab_CC_MINOR_UPGRADE_SECTION_RELOCATED").setVisible(true);
    }
    if (areYouRelocatingExistingCCSpace === false && submissionCategoryCode === "CC_UPGRADE") {
        formContext.ui.tabs.get("tab_MAJOR_CC_UPGRADE").sections.get("tab_MAJOR_CC_UPGRADE_SECTION_UPGRADE_RELOCATED").setVisible(false);
    }
    else if (areYouRelocatingExistingCCSpace === true && submissionCategoryCode === "CC_UPGRADE") {
        formContext.ui.tabs.get("tab_MAJOR_CC_UPGRADE").sections.get("tab_MAJOR_CC_UPGRADE_SECTION_UPGRADE_RELOCATED").setVisible(true);
    }
}
CAPS.Project.preFilterProjectGroupLookup = function (executionContext) {
    debugger;
    var formContext = executionContext.getFormContext();
    if (formContext.getControl("caps_projectcollection") === null)
        return;

    formContext.getControl("caps_projectcollection").addPreSearch(CAPS.Project.AddProjectGroupLookupFilter);
}
CAPS.Project.AddProjectGroupLookupFilter = function (executionContext) {
    debugger;
    var formContext = executionContext.getFormContext();
    var schoolFacility = CAPS.Project.GetLookup("caps_facility", formContext);
    var proposedSchoolFacility = formContext.getAttribute("caps_proposedfacility").getValue();
    var proposedSite = formContext.getAttribute("caps_proposedsite").getValue();

    if (schoolFacility !== undefined && schoolFacility.id !== null) {
        let fetchXml = "<filter type='and'><condition attribute='caps_schoolfacility' operator='eq' value='" + schoolFacility.id + "' /></filter>";
        formContext.getControl("caps_projectcollection").addCustomFilter(fetchXml);
    }
    else if (proposedSchoolFacility !== null) {
        let fetchXml = "<filter type='and'><condition attribute='caps_proposedschoolfacility' operator='eq' value='" + proposedSchoolFacility + "' /></filter>";
        formContext.getControl("caps_projectcollection").addCustomFilter(fetchXml);
    }
    else if (proposedSite !== null) {
        let fetchXml = "<filter type='and'><condition attribute='caps_proposedschoolfacility' operator='eq' value='" + proposedSite + "' /></filter>";
        formContext.getControl("caps_projectcollection").addCustomFilter(fetchXml);
    }

    //var existingChildCareFacility = formContext.getAttribute("caps_existingfacility").getValue();
    //var noSchoolFacility = "00000000-0000-0000-0000-000000000000";
    //if (existingChildCareFacility) {
    //    var schoolFacility = CAPS.Project.GetLookup("caps_facility", formContext);
    //    if (schoolFacility !== undefined) {
    //        var fetchXml = "<filter type='and'><condition attribute='caps_schoolfacility' operator='eq' value='" + schoolFacility.id + "' /></filter>";
    //        formContext.getControl("caps_projectcollection").addCustomFilter(fetchXml);
    //    }
    //    else if (schoolFacility === undefined) {
    //        var fetchXml = "<filter type='and'><condition attribute='caps_schoolfacility' operator='eq' value='" + noSchoolFacility + "' /></filter>";
    //        formContext.getControl("caps_projectcollection").addCustomFilter(fetchXml);
    //    }
    //}
    //else if (!existingChildCareFacility) {
    //    var proposedSchoolFacility = formContext.getAttribute("caps_proposedfacility").getValue();
    //    if (proposedSchoolFacility !== null) {
    //        var fetchXml = "<filter type='and'><condition attribute='caps_proposedschoolfacility' operator='eq' value='" + proposedSchoolFacility + "' /></filter>";
    //        formContext.getControl("caps_projectcollection").addCustomFilter(fetchXml);
    //    }
    //    else if (proposedSchoolFacility === null) {
    //        var fetchXml = "<filter type='and'><condition attribute='caps_proposedschoolfacility' operator='eq' value='" + " " + "' /></filter>";
    //        formContext.getControl("caps_projectcollection").addCustomFilter(fetchXml);
    //    }
    //}

}
CAPS.Project.GetLookup = function (fieldName, formContext) {
    var lookupFieldObject = formContext.data.entity.attributes.get(fieldName);
    if (lookupFieldObject !== null && lookupFieldObject.getValue() !== null && lookupFieldObject.getValue()[0] !== null) {
        var entityId = lookupFieldObject.getValue()[0].id;
        var entityName = lookupFieldObject.getValue()[0].entityType;
        var entityLabel = lookupFieldObject.getValue()[0].name;
        var obj = {
            id: entityId,
            type: entityName,
            value: entityLabel
        };
        return obj;
    }
}
CAPS.Project.RemoveCurlyBraces = function (str) {
    return str.replace(/[{}]/g, "");
}

