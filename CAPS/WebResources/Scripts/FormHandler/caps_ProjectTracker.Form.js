"use strict";

var CAPS = CAPS || {};
CAPS.ProjectTracker = CAPS.ProjectTracker || {
    Status: null
};

const MAX_BUDGET_NOTIFICATION = "Max_Budget_Notification";
const PROVINCIAL_BUDGET_NOTIFICATION = "Provincial_Budget_Notification";
const STALE_DATES_NOTIFICATION = "Stale_Dates_Notification";
const CLOSED_DATE_NOTIFICATION = "Closed_Date_Notification";

/**
 * Main function for Project Tracker.  This function calls all other form functions.
 * @param {any} executionContext the form's execution context.
 */
CAPS.ProjectTracker.gridContext = null;
CAPS.ProjectTracker.onLoad = function (executionContext) {
    var formContext = executionContext.getFormContext();
    CAPS.ProjectTracker.gridContext = formContext.getControl("Subgrid_5").getGrid(); // Project Cashflow subgrid
    //Show Summary Report
    // CAPS.ProjectTracker.showSummaryReport(formContext); // CAPS-1949 Get Rid of Summary Report

    //Format the form by showing and hiding the relevant sections and fields
    CAPS.ProjectTracker.showHideCategoryRelevantSections(formContext);

    CAPS.ProjectTracker.ShowHideProgressReport(executionContext);

    CAPS.ProjectTracker.ShowBudgetMissmatch(executionContext);

    CAPS.ProjectTracker.ShowStaleDates(executionContext);

    formContext.getAttribute("caps_maxpotential_fundingsource").addOnChange(CAPS.ProjectTracker.ShowBudgetMissmatch);
    formContext.getAttribute("caps_maxpotentialprojectbudget").addOnChange(CAPS.ProjectTracker.ShowBudgetMissmatch);
    formContext.getAttribute("caps_provincial").addOnChange(CAPS.ProjectTracker.ShowBudgetMissmatch);
    formContext.getAttribute("caps_totalprovincialbudget").addOnChange(CAPS.ProjectTracker.ShowBudgetMissmatch);

    formContext.getControl("Subgrid_5").addOnLoad(CAPS.ProjectTracker.CalculateTotalCashFlowFields);

    //Design Capacity Checks
    CAPS.ProjectTracker.ValidateKindergartenDesignCapacity(executionContext);
    formContext.getAttribute("caps_designcapacitykindergarten").addOnChange(CAPS.ProjectTracker.ValidateKindergartenDesignCapacity);

    CAPS.ProjectTracker.ValidateElementaryDesignCapacity(executionContext);
    formContext.getAttribute("caps_designcapacityelementary").addOnChange(CAPS.ProjectTracker.ValidateElementaryDesignCapacity);

    CAPS.ProjectTracker.ValidateSecondaryDesignCapacity(executionContext);
    formContext.getAttribute("caps_designcapacitysecondary").addOnChange(CAPS.ProjectTracker.ValidateSecondaryDesignCapacity);

    //Show/Hide Facilities Grid
    CAPS.ProjectTracker.ShowHideFacilities(executionContext);

    //Show/Hide Project Closure Tab
    CAPS.ProjectTracker.ShowHideProjectClosureTab(executionContext);
    formContext.getAttribute("caps_projectclosure").addOnChange(CAPS.ProjectTracker.ShowHideProjectClosureTab);

    //Register Cashflow Update Completed Button Event
    if (formContext.getAttribute("caps_updatecompletedbutton") != null) {
        formContext.getAttribute("caps_updatecompletedbutton").addOnChange(CAPS.ProjectTracker.SetCashflowCompletedOn);
    }

    //Show warning if closed date is set and state is active
    if (formContext.getAttribute("caps_dateprojectclosed") != null) {
        formContext.getAttribute("caps_dateprojectclosed").addOnChange(CAPS.ProjectTracker.ShowHideClosedDateWarning);
        if (formContext.getAttribute("statuscode").getValue() == 0) {
            CAPS.ProjectTracker.ShowHideClosedDateWarning(executionContext);
        }
    }

    //Add warning to onChange of status if moving from a preApproval to Approval phase
    if (formContext.getAttribute("statuscode") != null) {
        CAPS.ProjectTracker.Status = formContext.getAttribute("statuscode").getValue();
        formContext.getAttribute("statuscode").addOnChange(CAPS.ProjectTracker.ShowStatusChangeWarning);
    }

    CAPS.ProjectTracker.ShowHideTabs(executionContext);
    formContext.getAttribute("caps_submissioncategory").addOnChange(CAPS.ProjectTracker.ShowHideTabs);

}

/**
Called on the SD form only, this function calls all other form functions that are SD form specific.
*/
CAPS.ProjectTracker.onLoadSD = function (executionContext) {
    var formContext = executionContext.getFormContext();

    //get project phase
    var projectStatus = formContext.getAttribute("statuscode").getValue();

    if (projectStatus === 1) {
        //PDR development
        //Hide COA, Funding Sources, Budget Breakdown, Change in Design Capacity, and Provincial Budget
        formContext.ui.tabs.get("tab_general").sections.get("sec_coa").setVisible(false);
        formContext.ui.tabs.get("tab_general").sections.get("sec_other_funding").setVisible(false);
        formContext.ui.tabs.get("tab_general").sections.get("sec_budget_breakdown").setVisible(false);
        formContext.ui.tabs.get("tab_general").sections.get("section_designcapacity").setVisible(false);
        formContext.ui.tabs.get("tab_general").sections.get("section_districtoperating").setVisible(false);
        formContext.ui.tabs.get("tab_general").sections.get("sec_provincial_budget").setVisible(false);
    }
}

/**
Shows either the facility lookup or the facility sub-grid depending on if multiple facilities are specified for the project.
*/
CAPS.ProjectTracker.ShowHideFacilities = function (executionContext) {
    debugger;
    var formContext = executionContext.getFormContext();
    formContext.getControl("caps_facilitysite").setVisible(false);

    if (formContext.getAttribute("caps_submissioncategorycode").getValue() == "LEASE") {
        formContext.getControl("sgd_facilities").setVisible(false);
        formContext.getControl("caps_facility").setVisible(false);
    }
    else if (formContext.getAttribute("caps_submissioncategorycode").getValue() == "NEW_SCHOOL") {
        formContext.getControl("sgd_facilities").setVisible(false);
        formContext.getControl("caps_proposedfacility").setVisible(true);
    }
    else if (formContext.getAttribute("caps_submissioncategorycode").getValue() == "SITE_ACQUISITION") {
        formContext.getControl("sgd_facilities").setVisible(false);
        formContext.getControl("caps_proposedsite").setVisible(true);
    }
    else if (formContext.getAttribute("caps_submissioncategorycode").getValue() == "DEMOLITION") {
        formContext.getControl("sgd_facilities").setVisible(false);
        if (formContext.getAttribute("caps_facility").getValue() == null) {
            formContext.getControl("caps_otherfacility").setVisible(true);
        }
    }
    else if (formContext.getAttribute("caps_submissioncategorycode").getValue() === "AFG") {
        formContext.getAttribute("caps_facility").controls.forEach(control => control.setVisible(false));
    }
    else if (formContext.getAttribute("caps_submissioncategorycode").getValue() === "Major_CC_New_Spaces" ||
        formContext.getAttribute("caps_submissioncategorycode").getValue() === "CC_MAJOR_NEW_SPACES_INTEGRATED" ||
        formContext.getAttribute("caps_submissioncategorycode").getValue() === "CC_CONVERSION" || 
        formContext.getAttribute("caps_submissioncategorycode").getValue() === "CC_CONVERSION_MINOR" || 
        formContext.getAttribute("caps_submissioncategorycode").getValue() === "CC_UPGRADE" ||
        formContext.getAttribute("caps_submissioncategorycode").getValue() === "CC_UPGRADE_MINOR") {
        formContext.getAttribute("caps_facility").controls.forEach(control => control.setVisible(true));
    }
    else {

        var id = formContext.data.entity.getId().replace("{", "").replace("}", "");

        var fetchXml = [
                "<fetch top='50'>",
                "  <entity name='caps_facility'>",
                "    <link-entity name='caps_projecttracker_caps_facility' from='caps_facilityid' to='caps_facilityid' intersect='true'>",
                "      <filter>",
                "        <condition attribute='caps_projecttrackerid' operator='eq' value='", id, "'/>",
                "      </filter>",
                "    </link-entity>",
                "  </entity>",
                "</fetch>",
        ].join("");


        Xrm.WebApi.retrieveMultipleRecords("caps_facility", "?fetchXml=" + fetchXml).then(
                function success(result) {
                    if (result.entities.length > 0) {
                        formContext.getControl("sgd_facilities").setVisible(true);
                        formContext.getControl("caps_facility").setVisible(false);

                    }
                    else {
                        formContext.getControl("sgd_facilities").setVisible(false);
                        formContext.getControl("caps_facility").setVisible(true);
                    }
                },
                function (error) {
                    alert(error.message);
                }
                );
    }


}

/***
This function is called on load of the project cash flow grid and calculates the total provincial cash flow and total agency projected values.
*/
CAPS.ProjectTracker.CalculateTotalCashFlowFields = function (executionContext) {
    debugger;
    var formContext = executionContext.getFormContext();
    var id = formContext.data.entity.getId().replace("{", "").replace("}", "");
    Xrm.WebApi.retrieveMultipleRecords("caps_projectcashflow", "?$select=caps_totalactualdraws,caps_totalprovincial,caps_agencyactual,caps_totalagency,caps_thirdpartyactual,caps_thirdpartyforecast,statecode&$filter=caps_PTR/caps_projecttrackerid eq " + id).then(
        function success(result) {
            debugger;
            //var totalActualDraws = 0;
            var totalProvincial = 0;
            var totalAgency = 0;
            var total3rdParty = 0;
            for (var i = 0; i < result.entities.length; i++) {
                if (result.entities[i].statecode == 0) {
                    totalProvincial += result.entities[i].caps_totalprovincial;  // Total Provincial has display name of Provincial Forecast
                    totalAgency += result.entities[i].caps_totalagency; // caps_agencyactual was Other Funding Sources Actuals previously and is now repurpose to SD Funding Actual
                    total3rdParty += result.entities[i].caps_thirdpartyforecast; // new field from CAPS-1964
                }
                else {
                    totalProvincial += result.entities[i].caps_totalactualdraws; // Actual Draws values come from different child entity.
                    totalAgency += result.entities[i].caps_agencyactual; // caps_totalagency was Other Funding Sources Forecast previously and is now repurpose to SD Funding Forecast
                    total3rdParty += result.entities[i].caps_thirdpartyactual; // new field from CAPS-1964
                }
            }

            // perform operations on record retrieval
            formContext.getAttribute('caps_totalagencyprojected').setValue(totalAgency);
            formContext.getAttribute('caps_totalprovincialcashflow').setValue(totalProvincial); 
            formContext.getAttribute('caps_totalthirdpartyprojected').setValue(total3rdParty); // New field from CAPS-1966

        },
        function (error) {
            console.log(error.message);
            // handle error conditions
        }
    );
}

/***
This function validates that the kindergarten design capacity is a multiple of the specified design capacity and shows a warning if it isn't.
***/
CAPS.ProjectTracker.ValidateKindergartenDesignCapacity = function (executionContext) {
    
    var formContext = executionContext.getFormContext();
    var designCapacity = formContext.getAttribute("caps_designcapacitykindergarten").getValue();
    var validateDesignCapacityKRequest = new CAPS.ProjectTracker.ValidateDesignCapacityRequest("Kindergarten", designCapacity);

    Xrm.WebApi.online.execute(validateDesignCapacityKRequest).then(
    function (result) {
        if (result.ok) {

            return result.json().then(
                function (response) {
                    if (!response.IsValid) {
                        formContext.ui.setFormNotification(response.ValidationMessage, 'ERROR', 'KINDERGARTEN DESIGN WARNING');
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
CAPS.ProjectTracker.ValidateElementaryDesignCapacity = function (executionContext) {
    
    var formContext = executionContext.getFormContext();
    var designCapacity = formContext.getAttribute("caps_designcapacityelementary").getValue();
    var validateDesignCapacityKRequest = new CAPS.ProjectTracker.ValidateDesignCapacityRequest("Elementary", designCapacity);

    Xrm.WebApi.online.execute(validateDesignCapacityKRequest).then(
    function (result) {
        if (result.ok) {

            return result.json().then(
                function (response) {
                    if (!response.IsValid) {
                        formContext.ui.setFormNotification(response.ValidationMessage, 'ERROR', 'ELEMENTARY DESIGN WARNING');
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
CAPS.ProjectTracker.ValidateSecondaryDesignCapacity = function (executionContext) {
    
    var formContext = executionContext.getFormContext();
    var designCapacity = formContext.getAttribute("caps_designcapacitysecondary").getValue();
    var validateDesignCapacityKRequest = new CAPS.ProjectTracker.ValidateDesignCapacityRequest("Secondary", designCapacity);

    Xrm.WebApi.online.execute(validateDesignCapacityKRequest).then(
    function (result) {
        if (result.ok) {

            return result.json().then(
                function (response) {
                    if (!response.IsValid) {
                        formContext.ui.setFormNotification(response.ValidationMessage, 'ERROR', 'SECONDARY DESIGN WARNING');
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

CAPS.ProjectTracker.ValidateDesignCapacityRequest = function (capacityType, capacityCount) {
    this.Type = capacityType;
    this.Count = capacityCount;
};

CAPS.ProjectTracker.ValidateDesignCapacityRequest.prototype.getMetadata = function () {
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
 * Function to embed the SSRS Monthly Summary Report on the form.
 * @param {any} formContext the form context
 */
CAPS.ProjectTracker.showSummaryReport = function (formContext) {    
    var requestUrl = "/api/data/v9.1/EntityDefinitions?$filter=LogicalName eq 'caps_projecttracker'&$select=ObjectTypeCode";

    var globalContext = Xrm.Utility.getGlobalContext();

    var req = new XMLHttpRequest();
    req.open("GET", globalContext.getClientUrl() + requestUrl, true);
    req.setRequestHeader("OData-MaxVersion", "4.0");
    req.setRequestHeader("OData-Version", "4.0");
    req.setRequestHeader("Accept", "application/json");
    req.setRequestHeader("Content-Type", "application/json; charset=utf-8");
    req.onreadystatechange = function () {
        if (this.readyState === 4) {
            req.onreadystatechange = null;
            if (this.status === 200) {
                var result = JSON.parse(this.response);
                var objectTypeCode = result.value[0].ObjectTypeCode;
                //use retrieved objectTypeCode
                //Get iframe 
                var iframeObject = formContext.getControl("IFRAME_SummaryReport");

                if (iframeObject !== null) {
                    var strURL = "/crmreports/viewer/viewer.aspx"
                        + "?id=6c075f27-49ba-ea11-a812-000d3af42496"
                        + "&action=run"
                        + "&context=records"
                        + "&recordstype="+objectTypeCode
                        + "&records=" + formContext.data.entity.getId()
                        + "&helpID=Project%20Summary.rdl";

                    //Set URL of iframe
                    iframeObject.setSrc(strURL);
                }
                
            } else {
                var errorText = this.responseText;
                //handle error here
                //display error
                Xrm.Navigation.openErrorDialog({ message: errorText });
            }
        }
    };
    req.send();
};

CAPS.ProjectTracker.DisableGridReadOnlyFields = function () {
    var gridContext = CAPS.ProjectTracker.gridContext;
    if (gridContext == null) {
        return;
    }
    var selectedRows = gridContext.getSelectedRows();
    if (selectedRows.getLength() === 0) {
        // do nothing
        return;
    }
    // Can't Edit when select multiple rows.
    // Pick the first Row.
    var rowContext = selectedRows.get(0).getData();
    rowContext.getEntity().attributes.forEach(function (attr) {
        if (attr.getName() === "caps_name" ||
            attr.getName() === "caps_fiscalyear" ||
            attr.getName() === "caps_thirdpartyactual" ||
            attr.getName() === "caps_thirdpartyforecast" ||
            attr.getName() === "caps_agencyactual" ||
            attr.getName() === "caps_totalagency" ||
            attr.getName() === "caps_sdprovincialforecast" ||
            attr.getName() === "caps_totalactualdraws") {
            attr.controls.forEach(function (c) {
                c.setDisabled(true);
            });
        }
    });
};


/**
 * Shows and hides appropriate sections and fields for Major/Minor and AFG Projects.
 * @param {any} formContext the form context
 */
CAPS.ProjectTracker.showHideCategoryRelevantSections = function (formContext) {
    debugger;
    var submissionCategoryCode = formContext.getAttribute("caps_submissioncategorycode").getValue();

    if (submissionCategoryCode === 'AFG') {
        //hide procurement, hide other funding, hide both date sections, hide agency section
        formContext.getControl("caps_procurementmethod").setVisible(false);
        if (formContext.getControl("caps_totalagencyprojected") !== null) {
            formContext.getControl("caps_totalagencyprojected").setVisible(false);
        }
        if (formContext.getControl("caps_variancefromagencybudgeted") !== null) {
            formContext.getControl("caps_variancefromagencybudgeted").setVisible(false);
        }

        formContext.ui.tabs.get("tab_general").sections.get("sec_other_funding").setVisible(false);
        formContext.ui.tabs.get("tab_general").sections.get("sec_agreement_dates").setVisible(false);
        formContext.ui.tabs.get("tab_general").sections.get("sec_ministry_dates").setVisible(false);
        formContext.ui.tabs.get("tab_general").sections.get("sec_budget_breakdown").setVisible(false);
        formContext.ui.tabs.get("tab_general").sections.get("sec_provincial_budget").setVisible(false);
       

        if (formContext.ui.tabs.get("tab_general").sections.get("sec_cps") !== null) {
            formContext.ui.tabs.get("tab_general").sections.get("sec_cps").setVisible(false);
        }

        formContext.ui.tabs.get("tab_general").sections.get("sec_afg_budget").setVisible(true);
        formContext.ui.tabs.get("tab_general").sections.get("section_designcapacity").setVisible(false);
        formContext.ui.tabs.get("tab_general").sections.get("section_districtoperating").setVisible(false);

        if (formContext.getControl("caps_projecttype") !== null) {
            formContext.getControl("caps_projecttype").setVisible(false);
        }

        if (formContext.getControl("caps_cpsnumber") !== null) {
            formContext.getControl("caps_cpsnumber").setVisible(false);
        }
        if (formContext.getControl("caps_requestedfunding") !== null) {
            formContext.getControl("caps_requestedfunding").setVisible(false);
        }

        if (formContext.getControl("caps_relatedproject") !== null) {
            formContext.getControl("caps_relatedproject").setVisible(false);
        }

        if (formContext.getControl("caps_sdprojectnumber") !== null) {
            formContext.getControl("caps_sdprojectnumber").setVisible(false);
        }


    }
    else if (submissionCategoryCode === 'BUS' || submissionCategoryCode === 'SEP' || submissionCategoryCode === 'PEP' || submissionCategoryCode === 'CNCP') {
        //hide other funding, provincial funding > hide all but total provincial field
        formContext.ui.tabs.get("tab_general").sections.get("sec_other_funding").setVisible(false);
        formContext.ui.tabs.get("tab_general").sections.get("sec_budget_breakdown").setVisible(false);


        //formContext.getControl("caps_approvedreserve").setVisible(false);
        //formContext.getControl("caps_totalapproved").setVisible(false);
        //formContext.getControl("caps_unapprovedreserve").setVisible(false);
        //formContext.getControl("caps_totalprovincialbudget").setVisible(false);

        if (formContext.getControl("caps_totalagencyprojected") !== null) {
            formContext.getControl("caps_totalagencyprojected").setVisible(false);
        }
        if (formContext.getControl("caps_variancefromagencybudgeted") !== null) {
            formContext.getControl("caps_variancefromagencybudgeted").setVisible(false);
        }

        formContext.ui.tabs.get("tab_general").sections.get("section_designcapacity").setVisible(false);
        formContext.ui.tabs.get("tab_general").sections.get("section_districtoperating").setVisible(false);

        if (submissionCategoryCode === 'BUS') {
            formContext.getControl("caps_projecttype").setVisible(false);

            if (formContext.getControl("caps_cpsnumber") !== null) {
                formContext.getControl("caps_cpsnumber").setVisible(false);
            }
            if (formContext.getControl("caps_requestedfunding") !== null) {
                formContext.getControl("caps_requestedfunding").setVisible(false);
            }

            if (formContext.getControl("caps_relatedproject") !== null) {
                formContext.getControl("caps_relatedproject").setVisible(false);
            }

            if (formContext.getControl("caps_procurementmethod") !== null) {
                formContext.getControl("caps_procurementmethod").setVisible(false);
            }
        }

    }
    else if (submissionCategoryCode === 'SITE_ACQUISITION') {
        formContext.ui.tabs.get("tab_general").sections.get("section_designcapacity").setVisible(false);
        formContext.ui.tabs.get("tab_general").sections.get("section_districtoperating").setVisible(false);
    }
    else if (submissionCategoryCode === 'LEASE') {
        formContext.getControl("caps_facility").setVisible(false);
        formContext.getControl("caps_submission").setVisible(false);
        formContext.getControl("caps_sdprojectnumber").setVisible(false);
        formContext.getControl("caps_requestedfunding").setVisible(false);

        if (formContext.ui.tabs.get("tab_general").sections.get("section_designcapacity") != null) {
            formContext.ui.tabs.get("tab_general").sections.get("section_designcapacity").setVisible(false);
        }

        if (formContext.ui.tabs.get("tab_general").sections.get("section_districtoperating") != null) {
            formContext.ui.tabs.get("tab_general").sections.get("section_districtoperating").setVisible(false);
        }

        if (formContext.ui.tabs.get("tab_general").sections.get("sec_ministry_dates") != null) {
            formContext.ui.tabs.get("tab_general").sections.get("sec_ministry_dates").setVisible(false);
        }
        if (formContext.ui.tabs.get("tab_general").sections.get("sec_agreement_dates") != null) {
            formContext.ui.tabs.get("tab_general").sections.get("sec_agreement_dates").setVisible(false);
        }
        if (formContext.ui.tabs.get("tab_emr") != null) {
            formContext.ui.tabs.get("tab_emr").setVisible(false);
        }
    }
         
    var submissionCategory = CAPS.ProjectTracker.GetLookup("caps_submissioncategory", formContext);
    if (submissionCategory !== undefined) {
        Xrm.WebApi.retrieveRecord("caps_submissioncategory", CAPS.ProjectTracker.RemoveCurlyBraces(submissionCategory.id), "?$select=caps_type,caps_categorycode").then(
            function success(result) {
                var type = result.caps_type;
                var categoryCode = result.caps_categorycode;
                //Minor or BEP or AFG or BUS
                if (type === 200870001 || type === 200870003 || type === 200870002 || type === 200870004) {
                    formContext.ui.tabs.get("tab_general").sections.get("sec_Initiatives").setVisible(false);
                }
                //Major
                if (type === 200870000) {
                    formContext.ui.tabs.get("tab_general").sections.get("sec_Initiatives").setVisible(true);
                    formContext.ui.tabs.get("tab_general").sections.get("tab_general_cps").setVisible(true);
                }
                if (type !== 200870000) {
                    formContext.ui.tabs.get("tab_general").sections.get("tab_general_cps").setVisible(false);
                }
                

                if (categoryCode === 'Major_CC_New_Spaces' || categoryCode === 'CC_MAJOR_NEW_SPACES_INTEGRATED') {
                    formContext.getAttribute("caps_childcarefacility").controls.forEach(control => control.setVisible(true));
                    formContext.getAttribute("caps_proposedchildcarefacility").controls.forEach(control => control.setVisible(true));
                    formContext.getAttribute("caps_proposedschoolfacility").controls.forEach(control => control.setVisible(true));
                    formContext.getAttribute("caps_childcareconstructiontype").controls.forEach(control => control.setVisible(true));
                    formContext.getAttribute("caps_healthauthority").controls.forEach(control => control.setVisible(true));
                }
                if (categoryCode === 'CC_CONVERSION') {
                    formContext.getAttribute("caps_childcarefacility").controls.forEach(control => control.setVisible(true));
                    formContext.getAttribute("caps_proposedchildcarefacility").controls.forEach(control => control.setVisible(true));
                    formContext.getAttribute("caps_proposedschoolfacility").controls.forEach(control => control.setVisible(false));
                    formContext.getAttribute("caps_childcareconstructiontype").controls.forEach(control => control.setVisible(true));
                    formContext.getAttribute("caps_healthauthority").controls.forEach(control => control.setVisible(true));
                }
                if (categoryCode === 'CC_CONVERSION_MINOR') {
                    formContext.getAttribute("caps_childcarefacility").controls.forEach(control => control.setVisible(true));
                    formContext.getAttribute("caps_proposedchildcarefacility").controls.forEach(control => control.setVisible(true));
                    formContext.getAttribute("caps_proposedschoolfacility").controls.forEach(control => control.setVisible(false));
                    formContext.getAttribute("caps_childcareconstructiontype").controls.forEach(control => control.setVisible(false));
                    formContext.getAttribute("caps_healthauthority").controls.forEach(control => control.setVisible(true));
                }
                if (categoryCode === 'CC_UPGRADE') {
                    formContext.getAttribute("caps_childcarefacility").controls.forEach(control => control.setVisible(true));
                    formContext.getAttribute("caps_proposedchildcarefacility").controls.forEach(control => control.setVisible(false));
                    formContext.getAttribute("caps_proposedschoolfacility").controls.forEach(control => control.setVisible(false));
                    formContext.getAttribute("caps_childcareconstructiontype").controls.forEach(control => control.setVisible(true));
                    formContext.getAttribute("caps_healthauthority").controls.forEach(control => control.setVisible(false));
                }
                if (categoryCode === 'CC_UPGRADE_MINOR') {
                    formContext.getAttribute("caps_childcarefacility").controls.forEach(control => control.setVisible(true));
                    formContext.getAttribute("caps_proposedchildcarefacility").controls.forEach(control => control.setVisible(false));
                    formContext.getAttribute("caps_proposedschoolfacility").controls.forEach(control => control.setVisible(false));
                    formContext.getAttribute("caps_childcareconstructiontype").controls.forEach(control => control.setVisible(false));
                    formContext.getAttribute("caps_healthauthority").controls.forEach(control => control.setVisible(false));
                }
              

            },
            function (error) {
                console.log(error.message);

            }
        );
    }
}

CAPS.ProjectTracker.ShowHideTabs = function (executionContext) {
    var formContext = executionContext.getFormContext();
    var submissionCategory = CAPS.ProjectTracker.GetLookup("caps_submissioncategory", formContext);
    if (submissionCategory !== undefined) {
        Xrm.WebApi.retrieveRecord("caps_submissioncategory", CAPS.ProjectTracker.RemoveCurlyBraces(submissionCategory.id), "?$select=caps_ischildcare").then(
            function success(result) {
                var isChildCare = result.caps_ischildcare;
                if (isChildCare) {
                    // Show CC tab, Draw Requests tab and hide Certificates of Approval tab when CC
                    formContext.ui.tabs.get("tab_child_care").setVisible(true);
                    formContext.ui.tabs.get("tab_DrawRequests").setVisible(true);
                    formContext.ui.tabs.get("tab_CertificatesofApproval").setVisible(false);
                }
                else {
                    // Hide CC tab, Draw Requests tab and show Certificates of Approval tab when K-12
                    formContext.ui.tabs.get("tab_child_care").setVisible(false);
                    formContext.ui.tabs.get("tab_DrawRequests").setVisible(false);
                    formContext.ui.tabs.get("tab_CertificatesofApproval").setVisible(true);
                }

            },
            function (error) {
                console.log(error.message);

            }
        );
    }
}

/**
Shows the progress report for major projects, otherwise hides the tab.
*/
CAPS.ProjectTracker.ShowHideProgressReport = function (executionContext) {
    var formContext = executionContext.getFormContext();

    if (formContext.getAttribute("caps_showprogressreports").getValue()) {
        formContext.ui.tabs.get("tab_progressreports").setVisible(true);
    }
    else {
        formContext.ui.tabs.get("tab_progressreports").setVisible(false);
    }
};

/**
Shows a warning if the Max Potential Budget and Max Potential don't match or if the Provincial Funding Source or Total Provincial Budget don't match.
*/
CAPS.ProjectTracker.ShowBudgetMissmatch = function (executionContext) {
    var formContext = executionContext.getFormContext();

    if (formContext.getAttribute('caps_showprogressreports').getValue()) {

        if (formContext.getAttribute("caps_maxpotential_fundingsource").getValue() !== formContext.getAttribute("caps_maxpotentialprojectbudget").getValue()) {
            formContext.ui.setFormNotification('Max Potential Project Budget and Max Potential Funding Source don\'t match.', 'ERROR', MAX_BUDGET_NOTIFICATION);
        }
        else {
            formContext.ui.clearFormNotification(MAX_BUDGET_NOTIFICATION);
        }

        if (formContext.getAttribute("caps_provincial").getValue() !== formContext.getAttribute("caps_totalprovincialbudget").getValue()) {
            formContext.ui.setFormNotification('Provincial Funding Source and Total Provincial Budget don\'t match.', 'ERROR', PROVINCIAL_BUDGET_NOTIFICATION);
        }
        else {
            formContext.ui.clearFormNotification(PROVINCIAL_BUDGET_NOTIFICATION);
        }
    }
};

/**
Shows a warning if any Milestones dates are in the past but are not flagged as complete.
**/
CAPS.ProjectTracker.ShowStaleDates = function (executionContext) {
    var formContext = executionContext.getFormContext();

    var recordId = formContext.data.entity.getId();

    var fetchXML = "<fetch version=\"1.0\" output-format=\"xml-platform\" mapping=\"logical\" distinct=\"false\">"+
              "<entity name=\"caps_projectmilestone\">"+
                "<attribute name=\"caps_projectmilestoneid\" />"+
                "<attribute name=\"caps_name\" />"+
                "<attribute name=\"createdon\" />"+
                "<order attribute=\"caps_name\" descending=\"false\" />"+
                "<filter type=\"and\">"+
                  "<condition attribute=\"caps_complete\" operator=\"eq\" value=\"0\" />"+
                  "<condition attribute=\"caps_expectedactualdate\" operator=\"olderthan-x-days\" value=\"1\" />" +
                  "<condition attribute=\"caps_projecttracker\" operator=\"eq\"  value=\""+recordId+"\" />"+
                "</filter>"+
              "</entity>"+
            "</fetch>";


    Xrm.WebApi.retrieveMultipleRecords("caps_projectmilestone", "?fetchXml=" + fetchXML).then(
                  function success(result) {
                      if (result.entities.length > 0) {
                          formContext.ui.setFormNotification('One or more project milestone dates requires updating.', 'ERROR', STALE_DATES_NOTIFICATION);
                    
                      }
                      else {
                          formContext.ui.clearFormNotification(STALE_DATES_NOTIFICATION);
                      }
                  },
                  function (error) {
                      Xrm.Navigation.openErrorDialog({ message: error.message });
                  }
                  );

};

/***
Show/Hide Project Closure Tab based on if caps_projectclosure has a value (show if it does).
*/
CAPS.ProjectTracker.ShowHideProjectClosureTab = function (executionContext) {
    debugger;
    var formContext = executionContext.getFormContext();

    var projectClosure = formContext.getAttribute("caps_projectclosure").getValue();

    if (projectClosure != null) {
        //show tab
        formContext.ui.tabs.get("tab_projectclosure").setVisible(true);
    }
    else {
        formContext.ui.tabs.get("tab_projectclosure").setVisible(false);
    }
}

/***
Set's caps_datecashflowupdated to the current date.  Called on button click.
*/
CAPS.ProjectTracker.SetCashflowCompletedOn = function (executionContext) {
    
    var formContext = executionContext.getFormContext();
    let attribute = executionContext.getEventSource();
    let value = attribute.getValue();
    if (value === "Cashflow Update Completed") {
        var currentDate = new Date();
        formContext.getAttribute("caps_datecashflowupdated").setValue(currentDate);
    }
    // Clear the value and avoid to submit data
    attribute.setValue(null);
    formContext.data.entity.save();
}

/*
Show's a warning if the complete/cancelled date field is populated but the project isn't closed or cancelled.
*/
CAPS.ProjectTracker.ShowHideClosedDateWarning = function (executionContext) {
    var formContext = executionContext.getFormContext();

    if (formContext.getAttribute("statecode").getValue() == 0 && formContext.getAttribute("caps_dateprojectclosed").getValue() !== null) {
        formContext.ui.setFormNotification('The project has a complete/cancelled date but is not completed or cancelled. Either clear the field or complete/cancel the project.', 'ERROR', CLOSED_DATE_NOTIFICATION);
    }
    else {
        formContext.ui.clearFormNotification(CLOSED_DATE_NOTIFICATION);
    }
}

/*Shows a warning, if changing from a pre-approval status to an approved status.*/
CAPS.ProjectTracker.ShowStatusChangeWarning = function (executionContext) {
    var formContext = executionContext.getFormContext();
    var newStatus = formContext.getAttribute("statuscode").getValue();
    

    //if the status was PDR Development or Approval
    if (CAPS.ProjectTracker.Status == 1 || CAPS.ProjectTracker.Status == 200870000) {
        //and the new status is Design Development, Construction or Substatially Complete, show a warning
        if (newStatus == 200870001 || newStatus == 200870002 || newStatus == 200870003) {

            var confirmStrings = { text: "Changing to an approved phase will make this Project visible to the School District.  If you later move it back to an unapproved phase it will not be hidden again. Click OK to continue or Cancel to remain with the previous phase. ", title: "Confirm Status Change" };
            var confirmOptions = { height: 200, width: 450 };
            Xrm.Navigation.openConfirmDialog(confirmStrings, confirmOptions).then(
                function (success) {
                    if (success.confirmed) {
                        //save the change
                        formContext.data.entity.save();
                        CAPS.ProjectTracker.Status = newStatus;
                    }
                    else {
                        //revert back
                       
                        formContext.getAttribute("statuscode").setValue(CAPS.ProjectTracker.Status);
                        formContext.data.entity.save();
                    }

                },
                function (error) {
                    Xrm.Navigation.openErrorDialog({ message: error });
                });
        }
    }
    else {
        //Update the status variable
        CAPS.ProjectTracker.Status = newStatus;
    }
}

CAPS.ProjectTracker.ValidateBudgetPressureAndSavings = function (executionContext) {
    var formContext = executionContext.getFormContext();
    var budgetPressureAttribute = formContext.getAttribute("caps_budgetpressure");
    var budgetSavingsAttribute = formContext.getAttribute("caps_budgetsavings");
    if (budgetPressureAttribute == null || budgetSavingsAttribute == null) {
        return; // Ignore if attribute not found.
    }
    var budgetPressureValue = budgetPressureAttribute.getValue();
    var budgetSavingsValue = budgetSavingsAttribute.getValue();

    if (budgetPressureValue == true && budgetSavingsValue == true) {
        var message = "Budget Pressure and Budget Savings can’t both be YES";
        var alertStrings = { confirmButtonLabel: "OK", text: message, title: "Budget Pressure & Budget Savings Validation" };
        var alertOptions = { height: 120, width: 260 };
        Xrm.Navigation.openAlertDialog(alertStrings, alertOptions);
    }
}

CAPS.ProjectTracker.GetLookup = function (fieldName, formContext) {
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
CAPS.ProjectTracker.RemoveCurlyBraces = function (str) {
    return str.replace(/[{}]/g, "");
};