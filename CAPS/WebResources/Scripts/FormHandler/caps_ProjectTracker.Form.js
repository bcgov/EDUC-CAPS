"use strict";

var CAPS = CAPS || {};
CAPS.ProjectTracker = CAPS.ProjectTracker || {};

const MAX_BUDGET_NOTIFICATION = "Max_Budget_Notification";
const PROVINCIAL_BUDGET_NOTIFICATION = "Provincial_Budget_Notification";
const STALE_DATES_NOTIFICATION = "Stale_Dates_Notification";

/**
 * Main function for Project Tracker.  This function calls all other form functions.
 * @param {any} executionContext the form's execution context.
 */
CAPS.ProjectTracker.onLoad = function (executionContext) {
    var formContext = executionContext.getFormContext();

    //Show Summary Report
    CAPS.ProjectTracker.showSummaryReport(formContext);

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
        formContext.ui.tabs.get("tab_general").sections.get("sec_provincial_budget").setVisible(false);
    }
}

/**
Shows either the facility lookup or the facility sub-grid depending on if multiple facilities are specified for the project.
*/
CAPS.ProjectTracker.ShowHideFacilities = function (executionContext) {
    debugger;
    var formContext = executionContext.getFormContext();

    if (formContext.getAttribute("caps_submissioncategorycode").getValue() == "NEW_SCHOOL") {
        formContext.getControl("caps_facilitysite").setVisible(true);
    }
    else {
        formContext.getControl("caps_facilitysite").setVisible(false);
    }

    formContext.getControl("sgd_facilities").addOnLoad(function (gridContext) {
        debugger;
        var filteredRecordCount = gridContext.getFormContext().getControl("sgd_facilities").getGrid().getTotalRecordCount();

        if (filteredRecordCount > 0) {
            formContext.getControl("sgd_facilities").setVisible(true);
            formContext.getControl("caps_facility").setVisible(false);
        }
        else {
            formContext.getControl("sgd_facilities").setVisible(false);
            formContext.getControl("caps_facility").setVisible(true);
        }
    });


}

/***
This function is called on load of the project cash flow grid and calculates the total provincial cash flow and total agency projected values.
*/
CAPS.ProjectTracker.CalculateTotalCashFlowFields = function (executionContext) {
    debugger;
    var formContext = executionContext.getFormContext();
    var id = formContext.data.entity.getId().replace("{", "").replace("}", "");
    Xrm.WebApi.retrieveMultipleRecords("caps_projectcashflow", "?$select=caps_totalactualdraws,caps_totalprovincial,caps_totalagency,statecode&$filter=caps_PTR/caps_projecttrackerid eq " + id).then(
        function success(result) {
            debugger;
            //var totalActualDraws = 0;
            var totalProvincial = 0;
            var totalAgency = 0;
            for (var i = 0; i < result.entities.length; i++) {
                if (result.entities[i].statecode == 0) {
                    totalProvincial += result.entities[i].caps_totalprovincial;
                }
                else {
                    totalProvincial += result.entities[i].caps_totalactualdraws;
                }
                totalAgency += result.entities[i].caps_totalagency;
            }

            // perform operations on record retrieval
            formContext.getAttribute('caps_totalagencyprojected').setValue(totalAgency);
            formContext.getAttribute('caps_totalprovincialcashflow').setValue(totalProvincial);

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
                        formContext.ui.setFormNotification(response.ValidationMessage, 'WARNING', 'KINDERGARTEN DESIGN WARNING');
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
                        formContext.ui.setFormNotification(response.ValidationMessage, 'WARNING', 'ELEMENTARY DESIGN WARNING');
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
                        formContext.ui.setFormNotification(response.ValidationMessage, 'WARNING', 'SECONDARY DESIGN WARNING');
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
                        + "?id=ed1744bc-98a6-ea11-a813-000d3af42496"
                        + "&action=run"
                        + "&context=records"
                        + "&recordstype="+objectTypeCode
                        + "&records=" + formContext.data.entity.getId()
                        + "&helpID=Monthly%20Project%20Summary.rdl";

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


/**
 * Shows and hides appropriate sections and fields for Major/Minor and AFG Projects.
 * @param {any} formContext the form context
 */
CAPS.ProjectTracker.showHideCategoryRelevantSections = function (formContext) {
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

    }
    else if (submissionCategoryCode === 'BUS' || submissionCategoryCode === 'SEP' || submissionCategoryCode === 'PEP' || submissionCategoryCode === 'CNCP') {
        //hide other funding, provincial funding > hide all but total provincial field
        formContext.ui.tabs.get("tab_general").sections.get("sec_other_funding").setVisible(false);
        formContext.ui.tabs.get("tab_general").sections.get("sec_budget_breakdown").setVisible(false);


        formContext.getControl("caps_approvedreserve").setVisible(false);
        formContext.getControl("caps_totalapproved").setVisible(false);
        formContext.getControl("caps_unapprovedreserve").setVisible(false);
        formContext.getControl("caps_totalprovincialbudget").setVisible(false);
        
        if (formContext.getControl("caps_totalagencyprojected") !== null) {
            formContext.getControl("caps_totalagencyprojected").setVisible(false);
        }
        if (formContext.getControl("caps_variancefromagencybudgeted") !== null) {
            formContext.getControl("caps_variancefromagencybudgeted").setVisible(false);
        }

        formContext.ui.tabs.get("tab_general").sections.get("section_designcapacity").setVisible(false);

    }
    else if (submissionCategoryCode === 'SITE_ACQUISITION') {
        formContext.ui.tabs.get("tab_general").sections.get("section_designcapacity").setVisible(false);
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

    if (formContext.getAttribute("caps_maxpotential_fundingsource").getValue() !== formContext.getAttribute("caps_maxpotentialprojectbudget").getValue()) {
        formContext.ui.setFormNotification('Max Potential Project Budget and Max Potential Funding Source don\'t match.', 'WARNING', MAX_BUDGET_NOTIFICATION);
    }
    else {
        formContext.ui.clearFormNotification(MAX_BUDGET_NOTIFICATION);
    }

    if (formContext.getAttribute("caps_provincial").getValue() !== formContext.getAttribute("caps_totalprovincialbudget").getValue()) {
        formContext.ui.setFormNotification('Provincial Funding Source and Total Provincial Budget don\'t match.', 'WARNING', PROVINCIAL_BUDGET_NOTIFICATION);
    }
    else {
        formContext.ui.clearFormNotification(PROVINCIAL_BUDGET_NOTIFICATION);
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
                          formContext.ui.setFormNotification('One or more project milestone dates requires updating.', 'WARNING', STALE_DATES_NOTIFICATION);
                    
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

