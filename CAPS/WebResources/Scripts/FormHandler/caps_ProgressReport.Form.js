"use strict";

var CAPS = CAPS || {};
CAPS.ProgressReport = CAPS.ProgressReport || {};

/*
Main function for Project Tracker.  This function calls all other form functions.
*/
CAPS.ProgressReport.onLoad = function (executionContext) {
    var formContext = executionContext.getFormContext();
    var formType = formContext.ui.getFormType();
    if (formType == 2) { // Update Mode
        formContext.getControl("sgd_FutureCashFlow").addOnLoad(CAPS.ProgressReport.UpdateTotalFutureCashFlow);
        CAPS.ProgressReport.UpdateProjectBudgetValues(executionContext);
    }
    else if (formType == 1) // Create Mode
    {
        CAPS.ProgressReport.SetUseNewFormOnCreate(executionContext);
    }


    CAPS.ProgressReport.SwitchForm(executionContext);
    //embed report
    //CAPS.ProgressReport.ShowMonthlyReport(formContext);
};

/*
Sums up the total projected provincial forecast and projected actuals.
*/
CAPS.ProgressReport.UpdateTotalFutureCashFlow = function (executionContext) {
    var formContext = executionContext.getFormContext();
    var progressReportYearAttribute = formContext.getAttribute("caps_fiscalyear");
    var progressReportYearId = progressReportYearAttribute.getValue()[0].id.replace("{", "").replace("}", "");
    var id = formContext.data.entity.getId().replace("{", "").replace("}", "");
    Xrm.WebApi.retrieveRecord("edu_year", progressReportYearId, "?$select=edu_startyear").then(
        function SuccessRetrievingYear(data) {
            // Provincial Amount
            // Agency Amount is changed to SD Funding Sources Amount per CAPS-1958
            // Schema Name did not change.
            var progressReportStartYear = data.edu_startyear;
            var optionsOld = "?$select=caps_provincialamount,caps_sdfundingsourcesforecast,caps_thirdpartyforecast&$filter=caps_ProgressReport/caps_progressreportid eq " + id;
            var options = "?$select=caps_provincialamount,caps_sdprovincialforecast,caps_sdfundingsourcesforecast,caps_agencyamount,caps_thirdpartyforecast,caps_thirdpartyamount&$expand=caps_Year($select=edu_startyear)&$filter=_caps_progressreport_value eq " + id;
            Xrm.WebApi.retrieveMultipleRecords("caps_cashflowprojection", options).then(

                function success(result) {
                    var totalProvincial = 0;
                    var totalAgency = 0;
                    var total3rdParty = 0;
                    var startYear = 0;
                    var actualProvincial = 0;
                    var actual3rdParty = 0;
                    var actualSDSources = 0;
                    for (var i = 0; i < result.entities.length; i++) {
                        startYear = result.entities[i].caps_Year.edu_startyear;
                        if (startYear >= progressReportStartYear) {
                            // Existing Math has fiscal in the past calculated with a pre-populated amount.  
                            // We will only calculate the current and future fiscal years.
                            if (result.entities[i].caps_sdprovincialforecast != null)
                                totalProvincial += result.entities[i].caps_sdprovincialforecast;
                            if (result.entities[i].caps_sdfundingsourcesforecast != null)
                                totalAgency += result.entities[i].caps_sdfundingsourcesforecast;
                            if (result.entities[i].caps_thirdpartyforecast != null)
                                total3rdParty += result.entities[i].caps_thirdpartyforecast;
                        }
                        else {
                            if (result.entities[i].caps_provincialamount != null)
                                actualProvincial += result.entities[i].caps_provincialamount;
                            if (result.entities[i].caps_agencyamount != null)
                                actualSDSources += result.entities[i].caps_agencyamount;
                            if (result.entities[i].caps_thirdpartyamount != null)
                                actual3rdParty += result.entities[i].caps_thirdpartyamount;
                        }
                    }

                    // perform operations on record retrieval
                    formContext.getAttribute('caps_totalfutureprovincialprojection').setValue(totalProvincial);
                    formContext.getAttribute('caps_totalfutureagencyprojection').setValue(totalAgency);
                    formContext.getAttribute('caps_totalfuturethirdpartyprojection').setValue(total3rdParty);

                    formContext.getAttribute('caps_totalprovincialactualdraws').setValue(actualProvincial);
                    formContext.getAttribute('caps_totalthirdpartyactuals').setValue(actual3rdParty);
                    formContext.getAttribute('caps_totalagencyactuals').setValue(actualSDSources);
                },
                function (error) {
                    console.log(error.message);
                    // handle error conditions
                }
            );
        }
    );
};

/**
 * Function to embed the SSRS Monthly Progress Report on the form.
 * @param {any} formContext the form context
 */
// CAPS-1948 Remove Monthly Report from Form
/*
CAPS.ProgressReport.ShowMonthlyReport = function (formContext) {
    debugger;
    var requestUrl = "/api/data/v9.1/EntityDefinitions?$filter=LogicalName eq 'caps_progressreport'&$select=ObjectTypeCode";

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
                debugger;
                var result = JSON.parse(this.response);
                var objectTypeCode = result.value[0].ObjectTypeCode;
                //use retrieved objectTypeCode
                //Get iframe 
                var iframeObject = formContext.getControl("IFRAME_SummaryReport");

                if (iframeObject !== null) {
                    var strURL = "/crmreports/viewer/viewer.aspx"
                        + "?id=ef515e3b-3c2f-eb11-a813-000d3af43595"
                        + "&action=run"
                        + "&context=records"
                        + "&recordstype=" + objectTypeCode
                        + "&records=" + formContext.data.entity.getId()
                        + "&helpID=SD%20Project%20Progress%20Report.rdl";

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
*/

CAPS.ProgressReport.ValidateStatusComment = function (executionContext, formContext1) {
    var formContext = formContext1; // Take Form Context from Ribbon Control
    if (executionContext != null) {
        formContext = executionContext.getFormContext();
    }
    var statusCommentAttribute = formContext.getAttribute("caps_statuscomments");
    var confirmNoChangesAttribute = formContext.getAttribute("caps_confirmnochanges");
    if (statusCommentAttribute != null && confirmNoChangesAttribute != null) {
        var confirmNoChangeValue = confirmNoChangesAttribute.getValue();
        var statusCommentsValue = statusCommentAttribute.getValue();
        if (confirmNoChangeValue && (statusCommentsValue == null || statusCommentsValue.trim() == "")) {
            statusCommentAttribute.setRequiredLevel("required");
            return false;
        }
        else {
            statusCommentAttribute.setRequiredLevel("none");
            return true;
        }
    }
}

// 2 newly added 2-option for Confirmed Access for Cashflow and for Project Milestones
// Called by CAPS.ProgressReport.ValidateOccupancyDate on Submit Button click
CAPS.ProgressReport.ConfirmReviewedBeforeSubmission = function (formContext, message) {
    
    var needUpdate = false;
    var isConfimrMilestoneValid = false;
    var isConfirmCashflowValid = false;
    var isConfirmStatusCommentsValid = CAPS.ProgressReport.ValidateStatusComment(null, formContext);
    var confirmMessage = "Please ensure Confirm Review of required sections";
    var commentMessage = "You have indicated changes have been made to this report, please provide an explanation of these changes in the comment box in order to submit. If no changes have been made please set the Changes Made toggle to No.";
    var confirmReviewMessage = ""
    var confirmProjectMilestoneReviewAttribute = formContext.getAttribute("caps_confirmprojectmilestonereview");
    if (confirmProjectMilestoneReviewAttribute != null) {
        if (confirmProjectMilestoneReviewAttribute.getValue() == null) {
            confirmProjectMilestoneReviewAttribute.setValue(false);
            needUpdate = true;
        }
        else {
            isConfimrMilestoneValid = confirmProjectMilestoneReviewAttribute.getValue();
        }
    }

    var confirmCashflowReviewAttribute = formContext.getAttribute("caps_confirmcashflowreview");
    if (confirmCashflowReviewAttribute != null) {
        if (confirmCashflowReviewAttribute.getValue() == null) {
            confirmCashflowReviewAttribute.setValue(false);
            needUpdate = true;
        }
        else {
            isConfirmCashflowValid = confirmCashflowReviewAttribute.getValue();
        }
    }

    if (isConfimrMilestoneValid && isConfirmCashflowValid && isConfirmStatusCommentsValid) {
        CAPS.ProgressReport.ConfirmSubmit(formContext, message);
    }
    else {
        if (!isConfimrMilestoneValid || !isConfirmCashflowValid) {
            confirmReviewMessage += confirmMessage;
        }
        if (confirmReviewMessage != "" && !isConfirmStatusCommentsValid) {
            confirmReviewMessage += " and ";
        }
        if (!isConfirmStatusCommentsValid) {
            confirmReviewMessage += commentMessage;
        }
        confirmReviewMessage = confirmReviewMessage;
        var alertStrings = {
            confirmButtonLabel: "OK", text: confirmReviewMessage, title: "Confirmed Review before Submission"
        };
        var alertOptions = { height: 120, width: 260 };
        Xrm.Navigation.openAlertDialog(alertStrings, alertOptions);
        if (needUpdate) {
            formContext.data.entity.save(); // Save the record
        }
    }
}

// Technically, Project can have updated values during the month.
// Total Provincial Budget (caps_totalprovincialbudget) = Provincial (caps_provincial)
// Total SD Funding Budget (caps_totalagencybudget) = Restricted Capital (caps_restrictedcapital) + Land Capital (caps_landcapital) + AFG (caps_afg) + SD Other (caps_sdother)
// Total 3rd Party Budget (caps_totalthirdpartybudget) = Federal (caps_federal) + Other 3rd Party (caps_thirdparty)
CAPS.ProgressReport.UpdateProjectBudgetValues = function (executionContext, formContext1) {
    var formContext = formContext1; // Take Form Context from Ribbon Control
    if (executionContext != null) {
        formContext = executionContext.getFormContext();
    }
    var projectAttributeName = "caps_project";
    var projectAttribute = formContext.getAttribute(projectAttributeName);
    if (projectAttribute != null) {
        var projectValue = projectAttribute.getValue();
        if (projectValue != null) {
            var projectId = projectValue[0].id.replace("{", "").replace("}", "");
            var options = "?$select=caps_provincial,caps_restrictedcapital,caps_landcapital,caps_afg,caps_sdother,caps_federal,caps_thirdparty";
            Xrm.WebApi.retrieveRecord("caps_projecttracker", projectId, options).then(
                function Success(data) {
                    var updateData = false;

                    // Newly Added field Total 3rd Party Actuals is based on values submitted prior to current month.  
                    // However, this value may be NULL for older records.  This is set by cloud flows on record creation
                    // This is used in Calculated fields, and having 0 results in NULL in those calculated fields too.
                    var total3rdPartyActualsAttribute = formContext.getAttribute("caps_totalthirdpartyactuals");
                    if (total3rdPartyActualsAttribute != null && total3rdPartyActualsAttribute.getValue() == null) {
                        total3rdPartyActualsAttribute.setValue(0);
                        updateData = true;
                    }

                    var federal = data.caps_federal;
                    var other3rdParty = data.caps_thirdparty;
                    var total3rdPartyAttribute = formContext.getAttribute("caps_totalthirdpartybudget");
                    if (total3rdPartyAttribute != null &&
                        (total3rdPartyAttribute.getValue() == null || total3rdPartyAttribute.getValue() != (federal + other3rdParty))) {
                        total3rdPartyAttribute.setValue(federal + other3rdParty);
                        updateData = true;
                    }

                    var restrictedCapital = data.caps_restrictedcapital;
                    var landCapital = data.caps_landcapital;
                    var afg = data.caps_afg;
                    var sdOther = data.caps_sdother;
                    var totalSDFundingBudgetAttribute = formContext.getAttribute("caps_totalagencybudget");
                    if (totalSDFundingBudgetAttribute != null &&
                        (totalSDFundingBudgetAttribute.getValue() == null ||
                            totalSDFundingBudgetAttribute.getValue() != (restrictedCapital + landCapital + afg + sdOther))) {
                        totalSDFundingBudgetAttribute.setValue(restrictedCapital + landCapital + afg + sdOther);
                        updateData = true;
                    }

                    var provincial = data.caps_provincial;
                    var totalProvincialBudgetAttribute = formContext.getAttribute("caps_totalprovincialbudget");
                    if (totalProvincialBudgetAttribute != null &&
                        (totalProvincialBudgetAttribute.getValue() == null || totalProvincialBudgetAttribute.getValue() != provincial)) {
                        totalProvincialBudgetAttribute.setValue(provincial);
                        updateData = true;
                    }

                    if (updateData) {
                        formContext.data.entity.save(); // Save the record
                    }
                },
                function Error(err) {
                    var debug = err;
                }
            );
        }
    }
};

CAPS.ProgressReport.SetUseNewFormOnCreate = function (executionContext) {
    var formContext = executionContext.getFormContext();
    var useNewFormAttribute = formContext.getAttribute("caps_usenewform");
    if (useNewFormAttribute == null) {
        return; // Ignore if the attribute cannot be found.
    }
    useNewFormAttribute.setValue(true);
};

CAPS.ProgressReport.SwitchForm = function (executionContext) {
    var formContext = executionContext.getFormContext();
    var useNewFormAttribute = formContext.getAttribute("caps_usenewform");
    if (useNewFormAttribute == null) {
        return; // Ignore if the attribute cannot be found.
    }
    var useNewForm = useNewFormAttribute.getValue();
    var newFormName = "Information Form";
    var legacyFormName = "Legacy Form";
    var formName = "";
    if (useNewForm == true) {
        formName = newFormName;
    }
    else {
        formName = legacyFormName;
    }
    var clientContext = Xrm.Utility.getGlobalContext().client;
    try {
        // Mobile App only have 1 form.
        if (clientContext.getClient() == "Mobile") {
            return;
        }
        if (formContext.ui.formSelector.getCurrentItem().getLabel() !== formName) {
            var items = formContext.ui.formSelector.items;
            items.forEach(function (item, index) {
                var itemLabel = item.getLabel();
                if (itemLabel === formName) {
                    //navigate to the form
                    item.navigate();
                } //endif
            });

        } //endif
    }
    catch (ex) {
        // Do Nothing
    }
};

CAPS.ProgressReport.ValidateOccupancyDate = function (executionContext, formContext1, confrimSubmit) {
    var formContext = formContext1; // Take Form Context from Ribbon Control
    if (executionContext != null) {
        formContext = executionContext.getFormContext();
    }
    var occupancyDateAttributeName = "caps_projectedoccupancydate";
    var occupancyAttribute = formContext.getAttribute(occupancyDateAttributeName);
    if (occupancyAttribute == null) {
        return true;; // Ignore if attribute does not exist
    }
    // Cannot be in the past if contains data
    var occupancyDateValue = occupancyAttribute.getValue();
    if (occupancyDateValue == null) {
        if (confrimSubmit) {
            var message = "Occupancy Date does not contains value";
            CAPS.ProgressReport.ConfirmReviewedBeforeSubmission(formContext, message);
        }
        return true; // Ignore if it does not contains data
    }

    var occupancyDate = new Date(occupancyDateValue);

    // Set time to midnight
    occupancyDate.setHours(0, 0, 0, 0);
    // Get today's date at midnight
    var today = new Date();
    today.setHours(0, 0, 0, 0);
    // Check Project's Occupancy Date.
    // Check if Occupancy Date available and Is it already set as Complete.
    // If it is, does it equal to the Expected/Actual Date from milestone
    var projectAttributeName = "caps_project";
    var projectAttribute = formContext.getAttribute(projectAttributeName);
    if (projectAttribute != null) {
        var projectValue = projectAttribute.getValue();
        if (projectValue != null) {
            var projectId = projectValue[0].id.replace("{", "").replace("}", "");
            var occupancyMilestoneId = "1c455023-47a0-ea11-a812-000d3af42496";
            var notified = true;
            var options = "?$filter=_caps_milestone_value eq " + occupancyMilestoneId + " and _caps_projecttracker_value eq " + projectId + " and caps_complete eq true";
            Xrm.WebApi.retrieveMultipleRecords("caps_projectmilestone", options).then(
                function Success(data) {
                    // Do something
                    var results = data.entities;
                    if (results.length > 0) {
                        var expectedDateString = results[0].caps_expectedactualdate;
                        if (expectedDateString != null) {
                            var expectedDate = new Date(expectedDateString);
                            // This expected Date instantiation will get timezone offset.
                            expectedDate = new Date(expectedDate.getTime() + expectedDate.getTimezoneOffset() * 60000);
                            expectedDate.setHours(0, 0, 0, 0);
                            var doNotMatchYear = occupancyDate.getFullYear() != expectedDate.getFullYear();
                            var doNotMatchMonth = occupancyDate.getMonth() != expectedDate.getMonth();
                            var doNotMatchDay = occupancyDate.getDay() != expectedDate.getDay();
                            if (doNotMatchYear || doNotMatchMonth || doNotMatchDay) {
                                // Prompt
                                var message = "Occupancy Date is not matching Completed Occupancy Milestone's Expected/Actual Date - " + expectedDateString;
                                if (confrimSubmit) {
                                    CAPS.ProgressReport.ConfirmReviewedBeforeSubmission(formContext, message);
                                }
                                else {
                                    var alertStrings = { confirmButtonLabel: "OK", text: message, title: "Projected Occupancy Date Validation" };
                                    var alertOptions = { height: 120, width: 260 };
                                    Xrm.Navigation.openAlertDialog(alertStrings, alertOptions);
                                }
                            }
                            else {
                                var occupancyCompleteAttribute = formContext.getAttribute("caps_occupancydatecomplete");
                                // If it hasn't been locked already, lock the field if matching the Completed Occupancy Milestone
                                if (occupancyCompleteAttribute != null && occupancyCompleteAttribute.getValue() == false) {
                                    occupancyCompleteAttribute.setValue(true); 
                                }
                                if (confrimSubmit) {
                                    CAPS.ProgressReport.ConfirmReviewedBeforeSubmission(formContext, message);
                                }
                            }
                        }
                        else {
                            // If it were in the past
                            if (occupancyDate < today) {
                                // Prompt
                                var message = "Occupancy Date should not be in the past";
                                if (confrimSubmit) {
                                    CAPS.ProgressReport.ConfirmReviewedBeforeSubmission(formContext, message);
                                }
                                else {
                                    var alertStrings = { confirmButtonLabel: "OK", text: message, title: "Projected Occupancy Date Validation" };
                                    var alertOptions = { height: 120, width: 260 };
                                    Xrm.Navigation.openAlertDialog(alertStrings, alertOptions);
                                }
                            }
                            else {

                                if (confrimSubmit) {
                                    CAPS.ProgressReport.ConfirmReviewedBeforeSubmission(formContext, "");
                                }
                            }
                        }
                    }
                    else {
                        // If it were in the past
                        if (occupancyDate < today) {
                            // Prompt
                            var message = "Occupancy Date should not be in the past";
                            if (confrimSubmit) {
                                CAPS.ProgressReport.ConfirmReviewedBeforeSubmission(formContext, message);
                            }
                            else {
                                var alertStrings = { confirmButtonLabel: "OK", text: message, title: "Projected Occupancy Date Validation" };
                                var alertOptions = { height: 120, width: 260 };
                                Xrm.Navigation.openAlertDialog(alertStrings, alertOptions);
                            }
                        }
                        else {

                            if (confrimSubmit) {
                                CAPS.ProgressReport.ConfirmReviewedBeforeSubmission(formContext, "");
                            }
                        }
                    }
                    return notified;
                },
                function Error(err) {
                    var debug = err;
                }
            );
        }
    }

    
    // Parent Project whether it is Completed
};

