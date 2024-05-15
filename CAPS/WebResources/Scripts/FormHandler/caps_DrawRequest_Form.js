"use strict";

var CAPS = CAPS || {};
CAPS.DrawRequest = CAPS.DrawRequest || {};

//This JavaScript is to sum all draw request amounts except cancelled, and validate the amount that can be drawn on "amount" changes at draw request. 

CAPS.DrawRequest.form_onload = function (executionContext) {
    var formContext = executionContext.getFormContext();

    var amountField = formContext.getAttribute("caps_amount");
    var projectLookup = formContext.getAttribute("caps_project");
    var projectCode = formContext.getAttribute("caps_projectcode");
    var fiscalYear = formContext.getAttribute("caps_fiscalyear");
    var drawDate = formContext.getAttribute("caps_drawdate");
    var processDate = formContext.getAttribute("caps_processdate");
    var batchLookup = formContext.getAttribute("caps_batch");

    batchNotification(formContext);
    fetchBatchDetails(formContext);

    batchLookup.addOnChange(function () {
        batchNotification(formContext);
        fetchBatchDetails(formContext);
    });

    if (projectLookup) {
        projectLookup.addOnChange(CAPS.DrawRequest.updateRemainingBalance);
    }

    if (amountField) {
        amountField.addOnChange(function () {
            CAPS.DrawRequest.amountValidation(executionContext);
            CAPS.DrawRequest.updateRemainingBalance(executionContext);
            formContext.ui.refreshRibbon();
        });
    }

    CAPS.DrawRequest.updateRemainingBalance(executionContext);

    // Check all required fields before saving.
    if (projectLookup.getValue() && projectCode.getValue() && fiscalYear.getValue() &&
        drawDate.getValue() && processDate.getValue() && amountField.getValue()) {
        formContext.data.save();
    }

    fetchBatchDetails(formContext);

    function fetchBatchDetails(formContext) {

        function parseDate(dateStr) {
            var date = new Date(dateStr);
            return new Date(date.getTime() + date.getTimezoneOffset() * 60000);
        }

        var batchValue = formContext.getAttribute("caps_batch").getValue();
        if (!batchValue || batchValue.length === 0) {
            drawDate.setValue(null);
            processDate.setValue(null);
            fiscalYear.setValue(null);
            return;
        }

        var fetchXml = "<fetch top='1'>" +
            "<entity name='caps_batch'>" +
            "<attribute name='caps_drawdate' />" +
            "<attribute name='caps_submissiondeadline' />" +
            "<attribute name='caps_fiscalyear' />" +
            "<link-entity name='edu_year' from='edu_yearid' to='caps_fiscalyear' alias='fiscal'>" +
            "<attribute name='edu_yearid' />" +
            "<attribute name='edu_name' />" +
            "</link-entity>" +
            "<filter type='and'>" +
            "<condition attribute='caps_batchid' operator='eq' value='" + batchValue[0].id + "' />" +
            "</filter>" +
            "</entity>" +
            "</fetch>";

        var encodedFetchXml = encodeURIComponent(fetchXml);
        Xrm.WebApi.retrieveMultipleRecords("caps_batch", "?fetchXml=" + encodedFetchXml).then(
            function (result) {
                if (result.entities.length > 0) {
                    var batchRecord = result.entities[0];

                    var drawDateValue = batchRecord['caps_drawdate'] ? parseDate(batchRecord['caps_drawdate']) : null;
                    var processDateValue = batchRecord['caps_submissiondeadline'] ? parseDate(batchRecord['caps_submissiondeadline']) : null;
                    drawDate.setValue(drawDateValue);
                    processDate.setValue(processDateValue);

                    if (batchRecord['fiscal.edu_yearid']) {
                        fiscalYear.setValue([{
                            id: batchRecord['fiscal.edu_yearid'],
                            name: batchRecord['fiscal.edu_name'],
                            entityType: "edu_year"
                        }]);
                    } else {
                        fiscalYear.setValue(null);
                    }
                } else {
                    drawDate.setValue(null);
                    processDate.setValue(null);
                    fiscalYear.setValue(null);
                }
            },
            function (error) {
                console.error("Error retrieving batch details with FetchXML: " + error.message);
            }
        );
    }
};

//Display notification that a batch is required to proceed
function batchNotification(formContext) {
    var batch = formContext.getAttribute("caps_batch");
    var statecode = formContext.getAttribute("statecode").getValue();
    if (!batch || !batch.getValue() && statecode === 0) {
        formContext.ui.setFormNotification("Add the request to a batch to be able to mark it as ready to submit", "WARNING", "missingBatchNotification");
    } else {
        formContext.ui.clearFormNotification("missingBatchNotification");
    }
}

CAPS.DrawRequest.isFormDirty = function (primaryControl) {
    var formContext = primaryControl;
    // Use the getIsDirty method to check if the form has unsaved changes. This preents user from running the workflow through the button
    var isDirty = formContext.data.entity.getIsDirty();
    var showButton = false;

    if (!isDirty) {
        showButton = true;
    } else if (isDirty || formType === 1) {
        showButton = false;
    }

    return showButton;
};

//IF CREATE form and the total requested of the related draw requests to the same project === 0 then remaining draw reuest balance = the project's remaining draw request balance.
//ELSE IF existing record and total requested so far <> 0, then set remaining draw request balance as Total Approved of Project - Total Requested of related draw requets.
CAPS.DrawRequest.updateRemainingBalance = function (executionContext) {
    var formContext = executionContext.getFormContext();
    var projectLookup = formContext.getAttribute("caps_project").getValue();
    var currentDrawRequestId = formContext.data.entity.getId();
    var amount = formContext.getAttribute("caps_amount").getValue() || 0;
    var amountControl = formContext.getControl("caps_amount");
    var statecode = formContext.getAttribute("statecode").getValue();

    if (projectLookup && projectLookup.length > 0) {
        var projectId = projectLookup[0].id.replace(/[{}]/g, "");
        CAPS.DrawRequest.fetchTotalApproved(projectId, function (totalApproved) {
            CAPS.DrawRequest.fetchTotalRequested(projectId, currentDrawRequestId, false, function (totalRequested) {
                var remainingBalanceAttribute = formContext.getAttribute("caps_remainingdrawrequestbalance");

                var adjustedTotalRequested = totalRequested + amount;
                var remainingBalance = totalApproved - adjustedTotalRequested;
                remainingBalanceAttribute.setValue(remainingBalance);

                if (remainingBalance < 0 && statecode === 0) {
                    formContext.ui.setFormNotification("Please adjust the amount or project funds before making it eligible for submission.", "ERROR", "exceededBalance");
                    //amountControl.setNotification("Amount cannot exceed the remaining draw request balance.");
                } else {
                    formContext.ui.clearFormNotification("exceededBalance");
                    //amountControl.clearNotification();
                }
            });
        });
    }
};

CAPS.DrawRequest.fetchProjectRemainingBalance = function (projectId, callback) {
    var fetchXml = '<fetch>' +
        '<entity name="caps_projecttracker">' +
        '<attribute name="caps_remainingdrawrequestbalance"/>' +
        '<filter type="and">' +
        '<condition attribute="caps_projecttrackerid" operator="eq" value="' + projectId + '"/>' +
        '</filter>' +
        '</entity>' +
        '</fetch>';

    Xrm.WebApi.retrieveMultipleRecords("caps_projecttracker", "?fetchXml=" + encodeURIComponent(fetchXml)).then(
        function (response) {
            if (response.entities.length > 0) {
                var remainingDrawRequestBalance = response.entities[0].caps_remainingdrawrequestbalance || 0;
                callback(remainingDrawRequestBalance);
            } else {
                console.error("No project tracker found with the given ID.");
                callback(0);
            }
        },
        function (error) {
            console.error("Error fetching project tracker's remaining draw request balance: " + error.message);
            callback(0);
        }
    );
};

CAPS.DrawRequest.fetchTotalApproved = function (projectId, callback) {
    var fetchXml = '<fetch>' +
        '<entity name="caps_projecttracker">' +
        '<attribute name="caps_totalapproved"/>' +
        '<filter type="and">' +
        '<condition attribute="caps_projecttrackerid" operator="eq" value="' + projectId + '"/>' +
        '</filter>' +
        '</entity>' +
        '</fetch>';

    Xrm.WebApi.retrieveMultipleRecords("caps_projecttracker", "?fetchXml=" + encodeURIComponent(fetchXml)).then(
        function (response) {
            if (response.entities.length > 0) {
                var totalApproved = response.entities[0].caps_totalapproved || 0;
                callback(totalApproved);
            } else {
                console.error("No project tracker found with the given ID.");
                callback(0);
            }
        },
        function (error) {
            console.error("Error fetching total approved amount: " + error.message);
            callback(0);
        }
    );
};

CAPS.DrawRequest.fetchTotalRequested = function (projectId, currentDrawRequestId, includeCurrent, callback) {
    var fetchXml = '<fetch aggregate="true">' +
        '<entity name="caps_drawrequest">' +
        '<attribute name="caps_amount" aggregate="sum" alias="total_amount"/>' +
        '<filter type="and">' +
        '<condition attribute="statuscode" operator="ne" value="200870002"/>' + // Exclude cancelled requests
        '<condition attribute="caps_project" operator="eq" value="' + projectId + '"/>';

    if (!includeCurrent) {
        fetchXml += '<condition attribute="caps_drawrequestid" operator="ne" value="' + currentDrawRequestId + '"/>';
    }
    fetchXml += '</filter>' +
        '</entity>' +
        '</fetch>';

    Xrm.WebApi.retrieveMultipleRecords("caps_drawrequest", "?fetchXml=" + encodeURIComponent(fetchXml)).then(
        function (result) {
            var totalRequested = result.entities.length > 0 && result.entities[0].hasOwnProperty('total_amount') ? parseFloat(result.entities[0].total_amount) : 0;
            callback(totalRequested);
        },
        function (error) {
            console.error("Error fetching total requested amount: " + error.message);
            callback(0);
        }
    );
};

//Validate on Amount field changes in Draw Requests. 
CAPS.DrawRequest.amountValidation = function (executionContext) {
    var formContext = executionContext.getFormContext();
    var amountAttribute = formContext.getAttribute("caps_amount");
    var projectLookup = formContext.getAttribute("caps_project").getValue();
    var currentDrawRequestId = formContext.data.entity.getId();
    var amount = amountAttribute.getValue() || 0;
    var amountControl = formContext.getControl("caps_amount");
    var state = formContext.getAttribute("statecode").getValue();

    if (projectLookup && projectLookup.length > 0) {
        var projectId = projectLookup[0].id.replace(/[{}]/g, "");
        CAPS.DrawRequest.fetchTotalApproved(projectId, function (totalApproved) {
            CAPS.DrawRequest.fetchTotalRequested(projectId, currentDrawRequestId, function (totalRequested, isCurrentIncluded) {
                var formType = formContext.ui.getFormType();
                var remainingBalanceAttribute = formContext.getAttribute("caps_remainingdrawrequestbalance");

                var adjustedTotalRequested = totalRequested;
                if (isCurrentIncluded) {
                    adjustedTotalRequested -= amount;
                }

                var remainingBalance = totalApproved - adjustedTotalRequested;
                remainingBalance -= amount;

                remainingBalanceAttribute.setValue(remainingBalance);

                if (remainingBalance < 0 && state === 0) {
                    formContext.ui.setFormNotification("Please adjust the amount or project funds before making it eligible for submission.", "ERROR", "exceededBalance");
                    //amountControl.setNotification("Amount cannot exceed the remaining draw request balance.");
                } else {
                    formContext.ui.clearFormNotification("exceededBalance");
                    //amountControl.clearNotification();
                }
                formContext.data.save();
            }, true);
        });
    }
};
