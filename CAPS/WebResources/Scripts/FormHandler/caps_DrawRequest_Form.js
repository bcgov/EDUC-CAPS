﻿"use strict";

var CAPS = CAPS || {};
CAPS.DrawRequest = CAPS.DrawRequest || {};

//This JavaScript is to sum all draw request amounts except cancelled, and validate the amount that can be drawn on "amount" changes at draw request. 

//onLoad, check the project lookup 
CAPS.DrawRequest.form_onload = function (executionContext) {
    var formContext = executionContext.getFormContext();
    var amountField = formContext.getAttribute("caps_amount");
    var projectLookup = formContext.getAttribute("caps_project");

    if (projectLookup) {
        projectLookup.addOnChange(CAPS.DrawRequest.updateRemainingBalance);
        CAPS.DrawRequest.updateRemainingBalance(executionContext);
    }

    if (amountField) {
        amountField.addOnChange(CAPS.DrawRequest.amountValidation);
    }

    CAPS.DrawRequest.amountValidation(executionContext);
};

//IF CREATE form and the total requested of the related draw requests to the same project === 0 then remaining draw reuest balance = the project's remaining draw request balance.
//ELSE IF existing record and total requested so far <> 0, then set remaining draw request balance as Total Approved of Project - Total Requested of related draw requets.
CAPS.DrawRequest.updateRemainingBalance = function (executionContext) {
    var formContext = executionContext.getFormContext();
    var projectLookup = formContext.getAttribute("caps_project").getValue();
    var currentDrawRequestId = formContext.data.entity.getId();
    var formType = formContext.ui.getFormType();

    if (projectLookup && projectLookup.length > 0) {
        var projectId = projectLookup[0].id.replace(/[{}]/g, "");
        CAPS.DrawRequest.fetchTotalApproved(projectId, function (totalApproved) {
            CAPS.DrawRequest.fetchTotalRequested(projectId, currentDrawRequestId, function (totalRequested, isCurrentIncluded) {
                var remainingBalanceAttribute = formContext.getAttribute("caps_remainingdrawrequestbalance");

                if (formType === 1 && totalRequested === 0) {
                    CAPS.DrawRequest.fetchProjectRemainingBalance(projectId, function (projectRemainingBalance) {
                        remainingBalanceAttribute.setValue(projectRemainingBalance);
                    });
                } else {
                    var remainingBalance = totalApproved - totalRequested;
                    remainingBalanceAttribute.setValue(remainingBalance);
                }
            }, formType !== 1);
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

CAPS.DrawRequest.fetchTotalRequested = function (projectId, currentDrawRequestId, callback) {
    var fetchXml = '<fetch aggregate="true">' +
        '<entity name="caps_drawrequest">' +
        '<attribute name="caps_amount" aggregate="sum" alias="total_amount"/>' +
        '<filter type="and">' +
        '<condition attribute="statuscode" operator="ne" value="200870002"/>' +
        '<condition attribute="caps_project" operator="eq" value="' + projectId + '"/>' +
        '<condition attribute="caps_drawrequestid" operator="ne" value="' + currentDrawRequestId + '"/>' +
        '</filter>' +
        '</entity>' +
        '</fetch>';

    Xrm.WebApi.retrieveMultipleRecords("caps_drawrequest", "?fetchXml=" + encodeURIComponent(fetchXml)).then(
        function (result) {
            var totalRequested = result.entities.length > 0 && result.entities[0].hasOwnProperty('total_amount') ? parseFloat(result.entities[0].total_amount) : 0;
            callback(totalRequested);
        },
        function (error) {
            console.error("Error fetching total requested amount: " + error.message);
            callback(0, false);
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

                var amountControl = formContext.getControl("caps_amount");
                amountControl.clearNotification("Warning");
                if (0 > remainingBalance) {
                    amountControl.setNotification("Amount cannot exceed the remaining draw request balance.", "Warning");
                }
            }, true);
        });
    }
};
