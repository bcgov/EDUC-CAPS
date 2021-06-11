"use strict";

var CAPS = CAPS || {};
CAPS.COA = CAPS.COA || {};

const FORM_STATE = {
    UNDEFINED: 0,
    CREATE: 1,
    UPDATE: 2,
    READ_ONLY: 3,
    DISABLED: 4,
    BULK_EDIT: 6
};

/**
Main function for COA.  This function calls all other functions.
*/
CAPS.COA.onLoad = function (executionContext) {
    debugger;
    // Set variables
    var formContext = executionContext.getFormContext();
    var formState = formContext.ui.getFormType();

    //On Create
    if (formState === FORM_STATE.CREATE) {
        //Set Previous COA fields 
        CAPS.COA.SetPreviousCOAInformation(executionContext);

        formContext.getAttribute("caps_ptr").addOnChange(CAPS.COA.SetPreviousCOAInformation);
    }

    formContext.getAttribute("caps_previoustotalapprovedadvance").addOnChange(CAPS.COA.SetTotal);
    formContext.getAttribute("caps_changefrompreviouscoa").addOnChange(CAPS.COA.SetTotal);
}

/*
Called onLoad for new records, this function get's information from the previous COA and updates the relevant fields on the new COA.
*/
CAPS.COA.SetPreviousCOAInformation = function (executionContext) {
    var formContext = executionContext.getFormContext();
    //get Project
    var projectField = formContext.getAttribute("caps_ptr").getValue();

    if (projectField != null) {
        var project = projectField[0].id;

        var fetchXML = "<fetch version=\"1.0\" output-format=\"xml-platform\" mapping=\"logical\" distinct=\"true\">"+
                          "<entity name=\"caps_certificateofapproval\">"+
                            "<attribute name=\"caps_certificateofapprovalid\" />"+
                            "<attribute name=\"caps_name\" />"+
                            "<attribute name=\"caps_revisionnumber\" />" +
                            "<attribute name=\"caps_totalapprovedadvance\" />" +
                            "<attribute name=\"caps_expirydate\" />"+
                            "<order attribute=\"caps_name\" descending=\"false\" />"+
                            "<link-entity name=\"caps_projecttracker\" from=\"caps_currentcoa\" to=\"caps_certificateofapprovalid\" link-type=\"inner\" alias=\"ac\">"+
                              "<filter type=\"and\">"+
                                "<condition attribute=\"caps_projecttrackerid\" operator=\"eq\" value=\""+project+"\" />"+
                              "</filter>"+
                            "</link-entity>"+
                          "</entity>"+
                        "</fetch>";

        Xrm.WebApi.retrieveMultipleRecords("caps_certificateofapproval", "?fetchXml=" + fetchXML).then(
            function success(result) {
                if (result.entities.length === 1) {
                    var certificateNumber = result.entities[0]["caps_name"];
                    var revisionNumber = result.entities[0]["caps_revisionnumber"];
                    var previousTotal = result.entities[0]["caps_totalapprovedadvance"];
                    var expiryDate = result.entities[0]["caps_expirydate"];

                    formContext.getAttribute("caps_previouscertificatenumber").setValue(certificateNumber);
                    formContext.getAttribute("caps_previousrevisionnumber").setValue(revisionNumber);
                    formContext.getAttribute("caps_previoustotalapprovedadvance").setValue(previousTotal);

                    if (expiryDate != null) {
                        var parts = expiryDate.split('-');
                        var mydate = new Date(parts[0], parts[1] - 1, parts[2]);
                        formContext.getAttribute("caps_previousexpirydate").setValue(mydate);
                    }
                    
                }
                else {
                    //default previous value to 0
                    formContext.getAttribute("caps_previoustotalapprovedadvance").setValue(0);
                }
            },
            function (error) {
                alert(error.message);
            }
            );
    }
}

/*
Called onLoad and onChange of the increase/decrease field.  This function calculates the total COA amount.
*/
CAPS.COA.SetTotal = function (executionContext) {
    var formContext = executionContext.getFormContext();

    var previousAmount = formContext.getAttribute("caps_previoustotalapprovedadvance").getValue();
    var currentChange = formContext.getAttribute("caps_changefrompreviouscoa").getValue();

    formContext.getAttribute("caps_totalapprovedadvance").setValue(previousAmount + currentChange);
}