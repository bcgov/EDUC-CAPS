"use strict";

var CAPS = CAPS || {};
CAPS.Project = CAPS.Project || { GLOBAL_FORM_CONTEXT: null };

CAPS.Project.DefaultLookupIfSingle = function (formContext, fieldName, entityLogicalName, entityIdField, entityDisplayField) {
    //Get Default View for lookup
    var defaultView = formContext.getControl(fieldName).getDefaultView();

    //Get Fetch XML for the View
    Xrm.WebApi.retrieveRecord("savedquery", defaultView, "?$select=fetchxml").then(
        function success(result) {
            var fetchXML = result.fetchXML;

            //Now retrieve records from Fetch

            Xrm.WebApi.retrieveMultipleRecords(entityLogicalName, fetchXML).then(
                function success(result) { 
                    if (result.entities.length === 1) {
                        //only 1 result, so populate dropdown with it

                        formContext.getAttribute(fieldName).setValue([{ id: entityIdField, name: entityDisplayField, entityType: entityLogicalName}]);
                    }
                },
                function (error) { }
       
            );

        },
        function (error) {
            console.log(error.message);
            // handle error conditions
        }
    );
}