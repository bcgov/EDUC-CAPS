"use strict";

var CAPS = CAPS || {};
CAPS.Common = CAPS.Common || { GLOBAL_FORM_CONTEXT: null };

/**
 * This function calls the default lookup query for a field.  If the query returns only one result than the field is updated with that value.
 * @param {any} formContext Form Context
 * @param {any} fieldName Logical Name of the field
 * @param {any} entityLogicalName Logical Name of the entity
 * @param {any} entityIdField Logical Name of the ID field on the entity (usually the entity name followed by id)
 * @param {any} entityDisplayField Logical Name of the display field on the entity (usually name or prefix_name unless changed when entity is created)
 * @param {any} filterFetch Optional: filtering FetchXML for the lookup
 */
CAPS.Common.DefaultLookupIfSingle = function (formContext, fieldName, entityLogicalName, entityIdField, entityDisplayField, filterFetch) {
   
    //Get Default View for lookup
    var defaultView = formContext.getControl(fieldName).getDefaultView();

    //Get Fetch XML for the View
    Xrm.WebApi.retrieveRecord("savedquery", defaultView, "?$select=fetchxml").then(
        function success(result) {
 
            var fetchXML = result.fetchxml;

            if (filterFetch !== null && filterFetch.length > 0) {
                filterFetch = filterFetch + "</entity></fetch>";
                fetchXML = fetchXML.replace("</entity></fetch>", filterFetch);
            }

            //Now retrieve records from Fetch
            Xrm.WebApi.retrieveMultipleRecords(entityLogicalName, "?fetchXml=" +fetchXML).then(
                function success(result) { 

                    if (result.entities.length === 1) {
                        //only 1 result, so populate dropdown with it
                        var singleRecord = result.entities[0];
                        var recordId = singleRecord[entityIdField];
                        var recordValue = singleRecord[entityDisplayField];

                        formContext.getAttribute(fieldName).setValue([{ id: recordId, name: recordValue, entityType: entityLogicalName}]);
                    }
                },
                function (error) {
                }
       
            );

        },
        function (error) {
            console.log(error.message);
            // handle error conditions
        }
    );
}