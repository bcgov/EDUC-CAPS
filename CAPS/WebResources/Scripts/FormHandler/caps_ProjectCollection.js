"use strict";

/* INCLUDE CAPS.Common.js */

var CAPS = CAPS || {};
CAPS.ProjectCollection = CAPS.ProjectCollection || {
    
};

CAPS.ProjectCollection.OnLoad = function (executionContext) {
   
    
    CAPS.ProjectCollection.preFilterProjectRequestLookup(executionContext);
}

CAPS.ProjectCollection.preFilterProjectRequestLookup = function (executionContext) {
    
    var formContext = executionContext.getFormContext();
    if (formContext.getControl("caps_primaryprojectrequest") === null)
        return;

    formContext.getControl("caps_primaryprojectrequest").addPreSearch(CAPS.ProjectCollection.addProjectRequestLookupFilter);
}

CAPS.ProjectCollection.addProjectRequestLookupFilter = function (executionContext) {
   
    var formContext = executionContext.getFormContext();
    var projectCollectionID = formContext.data.entity.getId();
    var fetchXml = "<filter type='and'><condition attribute='caps_projectcollection' operator='eq' value='" + projectCollectionID + "' /></filter>";
    formContext.getControl("caps_primaryprojectrequest").addCustomFilter(fetchXml);
    
}

CAPS.ProjectCollection.GetLookup = function (fieldName, formContext) {
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
CAPS.ProjectCollection.RemoveCurlyBraces = function (str) {
    return str.replace(/[{}]/g, "");
}
