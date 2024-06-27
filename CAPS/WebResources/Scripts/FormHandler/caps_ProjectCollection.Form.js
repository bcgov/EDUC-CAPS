"use strict";

var CAPS = CAPS || {};
CAPS.ProjectCollection = CAPS.ProjectCollection || {

};

const FORM_STATE = {
    UNDEFINED: 0,
    CREATE: 1,
    UPDATE: 2,
    READ_ONLY: 3,
    DISABLED: 4,
    BULK_EDIT: 6
};

CAPS.ProjectCollection.OnLoad = function (executionContext) {
    var formContext = executionContext.getFormContext();
    var formState = formContext.ui.getFormType();
    if (formState === FORM_STATE.CREATE) {
        CAPS.ProjectCollection.FormatString(executionContext);
    }

    CAPS.ProjectCollection.preFilterProjectRequestLookup(executionContext);
    CAPS.ProjectCollection.FormatString(executionContext);   
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

CAPS.ProjectCollection.FormatString = function (executionContext) {
    var formContext = executionContext.getFormContext();
    var proposedSchoolFacility = formContext.getAttribute("caps_proposedschoolfacility").getValue();
    if (proposedSchoolFacility !== null) {

        var formatProposedSchoolFacility = proposedSchoolFacility.replace(/[^a-zA-Z ]/g, "%27");
        formContext.getAttribute("caps_formattedproposedschoolfacility").setValue(formatProposedSchoolFacility);
    }
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