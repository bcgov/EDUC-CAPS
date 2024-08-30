"use strict";

var CAPS = CAPS || {};
CAPS.ChildCareProgressReport = CAPS.ChildCareProgressReport || {};

// This is to handle show/hide/require fields depending on the selected values on CC Progress Report record

CAPS.ChildCareProgressReport.form_onload = function (executionContext) {
    var formContext = executionContext.getFormContext();

    // Check and apply logic for each field on form load
    CAPS.ChildCareProgressReport.checkChangesToLeaseAgreement(formContext);
    CAPS.ChildCareProgressReport.checkChangesWithDirectorshipNFP(formContext);
    CAPS.ChildCareProgressReport.checkAffordabilityInitiatives(formContext);
    CAPS.ChildCareProgressReport.checkLicenseCapacityChange(formContext);
    CAPS.ChildCareProgressReport.checkMonthsFullyClosed(formContext);

    // Add OnChange handlers for the fields
    formContext.getAttribute("caps_anynewchangestotheleaseagreement").addOnChange(function () {
        CAPS.ChildCareProgressReport.checkChangesToLeaseAgreement(formContext);
    });
    formContext.getAttribute("caps_ifnotforprofitanychangeswithdirectorship").addOnChange(function () {
        CAPS.ChildCareProgressReport.checkChangesWithDirectorshipNFP(formContext);
    });
    formContext.getAttribute("caps_affordabilityinitiativesyouareenrolledin").addOnChange(function () {
        CAPS.ChildCareProgressReport.checkAffordabilityInitiatives(formContext);
    });
    formContext.getAttribute("caps_achangeinlicensecapacitysincelastyear").addOnChange(function () {
        CAPS.ChildCareProgressReport.checkLicenseCapacityChange(formContext);
    });
    formContext.getAttribute("caps_anymonthsfullyclosedforallprogramsatcc").addOnChange(function () {
        CAPS.ChildCareProgressReport.checkMonthsFullyClosed(formContext);
    });
};

CAPS.ChildCareProgressReport.checkChangesToLeaseAgreement = function (formContext) {
    var isAnyNewChanges = formContext.getAttribute("caps_anynewchangestotheleaseagreement").getValue();

    if (isAnyNewChanges) {
        formContext.getControl("caps_commentanynewchangestotheleaseagreement").setVisible(true);
        formContext.getAttribute("caps_commentanynewchangestotheleaseagreement").setRequiredLevel("required");
    } else {
        formContext.getControl("caps_commentanynewchangestotheleaseagreement").setVisible(false);
        formContext.getAttribute("caps_commentanynewchangestotheleaseagreement").setRequiredLevel("none");
        formContext.getAttribute("caps_commentanynewchangestotheleaseagreement").setValue(null);
    }
};

CAPS.ChildCareProgressReport.checkChangesWithDirectorshipNFP = function (formContext) {
    var isChangesWithDirectorship = formContext.getAttribute("caps_ifnotforprofitanychangeswithdirectorship").getValue();

    if (isChangesWithDirectorship) {
        formContext.getControl("caps_commentanychangeswithdirectorshipnfp").setVisible(true);
        formContext.getAttribute("caps_commentanychangeswithdirectorshipnfp").setRequiredLevel("required");
    } else {
        formContext.getControl("caps_commentanychangeswithdirectorshipnfp").setVisible(false);
        formContext.getAttribute("caps_commentanychangeswithdirectorshipnfp").setRequiredLevel("none");
        formContext.getAttribute("caps_commentanychangeswithdirectorshipnfp").setValue(null);
    }
};

CAPS.ChildCareProgressReport.checkAffordabilityInitiatives = function (formContext) {
    var affordabilityInitiatives = formContext.getAttribute("caps_affordabilityinitiativesyouareenrolledin").getValue();
    var initiativeCode = 714430003; // Other

    if (affordabilityInitiatives && affordabilityInitiatives.includes(initiativeCode)) {
        formContext.getControl("caps_otheraffordabilityinitiatives").setVisible(true);
        formContext.getAttribute("caps_otheraffordabilityinitiatives").setRequiredLevel("required");
    } else {
        formContext.getControl("caps_otheraffordabilityinitiatives").setVisible(false);
        formContext.getAttribute("caps_otheraffordabilityinitiatives").setRequiredLevel("none");
        formContext.getAttribute("caps_otheraffordabilityinitiatives").setValue(null);
    }
};

CAPS.ChildCareProgressReport.checkLicenseCapacityChange = function (formContext) {
    var isLicenseCapacityChanged = formContext.getAttribute("caps_achangeinlicensecapacitysincelastyear").getValue();

    if (isLicenseCapacityChanged) {
        formContext.getControl("caps_changeinlicensedcapacity30monthstosa").setVisible(true);
        formContext.getControl("caps_changeinlicensedcapacitymultiage").setVisible(true);
        formContext.getControl("caps_changeinlicensedcapacitypreschool").setVisible(true);
        formContext.getControl("caps_changeinlicensedcapacitysasg").setVisible(true);
        formContext.getControl("caps_changeinlicensedcapacityschoolage").setVisible(true);
        formContext.getControl("caps_changeinlicensedcapacityunder36months").setVisible(true);

        formContext.getAttribute("caps_changeinlicensedcapacity30monthstosa").setRequiredLevel("required");
        formContext.getAttribute("caps_changeinlicensedcapacitymultiage").setRequiredLevel("required");
        formContext.getAttribute("caps_changeinlicensedcapacitypreschool").setRequiredLevel("required");
        formContext.getAttribute("caps_changeinlicensedcapacitysasg").setRequiredLevel("required");
        formContext.getAttribute("caps_changeinlicensedcapacityschoolage").setRequiredLevel("required");
        formContext.getAttribute("caps_changeinlicensedcapacityunder36months").setRequiredLevel("required");
    } else {
        formContext.getControl("caps_changeinlicensedcapacity30monthstosa").setVisible(false);
        formContext.getControl("caps_changeinlicensedcapacitymultiage").setVisible(false);
        formContext.getControl("caps_changeinlicensedcapacitypreschool").setVisible(false);
        formContext.getControl("caps_changeinlicensedcapacitysasg").setVisible(false);
        formContext.getControl("caps_changeinlicensedcapacityschoolage").setVisible(false);
        formContext.getControl("caps_changeinlicensedcapacityunder36months").setVisible(false);

        formContext.getAttribute("caps_changeinlicensedcapacity30monthstosa").setRequiredLevel("none");
        formContext.getAttribute("caps_changeinlicensedcapacitymultiage").setRequiredLevel("none");
        formContext.getAttribute("caps_changeinlicensedcapacitypreschool").setRequiredLevel("none");
        formContext.getAttribute("caps_changeinlicensedcapacitysasg").setRequiredLevel("none");
        formContext.getAttribute("caps_changeinlicensedcapacityschoolage").setRequiredLevel("none");
        formContext.getAttribute("caps_changeinlicensedcapacityunder36months").setRequiredLevel("none");

        formContext.getAttribute("caps_changeinlicensedcapacity30monthstosa").setValue(null);
        formContext.getAttribute("caps_changeinlicensedcapacitymultiage").setValue(null);
        formContext.getAttribute("caps_changeinlicensedcapacitypreschool").setValue(null);
        formContext.getAttribute("caps_changeinlicensedcapacitysasg").setValue(null);
        formContext.getAttribute("caps_changeinlicensedcapacityschoolage").setValue(null);
        formContext.getAttribute("caps_changeinlicensedcapacityunder36months").setValue(null);
    }
};

CAPS.ChildCareProgressReport.checkMonthsFullyClosed = function (formContext) {
    var isMonthsFullyClosed = formContext.getAttribute("caps_anymonthsfullyclosedforallprogramsatcc").getValue();

    if (isMonthsFullyClosed) {
        formContext.getControl("caps_monthsfullyclosedforallprogramsatchildcare").setVisible(true);
        formContext.getAttribute("caps_monthsfullyclosedforallprogramsatchildcare").setRequiredLevel("required");
    } else {
        formContext.getControl("caps_monthsfullyclosedforallprogramsatchildcare").setVisible(false);
        formContext.getAttribute("caps_monthsfullyclosedforallprogramsatchildcare").setRequiredLevel("none");
        formContext.getAttribute("caps_monthsfullyclosedforallprogramsatchildcare").setValue(null);
    }
};