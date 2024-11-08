"use strict";

var CAPS = CAPS || {};
CAPS.ChildCareQuarterlyReport = CAPS.ChildCareQuarterlyReport || {
   
};

CAPS.ChildCareQuarterlyReport.onLoad = function (executionContext) {
    
    var formContext = executionContext.getFormContext();
    CAPS.ChildCareQuarterlyReport.ShowHideFields(executionContext);
    formContext.getAttribute("caps_notdirectlyoperatingsecuredauthorizedoperator").addOnChange(CAPS.ChildCareQuarterlyReport.ShowHideFields);
    formContext.getAttribute("caps_whattypeofauthorizedoperator").addOnChange(CAPS.ChildCareQuarterlyReport.ShowHideFields);
    formContext.getAttribute("caps_receivedfundingfromadditionalsources").addOnChange(CAPS.ChildCareQuarterlyReport.ShowHideFields);
};

CAPS.ChildCareQuarterlyReport.ShowHideFields = function (executionContext) {

    var formContext = executionContext.getFormContext();
    if (formContext.getAttribute("caps_selfoperated").getValue() === true) {
        formContext.getControl("caps_notdirectlyoperatingsecuredauthorizedoperator").setVisible(false);
        formContext.getControl("caps_nameofauthorizedoperator").setVisible(false);
    }
    else {
        formContext.getControl("caps_notdirectlyoperatingsecuredauthorizedoperator").setVisible(true);
        formContext.getControl("caps_nameofauthorizedoperator").setVisible(true);
        //formContext.getAttribute("caps_nameofauthorizedoperator").setValue(null);
    }
    //If no Self Operated is No
    if (formContext.getAttribute("caps_notdirectlyoperatingsecuredauthorizedoperator").getValue() === 746660001 ||
        formContext.getAttribute("caps_notdirectlyoperatingsecuredauthorizedoperator").getValue() === null ||
        formContext.getAttribute("caps_selfoperated").getValue() === true) {
        formContext.getControl("caps_whattypeofauthorizedoperator").setVisible(false);
        formContext.getControl("caps_forprofitoperatorjustification").setVisible(false);
        formContext.getAttribute("caps_whattypeofauthorizedoperator").setValue(null);
        formContext.getAttribute("caps_forprofitoperatorjustification").setValue(null);
        formContext.getControl("caps_nameofauthorizedoperator").setVisible(false);
        formContext.getAttribute("caps_nameofauthorizedoperator").setValue(null);
    }
    else {
        formContext.getControl("caps_whattypeofauthorizedoperator").setVisible(true);
        //If what type of authorized operator is Public/Not-For-Profit
        if (formContext.getAttribute("caps_whattypeofauthorizedoperator").getValue() === 746660000 ||
            formContext.getAttribute("caps_whattypeofauthorizedoperator").getValue() === null) {
            formContext.getControl("caps_forprofitoperatorjustification").setVisible(false);
            formContext.getAttribute("caps_forprofitoperatorjustification").setValue(null);
            formContext.getControl("caps_nameofauthorizedoperator").setVisible(true);
            
        }
        else {
            formContext.getControl("caps_forprofitoperatorjustification").setVisible(true);
        }
    }
    //If Received funding from additional sources is No 
    if (formContext.getAttribute("caps_receivedfundingfromadditionalsources").getValue() === 746660001 ||
        formContext.getAttribute("caps_receivedfundingfromadditionalsources").getValue() === null) {
        formContext.getControl("caps_amountreceivedandfromwhom").setVisible(false);
        formContext.getAttribute("caps_amountreceivedandfromwhom").setValue(null);
    }
    else {
        formContext.getControl("caps_amountreceivedandfromwhom").setVisible(true);
    }

}
