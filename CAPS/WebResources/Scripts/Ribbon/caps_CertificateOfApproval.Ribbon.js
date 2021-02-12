"use strict";

var CAPS = CAPS || {};
CAPS.COA = CAPS.COA || {};


/**
 * Function to determine when the Close button should be shown.
 * @param {any} primaryControl primary control
 * @returns {boolean} true if should be shown, otherwise false.
 */
CAPS.COA.ShowClose = function (primaryControl) {
    debugger;
    var formContext = primaryControl;

    //TODO: check current state of coa, if approved (200,870,001)
    if (formContext.getAttribute("statuscode").getValue() !== 200870001) {
        return false;
    }

    var userRoles = Xrm.Utility.getGlobalContext().userSettings.roles;

    var showButton = false;

    userRoles.forEach(function hasFinancialDirectorRole(item, index) {
        if (item.name === "CAPS Financial Director Ministry User - Add On") {
            showButton = true;
        }
    });

    return showButton;
}

/**
 * Function to close the coa record.
 * @param {any} primaryControl primary control
 */
CAPS.COA.Close = function (primaryControl) {
    var formContext = primaryControl;

    var ptr = formContext.getAttribute("caps_ptr").getValue();

    if (ptr !== null && ptr[0] !== null) {
        var ptrID = ptr[0].id;

        var fetchXML = "<fetch version=\"1.0\" output-format=\"xml-platform\" mapping=\"logical\" distinct=\"false\">"+
                          "<entity name=\"caps_certificateofapproval\">"+
                            "<attribute name=\"caps_certificateofapprovalid\" />"+
                            "<attribute name=\"caps_name\" />"+
                            "<attribute name=\"createdon\" />"+
                            "<order attribute=\"caps_name\" descending=\"false\" />"+
                            "<filter type=\"and\">"+
                              "<condition attribute=\"statuscode\" operator=\"eq\" value=\"2\" />"+
                              "<condition attribute=\"caps_ptr\" operator=\"eq\"  value=\""+ptrID+"\" />"+
                            "</filter>"+
                          "</entity>"+
                        "</fetch>";

        Xrm.WebApi.retrieveMultipleRecords("caps_certificateofapproval", "?fetchXml=" + fetchXML).then(
        function success(result) {
            if (result.entities.length > 0) {

                let alertStrings = { confirmButtonLabel: "OK", text: "There is a submitted COA for this project.  Mark it as Approved or Rejected and then Close the COA.", title: "COA Closure" };
                let alertOptions = { height: 120, width: 260 };
                Xrm.Navigation.openAlertDialog(alertStrings, alertOptions).then(
                    function success(result) {
                        console.log("Alert dialog closed");
                    },
                    function (error) {
                        console.log(error.message);
                    }
                );
            }
            else {
                //Change status to Closed
                let confirmStrings = { text: "This will close the COA in CAPS, it will not close the COA in the Ministry of Finance's system.  This action cannot be undone.  Click OK to continue or Cancel to exit.", title: "Close Confirmation" };
                let confirmOptions = { height: 200, width: 450 };
                Xrm.Navigation.openConfirmDialog(confirmStrings, confirmOptions).then(
                    function (success) {
                        if (success.confirmed) {
                            debugger;
                            formContext.getAttribute("statecode").setValue(1);
                            formContext.getAttribute("statuscode").setValue(200870004);
                            formContext.data.entity.save();
                        }

                    });
            }
        },
        function (error) {
            alert(error.message);
        }
    );


    }
}
