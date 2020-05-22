

function RunOnSelected(executionContext) {
    var selected = executionContext.getFormContext().data.entity;
    var Id = selected.getId();

    var pageInput = {
        pageType: "entityrecord",
        entityName: "caps_projectcashflow",
        entityId: Id
    };
    var navigationOptions = {
        target: 2,
        height: { value: 80, unit: "%" },
        width: { value: 80, unit: "%" },
        position: 1
    };
    Xrm.Navigation.navigateTo(pageInput, navigationOptions).then(
        function success() { },
        function error() { }
    );
}


