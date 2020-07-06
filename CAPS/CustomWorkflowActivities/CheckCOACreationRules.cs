using CAPS.DataContext;
using Microsoft.Xrm.Sdk;
using Microsoft.Xrm.Sdk.Workflow;
using System;
using System.Activities;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace CustomWorkflowActivities
{
    /// <summary>
    /// This CWA is called on creation of a COA to confirm there are no other COAs in Draft or Submitted state.
    /// </summary>
    public class CheckCOACreationRules : CodeActivity
    {
        [Output("Error")]
        public OutArgument<bool> error { get; set; }

        [Output("ErrorMessage")]
        public OutArgument<string> errorMessage { get; set; }

        protected override void Execute(CodeActivityContext executionContext)
        {
            ITracingService tracingService = executionContext.GetExtension<ITracingService>();

            IWorkflowContext context = executionContext.GetExtension<IWorkflowContext>();
            IOrganizationServiceFactory serviceFactory = executionContext.GetExtension<IOrganizationServiceFactory>();
            IOrganizationService service = serviceFactory.CreateOrganizationService(context.UserId);

            tracingService.Trace("{0}{1}", "Start Custom Workflow Activity: CheckCOACreationRules", DateTime.Now.ToLongTimeString());

            //Get Project field
            Microsoft.Xrm.Sdk.Query.ColumnSet columns = new Microsoft.Xrm.Sdk.Query.ColumnSet("caps_ptr");
            var coaRecord = service.Retrieve(context.PrimaryEntityName, context.PrimaryEntityId, columns) as caps_CertificateofApproval;

            //might want to return an error if this is missing a Project
            if (coaRecord.caps_PTR == null || coaRecord.caps_PTR.Id == null) return;

            //Set default values for outputs
            var error = false;
            var errorMessage = "";

            
            using (var crmContext = new CrmServiceContext(service))
            {
                //Check if there are any other COAs for this project in a Draft or Submitted State
                var records = crmContext.caps_CertificateofApprovalSet.Where(r => r.caps_PTR.Id == coaRecord.caps_PTR.Id 
                            && r.caps_CertificateofApprovalId != context.PrimaryEntityId 
                            && (r.StatusCode.Value == (int)caps_CertificateofApproval_StatusCode.Draft || r.StatusCode.Value == (int)caps_CertificateofApproval_StatusCode.Submitted)).ToList();

                if (records.Count() > 0)
                {
                    //Start Year was already used so need to display error.
                    error = true;
                    errorMessage = "Unable to create COA as there is already a Draft or Submitted one for this Project.";
                }

                this.error.Set(executionContext, error);
                this.errorMessage.Set(executionContext, errorMessage);
            }
        }
    }
}
