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
    /// Called when a project is associated to a phased project group and the facility or year or phased project group changes.  
    /// This custom workflow activity verifies that no other project in the phased project group is for the same start year.
    /// </summary>
    public class CheckPhasedProjectYear : CodeActivity
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

            tracingService.Trace("{0}{1}", "Start Custom Workflow Activity: CheckPhasedProjectYear", DateTime.Now.ToLongTimeString());

            //Get Project & Phased Project Group Field
            Microsoft.Xrm.Sdk.Query.ColumnSet columns = new Microsoft.Xrm.Sdk.Query.ColumnSet("caps_phasedprojectgroup", "caps_startdate");
            var projectRecord = service.Retrieve(context.PrimaryEntityName, context.PrimaryEntityId, columns) as caps_Project;

            if (!projectRecord.caps_StartDate.HasValue) return;

            //Set default values for outputs
            var error = false;
            var errorMessage = "";

            var proejctStartYear = projectRecord.caps_StartDate.Value.Year;

            //get projects that are part of the phased project group
            using (var crmContext = new CrmServiceContext(service))
            {
                var records = crmContext.caps_ProjectSet.Where(r => r.caps_PhasedProjectGroup.Id == projectRecord.caps_PhasedProjectGroup.Id && r.caps_ProjectId != context.PrimaryEntityId);

                foreach (var record in records)
                {
                    //get start date
                    if (record.caps_StartDate.HasValue)
                    {
                        if (record.caps_StartDate.Value.Year == proejctStartYear)
                        {
                            //Start Year was already used so need to display error.
                            error = true;
                            errorMessage = "Another project in the phased project group is for the same year.";
                            break;
                        }
                    }
                }

                this.error.Set(executionContext, error);
                this.errorMessage.Set(executionContext, errorMessage);
            }
        }
    }
}
