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
    /// This CWA validates the capital plan by checking that all projects included in it are valid.
    /// </summary>
    public class ValidateCapitalPlan : CodeActivity
    {
        [Output("Valid")]
        public OutArgument<bool> valid { get; set; }


        protected override void Execute(CodeActivityContext executionContext)
        {
            ITracingService tracingService = executionContext.GetExtension<ITracingService>();

            IWorkflowContext context = executionContext.GetExtension<IWorkflowContext>();
            IOrganizationServiceFactory serviceFactory = executionContext.GetExtension<IOrganizationServiceFactory>();
            IOrganizationService service = serviceFactory.CreateOrganizationService(context.UserId);

            tracingService.Trace("{0}{1}", "Start Custom Workflow Activity: ValidateCapitalPlan", DateTime.Now.ToLongTimeString());

            var recordId = context.PrimaryEntityId;

            //VALIDATION STARTS HERE
            bool isValid = true;
            //StringBuilder validationMessage = new StringBuilder();

            //get all project requests
            using (var crmContext = new CrmServiceContext(service))
            {
                tracingService.Trace("{0}", "Loading data");
                var projectRequests = crmContext.caps_ProjectSet.Where(r=>r.caps_Submission.Id == recordId);

                foreach(var projectRequest in projectRequests)
                {
                    OrganizationRequest req = new OrganizationRequest();
                    req.RequestName = "caps_ValidateProjectRequest";
                    req.Parameters.Add("Target", new EntityReference(caps_Project.EntityLogicalName, projectRequest.Id));

                    OrganizationResponse response = service.Execute(req);

                    //check to see if the validation passed
                    if (!((bool?)response.Results["IsValid"]).GetValueOrDefault(true))
                    {
                        isValid = false;
                    }

                }
            }

            this.valid.Set(executionContext, isValid);
            //this.message.Set(executionContext, validationMessage.ToString());
        }
    }
}
