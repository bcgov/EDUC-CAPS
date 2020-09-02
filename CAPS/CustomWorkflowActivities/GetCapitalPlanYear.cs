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
    /// Get's the capital plan year from the related Call for Submission.
    /// </summary>
    public class GetCapitalPlanYear : CodeActivity
    {
        //Define the properties
        [Input("Capital Plan")]
        [ReferenceTarget("caps_submission")]
        public InArgument<EntityReference> capitalPlan { get; set; }

        [Output("Year")]
        [ReferenceTarget("edu_year")]
        public OutArgument<EntityReference> year { get; set; }

        protected override void Execute(CodeActivityContext executionContext)
        {
            ITracingService tracingService = executionContext.GetExtension<ITracingService>();

            IWorkflowContext context = executionContext.GetExtension<IWorkflowContext>();
            IOrganizationServiceFactory serviceFactory = executionContext.GetExtension<IOrganizationServiceFactory>();
            IOrganizationService service = serviceFactory.CreateOrganizationService(context.UserId);

            tracingService.Trace("{0}{1}", "Start Custom Workflow Activity: GetCapitalPlanYear", DateTime.Now.ToLongTimeString());

            //Check to see if the EntityReference has been set
            EntityReference capitalPlanReference = this.capitalPlan.Get(executionContext);
            if (capitalPlanReference == null)
            {
                throw new InvalidOperationException("Capital Plan has not been specified", new ArgumentNullException("Capital Plan"));
            }
            else if (capitalPlanReference.LogicalName != "caps_submission")
            {
                throw new InvalidOperationException("Input must reference a capital plan record",
                    new ArgumentException("Input must be of type capital plan", "Capital Plan"));
            }

            var capitalPlanRecord = service.Retrieve(capitalPlanReference.LogicalName, capitalPlanReference.Id, new Microsoft.Xrm.Sdk.Query.ColumnSet("caps_callforsubmission")) as caps_Submission;
            var callForSubmisisonRecord = service.Retrieve(capitalPlanRecord.caps_CallforSubmission.LogicalName, capitalPlanRecord.caps_CallforSubmission.Id, new Microsoft.Xrm.Sdk.Query.ColumnSet("caps_capitalplanyear")) as caps_CallForSubmission;

            tracingService.Trace("Year:{0}", callForSubmisisonRecord.caps_CapitalPlanYear.Id);

            this.year.Set(executionContext, new EntityReference(callForSubmisisonRecord.caps_CapitalPlanYear.LogicalName, callForSubmisisonRecord.caps_CapitalPlanYear.Id));
        }
    }
}
