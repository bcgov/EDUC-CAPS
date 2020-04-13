using CAPS.DataContext;
using Microsoft.Xrm.Sdk;
using Microsoft.Xrm.Sdk.Query;
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
    /// Called when a project is added to a capital plan, this CWA checks if the submission category has a limit and if it's been reached.
    /// If it has an error is returned.
    /// </summary>
    public class CheckSubmissionCategoryLimit : CodeActivity
    {
        //Define the properties
        [Input("Submission Category")]
        [ReferenceTarget("caps_submissioncategory")]
        public InArgument<EntityReference> submissonCategory { get; set; }

        [Input("Capital Plan")]
        [ReferenceTarget("caps_submission")]
        public InArgument<EntityReference> capitalPlan { get; set; }

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

            tracingService.Trace("{0}{1}", "Start Custom Workflow Activity: CheckSubmissionCategoryLimit", DateTime.Now.ToLongTimeString());

            //Check to see if the EntityReference has been set
            EntityReference submissonCategoryReference = this.submissonCategory.Get(executionContext);
            if (submissonCategoryReference == null)
            {
                throw new InvalidOperationException("Submission Category has not been specified", new ArgumentNullException("Submission Category"));
            }
            else if (submissonCategoryReference.LogicalName != "caps_submissioncategory")
            {
                throw new InvalidOperationException("Input must reference a submission category record",
                    new ArgumentException("Input must be of type submission category", "Submission Category"));
            }

            EntityReference capitalPlanReference = this.capitalPlan.Get(executionContext);
            if (capitalPlanReference == null)
            {
                throw new InvalidOperationException("Capital Plan has not been specified", new ArgumentNullException("Capital Plan"));
            }
            else if (capitalPlanReference.LogicalName != "caps_submission")
            {
                throw new InvalidOperationException("Input must reference a capital plan record",
                    new ArgumentException("Input must be of type capital plan", "Capita Plan"));
            }

            //Set default values for outputs
            var error = false;
            var errorMessage = "";

            //Get Submission Category
            Microsoft.Xrm.Sdk.Query.ColumnSet columns = new Microsoft.Xrm.Sdk.Query.ColumnSet("caps_projectlimit");
            var submissionCategoryRecord = service.Retrieve(submissonCategoryReference.LogicalName, submissonCategoryReference.Id, columns) as CAPS.DataContext.caps_SubmissionCategory;

            //if submission category has a count, check the count in the capital plan, otherwise return
            if (submissionCategoryRecord.caps_ProjectLimit.GetValueOrDefault(0) > 0)
            {
                tracingService.Trace("Submission Category Limit: {0}", submissionCategoryRecord.caps_ProjectLimit.Value.ToString());

                using (var crmContext = new CrmServiceContext(service))
                {
                    var records = crmContext.caps_ProjectSet.Where(r => r.caps_SubmissionCategory.Id == submissonCategoryReference.Id && r.caps_Submission.Id == capitalPlanReference.Id).ToList();

                    tracingService.Trace("Record Count: {0}", records.Count.ToString());

                    if (records.Count() > submissionCategoryRecord.caps_ProjectLimit.Value)
                    {
                        error = true;
                        errorMessage = string.Format("This project cannot be added to {0} as you have already reached the limit of {1} {2} Projects.  All subsequent projects will not be added.", capitalPlanReference.Name, submissionCategoryRecord.caps_ProjectLimit.Value.ToString(), submissonCategoryReference.Name);
                    }
                }
            }
            this.error.Set(executionContext, error);
            this.errorMessage.Set(executionContext, errorMessage);
        }
    }
}
