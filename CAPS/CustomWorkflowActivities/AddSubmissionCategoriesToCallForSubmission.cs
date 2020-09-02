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
    /// Called when call for submission is created.  This CWA adds submission categories with the same type (ie Major) to the call for submission.
    /// </summary>
    public class AddSubmissionCategoriesToCallForSubmission : CodeActivity
    {
        //Define the properties
        [Input("Call for Submission")]
        [ReferenceTarget("caps_callforsubmission")]
        public InArgument<EntityReference> callForSubmission { get; set; }

        protected override void Execute(CodeActivityContext executionContext)
        {
            ITracingService tracingService = executionContext.GetExtension<ITracingService>();

            IWorkflowContext context = executionContext.GetExtension<IWorkflowContext>();
            IOrganizationServiceFactory serviceFactory = executionContext.GetExtension<IOrganizationServiceFactory>();
            IOrganizationService service = serviceFactory.CreateOrganizationService(context.UserId);

            tracingService.Trace("{0}{1}", "Start Custom Workflow Activity: AddSubmissionCategoriesToCallForSubmission", DateTime.Now.ToLongTimeString());

            //Check to see if the EntityReference has been set
            EntityReference callForSubmissionReference = this.callForSubmission.Get(executionContext);
            if (callForSubmissionReference == null)
            {
                throw new InvalidOperationException("Call for Submission has not been specified", new ArgumentNullException("Call for Submission"));
            }
            else if (callForSubmissionReference.LogicalName != "caps_callforsubmission")
            {
                throw new InvalidOperationException("Input must reference a call for submission record",
                    new ArgumentException("Input must be of type call for submission", "Call for Submission"));
            }
            tracingService.Trace("Line:{0}", "41");

            var callForSubmisisonRecord = service.Retrieve(callForSubmissionReference.LogicalName, callForSubmissionReference.Id, new Microsoft.Xrm.Sdk.Query.ColumnSet("caps_callforsubmissiontype")) as caps_CallForSubmission;

            tracingService.Trace("Line:{0}", "45");

            using (var crmContext = new CrmServiceContext(service))
            {
                var relatedEntities = new EntityReferenceCollection();

                //get active school districts
                var submissionCategories = crmContext.caps_SubmissionCategorySet.Where(r => r.StateCode == caps_SubmissionCategoryState.Active && r.caps_CallforSubmissionType == callForSubmisisonRecord.caps_CallforSubmissionType);

                tracingService.Trace("Line:{0}", "45");

                foreach (var submissionCategory in submissionCategories)
                {
                    tracingService.Trace("category:{0}", submissionCategory.caps_Name);

                    relatedEntities.Add(new EntityReference(submissionCategory.LogicalName, submissionCategory.Id));
                }

                Relationship relationship = new Relationship(caps_CallForSubmission_caps_SubmissionCat.EntitySchemaName);
                service.Associate(callForSubmissionReference.LogicalName, callForSubmissionReference.Id, relationship, relatedEntities);
            }

        }
    }
}
