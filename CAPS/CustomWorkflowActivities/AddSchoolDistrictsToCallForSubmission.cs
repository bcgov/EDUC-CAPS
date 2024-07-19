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
    /// Called when call for submission is created.  This custom workflow activity associates all active school districts to the call for submission.
    /// </summary>
    public class AddSchoolDistrictsToCallForSubmission : CodeActivity
    {
        //Define the properties
        [Input("Call for Submission")]
        [ReferenceTarget("caps_callforsubmission")]
        public InArgument<EntityReference> callForSubmission { get; set; }

        [Input("Call Type")]
        [AttributeTarget("caps_callforsubmission", "caps_callforsubmissiontype")]
        public InArgument<OptionSetValue> callType { get; set; }

        protected override void Execute(CodeActivityContext executionContext)
        {
            ITracingService tracingService = executionContext.GetExtension<ITracingService>();

            IWorkflowContext context = executionContext.GetExtension<IWorkflowContext>();
            IOrganizationServiceFactory serviceFactory = executionContext.GetExtension<IOrganizationServiceFactory>();
            IOrganizationService service = serviceFactory.CreateOrganizationService(context.UserId);

            tracingService.Trace("{0}{1}", "Start Custom Workflow Activity: AddSchoolDistrictsToCallForSubmission", DateTime.Now.ToLongTimeString());

            //Check to see if the EntityReference has been set
            EntityReference callForSubmissionReference = this.callForSubmission.Get(executionContext);
            OptionSetValue callTypeOptionSet = this.callType.Get(executionContext);
            
            if (callForSubmissionReference == null)
            {
                throw new InvalidOperationException("Call for Submission has not been specified", new ArgumentNullException("Call for Submission"));
            }
            else if (callForSubmissionReference.LogicalName != "caps_callforsubmission")
            {
                throw new InvalidOperationException("Input must reference a call for submission record",
                    new ArgumentException("Input must be of type call for submission", "Call for Submission"));
            }

            using (var crmContext = new CrmServiceContext(service))
            {
                var relatedEntities = new EntityReferenceCollection();
                IQueryable<edu_schooldistrict> schoolDistricts;
                //If call type is CC-AFG
                if (callTypeOptionSet.Value == 385610001)
                {
                    //Get Active School Districts with CC AFG Ratio greater than 0
                    schoolDistricts = crmContext.edu_schooldistrictSet.Where(r => r.StateCode == edu_schooldistrictState.Active
                                                                                   && r.caps_CCAFGRatio > 0);
                }
                else 
                {
                    //Get active school districts
                    schoolDistricts = crmContext.edu_schooldistrictSet.Where(r => r.StateCode == edu_schooldistrictState.Active);
                }
                

                foreach(var schoolDistrict in schoolDistricts)
                {
                    relatedEntities.Add(new EntityReference(edu_schooldistrict.EntityLogicalName, schoolDistrict.Id));
                }

                Relationship relationship = new Relationship(caps_CallForSubmission_edu_schooldistrict.EntitySchemaName);
                service.Associate(callForSubmissionReference.LogicalName, callForSubmissionReference.Id, relationship, relatedEntities);
            }

        }
    }
}
