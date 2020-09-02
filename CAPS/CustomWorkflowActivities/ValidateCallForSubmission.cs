using CAPS.DataContext;
using Microsoft.Xrm.Sdk;
using Microsoft.Xrm.Sdk.Workflow;
using Microsoft.Xrm.Sdk.Query;
using System;
using System.Activities;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace CustomWorkflowActivities
{
    /// <summary>
    /// This CWA validates the call for submission 
    /// including checking the type on the selected submission categories matches the one on call for submission.
    /// </summary>
    public class ValidateCallForSubmission : CodeActivity
    {
        [Output("Valid")]
        public OutArgument<bool> valid { get; set; }

        [Output("Validate Message")]
        public OutArgument<string> message { get; set; }

        protected override void Execute(CodeActivityContext executionContext)
        {
            ITracingService tracingService = executionContext.GetExtension<ITracingService>();

            IWorkflowContext context = executionContext.GetExtension<IWorkflowContext>();
            IOrganizationServiceFactory serviceFactory = executionContext.GetExtension<IOrganizationServiceFactory>();
            IOrganizationService service = serviceFactory.CreateOrganizationService(context.UserId);

            tracingService.Trace("{0}{1}", "Start Custom Workflow Activity: ValidateCallForSubmission", DateTime.Now.ToLongTimeString());

            var recordId = context.PrimaryEntityId;

            var callForSubmission = service.Retrieve(context.PrimaryEntityName, context.PrimaryEntityId, new ColumnSet("caps_callforsubmissiontype")) as caps_CallForSubmission;

            //VALIDATION STARTS HERE
            bool isValid = true;
            StringBuilder validationMessage = new StringBuilder();


            var fetchXML = "<fetch version = '1.0' output-format = 'xml-platform' mapping = 'logical' distinct = 'true' >" +
                            "<entity name = 'caps_submissioncategory' >" +
                                "<attribute name = 'caps_submissioncategoryid' />" +
                                "<attribute name = 'caps_name' />" +
                                "<attribute name = 'createdon' />" +
                                "<order attribute = 'caps_name' descending = 'false' />" +
                                "<filter type = 'and' >" +
                                    "<condition attribute = 'caps_callforsubmissiontype' operator= 'ne' value = '"+callForSubmission.caps_CallforSubmissionType.Value+"' />" +
                                "</filter>" +
                                "<link-entity name = 'caps_callforsubmission_caps_submissioncat' from = 'caps_submissioncategoryid' to = 'caps_submissioncategoryid' visible = 'false' intersect = 'true' >" +
                                    "<link-entity name = 'caps_callforsubmission' from = 'caps_callforsubmissionid' to = 'caps_callforsubmissionid' alias = 'ab' >" +
                                    "<filter type = 'and' >" +
                                        "<condition attribute = 'caps_callforsubmissionid' operator= 'eq' value = '{"+recordId+"}' />" +
                                    "</filter >" +
                                    "</link-entity >" +
                                "</link-entity >" +
                            "</entity >" +
                            "</fetch >";

            tracingService.Trace("Fetch XML: {0}", fetchXML);
            var submissionCategories = service.RetrieveMultiple(new FetchExpression(fetchXML));

            if (submissionCategories.Entities.Count() > 0)
            {
                isValid = false;
                validationMessage.AppendLine("One or more related submission categories is for the wrong type.");
            }

            //if this is an AFG project, confirm that the AFG %s for SDs = 100 

            if (callForSubmission.caps_CallforSubmissionType.Value == (int)caps_CallforSubmissionType.AFG)
            {
                using (var crmContext= new CrmServiceContext(service))
                {
                    var afgFundingTotal = crmContext.edu_schooldistrictSet.Where(r => r.StateCode == edu_schooldistrictState.Active).AsEnumerable().Sum(r=>r.caps_AFGFundingPercentage);

                    tracingService.Trace("AFG Total: {0}", afgFundingTotal);

                    if (afgFundingTotal != 1M)
                    {
                        isValid = false;
                        validationMessage.AppendLine("AFG total allocation does not equal 100%.");
                    }

                }
            }

            this.valid.Set(executionContext, isValid);
            this.message.Set(executionContext, validationMessage.ToString());
        }
    }
}
