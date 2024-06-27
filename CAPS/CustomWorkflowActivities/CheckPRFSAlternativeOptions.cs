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
using System.Windows;
namespace CustomWorkflowActivities
{
    /// <summary>
    /// Called via javascript on project request form when specific PRFS fields are changed.  
    /// This CWA returns true if there are expenditures in the first 3 years and there are no related PRFS Alternate Option records.
    /// </summary>
    public class CheckPRFSAlternativeOptions : CodeActivity
    {
        [Input("Capital Plan")]
        [ReferenceTarget("caps_submission")]
        public InArgument<EntityReference> capitalPlan { get; set; }

        [Output("Display Error")]
        public OutArgument<bool> displayError { get; set; }

        protected override void Execute(CodeActivityContext executionContext)
        {
            ITracingService tracingService = executionContext.GetExtension<ITracingService>();

            IWorkflowContext context = executionContext.GetExtension<IWorkflowContext>();
            IOrganizationServiceFactory serviceFactory = executionContext.GetExtension<IOrganizationServiceFactory>();
            IOrganizationService service = serviceFactory.CreateOrganizationService(context.UserId);

            tracingService.Trace("{0}{1}", "Start Custom Workflow Activity: CheckPRFSSurroundingSchools", DateTime.Now.ToLongTimeString());

            var recordId = context.PrimaryEntityId;

            var projectRequestRecord = service.Retrieve(context.PrimaryEntityName, recordId, new ColumnSet("caps_municipalrequirements", "caps_submissioncategorycode")) as caps_Project;

            EntityReference capitalPlanReference = this.capitalPlan.Get(executionContext);

            var capitalPlan = service.Retrieve(capitalPlanReference.LogicalName, capitalPlanReference.Id, new ColumnSet("caps_callforsubmission")) as caps_Submission;

            var callForSubmission = service.Retrieve(capitalPlan.caps_CallforSubmission.LogicalName, capitalPlan.caps_CallforSubmission.Id, new ColumnSet("caps_capitalplanyear")) as caps_CallForSubmission;

            var capitalPlanYear = service.Retrieve(callForSubmission.caps_CapitalPlanYear.LogicalName, callForSubmission.caps_CapitalPlanYear.Id, new Microsoft.Xrm.Sdk.Query.ColumnSet("edu_startyear")) as edu_Year;

            var startYear = capitalPlanYear.edu_StartYear.Value;

            var fetchXML = "<fetch version = '1.0' output-format = 'xml-platform' mapping = 'logical' distinct = 'false' >" +
                            "<entity name = 'caps_estimatedyearlycapitalexpenditure' >" +
                                "<attribute name = 'caps_estimatedyearlycapitalexpenditureid' />" +
                                "<attribute name = 'caps_name' />" +
                                "<attribute name = 'createdon' />" +
                                "<order attribute = 'caps_name' descending = 'false' />" +
                                "<filter type = 'and' >" +
                                    "<condition attribute = 'statecode' operator= 'eq' value = '0' />" +
                                    "<condition attribute = 'caps_project' operator= 'eq' value = '{" + recordId + "}' />" +
                                    "<condition attribute = 'caps_yearlyexpenditure' operator='not-null' />" +
                                    "<condition attribute = 'caps_yearlyexpenditure' operator='gt' value='0' />" +
                                "</filter>" +
                                "<link-entity name='edu_year' from='edu_yearid' to='caps_year' link-type='inner' alias='ad' >" +
                                    "<filter type = 'and' >" +
                                        "<condition attribute = 'edu_startyear' operator= 'ge' value='" + startYear + "' />" +
                                        "<condition attribute = 'edu_startyear' operator= 'le' value = '" + (startYear + 2) + "' />" +
                                        "<condition attribute = 'edu_type' operator= 'eq' value = '757500000' />" +
                                    "</filter></link-entity></entity></fetch>";

            tracingService.Trace("{0}", "Calling Fetch");
            //Find out if there is cash flow in any of those 5 years
            var estimatedExpenditures = service.RetrieveMultiple(new FetchExpression(fetchXML));

            if (estimatedExpenditures.Entities.Count() > 0 || projectRequestRecord.caps_SubmissionCategoryCode == "LEASE")
            {
                //check if there is at least one alternative option
                var fetchAlterativeOptions = "<fetch version=\"1.0\" output-format=\"xml-platform\" mapping=\"logical\" distinct=\"false\">"+
                                    "<entity name=\"caps_prfsalternativeoption\" > "+
                                       "<attribute name = \"caps_prfsalternativeoptionid\" /> "+
                                        "<attribute name = \"caps_name\" /> "+
                                         "<attribute name = \"createdon\" /> "+
                                          "<order attribute = \"caps_name\" descending = \"false\" /> "+
                                            "<filter type = \"and\" > "+
                                                "<condition attribute = \"statuscode\" operator= \"in\" >" +
                                                    "<value>1</value>" +
                                                     "<value>200870000</value>" +
                                                "</condition >" +
                                                "<condition attribute = \"caps_projectrequest\" operator= \"eq\" value = \"{" + recordId+"}\" /> "+
                                                        "</filter>" +
                                                    "</entity> " +
                                             "</fetch>";

                tracingService.Trace("Fetch: {0}", fetchAlterativeOptions);

                var alternativeOptions = service.RetrieveMultiple(new FetchExpression(fetchAlterativeOptions));

                if (alternativeOptions.Entities.Count() > 0)
                {
                    tracingService.Trace("{0}", "Alternative Options records found");
                    this.displayError.Set(executionContext, false);
                }
                else
                {
                    tracingService.Trace("{0}", "No Alternative Options records found");
                    this.displayError.Set(executionContext, true);
                }
            }
            else
            {

                tracingService.Trace("{0}", "No cashflow in the first 3 years");
                this.displayError.Set(executionContext, false);
            }
        }
    }
}
