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
    /// Called via javascript on project request form when specific PRFS fields are changed.  
    /// This CWA returns true if there are expenditures in the first 3 years and there are no Surrounding School records.
    /// </summary>
    public class CheckPRFSSurroundingSchools : CodeActivity
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

            if (estimatedExpenditures.Entities.Count() > 0)
            {
                //check if there is at least one surrounding school
                var fetchFacilties = "<fetch version=\"1.0\" output-format=\"xml-platform\" mapping=\"logical\" distinct=\"true\">"+
                                        "<entity name=\"caps_facility\">"+
                                           "<attribute name=\"caps_facilityid\" /> "+
                                            "<attribute name =\"caps_name\" /> "+
                                             "<attribute name = \"createdon\" /> "+
                                              "<order attribute = \"caps_name\" descending = \"false\" /> "+
                                                 "<link-entity name = \"caps_project_caps_facilityss\" from = \"caps_facilityid\" to = \"caps_facilityid\" visible = \"false\" intersect = \"true\" > "+
                                                              "<link-entity name = \"caps_project\" from = \"caps_projectid\" to = \"caps_projectid\" alias=\"ad\" > "+
                                                     "<filter type = \"and\" > "+
                                                     "<condition attribute = 'statecode' operator= 'eq' value = '0' />" +
                                                        "<condition attribute = \"caps_projectid\" operator= \"eq\"  value = \"{" +recordId+"}\" /> "+
                                                              "</filter > " +
                                                            "</link-entity > " +
                                                          "</link-entity > " +
                                                        "</entity> " +
                                                      "</fetch> ";

                tracingService.Trace("Fetch: {0}", fetchFacilties);
                var surroundingSchools = service.RetrieveMultiple(new FetchExpression(fetchFacilties));

                if (surroundingSchools.Entities.Count() > 0)
                {
                    tracingService.Trace("{0}", "Surrounding School records found");
                    this.displayError.Set(executionContext, false);
                }
                else
                {
                    tracingService.Trace("{0}", "No Surrounding School records found");
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
