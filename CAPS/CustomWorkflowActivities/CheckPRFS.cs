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
    /// Called via javascript on project request form when specific PRFS fields are changed.  This CWA returns true if there are expenditures in the first 3 years.
    /// </summary>
    public class CheckPRFS : CodeActivity
    {
        [Input("Capital Plan")]
        [ReferenceTarget("caps_submission")]
        public InArgument<EntityReference> capitalPlan { get; set; }

        [Output("Contains Expenditures in First 3 Years")]
        public OutArgument<bool> containsExpenditures { get; set; }

        protected override void Execute(CodeActivityContext executionContext)
        {
            ITracingService tracingService = executionContext.GetExtension<ITracingService>();

            IWorkflowContext context = executionContext.GetExtension<IWorkflowContext>();
            IOrganizationServiceFactory serviceFactory = executionContext.GetExtension<IOrganizationServiceFactory>();
            IOrganizationService service = serviceFactory.CreateOrganizationService(context.UserId);

            tracingService.Trace("{0}{1}", "Start Custom Workflow Activity: CheckPRFS", DateTime.Now.ToLongTimeString());

            var recordId = context.PrimaryEntityId;

            var projectRequestRecord = service.Retrieve(context.PrimaryEntityName, recordId, new ColumnSet("caps_projectrationale", "caps_scopeofwork", "caps_tempaccommodationandbusingplan", "caps_municipalrequirements")) as caps_Project;

            //tracingService.Trace("Project Rationale: {0}", projectRequestRecord.caps_ProjectRationale);
            //tracingService.Trace("Scope of Work: {0}", projectRequestRecord.caps_ScopeofWork);
            //tracingService.Trace("Temp Accomodation Plan: {0}", projectRequestRecord.caps_TempAccommodationandBusingPlan);
            //tracingService.Trace("Municipal Requirements: {0}", projectRequestRecord.caps_MunicipalRequirements);

            ////if the 4 PRFS fields are filled out then there is no need to go further
            //if (!string.IsNullOrWhiteSpace(projectRequestRecord.caps_ProjectRationale)
            //    && !string.IsNullOrWhiteSpace(projectRequestRecord.caps_ScopeofWork)
            //    && !string.IsNullOrWhiteSpace(projectRequestRecord.caps_TempAccommodationandBusingPlan)
            //    && !string.IsNullOrWhiteSpace(projectRequestRecord.caps_MunicipalRequirements))
            //{
            //    tracingService.Trace("{0}", "PRFS fields filled out, returned true.");
            //    this.valid.Set(executionContext, true);
            //    return;
            //}

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

            if (estimatedExpenditures.Entities.Count() < 1)
            {
                tracingService.Trace("{0}", "Doesn't contain expendiures, return false");
                this.containsExpenditures.Set(executionContext, false);
            }
            else
            {
                tracingService.Trace("{0}", "Contains expenditure, return true");
                this.containsExpenditures.Set(executionContext, true);
            }
        }
    }
}
