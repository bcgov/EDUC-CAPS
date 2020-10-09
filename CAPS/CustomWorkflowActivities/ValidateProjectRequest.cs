﻿using CAPS.DataContext;
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
    /// This CWA validates the project request when triggered, when the project request is added to a capital plan and before the capital plan is submitted.
    /// </summary>
    public class ValidateProjectRequest : CodeActivity
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

            tracingService.Trace("{0}{1}", "Start Custom Workflow Activity: ValidateProjectRequest", DateTime.Now.ToLongTimeString());

            var recordId = context.PrimaryEntityId;

            tracingService.Trace("{0}", "Loading data");
            //get Project Request
            var projectRequest = service.Retrieve(context.PrimaryEntityName, context.PrimaryEntityId, new ColumnSet(true)) as caps_Project;

            //get Submission Category
            var submissionCategory = service.Retrieve(projectRequest.caps_SubmissionCategory.LogicalName, projectRequest.caps_SubmissionCategory.Id, new ColumnSet(true)) as caps_SubmissionCategory;

            caps_Submission capitalPlan = new caps_Submission();
            caps_CallForSubmission callForSubmission = new caps_CallForSubmission();

            if (projectRequest.caps_Submission != null)
            {
                //get Capital Plan
                capitalPlan = service.Retrieve(projectRequest.caps_Submission.LogicalName, projectRequest.caps_Submission.Id, new ColumnSet("caps_callforsubmission")) as caps_Submission;

                //get Call for Submission
                callForSubmission = service.Retrieve(capitalPlan.caps_CallforSubmission.LogicalName, capitalPlan.caps_CallforSubmission.Id, new ColumnSet("caps_capitalplanyear")) as caps_CallForSubmission;
            }

            //VALIDATION STARTS HERE
            bool isValid = true;
            StringBuilder validationMessage = new StringBuilder();

            decimal totalProjectCost = projectRequest.caps_TotalProjectCost.GetValueOrDefault(0);

            #region Check if Published
            //check if published
            if (!projectRequest.caps_PublishProject.GetValueOrDefault(false))
            {
                isValid = false;
                validationMessage.AppendLine("Project Request is a part of a capital plan, it needs to be published in order to be submitted.");
            } 
            #endregion

            tracingService.Trace("{0}", "Run Schedule B");
            #region Run Schedule B
            //Run Schedule B for any project requests that require it
            if (projectRequest.caps_RequiresScheduleB.GetValueOrDefault(false))
            {
                //Run Schedule B
                try
                {
                    totalProjectCost = CalculateScheduleB.RunCalculation(tracingService, context, service, recordId);

                    if (projectRequest.caps_TotalProjectCost != totalProjectCost)
                    {
                        //need to update project request
                        var recordToUpdate = new caps_Project();
                        recordToUpdate.Id = projectRequest.Id;
                        recordToUpdate.caps_TotalProjectCost = totalProjectCost;
                        recordToUpdate.caps_ScheduleBTotal = totalProjectCost;
                        service.Update(recordToUpdate);
                    }
                }
                catch(Exception ex)
                {
                    isValid = false;
                    validationMessage.AppendLine("Schedule B failed to run sucessfully.  Please manually trigger the process by clicking the button to see the detailed error.");
                }
            }
            #endregion

            tracingService.Trace("{0}", "Check Cash Fully Allocated");
            #region Check Cash Fully Allocated
            //Check that cash flow is fully allocated if needed
            if (submissionCategory.caps_RequireCostAllocation.GetValueOrDefault(false))
            {
                //Get Estimated Expenditures
                using (var crmContext = new CrmServiceContext(service))
                {
                    decimal? estimatedExpenditures = crmContext.caps_EstimatedYearlyCapitalExpenditureSet.Where(r => r.StateCode == caps_EstimatedYearlyCapitalExpenditureState.Active && r.caps_Project.Id == recordId).AsEnumerable().Sum(r => r.caps_YearlyExpenditure);

                    if (Math.Round(totalProjectCost) != estimatedExpenditures)
                    {
                        tracingService.Trace("Total Project Cost: {0}", totalProjectCost);
                        tracingService.Trace("Total Estimated Expenditures: {0}", estimatedExpenditures);
                        isValid = false;
                        validationMessage.AppendLine("Total Project Cost is not fully allocated.");
                    }
                }
            }
            #endregion

            tracingService.Trace("{0}", "Facility is BEP Eligible");
            #region Check if Facility is BEP Eligible
            if (submissionCategory.caps_CategoryCode == "BEP")
            {
                //check if related facility still set as BEP Eligible
                var facility = service.Retrieve(projectRequest.caps_Facility.LogicalName, projectRequest.caps_Facility.Id, new ColumnSet("caps_bepeligible")) as caps_Facility;

                if (!facility.caps_BEPEligible.GetValueOrDefault(false))
                {
                    isValid = false;
                    validationMessage.AppendLine("This BEP project request is for a facility that is not eligible.");
                }

            }
            #endregion
            
            #region Check if BUS is eligible for replacement
            //Check if BUS is eligible for replacement
            if (submissionCategory.caps_CategoryCode == "BUS"
                && projectRequest.caps_bus != null)
            {
                //get related bus record
                var bus = service.Retrieve(projectRequest.caps_bus.LogicalName, projectRequest.caps_bus.Id, new ColumnSet("caps_nonreplaceable")) as caps_Bus;

                if (bus.caps_NonReplaceable.GetValueOrDefault(false))
                {
                    isValid = false;
                    validationMessage.AppendLine("This Bus project request is for a bus that is not eligible for replacement.");
                }
            }
            #endregion

            #region Check if SEP and CNCP have at least one Facility selected
            if (submissionCategory.caps_CategoryCode == "CNCP"
        || submissionCategory.caps_CategoryCode == "SEP")
            {
                if (projectRequest.caps_MultipleFacilities.GetValueOrDefault(false))
                {
                    //check that there is at least 1 related facility
                    QueryExpression query = new QueryExpression("caps_project_caps_facility");
                    //query.ColumnSet.AddColumns(true);
                    query.Criteria = new FilterExpression();
                    query.Criteria.AddCondition(caps_Project_caps_Facility.Fields.caps_projectid, ConditionOperator.Equal, recordId);

                    EntityCollection relatedFacilities = service.RetrieveMultiple(query);

                    if (relatedFacilities.Entities.Count() < 1)
                    {
                        isValid = false;
                        validationMessage.AppendLine("Project request is part of a capital plan, it needs at least one facility in order to be submitted.");
                    }
                }
            } 
            #endregion

            tracingService.Trace("{0}", "Minor & AFG in correct year");
            #region Check if minor project year matches call for submission
            //Check if minor projects' project year matches call for submission
            if (submissionCategory.caps_CallforSubmissionType.Value == (int)caps_CallforSubmissionType.Minor
                || submissionCategory.caps_CallforSubmissionType.Value == (int)caps_CallforSubmissionType.AFG)
            {
                if (projectRequest.caps_Submission != null
                    && projectRequest.caps_Projectyear.Id != callForSubmission.caps_CapitalPlanYear.Id)
                {
                    tracingService.Trace("Project Year: {0}", projectRequest.caps_Projectyear.Id);
                    tracingService.Trace("Capital Plan Year: {0}", callForSubmission.caps_CapitalPlanYear.Id);
                    //Project request is not in correct year
                    isValid = false;
                    validationMessage.AppendLine("Minor project requests added to a capital plan must start in same year in order to be submitted.");
                }
            }
            #endregion

            #region Check if Major Project has cash flow in first 5 years
            //Check if Major Project has cash flow in first 5 years
            if (submissionCategory.caps_type.Value == (int)caps_submissioncategory_type.Major)
            {
                if (projectRequest.caps_Submission != null
                    && projectRequest.caps_Projectyear != callForSubmission.caps_CapitalPlanYear)
                {
                    //get capital plan year
                    var capitalPlanYear = service.Retrieve(callForSubmission.caps_CapitalPlanYear.LogicalName, callForSubmission.caps_CapitalPlanYear.Id, new ColumnSet("edu_startyear")) as edu_Year;

                    int startYear = capitalPlanYear.edu_StartYear.Value;

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
                                                "<condition attribute = 'edu_startyear' operator= 'le' value = '" + (startYear + 4) + "' />" +
                                                "<condition attribute = 'edu_type' operator= 'eq' value = '757500000' />" +
                                            "</filter></link-entity></entity></fetch>";

                    //Find out if there is cash flow in any of those 5 years
                    var estimatedExpenditures = service.RetrieveMultiple(new FetchExpression(fetchXML));

                    if (estimatedExpenditures.Entities.Count() < 1)
                    {
                        isValid = false;
                        validationMessage.AppendLine("There are no estimated yearly expenditures in the first 5 years.  Please adjust or remove the project request from the capital plan.");
                    }
                }
            } 
            #endregion

            //Check if major project has occupancy year before anticipated start year

            if (submissionCategory.caps_type.Value == (int)caps_submissioncategory_type.Major 
                && projectRequest.caps_AnticipatedOccupancyYear != null
                && projectRequest.caps_Projectyear != null)
            {
                //get both anticipated occupancy year and project year
                var anticipatedOccupancyYear = service.Retrieve(projectRequest.caps_AnticipatedOccupancyYear.LogicalName, projectRequest.caps_AnticipatedOccupancyYear.Id, new ColumnSet("edu_startyear")) as edu_Year;

                var anticipatedStartYear = service.Retrieve(projectRequest.caps_Projectyear.LogicalName, projectRequest.caps_Projectyear.Id, new ColumnSet("edu_startyear")) as edu_Year;
                
                if (anticipatedOccupancyYear.edu_StartYear < anticipatedStartYear.edu_StartYear)
                {
                    isValid = false;
                    validationMessage.AppendLine("Anticipated Occupancy Year should be equal or After Anticipated Project Start Year.");
                }
                  
            }

            this.valid.Set(executionContext, isValid);
            this.message.Set(executionContext, validationMessage.ToString());
        }
    }
}