using Microsoft.Xrm.Sdk;
using Microsoft.Xrm.Sdk.Query;
using Microsoft.Xrm.Sdk.Workflow;
using System;
using System.Activities;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using CAPS.DataContext;

namespace CustomWorkflowActivities
{
    /// <summary>
    /// This CWA runs on a facility record.  It calculates historical, current and future capacities (design and operating)
    /// and updates the related Capacity Reporting records.  This data is used by the capacity reports.
    /// </summary>
    public class SetCapacityReporting : CodeActivity
    {
        protected override void Execute(CodeActivityContext executionContext)
        {
            ITracingService tracingService = executionContext.GetExtension<ITracingService>();

            IWorkflowContext context = executionContext.GetExtension<IWorkflowContext>();
            IOrganizationServiceFactory serviceFactory = executionContext.GetExtension<IOrganizationServiceFactory>();
            IOrganizationService service = serviceFactory.CreateOrganizationService(context.UserId);

            tracingService.Trace("{0}{1}", "Start Custom Workflow Activity: SetCapacityReporting", DateTime.Now.ToLongTimeString());

            var recordId = context.PrimaryEntityId;

            var enrolmentK = 0;
            var enrolmentE = 0;
            var enrolmentS = 0;

            //Get Facility record with all current design capacity fields & school district
            var columns = new ColumnSet("caps_adjusteddesigncapacityelementary",
                                    "caps_adjusteddesigncapacitykindergarten",
                                    "caps_adjusteddesigncapacitysecondary",
                                    "caps_schooldistrict",
                                    "caps_currentenrolment",
                                    "caps_lowestgrade",
                                    "caps_highestgrade");
            var facilityRecord = service.Retrieve(context.PrimaryEntityName, context.PrimaryEntityId, columns) as caps_Facility;

            var capacity = new Services.CapacityFactors();
            capacity.DesignKindergarten = GetBudgetCalculationValue(service, "Design Capacity Kindergarten");
            capacity.DesignElementary = GetBudgetCalculationValue(service, "Design Capacity Elementary");
            capacity.DesignSecondary = GetBudgetCalculationValue(service, "Design Capacity Secondary");

            capacity.OperatingKindergarten = GetBudgetCalculationValue(service, "Kindergarten Operating Capacity");
            capacity.OperatingLowElementary = GetBudgetCalculationValue(service, "Elementary Lower Operating Capacity");
            capacity.OperatingHighElementary = GetBudgetCalculationValue(service, "Elementary Upper Operating Capacity");
            capacity.OperatingSecondary = GetBudgetCalculationValue(service, "Secondary Operating Capacity");

            Services.OperatingCapacity capacityService = new Services.OperatingCapacity(service, tracingService, capacity);

            tracingService.Trace("Line: {0}", "40");
            if (facilityRecord.caps_CurrentEnrolment != null)
            {
                var enrolmentRecord = service.Retrieve(caps_facilityenrolment.EntityLogicalName, facilityRecord.caps_CurrentEnrolment.Id, new ColumnSet("caps_sumofkindergarten", "caps_sumofelementary", "caps_sumofsecondary")) as caps_facilityenrolment;

                enrolmentK = enrolmentRecord.caps_SumofKindergarten.GetValueOrDefault(0);
                enrolmentE = enrolmentRecord.caps_SumofElementary.GetValueOrDefault(0);
                enrolmentS = enrolmentRecord.caps_SumofSecondary.GetValueOrDefault(0);

                tracingService.Trace("Line: {0}", "49");
            }

            //Get Curretn Enrolment Projections
            QueryExpression enrolmentQuery = new QueryExpression("caps_enrolmentprojections_sd");
            enrolmentQuery.ColumnSet.AllColumns = true;
            enrolmentQuery.Criteria = new FilterExpression();
            enrolmentQuery.Criteria.AddCondition("caps_facility", ConditionOperator.Equal, recordId);
            enrolmentQuery.Criteria.AddCondition("statuscode", ConditionOperator.Equal, 200870000);
            EntityCollection enrolmentProjectionRecords = service.RetrieveMultiple(enrolmentQuery);
            var enrolmentProjectionList = enrolmentProjectionRecords.Entities.Select(r => r.ToEntity<caps_EnrolmentProjections_SD>()).ToList();


            //Get Projects for this facility with occupancy dates in the future
            QueryExpression projectQuery = new QueryExpression("caps_projecttracker");
            projectQuery.ColumnSet.AddColumns("caps_designcapacitykindergarten", "caps_designcapacityelementary", "caps_designcapacitysecondary");
            projectQuery.Criteria = new FilterExpression();
            projectQuery.Criteria.AddCondition("caps_facility", ConditionOperator.Equal, recordId);
            projectQuery.Criteria.AddCondition("statecode", ConditionOperator.Equal, 0);
            LinkEntity milestoneLink = projectQuery.AddLink("caps_projectmilestone", "caps_projecttrackerid", "caps_projecttracker", JoinOperator.Inner);
            milestoneLink.Columns.AddColumn("caps_expectedactualdate");
            milestoneLink.EntityAlias = "milestone";
            milestoneLink.LinkCriteria = new FilterExpression();
            milestoneLink.LinkCriteria.AddCondition("caps_name", ConditionOperator.Equal, "Occupancy");
            milestoneLink.LinkCriteria.AddCondition("caps_expectedactualdate", ConditionOperator.OnOrAfter, DateTime.Now);

            EntityCollection projectRecords = service.RetrieveMultiple(projectQuery);

            

            //Get Capacity Reporting Records
            QueryExpression query = new QueryExpression("caps_capacityreporting");
            query.ColumnSet.AllColumns = true;
            LinkEntity yearLink = query.AddLink("edu_year", "caps_schoolyear", "edu_yearid");
            yearLink.Columns.AddColumns("edu_yearid", "statuscode", "edu_enddate");
            yearLink.EntityAlias = "year";
            query.Criteria = new FilterExpression();
            query.Criteria.AddCondition("year", "edu_type", ConditionOperator.Equal, 757500001);
            query.Criteria.AddCondition("statecode", ConditionOperator.Equal, 0);
            query.Criteria.AddCondition("caps_facility", ConditionOperator.Equal, recordId);
            EntityCollection reportingRecords = service.RetrieveMultiple(query);

            foreach (var reportingRecord in reportingRecords.Entities)
            {
                tracingService.Trace("Line: {0}", "71");
                //Get Type - Historical, Current, Future
                var recordType = (OptionSetValue)((AliasedValue)reportingRecord["year.statuscode"]).Value;

                if (recordType.Value == 757500001)
                {
                    tracingService.Trace("Line: {0}", "77");
                    //Historical
                    #region Historical

                    //For each record, get top 1 Facility History record where year snapshot equals the school year record, order by created on date descending
                    tracingService.Trace("Year Value: {0}", ((AliasedValue)reportingRecord["year.edu_yearid"]).Value);
                    var schoolYearRecord = ((AliasedValue)reportingRecord["year.edu_yearid"]).Value;

                    QueryExpression snapshotQuery = new QueryExpression("caps_facilityhistory");
                    snapshotQuery.ColumnSet.AllColumns = true;

                    snapshotQuery.Criteria = new FilterExpression();
                    snapshotQuery.Criteria.AddCondition("caps_facility", ConditionOperator.Equal, recordId);
                    snapshotQuery.Criteria.AddCondition("caps_year", ConditionOperator.Equal, schoolYearRecord);
                    snapshotQuery.AddOrder("createdon", OrderType.Descending);
                    snapshotQuery.TopCount = 1;

                    EntityCollection snapshotRecords = service.RetrieveMultiple(snapshotQuery);

                    //Get top record
                    if (snapshotRecords.Entities.Count() == 1)
                    {
                        tracingService.Trace("Line: {0}", "93");
                        var topRecord = snapshotRecords.Entities[0] as caps_FacilityHistory;

                        //Update the capacity record
                        var recordToUpdate = new caps_CapacityReporting();
                        recordToUpdate.Id = reportingRecord.Id;

                        recordToUpdate.caps_Kindergarten_designcapacity = topRecord.caps_DesignCapacityKindergarten;
                        recordToUpdate.caps_Elementary_designcapacity = topRecord.caps_DesignCapacityElementary;
                        recordToUpdate.caps_Secondary_designcapacity = topRecord.caps_DesignCapacitySecondary;

                        recordToUpdate.caps_Kindergarten_operatingcapacity = topRecord.caps_OperatingCapacityKindergarten;
                        recordToUpdate.caps_Elementary_operatingcapacity = topRecord.caps_OperatingCapacityElementary;
                        recordToUpdate.caps_Secondary_operatingcapacity = topRecord.caps_OperatingCapacitySecondary;

                        recordToUpdate.caps_Kindergarten_enrolment = 0;
                        recordToUpdate.caps_Elementary_enrolment = 0;
                        recordToUpdate.caps_Secondary_enrolment = 0;

                        service.Update(recordToUpdate);
                    } 
                    #endregion
                }
                else if (recordType.Value == 1)
                {
                    //Current
                    #region Current
                    tracingService.Trace("Line: {0}", "123");

                    //Update the capacity record
                    var recordToUpdate = new caps_CapacityReporting();
                    recordToUpdate.Id = reportingRecord.Id;

                    recordToUpdate.caps_Kindergarten_designcapacity = facilityRecord.caps_AdjustedDesignCapacityKindergarten;
                    recordToUpdate.caps_Elementary_designcapacity = facilityRecord.caps_AdjustedDesignCapacityElementary;
                    recordToUpdate.caps_Secondary_designcapacity = facilityRecord.caps_AdjustedDesignCapacitySecondary;

                    //Note: Operating capacity is kept up to date on the facility record so no calculation is needed
                    recordToUpdate.caps_Kindergarten_operatingcapacity = facilityRecord.caps_OperatingCapacityKindergarten;
                    recordToUpdate.caps_Elementary_operatingcapacity = facilityRecord.caps_OperatingCapacityElementary;
                    recordToUpdate.caps_Secondary_operatingcapacity = facilityRecord.caps_OperatingCapacitySecondary;

                    recordToUpdate.caps_Kindergarten_enrolment = enrolmentK;
                    recordToUpdate.caps_Elementary_enrolment = enrolmentE;
                    recordToUpdate.caps_Secondary_enrolment = enrolmentS;

                    service.Update(recordToUpdate); 
                    #endregion
                }
                else {
                    //Future
                    #region Future
                    tracingService.Trace("Line: {0}", "144");

                    var endDate = (DateTime?)((AliasedValue)reportingRecord["year.edu_enddate"]).Value;

                    var recordToUpdate = new caps_CapacityReporting();
                    recordToUpdate.Id = reportingRecord.Id;

                    //get matching enrolment projection record
                    var matchingProjection = enrolmentProjectionList.FirstOrDefault(r => r.caps_ProjectionYear.Id == reportingRecord.GetAttributeValue<EntityReference>("caps_schoolyear").Id);

                    var kDesign = facilityRecord.caps_AdjustedDesignCapacityKindergarten.GetValueOrDefault(0);
                    var eDesign = facilityRecord.caps_AdjustedDesignCapacityElementary.GetValueOrDefault(0);
                    var sDesign = facilityRecord.caps_AdjustedDesignCapacitySecondary.GetValueOrDefault(0);

                    if (endDate.HasValue)
                    {
                        //loop project records
                        foreach(var projectRecord in projectRecords.Entities)
                        {
                            if ((DateTime?)((AliasedValue)projectRecord["milestone.caps_expectedactualdate"]).Value <= endDate)
                            {
                                //add design capacity on
                                kDesign += Convert.ToInt32(projectRecord.GetAttributeValue<decimal?>("caps_designcapacitykindergarten").GetValueOrDefault(0));
                                eDesign += Convert.ToInt32(projectRecord.GetAttributeValue<decimal?>("caps_designcapacityelementary").GetValueOrDefault(0));
                                sDesign += Convert.ToInt32(projectRecord.GetAttributeValue<decimal?>("caps_designcapacitysecondary").GetValueOrDefault(0));
                            }
                        }
                    }
                    var lowestGrade = facilityRecord.caps_LowestGrade.Value;
                    var highestGrade = facilityRecord.caps_HighestGrade.Value;

                    recordToUpdate.caps_Kindergarten_designcapacity = kDesign;
                    recordToUpdate.caps_Elementary_designcapacity = eDesign;
                    recordToUpdate.caps_Secondary_designcapacity = sDesign;

                    //Get operating capacity for the design numbers
                    var result = capacityService.Calculate(kDesign, eDesign, sDesign, lowestGrade, highestGrade);

                    recordToUpdate.caps_Kindergarten_operatingcapacity = result.KindergartenCapacity;
                    recordToUpdate.caps_Elementary_operatingcapacity = result.ElementaryCapacity;
                    recordToUpdate.caps_Secondary_operatingcapacity = result.SecondaryCapacity;

                    recordToUpdate.caps_Kindergarten_enrolment = (matchingProjection != null) ? matchingProjection.caps_EnrolmentProjectionKindergarten : enrolmentK;
                    recordToUpdate.caps_Elementary_enrolment = (matchingProjection != null) ? matchingProjection.caps_EnrolmentProjectionElementary : enrolmentE;
                    recordToUpdate.caps_Secondary_enrolment = (matchingProjection != null) ? matchingProjection.caps_EnrolmentProjectionSecondary : enrolmentS;

                    service.Update(recordToUpdate); 
                    #endregion
                }
            }

        }

        private decimal GetBudgetCalculationValue(IOrganizationService service, string name)
        {

            FilterExpression filterName = new FilterExpression();
            filterName.Conditions.Add(new ConditionExpression("caps_name", ConditionOperator.Equal, name));

            QueryExpression query = new QueryExpression("caps_budgetcalc_value");
            query.ColumnSet.AddColumns("caps_value");
            query.Criteria.AddFilter(filterName);

            EntityCollection results = service.RetrieveMultiple(query);

            if (results.Entities.Count != 1) throw new Exception("Missing Budget Calculation Value: " + name);

            return results.Entities[0].GetAttributeValue<decimal>("caps_value");
        }
    }
}
