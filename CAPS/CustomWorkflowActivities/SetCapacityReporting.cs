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

            decimal? enrolmentK = 0;
            decimal? enrolmentE = 0;
            decimal? enrolmentS = 0;

            //Get Facility record with all current design capacity fields & school district
            var columns = new ColumnSet("caps_adjusteddesigncapacityelementary",
                                    "caps_adjusteddesigncapacitykindergarten",
                                    "caps_adjusteddesigncapacitysecondary",
                                    "caps_schooldistrict",
                                    "caps_currentenrolment",
                                    "caps_lowestgrade",
                                    "caps_highestgrade",
                                    "caps_name",
                                    "caps_districtoperatingcapacitykindergarten",
                                    "caps_districtoperatingcapacityelementary",
                                    "caps_districtoperatingcapacitysecondary",
                                    "caps_operatingcapacitykindergarten",
                                    "caps_operatingcapacityelementary",
                                    "caps_operatingcapacitysecondary");
            var facilityRecord = service.Retrieve(context.PrimaryEntityName, context.PrimaryEntityId, columns) as caps_Facility;

            tracingService.Trace("Facility Name:{0}", facilityRecord.caps_Name);

            var capacity = new Services.CapacityFactors(service);

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

            //Get Current Enrolment Projections
            QueryExpression enrolmentQuery = new QueryExpression("caps_enrolmentprojections_sd");
            enrolmentQuery.ColumnSet.AllColumns = true;
            enrolmentQuery.Criteria = new FilterExpression();
            enrolmentQuery.Criteria.AddCondition("caps_facility", ConditionOperator.Equal, recordId);
            enrolmentQuery.Criteria.AddCondition("statuscode", ConditionOperator.Equal, 200870000);
            EntityCollection enrolmentProjectionRecords = service.RetrieveMultiple(enrolmentQuery);
            var enrolmentProjectionList = enrolmentProjectionRecords.Entities.Select(r => r.ToEntity<caps_EnrolmentProjections_SD>()).ToList();


            //Get Projects for this facility with occupancy dates in the future
            QueryExpression projectQuery = new QueryExpression("caps_projecttracker");
            projectQuery.ColumnSet.AddColumns("caps_designcapacitykindergarten", "caps_designcapacityelementary", "caps_designcapacitysecondary", "caps_districtoperatingcapacitykindergarten", "caps_districtoperatingcapacityelementary", "caps_districtoperatingcapacitysecondary");
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
                    //Historical
                    #region Historical

                    //For each record, get top 1 Facility History record where year snapshot equals the school year record, order by created on date descending
                    tracingService.Trace("Year Value: {0}", ((AliasedValue)reportingRecord["year.edu_yearid"]).Value);
                    var schoolYearRecord = ((AliasedValue)reportingRecord["year.edu_yearid"]).Value;

                    QueryExpression snapshotQuery = new QueryExpression("caps_facilityhistory");
                    snapshotQuery.ColumnSet.AllColumns = true;
                    LinkEntity enrolmentLink = snapshotQuery.AddLink("caps_facilityenrolment", "caps_facilityenrolment", "caps_facilityenrolmentid");
                    enrolmentLink.EntityAlias = "enrolment";
                    enrolmentLink.Columns.AddColumns("caps_sumofkindergarten", "caps_sumofelementary", "caps_sumofsecondary");

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

                        decimal? historicalEnrolmentK = (!reportingRecord.Contains("enrolment.caps_sumofkindergarten")) ? 0 : ((int?)((AliasedValue)reportingRecord["enrolment.caps_sumofkindergarten"]).Value).GetValueOrDefault(0);
                        decimal? historicalEnrolmentE = (!reportingRecord.Contains("enrolment.caps_sumofelementary")) ? 0 : ((int?)((AliasedValue)reportingRecord["enrolment.caps_sumofelementary"]).Value).GetValueOrDefault(0);
                        decimal? historicalEnrolmentS = (!reportingRecord.Contains("enrolment.caps_sumofsecondary")) ? 0 : ((int?)((AliasedValue)reportingRecord["enrolment.caps_sumofsecondary"]).Value).GetValueOrDefault(0);

                        recordToUpdate.caps_Kindergarten_enrolment = historicalEnrolmentK;
                        recordToUpdate.caps_Elementary_enrolment = historicalEnrolmentE;
                        recordToUpdate.caps_Secondary_enrolment = historicalEnrolmentS;

                        recordToUpdate.caps_Kindergarten_designcapacity = topRecord.caps_DesignCapacityKindergarten;
                        recordToUpdate.caps_Elementary_designcapacity = topRecord.caps_DesignCapacityElementary;
                        recordToUpdate.caps_Secondary_designcapacity = topRecord.caps_DesignCapacitySecondary;

                        recordToUpdate.caps_Kindergarten_designutilization = (topRecord.caps_DesignCapacityKindergarten == 0) ? 0 : historicalEnrolmentK / topRecord.caps_DesignCapacityKindergarten * 100;
                        recordToUpdate.caps_Elementary_designutilization = (topRecord.caps_DesignCapacityElementary == 0) ? 0 : historicalEnrolmentE / topRecord.caps_DesignCapacityElementary * 100;
                        recordToUpdate.caps_Secondary_designutilization = (topRecord.caps_DesignCapacitySecondary == 0) ? 0 : historicalEnrolmentS / topRecord.caps_DesignCapacitySecondary * 100;

                        recordToUpdate.caps_Kindergarten_operatingcapacity = topRecord.caps_OperatingCapacityKindergarten;
                        recordToUpdate.caps_Elementary_operatingcapacity = topRecord.caps_OperatingCapacityElementary;
                        recordToUpdate.caps_Secondary_operatingcapacity = topRecord.caps_OperatingCapacitySecondary;

                        recordToUpdate.caps_Kindergarten_operatingutilization = (topRecord.caps_OperatingCapacityKindergarten == 0) ? 0 : historicalEnrolmentK / topRecord.caps_OperatingCapacityKindergarten * 100;
                        recordToUpdate.caps_Elementary_operatingutilization = (topRecord.caps_OperatingCapacityElementary == 0) ? 0 : historicalEnrolmentE / topRecord.caps_OperatingCapacityElementary * 100;
                        recordToUpdate.caps_Secondary_operatingutilization = (topRecord.caps_OperatingCapacitySecondary == 0) ? 0 : historicalEnrolmentS / topRecord.caps_OperatingCapacitySecondary * 100;

                        recordToUpdate.caps_kindergarten_operatingcapacity_district = topRecord.caps_DistrictOperatingCapacityKindergarten;
                        recordToUpdate.caps_elementary_operatingcapacity_district = topRecord.caps_DistrictOperatingCapacityElementary;
                        recordToUpdate.caps_secondary_operatingcapacity_district = topRecord.caps_DistrictOperatingCapacitySecondary;

                        recordToUpdate.caps_kindergarten_operatingutilizatio_district = (topRecord.caps_DistrictOperatingCapacityKindergarten == 0) ? 0 : historicalEnrolmentK / topRecord.caps_DistrictOperatingCapacityKindergarten * 100;
                        recordToUpdate.caps_elementary_operatingutilization_district = (topRecord.caps_DistrictOperatingCapacityElementary == 0) ? 0 : historicalEnrolmentE / topRecord.caps_DistrictOperatingCapacityElementary * 100;
                        recordToUpdate.caps_secondary_operatingutilization_district = (topRecord.caps_DistrictOperatingCapacitySecondary == 0) ? 0 : historicalEnrolmentS / topRecord.caps_DistrictOperatingCapacitySecondary * 100;

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

                    recordToUpdate.caps_Kindergarten_enrolment = enrolmentK;
                    recordToUpdate.caps_Elementary_enrolment = enrolmentE;
                    recordToUpdate.caps_Secondary_enrolment = enrolmentS;

                    recordToUpdate.caps_Kindergarten_designcapacity = facilityRecord.caps_AdjustedDesignCapacityKindergarten;
                    recordToUpdate.caps_Elementary_designcapacity = facilityRecord.caps_AdjustedDesignCapacityElementary;
                    recordToUpdate.caps_Secondary_designcapacity = facilityRecord.caps_AdjustedDesignCapacitySecondary;

                    recordToUpdate.caps_Kindergarten_designutilization = (facilityRecord.caps_AdjustedDesignCapacityKindergarten == 0) ? 0 : enrolmentK / facilityRecord.caps_AdjustedDesignCapacityKindergarten * 100;
                    recordToUpdate.caps_Elementary_designutilization = (facilityRecord.caps_AdjustedDesignCapacityElementary == 0) ? 0 : enrolmentE / facilityRecord.caps_AdjustedDesignCapacityElementary * 100;
                    recordToUpdate.caps_Secondary_designutilization = (facilityRecord.caps_AdjustedDesignCapacitySecondary == 0) ? 0 : enrolmentS / facilityRecord.caps_AdjustedDesignCapacitySecondary * 100;

                    //Note: Operating capacity is kept up to date on the facility record so no calculation is needed
                    recordToUpdate.caps_Kindergarten_operatingcapacity = facilityRecord.caps_OperatingCapacityKindergarten;
                    recordToUpdate.caps_Elementary_operatingcapacity = facilityRecord.caps_OperatingCapacityElementary;
                    recordToUpdate.caps_Secondary_operatingcapacity = facilityRecord.caps_OperatingCapacitySecondary;

                    recordToUpdate.caps_Kindergarten_operatingutilization = (facilityRecord.caps_OperatingCapacityKindergarten == 0) ? 0 : enrolmentK / facilityRecord.caps_OperatingCapacityKindergarten * 100;
                    recordToUpdate.caps_Elementary_operatingutilization = (facilityRecord.caps_OperatingCapacityElementary == 0) ? 0 : enrolmentE / facilityRecord.caps_OperatingCapacityElementary * 100;
                    recordToUpdate.caps_Secondary_operatingutilization = (facilityRecord.caps_OperatingCapacitySecondary == 0) ? 0 : enrolmentS / facilityRecord.caps_OperatingCapacitySecondary * 100;

                    recordToUpdate.caps_kindergarten_operatingcapacity_district = facilityRecord.caps_DistrictOperatingCapacityKindergarten;
                    recordToUpdate.caps_elementary_operatingcapacity_district = facilityRecord.caps_DistrictOperatingCapacityElementary;
                    recordToUpdate.caps_secondary_operatingcapacity_district = facilityRecord.caps_DistrictOperatingCapacitySecondary;

                    recordToUpdate.caps_kindergarten_operatingutilizatio_district = (facilityRecord.caps_DistrictOperatingCapacityKindergarten == 0) ? 0 : enrolmentK / facilityRecord.caps_DistrictOperatingCapacityKindergarten * 100;
                    recordToUpdate.caps_elementary_operatingutilization_district = (facilityRecord.caps_DistrictOperatingCapacityElementary == 0) ? 0 : enrolmentE / facilityRecord.caps_DistrictOperatingCapacityElementary * 100;
                    recordToUpdate.caps_secondary_operatingutilization_district = (facilityRecord.caps_DistrictOperatingCapacitySecondary == 0) ? 0 : enrolmentS / facilityRecord.caps_DistrictOperatingCapacitySecondary * 100;

                    service.Update(recordToUpdate); 
                    #endregion
                }
                else {
                    //Future
                    #region Future
                    var endDate = (DateTime?)((AliasedValue)reportingRecord["year.edu_enddate"]).Value;

                    var recordToUpdate = new caps_CapacityReporting();
                    recordToUpdate.Id = reportingRecord.Id;

                    //get matching enrolment projection record
                    var matchingProjection = enrolmentProjectionList.FirstOrDefault(r => r.caps_ProjectionYear.Id == reportingRecord.GetAttributeValue<EntityReference>("caps_schoolyear").Id);

                    var kDesign = facilityRecord.caps_AdjustedDesignCapacityKindergarten.GetValueOrDefault(0);
                    var eDesign = facilityRecord.caps_AdjustedDesignCapacityElementary.GetValueOrDefault(0);
                    var sDesign = facilityRecord.caps_AdjustedDesignCapacitySecondary.GetValueOrDefault(0);

                    var kDistrict = facilityRecord.caps_DistrictOperatingCapacityKindergarten.GetValueOrDefault(0);
                    var eDistrict = facilityRecord.caps_DistrictOperatingCapacityElementary.GetValueOrDefault(0);
                    var sDistrict = facilityRecord.caps_DistrictOperatingCapacitySecondary.GetValueOrDefault(0);

                    tracingService.Trace("Design Capacity - K:{0}; E:{1}; S:{2}", kDesign, eDesign, sDesign);

                    if (endDate.HasValue)
                    {
                        //loop project records
                        foreach (var projectRecord in projectRecords.Entities)
                        {
                            if ((DateTime?)((AliasedValue)projectRecord["milestone.caps_expectedactualdate"]).Value <= endDate)
                            {
                                //add design capacity on
                                kDesign += Convert.ToInt32(projectRecord.GetAttributeValue<decimal?>("caps_designcapacitykindergarten").GetValueOrDefault(0));
                                eDesign += Convert.ToInt32(projectRecord.GetAttributeValue<decimal?>("caps_designcapacityelementary").GetValueOrDefault(0));
                                sDesign += Convert.ToInt32(projectRecord.GetAttributeValue<decimal?>("caps_designcapacitysecondary").GetValueOrDefault(0));

                                //add district operating capacity
                                kDistrict += Convert.ToInt32(projectRecord.GetAttributeValue<decimal?>("caps_districtoperatingcapacitykindergarten").GetValueOrDefault(0));
                                eDistrict += Convert.ToInt32(projectRecord.GetAttributeValue<decimal?>("caps_districtoperatingcapacityelementary").GetValueOrDefault(0));
                                sDistrict += Convert.ToInt32(projectRecord.GetAttributeValue<decimal?>("caps_districtoperatingcapacitysecondary").GetValueOrDefault(0));
                            }
                        }
                    }

                    tracingService.Trace("Updated Design Capacity - K:{0}; E:{1}; S:{2}", kDesign, eDesign, sDesign);

                    decimal? futureEnrolmentK = (matchingProjection != null) ? matchingProjection.caps_EnrolmentProjectionKindergarten : enrolmentK;
                    decimal? futureEnrolmentE = (matchingProjection != null) ? matchingProjection.caps_EnrolmentProjectionElementary : enrolmentE;
                    decimal? futureEnrolmentS = (matchingProjection != null) ? matchingProjection.caps_EnrolmentProjectionSecondary : enrolmentS;

                    recordToUpdate.caps_Kindergarten_enrolment = futureEnrolmentK;
                    recordToUpdate.caps_Elementary_enrolment = futureEnrolmentE;
                    recordToUpdate.caps_Secondary_enrolment = futureEnrolmentS;

                    recordToUpdate.caps_Kindergarten_designcapacity = kDesign;
                    recordToUpdate.caps_Elementary_designcapacity = eDesign;
                    recordToUpdate.caps_Secondary_designcapacity = sDesign;

                    tracingService.Trace("Enrolment - K:{0}; E:{1}; S:{2}", futureEnrolmentK, futureEnrolmentE, futureEnrolmentS);

                    recordToUpdate.caps_Kindergarten_designutilization = (kDesign == 0) ? 0 : futureEnrolmentK / kDesign * 100;
                    recordToUpdate.caps_Elementary_designutilization = (eDesign == 0) ? 0 : futureEnrolmentE / eDesign * 100;
                    recordToUpdate.caps_Secondary_designutilization = (sDesign == 0) ? 0 : futureEnrolmentS / sDesign * 100;

                    //no sense in running this if the grade range isn't specified
                    if (facilityRecord.caps_LowestGrade != null && facilityRecord.caps_HighestGrade != null)
                    {
                        //Get operating capacity for the design numbers
                        var lowestGrade = facilityRecord.caps_LowestGrade.Value;
                        var highestGrade = facilityRecord.caps_HighestGrade.Value;

                        var result = capacityService.Calculate(kDesign, eDesign, sDesign, lowestGrade, highestGrade);

                        recordToUpdate.caps_Kindergarten_operatingcapacity = result.KindergartenCapacity;
                        recordToUpdate.caps_Elementary_operatingcapacity = result.ElementaryCapacity;
                        recordToUpdate.caps_Secondary_operatingcapacity = result.SecondaryCapacity;

                        recordToUpdate.caps_Kindergarten_operatingutilization = (result.KindergartenCapacity == 0) ? 0 : futureEnrolmentK / result.KindergartenCapacity * 100;
                        recordToUpdate.caps_Elementary_operatingutilization = (result.ElementaryCapacity == 0) ? 0 : futureEnrolmentE / result.ElementaryCapacity * 100;
                        recordToUpdate.caps_Secondary_operatingutilization = (result.SecondaryCapacity == 0) ? 0 : futureEnrolmentS / result.SecondaryCapacity * 100;
                    }

                    recordToUpdate.caps_kindergarten_operatingcapacity_district = kDistrict;
                    recordToUpdate.caps_elementary_operatingcapacity_district = eDistrict;
                    recordToUpdate.caps_secondary_operatingcapacity_district = sDistrict;

                    recordToUpdate.caps_kindergarten_operatingutilizatio_district = (kDistrict == 0) ? 0 : futureEnrolmentK / kDistrict * 100;
                    recordToUpdate.caps_elementary_operatingutilization_district = (eDistrict == 0) ? 0 : futureEnrolmentE / eDistrict * 100;
                    recordToUpdate.caps_secondary_operatingutilization_district = (sDistrict == 0) ? 0 : futureEnrolmentS / sDistrict * 100;

                    service.Update(recordToUpdate);
                }
                    #endregion
                
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
