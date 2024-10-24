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
    /// This CWA runs on a child care facility record.  It calculates historical, current and future capacities (design and operating)
    /// and updates the related Capacity Reporting records.  This data is used by the capacity reports.
    /// </summary>
    public class SetChildCareCapacityReporting : CodeActivity
    {
        protected override void Execute(CodeActivityContext executionContext)
        {
            ITracingService tracingService = executionContext.GetExtension<ITracingService>();

            IWorkflowContext context = executionContext.GetExtension<IWorkflowContext>();
            IOrganizationServiceFactory serviceFactory = executionContext.GetExtension<IOrganizationServiceFactory>();
            IOrganizationService service = serviceFactory.CreateOrganizationService(context.UserId);

            tracingService.Trace("{0}{1}", "Start Custom Workflow Activity: SetChildCareCapacityReporting", DateTime.Now.ToLongTimeString());

            var recordId = context.PrimaryEntityId;

            decimal? capacityUnder36Months = 0;
            decimal? capacity30MonthsSchoolAge = 0;
            decimal? capacityPreSchool = 0;
            decimal? capacityMultiAge = 0;
            decimal? capacitySchoolAge = 0;
            decimal? capacitySASG = 0;

            //Get Child Care Facility record with all current design capacity fields & school district
            var columns = new ColumnSet("caps_currentchildcareactualenrolment",
                                        "caps_name",
                                        "caps_usefutureforutilization",
                                        "caps_capacityunder36months",
                                        "caps_capacity30monthstoschoolage",
                                        "caps_capacitypreschool",
                                        "caps_capacitymultiage",
                                        "caps_capacityschoolage",
                                        "caps_capacityschoolageschoolgrounds",
                                        "caps_licensedcapacityunder36months",
                                        "caps_licensedcapacity30monthstoschoolage",
                                        "caps_licensedcapacitypreschool",
                                        "caps_licensedcapacitymultiage",
                                        "caps_licensedcapacityschoolage",
                                        "caps_licensedcapacitysasg");
            var childCareFacilityRecord = service.Retrieve(context.PrimaryEntityName, context.PrimaryEntityId, columns) as caps_Childcare;

            tracingService.Trace("Child Care Facility Name:{0}", childCareFacilityRecord.caps_Name);

            //var capacity = new Services.CapacityFactors(service);

            //Services.OperatingCapacity capacityService = new Services.OperatingCapacity(service, tracingService, capacity);

            //tracingService.Trace("Line: {0}", "40");
            //Get Current Child Care Enrolment Projections
            QueryExpression childCareEnrolmentProjQuery = new QueryExpression("caps_childcareenrolmentprojection");
            childCareEnrolmentProjQuery.ColumnSet.AllColumns = true;
            LinkEntity enrolmentYearLink = childCareEnrolmentProjQuery.AddLink("edu_year", "caps_schoolyear", "edu_yearid");
            enrolmentYearLink.Columns.AddColumns("edu_yearid", "statuscode", "edu_enddate");
            enrolmentYearLink.EntityAlias = "year";
            childCareEnrolmentProjQuery.Criteria = new FilterExpression();
            childCareEnrolmentProjQuery.Criteria.AddCondition("caps_childcarefacility", ConditionOperator.Equal, recordId);
            tracingService.Trace("Use For Future Utilization  {0}", childCareFacilityRecord.caps_UseFutureForUtilization);
            if (childCareFacilityRecord.caps_UseFutureForUtilization == true)
            {
                tracingService.Trace("Line 74");
                //submitted
                childCareEnrolmentProjQuery.Criteria.AddCondition("statuscode", ConditionOperator.Equal, 2);
            }
            else
            {
                tracingService.Trace("Line 80");
                //current
                childCareEnrolmentProjQuery.Criteria.AddCondition("statuscode", ConditionOperator.Equal, 714430001);
            }

            EntityCollection ccEnrolmentProjectionRecords = service.RetrieveMultiple(childCareEnrolmentProjQuery);
            var enrolmentProjectionList = ccEnrolmentProjectionRecords.Entities.Select(r => r.ToEntity<caps_ChildCareEnrolmentProjection>()).ToList();

            tracingService.Trace("Enrolment Projection List : {0}", ccEnrolmentProjectionRecords.Entities.Count);
            tracingService.Trace("Line 86");
            tracingService.Trace("Current Actual Enrolment:{0}", childCareFacilityRecord.caps_CurrentChildCareActualEnrolment);
            if (childCareFacilityRecord.caps_CurrentChildCareActualEnrolment != null)
            {
                var ccAcctualEnrolmentColumns = new ColumnSet("caps_capacityunder36months", "caps_capacity30monthstoschoolage",
                    "caps_capacitypreschool", "caps_capacitymultiage", "caps_capacityschoolage", "caps_capacitysasg");
                var ccActualEnrolmentRecord = service.Retrieve(caps_ChildCareActualEnrolment.EntityLogicalName, childCareFacilityRecord.caps_CurrentChildCareActualEnrolment.Id, ccAcctualEnrolmentColumns) as caps_ChildCareActualEnrolment;

                capacityUnder36Months = ccActualEnrolmentRecord.caps_CapacityUnder36Months.GetValueOrDefault(0);
                capacity30MonthsSchoolAge = ccActualEnrolmentRecord.caps_Capacity30MonthstoSchoolAge.GetValueOrDefault(0);
                capacityPreSchool = ccActualEnrolmentRecord.caps_CapacityPreschool.GetValueOrDefault(0);
                capacityMultiAge = ccActualEnrolmentRecord.caps_CapacityMultiAge.GetValueOrDefault(0);
                capacitySchoolAge = ccActualEnrolmentRecord.caps_CapacitySchoolAge.GetValueOrDefault(0);
                capacitySASG = ccActualEnrolmentRecord.caps_CapacitySASG.GetValueOrDefault(0);
                tracingService.Trace("Line: {0}", "99");
            }
            else if (childCareFacilityRecord.caps_CurrentChildCareActualEnrolment == null)
            {
                tracingService.Trace("Line: {0}", "105");
                //TODO: code to get latest child care projection for the current year
                foreach (Entity projection in ccEnrolmentProjectionRecords.Entities)
                {
                    tracingService.Trace("Line 106");
                    var yearStatus = (OptionSetValue)((AliasedValue)projection["year.statuscode"]).Value;
                    if (yearStatus.Value == 1)
                    {
                        //Current Year projections
                        tracingService.Trace("Line 112");
                        capacityUnder36Months = (!projection.Contains("caps_under36months")) ? 0 : projection.GetAttributeValue<decimal?>("caps_under36months").GetValueOrDefault(0);
                        capacity30MonthsSchoolAge = (!projection.Contains("caps_monthstoschoolage")) ? 0 : projection.GetAttributeValue<decimal?>("caps_monthstoschoolage").GetValueOrDefault(0);
                        capacityPreSchool = (!projection.Contains("caps_preschool")) ? 0 : projection.GetAttributeValue<decimal?>("caps_preschool").GetValueOrDefault(0);
                        capacityMultiAge = (!projection.Contains("caps_multiage")) ? 0 : projection.GetAttributeValue<decimal?>("caps_multiage").GetValueOrDefault(0);
                        capacitySchoolAge = (!projection.Contains("caps_schoolage")) ? 0 : projection.GetAttributeValue<decimal?>("caps_schoolage").GetValueOrDefault(0);
                        capacitySASG = (!projection.Contains("caps_schoolageonschoolgrounds")) ? 0 : projection.GetAttributeValue<decimal?>("caps_schoolageonschoolgrounds").GetValueOrDefault(0);
                        break;
                    }
                }
                tracingService.Trace("Line 120");
                //Set Current Year Projections
                var childCareFacilityToUpdate = new caps_Childcare();
                childCareFacilityToUpdate.Id = recordId;
                childCareFacilityToUpdate.caps_Under36Months_CurrentEnrolment = capacityUnder36Months;
                childCareFacilityToUpdate.caps_30MonthstoSchoolAge_CurrentEnrolment = capacity30MonthsSchoolAge;
                childCareFacilityToUpdate.caps_Preschool_CurrentEnrolment = capacityPreSchool;
                childCareFacilityToUpdate.caps_MultiAge_CurrentEnrolment = capacityMultiAge;
                childCareFacilityToUpdate.caps_SchoolAge_CurrentEnrolment = capacitySchoolAge;
                childCareFacilityToUpdate.caps_SASG_CurrentEnrolment = capacitySASG;

                service.Update(childCareFacilityToUpdate);

            }




            //Get Projects for this child care facility with occupancy dates in the future
            QueryExpression projectQuery = new QueryExpression("caps_projecttracker");
            projectQuery.ColumnSet.AddColumns("caps_nettotalunder36months", "caps_nettotal30monthstoschoolage", "caps_nettotalpreschool", "caps_nettotalmultiage", "caps_nettotalschoolage", "caps_nettotalsasg");
            projectQuery.Criteria = new FilterExpression();
            projectQuery.Criteria.AddCondition("caps_childcarefacility", ConditionOperator.Equal, recordId);
            projectQuery.Criteria.AddCondition("statecode", ConditionOperator.Equal, 0);

            LinkEntity milestoneLink = projectQuery.AddLink("caps_projectmilestone", "caps_projecttrackerid", "caps_projecttracker", JoinOperator.Inner);
            milestoneLink.Columns.AddColumn("caps_expectedactualdate");
            milestoneLink.EntityAlias = "milestone";
            milestoneLink.LinkCriteria = new FilterExpression();
            milestoneLink.LinkCriteria.AddCondition("caps_name", ConditionOperator.Equal, "Occupancy");
            milestoneLink.LinkCriteria.AddCondition("caps_expectedactualdate", ConditionOperator.OnOrAfter, DateTime.Now);

            EntityCollection projectRecords = service.RetrieveMultiple(projectQuery);

            //Get Child Care Capacity Reporting Records
            QueryExpression query = new QueryExpression("caps_childcarecapacityreporting");
            query.ColumnSet.AllColumns = true;
            LinkEntity yearLink = query.AddLink("edu_year", "caps_schoolyear", "edu_yearid");
            yearLink.Columns.AddColumns("edu_yearid", "statuscode", "edu_enddate");
            yearLink.EntityAlias = "year";
            query.Criteria = new FilterExpression();
            query.Criteria.AddCondition("year", "edu_type", ConditionOperator.Equal, 757500001); //School
            query.Criteria.AddCondition("statecode", ConditionOperator.Equal, 0);
            query.Criteria.AddCondition("caps_childcarefacility", ConditionOperator.Equal, recordId);
            EntityCollection reportingRecords = service.RetrieveMultiple(query);

            foreach (var reportingRecord in reportingRecords.Entities)
            {
                tracingService.Trace("Line: {0}", "71");
                //Get Type - Historical, Current, Future
                var recordType = (OptionSetValue)((AliasedValue)reportingRecord["year.statuscode"]).Value;
                tracingService.Trace("Record Type:{0}", recordType.Value);

                //Historical
                if (recordType.Value == 757500001)
                {
                    //Historical
                    #region Historical

                    //For each record, get top 1 Facility History record where year snapshot equals the school year record, order by created on date descending

                    var schoolYearRecord = ((AliasedValue)reportingRecord["year.edu_yearid"]).Value;
                    if (schoolYearRecord != null)
                    {

                        QueryExpression snapshotQuery = new QueryExpression("caps_childcarefacilityhistory");
                        snapshotQuery.ColumnSet.AllColumns = true;
                        LinkEntity enrolmentLink = snapshotQuery.AddLink("caps_childcareactualenrolment", "caps_childcareactualenrolment", "caps_childcareactualenrolmentid");
                        enrolmentLink.EntityAlias = "enrolment";
                        enrolmentLink.Columns.AddColumns("caps_capacityunder36months", "caps_capacity30monthstoschoolage", "caps_capacitypreschool", "caps_capacitymultiage", "caps_capacityschoolage", "caps_capacitysasg");

                        snapshotQuery.Criteria = new FilterExpression();
                        snapshotQuery.Criteria.AddCondition("caps_childcarefacility", ConditionOperator.Equal, recordId);
                        snapshotQuery.Criteria.AddCondition("caps_schoolyear", ConditionOperator.Equal, schoolYearRecord);
                        snapshotQuery.AddOrder("createdon", OrderType.Descending);
                        snapshotQuery.TopCount = 1;

                        EntityCollection snapshotRecords = service.RetrieveMultiple(snapshotQuery);

                        //Get top record
                        if (snapshotRecords.Entities.Count() == 1)
                        {
                            tracingService.Trace("Line: {0}", "93");
                            var topRecord = snapshotRecords.Entities[0] as caps_ChildCareFacilityHistory;
                            tracingService.Trace("Top Record:{0}", topRecord);
                            //Update the capacity record
                            var recordToUpdate = new caps_ChildCareCapacityReporting();
                            recordToUpdate.Id = reportingRecord.Id;

                            tracingService.Trace("Historical Enrol Under 36 Months:{0}", ((decimal?)((AliasedValue)topRecord["enrolment.caps_capacityunder36months"]).Value).GetValueOrDefault(0));

                            decimal? hEnrolUnder36Months = (!topRecord.Contains("enrolment.caps_capacityunder36months")) ? 0 : ((decimal?)((AliasedValue)topRecord["enrolment.caps_capacityunder36months"]).Value).GetValueOrDefault(0);
                            decimal? hEnrol30MonthsToSchoolAge = (!topRecord.Contains("enrolment.caps_capacity30monthstoschoolage")) ? 0 : ((decimal?)((AliasedValue)topRecord["enrolment.caps_capacity30monthstoschoolage"]).Value).GetValueOrDefault(0);
                            decimal? hEnrolPreSchool = (!topRecord.Contains("enrolment.caps_capacitypreschool")) ? 0 : ((decimal?)((AliasedValue)topRecord["enrolment.caps_capacitypreschool"]).Value).GetValueOrDefault(0);
                            decimal? hEnrolMultiAge = (!topRecord.Contains("enrolment.caps_capacitymultiage")) ? 0 : ((decimal?)((AliasedValue)topRecord["enrolment.caps_capacitymultiage"]).Value).GetValueOrDefault(0);
                            decimal? hEnrolSchoolAge = (!topRecord.Contains("enrolment.caps_capacityschoolage")) ? 0 : ((decimal?)((AliasedValue)topRecord["enrolment.caps_capacityschoolage"]).Value).GetValueOrDefault(0);
                            decimal? hEnrolSASG = (!topRecord.Contains("enrolment.caps_capacitysasg")) ? 0 : ((decimal?)((AliasedValue)topRecord["enrolment.caps_capacitysasg"]).Value).GetValueOrDefault(0);

                            tracingService.Trace("Line:{0}", "213");

                            decimal? hDesignUnder36Months = (!topRecord.Contains("caps_capacityunder36months")) ? 0 : ((decimal?)topRecord["caps_capacityunder36months"]).GetValueOrDefault(0);
                            decimal? hDesign30MonthsToSchoolAge = (!topRecord.Contains("caps_capacity30monthstoschoolage")) ? 0 : ((decimal?)topRecord["caps_capacity30monthstoschoolage"]).GetValueOrDefault(0);
                            decimal? hDesignPreSchool = (!topRecord.Contains("caps_capacitypreschool")) ? 0 : ((decimal?)topRecord["caps_capacitypreschool"]).GetValueOrDefault(0);
                            decimal? hDesignMultiAge = (!topRecord.Contains("caps_capacitymultiage")) ? 0 : ((decimal?)topRecord["caps_capacitymultiage"]).GetValueOrDefault(0);
                            decimal? hDesignSchoolAge = (!topRecord.Contains("caps_capacityschoolage")) ? 0 : ((decimal?)topRecord["caps_capacityschoolage"]).GetValueOrDefault(0);
                            decimal? hDesignSASG = (!topRecord.Contains("caps_capacityschoolageschoolgrounds")) ? 0 : ((decimal?)topRecord["caps_capacityschoolageschoolgrounds"]).GetValueOrDefault(0);

                            tracingService.Trace("Line:{0}", "222");

                            decimal? hLicensedUnder36Months = (!topRecord.Contains("caps_licensedcapacityunder36months")) ? 0 : ((decimal?)topRecord["caps_licensedcapacityunder36months"]).GetValueOrDefault(0);
                            decimal? hLicensed30MonthsToSchoolAge = (!topRecord.Contains("caps_licensedcapacity30monthstoschoolage")) ? 0 : ((decimal?)topRecord["caps_licensedcapacity30monthstoschoolage"]).GetValueOrDefault(0);
                            decimal? hLicensedPreSchool = (!topRecord.Contains("caps_licensedcapacitypreschool")) ? 0 : ((decimal?)topRecord["caps_licensedcapacitypreschool"]).GetValueOrDefault(0);
                            decimal? hLicensedMultiAge = (!topRecord.Contains("caps_licensedcapacitymultiage")) ? 0 : ((decimal?)topRecord["caps_licensedcapacitymultiage"]).GetValueOrDefault(0);
                            decimal? hLicensedSchoolAge = (!topRecord.Contains("caps_licensedcapacityschoolage")) ? 0 : ((decimal?)topRecord["caps_licensedcapacityschoolage"]).GetValueOrDefault(0);
                            decimal? hLicensedSASG = (!topRecord.Contains("caps_licensedcapacitysasg")) ? 0 : ((decimal?)topRecord["caps_licensedcapacitysasg"]).GetValueOrDefault(0);


                            tracingService.Trace("Line:{0}", "232");
                            //Updating Enrolment
                            recordToUpdate.caps_Under36months_Enrolment = hEnrolUnder36Months;
                            recordToUpdate.caps_MonthstoSchoolAge_Enrolment = hEnrol30MonthsToSchoolAge;
                            recordToUpdate.caps_Preschool_Enrolment = hEnrolPreSchool;
                            recordToUpdate.caps_MultiAge_Enrolment = hEnrolMultiAge;
                            recordToUpdate.caps_SchoolAge_Enrolment = hEnrolSchoolAge;
                            recordToUpdate.caps_SchoolAgeonSchoolGrounds_Enrolment = hEnrolSASG;

                            //Updating Design Capacity Section
                            recordToUpdate.caps_CapacityUnder36Months = hDesignUnder36Months;
                            recordToUpdate.caps_Capacity30MonthstoSchoolAge = hDesign30MonthsToSchoolAge;
                            recordToUpdate.caps_CapacityPreschool = hDesignPreSchool;
                            recordToUpdate.caps_CapacityMultiAge = hDesignMultiAge;
                            recordToUpdate.caps_CapacitySchoolAge = hDesignSchoolAge;
                            recordToUpdate.caps_CapacitySchoolAgeonSchoolGrounds = hDesignSASG;

                            //Updating Licensed Capacity
                            recordToUpdate.caps_LicensedUnder36Months = hLicensedUnder36Months;
                            recordToUpdate.caps_Licensed30MonthstoSchoolAge = hLicensed30MonthsToSchoolAge;
                            recordToUpdate.caps_LicensedPreschool = hLicensedPreSchool;
                            recordToUpdate.caps_LicensedMultiAge = hLicensedMultiAge;
                            recordToUpdate.caps_LicensedSchoolAge = hLicensedSchoolAge;
                            recordToUpdate.caps_LicensedSchoolAgeonSchoolGrounds = hLicensedSASG;

                            service.Update(recordToUpdate);
                        }
                    }
                        #endregion
                }
                else if (recordType.Value == 1)
                {
                    //Current
                    #region Current
                    tracingService.Trace("Line: {0}", "123");

                    //Update the capacity record
                    var recordToUpdate = new caps_ChildCareCapacityReporting();
                    recordToUpdate.Id = reportingRecord.Id;

                    recordToUpdate.caps_Under36months_Enrolment = capacityUnder36Months;
                    recordToUpdate.caps_MonthstoSchoolAge_Enrolment = capacity30MonthsSchoolAge;
                    recordToUpdate.caps_Preschool_Enrolment = capacityPreSchool;
                    recordToUpdate.caps_MultiAge_Enrolment = capacityMultiAge;
                    recordToUpdate.caps_SchoolAge_Enrolment = capacitySchoolAge;
                    recordToUpdate.caps_SchoolAgeonSchoolGrounds_Enrolment = capacitySASG;

                    //Licensed and Design will always come from the CC Facility

                    //Updating Future DESIGN CAPACITY
                    recordToUpdate.caps_CapacityUnder36Months = (childCareFacilityRecord.caps_CapacityUnder36Months == 0) ? 0 : childCareFacilityRecord.caps_CapacityUnder36Months;
                    recordToUpdate.caps_Capacity30MonthstoSchoolAge = (childCareFacilityRecord.caps_Capacity30MonthstoSchoolAge == 0) ? 0 : childCareFacilityRecord.caps_Capacity30MonthstoSchoolAge;
                    recordToUpdate.caps_CapacityPreschool = (childCareFacilityRecord.caps_CapacityPreschool == 0) ? 0 : childCareFacilityRecord.caps_CapacityPreschool;
                    recordToUpdate.caps_CapacityMultiAge = (childCareFacilityRecord.caps_CapacityMultiAge == 0) ? 0 : childCareFacilityRecord.caps_CapacityMultiAge;
                    recordToUpdate.caps_CapacitySchoolAge = (childCareFacilityRecord.caps_CapacitySchoolAge == 0) ? 0 : childCareFacilityRecord.caps_CapacitySchoolAge;
                    recordToUpdate.caps_CapacitySchoolAgeonSchoolGrounds = (childCareFacilityRecord.caps_CapacitySchoolAgeSchoolGrounds == 0) ? 0 : childCareFacilityRecord.caps_CapacitySchoolAgeSchoolGrounds;

                    //Updating Future LICENSED CAPACITY
                    recordToUpdate.caps_LicensedUnder36Months = (childCareFacilityRecord.caps_LicensedCapacityUnder36Months == 0) ? 0 : childCareFacilityRecord.caps_LicensedCapacityUnder36Months; ;
                    recordToUpdate.caps_Licensed30MonthstoSchoolAge = (childCareFacilityRecord.caps_LicensedCapacity30MonthstoSchoolAge == 0) ? 0 : childCareFacilityRecord.caps_LicensedCapacity30MonthstoSchoolAge;
                    recordToUpdate.caps_LicensedPreschool = (childCareFacilityRecord.caps_LicensedCapacityPreschool == 0) ? 0 : childCareFacilityRecord.caps_LicensedCapacityPreschool;
                    recordToUpdate.caps_LicensedMultiAge = (childCareFacilityRecord.caps_LicensedCapacityMultiAge == 0) ? 0 : childCareFacilityRecord.caps_LicensedCapacityMultiAge;
                    recordToUpdate.caps_LicensedSchoolAge = (childCareFacilityRecord.caps_LicensedCapacitySchoolAge == 0) ? 0 : childCareFacilityRecord.caps_LicensedCapacitySchoolAge;
                    recordToUpdate.caps_LicensedSchoolAgeonSchoolGrounds = (childCareFacilityRecord.caps_LicensedCapacitySASG == 0) ? 0 : childCareFacilityRecord.caps_LicensedCapacitySASG;


                    service.Update(recordToUpdate);
                    #endregion
                }
                else if (recordType.Value == 757500000)
                {
                    var recordToUpdate = new caps_ChildCareCapacityReporting();
                    recordToUpdate.Id = reportingRecord.Id;

                    tracingService.Trace("Enrolment Projection Count:{0}", ccEnrolmentProjectionRecords.Entities.Count);

                    foreach(var Projection in ccEnrolmentProjectionRecords.Entities)
                    {
                        var schoolYear = ((EntityReference)Projection["caps_schoolyear"]).Id;
                        var reportingYear = reportingRecord.GetAttributeValue<EntityReference>("caps_schoolyear").Id;
                        tracingService.Trace("Enrolment School Year:{0}", schoolYear);
                        tracingService.Trace("Reporting Year: {0}", reportingRecord.GetAttributeValue<EntityReference>("caps_schoolyear").Id);
                        if(schoolYear == reportingYear)
                        {
                            decimal? fEnrolUnder36Month = (decimal?)Projection["caps_under36months"] == 0 ? 0 : (decimal?)Projection["caps_under36months"];
                            decimal? fEnrolMonthsToSchoolAge = (decimal?)Projection["caps_monthstoschoolage"] == 0 ? 0 : (decimal?)Projection["caps_monthstoschoolage"];
                            decimal? fEnrolPreSchool = (decimal?)Projection["caps_preschool"] == 0 ? 0 : (decimal?)Projection["caps_preschool"];
                            decimal? fEnrolMultiAge = (decimal?)Projection["caps_multiage"] == 0 ? 0 : (decimal?)Projection["caps_multiage"];
                            decimal? fEnrolSchoolAge = (decimal?)Projection["caps_schoolage"] == 0 ? 0 : (decimal?)Projection["caps_schoolage"];
                            decimal? fEnrolSASG = (decimal?)Projection["caps_schoolageonschoolgrounds"] == 0 ? 0 : (decimal?)Projection["caps_schoolageonschoolgrounds"];
                            //Updating Future ENROLMENT
                            recordToUpdate.caps_Under36months_Enrolment = fEnrolUnder36Month;
                            recordToUpdate.caps_MonthstoSchoolAge_Enrolment = fEnrolMonthsToSchoolAge;
                            recordToUpdate.caps_Preschool_Enrolment = fEnrolPreSchool;
                            recordToUpdate.caps_MultiAge_Enrolment = fEnrolMultiAge;
                            recordToUpdate.caps_SchoolAge_Enrolment = fEnrolSchoolAge;
                            recordToUpdate.caps_SchoolAgeonSchoolGrounds_Enrolment = fEnrolSASG;
                        }

                    }

                    //get matching enrolment projection record
                    var matchingProjection = enrolmentProjectionList.FirstOrDefault(r => r.caps_SchoolYear.Id == reportingRecord.GetAttributeValue<EntityReference>("caps_schoolyear").Id);
                   

                    //For enrolment values we always get it from the CC Enrolment Projection record
                    //decimal? fEnrolUnder36Month = (matchingProjection.caps_Under36Months == 0) ? 0 : matchingProjection.caps_Under36Months;
                    //tracingService.Trace("Future Under 36 months:{0}", fEnrolUnder36Month);
                    //decimal? fEnrolMonthsToSchoolAge = (matchingProjection.caps_MonthstoSchoolAge == 0) ? 0 : matchingProjection.caps_MonthstoSchoolAge;
                    //tracingService.Trace("Future Enrol Months To School Age:{0}", fEnrolMonthsToSchoolAge);
                    //decimal? fEnrolPreSchool = (matchingProjection.caps_Preschool == 0) ? 0 : matchingProjection.caps_Preschool;
                    //decimal? fEnrolMultiAge = (matchingProjection.caps_MultiAge == 0) ? 0 : matchingProjection.caps_MultiAge;
                    //decimal? fEnrolSchoolAge = (matchingProjection.caps_SchoolAge == 0) ? 0 : matchingProjection.caps_SchoolAge;
                    //decimal? fEnrolSASG = (matchingProjection.caps_SchoolAgeonSchoolGrounds == 0) ? 0 : matchingProjection.caps_SchoolAgeonSchoolGrounds;
                    //tracingService.Trace("SASG:{0}", matchingProjection.caps_SchoolAgeonSchoolGrounds);
                    //tracingService.Trace("Enrolment ID:{0}", matchingProjection.caps_ChildCareEnrolmentProjectionId);

                    //For Design We need to use the same project query that is linked to milestones and get the fields for the Design from the total section of the Child Care tab of the project and 
                    //combine it with the Design on the CC Facility Design fields
                    var fDesignUnder36Months = childCareFacilityRecord.caps_CapacityUnder36Months.GetValueOrDefault(0);
                    var fDesign30MonthstoSchoolAge = childCareFacilityRecord.caps_Capacity30MonthstoSchoolAge.GetValueOrDefault(0);
                    var fDesignPreSchool = childCareFacilityRecord.caps_CapacityPreschool.GetValueOrDefault(0);
                    var fDesignMultiAge = childCareFacilityRecord.caps_CapacityMultiAge.GetValueOrDefault(0);
                    var fDesignSchoolAge = childCareFacilityRecord.caps_CapacitySchoolAge.GetValueOrDefault(0);
                    var fDesignSASG = childCareFacilityRecord.caps_CapacitySchoolAgeSchoolGrounds.GetValueOrDefault(0);

                    //For Licensed we get the values from the Licensed section of the Child Care Facility. However, if the licensed values change on CC Facility we need to update current and future CC capacity reporting fields.
                    var fLicensedUnder36Months = childCareFacilityRecord.caps_LicensedCapacityUnder36Months.GetValueOrDefault(0);
                    var fLicensed30MonthstoSchoolAge = childCareFacilityRecord.caps_LicensedCapacity30MonthstoSchoolAge.GetValueOrDefault(0);
                    var fLicensedPreSchool = childCareFacilityRecord.caps_LicensedCapacityPreschool.GetValueOrDefault(0);
                    var fLicensedMultiAge = childCareFacilityRecord.caps_LicensedCapacityMultiAge.GetValueOrDefault(0);
                    var fLicensedSchoolAge = childCareFacilityRecord.caps_LicensedCapacitySchoolAge.GetValueOrDefault(0);
                    var fLicensedSASG = childCareFacilityRecord.caps_LicensedCapacitySASG.GetValueOrDefault(0);

                    tracingService.Trace("Design Capacity - Under36Months:{0}; 30MonthToSchoolAge:{1}; PreSchool:{2}; MultiAge:{3}; SchoolAge:{4}; SASG:{5}",
                        fDesignUnder36Months, fDesign30MonthstoSchoolAge, fDesignPreSchool, fDesignMultiAge, fDesignSchoolAge, fDesignSASG);
                    tracingService.Trace("Licensed Capacity - Under36Months:{0}; 30MonthToSchoolAge:{1}; PreSchool:{2}; MultiAge:{3}; SchoolAge:{4}; SASG:{5}",
                        fLicensedUnder36Months, fLicensed30MonthstoSchoolAge, fLicensedPreSchool, fLicensedMultiAge, fLicensedSchoolAge, fLicensedSASG);

                    var endDate = (DateTime?)((AliasedValue)reportingRecord["year.edu_enddate"]).Value;
                    if (endDate.HasValue)
                    {
                        //loop project records
                        foreach (var projectRecord in projectRecords.Entities)
                        {
                            if ((DateTime?)((AliasedValue)projectRecord["milestone.caps_expectedactualdate"]).Value <= endDate)
                            {
                                //Design capacity for future
                                fDesignUnder36Months += Convert.ToInt32(projectRecord.GetAttributeValue<decimal?>("caps_nettotalunder36months").GetValueOrDefault(0));
                                fDesign30MonthstoSchoolAge += Convert.ToInt32(projectRecord.GetAttributeValue<decimal?>("caps_nettotal30monthstoschoolage").GetValueOrDefault(0));
                                fDesignPreSchool += Convert.ToInt32(projectRecord.GetAttributeValue<decimal?>("caps_nettotalpreschool").GetValueOrDefault(0));
                                fDesignMultiAge += Convert.ToInt32(projectRecord.GetAttributeValue<decimal?>("caps_nettotalmultiage").GetValueOrDefault(0));
                                fDesignSchoolAge += Convert.ToInt32(projectRecord.GetAttributeValue<decimal?>("caps_nettotalschoolage").GetValueOrDefault(0));
                                fDesignSASG += Convert.ToInt32(projectRecord.GetAttributeValue<decimal?>("caps_nettotalsasg").GetValueOrDefault(0));

                            }
                        }
                    }

                    //tracingService.Trace("Updated Design Capacity - K:{0}; E:{1}; S:{2}", kDesign, eDesign, sDesign);
                    
                    

                    //Updating Future DESIGN CAPACITY
                    recordToUpdate.caps_CapacityUnder36Months = fDesignUnder36Months;
                    recordToUpdate.caps_Capacity30MonthstoSchoolAge = fDesign30MonthstoSchoolAge;
                    recordToUpdate.caps_CapacityPreschool = fDesignPreSchool;
                    recordToUpdate.caps_CapacityMultiAge = fDesignMultiAge;
                    recordToUpdate.caps_CapacitySchoolAge = fDesignSchoolAge;
                    recordToUpdate.caps_CapacitySchoolAgeonSchoolGrounds = fDesignSASG;

                    //Updating Future LICENSED CAPACITY
                    recordToUpdate.caps_LicensedUnder36Months = fLicensedUnder36Months;
                    recordToUpdate.caps_Licensed30MonthstoSchoolAge = fLicensed30MonthstoSchoolAge;
                    recordToUpdate.caps_LicensedPreschool = fLicensedPreSchool;
                    recordToUpdate.caps_LicensedMultiAge = fLicensedMultiAge;
                    recordToUpdate.caps_LicensedSchoolAge = fLicensedSchoolAge;
                    recordToUpdate.caps_LicensedSchoolAgeonSchoolGrounds = fLicensedSASG;

                    //tracingService.Trace("Enrolment - K:{0}; E:{1}; S:{2}", futureEnrolmentK, futureEnrolmentE, futureEnrolmentS);
                    service.Update(recordToUpdate);
                }
                    

            }
        }
    }
}

