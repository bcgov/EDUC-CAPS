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
            if (childCareFacilityRecord.caps_UseFutureForUtilization.GetValueOrDefault(false))
            {
                //submitted
                childCareEnrolmentProjQuery.Criteria.AddCondition("statuscode", ConditionOperator.Equal, 2);
            }
            else
            {
                //current
                childCareEnrolmentProjQuery.Criteria.AddCondition("statuscode", ConditionOperator.Equal, 714430001);
            }

            EntityCollection ccEnrolmentProjectionRecords = service.RetrieveMultiple(childCareEnrolmentProjQuery);
            var enrolmentProjectionList = ccEnrolmentProjectionRecords.Entities.Select(r => r.ToEntity<caps_ChildCareEnrolmentProjection>()).ToList();


            if (childCareFacilityRecord.caps_CurrentChildCareActualEnrolment != null)
            {
                var ccAcctualEnrolmentColumns = new ColumnSet("caps_capacityunder36months", "caps_capacity30monthstoschoolage", 
                    "caps_capacitypreschool", "caps_capacitymultiage", "caps_capacityschoolage", "caps_capacitysasg");
                var ccActualEnrolmentRecord = service.Retrieve(caps_ChildCareActualEnrolment.EntityLogicalName, childCareFacilityRecord.caps_CurrentChildCareActualEnrolment.Id, ccAcctualEnrolmentColumns) as caps_ChildCareActualEnrolment;

                capacityUnder36Months = ccActualEnrolmentRecord.caps_CapacityUnder36Months;
                capacity30MonthsSchoolAge = ccActualEnrolmentRecord.caps_Capacity30MonthstoSchoolAge;
                capacityPreSchool = ccActualEnrolmentRecord.caps_CapacityPreschool;
                capacityMultiAge = ccActualEnrolmentRecord.caps_CapacityMultiAge;
                capacitySchoolAge = ccActualEnrolmentRecord.caps_CapacitySchoolAge;
                capacitySASG = ccActualEnrolmentRecord.caps_CapacitySASG;
                //tracingService.Trace("Line: {0}", "49");
            }
            else
            {
                //TODO: code to get latest child care projection for the current year
                foreach (Entity projection in ccEnrolmentProjectionRecords.Entities)
                {
                    var yearStatus = (OptionSetValue)((AliasedValue)projection["year.statuscode"]).Value;
                    if (yearStatus.Value == 1)
                    {
                        //Current Year projections
                        capacityUnder36Months = projection.GetAttributeValue<int?>("caps_under36months").GetValueOrDefault(0);
                        capacity30MonthsSchoolAge = projection.GetAttributeValue<int?>("caps_monthstoschoolage").GetValueOrDefault(0);
                        capacityPreSchool = projection.GetAttributeValue<int?>("caps_preschool").GetValueOrDefault(0);
                        capacityMultiAge = projection.GetAttributeValue<int?>("caps_multiage").GetValueOrDefault(0);
                        capacitySchoolAge = projection.GetAttributeValue<int?>("caps_schoolage").GetValueOrDefault(0);
                        capacitySASG = projection.GetAttributeValue<int?>("caps_schoolageonschoolgrounds").GetValueOrDefault(0);
                        break;
                    }
                }

                //Set Current Year Projections
                var childCareFacilityToUpdate = new caps_Childcare();
                childCareFacilityToUpdate.Id = recordId;
                childCareFacilityToUpdate.caps_Under36Months_CurrentEnrolment = (int)capacityUnder36Months;
                childCareFacilityToUpdate.caps_30MonthstoSchoolAge_CurrentEnrolment = (int)capacity30MonthsSchoolAge;
                childCareFacilityToUpdate.caps_Preschool_CurrentEnrolment = (int)capacityPreSchool;
                childCareFacilityToUpdate.caps_MultiAge_CurrentEnrolment = (int)capacityMultiAge;
                childCareFacilityToUpdate.caps_SchoolAge_CurrentEnrolment = (int)capacitySchoolAge;
                childCareFacilityToUpdate.caps_SASG_CurrentEnrolment = (int)capacitySASG;

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

                //Historical
                if (recordType.Value == 757500001)
                {
                    //Historical
                    #region Historical

                    //For each record, get top 1 Facility History record where year snapshot equals the school year record, order by created on date descending
                    tracingService.Trace("Year Value: {0}", ((AliasedValue)reportingRecord["year.edu_yearid"]).Value);
                    var schoolYearRecord = ((AliasedValue)reportingRecord["year.edu_yearid"]).Value;

                    QueryExpression snapshotQuery = new QueryExpression("caps_childcarefacilityhistory");
                    snapshotQuery.ColumnSet.AllColumns = true;
                    LinkEntity enrolmentLink = snapshotQuery.AddLink("caps_childcareactualenrolment", "caps_childcareactualenrolment", "caps_childcareactualenrolmentid");
                    enrolmentLink.EntityAlias = "enrolment";
                    enrolmentLink.Columns.AddColumns("caps_capacityunder36months", "caps_capacity30monthstoschoolage", "caps_capacitypreschool","caps_capacitymultiage", "caps_capacityschoolage", "caps_capacitysasg");

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
                    recordToUpdate.caps_CapacityUnder36Months = childCareFacilityRecord.caps_CapacityUnder36Months;
                    recordToUpdate.caps_Capacity30MonthstoSchoolAge = childCareFacilityRecord.caps_Capacity30MonthstoSchoolAge;
                    recordToUpdate.caps_CapacityPreschool = childCareFacilityRecord.caps_CapacityPreschool;
                    recordToUpdate.caps_CapacityMultiAge = childCareFacilityRecord.caps_CapacityMultiAge;
                    recordToUpdate.caps_CapacitySchoolAge = childCareFacilityRecord.caps_CapacitySchoolAge;
                    recordToUpdate.caps_CapacitySchoolAgeonSchoolGrounds = childCareFacilityRecord.caps_CapacitySchoolAgeSchoolGrounds;

                    //Updating Future LICENSED CAPACITY
                    recordToUpdate.caps_LicensedUnder36Months = childCareFacilityRecord.caps_LicensedCapacityUnder36Months; ;
                    recordToUpdate.caps_Licensed30MonthstoSchoolAge = childCareFacilityRecord.caps_LicensedCapacity30MonthstoSchoolAge;
                    recordToUpdate.caps_LicensedPreschool = childCareFacilityRecord.caps_LicensedCapacityPreschool;
                    recordToUpdate.caps_LicensedMultiAge = childCareFacilityRecord.caps_LicensedCapacityMultiAge;
                    recordToUpdate.caps_LicensedSchoolAge = childCareFacilityRecord.caps_LicensedCapacitySchoolAge;
                    recordToUpdate.caps_LicensedSchoolAgeonSchoolGrounds = childCareFacilityRecord.caps_LicensedCapacitySASG;


                    service.Update(recordToUpdate);
                    #endregion
                }
                else
                {
                    //Future
                    #region Future
                    var endDate = (DateTime?)((AliasedValue)reportingRecord["year.edu_enddate"]).Value;

                    var recordToUpdate = new caps_ChildCareCapacityReporting();
                    recordToUpdate.Id = reportingRecord.Id;
                    
                    //get matching enrolment projection record
                    var matchingProjection = enrolmentProjectionList.FirstOrDefault(r => r.caps_SchoolYear.Id == reportingRecord.GetAttributeValue<EntityReference>("caps_schoolyear").Id);
                    
                    //For enrolment values we always get it from the CC Enrolment Projection record
                    decimal? fEnrolUnder36Month = (matchingProjection != null) ? matchingProjection.caps_Under36Months : 0;
                    decimal? fEnrolMonthsToSchoolAge = (matchingProjection != null) ? matchingProjection.caps_MonthstoSchoolAge : 0;
                    decimal? fEnrolPreSchool = (matchingProjection != null) ? matchingProjection.caps_Preschool : 0;
                    decimal? fEnrolMultiAge = (matchingProjection != null) ? matchingProjection.caps_MultiAge : 0;
                    decimal? fEnrolSchoolAge = (matchingProjection != null) ? matchingProjection.caps_SchoolAge : 0;
                    decimal? fEnrolSASG = (matchingProjection != null) ? matchingProjection.caps_SchoolAgeonSchoolGrounds : 0;

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

                    //Updating Future ENROLMENT
                    recordToUpdate.caps_Under36months_Enrolment = fEnrolUnder36Month;
                    recordToUpdate.caps_MonthstoSchoolAge_Enrolment = fEnrolMonthsToSchoolAge;
                    recordToUpdate.caps_Preschool_Enrolment = fEnrolPreSchool;
                    recordToUpdate.caps_MultiAge_Enrolment = fEnrolMultiAge;
                    recordToUpdate.caps_SchoolAge_Enrolment = fEnrolSchoolAge;
                    recordToUpdate.caps_SchoolAgeonSchoolGrounds_Enrolment = fLicensedSASG;
                    
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
                #endregion

            }
        }
    }
}
