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
                                    "caps_currentenrolment");
            var facilityRecord = service.Retrieve(context.PrimaryEntityName, context.PrimaryEntityId, columns) as caps_Facility;

            tracingService.Trace("Line: {0}", "40");
            if (facilityRecord.caps_CurrentEnrolment != null)
            {
                var enrolmentRecord = service.Retrieve(caps_facilityenrolment.EntityLogicalName, facilityRecord.caps_CurrentEnrolment.Id, new ColumnSet("caps_sumofkindergarten", "caps_sumofelementary", "caps_sumofsecondary")) as caps_facilityenrolment;

                enrolmentK = enrolmentRecord.caps_SumofKindergarten.GetValueOrDefault(0);
                enrolmentE = enrolmentRecord.caps_SumofElementary.GetValueOrDefault(0);
                enrolmentS = enrolmentRecord.caps_SumofSecondary.GetValueOrDefault(0);

                tracingService.Trace("Line: {0}", "49");
            }

            //Get Operating Design Capacity factors for the school district
            var schoolDistrictRecord = service.Retrieve(edu_schooldistrict.EntityLogicalName, facilityRecord.caps_SchoolDistrict.Id, new ColumnSet(true)) as edu_schooldistrict;

            //Get Enrolment Projections
            QueryExpression enrolmentQuery = new QueryExpression("caps_enrolmentprojections_sd");
            enrolmentQuery.ColumnSet.AllColumns = true;

            enrolmentQuery.Criteria = new FilterExpression();
            enrolmentQuery.Criteria.AddCondition("caps_facility", ConditionOperator.Equal, recordId);
            enrolmentQuery.Criteria.AddCondition("statuscode", ConditionOperator.Equal, 200870000);

            EntityCollection enrolmentProjectionRecords = service.RetrieveMultiple(enrolmentQuery);

            var enrolmentProjectionList = enrolmentProjectionRecords.Entities.Select(r => r.ToEntity<caps_EnrolmentProjections_SD>()).ToList();

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

                    //For each record, get top 1 Facility History record where created on date < year end date, order by created on date descending
                    var schoolYearEndDate = ((AliasedValue)reportingRecord["year.edu_enddate"]).Value as DateTime?;

                    QueryExpression snapshotQuery = new QueryExpression("caps_facilityhistory");
                    snapshotQuery.ColumnSet.AllColumns = true;

                    snapshotQuery.Criteria = new FilterExpression();
                    snapshotQuery.Criteria.AddCondition("caps_facility", ConditionOperator.Equal, recordId);
                    snapshotQuery.Criteria.AddCondition("createdon", ConditionOperator.LessEqual, schoolYearEndDate);
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
                    var recordToUpdate = new caps_CapacityReporting();
                    recordToUpdate.Id = reportingRecord.Id;

                    //get matching enrolment projection record
                    var matchingProjection = enrolmentProjectionList.FirstOrDefault(r => r.caps_ProjectionYear.Id == reportingRecord.GetAttributeValue<EntityReference>("caps_schoolyear").Id);

                    recordToUpdate.caps_Kindergarten_designcapacity = facilityRecord.caps_AdjustedDesignCapacityKindergarten;
                    recordToUpdate.caps_Elementary_designcapacity = facilityRecord.caps_AdjustedDesignCapacityElementary;
                    recordToUpdate.caps_Secondary_designcapacity = facilityRecord.caps_AdjustedDesignCapacitySecondary;

                    recordToUpdate.caps_Kindergarten_operatingcapacity = facilityRecord.caps_OperatingCapacityKindergarten;
                    recordToUpdate.caps_Elementary_operatingcapacity = facilityRecord.caps_OperatingCapacityElementary;
                    recordToUpdate.caps_Secondary_operatingcapacity = facilityRecord.caps_OperatingCapacitySecondary;

                    recordToUpdate.caps_Kindergarten_enrolment = (matchingProjection != null) ? matchingProjection.caps_EnrolmentProjectionKindergarten : enrolmentK;
                    recordToUpdate.caps_Elementary_enrolment = (matchingProjection != null) ? matchingProjection.caps_EnrolmentProjectionElementary : enrolmentE;
                    recordToUpdate.caps_Secondary_enrolment = (matchingProjection != null) ? matchingProjection.caps_EnrolmentProjectionSecondary : enrolmentS;

                    service.Update(recordToUpdate); 
                    #endregion
                }
            }


            //CURRENT
            //Get Capacity Reporting records with year details for year marked Current
            //TODO LATER: Get current enrolment data
            //Get all Projects

            //FUTURE
            //Get Capacity Reporting records with year details for years marked Future
        }
    }
}
