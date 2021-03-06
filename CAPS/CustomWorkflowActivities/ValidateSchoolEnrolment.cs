﻿using CAPS.DataContext;
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
    /// Called on School Enrolment Collection via Workflow, this CWA attempts to find matching facility records.
    /// </summary>
    public class ValidateSchoolEnrolment : CodeActivity
    {
        protected override void Execute(CodeActivityContext executionContext)
        {
            ITracingService tracingService = executionContext.GetExtension<ITracingService>();

            IWorkflowContext context = executionContext.GetExtension<IWorkflowContext>();
            IOrganizationServiceFactory serviceFactory = executionContext.GetExtension<IOrganizationServiceFactory>();
            IOrganizationService service = serviceFactory.CreateOrganizationService(context.UserId);

            tracingService.Trace("{0}{1}", "Start Custom Workflow Activity: ValidateSchoolEnrolment", DateTime.Now.ToLongTimeString());

            var recordId = context.PrimaryEntityId;

            bool isCollectionValid = true;
            string collectionErrorMessage = string.Empty;

            //Get all facilities in the system
            //FilterExpression filterFacilities = new FilterExpression();
            //filterFacilities.FilterOperator = LogicalOperator.Or;
            //filterFacilities.Conditions.Add(new ConditionExpression("caps_sldfacilitycode", ConditionOperator.Equal, record.caps_MinistryCode));
            //filterFacilities.Conditions.Add(new ConditionExpression("caps_ministryassetnumber", ConditionOperator.Equal, record.caps_AssetNumber));

            QueryExpression queryFacility = new QueryExpression("caps_facility");
            queryFacility.ColumnSet.AddColumns("caps_sldfacilitycode", "caps_ministryassetnumber", "statecode");
            //queryFacility.Criteria.AddFilter(filterFacilities);

            EntityCollection facilityRecords = service.RetrieveMultiple(queryFacility);

            var arrayFacilities = facilityRecords.Entities.Select(r => r.ToEntity<caps_Facility>()).ToArray();

            //Get all School Enrolment records for the collection
            FilterExpression filterName = new FilterExpression();
            filterName.Conditions.Add(new ConditionExpression("caps_schoolenrolmentcollection", ConditionOperator.Equal, recordId));


            QueryExpression query = new QueryExpression("edu_schoolenrolment");
            query.ColumnSet.AddColumns("caps_ministrycode", "caps_assetnumber");

            query.Criteria.AddFilter(filterName);

            EntityCollection records = service.RetrieveMultiple(query);

            //var records = crmContext.edu_SchoolEnrolmentSet.Where(r=>r.caps_SchoolEnrolmentCollection.Id == recordId);

            //for each school enrolment
            foreach(edu_SchoolEnrolment record in records.Entities)
            {
                bool isSchoolRecordValid = true;
                string schoolRecordErrorMessage = string.Empty;
                Guid facilityId = new Guid();
                //Get facilities where school code matches or where ministry asset number matches

                /*
                FilterExpression filterFacilities = new FilterExpression();
                filterFacilities.FilterOperator = LogicalOperator.Or;
                filterFacilities.Conditions.Add(new ConditionExpression("caps_sldfacilitycode", ConditionOperator.Equal, record.caps_MinistryCode));
                filterFacilities.Conditions.Add(new ConditionExpression("caps_ministryassetnumber", ConditionOperator.Equal, record.caps_AssetNumber));

                QueryExpression queryFacility = new QueryExpression("caps_facility");
                queryFacility.ColumnSet.AddColumns("caps_sldfacilitycode", "caps_ministryassetnumber", "statecode");
                queryFacility.Criteria.AddFilter(filterFacilities);

                EntityCollection facilityRecords = service.RetrieveMultiple(queryFacility);
                */

                var facilities = arrayFacilities.Where(r => r.caps_SLDFacilityCode == record.caps_MinistryCode || r.caps_MinistryAssetNumber == record.caps_AssetNumber).AsEnumerable();

                //var facilities = crmContext.caps_FacilitySet.Where(r=>r.caps_SLDFacilityCode == record.caps_MinistryCode || r.caps_MinistryAssetNumber == record.caps_AssetNumber).AsEnumerable();

                if (facilities.Count() == 0)
                {
                    isSchoolRecordValid = false;
                    schoolRecordErrorMessage = "No Match Found.";
                }
                else if (facilities.Count() > 1)
                {
                    isSchoolRecordValid = false;
                    schoolRecordErrorMessage = "Multiple Facility Matches Found.  Check both ministry code and asset number.";
                }
                else if (facilities.Count() == 1)
                {
                    var facility = facilities.First(); //facilities.First();
                    facilityId = facility.Id;

                    if (facility.caps_SLDFacilityCode != record.caps_MinistryCode)
                    {
                        isSchoolRecordValid = false;
                        schoolRecordErrorMessage = "Asset # match but Ministry Code Different";
                    }
                    else if (facility.caps_MinistryAssetNumber != record.caps_AssetNumber)
                    {
                        isSchoolRecordValid = false;
                        schoolRecordErrorMessage = "Ministry Code match but Asset # Different.";
                    }
                    else if ((int)facility.StateCode.Value == 1)
                    {
                        isSchoolRecordValid = false;
                        schoolRecordErrorMessage = "Status Missmatch. Facility is Inactive but Active in SLD.";
                    }
                }

                if (schoolRecordErrorMessage.Length > 100)
                {
                    schoolRecordErrorMessage = schoolRecordErrorMessage.Substring(0, 100);
                }


                var recordToUpdate = new edu_SchoolEnrolment();
                recordToUpdate.Id = record.Id;
                if (!isSchoolRecordValid) {
                    recordToUpdate.caps_ValidationStatus = new OptionSetValue(100000001);
                    recordToUpdate.caps_ValidationReason = schoolRecordErrorMessage;
                }
                else
                {
                    recordToUpdate.caps_ValidationStatus = new OptionSetValue(100000000);
                    recordToUpdate.caps_ValidationReason = schoolRecordErrorMessage;
                }
                if (facilityId != Guid.Empty)
                {
                    recordToUpdate.caps_Facility = new EntityReference(caps_Facility.EntityLogicalName, facilityId);
                }

                service.Update(recordToUpdate);

                if(!isSchoolRecordValid)
                {
                    isCollectionValid = false;
                    collectionErrorMessage = "One or more School Enrolment records failed.\r\n";
                }
            }

            //Now see if any facilities are missing from the list
            var fetchXml = "<fetch version=\"1.0\" output-format=\"xml-platform\" mapping=\"logical\" distinct=\"true\">" +
                            "<entity name=\"caps_facility\" > " +
                                "<attribute name=\"caps_facilityid\" /> " +
                                "<attribute name=\"caps_name\" /> " +
                                    "<attribute name=\"createdon\" /> " +
                                    "<order attribute=\"caps_name\" descending=\"false\" /> " +
                                        "<filter type=\"and\" > " +
                                        "<condition attribute=\"statecode\" operator=\"eq\" value=\"0\" /> " +
                                        "<condition attribute=\"caps_isschool\" operator=\"eq\" value=\"1\" /> " +
                                            "</filter> " +
                                            "<link-entity name=\"edu_schoolenrolment\" from=\"caps_facility\" to=\"caps_facilityid\" link-type=\"outer\" alias=\"ab\" /> " +
                        "<filter type=\"and\" > " +
                            "<condition entityname=\"ab\" attribute=\"caps_facility\" operator=\"null\" /> " +
                                "</filter> " +
                            "</entity> " +
                            "</fetch> ";

            var missingFacilities = service.RetrieveMultiple(new FetchExpression(fetchXml));

            if (missingFacilities.Entities.Count > 0)
            {
                isCollectionValid = false;

                collectionErrorMessage += "The following facilities were not included in the school enrolment data:\r\n";

                foreach(var facility in missingFacilities.Entities)
                {
                    collectionErrorMessage += facility.GetAttributeValue<string>("caps_name")+"\r\n";
                }

            }

            if (collectionErrorMessage.Length > 4000)
            {
                collectionErrorMessage = collectionErrorMessage.Substring(0, 4000);
            }

            //Update Collection Record
            var collectionToUpdate = new caps_SchoolEnrolmentCollection();
            collectionToUpdate.Id = recordId;
            collectionToUpdate.caps_ValidationStatus = (isCollectionValid) ? new OptionSetValue(100000000) : new OptionSetValue(100000001);
            collectionToUpdate.caps_ValidationReason = collectionErrorMessage;

            service.Update(collectionToUpdate);

            tracingService.Trace("{0}{1}", "END Custom Workflow Activity: ValidateSchoolEnrolment", DateTime.Now.ToLongTimeString());

        }
        
    }
}
