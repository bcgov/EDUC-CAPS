using CAPS.DataContext;
using CustomWorkflowActivities.Services;
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
    /// Calculates the cost of a full replacement using the preliminary budget functionality.  Takes into account location, design capacity and NLC.
    /// </summary>
    public class CalculateFacilityReplacementCost : CodeActivity
    {
        [Output("Error")]
        public OutArgument<bool> error { get; set; }

        [Output("ErrorMessage")]
        public OutArgument<string> errorMessage { get; set; }

        protected override void Execute(CodeActivityContext executionContext)
        {
            ITracingService tracingService = executionContext.GetExtension<ITracingService>();

            IWorkflowContext context = executionContext.GetExtension<IWorkflowContext>();
            IOrganizationServiceFactory serviceFactory = executionContext.GetExtension<IOrganizationServiceFactory>();
            IOrganizationService service = serviceFactory.CreateOrganizationService(context.UserId);

            tracingService.Trace("{0}{1}", "Start Custom Workflow Activity: CalculateFacilityReplacementCost", DateTime.Now.ToLongTimeString());

            var recordId = context.PrimaryEntityId;

            try
            {
                //Get all facilities for the school district
                FilterExpression filterName = new FilterExpression();
                filterName.Conditions.Add(new ConditionExpression("caps_schooldistrict", ConditionOperator.Equal, recordId));
                filterName.Conditions.Add(new ConditionExpression("caps_isschool", ConditionOperator.Equal, true));
                filterName.Conditions.Add(new ConditionExpression("statecode", ConditionOperator.Equal, 0));
                filterName.Conditions.Add(new ConditionExpression("caps_geographicalschooldistrict", ConditionOperator.NotNull));
                filterName.Conditions.Add(new ConditionExpression("caps_communitylocation", ConditionOperator.NotNull));

                QueryExpression query = new QueryExpression("caps_facility");
                query.ColumnSet.AddColumns("caps_communitylocation"
                                            , "caps_designcapacitykindergarten"
                                            , "caps_strongstartcapacitykindergarten"
                                            , "caps_designcapacityelementary"
                                            , "caps_strongstartcapacityelementary"
                                            , "caps_designcapacitysecondary"
                                            , "caps_strongstartcapacitysecondary"
                                            , "caps_currentfacilitytype"
                                            , "caps_geographicalschooldistrict"
                                            , "caps_neighbourhoodlearningcentre");

                query.Criteria.AddFilter(filterName);

                EntityCollection results = service.RetrieveMultiple(query);

                tracingService.Trace("Line: {0}", "58");

                foreach (caps_Facility facility in results.Entities)
                {
                    var communityLocationRecord = service.Retrieve(facility.caps_CommunityLocation.LogicalName, facility.caps_CommunityLocation.Id, new ColumnSet("caps_projectlocationfactor")) as caps_BudgetCalc_CommunityLocation;
                    tracingService.Trace("Line: {0}", "63");
                    var hostSchoolDistrictRecord = service.Retrieve(facility.caps_GeographicalSchoolDistrict.LogicalName, facility.caps_GeographicalSchoolDistrict.Id, new ColumnSet("caps_freightrateallowance")) as edu_schooldistrict;
                    tracingService.Trace("Line: {0}", "65");
                    var facilityTypeRecord = service.Retrieve(facility.caps_CurrentFacilityType.LogicalName, facility.caps_CurrentFacilityType.Id, new ColumnSet("caps_schooltype")) as caps_FacilityType;
                    tracingService.Trace("Line: {0}", "67");
                    var adjustedDesignK = (int)facility.caps_DesignCapacityKindergarten.GetValueOrDefault(0) + facility.caps_StrongStartCapacityKindergarten.GetValueOrDefault(0);
                    var adjustedDesignE = (int)facility.caps_DesignCapacityElementary.GetValueOrDefault(0) + facility.caps_StrongStartCapacityElementary.GetValueOrDefault(0);
                    var adjustedDesignS = (int)facility.caps_DesignCapacitySecondary.GetValueOrDefault(0) + facility.caps_StrongStartCapacitySecondary.GetValueOrDefault(0);

                    tracingService.Trace("Line: {0}", "72");
                    var scheduleB = new ScheduleB(service, tracingService);
                    //set parameters
                    scheduleB.SchoolType = facilityTypeRecord.caps_SchoolType.Id;
                    scheduleB.BudgetCalculationType = (int)caps_BudgetCalculationType.Replacement;
                    scheduleB.IncludeNLC = facility.caps_NeighbourhoodLearningCentre.GetValueOrDefault(false);
                    scheduleB.ProjectLocationFactor = communityLocationRecord.caps_ProjectLocationFactor.GetValueOrDefault(1);

                    scheduleB.ExistingAndDecreaseDesignCapacity = new Services.DesignCapacity(0, 0, 0);
                    scheduleB.ExtraSpaceAllocation = 0;
                    scheduleB.ApprovedDesignCapacity = new Services.DesignCapacity(adjustedDesignK, adjustedDesignE, adjustedDesignS);

                    scheduleB.MunicipalFees = 0;
                    //scheduleB.ConstructionNonStructuralSeismicUpgrade = 0;
                    scheduleB.ConstructionSeismicUpgrade = 0;
                    //scheduleB.ConstructionSPIRAdjustment = 0;

                    tracingService.Trace("Line: {0}", "89");
                    //scheduleB.SPIRFees = 0;

                    scheduleB.FreightRateAllowance = hostSchoolDistrictRecord.caps_FreightRateAllowance.GetValueOrDefault(0);

                    //Supplemental Items
                    scheduleB.Demolition = 0;
                    scheduleB.AbnormalTopography = 0;
                    scheduleB.TempAccommodation = 0;
                    scheduleB.OtherSupplemental = 0;

                    //call Calculate
                    tracingService.Trace("CalculateScheduleB: {0}", "Call Calculate Function");
                    CalculationResult result = scheduleB.Calculate();

                    tracingService.Trace("Line: {0}", "104");
                    //Update the facility
                    var recordToUpdate = new caps_Facility();
                    recordToUpdate.Id = facility.Id;
                    recordToUpdate.caps_PreliminaryReplacementValue = result.Total;
                    recordToUpdate.caps_PreliminaryReplacementCalculatedOn = DateTime.Now;
                    service.Update(recordToUpdate);
                

                }

                //Update Project Request
                this.error.Set(executionContext, false);
            }
            catch (Exception ex)
            {
                tracingService.Trace("Error Details: {0}", ex.Message);
                //might want to also include error message
                this.error.Set(executionContext, true);
                this.errorMessage.Set(executionContext, ex.Message);
            }

        }
    }
}
