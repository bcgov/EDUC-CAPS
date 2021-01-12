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
    /// Called from action on Project Request.  This CWA calculates the schedule b cost and returns the amount as total.
    /// </summary>
    public class CalculateScheduleB : CodeActivity
    {
        [Output("Total")]
        public OutArgument<decimal> total { get; set; }

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

            tracingService.Trace("{0}{1}", "Start Custom Workflow Activity: CalculateScheduleB", DateTime.Now.ToLongTimeString());

            var recordId = context.PrimaryEntityId;

            try
            {
                decimal total = RunCalculation(tracingService, context, service, recordId);

                //Update Project Request
                this.total.Set(executionContext, total);
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

        internal static decimal RunCalculation(ITracingService tracingService, IWorkflowContext context, IOrganizationService service, Guid recordId)
        {
            tracingService.Trace("{0}{1}", "Regarding Record ID: ", recordId);

            var columns = new ColumnSet("caps_schooltype"
                                        , "caps_projecttype"
                                        , "caps_includenlc"
                                        , "caps_communitylocation"
                                        , "caps_municipalfees"
                                        , "caps_constructioncostsnonstructuralseismicupgr"
                                        , "caps_constructioncostsspir"
                                        , "caps_constructioncostsspiradjustments"
                                        , "caps_seismicprojectidentificationreportfees"
                                        , "caps_hostschooldistrict"
                                        , "caps_changeindesigncapacitykpositive"
                                        , "caps_changeindesigncapacityepositive"
                                        , "caps_changeindesigncapacityspositive"
                                        , "caps_changeindesigncapacityknegative"
                                        , "caps_changeindesigncapacityenegative"
                                        , "caps_changeindesigncapacitysnegative"
                                        , "caps_facility"
                                        , "caps_demolitioncost"
                                        , "caps_abnormaltopographycost"
                                        , "caps_temporaryaccommodationcost"
                                        , "caps_othercost"
                                        , "caps_schbadditionalspaceallocation");



            var projectRequestRecord = service.Retrieve(context.PrimaryEntityName, context.PrimaryEntityId, columns) as caps_Project;

            var projectTypeRecord = service.Retrieve(projectRequestRecord.caps_ProjectType.LogicalName, projectRequestRecord.caps_ProjectType.Id, new ColumnSet("caps_budgetcalculationtype")) as caps_ProjectType;

            var communityLocationRecord = service.Retrieve(projectRequestRecord.caps_CommunityLocation.LogicalName, projectRequestRecord.caps_CommunityLocation.Id, new ColumnSet("caps_projectlocationfactor")) as caps_BudgetCalc_CommunityLocation;

            var hostSchoolDistrictRecord = service.Retrieve(projectRequestRecord.caps_HostSchoolDistrict.LogicalName, projectRequestRecord.caps_HostSchoolDistrict.Id, new ColumnSet("caps_freightrateallowance")) as edu_schooldistrict;

            int adjustedDesignK = 0;
            int adjustedDesignE = 0;
            int adjustedDesignS = 0;
            if (projectRequestRecord.caps_Facility != null)
            {
                tracingService.Trace("Facility: {0}", projectRequestRecord.caps_Facility.Id);

                var facilityColumnSet = new ColumnSet("caps_designcapacitykindergarten"
                                                        , "caps_strongstartcapacitykindergarten"
                                                        , "caps_designcapacityelementary"
                                                        , "caps_strongstartcapacityelementary"
                                                        , "caps_designcapacitysecondary"
                                                        , "caps_strongstartcapacitysecondary");
                var facilityRecord = service.Retrieve(projectRequestRecord.caps_Facility.LogicalName, projectRequestRecord.caps_Facility.Id, new ColumnSet(true)) as caps_Facility;

                adjustedDesignK = (int)facilityRecord.caps_DesignCapacityKindergarten.GetValueOrDefault(0) + facilityRecord.caps_StrongStartCapacityKindergarten.GetValueOrDefault(0);
                adjustedDesignE = (int)facilityRecord.caps_DesignCapacityElementary.GetValueOrDefault(0) + facilityRecord.caps_StrongStartCapacityElementary.GetValueOrDefault(0);
                adjustedDesignS = (int)facilityRecord.caps_DesignCapacitySecondary.GetValueOrDefault(0) + facilityRecord.caps_StrongStartCapacitySecondary.GetValueOrDefault(0);
            }

            int increaseDesignK = (int)projectRequestRecord.caps_ChangeinDesignCapacityKPositive.GetValueOrDefault(0);
            int increaseDesignE = (int)projectRequestRecord.caps_ChangeinDesignCapacityEPositive.GetValueOrDefault(0);
            int increaseDesignS = (int)projectRequestRecord.caps_ChangeinDesignCapacitySPositive.GetValueOrDefault(0);

            int decreaseDesignK = (int)projectRequestRecord.caps_ChangeinDesignCapacityKNegative.GetValueOrDefault(0);
            int decreaseDesignE = (int)projectRequestRecord.caps_ChangeinDesignCapacityENegative.GetValueOrDefault(0);
            int decreaseDesignS = (int)projectRequestRecord.caps_ChangeinDesignCapacitySNegative.GetValueOrDefault(0);

            int subtotalDesignK = adjustedDesignK + decreaseDesignK;
            int subtotalDesignE = adjustedDesignE + decreaseDesignE;
            int subtotalDesignS = adjustedDesignS + decreaseDesignS;

            tracingService.Trace("CalculateScheduleB: {0}", "Populate Variable");

            var scheduleB = new ScheduleB(service, tracingService);
            //set parameters
            scheduleB.SchoolType = projectRequestRecord.caps_SchoolType.Id;
            scheduleB.BudgetCalculationType = projectTypeRecord.caps_BudgetCalculationType.Value;
            scheduleB.IncludeNLC = projectRequestRecord.caps_IncludeNLC.GetValueOrDefault(false);
            scheduleB.ProjectLocationFactor = communityLocationRecord.caps_ProjectLocationFactor.GetValueOrDefault(1);

            tracingService.Trace("Facility - K:{0} E:{1} S:{2}", adjustedDesignK, adjustedDesignE, adjustedDesignS);
            tracingService.Trace("Subtotal - K:{0} E:{1} S:{2}", subtotalDesignK, subtotalDesignE, subtotalDesignS);
            tracingService.Trace("Approved - K:{0} E:{1} S:{2}", subtotalDesignK + increaseDesignK, subtotalDesignE + increaseDesignE, subtotalDesignS + increaseDesignS);

            scheduleB.ExistingAndDecreaseDesignCapacity = new Services.DesignCapacity(subtotalDesignK, subtotalDesignE, subtotalDesignS);
            scheduleB.ExtraSpaceAllocation = projectRequestRecord.caps_SchBAdditionalSpaceAllocation;
            scheduleB.ApprovedDesignCapacity = new Services.DesignCapacity(subtotalDesignK + increaseDesignK, subtotalDesignE + increaseDesignE, subtotalDesignS + increaseDesignS);

            scheduleB.MunicipalFees = projectRequestRecord.caps_MunicipalFees.GetValueOrDefault(0);
            scheduleB.ConstructionNonStructuralSeismicUpgrade = projectRequestRecord.caps_ConstructionCostsNonStructuralSeismicUpgr;
            scheduleB.ConstructionSeismicUpgrade = projectRequestRecord.caps_ConstructionCostsSPIR;
            scheduleB.ConstructionSPIRAdjustment = projectRequestRecord.caps_ConstructionCostsSPIRAdjustments;

            scheduleB.SPIRFees = projectRequestRecord.caps_SeismicProjectIdentificationReportFees;

            scheduleB.FreightRateAllowance = hostSchoolDistrictRecord.caps_FreightRateAllowance.GetValueOrDefault(0);

            //Supplemental Items
            scheduleB.Demolition = projectRequestRecord.caps_DemolitionCost.GetValueOrDefault(0);
            scheduleB.AbnormalTopography = projectRequestRecord.caps_AbnormalTopographyCost.GetValueOrDefault(0);
            scheduleB.TempAccommodation = projectRequestRecord.caps_TemporaryAccommodationCost.GetValueOrDefault(0);
            scheduleB.OtherSupplemental = projectRequestRecord.caps_OtherCost.GetValueOrDefault(0);

            //call Calculate
            tracingService.Trace("CalculateScheduleB: {0}", "Call Calculate Function");
            CalculationResult result = scheduleB.Calculate();

            //Update PRFS Option with Calculations
            var recordToUpdate = new caps_Project();
            recordToUpdate.Id = recordId;
            //Section 2
            recordToUpdate.caps_SchBSpaceAllocationNewReplacement = result.SpaceAllocationNewReplacement;
            //recordToUpdate.caps_SchBSpaceAllocationNLC = result.SpaceAllocationNLC;

            //Section 3
            recordToUpdate.caps_SchBBaseBudgetRate = result.BaseBudgetRate;
            recordToUpdate.caps_SchBProjectSizeFactor = result.ProjectSizeFactor;
            recordToUpdate.caps_SchBProjectLocationFactor = result.ProjectLocationFactor;
            recordToUpdate.caps_SchBUnitRate = result.UnitRate;

            //Section 4
            recordToUpdate.caps_SchBConstructionNewSpaceReplacement = result.ConstructionNewReplacement;
            recordToUpdate.caps_SchBConstructionRenovations = result.ConstructionRenovation;
            recordToUpdate.caps_SchBSiteDevelopmentAllowance = result.SiteDevelopmentAllowance;
            recordToUpdate.caps_SchBSiteDevelopmentLocationAllowance = result.SiteDevelopmentLocationAllowance;

            //Section 5
            recordToUpdate.caps_schbdesignfees = result.DesignFees;
            recordToUpdate.caps_SchBPostContractNewReplacement = result.PostContractNewReplacement;
            recordToUpdate.caps_SchBPostContractRenovations = result.PostContractRenovation;
            recordToUpdate.caps_SchBPostContractSeismic = result.PostContractSeismic;
            //recordToUpdate.caps_SchBMunicipalFees = result.MunicipalFees;
            recordToUpdate.caps_SchBEquipmentNew = result.EquipmentNew;
            recordToUpdate.caps_SchBEquipmentReplacement = result.EquipmentReplacement;
            recordToUpdate.caps_SchBProjectManagementFees = result.ProjectManagement;
            recordToUpdate.caps_SchBLiabilityInsurance = result.LiabilityInsurance;
            recordToUpdate.caps_SchBPayableTaxes = result.PayableTaxes;
            recordToUpdate.caps_SchBRiskReserve = result.RiskReserve;
            recordToUpdate.caps_SchBNLCBudgetAmount = result.NLCBudgetAmount;

            service.Update(recordToUpdate);


            return result.Total;
        }
    }
}
