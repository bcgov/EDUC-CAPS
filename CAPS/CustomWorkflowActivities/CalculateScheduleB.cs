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
    public class CalculateScheduleB : CodeActivity
    {
        [Output("Total")]
        public OutArgument<decimal> total { get; set; }

        [Output("Error")]
        public OutArgument<bool> error { get; set; }

        protected override void Execute(CodeActivityContext executionContext)
        {
            ITracingService tracingService = executionContext.GetExtension<ITracingService>();

            IWorkflowContext context = executionContext.GetExtension<IWorkflowContext>();
            IOrganizationServiceFactory serviceFactory = executionContext.GetExtension<IOrganizationServiceFactory>();
            IOrganizationService service = serviceFactory.CreateOrganizationService(context.UserId);

            tracingService.Trace("{0}{1}", "Start Custom Workflow Activity: CalculateScheduleB", DateTime.Now.ToLongTimeString());

            var recordId = context.PrimaryEntityId;

            tracingService.Trace("{0}{1}", "Regarding Record ID: ", recordId);

            var columns = new ColumnSet("caps_schooltype"
                                        , "caps_projecttype"
                                        , "caps_nlc"
                                        , "caps_communitylocation"
                                        , "caps_municipalfees"
                                        , "caps_constructioncostsnonstructuralseismicupgr"
                                        , "caps_constructioncostsspir"
                                        , "caps_constructioncostsspiradjustments"
                                        , "caps_seismicprojectidentificationreportfees"
                                        , "caps_hostschooldistrict");

            var projectRequestRecord = service.Retrieve(context.PrimaryEntityName, context.PrimaryEntityId, columns) as caps_Project;

            var projectTypeRecord = service.Retrieve(projectRequestRecord.caps_ProjectType.LogicalName, projectRequestRecord.caps_ProjectType.Id, new ColumnSet("caps_budgetcalculationtype")) as caps_ProjectType;

            var communityLocationRecord = service.Retrieve(projectRequestRecord.caps_CommunityLocation.LogicalName, projectRequestRecord.caps_CommunityLocation.Id, new ColumnSet("caps_projectlocationfactor")) as caps_BudgetCalc_CommunityLocation;

            var hostSchoolDistrictRecord = service.Retrieve(projectRequestRecord.caps_HostSchoolDistrict.LogicalName, projectRequestRecord.caps_HostSchoolDistrict.Id, new ColumnSet("caps_freightrateallowance")) as edu_schooldistrict;

            tracingService.Trace("CalculateScheduleB: {0}", "Populate Variable");
            var scheduleB = new Services.ScheduleB(service, tracingService);
            //set parameters
            scheduleB.SchoolType = projectRequestRecord.caps_SchoolType.Id;
            scheduleB.BudgetCalculationType = projectTypeRecord.caps_BudgetCalculationType.Value;
            scheduleB.NLC = projectRequestRecord.caps_NLC;
            scheduleB.ProjectLocationFactor = communityLocationRecord.caps_ProjectLocationFactor.GetValueOrDefault(1);

            scheduleB.IncreaseInDesignCapacity = new Services.DesignCapacity(40, 200, 0);
            scheduleB.ExistingDesignCapacity = new Services.DesignCapacity(0, 0, 0);
            scheduleB.SeismicUpgradeSpaceAllocation = projectRequestRecord.caps_SeismicUpgradeSpaceAllocation.GetValueOrDefault(0);

            scheduleB.MunicipalFees = projectRequestRecord.caps_MunicipalFees.GetValueOrDefault(0);
            scheduleB.ConstructionNonStructuralSeismicUpgrade = projectRequestRecord.caps_ConstructionCostsNonStructuralSeismicUpgr;
            scheduleB.ConstructionSeismicUpgrade = projectRequestRecord.caps_ConstructionCostsSPIR;
            scheduleB.ConstructionSPIRAdjustment = projectRequestRecord.caps_ConstructionCostsSPIRAdjustments;

            scheduleB.SPIRFees = projectRequestRecord.caps_SeismicProjectIdentificationReportFees;

            scheduleB.FreightRateAllowance = hostSchoolDistrictRecord.caps_FreightRateAllowance.GetValueOrDefault(0);
            //scheduleB.Community

            //call Calculate
            tracingService.Trace("CalculateScheduleB: {0}", "Call Calculate Function");
            var total = scheduleB.Calculate();

            //Update Project Request
            this.total.Set(executionContext, total);

        }
    }
}
