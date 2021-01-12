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
    /// <summary>
    /// The CWA runs on Project Request and calculates the change in operating capacity.
    /// </summary>
    public class CalculateProjectRequestOperatingCapacity: CodeActivity
    {
        protected override void Execute(CodeActivityContext executionContext)
        {
            ITracingService tracingService = executionContext.GetExtension<ITracingService>();

            IWorkflowContext context = executionContext.GetExtension<IWorkflowContext>();
            IOrganizationServiceFactory serviceFactory = executionContext.GetExtension<IOrganizationServiceFactory>();
            IOrganizationService service = serviceFactory.CreateOrganizationService(context.UserId);

            tracingService.Trace("{0}{1}", "Start Custom Workflow Activity: CalculateProjectRequestOperatingCapacity", DateTime.Now.ToLongTimeString());

            var recordId = context.PrimaryEntityId;

            var startingDesign_K = 0;
            var startingDesign_E = 0;
            var startingDesign_S = 0;

            var startingOperating_K = 0;
            var startingOperating_E = 0;
            var startingOperating_S = 0;

            //Get Project Request
            var projectRequestColumns = new ColumnSet("caps_facility",
                "caps_changeindesigncapacitykindergarten",
                "caps_changeindesigncapacityelementary",
                "caps_changeindesigncapacitysecondary",
                "caps_futurelowestgrade",
                "caps_futurehighestgrade");
            var projectRequest = service.Retrieve(caps_Project.EntityLogicalName, recordId, projectRequestColumns) as caps_Project;

            if (projectRequest.caps_FutureLowestGrade != null && projectRequest.caps_FutureHighestGrade != null)
            {

                //if facility exists, then retrieve it
                if (projectRequest.caps_Facility != null && projectRequest.caps_Facility.Id != Guid.Empty)
                {
                    var facility = service.Retrieve(caps_Facility.EntityLogicalName, projectRequest.caps_Facility.Id, new ColumnSet("caps_adjusteddesigncapacitykindergarten", "caps_adjusteddesigncapacityelementary", "caps_adjusteddesigncapacitysecondary", "caps_operatingcapacitykindergarten", "caps_operatingcapacityelementary", "caps_operatingcapacitysecondary")) as caps_Facility;

                    if (facility != null)
                    {
                        startingDesign_K = facility.caps_AdjustedDesignCapacityKindergarten.GetValueOrDefault(0);
                        startingDesign_E = facility.caps_AdjustedDesignCapacityElementary.GetValueOrDefault(0);
                        startingDesign_S = facility.caps_AdjustedDesignCapacitySecondary.GetValueOrDefault(0);

                        startingOperating_K = Convert.ToInt32(facility.caps_OperatingCapacityKindergarten.GetValueOrDefault(0));
                        startingOperating_E = Convert.ToInt32(facility.caps_OperatingCapacityElementary.GetValueOrDefault(0));
                        startingOperating_S = Convert.ToInt32(facility.caps_OperatingCapacitySecondary.GetValueOrDefault(0));
                    }
                }

                var changeInDesign_K = startingDesign_K + projectRequest.caps_ChangeinDesignCapacityKindergarten.GetValueOrDefault(0);
                var changeInDesign_E = startingDesign_E + projectRequest.caps_ChangeinDesignCapacityElementary.GetValueOrDefault(0);
                var changeInDesign_S = startingDesign_S + projectRequest.caps_ChangeinDesignCapacitySecondary.GetValueOrDefault(0);

                //Get Operating Capacity Values
                var capacity = new Services.CapacityFactors(service);

                Services.OperatingCapacity capacityService = new Services.OperatingCapacity(service, tracingService, capacity);

                var result = capacityService.Calculate(changeInDesign_K, changeInDesign_E, changeInDesign_S, projectRequest.caps_FutureLowestGrade.Value, projectRequest.caps_FutureHighestGrade.Value);

                var changeInOperating_K = Convert.ToInt32(result.KindergartenCapacity) - startingOperating_K;
                var changeInOperating_E = Convert.ToInt32(result.ElementaryCapacity) - startingOperating_E;
                var changeInOperating_S = Convert.ToInt32(result.SecondaryCapacity) - startingOperating_S;

                var recordToUpdate = new caps_Project();
                recordToUpdate.Id = recordId;
                recordToUpdate.caps_ChangeinOperatingCapacityKindergarten = changeInOperating_K;
                recordToUpdate.caps_ChangeinOperatingCapacityElementary = changeInOperating_E;
                recordToUpdate.caps_ChangeinOperatingCapacitySecondary = changeInOperating_S;
                service.Update(recordToUpdate);
            }
            else {
                //blank out operating capacity
                var recordToUpdate = new caps_Project();
                recordToUpdate.Id = recordId;
                recordToUpdate.caps_ChangeinOperatingCapacityKindergarten = null;
                recordToUpdate.caps_ChangeinOperatingCapacityElementary = null;
                recordToUpdate.caps_ChangeinOperatingCapacitySecondary = null;
                service.Update(recordToUpdate);
            }
        }
    }
}
