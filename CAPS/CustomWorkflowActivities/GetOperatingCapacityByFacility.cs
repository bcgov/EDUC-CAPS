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
    /// Calculates the current operating capacity for a facility.
    /// </summary>
    public class GetOperatingCapacityByFacility : CodeActivity
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

            tracingService.Trace("{0}{1}", "Start Custom Workflow Activity: GetOperatingCapacityByFacility", DateTime.Now.ToLongTimeString());

            var recordId = context.PrimaryEntityId;

            try
            {
                //get Facility
                var columns = new ColumnSet("caps_adjusteddesigncapacitykindergarten",
                                        "caps_adjusteddesigncapacityelementary",
                                        "caps_adjusteddesigncapacitysecondary",
                                        "caps_lowestgrade",
                                        "caps_highestgrade");

                var facilityRecord = service.Retrieve(context.PrimaryEntityName, context.PrimaryEntityId, columns) as caps_Facility;

                var capacity = new Services.CapacityFactors(service);

                Services.OperatingCapacity capacityService = new Services.OperatingCapacity(service, tracingService, capacity);

                var kDesign = facilityRecord.caps_AdjustedDesignCapacityKindergarten.GetValueOrDefault(0);
                var eDesign = facilityRecord.caps_AdjustedDesignCapacityElementary.GetValueOrDefault(0);
                var sDesign = facilityRecord.caps_AdjustedDesignCapacitySecondary.GetValueOrDefault(0);
                var lowestGrade = facilityRecord.caps_LowestGrade.Value;
                var highestGrade = facilityRecord.caps_HighestGrade.Value;

                var result = capacityService.Calculate(kDesign, eDesign, sDesign, lowestGrade, highestGrade);

                //Update Facility
                var recordToUpdate = new caps_Facility();
                recordToUpdate.Id = context.PrimaryEntityId;
                recordToUpdate.caps_OperatingCapacityKindergarten = result.KindergartenCapacity;
                recordToUpdate.caps_OperatingCapacityElementary = result.ElementaryCapacity;
                recordToUpdate.caps_OperatingCapacitySecondary = result.SecondaryCapacity;
                service.Update(recordToUpdate);

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
