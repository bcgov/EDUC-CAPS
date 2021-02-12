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
    /// Validates the design capacity is a multiple of the specified  classroom design capacity in Preliminary Budget > Budget Calculation Factors.
    /// </summary>
    public class ValidateDesignCapacity : CodeActivity
    {
        [Input("Type")]
        public InArgument<string> Type { get; set; }

        [Input("Design Count")]
        public InArgument<int> DesignCount { get; set; }

        [Output("Error")]
        public OutArgument<bool> Error { get; set; }

        [Output("Capacity Guideline")]
        public OutArgument<string> CapacityGuideline { get; set; }

        protected override void Execute(CodeActivityContext executionContext)
        {
            ITracingService tracingService = executionContext.GetExtension<ITracingService>();

            IWorkflowContext context = executionContext.GetExtension<IWorkflowContext>();
            IOrganizationServiceFactory serviceFactory = executionContext.GetExtension<IOrganizationServiceFactory>();
            IOrganizationService service = serviceFactory.CreateOrganizationService(context.UserId);

            tracingService.Trace("{0}{1}", "Start Custom Workflow Activity: ValidateDesignCapacity", DateTime.Now.ToLongTimeString());

            string type = this.Type.Get(executionContext);
            int count = this.DesignCount.Get(executionContext);
            string[] typeArray = { "Kindergarten", "Elementary", "Secondary"};
            if (!typeArray.Contains(type))
            {
                throw new InvalidOperationException("Type is not Kindergarten, Elementary or Secondary", new ArgumentNullException("Type"));
            }

            tracingService.Trace("Type:{0};Count:{1};", type, count);

            //get Design capacity values
            decimal designCapacityGuideline = 0;
            switch (type)
            {
                case "Kindergarten":
                    designCapacityGuideline = GetBudgetCalculationValue(service, "Design Capacity Kindergarten");
                    break;
                case "Elementary":
                    designCapacityGuideline = GetBudgetCalculationValue(service, "Design Capacity Elementary");
                    break;
                case "Secondary":
                    designCapacityGuideline = GetBudgetCalculationValue(service, "Design Capacity Secondary");
                    break;
            }

            tracingService.Trace("Design Guideline:{0}", designCapacityGuideline);

            //take count mod design capacity, if remainder isn't 0 then there is an error
            var remainder = Convert.ToDecimal(count) % designCapacityGuideline;

            if (remainder != 0)
            {
                this.Error.Set(executionContext, true);
                this.CapacityGuideline.Set(executionContext, designCapacityGuideline.ToString("#"));
                //this.ErrorMessage.Set(executionContext, "The design capacity is not divisible by "+designCapacityGuideline);
            }
            else
            {
                this.Error.Set(executionContext, false);
                //this.ErrorMessage.Set(executionContext, "");
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
