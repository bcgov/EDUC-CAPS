using CAPS.DataContext;
using Microsoft.Xrm.Sdk;
using Microsoft.Xrm.Sdk.Query;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace CustomWorkflowActivities.Services
{
    internal class CapacityFactors
    {
        internal decimal DesignKindergarten { get; set; }
        internal decimal DesignElementary { get; set; }
        internal decimal DesignSecondary { get; set; }
        internal decimal OperatingKindergarten { get; set; }
        internal decimal OperatingLowElementary { get; set; }
        internal decimal OperatingHighElementary { get; set; }
        internal decimal OperatingSecondary { get; set; }

        internal CapacityFactors(IOrganizationService service) {
            DesignKindergarten = GetBudgetCalculationValue(service, "Design Capacity Kindergarten");
            DesignElementary = GetBudgetCalculationValue(service, "Design Capacity Elementary");
            DesignSecondary = GetBudgetCalculationValue(service, "Design Capacity Secondary");

            OperatingKindergarten = GetBudgetCalculationValue(service, "Kindergarten Operating Capacity");
            OperatingLowElementary = GetBudgetCalculationValue(service, "Elementary Lower Operating Capacity");
            OperatingHighElementary = GetBudgetCalculationValue(service, "Elementary Upper Operating Capacity");
            OperatingSecondary = GetBudgetCalculationValue(service, "Secondary Operating Capacity");
        }

        internal CapacityFactors(decimal design_k, decimal design_e, decimal design_s, decimal operating_k, decimal operating_low_e, decimal operating_high_e, decimal operating_s)
        {
            DesignKindergarten = design_k;
            DesignElementary = design_e;
            DesignSecondary = design_s;

            OperatingKindergarten = operating_k;
            OperatingLowElementary = operating_low_e;
            OperatingHighElementary = operating_high_e;
            OperatingSecondary = operating_s;
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

    internal class OperatingCapacity
    {
        CapacityFactors Capacity { get; set; }

        IOrganizationService service { get; set; }
        ITracingService tracingService { get; set; }

        internal class CalculationResult
        {
            internal decimal KindergartenCapacity { get; set; }
            internal decimal ElementaryCapacity { get; set; }
            internal decimal SecondaryCapacity { get; set; }
        }


        /// <summary>
        /// Default constructor for Schedule B
        /// </summary>
        /// <param name="s">IOrganizationService</param>
        /// <param name="t">ITracingService</param>
        public OperatingCapacity(IOrganizationService s, ITracingService t, CapacityFactors capacity)
        {
            this.service = s;
            this.tracingService = t;
            this.Capacity = capacity;
        }

        public CalculationResult Calculate(decimal DesignKindergarten, decimal DesignElementary, decimal DesignSecondary, int LowestGrade, int HighestGrade)
        {
            tracingService.Trace("{0}", "Starting OperatingCapacity.Calculate");
            tracingService.Trace("K:{0}, E: {1}, S: {2}", DesignKindergarten, DesignElementary, DesignSecondary);
            tracingService.Trace("Grade Low:{0}, High: {1}", LowestGrade, HighestGrade);

            CalculationResult result = new CalculationResult();

            //Calculate Kindergarten Capacity
            if (DesignKindergarten != 0)
            {
                result.KindergartenCapacity = Math.Round(DesignKindergarten / Capacity.DesignKindergarten * Capacity.OperatingKindergarten, 0, MidpointRounding.AwayFromZero);
            }

            //Calculate Secondary Capacity
            if (DesignSecondary != 0)
            {
                result.SecondaryCapacity = Math.Round(DesignSecondary / Capacity.DesignSecondary * Capacity.OperatingSecondary, 0, MidpointRounding.AwayFromZero);
            }

            //Calculate Elementary Capacity
            //Get Grade Range
            int lowGrade = LowestGrade - 100000000;
            int highGrade = HighestGrade - 100000000;

            int lowECount = 0;
            int highECount = 0;

            if (lowGrade < 8 && highGrade > 0)
            {
                //we need to calculate elementary
                //get number of low and high elementary grades
                for (int x= 1; x<=7; x++)
                {
                    if (x >= lowGrade && x <=highGrade)
                    {
                        if (x <=3)
                        {
                            lowECount++;
                        }
                        else
                        {
                            highECount++;
                        }
                    }
                }

                if (lowECount == 0) {
                    // all high E
                    result.ElementaryCapacity = Math.Round(DesignElementary / Capacity.DesignElementary * Capacity.OperatingHighElementary, 0, MidpointRounding.AwayFromZero);
                }
                else if (highECount == 0) {
                    result.ElementaryCapacity = Math.Round(DesignElementary / Capacity.DesignElementary * Capacity.OperatingLowElementary, 0, MidpointRounding.AwayFromZero);
                }
                else {
                    var calcOperatingCapacity = ((lowECount * Capacity.OperatingLowElementary) + (highECount * Capacity.OperatingHighElementary)) / (lowECount + highECount);

                    result.ElementaryCapacity = Math.Round(DesignElementary / Capacity.DesignElementary * calcOperatingCapacity, 0, MidpointRounding.AwayFromZero);
                }
            }

            return result;
        }

    }




}
