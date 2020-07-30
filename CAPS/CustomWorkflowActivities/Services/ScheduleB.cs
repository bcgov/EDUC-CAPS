using System;
using CAPS.DataContext;
using Microsoft.Xrm.Sdk;
using Microsoft.Xrm.Sdk.Query;

namespace CustomWorkflowActivities.Services
{
    internal class DesignCapacity
    {
        internal DesignCapacity() { }
        internal DesignCapacity(int k, int e, int s)
        {
            Kindergarten = k;
            Elementary = e;
            Secondary = s;
        }
        internal int Kindergarten { get; set; }
        internal int Elementary { get; set; }
        internal int Secondary { get; set; }

        internal int Total()
        {
            return Kindergarten + Elementary + Secondary;
        }
    }
    internal class ScheduleB
    {
        internal Guid SchoolType { get; set; } //Required for all
        internal int BudgetCalculationType { get; set; }
        internal DesignCapacity ExistingDesignCapacity { get; set; }
        internal DesignCapacity IncreaseInDesignCapacity { get; set; }
        internal DesignCapacity ApprovedDesignCapacity { get; set; }
        internal decimal SeismicUpgradeSpaceAllocation { get; set; }
        internal decimal ProjectLocationFactor { get; set; }
        internal decimal? NLC { get; set; }
        internal decimal MunicipalFees { get; set; }
        internal Guid Community { get; set; }
        internal decimal? ConstructionSeismicUpgrade { get; set; }
        internal decimal? ConstructionSPIRAdjustment { get; set; }
        internal decimal? ConstructionNonStructuralSeismicUpgrade { get; set; }
        internal decimal? SPIRFees { get; set; } //5.1b
        internal decimal FreightRateAllowance { get; set; }
        internal decimal Demolition { get; set; }
        internal decimal AbnormalTopography { get; set; }
        internal decimal TempAccommodation { get; set; }
        internal decimal OtherSupplemental { get; set; }

        IOrganizationService service { get; set; }
        ITracingService tracingService { get; set; }


        public ScheduleB(IOrganizationService s, ITracingService t)
        {
            this.service = s;
            this.tracingService = t;
        }
        internal decimal Calculate()
        {
            tracingService.Trace("{0}", "Starting ScheduleB.Calculate");

            //variables
            decimal spaceAllocationExisting = 0;
            decimal constructionRenovations = 0;
            decimal constructionSiteDevelopmentAllowance = 0;
            decimal softCostContingencyRenovations = 0;
            decimal softCostContingencySeismic = 0;
            decimal softCostEquipmentAllowance = 0;
            decimal constructionTotalConstructionBudget = 0;
            decimal spaceAllocationNewReplacement = 0;

            //Default some numbers to 0, cased on Budget Calculation type
            //if partial seismic, set NLC to 0
            if (BudgetCalculationType == (int)caps_BudgetCalculationType.PartialSeismic)
            {
                NLC = 0;
            }

            if (BudgetCalculationType == (int)caps_BudgetCalculationType.New ||
                BudgetCalculationType == (int)caps_BudgetCalculationType.Replacement ||
                BudgetCalculationType == (int)caps_BudgetCalculationType.Addition ||
                BudgetCalculationType == (int)caps_BudgetCalculationType.PartialReplacement)
            {
                //default seismic fields to 0
                ConstructionSeismicUpgrade = 0;
                ConstructionSPIRAdjustment = 0;
                ConstructionNonStructuralSeismicUpgrade = 0;
            }

            //Get School Type
            var schoolTypeRecord = GetSchoolType(SchoolType);

            #region 2. Space Allocations for Capital Budget -- DONE
            tracingService.Trace("{0}", "Section 2");
            //Get existing space allocation for addition and partial replacement
            if (BudgetCalculationType == (int)caps_BudgetCalculationType.Addition ||
                BudgetCalculationType == (int)caps_BudgetCalculationType.PartialReplacement)
            {
                spaceAllocationExisting = CalculateSpaceAllocation(ExistingDesignCapacity, SchoolType);
            }

            if (BudgetCalculationType == (int)caps_BudgetCalculationType.Addition)
            {
                spaceAllocationNewReplacement = CalculateSpaceAllocation(ApprovedDesignCapacity, SchoolType);
            }
            else
            {
                spaceAllocationNewReplacement = CalculateSpaceAllocation(IncreaseInDesignCapacity, SchoolType);
            }


            //Get total square meterage
            decimal spaceAllocationTotalNewReplacement = (NLC.GetValueOrDefault(0) * spaceAllocationNewReplacement) + spaceAllocationNewReplacement;

            //DB: Total space allocation doesn't seem to be needed
            //decimal spaceAllocationTotal = spaceAllocationExisting + spaceAllocationTotalNewReplacement;
            #endregion

            #region 3. Construction Unit Rate -- DONE
            tracingService.Trace("{0}", "Section 3");
            //Get Base Budget Rate based on School Type
            decimal constructionBaseBudgetRate = schoolTypeRecord.caps_BaseBudgetRate.Value;
            //Get Project Size Factor based on square meterage and school type
            decimal constructionProjectSizeFactor = CalculateProjectSizeFactor(SchoolType, spaceAllocationNewReplacement);

            //Set Unit Rate = Base Budget Rate * Project Size Factor * Project Location Factor
            decimal constructionUnitRate = constructionBaseBudgetRate * constructionProjectSizeFactor * ProjectLocationFactor; 
            #endregion

            #region 4. Construction Items -- NOT DONE waiting on space allocation question
            tracingService.Trace("{0}", "Section 4");
            //Set Construction: New/Replacement = total space allocation * unit rate
            decimal constructionNewReplacement = spaceAllocationTotalNewReplacement * constructionUnitRate;                       

            //Only for Addition/Partial Replacement & Partial Seismic (4.2)
            if (BudgetCalculationType == (int)caps_BudgetCalculationType.Addition ||
                    BudgetCalculationType == (int)caps_BudgetCalculationType.PartialReplacement || 
                    BudgetCalculationType == (int)caps_BudgetCalculationType.PartialSeismic)
            {
                constructionRenovations = CalculateConstructionRenovationFactor(SchoolType, spaceAllocationNewReplacement);
            }

            //Get Site Development Allowance based on Project Type, School Type and Square meterage (4.3)
            if (BudgetCalculationType == (int)caps_BudgetCalculationType.New)
            {
                if (IncreaseInDesignCapacity.Total() >= 1500)
                {
                    constructionSiteDevelopmentAllowance = schoolTypeRecord.caps_NewSchoolGreaterThan1500.Value;
                }
                else
                {
                    constructionSiteDevelopmentAllowance = schoolTypeRecord.caps_NewSchoolLessThan1500.Value;
                }
            }
            else if (BudgetCalculationType == (int)caps_BudgetCalculationType.Replacement ||
                BudgetCalculationType == (int)caps_BudgetCalculationType.PartialReplacement)
            {
                constructionSiteDevelopmentAllowance = schoolTypeRecord.caps_Replacement.Value;
            }
            else if (BudgetCalculationType == (int)caps_BudgetCalculationType.Addition)
            {
                if (spaceAllocationNewReplacement > 2000)
                {
                    constructionSiteDevelopmentAllowance = schoolTypeRecord.caps_Additionfor2000m2.Value;
                }
                else if (spaceAllocationNewReplacement > 1000)
                {
                    constructionSiteDevelopmentAllowance = schoolTypeRecord.caps_Additionfor1000m2.Value;
                }
                else if (spaceAllocationNewReplacement > 500)
                {
                    constructionSiteDevelopmentAllowance = schoolTypeRecord.caps_Additionfor2000m2.Value;
                }
            }
            
            //Set Site Development Location Allowance = (Project Location Factor -1) * Site Development Allowance (4.4)
            decimal constructionSiteDevelopmentLocationAllowance = (ProjectLocationFactor - 1) * constructionSiteDevelopmentAllowance;

            //Set Total Construction Budget = All fields in region
            if (BudgetCalculationType == (int)caps_BudgetCalculationType.SeismicUpgrade)
            {
                constructionTotalConstructionBudget = ConstructionSeismicUpgrade.GetValueOrDefault(0) + ConstructionSPIRAdjustment.GetValueOrDefault(0) + ConstructionNonStructuralSeismicUpgrade.GetValueOrDefault(0);
            }
            else
            {
                constructionTotalConstructionBudget = constructionNewReplacement + constructionRenovations + constructionSiteDevelopmentAllowance + constructionSiteDevelopmentLocationAllowance + ConstructionSeismicUpgrade.GetValueOrDefault(0) + ConstructionSPIRAdjustment.GetValueOrDefault(0) + ConstructionNonStructuralSeismicUpgrade.GetValueOrDefault(0);
            }
            #endregion

            #region 5. Owner's Cost Items
            tracingService.Trace("{0}", "Section 5");
            //Set Design Fees = Reports and Studies Allowance % * Total Construction Budget + Base Reports and Studies Allowance (5.1a)
            decimal softCostDesignFees = (schoolTypeRecord.caps_ReportsandStudiesDesignFees.Value * constructionTotalConstructionBudget) + GetBudgetCalculationValue("Reports and Studies Allowance");

            //Set Post-Contract Contingency = Construction % * Total Construction Budget (5.2)
            decimal softCostContingencyNewReplacement = GetBudgetCalculationValue("Post Contract Contingency New/Replacement Space") * constructionTotalConstructionBudget;

            //Set Post-Contract Contingency - Renovatiions (5.3)
            if (BudgetCalculationType == (int)caps_BudgetCalculationType.Addition ||
                BudgetCalculationType == (int)caps_BudgetCalculationType.PartialReplacement ||
                BudgetCalculationType == (int)caps_BudgetCalculationType.PartialSeismic)
            {
                softCostContingencyRenovations = schoolTypeRecord.caps_SeismicReportsandStudiesDesignFees.Value * constructionRenovations;
            }
            //Set Post-Contract Contingency - Seismic (5.3a)
            if (BudgetCalculationType == (int)caps_BudgetCalculationType.PartialSeismic ||
                BudgetCalculationType == (int)caps_BudgetCalculationType.SeismicUpgrade)
            {
                softCostContingencySeismic = schoolTypeRecord.caps_SeismicReportsandStudiesDesignFees.Value * (ConstructionSeismicUpgrade.GetValueOrDefault(0) + ConstructionSPIRAdjustment.GetValueOrDefault(0) + ConstructionNonStructuralSeismicUpgrade.GetValueOrDefault(0));
            }


            //**Get Equipment Allowance Percentage
            //**Get School District Location Freight Percentage
            //Set Equipment Allowance - New Space (5.5)
            if (BudgetCalculationType == (int)caps_BudgetCalculationType.New ||
                BudgetCalculationType == (int)caps_BudgetCalculationType.Addition)
            {
                softCostEquipmentAllowance = (constructionBaseBudgetRate * spaceAllocationNewReplacement * schoolTypeRecord.caps_EquipmentAllowanceNewSpace.Value) + (constructionBaseBudgetRate * spaceAllocationNewReplacement * schoolTypeRecord.caps_EquipmentAllowanceNewSpace.Value * FreightRateAllowance);
            }
            else if (BudgetCalculationType == (int)caps_BudgetCalculationType.Replacement ||
                BudgetCalculationType == (int)caps_BudgetCalculationType.PartialReplacement ||
                BudgetCalculationType == (int)caps_BudgetCalculationType.PartialSeismic)
            {
                softCostEquipmentAllowance = (constructionBaseBudgetRate * spaceAllocationNewReplacement * schoolTypeRecord.caps_EquipmentAllowanceReplacementSpace.Value * 0.03530M) + (constructionBaseBudgetRate * spaceAllocationNewReplacement * schoolTypeRecord.caps_EquipmentAllowanceNewSpace.Value * FreightRateAllowance); ;
            }

            //Waiting on Damien
            decimal softCostWrapUpLiabilityInsurance = 0; // constructionBaseBudgetRate * 

            decimal softCostPayableTaxes = (constructionTotalConstructionBudget + softCostDesignFees + softCostContingencyNewReplacement + softCostContingencyRenovations + softCostEquipmentAllowance) * GetBudgetCalculationValue("Payable Taxes");

            decimal totalOwnersCost = constructionTotalConstructionBudget + softCostDesignFees + SPIRFees.GetValueOrDefault(0) + softCostContingencyNewReplacement + softCostContingencyRenovations + softCostContingencySeismic + MunicipalFees + softCostEquipmentAllowance + softCostWrapUpLiabilityInsurance;

            var projectManagementFee = CalculateProjectManagementFeeAllowance(totalOwnersCost);
            #endregion

            decimal supplementalCosts = Demolition + AbnormalTopography + TempAccommodation + OtherSupplemental;

            return totalOwnersCost + projectManagementFee + supplementalCosts;
        }

        private decimal CalculateSpaceAllocation(DesignCapacity designCapacityRecord,  Guid schoolType)
        {
            FilterExpression filterAnd = new FilterExpression();
            filterAnd.Conditions.Add(new ConditionExpression("caps_schooltype", ConditionOperator.Equal, schoolType));
            filterAnd.Conditions.Add(new ConditionExpression("caps_designcapacity", ConditionOperator.GreaterEqual, designCapacityRecord.Elementary + designCapacityRecord.Secondary));

            QueryExpression query = new QueryExpression("caps_budgetcalc_spaceallocation");
            query.ColumnSet.AddColumns("caps_designcapacity", "caps_spaceallocation");
            query.Criteria.AddFilter(filterAnd);
            query.AddOrder("caps_designcapacity", OrderType.Ascending);

            EntityCollection results = service.RetrieveMultiple(query);

            if (results.Entities.Count < 1) throw new Exception(string.Format("There is no space allocation record for the school type: {0} and design capacity: {1}.", schoolType, designCapacityRecord.Elementary + designCapacityRecord.Secondary));

            var designCapacity = results.Entities[0].GetAttributeValue<decimal>("caps_designcapacity");

            if (designCapacityRecord.Kindergarten > 0)
            {
                //now add kindergarten
                var kClassSize = GetBudgetCalculationValue("Design Capacity Kindergarten");
                var kRoomSize = GetBudgetCalculationValue("Design Capacity Kindergarten");

                var kClassNumber = (int)Math.Ceiling(designCapacityRecord.Kindergarten / kClassSize);

                return designCapacity + (kRoomSize * kClassNumber);
            }

            return designCapacity;
        }

        private decimal GetBudgetCalculationValue(string name)
        {

            FilterExpression filterName = new FilterExpression();
            filterName.Conditions.Add(new ConditionExpression("caps_name", ConditionOperator.Equal, name));

            QueryExpression query = new QueryExpression("caps_budgetcalc_value");
            query.ColumnSet.AddColumns("caps_value");
            query.Criteria.AddFilter(filterName);

            EntityCollection results = service.RetrieveMultiple(query);

            if (results.Entities.Count != 1) throw new Exception("Missing Budget Calculation Value: "+name);

            return results.Entities[0].GetAttributeValue<decimal>("caps_value");
        }

        private caps_BudgetCalc_SchoolType GetSchoolType(Guid schoolType)
        {
            return service.Retrieve(caps_BudgetCalc_SchoolType.EntityLogicalName, schoolType, new ColumnSet(true)) as caps_BudgetCalc_SchoolType;
        }

        private decimal CalculateProjectSizeFactor(Guid schoolType, decimal spaceAllocation)
        {
            FilterExpression filterAnd = new FilterExpression();
            filterAnd.Conditions.Add(new ConditionExpression("caps_schooltype", ConditionOperator.Equal, schoolType));
            filterAnd.Conditions.Add(new ConditionExpression("caps_upperlimitspaceallocation", ConditionOperator.GreaterEqual, spaceAllocation));

            QueryExpression query = new QueryExpression("caps_budgetcalc_projectsizefactor");
            query.ColumnSet.AddColumns("caps_factor");
            query.Criteria.AddFilter(filterAnd);
            query.AddOrder("caps_upperlimitspaceallocation", OrderType.Ascending);

            EntityCollection results = service.RetrieveMultiple(query);

            if (results.Entities.Count < 1) throw new Exception(string.Format("There is no project size factor record for the school type: {0} and design capacity: {1}.", schoolType, spaceAllocation));

            return results.Entities[0].GetAttributeValue<decimal>("caps_factor");
        }

        private decimal CalculateConstructionRenovationFactor(Guid schoolType, decimal spaceAllocation)
        {
            FilterExpression filterAnd = new FilterExpression();
            filterAnd.Conditions.Add(new ConditionExpression("caps_schooltype", ConditionOperator.Equal, schoolType));
            filterAnd.Conditions.Add(new ConditionExpression("caps_upperlimitspaceallocation", ConditionOperator.GreaterEqual, spaceAllocation));

            QueryExpression query = new QueryExpression("caps_budgetcalc_constructionrenovationfactor");
            query.ColumnSet.AddColumns("caps_factor");
            query.Criteria.AddFilter(filterAnd);
            query.AddOrder("caps_upperlimitspaceallocation", OrderType.Ascending);

            EntityCollection results = service.RetrieveMultiple(query);

            if (results.Entities.Count < 1) throw new Exception(string.Format("There is no construction renovation factor record for the school type: {0} and design capacity: {1}.", schoolType, spaceAllocation));

            return results.Entities[0].GetAttributeValue<decimal>("caps_factor");
        }

        private decimal CalculateProjectManagementFeeAllowance(decimal projectCost)
        {
            FilterExpression filterAnd = new FilterExpression();
            filterAnd.Conditions.Add(new ConditionExpression("caps_upperlimitcapitalprojectfunding", ConditionOperator.GreaterEqual, projectCost));

            QueryExpression query = new QueryExpression("caps_budgetcalc_projectmanagementfees");
            query.ColumnSet.AddColumns("caps_projectmanagementfees");
            query.Criteria.AddFilter(filterAnd);
            query.AddOrder("caps_upperlimitcapitalprojectfunding", OrderType.Ascending);

            EntityCollection results = service.RetrieveMultiple(query);

            if (results.Entities.Count < 1) throw new Exception(string.Format("There is no project management fee allowance record for the project cost: {0}.", projectCost));

            return results.Entities[0].GetAttributeValue<decimal>("caps_projectmanagementfees");
        }
    }
}
