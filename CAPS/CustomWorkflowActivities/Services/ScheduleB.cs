using System;
using CAPS.DataContext;
using Microsoft.Xrm.Sdk;
using Microsoft.Xrm.Sdk.Query;

namespace CustomWorkflowActivities.Services
{
    internal class CalculationResult
    {
        internal decimal SpaceAllocationNewReplacement {get; set;}
        //internal decimal SpaceAllocationNLC {get; set;}
        internal decimal BaseBudgetRate { get; set; }
        internal decimal ProjectSizeFactor { get; set; }
        internal decimal ProjectLocationFactor { get; set; }
        internal decimal UnitRate { get; set; }
        internal decimal ConstructionNewReplacement { get; set; }
        internal decimal ConstructionRenovation { get; set; }
        internal decimal SiteDevelopmentAllowance {get; set;}
        internal decimal SiteDevelopmentLocationAllowance {get; set;}
        internal decimal DesignFees {get; set;}
        internal decimal SPIRFees {get; set;}
        internal decimal PostContractNewReplacement {get; set;}
        internal decimal PostContractRenovation {get; set;}
        internal decimal PostContractSeismic { get; set; }
        //internal decimal MunicipalFees {get; set;}
        internal decimal EquipmentNew { get; set; }
        internal decimal EquipmentReplacement { get; set; }
        internal decimal ProjectManagement {get; set;}
        internal decimal LiabilityInsurance {get; set;}
        internal decimal PayableTaxes {get; set;}
        internal decimal RiskReserve { get; set; }
        internal decimal NLCBudgetAmount { get; set;}

    internal decimal Total {get; set;}


    //internal decimal  {get; set;}
}
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
        internal DesignCapacity ExistingAndDecreaseDesignCapacity { get; set; }
        internal DesignCapacity ApprovedDesignCapacity { get; set; }
        internal decimal? ExtraSpaceAllocation { get; set; }
        internal decimal ProjectLocationFactor { get; set; }
        //internal decimal? NLC { get; set; }
        internal bool IncludeNLC { get; set; }
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

        /// <summary>
        /// Default constructor for Schedule B
        /// </summary>
        /// <param name="s">IOrganizationService</param>
        /// <param name="t">ITracingService</param>
        public ScheduleB(IOrganizationService s, ITracingService t)
        {
            this.service = s;
            this.tracingService = t;
        }

        /// <summary>
        /// Function to calculate schedule b total
        /// </summary>
        /// <returns>Schedule B Total</returns>
        internal CalculationResult Calculate()
        {
            CalculationResult result = new CalculationResult();
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
            decimal softCostWrapUpLiabilityInsurance = 0;
            decimal softCostDesignFees = 0;
            decimal softCostContingencyNewReplacement = 0;
            decimal NLCAmount = 0;

            //Default some numbers to 0, cased on Budget Calculation type
            //if partial seismic, set NLC to 0
            //if (BudgetCalculationType == (int)caps_BudgetCalculationType.PartialSeismic)
            //{
            //    NLC = 0;
            //}

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
            spaceAllocationNewReplacement = CalculateSpaceAllocation(ApprovedDesignCapacity, SchoolType, schoolTypeRecord.caps_Name);

            //Get existing space allocation for addition and partial replacement less any reductions
            if (BudgetCalculationType == (int)caps_BudgetCalculationType.Addition ||
                BudgetCalculationType == (int)caps_BudgetCalculationType.PartialReplacement ||
                BudgetCalculationType == (int)caps_BudgetCalculationType.PartialSeismic)
            {
                spaceAllocationExisting = CalculateSpaceAllocation(ExistingAndDecreaseDesignCapacity, SchoolType, schoolTypeRecord.caps_Name);
                spaceAllocationNewReplacement = spaceAllocationNewReplacement - spaceAllocationExisting;
            }

            tracingService.Trace("Space Allocation - New/Replacement: {0}", spaceAllocationNewReplacement);


            result.SpaceAllocationNewReplacement = spaceAllocationNewReplacement;

            //add extra space allocation
            spaceAllocationNewReplacement += ExtraSpaceAllocation.GetValueOrDefault(0);
            //Get total square meterage
            decimal spaceAllocationTotalNewReplacement = spaceAllocationNewReplacement;

            tracingService.Trace("Space Allocation - Total New/Replacement: {0}", spaceAllocationTotalNewReplacement);
            //DB: Total space allocation doesn't seem to be needed
            //decimal spaceAllocationTotal = spaceAllocationExisting + spaceAllocationTotalNewReplacement;
            #endregion

            #region 3. Construction Unit Rate -- DONE
            tracingService.Trace("{0}", "Section 3");
            //Get Base Budget Rate based on School Type
            decimal constructionBaseBudgetRate = schoolTypeRecord.caps_BaseBudgetRate.Value;

            result.BaseBudgetRate = constructionBaseBudgetRate;
            tracingService.Trace("Base Budget Rate: {0}", constructionBaseBudgetRate);

            //Get Project Size Factor based on square meterage and school type
            decimal constructionProjectSizeFactor = CalculateProjectSizeFactor(SchoolType, schoolTypeRecord.caps_Name, spaceAllocationNewReplacement);

            tracingService.Trace("Project Size Factor: {0}", constructionProjectSizeFactor);
            tracingService.Trace("Project Location Factor: {0}", ProjectLocationFactor);
            result.ProjectSizeFactor = constructionProjectSizeFactor;
            result.ProjectLocationFactor = ProjectLocationFactor;

            //Set Unit Rate = Base Budget Rate * Project Size Factor * Project Location Factor
            decimal constructionUnitRate = constructionBaseBudgetRate * constructionProjectSizeFactor * ProjectLocationFactor;

            tracingService.Trace("Unit Rate: {0}", constructionUnitRate);
            result.UnitRate = constructionUnitRate;
            #endregion

            #region 4. Construction Items -- DONE
            tracingService.Trace("{0}", "Section 4");
            //Set Construction: New/Replacement = total space allocation * unit rate
            decimal constructionNewReplacement = spaceAllocationTotalNewReplacement * constructionUnitRate;

            tracingService.Trace("Construction - New/Replacement: {0}", constructionNewReplacement);
            result.ConstructionNewReplacement = constructionNewReplacement;

            //Only for Addition/Partial Replacement & Partial Seismic (4.2)
            if (BudgetCalculationType == (int)caps_BudgetCalculationType.Addition ||
                    BudgetCalculationType == (int)caps_BudgetCalculationType.PartialReplacement || 
                    BudgetCalculationType == (int)caps_BudgetCalculationType.PartialSeismic)
            {
                constructionRenovations = CalculateConstructionRenovationFactor(SchoolType, schoolTypeRecord.caps_Name, spaceAllocationNewReplacement) * constructionNewReplacement;
            }

            tracingService.Trace("Construction - Renovations: {0}", constructionRenovations);
            result.ConstructionRenovation = constructionRenovations;

            //Get Site Development Allowance based on Project Type, School Type and Square meterage (4.3)
            if (BudgetCalculationType == (int)caps_BudgetCalculationType.New)
            {
                if (ApprovedDesignCapacity.Total() >= 1500)
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
            else if (BudgetCalculationType == (int)caps_BudgetCalculationType.Addition ||
                BudgetCalculationType == (int)caps_BudgetCalculationType.PartialSeismic)
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
                    constructionSiteDevelopmentAllowance = schoolTypeRecord.caps_Additionfor500m2.Value;
                }
            }

            tracingService.Trace("Site Development Allowance: {0}", constructionSiteDevelopmentAllowance);
            result.SiteDevelopmentAllowance = constructionSiteDevelopmentAllowance;

            //Set Site Development Location Allowance = (Project Location Factor -1) * Site Development Allowance (4.4)
            decimal constructionSiteDevelopmentLocationAllowance = (ProjectLocationFactor - 1) * constructionSiteDevelopmentAllowance;

            tracingService.Trace("Site Development Location Allowance: {0}", constructionSiteDevelopmentLocationAllowance);
            result.SiteDevelopmentLocationAllowance = constructionSiteDevelopmentLocationAllowance;

            //Set Total Construction Budget = All fields in region
            if (BudgetCalculationType == (int)caps_BudgetCalculationType.SeismicUpgrade)
            {
                constructionTotalConstructionBudget = ConstructionSeismicUpgrade.GetValueOrDefault(0) + ConstructionSPIRAdjustment.GetValueOrDefault(0) + ConstructionNonStructuralSeismicUpgrade.GetValueOrDefault(0);
            }
            else
            {
                constructionTotalConstructionBudget = constructionNewReplacement + constructionRenovations + constructionSiteDevelopmentAllowance + constructionSiteDevelopmentLocationAllowance + ConstructionSeismicUpgrade.GetValueOrDefault(0) + ConstructionSPIRAdjustment.GetValueOrDefault(0) + ConstructionNonStructuralSeismicUpgrade.GetValueOrDefault(0);
            }

            tracingService.Trace("Total Construction Budget: {0}", constructionTotalConstructionBudget);
            #endregion

            #region 5. Owner's Cost Items
            tracingService.Trace("{0}", "Section 5");
            //Set Design Fees = Reports and Studies Allowance % * Total Construction Budget + Base Reports and Studies Allowance (5.1a)
            if (BudgetCalculationType == (int)caps_BudgetCalculationType.SeismicUpgrade)
            {
                softCostDesignFees = (schoolTypeRecord.caps_SeismicReportsandStudiesDesignFees.GetValueOrDefault(0) * constructionTotalConstructionBudget) + GetBudgetCalculationValue("Reports and Studies Allowance");
            }
            else
            {
                softCostDesignFees = (schoolTypeRecord.caps_ReportsandStudiesDesignFees.GetValueOrDefault(0) * constructionTotalConstructionBudget) + GetBudgetCalculationValue("Reports and Studies Allowance");
            }

            tracingService.Trace("Design Fees: {0}", softCostDesignFees);
            result.DesignFees = softCostDesignFees;

            //Set Post-Contract Contingency = Construction % * Total Construction Budget (5.2)

            if (BudgetCalculationType == (int)caps_BudgetCalculationType.New ||
                    BudgetCalculationType == (int)caps_BudgetCalculationType.Replacement)
            {
                softCostContingencyNewReplacement = GetBudgetCalculationValue("Post Contract Contingency New/Replacement Space") * constructionTotalConstructionBudget;
            }
            else if (BudgetCalculationType == (int)caps_BudgetCalculationType.Addition ||
                BudgetCalculationType == (int)caps_BudgetCalculationType.PartialReplacement ||
                BudgetCalculationType == (int)caps_BudgetCalculationType.PartialSeismic
                )
            {
                softCostContingencyNewReplacement = GetBudgetCalculationValue("Post Contract Contingency New/Replacement Space") * (constructionNewReplacement + constructionSiteDevelopmentAllowance + constructionSiteDevelopmentLocationAllowance);
            }

            tracingService.Trace("Post-Contract Contingency - New/Replacement: {0}", softCostContingencyNewReplacement);
            result.PostContractNewReplacement = softCostContingencyNewReplacement;

            //Set Post-Contract Contingency - Renovations (5.3)
            if (BudgetCalculationType == (int)caps_BudgetCalculationType.Addition ||
                BudgetCalculationType == (int)caps_BudgetCalculationType.PartialReplacement ||
                BudgetCalculationType == (int)caps_BudgetCalculationType.PartialSeismic)
            {
                softCostContingencyRenovations = GetBudgetCalculationValue("Post Contract Contingency Renovations") * constructionRenovations;
            }

            //Set Post-Contract Contingency - Seismic (5.3a)
            if (BudgetCalculationType == (int)caps_BudgetCalculationType.PartialSeismic ||
                BudgetCalculationType == (int)caps_BudgetCalculationType.SeismicUpgrade)
            {
                
                softCostContingencySeismic = GetBudgetCalculationValue("Post Contract Contingency Seismic") * (ConstructionSeismicUpgrade.GetValueOrDefault(0) + ConstructionSPIRAdjustment.GetValueOrDefault(0) + ConstructionNonStructuralSeismicUpgrade.GetValueOrDefault(0));
            }

            tracingService.Trace("Post-Contract Contingency - Renovations: {0}", softCostContingencyRenovations);
            tracingService.Trace("Post-Contract Contingency - Seismic: {0}", softCostContingencySeismic);
            result.PostContractRenovation = softCostContingencyRenovations;
            result.PostContractSeismic = softCostContingencySeismic;


            //**Get Equipment Allowance Percentage
            //**Get School District Location Freight Percentage
            //Set Equipment Allowance - New Space (5.5)
            if (BudgetCalculationType == (int)caps_BudgetCalculationType.New ||
                BudgetCalculationType == (int)caps_BudgetCalculationType.Addition)
            {
                softCostEquipmentAllowance = (constructionBaseBudgetRate * spaceAllocationNewReplacement * schoolTypeRecord.caps_EquipmentAllowanceNewSpace.Value) + (constructionBaseBudgetRate * spaceAllocationNewReplacement * schoolTypeRecord.caps_EquipmentAllowanceNewSpace.Value * FreightRateAllowance);
                result.EquipmentNew = softCostEquipmentAllowance;
            }
            else if (BudgetCalculationType == (int)caps_BudgetCalculationType.Replacement ||
                BudgetCalculationType == (int)caps_BudgetCalculationType.PartialReplacement ||
                BudgetCalculationType == (int)caps_BudgetCalculationType.PartialSeismic)
            {
                softCostEquipmentAllowance = (constructionBaseBudgetRate * spaceAllocationNewReplacement * schoolTypeRecord.caps_EquipmentAllowanceReplacementSpace.Value) + (constructionBaseBudgetRate * spaceAllocationNewReplacement * schoolTypeRecord.caps_EquipmentAllowanceNewSpace.Value * FreightRateAllowance);
                result.EquipmentReplacement = softCostEquipmentAllowance;
            }

            tracingService.Trace("Equipment Allowance: {0}", softCostEquipmentAllowance);

            //Wrap-Up Liability Insurance (5.8)
            if (BudgetCalculationType == (int)caps_BudgetCalculationType.PartialSeismic ||
                BudgetCalculationType == (int)caps_BudgetCalculationType.SeismicUpgrade)
            {
                softCostWrapUpLiabilityInsurance = constructionTotalConstructionBudget / 1000 * GetBudgetCalculationValue("Wrap Up Liability Insurance Renovations or Seismic");
            }
            else
            {
                softCostWrapUpLiabilityInsurance = constructionTotalConstructionBudget / 1000 * GetBudgetCalculationValue("Wrap Up Liability Insurance New or Replacement");
            }

            tracingService.Trace("Liability Insurance: {0}", softCostWrapUpLiabilityInsurance);
            result.LiabilityInsurance = softCostWrapUpLiabilityInsurance;

            decimal softCostPayableTaxes = (constructionTotalConstructionBudget + softCostDesignFees + softCostContingencyNewReplacement + softCostContingencyRenovations + softCostContingencySeismic+ softCostEquipmentAllowance) * GetBudgetCalculationValue("Payable Taxes");

            tracingService.Trace("Payable Taxes: {0}", softCostPayableTaxes);
            result.PayableTaxes = softCostPayableTaxes;

            decimal totalOwnersCost = softCostDesignFees + SPIRFees.GetValueOrDefault(0) + softCostContingencyNewReplacement + softCostContingencyRenovations + softCostContingencySeismic + MunicipalFees + softCostEquipmentAllowance + softCostWrapUpLiabilityInsurance + softCostPayableTaxes;

            tracingService.Trace("Total Owners Costs: {0}", totalOwnersCost);
            

            var projectManagementFee = CalculateProjectManagementFeeAllowance(totalOwnersCost + constructionTotalConstructionBudget);
            result.ProjectManagement = projectManagementFee;
            tracingService.Trace("Project Management Fee: {0}", projectManagementFee);
            #endregion

            decimal supplementalCosts = Demolition + AbnormalTopography + TempAccommodation + OtherSupplemental;

            //UPDATE: Add NLC to supplemental
            if ((BudgetCalculationType == (int)caps_BudgetCalculationType.New || BudgetCalculationType == (int)caps_BudgetCalculationType.Replacement)
                && IncludeNLC)
            {
                //get ProjectLocationFactor and NLC Amount
                NLCAmount = CalculateNLCAmount(ExistingAndDecreaseDesignCapacity, SchoolType, schoolTypeRecord.caps_Name);
                NLCAmount = NLCAmount * ProjectLocationFactor;
            }
            result.NLCBudgetAmount = NLCAmount;

            tracingService.Trace("Supplemental Costs: {0}", supplementalCosts);

            var subTotal = totalOwnersCost + constructionTotalConstructionBudget + projectManagementFee + supplementalCosts + NLCAmount;
            
            var riskReserve = subTotal * GetBudgetCalculationValue("Risk Reserve and Escalation");

            result.RiskReserve = riskReserve;

            result.Total = subTotal + riskReserve;

            return result;
        }

        /// <summary>
        /// Get Space Allocation by design capacity and school type
        /// </summary>
        /// <param name="designCapacityRecord">Design capacity of K, E & S</param>
        /// <param name="schoolType">Type of School</param>
        /// <returns>Square metre space allocation</returns>
        private decimal CalculateSpaceAllocation(DesignCapacity designCapacityRecord,  Guid schoolType, string schoolTypeName)
        {
            FilterExpression filterAnd = new FilterExpression();
            filterAnd.Conditions.Add(new ConditionExpression("caps_schooltype", ConditionOperator.Equal, schoolType));
            filterAnd.Conditions.Add(new ConditionExpression("caps_designcapacity", ConditionOperator.GreaterEqual, designCapacityRecord.Elementary + designCapacityRecord.Secondary));

            QueryExpression query = new QueryExpression("caps_budgetcalc_spaceallocation");
            query.ColumnSet.AddColumns("caps_designcapacity", "caps_spaceallocation");
            query.Criteria.AddFilter(filterAnd);
            query.AddOrder("caps_designcapacity", OrderType.Ascending);

            EntityCollection results = service.RetrieveMultiple(query);

            tracingService.Trace("Count of Results: {0}", results.Entities.Count);

            if (results.Entities.Count < 1) throw new Exception(string.Format("There is no space allocation record for the school type: {0} and design capacity: {1}.", schoolTypeName, designCapacityRecord.Elementary + designCapacityRecord.Secondary));

            var designCapacity = results.Entities[0].GetAttributeValue<decimal>("caps_spaceallocation");

            tracingService.Trace("Design Capacity: {0}", designCapacity);

            if (designCapacityRecord.Kindergarten > 0)
            {
                //now add kindergarten
                var kClassSize = GetBudgetCalculationValue("Design Capacity Kindergarten");
                var kRoomSize = GetBudgetCalculationValue("Kindergarten Space Allocation");

                tracingService.Trace("K Class Size: {0}", kClassSize);
                tracingService.Trace("K Room Size: {0}", kRoomSize);

                var kClassNumber = (int)Math.Ceiling(designCapacityRecord.Kindergarten / kClassSize);

                tracingService.Trace("K Class Number: {0}", kClassNumber);

                return designCapacity + (kRoomSize * kClassNumber);
            }

            return designCapacity;
        }

        /// <summary>
        /// Get NLC amount by design capacity and school type
        /// </summary>
        /// <param name="designCapacityRecord">Design capacity of K, E & S</param>
        /// <param name="schoolType">Type of School</param>
        /// <returns>NLC dollar amount</returns>
        private decimal CalculateNLCAmount(DesignCapacity designCapacityRecord, Guid schoolType, string schoolTypeName)
        {
            FilterExpression filterAnd = new FilterExpression();
            filterAnd.Conditions.Add(new ConditionExpression("caps_schooltype", ConditionOperator.Equal, schoolType));
            filterAnd.Conditions.Add(new ConditionExpression("caps_designcapacity", ConditionOperator.GreaterEqual, designCapacityRecord.Elementary + designCapacityRecord.Secondary));

            QueryExpression query = new QueryExpression("caps_nlcfactor");
            query.ColumnSet.AddColumns("caps_designcapacity", "caps_budget");
            query.Criteria.AddFilter(filterAnd);
            query.AddOrder("caps_designcapacity", OrderType.Ascending);

            EntityCollection results = service.RetrieveMultiple(query);

            tracingService.Trace("Count of Results: {0}", results.Entities.Count);

            if (results.Entities.Count < 1) throw new Exception(string.Format("There is no NLC Factor record for the school type: {0} and design capacity: {1}.", schoolTypeName, designCapacityRecord.Elementary + designCapacityRecord.Secondary));

            var nlcBudget = results.Entities[0].GetAttributeValue<int>("caps_budget");

            tracingService.Trace("NLC Amount: {0}", nlcBudget);

            return nlcBudget;
        }

        /// <summary>
        /// Lookup Budget Calculation Value by Name
        /// </summary>
        /// <param name="name">Name of record</param>
        /// <returns>Value of record</returns>
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

        /// <summary>
        /// Get school type record and all fields by School Type Guid
        /// </summary>
        /// <param name="schoolType">School Type Guid</param>
        /// <returns>Full School Type record</returns>
        private caps_BudgetCalc_SchoolType GetSchoolType(Guid schoolType)
        {
            return service.Retrieve(caps_BudgetCalc_SchoolType.EntityLogicalName, schoolType, new ColumnSet(true)) as caps_BudgetCalc_SchoolType;
        }

        /// <summary>
        /// Lookup Project Size Factory by school type and space allocation (m²)
        /// </summary>
        /// <param name="schoolType">school type guid</param>
        /// <param name="schoolTypeName">school type name</param>
        /// <param name="spaceAllocation">Space allocation in metres squared</param>
        /// <returns></returns>
        private decimal CalculateProjectSizeFactor(Guid schoolType, string schoolTypeName, decimal spaceAllocation)
        {
            FilterExpression filterAnd = new FilterExpression();
            filterAnd.Conditions.Add(new ConditionExpression("caps_schooltype", ConditionOperator.Equal, schoolType));
            filterAnd.Conditions.Add(new ConditionExpression("caps_upperlimitspaceallocation", ConditionOperator.GreaterEqual, spaceAllocation));

            QueryExpression query = new QueryExpression("caps_budgetcalc_projectsizefactor");
            query.ColumnSet.AddColumns("caps_factor");
            query.Criteria.AddFilter(filterAnd);
            query.AddOrder("caps_upperlimitspaceallocation", OrderType.Ascending);

            EntityCollection results = service.RetrieveMultiple(query);

            if (results.Entities.Count < 1) throw new Exception(string.Format("There is no project size factor record for the school type: {0} and design capacity: {1}.", schoolTypeName, spaceAllocation));

            return results.Entities[0].GetAttributeValue<decimal>("caps_factor");
        }

        /// <summary>
        /// Lookup Construction Renovation Factory by school type and space allocation (m²)
        /// </summary>
        /// <param name="schoolType">school type guid</param>
        /// <param name="schoolTypeName">school type name</param>
        /// <param name="spaceAllocation">Space allocation in metres squared</param>
        /// <returns></returns>
        private decimal CalculateConstructionRenovationFactor(Guid schoolType, string schoolTypeName, decimal spaceAllocation)
        {
            FilterExpression filterAnd = new FilterExpression();
            filterAnd.Conditions.Add(new ConditionExpression("caps_schooltype", ConditionOperator.Equal, schoolType));
            filterAnd.Conditions.Add(new ConditionExpression("caps_upperlimitspaceallocation", ConditionOperator.GreaterEqual, spaceAllocation));

            QueryExpression query = new QueryExpression("caps_budgetcalc_constructionrenovationfactor");
            query.ColumnSet.AddColumns("caps_factor");
            query.Criteria.AddFilter(filterAnd);
            query.AddOrder("caps_upperlimitspaceallocation", OrderType.Ascending);

            EntityCollection results = service.RetrieveMultiple(query);

            if (results.Entities.Count < 1) throw new Exception(string.Format("There is no construction renovation factor record for the school type: {0} and design capacity: {1}.", schoolTypeName, spaceAllocation));

            return results.Entities[0].GetAttributeValue<decimal>("caps_factor");
        }

        /// <summary>
        /// Calculate Project Management Fee Allowance by project cost
        /// </summary>
        /// <param name="projectCost">sub total of project cost</param>
        /// <returns>project management fee</returns>
        private decimal CalculateProjectManagementFeeAllowance(decimal projectCost)
        {
            FilterExpression filterAnd = new FilterExpression();
            filterAnd.Conditions.Add(new ConditionExpression("caps_upperlimitcapitalprojectfunding", ConditionOperator.GreaterEqual, projectCost));

            QueryExpression query = new QueryExpression("caps_budgetcalc_projectmanagementfees");
            query.ColumnSet.AddColumns("caps_projectmanagementfees");
            query.Criteria.AddFilter(filterAnd);
            query.AddOrder("caps_upperlimitcapitalprojectfunding", OrderType.Ascending);

            EntityCollection results = service.RetrieveMultiple(query);

            if (results.Entities.Count < 1) throw new Exception(string.Format("There is no project management fee allowance record for the project cost: {0}.", projectCost.ToString("C")));

            return results.Entities[0].GetAttributeValue<decimal>("caps_projectmanagementfees");
        }
    }
}
