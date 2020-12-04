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
    /// Updates the operating capacity on all active facilities and their capacity reporting records.
    /// </summary>
    public class SetOperatingCapacity : CodeActivity
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

            tracingService.Trace("{0}{1}", "Start Custom Workflow Activity: SetOperatingCapacity", DateTime.Now.ToLongTimeString());

            var recordId = context.PrimaryEntityId;

            try
            {
                //Get Global Capacity values
                var capacity = new Services.CapacityFactors();
                capacity.DesignKindergarten = GetBudgetCalculationValue(service, "Design Capacity Kindergarten");
                capacity.DesignElementary = GetBudgetCalculationValue(service, "Design Capacity Elementary");
                capacity.DesignSecondary = GetBudgetCalculationValue(service, "Design Capacity Secondary");

                capacity.OperatingKindergarten = GetBudgetCalculationValue(service, "Kindergarten Operating Capacity");
                capacity.OperatingLowElementary = GetBudgetCalculationValue(service, "Elementary Lower Operating Capacity");
                capacity.OperatingHighElementary = GetBudgetCalculationValue(service, "Elementary Upper Operating Capacity");
                capacity.OperatingSecondary = GetBudgetCalculationValue(service, "Secondary Operating Capacity");

                Services.OperatingCapacity capacityService = new Services.OperatingCapacity(service, tracingService, capacity);

                #region Update Facilities
                //get Facilities
                var fetchXML = "<fetch version=\"1.0\" output-format=\"xml-platform\" mapping=\"logical\" distinct=\"false\" >" +
                           "<entity name=\"caps_facility\">" +
                                "<attribute name=\"caps_facilityid\" /> " +
                                "<attribute name=\"caps_name\" /> " +
                                "<attribute name=\"caps_lowestgrade\" /> " +
                                "<attribute name=\"caps_highestgrade\" /> " +
                                "<attribute name=\"caps_adjusteddesigncapacitysecondary\" /> " +
                                "<attribute name=\"caps_adjusteddesigncapacitykindergarten\" /> " +
                                "<attribute name=\"caps_adjusteddesigncapacityelementary\" /> " +
                                "<order attribute=\"caps_name\" descending=\"false\" /> " +
                                    "<filter type=\"and\" > " +
                                           "<condition attribute=\"statecode\" operator=\"eq\" value=\"0\" /> " +
                                    "</filter> " +
                                "<link-entity name=\"caps_facilitytype\" from=\"caps_facilitytypeid\" to=\"caps_currentfacilitytype\" link-type=\"inner\" alias=\"ac\" > " +
                                        "<filter type=\"and\" > " +
                                            "<condition attribute=\"caps_schooltype\" operator=\"not-null\" /> " +
                                        "</filter> " +
                                "</link-entity> " +
                            "</entity> " +
                            "</fetch>";


                EntityCollection results = service.RetrieveMultiple(new FetchExpression(fetchXML));

                foreach (caps_Facility facilityRecord in results.Entities)
                {
                    tracingService.Trace("Facility:{0}", facilityRecord.caps_Name);

                    var kDesign = facilityRecord.caps_AdjustedDesignCapacityKindergarten.GetValueOrDefault(0);
                    var eDesign = facilityRecord.caps_AdjustedDesignCapacityElementary.GetValueOrDefault(0);
                    var sDesign = facilityRecord.caps_AdjustedDesignCapacitySecondary.GetValueOrDefault(0);
                    var lowestGrade = facilityRecord.caps_LowestGrade.Value;
                    var highestGrade = facilityRecord.caps_HighestGrade.Value;

                    var result = capacityService.Calculate(kDesign, eDesign, sDesign, lowestGrade, highestGrade);

                    //Update Facility
                    var recordToUpdate = new caps_Facility();
                    recordToUpdate.Id = facilityRecord.Id;
                    recordToUpdate.caps_OperatingCapacityKindergarten = result.KindergartenCapacity;
                    recordToUpdate.caps_OperatingCapacityElementary = result.ElementaryCapacity;
                    recordToUpdate.caps_OperatingCapacitySecondary = result.SecondaryCapacity;
                    service.Update(recordToUpdate);
                } 
                #endregion

                //Update Capacity Reporting
                var capacityFetchXML = "<fetch version=\"1.0\" output-format=\"xml-platform\" mapping=\"logical\" distinct=\"false\" >"+
                                       "<entity name=\"caps_capacityreporting\" > "+
                                        "<attribute name=\"caps_capacityreportingid\" /> "+
                                        "<attribute name=\"caps_secondary_designcapacity\" /> "+
                                        "<attribute name=\"caps_kindergarten_designcapacity\" /> "+
                                        "<attribute name=\"caps_elementary_designcapacity\" /> "+
                                         "<order attribute=\"caps_secondary_designutilization\" descending=\"false\" /> "+
                                            "<link-entity name=\"caps_facility\" from=\"caps_facilityid\" to=\"caps_facility\" visible=\"false\" link-type=\"outer\" alias=\"facility\" > "+
                                                "<attribute name=\"caps_lowestgrade\" /> "+
                                                "<attribute name=\"caps_highestgrade\" /> "+
                                            "</link-entity> "+
                                                "<link-entity name=\"edu_year\" from=\"edu_yearid\" to=\"caps_schoolyear\" link-type=\"inner\" alias=\"ab\" >"+
                                                "<filter type=\"and\"> "+
                                                    "<condition attribute=\"statuscode\" operator=\"in\">"+
                                                       "<value>1</value> "+
                                                       "<value>757500000</value> "+
                                                     "</condition>"+
                                                   "</filter>"+
                                                 "</link-entity>"+
                                            "</entity> " +
                                     "</fetch> ";

                EntityCollection capacityResults = service.RetrieveMultiple(new FetchExpression(capacityFetchXML));

                foreach (caps_CapacityReporting capacityRecord in capacityResults.Entities)
                {
                    var kDesign = capacityRecord.caps_Kindergarten_designcapacity.GetValueOrDefault(0);
                    var eDesign = capacityRecord.caps_Elementary_designcapacity.GetValueOrDefault(0);
                    var sDesign = capacityRecord.caps_Secondary_designcapacity.GetValueOrDefault(0);
                    var lowestGrade = ((OptionSetValue)((AliasedValue)capacityRecord["facility.caps_lowestgrade"]).Value).Value; 
                    var highestGrade = ((OptionSetValue)((AliasedValue)capacityRecord["facility.caps_highestgrade"]).Value).Value;

                    var result = capacityService.Calculate(kDesign, eDesign, sDesign, lowestGrade, highestGrade);

                    //Update Capacity Reporting
                    var recordToUpdate = new caps_CapacityReporting();
                    recordToUpdate.Id = capacityRecord.Id;
                    recordToUpdate.caps_Kindergarten_operatingcapacity = result.KindergartenCapacity;
                    recordToUpdate.caps_Elementary_operatingcapacity = result.ElementaryCapacity;
                    recordToUpdate.caps_Secondary_operatingcapacity = result.SecondaryCapacity;
                    service.Update(recordToUpdate);
                }

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
