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
    /// Updates the operating capacity on all project request records.
    /// </summary>
    public class SetOperatingCapacityForProjectRequests : CodeActivity
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
                var capacity = new Services.CapacityFactors(service);

                Services.OperatingCapacity capacityService = new Services.OperatingCapacity(service, tracingService, capacity);

                #region Update Draft Project Requests
                //Update Draft Project Requests
                var projectRequestFetchXML = "<fetch version=\"1.0\" output-format=\"xml-platform\" mapping=\"logical\" distinct=\"false\">" +
                                            "<entity name=\"caps_project\" > " +
                                               "<attribute name=\"caps_facility\" /> " +
                                                "<attribute name=\"caps_changeindesigncapacitykindergarten\" /> " +
                                                 "<attribute name=\"caps_changeindesigncapacityelementary\" /> " +
                                                  "<attribute name=\"caps_changeindesigncapacitysecondary\" /> " +
                                                  "<attribute name=\"caps_futurelowestgrade\" /> " +
                                                  "<attribute name=\"caps_futurehighestgrade\" /> " +
                                                    "<order attribute=\"caps_projectcode\" descending=\"false\" /> " +
                                                       "<filter type=\"and\" > " +
                                                          "<condition attribute=\"statuscode\" operator=\"eq\" value=\"1\" /> " +
                                                              "<filter type=\"or\" > " +
                                                                 "<condition attribute=\"caps_changeinoperatingcapacitykindergarten\" operator=\"not-null\" /> " +
                                                                    "<condition attribute=\"caps_changeinoperatingcapacityelementary\" operator=\"not-null\" /> " +
                                                                       "<condition attribute=\"caps_changeinoperatingcapacitysecondary\" operator=\"not-null\" /> " +
                                                                        "</filter> " +
                                                                      "</filter> " +
                                                                    "</entity> " +
                                                                  "</fetch> ";

                EntityCollection projectResults = service.RetrieveMultiple(new FetchExpression(projectRequestFetchXML));

                foreach (caps_Project projectRecord in projectResults.Entities)
                {
                    if (projectRecord.caps_FutureLowestGrade != null && projectRecord.caps_FutureHighestGrade != null)
                    {
                        var startingDesign_K = 0;
                        var startingDesign_E = 0;
                        var startingDesign_S = 0;

                        //if facility exists, then retrieve it
                        if (projectRecord.caps_Facility != null && projectRecord.caps_Facility.Id != Guid.Empty)
                        {
                            var facility = service.Retrieve(caps_Facility.EntityLogicalName, projectRecord.caps_Facility.Id, new ColumnSet("caps_adjusteddesigncapacitykindergarten", "caps_adjusteddesigncapacityelementary", "caps_adjusteddesigncapacitysecondary")) as caps_Facility;

                            if (facility != null)
                            {
                                startingDesign_K = facility.caps_AdjustedDesignCapacityKindergarten.GetValueOrDefault(0);
                                startingDesign_E = facility.caps_AdjustedDesignCapacityElementary.GetValueOrDefault(0);
                                startingDesign_S = facility.caps_AdjustedDesignCapacitySecondary.GetValueOrDefault(0);
                            }
                        }

                        var changeInDesign_K = startingDesign_K + projectRecord.caps_ChangeinDesignCapacityKindergarten.GetValueOrDefault(0);
                        var changeInDesign_E = startingDesign_E + projectRecord.caps_ChangeinDesignCapacityElementary.GetValueOrDefault(0);
                        var changeInDesign_S = startingDesign_S + projectRecord.caps_ChangeinDesignCapacitySecondary.GetValueOrDefault(0);

                        var result = capacityService.Calculate(changeInDesign_K, changeInDesign_E, changeInDesign_S, projectRecord.caps_FutureLowestGrade.Value, projectRecord.caps_FutureHighestGrade.Value);

                        var recordToUpdate = new caps_Project();
                        recordToUpdate.Id = recordId;
                        recordToUpdate.caps_ChangeinOperatingCapacityKindergarten = Convert.ToInt32(result.KindergartenCapacity);
                        recordToUpdate.caps_ChangeinOperatingCapacityElementary = Convert.ToInt32(result.ElementaryCapacity);
                        recordToUpdate.caps_ChangeinOperatingCapacitySecondary = Convert.ToInt32(result.SecondaryCapacity);
                        service.Update(recordToUpdate);
                    }
                    else
                    {
                        //blank out operating capacity
                        var recordToUpdate = new caps_Project();
                        recordToUpdate.Id = recordId;
                        recordToUpdate.caps_ChangeinOperatingCapacityKindergarten = null;
                        recordToUpdate.caps_ChangeinOperatingCapacityElementary = null;
                        recordToUpdate.caps_ChangeinOperatingCapacitySecondary = null;
                        service.Update(recordToUpdate);
                    }
                }
                #endregion


                this.error.Set(executionContext, false);
            }
            catch (Exception ex)
            {
                tracingService.Trace("Error Details: {0}", ex.Message);
                //might want to also include error message
                this.error.Set(executionContext, true);
                this.errorMessage.Set(executionContext, ex.Message);
            }

            tracingService.Trace("{0}{1}", "End Custom Workflow Activity: SetOperatingCapacity", DateTime.Now.ToLongTimeString());
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
