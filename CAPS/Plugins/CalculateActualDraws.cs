using CAPS.DataContext;
using Microsoft.Xrm.Sdk;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace Plugins
{
    /// <summary>
    /// This plugin runs on change of actual draws.  It finds the previous and current related project, and adjusts the actual draw amount for each project cash flow.
    /// Register on async Post of events (Create, Update, Delete)
    /// </summary>
    public class CalculateActualDraws : IPlugin
    {
        public void Execute(IServiceProvider serviceProvider)
        {
            IPluginExecutionContext context = (IPluginExecutionContext)serviceProvider.GetService(typeof(IPluginExecutionContext));
            IOrganizationServiceFactory serviceFactory = (IOrganizationServiceFactory)serviceProvider.GetService(typeof(IOrganizationServiceFactory));
            IOrganizationService service = serviceFactory.CreateOrganizationService(context.UserId);
            ITracingService tracingService = (ITracingService)serviceProvider.GetService(typeof(ITracingService));

            tracingService.Trace("{0}", "CalculateActualDraws Plug-in");

            // The InputParameters collection contains all the data passed in the message request.
            if (context.InputParameters.Contains("Target") 
                && (context.InputParameters["Target"] is Entity || context.InputParameters["Target"] is EntityReference))
            {
                tracingService.Trace("{0}", "Line 25");

                var projectsToUpdate = new EntityReferenceCollection();

                //Keep CRUD in mind, need to handle all scenarios
                //Get previous project (if exists) and current project
                if (context.MessageName == "Create" || context.MessageName == "Update")
                {
                    Entity entity = (Entity)context.InputParameters["Target"];

                    if (entity.LogicalName != caps_ActualDraw.EntityLogicalName)
                        return;

                    if (context.MessageName == "Create" && entity.Contains("caps_project"))
                    {
                        projectsToUpdate.Add(entity.GetAttributeValue<EntityReference>("caps_project"));
                    }

                    if (context.MessageName == "Update")
                    {
                        Entity preImage = context.PreEntityImages["preImage"];

                        if (preImage.Contains("caps_project"))
                        {
                            projectsToUpdate.Add(preImage.GetAttributeValue<EntityReference>("caps_project"));
                        }
                        if (entity.Contains("caps_project"))
                        {
                            projectsToUpdate.Add(entity.GetAttributeValue<EntityReference>("caps_project"));
                        }
                    }
                }
                tracingService.Trace("Message Name: {0}", "context.MessageName");

                if (context.MessageName == "Delete")
                {
                    EntityReference entity = (EntityReference)context.InputParameters["Target"];

                    if (entity.LogicalName != caps_ActualDraw.EntityLogicalName)
                        return;

                    tracingService.Trace("{0}", "Delete");
                    Entity preImage = context.PreEntityImages["preImage"];

                    if (preImage.Contains("caps_project"))
                    {
                        tracingService.Trace("{0}", "Have pre-Image");
                        projectsToUpdate.Add(preImage.GetAttributeValue<EntityReference>("caps_project"));
                    }

                }

                tracingService.Trace("{0}", "Line 66");

                foreach (var projectToUpdate in projectsToUpdate)
                {
                    //get all project cash flow records for the project, then for each year on those breakdown the quarters, get the actual draws and update.
                    //might need another plugin on project cash flow record creation too.
                    using (var crmContext = new CrmServiceContext(service))
                    {
                        var projectCashFlowRecords = crmContext.caps_projectcashflowSet.Where(p => p.caps_PTR.Id == projectToUpdate.Id);

                        //for each cash flow record
                        foreach(var cashFlowRecord in projectCashFlowRecords)
                        {
                            //get Fiscal Year
                            var fiscalYear = cashFlowRecord.caps_FiscalYear;

                            //get all actual draws for this project and this year
                            var actualDrawRecords = crmContext.caps_ActualDrawSet.Where(r => r.caps_Project.Id == projectToUpdate.Id && r.caps_FiscalYear.Id == fiscalYear.Id);

                            //decimal Q1, Q2, Q3, Q4;
                            //Q1 = Q2 = Q3 = Q4 = 0;
                            decimal total = 0;
                            
                            foreach(var actualDrawRecord in actualDrawRecords)
                            {
                                if (actualDrawRecord.caps_DrawDate.HasValue)
                                {
                                    total += actualDrawRecord.caps_Amount.GetValueOrDefault(0);

                                }
                            }
                            tracingService.Trace("{0}", total);

                            //Now update the cash flow record
                            var cashFlowToUpdate = new caps_projectcashflow();
                            cashFlowToUpdate.Id = cashFlowRecord.Id;
                            cashFlowToUpdate.caps_TotalActualDraws = total;
                            service.Update(cashFlowToUpdate);

                        }

                    }
                }
            }
        }
    }
}
