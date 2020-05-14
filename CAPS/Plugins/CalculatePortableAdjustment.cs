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
    /// This plugin runs on change of portable.  It finds the previous and current related facility, and adjusts the portable capacity.
    /// Register on async Post of events (Create, Update, Delete)
    /// </summary>
    public class CalculatePortableAdjustment : IPlugin
    {
        public void Execute(IServiceProvider serviceProvider)
        {
            IPluginExecutionContext context = (IPluginExecutionContext)serviceProvider.GetService(typeof(IPluginExecutionContext));
            IOrganizationServiceFactory serviceFactory = (IOrganizationServiceFactory)serviceProvider.GetService(typeof(IOrganizationServiceFactory));
            IOrganizationService service = serviceFactory.CreateOrganizationService(context.UserId);
            ITracingService tracingService = (ITracingService)serviceProvider.GetService(typeof(ITracingService));

            tracingService.Trace("{0}", "CalculatePortableAdjustment Plug-in");

            // The InputParameters collection contains all the data passed in the message request.
            if (context.InputParameters.Contains("Target") && context.InputParameters["Target"] is Entity)
            {
                Entity entity = (Entity)context.InputParameters["Target"];

                if (entity.LogicalName != caps_Portable.EntityLogicalName)
                    return;

                var facilitiesToUpdate = new EntityReferenceCollection();

                //Keep CRUD in mind, need to handle all scenarios
                //Get previous facility (if exists) and current facility
                if (context.MessageName == "Create" && entity.Contains("caps_facility"))
                {
                    facilitiesToUpdate.Add(entity.GetAttributeValue<EntityReference>("caps_facility"));
                }

                if (context.MessageName == "Update")
                {
                    Entity preImage = context.PreEntityImages["preImage"];

                    if (preImage.Contains("caps_facility"))
                    {
                        facilitiesToUpdate.Add(preImage.GetAttributeValue<EntityReference>("caps_facility"));
                    }
                    if (entity.Contains("caps_facility"))
                    {
                        facilitiesToUpdate.Add(entity.GetAttributeValue<EntityReference>("caps_facility"));
                    }
                }

                if (context.MessageName == "Delete")
                {
                    //Entity preImage = context.PreEntityImages["preImage"];

                    if (entity.Contains("caps_facility"))
                    {
                        facilitiesToUpdate.Add(entity.GetAttributeValue<EntityReference>("caps_facility"));
                    }
                }

                tracingService.Trace("{0}", "Line 66");

                foreach (var facilityToUpdate in facilitiesToUpdate)
                {
                    var kindergardenCount = 0;
                    var elementaryCount = 0;
                    var secondaryCount = 0;
                    //For each facility, get all strong starts
                    using (var crmContext = new CrmServiceContext(service))
                    {
                        var portableRecords = crmContext.caps_PortableSet.Where(r => r.caps_Facility.Id == facilityToUpdate.Id && r.StateCode == caps_PortableState.Active);

                        tracingService.Trace("{0}", "Line 75");
                        foreach (var portableRecord in portableRecords)
                        {
                            if (portableRecord.Contains("caps_classtype"))
                            {
                                if (portableRecord.GetAttributeValue<OptionSetValue>("caps_classtype").Value == (int)caps_ClassType.Kindergarten)
                                {
                                    tracingService.Trace("{0}", "Line 82");
                                    kindergardenCount += portableRecord.caps_Capacity.GetValueOrDefault(0);
                                }

                                if (portableRecord.GetAttributeValue<OptionSetValue>("caps_classtype").Value == (int)caps_ClassType.Elementary)
                                {
                                    tracingService.Trace("{0}", "Line 88");
                                    elementaryCount += portableRecord.caps_Capacity.GetValueOrDefault(0);
                                }

                                if (portableRecord.GetAttributeValue<OptionSetValue>("caps_classtype").Value == (int)caps_ClassType.Secondary)
                                {
                                    tracingService.Trace("{0}", "Line 94");
                                    secondaryCount += portableRecord.caps_Capacity.GetValueOrDefault(0);
                                }
                            }
                        }

                        tracingService.Trace("{0} - K:{1}; E:{2}; S:{3}", facilityToUpdate.Name, kindergardenCount, elementaryCount, secondaryCount);
                        
                        //Update the facility
                        var recordToUpdate = new caps_Facility();
                        recordToUpdate.Id = facilityToUpdate.Id;
                        recordToUpdate.caps_PortableCapacityElementary = elementaryCount;
                        recordToUpdate.caps_PortableCapacityKindergarten = kindergardenCount;
                        recordToUpdate.caps_PortableCapacitySecondary = secondaryCount;
                        service.Update(recordToUpdate);
                    }
                }
            }
        }
    }
}
