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
    /// This plugin runs on change of design capacity fields on facility.  It finds the previous and current family of schools and updates the totals.
    /// </summary>
    /*public class CalculateFamilyOfSchoolsEnrolment : IPlugin
    {
        public void Execute(IServiceProvider serviceProvider)
        {
            IPluginExecutionContext context = (IPluginExecutionContext)serviceProvider.GetService(typeof(IPluginExecutionContext));
            IOrganizationServiceFactory serviceFactory = (IOrganizationServiceFactory)serviceProvider.GetService(typeof(IOrganizationServiceFactory));
            IOrganizationService service = serviceFactory.CreateOrganizationService(context.UserId);
            ITracingService tracingService = (ITracingService)serviceProvider.GetService(typeof(ITracingService));

            tracingService.Trace("{0}", "CalculateStrongStartAdjustment Plug-in");

            // The InputParameters collection contains all the data passed in the message request.
            if (context.InputParameters.Contains("Target")
                && (context.InputParameters["Target"] is Entity || context.InputParameters["Target"] is EntityReference))
            {
                var familiesToUpdate = new EntityReferenceCollection();

                if (context.MessageName == "Create" || context.MessageName == "Update")
                {
                    Entity entity = (Entity)context.InputParameters["Target"];

                    if (entity.LogicalName != caps_Facility.EntityLogicalName)
                        return;

                    //Keep CRUD in mind, need to handle all scenarios
                    //Get previous family of schools (if exists) and current family of schools
                    if (context.MessageName == "Create" && entity.Contains("caps_familyofschools"))
                    {
                        familiesToUpdate.Add(entity.GetAttributeValue<EntityReference>("caps_familyofschools"));
                    }

                    if (context.MessageName == "Update")
                    {
                        Entity preImage = context.PreEntityImages["preImage"];

                        if (preImage.Contains("caps_familyofschools"))
                        {
                            familiesToUpdate.Add(preImage.GetAttributeValue<EntityReference>("caps_familyofschools"));
                        }
                        if (entity.Contains("caps_familyofschools"))
                        {
                            familiesToUpdate.Add(entity.GetAttributeValue<EntityReference>("caps_familyofschools"));
                        }
                    }
                }
                if (context.MessageName == "Delete")
                {
                    EntityReference entity = (EntityReference)context.InputParameters["Target"];

                    if (entity.LogicalName != caps_Facility.EntityLogicalName)
                        return;

                    Entity preImage = context.PreEntityImages["preImage"];

                    if (preImage.Contains("caps_familyofschools"))
                    {
                        familiesToUpdate.Add(preImage.GetAttributeValue<EntityReference>("caps_familyofschools"));
                    }
                }

                tracingService.Trace("{0}", "Line 66");

                foreach (var familyToUpdate in familiesToUpdate)
                {
                    var actualKCount = 0;
                    var actualECount = 0;
                    var actualSCount = 0;
                    var designKCount = 0;
                    var designECount = 0;
                    var designSCount = 0;
                    var operatingKCount = 0;
                    var operatingECount = 0;
                    var operatingSCount = 0;
                    //For each facility, get all strong starts
                    using (var crmContext = new CrmServiceContext(service))
                    {
                        var facilityRecords = crmContext.caps_FacilitySet.Where(r => r.caps_FamilyofSchools.Id == familyToUpdate.Id && r.StateCode == caps_FacilityState.Active);

                        foreach (var facilityRecord in facilityRecords)
                        {
                            actualKCount += facilityRecord.caps_Kindergarten_currentenrolment.GetValueOrDefault(0);
                            actualECount += facilityRecord.caps_Elementary_currentenrolment.GetValueOrDefault(0);
                            actualSCount += facilityRecord.caps_Secondary_currentenrolment.GetValueOrDefault(0);

                            designKCount += facilityRecord.caps_AdjustedDesignCapacityKindergarten.GetValueOrDefault(0);
                            designECount += facilityRecord.caps_AdjustedDesignCapacityElementary.GetValueOrDefault(0);
                            designSCount += facilityRecord.caps_AdjustedDesignCapacitySecondary.GetValueOrDefault(0);

                            operatingKCount += (int)facilityRecord.caps_OperatingCapacityKindergarten.GetValueOrDefault(0);
                            operatingECount += (int)facilityRecord.caps_OperatingCapacityElementary.GetValueOrDefault(0);
                            operatingSCount += (int)facilityRecord.caps_OperatingCapacitySecondary.GetValueOrDefault(0);

                        }

                        //tracingService.Trace("{0} - K:{1}; E:{2}; S:{3}", facilityToUpdate.Name, kindergardenCount, elementaryCount, secondaryCount);
                        //Update the facility
                        var recordToUpdate = new caps_FamilyofSchools();
                        recordToUpdate.Id = familyToUpdate.Id;
                        recordToUpdate.caps_CurrentEnrolmentKindergarten = actualKCount;
                        recordToUpdate.caps_CurrentEnrolmentElementary = actualECount;
                        recordToUpdate.caps_CurrentEnrolmentSecondary = actualSCount;
                        recordToUpdate.caps_CurrentEnrolmentTotal = actualKCount + actualECount + actualSCount;

                        recordToUpdate.caps_CapacityDesignKindergarten = designKCount;
                        recordToUpdate.caps_CapacityDesignElementary = designECount;
                        recordToUpdate.caps_CapacityDesignSecondary = designSCount;
                        recordToUpdate.caps_CapacityDesignTotal = designKCount + designECount + designSCount;

                        recordToUpdate.caps_CapacityOperatingKindergarten = operatingKCount;
                        recordToUpdate.caps_CapacityOperatingElementary = operatingECount;
                        recordToUpdate.caps_CapacityOperatingSecondary = operatingSCount;
                        recordToUpdate.caps_CapacityOperatingTotal = operatingKCount + operatingECount + operatingSCount;

                        service.Update(recordToUpdate);
                    }
                }
            }
        }
    }*/
}
