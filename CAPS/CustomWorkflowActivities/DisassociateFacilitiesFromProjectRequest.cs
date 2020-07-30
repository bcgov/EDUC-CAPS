using Microsoft.Xrm.Sdk;
using Microsoft.Xrm.Sdk.Query;
using Microsoft.Xrm.Sdk.Workflow;
using System;
using System.Activities;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using CAPS.DataContext;

namespace CustomWorkflowActivities
{
    public class DisassociateFacilitiesFromProjectRequest : CodeActivity
    {
        protected override void Execute(CodeActivityContext executionContext)
        {
            ITracingService tracingService = executionContext.GetExtension<ITracingService>();

            IWorkflowContext context = executionContext.GetExtension<IWorkflowContext>();
            IOrganizationServiceFactory serviceFactory = executionContext.GetExtension<IOrganizationServiceFactory>();
            IOrganizationService service = serviceFactory.CreateOrganizationService(context.UserId);

            tracingService.Trace("{0}{1}", "Start Custom Workflow Activity: DisassociateFacilitiesFromProjectRequest", DateTime.Now.ToLongTimeString());

            var recordId = context.PrimaryEntityId;


            string fetchXml = @"<fetch version='1.0' mapping='logical'>" +
                                "<entity name='caps_facility' >"+
                                "<attribute name='caps_facilityid' />"+
                                "<link-entity name='caps_project_caps_facility' from='caps_facilityid' to='caps_facilityid' visible='false' intersect='true' >" +
                                "<link-entity name='caps_project' from='caps_projectid' to='caps_projectid' alias='ac' >"+
                                     "<filter type='and' >"+
                                        "<condition attribute = 'caps_projectid' operator= 'eq'  value = '{"+recordId+"}' />"+
                                              "</filter>"+
                                            "</link-entity>"+
                                          "</link-entity>"+
                                        "</entity>"+
                                      "</fetch> ";

            tracingService.Trace("Fetch: {0}", fetchXml);

            EntityCollection collRecords = service.RetrieveMultiple(new FetchExpression(fetchXml));

            if (collRecords != null && collRecords.Entities != null && collRecords.Entities.Count > 0)
            {
                EntityReferenceCollection collection = new EntityReferenceCollection();
                foreach (var entity in collRecords.Entities)
                {
                    var reference = new EntityReference("caps_facility", entity.GetAttributeValue<Guid>("caps_facilityid"));
                    collection.Add(reference); //Create a collection of entity references
                }
                Relationship relationship = new Relationship("caps_Project_caps_Facility"); //schema name of N:N relationship
                service.Disassociate("caps_project", recordId, relationship, collection); //Pass the entity reference collections to be disassociated from the specific Email Send record
            }

        }
    }
}
