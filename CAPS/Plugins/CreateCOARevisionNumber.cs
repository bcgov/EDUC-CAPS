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
    /// This plugin runs on pre-create of COA.
    /// </summary>
    public class CreateCOARevisionNumber : IPlugin
    {
        public void Execute(IServiceProvider serviceProvider)
        {
            IPluginExecutionContext context = (IPluginExecutionContext)serviceProvider.GetService(typeof(IPluginExecutionContext));
            IOrganizationServiceFactory serviceFactory = (IOrganizationServiceFactory)serviceProvider.GetService(typeof(IOrganizationServiceFactory));
            IOrganizationService service = serviceFactory.CreateOrganizationService(context.UserId);
            ITracingService tracingService = (ITracingService)serviceProvider.GetService(typeof(ITracingService));

            tracingService.Trace("{0}", "CreateCOARevisionNumber Plug-in");

            // The InputParameters collection contains all the data passed in the message request.
            if (context.InputParameters.Contains("Target") && context.InputParameters["Target"] is Entity)
            {
                Entity entity = (Entity)context.InputParameters["Target"];

                if (entity.LogicalName != caps_CertificateofApproval.EntityLogicalName || context.MessageName != "Create")
                    return;

                if (!entity.Contains("caps_ptr")) throw new Exception("Project is a required field.");
                //Get COA PTR
                var projectTracker = entity.GetAttributeValue<EntityReference>("caps_ptr");


                using (var crmContext = new CrmServiceContext(service))
                {
                    //Check if there are any other COAs for this project in a Draft  or Submitted State
                    var records = crmContext.caps_CertificateofApprovalSet.Where(r => r.caps_PTR.Id == projectTracker.Id
                                && r.caps_CertificateofApprovalId != context.PrimaryEntityId
                                && r.StatusCode.Value != (int)caps_CertificateofApproval_StatusCode.Cancelled
                                && r.StatusCode.Value !=(int)caps_CertificateofApproval_StatusCode.Rejected).ToList();

                    var revisionNumber = 1;
                    if (records.Count() > 0)
                    {
                        //Get top value
                        var topRecord = records.OrderByDescending(r => r.caps_RevisionNumber).Take(1).ToList();
                        revisionNumber = topRecord[0].caps_RevisionNumber.Value + 1;
                    }
                    entity["caps_revisionnumber"] = revisionNumber;

                }
            }
        }
    }
}
