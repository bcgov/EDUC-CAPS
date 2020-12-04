using Microsoft.Xrm.Sdk;
using Microsoft.Xrm.Sdk.Workflow;
using System;
using System.Activities;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using CAPS.DataContext;
using Microsoft.Xrm.Sdk.Query;
using Microsoft.Xrm.Sdk.Messages;

namespace CustomWorkflowActivities
{
    /// <summary>
    /// This action bulk deletes all Capacity Reporting records.
    /// </summary>
    public class DeleteAllCapacityReportingRecords : CodeActivity
    {
        protected override void Execute(CodeActivityContext executionContext)
        {
            ITracingService tracingService = executionContext.GetExtension<ITracingService>();

            IWorkflowContext context = executionContext.GetExtension<IWorkflowContext>();
            IOrganizationServiceFactory serviceFactory = executionContext.GetExtension<IOrganizationServiceFactory>();
            IOrganizationService service = serviceFactory.CreateOrganizationService(context.UserId);

            tracingService.Trace("{0}{1}", "Start Custom Workflow Activity: DeleteAllCapacityReportingRecords", DateTime.Now.ToLongTimeString());

            int recordCount = 0;

            do
            {
                // Create an ExecuteMultipleRequest object.
                ExecuteMultipleRequest requestWithResults = new ExecuteMultipleRequest()
                {
                    // Assign settings that define execution behavior: continue on error, return responses. 
                    Settings = new ExecuteMultipleSettings()
                    {
                        ContinueOnError = true,
                        ReturnResponses = true
                    },
                    // Create an empty organization request collection.
                    Requests = new OrganizationRequestCollection()
                };

                QueryExpression query = new QueryExpression(caps_CapacityReporting.EntityLogicalName);
                query.ColumnSet.AddColumn("caps_capacityreportingid");
                query.TopCount = 1000;

                EntityCollection reportingRecords = service.RetrieveMultiple(query);
                recordCount = reportingRecords.Entities.Count();

                foreach (var reportingRecord in reportingRecords.Entities)
                {
                    DeleteRequest deleteRequest = new DeleteRequest();
                    deleteRequest.Target = new EntityReference(caps_CapacityReporting.EntityLogicalName, reportingRecord.Id);
                    requestWithResults.Requests.Add(deleteRequest);
                }

                ExecuteMultipleResponse responseWithResults =
                    (ExecuteMultipleResponse)service.Execute(requestWithResults);

            } while (recordCount > 0);
        }
    }
}
