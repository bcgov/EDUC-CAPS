using CAPS.DataContext;
using Microsoft.Xrm.Sdk;
using Microsoft.Xrm.Sdk.Workflow;
using System;
using System.Activities;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace CustomWorkflowActivities
{
    public class GetTeamByName : CodeActivity
    {
        [Input("TeamName")]
        public InArgument<string> teamName { get; set; }

        [Output("Team")]
        [ReferenceTarget("team")]
        public OutArgument<EntityReference> team { get; set; }

        protected override void Execute(CodeActivityContext executionContext)
        {
            ITracingService tracingService = executionContext.GetExtension<ITracingService>();

            IWorkflowContext context = executionContext.GetExtension<IWorkflowContext>();
            IOrganizationServiceFactory serviceFactory = executionContext.GetExtension<IOrganizationServiceFactory>();
            IOrganizationService service = serviceFactory.CreateOrganizationService(context.UserId);

            tracingService.Trace("{0}{1}", "Start Custom Workflow Activity: GetTeamByName", DateTime.Now.ToLongTimeString());

            using (var crmContext = new CrmServiceContext(service))
            {
                var records = crmContext.TeamSet.Where(r => r.Name == this.teamName.Get(executionContext)).ToList();

                if (records.Count == 1)
                {
                    EntityReference teamReference = new EntityReference(Team.EntityLogicalName, records[0].Id);
                    this.team.Set(executionContext, teamReference);
                }
                else
                {
                    throw new Exception("Team not found or multiple teams with the same name found.");
                }
            }
        }
    }
}
