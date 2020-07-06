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
    /// <summary>
    /// Called by flow via action to get next project number for AFG and BUS projects.
    /// </summary>
    public class GetNextProjectRequestNumber : CodeActivity
    {
        [Output("ProjectRequestNumber")]
        public OutArgument<string> projectRequestNumber { get; set; }

        protected override void Execute(CodeActivityContext executionContext)
        {
            ITracingService tracingService = executionContext.GetExtension<ITracingService>();

            IWorkflowContext context = executionContext.GetExtension<IWorkflowContext>();
            IOrganizationServiceFactory serviceFactory = executionContext.GetExtension<IOrganizationServiceFactory>();
            IOrganizationService service = serviceFactory.CreateOrganizationService(context.UserId);

            tracingService.Trace("{0}{1}", "Start Custom Workflow Activity: GetNextProjectRequestNumber", DateTime.Now.ToLongTimeString());

            //Get Project field
            var projectToCreate = new caps_Project();
            var projectId = service.Create(projectToCreate);

            //get project Id
            Microsoft.Xrm.Sdk.Query.ColumnSet columns = new Microsoft.Xrm.Sdk.Query.ColumnSet("caps_projectnumber");
            var projectRecord = service.Retrieve(caps_Project.EntityLogicalName, projectId, columns) as caps_Project;

            var projectNumber = projectRecord.caps_ProjectNumber;

            tracingService.Trace("Project Number:{0}", projectNumber);

            projectRequestNumber.Set(executionContext, projectNumber);

            //delete project
            service.Delete(caps_Project.EntityLogicalName, projectId);

        }
    }
}
