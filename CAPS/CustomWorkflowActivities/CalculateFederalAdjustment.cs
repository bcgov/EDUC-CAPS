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
    /// This CWA validates the total federal adjustment amount on all actual draws associated to a project.  
    /// It returns true if the amount exceeds the federal funding amount and otherwise returns false.
    /// </summary>
    public class CalculateFederalAdjustment : CodeActivity
    {
        [Input("Project")]
        [ReferenceTarget("caps_projecttracker")]
        public InArgument<EntityReference> project { get; set; }

        [Output("Exceeds Total")]
        public OutArgument<bool> exceedsTotal { get; set; }

        protected override void Execute(CodeActivityContext executionContext)
        {
            ITracingService tracingService = executionContext.GetExtension<ITracingService>();

            IWorkflowContext context = executionContext.GetExtension<IWorkflowContext>();
            IOrganizationServiceFactory serviceFactory = executionContext.GetExtension<IOrganizationServiceFactory>();
            IOrganizationService service = serviceFactory.CreateOrganizationService(context.UserId);

            tracingService.Trace("{0}{1}", "Start Custom Workflow Activity: CalculateFederalAdjustment", DateTime.Now.ToLongTimeString());

            var recordId = context.PrimaryEntityId;

            EntityReference projectReference = this.project.Get(executionContext);

            var projectRecord = service.Retrieve(projectReference.LogicalName, projectReference.Id, new ColumnSet("caps_federal")) as caps_ProjectTracker;

            //Get total of all federal adjustments for the project
            
            var fetchAdjustments = "<fetch version=\"1.0\" output-format = \"xml-platform\" mapping = \"logical\" distinct =\"false\" >"+
                                "<entity name =\"caps_actualdraw\" >"+               
                                    "<attribute name =\"caps_amount\" />"+
                                    "<attribute name = \"caps_certificateofapproval\" />"+
                                    "<order attribute = \"caps_drawdate\" descending = \"true\" />"+
                                    "<filter type = \"and\" >"+
                                   "<condition attribute = \"caps_project\" operator= \"eq\" value = \"{"+ projectReference.Id + "}\" />"+
                                   "<condition attribute = \"statuscode\" operator= \"eq\" value = \"200870001\" />"+
                                    "</filter >"+
                                    "</entity >"+
                                    "</fetch >";

            //Summ all adjustment values
            decimal totalAdjustments = 0;

            var estimatedExpenditures = service.RetrieveMultiple(new FetchExpression(fetchAdjustments));

            foreach(var draw in estimatedExpenditures.Entities)
            {
                totalAdjustments += draw.GetAttributeValue<decimal?>("caps_amount").GetValueOrDefault(0);
            }

            tracingService.Trace("Adjustment Total: {0}", totalAdjustments);

            if (projectRecord.caps_Federal.HasValue &&  Math.Abs(totalAdjustments) > projectRecord.caps_Federal.Value)
            {
                //return error
                this.exceedsTotal.Set(executionContext, true);
            }
            else
            {
                this.exceedsTotal.Set(executionContext, false);
            }

            
            

        }
    }
}
