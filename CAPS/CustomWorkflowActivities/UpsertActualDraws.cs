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
    /// Called via Action, this CWA takes in COA draw information from CSV file and creates or updates actual draw records.
    /// </summary>
    public class UpsertActualDraws : CodeActivity
    {
        [Input("COA Number")]
        public InArgument<string> coa { get; set; }

        [Input("Revision Number")]
        public InArgument<string> revision { get; set; }

        [Input("Process Date")]
        public InArgument<string> processDate { get; set; }

        [Input("Payment Date")]
        public InArgument<string> paymentDate { get; set; }

        [Input("Amount")]
        public InArgument<string> amount { get; set; }

        [Input("Status")]
        public InArgument<string> status { get; set; }

        [Output("Failed")]
        public OutArgument<bool> failed { get; set; }

        [Output("FailureMessage")]
        public OutArgument<string> failureMessage { get; set; }

        protected override void Execute(CodeActivityContext executionContext)
        {
            ITracingService tracingService = executionContext.GetExtension<ITracingService>();

            IWorkflowContext context = executionContext.GetExtension<IWorkflowContext>();
            IOrganizationServiceFactory serviceFactory = executionContext.GetExtension<IOrganizationServiceFactory>();
            IOrganizationService service = serviceFactory.CreateOrganizationService(context.UserId);

            tracingService.Trace("{0}{1}", "Start Custom Workflow Activity: UpsertActualDraws", DateTime.Now.ToLongTimeString());
            try
            {
                //Get values from inputs
                var coaValue = this.coa.Get(executionContext);
                var revisionValue = this.revision.Get(executionContext);
                var processDateValue = this.processDate.Get(executionContext);
                var paymentDateValue = this.paymentDate.Get(executionContext);
                var amountValue = this.amount.Get(executionContext);
                var statusValue = (this.status.Get(executionContext) == "ACTUAL") ? (int)caps_ActualDraw_StatusCode.Approved : (int)caps_ActualDraw_StatusCode.Pending;


                tracingService.Trace("Processed Date: {0}", processDateValue);
                tracingService.Trace("Paid Date: {0}", paymentDateValue);

                //Parse values
                var dateProcessed = DateTime.ParseExact(processDateValue, "yyyyMMdd", System.Globalization.CultureInfo.InvariantCulture);
                tracingService.Trace("Line {0}", "68");
                var datePaid = DateTime.ParseExact(paymentDateValue, "yyyyMMdd", System.Globalization.CultureInfo.InvariantCulture);
                tracingService.Trace("Line {0}", "70");

                tracingService.Trace("Processed Date: {0}", dateProcessed.ToShortDateString());
                tracingService.Trace("Paid Date: {0}", datePaid.ToShortDateString());

                decimal paymentAmount = Convert.ToDecimal(amountValue) / 100;

                tracingService.Trace("Amount: {0}", paymentAmount);
                tracingService.Trace("COA: {0}", coaValue);
                tracingService.Trace("Revision: {0}", revisionValue);

                //Get Project by COA number
                QueryExpression query = new QueryExpression("caps_certificateofapproval");
                query.ColumnSet.AddColumns("caps_ptr");

                FilterExpression filterName = new FilterExpression();
                filterName.Conditions.Add(new ConditionExpression("caps_name", ConditionOperator.Equal, coaValue));
                filterName.Conditions.Add(new ConditionExpression("caps_ptr", ConditionOperator.NotNull));
                filterName.Conditions.Add(new ConditionExpression("caps_revisionnumber", ConditionOperator.Equal, revisionValue));
                filterName.Conditions.Add(new ConditionExpression("statuscode", ConditionOperator.In, (int)caps_CertificateofApproval_StatusCode.Approved, (int)caps_CertificateofApproval_StatusCode.Revised, (int)caps_CertificateofApproval_StatusCode.Closed ));
                query.Criteria.AddFilter(filterName);

                EntityCollection results = service.RetrieveMultiple(query);

                //Get top record and get project info
                if (results.Entities.Count > 0)
                {
                    var project = results.Entities[0].GetAttributeValue<EntityReference>("caps_ptr");
                    var coa = results.Entities[0].Id;

                    //Get School District Team from SD record
                    var fetchXML = "<fetch version=\"1.0\" output-format=\"xml-platform\" mapping=\"logical\" distinct=\"true\">"+
                                    "<entity name=\"edu_schooldistrict\" > "+
                                       "<attribute name=\"edu_name\" /> "+
                                            "<attribute name = \"caps_sddefaultteam\" /> "+
                                             "<order attribute = \"edu_name\" descending=\"false\" /> "+
                                                "<link-entity name = \"caps_projecttracker\" from=\"caps_schooldistrict\" to=\"edu_schooldistrictid\" link-type=\"inner\" alias=\"ac\" > "+
                                                               "<filter type=\"and\" > "+
                                                                  "<condition attribute=\"caps_projecttrackerid\" operator= \"eq\"  value = \"{"+project.Id+"}\" /> "+
                                                                        "</filter> "+
                                                                      "</link-entity> "+
                                                                    "</entity> "+
                                                                  "</fetch> ";
                    var sdRecords = service.RetrieveMultiple(new FetchExpression(fetchXML));


                    tracingService.Trace("Project: {0}", project.Id);

                    //check if actual draw already exists
                    QueryExpression queryDraw = new QueryExpression("caps_actualdraw");
                    queryDraw.ColumnSet.AddColumn("caps_actualdrawid");

                    FilterExpression filterDraw = new FilterExpression();
                    filterDraw.Conditions.Add(new ConditionExpression("caps_project", ConditionOperator.Equal, project.Id));
                    filterDraw.Conditions.Add(new ConditionExpression("caps_drawdate", ConditionOperator.Equal, datePaid));
                    queryDraw.Criteria.AddFilter(filterDraw);

                    EntityCollection resultsDraw = service.RetrieveMultiple(queryDraw);

                    if (resultsDraw.Entities.Count > 0)
                    {
                        tracingService.Trace("Update Existing Record:{0}", resultsDraw.Entities[0].Id);
                        //Update Record
                        caps_ActualDraw recordToUpdate = new caps_ActualDraw();
                        recordToUpdate.Id = resultsDraw.Entities[0].Id;
                        recordToUpdate.caps_Project = project;
                        recordToUpdate.caps_CertificateOfApproval = new EntityReference(caps_CertificateofApproval.EntityLogicalName, coa);
                        recordToUpdate.caps_DrawDate = datePaid;
                        recordToUpdate.caps_ProcessDate = dateProcessed;
                        recordToUpdate.caps_Amount = paymentAmount;
                        recordToUpdate.StatusCode = new OptionSetValue(statusValue);
                        if (sdRecords.Entities.Count > 0)
                        {
                            recordToUpdate.OwnerId = sdRecords.Entities[0].GetAttributeValue<EntityReference>("caps_sddefaultteam");
                        }
                        service.Update(recordToUpdate);
                    }
                    else
                    {
                        tracingService.Trace("New Record: {0}", "created");
                        //Create new record
                        caps_ActualDraw recordToCreate = new caps_ActualDraw();
                        recordToCreate.caps_Project = project;
                        recordToCreate.caps_CertificateOfApproval = new EntityReference(caps_CertificateofApproval.EntityLogicalName, coa);
                        recordToCreate.caps_DrawDate = datePaid;
                        recordToCreate.caps_ProcessDate = dateProcessed;
                        recordToCreate.caps_Amount = paymentAmount;
                        recordToCreate.StatusCode = new OptionSetValue(statusValue);
                        if (sdRecords.Entities.Count > 0)
                        {
                            recordToCreate.OwnerId = sdRecords.Entities[0].GetAttributeValue<EntityReference>("caps_sddefaultteam");
                        }

                        //Get fiscal year
                        QueryExpression queryYear = new QueryExpression("edu_year");
                        queryYear.ColumnSet.AddColumn("edu_yearid");

                        FilterExpression filterYear = new FilterExpression();
                        filterYear.Conditions.Add(new ConditionExpression("edu_startdate", ConditionOperator.OnOrBefore, datePaid));
                        filterYear.Conditions.Add(new ConditionExpression("edu_enddate", ConditionOperator.OnOrAfter, datePaid));
                        filterYear.Conditions.Add(new ConditionExpression("edu_type", ConditionOperator.Equal, 757500000));
                        queryYear.Criteria.AddFilter(filterYear);

                        EntityCollection resultsYear = service.RetrieveMultiple(queryYear);

                        tracingService.Trace("Number of Year Records:{0}", resultsYear.Entities.Count);

                        if (resultsYear.Entities.Count == 1)
                        {
                            recordToCreate.caps_FiscalYear = new EntityReference(edu_Year.EntityLogicalName, resultsYear.Entities[0].Id);
                        }

                        service.Create(recordToCreate);
                    }
                    this.failed.Set(executionContext, false);
                }
                else
                {
                    this.failed.Set(executionContext, true);
                    this.failureMessage.Set(executionContext, "Unable to find matching COA for: " + coaValue+" Revision: "+revisionValue);
                }
            }
            catch(Exception ex)
            {
                this.failed.Set(executionContext, true);
                this.failureMessage.Set(executionContext, "Error:" + ex.Message);
            }
        }
    }
}
