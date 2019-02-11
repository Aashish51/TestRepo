using Microsoft.Xrm.Sdk;
using Microsoft.Xrm.Sdk.Query;
using System;

namespace SeeLogic.SurveyResponseOpportunity
{
    public class Opportunity : IPlugin
    {
        #region Secure/Unsecure Configuration Setup
        private string _secureConfig = null;
        private string _unsecureConfig = null;

        public Opportunity(string unsecureConfig, string secureConfig)
        {
            _secureConfig = secureConfig;
            _unsecureConfig = unsecureConfig;
        }
        #endregion
        public void Execute(IServiceProvider serviceProvider)
        {
            ITracingService tracer = (ITracingService)serviceProvider.GetService(typeof(ITracingService));
            IPluginExecutionContext context = (IPluginExecutionContext)serviceProvider.GetService(typeof(IPluginExecutionContext));
            IOrganizationServiceFactory factory = (IOrganizationServiceFactory)serviceProvider.GetService(typeof(IOrganizationServiceFactory));
            IOrganizationService service = factory.CreateOrganizationService(context.UserId);

            try
            {
                Entity entity = (Entity)context.InputParameters["Target"];

                string surveyName = entity.GetAttributeValue<string>("msdyn_name");

                string lastName = entity.GetAttributeValue<string>("msdyn_lastnameresponse");

                string email = entity.GetAttributeValue<string>("msdyn_emailresponse");

                string company = entity.GetAttributeValue<string>("msdyn_companyresponse");

                FilterExpression codeFilter = new FilterExpression(LogicalOperator.And);


                if (!string.IsNullOrEmpty(lastName))
                    codeFilter.AddCondition("lastname", ConditionOperator.Equal, lastName);

                if (!string.IsNullOrEmpty(email))
                    codeFilter.AddCondition("emailaddress1", ConditionOperator.Equal, email);

                QueryExpression query = new QueryExpression
                {

                    EntityName = "contact",

                    ColumnSet = new ColumnSet(true),  // we assume you want to retrieve all the fields

                    Criteria = codeFilter

                };

                FilterExpression accountFilter = new FilterExpression(LogicalOperator.And);
                if (!string.IsNullOrEmpty(company))
                { 
                    accountFilter.AddCondition("name", ConditionOperator.Equal, company);

                }

                QueryExpression accountQuery = new QueryExpression
                {

                    EntityName = "account",

                    ColumnSet = new ColumnSet(true),  // we assume you want to retrieve all the fields

                    Criteria = accountFilter

                };

                //queryAccounts.ColumnSet = new ColumnSet(true);
                EntityCollection contact = service.RetrieveMultiple(query);
                EntityCollection account = service.RetrieveMultiple(accountQuery);

                Entity opportunity = new Entity("opportunity");

                opportunity["name"] = surveyName;
                opportunity["slgc_fromvoc"] = true;
                
                int opportunityTypeValue = 0;

                if (surveyName.ToLower().Contains("knowledge transfer"))
                {
                    opportunityTypeValue = 297950002;
                }
                else if (surveyName.ToLower().Contains("re-use"))
                {
                    opportunityTypeValue = 297950000;
                }
                else if (surveyName.ToLower().Contains("speaker"))
                {
                    opportunityTypeValue = 297950001;
                }


                if (opportunityTypeValue>0)
                {
                    opportunity["slgc_ipopportunitytype"] = new OptionSetValue(opportunityTypeValue);
                }
                if(entity.Id!=null && entity.Id!=new Guid())
                {
                    opportunity["slgc_surveyresponse"] = new EntityReference("msdyn_surveyresponse", entity.Id);
                }
                     

                if (contact.Entities.Count > 0)
                    opportunity["parentcontactid"] = new EntityReference("contact", Guid.Parse(contact.Entities[0]["contactid"].ToString()));

                if (account.Entities.Count > 0 &&  !string.IsNullOrEmpty(company))
                    opportunity["parentaccountid"] = new EntityReference("account", Guid.Parse(account.Entities[0]["accountid"].ToString()));

                opportunity["slgc_type"] = new OptionSetValue(297950001);
                service.Create(opportunity);
                 

                //TODO: Do stuff
            }
            catch (Exception e)
            {
                throw new InvalidPluginExecutionException(e.Message);
            }
        }
    }
}