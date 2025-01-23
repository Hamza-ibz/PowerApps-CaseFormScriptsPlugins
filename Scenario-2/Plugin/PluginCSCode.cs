using System;
using Microsoft.Xrm.Sdk;
using Microsoft.Xrm.Sdk.Query;

namespace Plugin
{
    public class PluginCSCode : IPlugin
    {
        public void Execute(IServiceProvider serviceProvider)
        {
            // Extract services
            var tracingService = (ITracingService)serviceProvider.GetService(typeof(ITracingService));
            var pluginExecutionContext = (IPluginExecutionContext)serviceProvider.GetService(typeof(IPluginExecutionContext));
            var serviceFactory = (IOrganizationServiceFactory)serviceProvider.GetService(typeof(IOrganizationServiceFactory));
            var organizationService = serviceFactory.CreateOrganizationService(pluginExecutionContext.UserId);

            tracingService.Trace("PluginCSCode execution started.");

            try
            {
                ExecutePlugin(pluginExecutionContext, organizationService, tracingService);
                tracingService.Trace("PluginCSCode executed successfully.");
            }
            catch (Exception ex)
            {
                tracingService.Trace("Error: {0}", ex.Message);
                throw;
            }
        }

        private void ExecutePlugin(IPluginExecutionContext context, IOrganizationService service, ITracingService tracingService)
        {
            if (!IsValidContext(context))
            {
                tracingService.Trace("Invalid context: Not an 'incident' entity or missing target.");
                return;
            }

            Entity targetEntity = (Entity)context.InputParameters["Target"];

            ValidateCustomerId(targetEntity);

            Guid customerId = ((EntityReference)targetEntity["customerid"]).Id;

            tracingService.Trace("Validated Customer ID: {0}", customerId);

            if (HasActiveCases(service, customerId))
            {
                tracingService.Trace("Active case exists for Customer ID: {0}", customerId);
                throw new InvalidPluginExecutionException("A case for this customer is already active. Unable to create Case.");
            }

            tracingService.Trace("No active cases found for Customer ID: {0}. Proceeding.", customerId);
        }

        private bool IsValidContext(IPluginExecutionContext context)
        {
            return context.InputParameters.Contains("Target") &&
                   context.InputParameters["Target"] is Entity entity &&
                   entity.LogicalName == "incident";
        }

        private void ValidateCustomerId(Entity entity)
        {
            if (!entity.Attributes.Contains("customerid") || !(entity["customerid"] is EntityReference))
            {
                throw new InvalidPluginExecutionException("The Customer ID is required to create a case.");
            }
        }

        private bool HasActiveCases(IOrganizationService service, Guid customerId)
        {
            var query = CreateActiveCasesQuery(customerId);
            var results = service.RetrieveMultiple(query);
            return results.Entities.Count > 0;
        }

        private QueryExpression CreateActiveCasesQuery(Guid customerId)
        {
            return new QueryExpression("incident")
            {
                ColumnSet = new ColumnSet("statecode"),
                Criteria =
                {
                    Conditions =
                    {
                        new ConditionExpression("customerid", ConditionOperator.Equal, customerId),
                        new ConditionExpression("statecode", ConditionOperator.Equal, 0) // Active cases
                    }
                }
            };
        }
    }
}










