using System;
using Microsoft.Xrm.Sdk;
using Microsoft.Xrm.Sdk.Query;

namespace Plugin
{
    public class PluginCSCode : IPlugin
    {
        public void Execute(IServiceProvider serviceProvider)
        {
            // Initialize plugin services
            var (tracingService, pluginExecutionContext, organizationService) = InitializeServices(serviceProvider);

            tracingService.Trace("PluginCSCode execution started.");

            try
            {
                // Execute the main plugin logic
                ExecutePlugin(pluginExecutionContext, organizationService, tracingService);
                tracingService.Trace("PluginCSCode executed successfully. Case Created.");
            }
            catch (Exception ex)
            {
                // Log and re-throw any exceptions encountered during execution
                tracingService.Trace("Error: {0}", ex.Message);
                throw;
            }
        }

        /// <summary>
        /// Initializes the necessary services for the plugin.
        /// </summary>
        /// <param name="serviceProvider">Service provider from the plugin execution context.</param>
        /// <returns>Tuple containing tracing service, plugin execution context, and organization service.</returns>
        private (ITracingService tracingService, IPluginExecutionContext pluginExecutionContext, IOrganizationService organizationService)
            InitializeServices(IServiceProvider serviceProvider)
        {
            // Extract tracing service for logging plugin execution details
            var tracingService = (ITracingService)serviceProvider.GetService(typeof(ITracingService));
            // Extract plugin execution context to get information about the triggering event
            var pluginExecutionContext = (IPluginExecutionContext)serviceProvider.GetService(typeof(IPluginExecutionContext));
            // Extract organization service factory to interact with Dynamics 365 data
            var serviceFactory = (IOrganizationServiceFactory)serviceProvider.GetService(typeof(IOrganizationServiceFactory));
            // Create an organization service instance to interact with CRM
            var organizationService = serviceFactory.CreateOrganizationService(pluginExecutionContext.UserId);

            return (tracingService, pluginExecutionContext, organizationService);
        }

        /// <summary>
        /// Main execution logic for the plugin.
        /// </summary>
        /// <param name="context">Plugin execution context.</param>
        /// <param name="service">Organization service to interact with CRM data.</param>
        /// <param name="tracingService">Tracing service for logging details.</param>
        private void ExecutePlugin(IPluginExecutionContext context, IOrganizationService service, ITracingService tracingService)
        {
            // Check if the context is valid (must be an "incident" entity with a target)
            if (!IsValidContext(context))
            {
                tracingService.Trace("Invalid context: Not an 'incident' entity or missing target.");
                return;
            }

            // Get the target entity from the context
            Entity targetEntity = (Entity)context.InputParameters["Target"];

            // Validate that the "customerid" field exists and is valid
            ValidateCustomerId(targetEntity);

            // Extract the Customer ID from the "customerid" attribute
            Guid customerId = ((EntityReference)targetEntity["customerid"]).Id;

            tracingService.Trace("Validated Customer ID: {0}", customerId);

            // Check if there are any active cases for the customer
            if (HasActiveCases(service, customerId))
            {
                tracingService.Trace("Active case exists for Customer ID: {0}", customerId);
                // Throw an exception if an active case already exists for the customer
                throw new InvalidPluginExecutionException("A case for this customer is already active. Unable to create Case.");
            }

            tracingService.Trace("No active cases found for Customer ID: {0}. Proceeding.", customerId);
        }

        /// <summary>
        /// Validates the context to ensure it contains a valid "incident" entity.
        /// </summary>
        /// <param name="context">Plugin execution context.</param>
        /// <returns>True if the context is valid; otherwise, false.</returns>
        private bool IsValidContext(IPluginExecutionContext context)
        {
            return context.InputParameters.Contains("Target") &&
                   context.InputParameters["Target"] is Entity entity &&
                   entity.LogicalName == "incident";
        }

        /// <summary>
        /// Validates that the "customerid" attribute exists and is valid in the entity.
        /// </summary>
        /// <param name="entity">The target entity.</param>
        private void ValidateCustomerId(Entity entity)
        {
            if (!entity.Attributes.Contains("customerid") || !(entity["customerid"] is EntityReference))
            {
                // Throw an exception if the "customerid" attribute is missing or invalid
                throw new InvalidPluginExecutionException("The Customer ID is required to create a case.");
            }
        }

        /// <summary>
        /// Checks if there are any active cases for the specified customer.
        /// </summary>
        /// <param name="service">Organization service to interact with CRM data.</param>
        /// <param name="customerId">The Customer ID to check for active cases.</param>
        /// <returns>True if there are active cases; otherwise, false.</returns>
        private bool HasActiveCases(IOrganizationService service, Guid customerId)
        {
            // Create and execute the query to check for active cases
            var query = CreateActiveCasesQuery(customerId);
            var results = service.RetrieveMultiple(query);
            // Return true if any active cases are found
            return results.Entities.Count > 0;
        }

        /// <summary>
        /// Creates a query to retrieve active cases for the specified customer.
        /// </summary>
        /// <param name="customerId">The Customer ID to filter active cases.</param>
        /// <returns>A QueryExpression object to retrieve active cases.</returns>
        private QueryExpression CreateActiveCasesQuery(Guid customerId)
        {
            return new QueryExpression("incident")
            {
                ColumnSet = new ColumnSet("statecode"), // Retrieve only the "statecode" field to minimize data retrieval
                Criteria =
                {
                    Conditions =
                    {
                        new ConditionExpression("customerid", ConditionOperator.Equal, customerId), // Filter by Customer ID
                        new ConditionExpression("statecode", ConditionOperator.Equal, 0) // Filter by active cases (statecode = 0)
                    }
                }
            };
        }
    }
}








//using System;
//using Microsoft.Xrm.Sdk;
//using Microsoft.Xrm.Sdk.Query;

//namespace Plugin
//{
//    public class PluginCSCode : IPlugin
//    {
//        public void Execute(IServiceProvider serviceProvider)
//        {
//            // Extract services
//            var tracingService = (ITracingService)serviceProvider.GetService(typeof(ITracingService));
//            var pluginExecutionContext = (IPluginExecutionContext)serviceProvider.GetService(typeof(IPluginExecutionContext));
//            var serviceFactory = (IOrganizationServiceFactory)serviceProvider.GetService(typeof(IOrganizationServiceFactory));
//            var organizationService = serviceFactory.CreateOrganizationService(pluginExecutionContext.UserId);

//            tracingService.Trace("PluginCSCode execution started.");

//            try
//            {
//                ExecutePlugin(pluginExecutionContext, organizationService, tracingService);
//                tracingService.Trace("PluginCSCode executed successfully.");
//            }
//            catch (Exception ex)
//            {
//                tracingService.Trace("Error: {0}", ex.Message);
//                throw;
//            }
//        }

//        private void ExecutePlugin(IPluginExecutionContext context, IOrganizationService service, ITracingService tracingService)
//        {
//            if (!IsValidContext(context))
//            {
//                tracingService.Trace("Invalid context: Not an 'incident' entity or missing target.");
//                return;
//            }

//            Entity targetEntity = (Entity)context.InputParameters["Target"];

//            ValidateCustomerId(targetEntity);

//            Guid customerId = ((EntityReference)targetEntity["customerid"]).Id;

//            tracingService.Trace("Validated Customer ID: {0}", customerId);

//            if (HasActiveCases(service, customerId))
//            {
//                tracingService.Trace("Active case exists for Customer ID: {0}", customerId);
//                throw new InvalidPluginExecutionException("A case for this customer is already active. Unable to create Case.");
//            }

//            tracingService.Trace("No active cases found for Customer ID: {0}. Proceeding.", customerId);
//        }

//        private bool IsValidContext(IPluginExecutionContext context)
//        {
//            return context.InputParameters.Contains("Target") &&
//                   context.InputParameters["Target"] is Entity entity &&
//                   entity.LogicalName == "incident";
//        }

//        private void ValidateCustomerId(Entity entity)
//        {
//            if (!entity.Attributes.Contains("customerid") || !(entity["customerid"] is EntityReference))
//            {
//                throw new InvalidPluginExecutionException("The Customer ID is required to create a case.");
//            }
//        }

//        private bool HasActiveCases(IOrganizationService service, Guid customerId)
//        {
//            var query = CreateActiveCasesQuery(customerId);
//            var results = service.RetrieveMultiple(query);
//            return results.Entities.Count > 0;
//        }

//        private QueryExpression CreateActiveCasesQuery(Guid customerId)
//        {
//            return new QueryExpression("incident")
//            {
//                ColumnSet = new ColumnSet("statecode"),
//                Criteria =
//                {
//                    Conditions =
//                    {
//                        new ConditionExpression("customerid", ConditionOperator.Equal, customerId),
//                        new ConditionExpression("statecode", ConditionOperator.Equal, 0) // Active cases
//                    }
//                }
//            };
//        }
//    }
//}










