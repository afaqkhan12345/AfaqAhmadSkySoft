using System;
using System.ServiceModel;

using Microsoft.Xrm.Sdk;
using Microsoft.Xrm.Sdk.Query;

namespace UpdateAccountName
{
    public class OverrideAccountNameFromLead : IPlugin
    {
        public void Execute(IServiceProvider serviceProvider)
        {
            IPluginExecutionContext context = (IPluginExecutionContext)serviceProvider.GetService(typeof(IPluginExecutionContext));
            IOrganizationServiceFactory serviceFactory = (IOrganizationServiceFactory)serviceProvider.GetService(typeof(IOrganizationServiceFactory));
            IOrganizationService service = serviceFactory.CreateOrganizationService(context.UserId);
            ITracingService tracing = (ITracingService)serviceProvider.GetService(typeof(ITracingService));

            tracing.Trace("Plugin execution started.");

            try
            {
                tracing.Trace("MessageName: {0}", context.MessageName);

                if (context.MessageName != "QualifyLead" || !context.InputParameters.Contains("LeadId"))
                {
                    tracing.Trace("Not QualifyLead message or LeadId not found.");
                    return;
                }

                var leadRef = (EntityReference)context.InputParameters["LeadId"];
                tracing.Trace("Lead ID: {0}", leadRef.Id);

                var lead = service.Retrieve("lead", leadRef.Id, new ColumnSet("tp_groupname"));
                tracing.Trace("Lead retrieved successfully.");

                string groupName = lead.GetAttributeValue<string>("tp_groupname");
                tracing.Trace("Group Name retrieved: {0}", groupName);

                if (string.IsNullOrEmpty(groupName))
                {
                    tracing.Trace("Group Name is null or empty. Exiting plugin.");
                    return;
                }

                if (context.OutputParameters.Contains("CreatedEntities"))
                {
                    var createdRefs = (EntityReferenceCollection)context.OutputParameters["CreatedEntities"];
                    tracing.Trace("CreatedEntities found: {0} records", createdRefs.Count);

                    foreach (var entityRef in createdRefs)
                    {
                        tracing.Trace("Entity LogicalName: {0}, ID: {1}", entityRef.LogicalName, entityRef.Id);

                        if (entityRef.LogicalName == "account")
                        {
                            tracing.Trace("Updating Account name to Group Name.");

                            var accountUpdate = new Entity("account", entityRef.Id);
                            accountUpdate["name"] = groupName;

                            service.Update(accountUpdate);

                            tracing.Trace("Account updated successfully.");
                        }
                    }
                }
                else
                {
                    tracing.Trace("No CreatedEntities found in OutputParameters.");
                }
            }
            catch (FaultException<OrganizationServiceFault> ex)
            {
                tracing.Trace("OrganizationServiceFault: {0}", ex.Message);
                throw new InvalidPluginExecutionException("Error in OverrideAccountNameFromLead plugin.", ex);
            }
            catch (Exception ex)
            {
                tracing.Trace("Unexpected Exception: {0}", ex.Message);
                throw new InvalidPluginExecutionException("Unexpected error in OverrideAccountNameFromLead plugin.", ex);
            }

            tracing.Trace("Plugin execution ended.");
        }
    }
}
