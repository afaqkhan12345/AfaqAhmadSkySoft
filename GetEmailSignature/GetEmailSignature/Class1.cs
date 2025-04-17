using System;
using System.ServiceModel;
using System.Xml;

using Microsoft.Xrm.Sdk;
using Microsoft.Xrm.Sdk.Query;

namespace GetEmailSignature
{
    public class Class1 : IPlugin
    {
        public void Execute(IServiceProvider serviceProvider)
        {
            // Extract the tracing service for debugging
            ITracingService tracingService = (ITracingService)serviceProvider.GetService(typeof(ITracingService));
            tracingService.Trace("Plugin execution started.");

            // Obtain the execution context from the service provider
            IPluginExecutionContext context = (IPluginExecutionContext)serviceProvider.GetService(typeof(IPluginExecutionContext));
            tracingService.Trace("Plugin context retrieved.");

            // Get the organization service for database operations
            IOrganizationServiceFactory serviceFactory = (IOrganizationServiceFactory)serviceProvider.GetService(typeof(IOrganizationServiceFactory));
            IOrganizationService service = serviceFactory.CreateOrganizationService(context.UserId);
            tracingService.Trace("Organization service initialized.");

            try
            {
                // Ensure the plugin is triggered by the correct action
                if (context.MessageName.ToLower() == "ssc_retrieveemailsignature")
                {
                    tracingService.Trace("Custom action 'ssc_RetrieveEmailSignature' triggered.");

                    // Retrieve the input parameter (SignatureId)
                    if (context.InputParameters.Contains("SignatureId") && context.InputParameters["SignatureId"] != null)
                    {
                        string signatureId = context.InputParameters["SignatureId"].ToString();
                        tracingService.Trace("SignatureId retrieved from input parameters: " + signatureId);

                        // Retrieve the email signature details from the database
                        Entity signature = RetrieveEmailSignature(service, signatureId, tracingService);

                        if (signature != null)
                        {
                            if (signature.Contains("body") && signature["body"] != null)
                            {
                                string emailSignatureBody = signature["body"].ToString();

                                // Trace the actual body content
                                tracingService.Trace("Email Signature Body: " + emailSignatureBody);

                                // Parse the XML/XSLT and extract the resolved content
                                string resolvedSignature = ExtractResolvedContent(emailSignatureBody);
                                tracingService.Trace("Resolved Signature: " + resolvedSignature);

                                // Set the output parameter (SignatureDetails)
                                context.OutputParameters["SignatureDetails"] = resolvedSignature;
                                tracingService.Trace("SignatureDetails output parameter set.");
                            }
                            else
                            {
                                tracingService.Trace("Body field is null or missing in the email signature.");
                                throw new InvalidPluginExecutionException("Body field is null or missing in the email signature.");
                            }
                        }
                        else
                        {
                            tracingService.Trace("No email signature found for SignatureId: " + signatureId);
                            throw new InvalidPluginExecutionException("Email signature not found.");
                        }
                    }
                    else
                    {
                        tracingService.Trace("SignatureId input parameter is missing.");
                        throw new InvalidPluginExecutionException("SignatureId input parameter is missing.");
                    }
                }
                else
                {
                    tracingService.Trace("Plugin triggered by incorrect message: " + context.MessageName);
                }
            }
            catch (Exception ex)
            {
                tracingService.Trace("Error in GetEmailSignature plugin: {0}", ex.ToString());
                throw new InvalidPluginExecutionException("An error occurred in the GetEmailSignature plugin.", ex);
            }
            finally
            {
                tracingService.Trace("Plugin execution completed.");
            }
        }

        private Entity RetrieveEmailSignature(IOrganizationService service, string signatureId, ITracingService tracingService)
        {
            tracingService.Trace("Retrieving email signature for SignatureId: " + signatureId);

            // Query to retrieve the email signature entity record
            QueryExpression query = new QueryExpression("emailsignature"); // Entity logical name
            query.ColumnSet = new ColumnSet("body"); // Fetch the correct field "body"
            query.Criteria.AddCondition("emailsignatureid", ConditionOperator.Equal, signatureId); // Primary key field

            tracingService.Trace("Querying email signature entity...");

            EntityCollection results = service.RetrieveMultiple(query);

            if (results.Entities.Count > 0)
            {
                tracingService.Trace("Email signature found.");
                return results.Entities[0];
            }
            else
            {
                tracingService.Trace("No email signature found for SignatureId: " + signatureId);
                return null;
            }
        }

        private string ExtractResolvedContent(string xmlContent)
        {
            try
            {
                // Load the XML content into an XmlDocument
                XmlDocument xmlDoc = new XmlDocument();
                xmlDoc.LoadXml(xmlContent);

                // Extract the resolved content from the CDATA section
                XmlNode cdataNode = xmlDoc.SelectSingleNode("//xsl:template/text()", GetNamespaceManager(xmlDoc));
                if (cdataNode != null)
                {
                    return cdataNode.Value;
                }
            }
            catch (Exception ex)
            {
                // Log the error and return the original content if parsing fails
                ITracingService tracingService = new MockTracingService();
                tracingService.Trace("Error parsing XML/XSLT: " + ex.ToString());
            }

            return xmlContent; // Return the original content if parsing fails
        }

        private XmlNamespaceManager GetNamespaceManager(XmlDocument xmlDoc)
        {
            XmlNamespaceManager nsmgr = new XmlNamespaceManager(xmlDoc.NameTable);
            nsmgr.AddNamespace("xsl", "http://www.w3.org/1999/XSL/Transform");
            return nsmgr;
        }
    }

    // Mock tracing service for use in the ExtractResolvedContent method
    public class MockTracingService : ITracingService
    {
        public void Trace(string format, params object[] args)
        {
            // Do nothing (mock implementation)
        }
    }
}