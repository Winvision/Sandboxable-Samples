using System;
using System.ServiceModel;
using Microsoft.Xrm.Sdk;
using Newtonsoft.Json;
using Sandboxable.Microsoft.WindowsAzure.Storage.Auth;
using Sandboxable.Microsoft.WindowsAzure.Storage.Queue;

namespace Sandboxable.Samples.Crm.Queue
{
    /// <summary>
    /// A sample class to show that it's very easy to put messages on an Azure queue from a CRM plug-in.
    /// </summary>
    /// <remarks>
    /// Based on sample https://msdn.microsoft.com/en-us/library/gg328263.aspx
    /// </remarks>
    public class SandboxableSampleAzureQueueCrmPlugin : IPlugin
    {
        /// <summary>
        /// Holds the secure string.
        /// </summary>
        private readonly string secureString;

        /// <summary>
        /// Holds the sample queue name.
        /// </summary>
        private const string QueueName = "samplecrmqueue";

        /// <summary>
        /// Initializes a new instance of the <see cref="SandboxableSampleAzureQueueCrmPlugin"/> class with the specified unsecure string and secure string.
        /// </summary>
        /// <param name="unsecureString">Contains the unsecure configuration.</param>
        /// <param name="secureString">Contains the secure configuration</param>
        public SandboxableSampleAzureQueueCrmPlugin(string unsecureString, string secureString)
        {
            // Store the secure string so it's available on execution
            this.secureString = secureString;
        }

        public void Execute(IServiceProvider serviceProvider)
        {
            // Extract the tracing service for use in debugging sandboxed plug-ins
            ITracingService tracingService = (ITracingService)serviceProvider.GetService(typeof(ITracingService));

            // Obtain the execution context from the service provider
            IPluginExecutionContext context = (IPluginExecutionContext)serviceProvider.GetService(typeof(IPluginExecutionContext));

            if (context.InputParameters.Contains("Target") && context.InputParameters["Target"] is Entity)
            {
                // Obtain the target entity from the input parameters
                Entity entity = (Entity)context.InputParameters["Target"];

                try
                {
                    // Parse the settings
                    PluginSettings pluginSettings = JsonConvert.DeserializeObject<PluginSettings>(this.secureString);

                    // Build credentials and create a queue client
                    StorageCredentials storageCredentials = new StorageCredentials(pluginSettings.AccountName, pluginSettings.Key);
                    Uri baseUri = new Uri($"https://{pluginSettings.AccountName.ToLowerInvariant()}.queue.core.windows.net");
                    CloudQueueClient queueClient = new CloudQueueClient(baseUri, storageCredentials);

                    // Create a message to store at the queue
                    var messageData = new
                    {
                        context.UserId,
                        context.MessageName,
                        entity.LogicalName,
                        entity.Id,
                        entity.Attributes
                    };
                    CloudQueueMessage queueMessage = new CloudQueueMessage(JsonConvert.SerializeObject(messageData));

                    // Get a reference to the queue, create the queue if it doesn't exist yet
                    CloudQueue queue = queueClient.GetQueueReference(QueueName.ToLowerInvariant());
                    queue.CreateIfNotExists();

                    // Add the message to the queue
                    queue.AddMessage(queueMessage);
                }
                catch (FaultException<OrganizationServiceFault> ex)
                {
                    throw new InvalidPluginExecutionException("An error occurred in Sandboxable Sample Azure Queue CRM Plugin", ex);
                }
                catch (Exception ex)
                {
                    tracingService.Trace("Sandboxable Sample Azure Queue CRM Plugin: {0}", ex.ToString());
                    throw;
                }
            }
        }

        /// <summary>
        /// The settings for this plugin.
        /// </summary>
        public class PluginSettings
        {
            /// <summary>
            /// Gets or sets the base64 encoded connection key.
            /// </summary>
            public string Key { get; set; }

            /// <summary>
            /// Gets or sets the Azure storage account name.
            /// </summary>
            public string AccountName { get; set; }
        }
    }
}
