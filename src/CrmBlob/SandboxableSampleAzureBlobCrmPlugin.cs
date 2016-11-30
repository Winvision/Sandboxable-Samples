using System;
using System.ServiceModel;
using Microsoft.Xrm.Sdk;
using Microsoft.Xrm.Sdk.Query;
using Newtonsoft.Json;
using Sandboxable.Microsoft.WindowsAzure.Storage.Auth;
using Sandboxable.Microsoft.WindowsAzure.Storage.Blob;

namespace Sandboxable.Samples.Crm.Blob
{
    /// <summary>
    /// A sample class to show that it's very easy to put files on Azure blob storage from a CRM plug-in.
    /// </summary>
    /// <remarks>
    /// Based on sample https://msdn.microsoft.com/en-us/library/gg328263.aspx
    /// </remarks>
    public class SandboxableSampleAzureBlobCrmPlugin : IPlugin
    {
        /// <summary>
        /// Holds the secure string.
        /// </summary>
        private readonly string secureString;

        /// <summary>
        /// Holds the sample folder name.
        /// </summary>
        private const string FolderName = "samplecrmfolder";

        /// <summary>
        /// Initializes a new instance of the <see cref="SandboxableSampleAzureBlobCrmPlugin"/> class with the specified unsecure string and secure string.
        /// </summary>
        /// <param name="unsecureString">Contains the unsecure configuration.</param>
        /// <param name="secureString">Contains the secure configuration</param>
        public SandboxableSampleAzureBlobCrmPlugin(string unsecureString, string secureString)
        {
            // Store the secure string so it's available on execution
            this.secureString = secureString;
        }

        public async void Execute(IServiceProvider serviceProvider)
        {
            // Extract the tracing service for use in debugging sandboxed plug-ins
            ITracingService tracingService = (ITracingService)serviceProvider.GetService(typeof(ITracingService));

            // Obtain the execution context from the service provider
            IPluginExecutionContext context = (IPluginExecutionContext)serviceProvider.GetService(typeof(IPluginExecutionContext));

            // Get an instance of the organization service
            IOrganizationServiceFactory serviceFactory = (IOrganizationServiceFactory)serviceProvider.GetService(typeof(IOrganizationServiceFactory));
            IOrganizationService organizationService = serviceFactory.CreateOrganizationService(context.UserId);

            // Get user name
            Entity userEntity = organizationService.Retrieve("systemuser", context.InitiatingUserId, new ColumnSet("fullname"));

            if (context.PreEntityImages.ContainsKey("Target"))
            {
                // Obtain the target entity from the pre image
                Entity entity = context.PreEntityImages["Target"];

                try
                {
                    // Parse the settings
                    PluginSettings pluginSettings = JsonConvert.DeserializeObject<PluginSettings>(this.secureString);

                    // Build credentials and create a blob client
                    StorageCredentials storageCredentials = new StorageCredentials(pluginSettings.AccountName, pluginSettings.Key);
                    Uri baseUri = new Uri($"https://{pluginSettings.AccountName.ToLowerInvariant()}.blob.core.windows.net");
                    CloudBlobClient blobClient = new CloudBlobClient(baseUri, storageCredentials);

                    // Make sure our root sample directory exists
                    CloudBlobContainer container = blobClient.GetContainerReference(FolderName.ToLowerInvariant());
                    await container.CreateIfNotExistsAsync();

                    // Make sure there is a container for this specific entity
                    CloudBlobDirectory entityDirectory = container.GetDirectoryReference(entity.LogicalName.ToLowerInvariant());
                    await entityDirectory.Container.CreateIfNotExistsAsync();
                    
                    string fullName = userEntity.GetAttributeValue<string>("fullname");

                    // Create a message to store at the queue
                    var messageData = new
                    {
                        context.UserId,
                        fullName,
                        context.MessageName,
                        entity.LogicalName,
                        entity.Id,
                        entity.Attributes
                    };

                    // Get a reference to the blob
                    CloudBlockBlob blob = entityDirectory.GetBlockBlobReference(entity.Id.ToString("N").ToUpperInvariant() + ".json");

                    // Set properties
                    blob.Properties.ContentType = "application/json";

                    // Add metadata
                    blob.Metadata["User_ID"] = context.UserId.ToString("B").ToLowerInvariant();
                    blob.Metadata["User_Fullname"] = fullName;
                    blob.Metadata["Deletion_Date"] = context.OperationCreatedOn.ToString("O");

                    // Upload the content
                    await blob.UploadTextAsync(JsonConvert.SerializeObject(messageData, Formatting.Indented));
                }
                catch (FaultException<OrganizationServiceFault> exception)
                {
                    throw new InvalidPluginExecutionException("An error occurred in Sandboxable Sample Azure Blob CRM Plugin", exception);
                }
                catch (Exception exception)
                {
                    tracingService.Trace("Sandboxable Sample Azure Blob CRM Plugin: {0}", exception.ToString());
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