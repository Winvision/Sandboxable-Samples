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

        public void Execute(IServiceProvider serviceProvider)
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
            string fullName = userEntity.GetAttributeValue<string>("fullname");

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
                    container.CreateIfNotExists();

                    // Get a directory reference for this specific entity
                    CloudBlobDirectory entityDirectory = container.GetDirectoryReference(entity.LogicalName.ToLowerInvariant());
                    
                    // Get a reference to the blob
                    string fileName = entity.Id.ToString("N").ToUpperInvariant() + ".json";
                    CloudBlockBlob blob = entityDirectory.GetBlockBlobReference(fileName);

                    // Set properties
                    blob.Properties.ContentType = "application/json";

                    // Add metadata
                    blob.Metadata["userid"] = context.UserId.ToString("B").ToLowerInvariant();
                    blob.Metadata["userfullname"] = fullName;
                    blob.Metadata["deletiondate"] = context.OperationCreatedOn.ToString("O");

                    // Create the file content to store in the blob
                    var blobData = new
                    {
                        context.UserId,
                        FullName = fullName,
                        context.MessageName,
                        entity.LogicalName,
                        entity.Id,
                        entity.Attributes
                    };

                    // Upload the content
                    blob.UploadText(JsonConvert.SerializeObject(blobData, Formatting.Indented));
                }
                catch (FaultException<OrganizationServiceFault> exception)
                {
                    throw new InvalidPluginExecutionException("An error occurred in Sandboxable Sample Azure Blob CRM Plugin", exception);
                }
                catch (Exception exception)
                {
                    tracingService.Trace($"Sandboxable Sample Azure Blob CRM Plugin: {exception}");
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