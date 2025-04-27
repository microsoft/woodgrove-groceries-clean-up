using Azure.Identity;
using Microsoft.ApplicationInsights;
using Microsoft.ApplicationInsights.DataContracts;
using Microsoft.Extensions.Logging;
using Microsoft.Graph;
using Microsoft.Graph.Models;
using Microsoft.Identity.Client;
using System.Security.Cryptography.X509Certificates;

public class GraphService
{
    private readonly GraphServiceClient? _graphClient;
    private readonly ILogger _logger;
    private TelemetryClient _telemetryClient;

    public GraphService(ILogger logger, TelemetryClient tc)
    {
        _logger = logger;
        _telemetryClient = tc;

        var scopes = new[] { "https://graph.microsoft.com/.default" };

        // Values from app registration
        string tenantId = Environment.GetEnvironmentVariable("TenantId") ?? throw new InvalidOperationException("TenantId environment variable is not set.");
        string clientId = Environment.GetEnvironmentVariable("ClientId") ?? throw new InvalidOperationException("ClientId environment variable is not set.");
        string certificateThumbprint = Environment.GetEnvironmentVariable("CertificateThumbprint") ?? throw new InvalidOperationException("CertificateThumbprint environment variable is not set.");

        // Load the certificate from the certificate store
        var certificate = LoadCertificateFromStore(certificateThumbprint);
        if (certificate == null)
        {
            _logger.LogError("Certificate not found.");
            return;
        }

        // using Azure.Identity;
        var options = new ClientCertificateCredentialOptions
        {
            AuthorityHost = AzureAuthorityHosts.AzurePublicCloud,
        };

        // Create a ClientCertificateCredential
        var clientCertificateCredential = new ClientCertificateCredential(tenantId, clientId, certificate);

        // Create a GraphServiceClient
        _graphClient = new GraphServiceClient(clientCertificateCredential, scopes);
    }


    static X509Certificate2 LoadCertificateFromStore(string thumbprint)
    {
        if (string.IsNullOrWhiteSpace(thumbprint))
        {
            throw new ArgumentException("CertificateThumbprint environment variable should not be empty.");
        }

        using (var store = new X509Store(StoreLocation.CurrentUser)) // Or StoreLocation.LocalMachine
        {
            store.Open(OpenFlags.ReadOnly);
            var certificates = store.Certificates.Find(
                X509FindType.FindByThumbprint,
                thumbprint,
                validOnly: false // Set to true if you only want valid certificates
            );

            if (certificates.Count > 0)
            {
                return certificates[0];
            }

            throw new InvalidOperationException("No certificate found with the specified thumbprint.");
        }
    }

    // Example method to get a user by ID
    public async Task CleanUpDormantAccountsAsync()
    {
        using (_telemetryClient.StartOperation<RequestTelemetry>("Clean up operation"))
        {

            List<string> protectedAccounts = new List<string>();

            // Get the admin security group ID from the environment variable
            var adminGroupId = Environment.GetEnvironmentVariable("AdminGroupId");
            if (!string.IsNullOrEmpty(adminGroupId))
            {
                protectedAccounts = await GetGroupMembersAsync(protectedAccounts, adminGroupId);
            }

            // Get the exclusive demos group ID from the environment variable
            var exclusiveDemosGroupId = Environment.GetEnvironmentVariable("ExclusiveDemosGroupId");
            if (!string.IsNullOrEmpty(exclusiveDemosGroupId))
            {
                protectedAccounts = await GetGroupMembersAsync(protectedAccounts, exclusiveDemosGroupId);
            }

            // Get the users to be deleted
            List<string> usersToBeDeleted = await GetDormantAccounts(protectedAccounts);

            // Delete users in batches 
            await DeleteUsersInBatchAsync(usersToBeDeleted);
        }
    }

    private async Task DeleteUsersInBatchAsync(List<string> users)
    {

        // Check if the GraphServiceClient is initialized
        if (_graphClient == null)
        {
            _logger.LogError("GraphServiceClient is not initialized.");
            return;
        }

        try
        {
            // Create the batch request content
            var batchRequestContent = new BatchRequestContentCollection(_graphClient);

            // Add delete requests to the batch
            for (int i = 0; i < users.Count; i++)
            {
                string userId = users[i];

                // Build the DELETE request for each user
                var request = _graphClient.Users[userId].ToDeleteRequestInformation();

                await batchRequestContent.AddBatchRequestStepAsync(request, Guid.NewGuid().ToString());
                _logger.LogInformation($"The user {userId} will be deleted");

                Dictionary<string, string> prop = new Dictionary<string, string>();
                prop.Add("UserId", userId);
                _telemetryClient.TrackEvent("User deleted", prop, null);

                if (batchRequestContent.BatchRequestSteps.Count == 20 || i == (users.Count - 1))
                {
                    try
                    {
                        var returnedResponse = await _graphClient.Batch.PostAsync(batchRequestContent);

                        // var events = await returnedResponse
                        //     .GetResponseByIdAsync<EventCollectionResponse>(eventsRequestId);
                        // _logger.LogInformation($"{events.Value?.Count} Users have been deleted.");

                        // Initiate the collection
                        Thread.Sleep(3000);
                        batchRequestContent = new BatchRequestContentCollection(_graphClient);
                    }
                    catch (Exception ex)
                    {
                        _logger.LogError($"Batch operation failed: {ex.Message}");
                    }
                }
            }

        }
        catch (ServiceException ex)
        {
            _logger.LogError($"Error deleting users: {ex.Message}");
        }

    }

    public async Task<List<string>> GetDormantAccounts(List<string> protectedAccounts)
    {
        _logger.LogInformation("Search for all dormant accounts in the directory...");
        int skippedUserCount = 0;
        List<string> usersToDelete = new List<string>();

        // Check if the GraphServiceClient is initialized
        if (_graphClient == null)
        {
            _logger.LogError("GraphServiceClient is not initialized.");
            return usersToDelete;
        }

        // Format the cutoff date to ISO 8601 format
        string formattedDate = DateTime.UtcNow.AddDays(-30).ToString("yyyy-MM-ddTHH:mm:ssZ");

        try
        {
            // Get all users
            var users = await _graphClient.Users
            .GetAsync(requestConfiguration =>
                    {
                        requestConfiguration.QueryParameters.Select = new string[] { "id, DisplayName" };
                        requestConfiguration.QueryParameters.Filter = $"signInActivity/lastSignInDateTime le {formattedDate}";
                    });

            // Iterate over all the users in the directory
            var pageIterator = PageIterator<User, UserCollectionResponse>
                .CreatePageIterator(
                    _graphClient,
                    users!,
                    // Callback executed for each user in the collection
                    (user) =>
                    {
                        // Delete only test users
                        if (protectedAccounts.Contains(user.Id!))
                        {
                            //_logger.LogInformation("**** Skipping " + user.Id);
                            skippedUserCount++;
                        }
                        else
                        {
                            //_logger.LogInformation("Deleting " + user.Id);
                            usersToDelete.Add(user.Id!);
                        }

                        return true;
                    },

                    (req) =>
                    {
                        // Used to configure subsequent page requests
                        _logger.LogInformation($"{usersToDelete.Count} users will be deleted. {skippedUserCount} users will be skipped");
                        _logger.LogInformation($"Waiting and reading next page of users...");
                        Thread.Sleep(3000);
                        return req;
                    }
                );

            await pageIterator.IterateAsync();
        }
        catch (Exception ex)
        {
            _logger.LogError(ex.Message);
        }

        _logger.LogInformation($"The search for dormant accounts has been completed. {usersToDelete.Count} users will be deleted. {skippedUserCount} users will be skipped");
        Dictionary<string, double> metrics = new Dictionary<string, double>();
        metrics.Add("Delete", usersToDelete.Count);
        metrics.Add("Skip", skippedUserCount);
        _telemetryClient.TrackEvent("Search completed", null, metrics);
        return usersToDelete;
    }

    public async Task<List<string>> GetGroupMembersAsync(List<string> protectedAccounts, string groupID)
    {
        // Check if the GraphServiceClient is initialized
        if (_graphClient == null)
        {
            _logger.LogError("GraphServiceClient is not initialized.");
            return protectedAccounts;
        }

        int membersCount = 0;

        try
        {
            // Get the members of the group
            var members = await _graphClient.Groups[groupID].Members
                .GetAsync(requestConfiguration =>
                    {
                        requestConfiguration.QueryParameters.Top = 999;
                        requestConfiguration.QueryParameters.Select = new string[] { "id, DisplayName" };
                    });

            // Check if the group has members and iterate through them
            if (members != null && members.Value != null)
            {
                foreach (var user in members.Value)
                {
                    // Check if the user is already in the protected accounts list
                    if (user.Id != null && !protectedAccounts.Contains(user.Id))
                    {
                        // Add the user ID to the protected accounts list
                        protectedAccounts.Add(user.Id);
                        membersCount++;
                    }
                }
            }
        }
        catch (ServiceException ex)
        {
            _logger.LogError($"Error getting group members: {ex.Message}");
        }

        _logger.LogInformation($"The security group '{groupID}' has {membersCount} distinct protected accounts.");

        return protectedAccounts;
    }
}