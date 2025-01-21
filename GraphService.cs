using Azure.Identity;
using Microsoft.Extensions.Logging;
using Microsoft.Graph;
using Microsoft.Graph.Models;
using Microsoft.Identity.Client;
using System.Security.Cryptography.X509Certificates;

public class GraphService
{
    private readonly GraphServiceClient _graphClient;
    private readonly ILogger _logger;

    public GraphService(ILogger logger)
    {
        _logger = logger;

        var scopes = new[] { "https://graph.microsoft.com/.default" };

        // Values from app registration
        string tenantId = Environment.GetEnvironmentVariable("TenantId");
        string clientId = Environment.GetEnvironmentVariable("ClientId");
        string certificateThumbprint = Environment.GetEnvironmentVariable("CertificateThumbprint");

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

            return certificates.Count > 0 ? certificates[0] : null;
        }
    }

    // Example method to get a user by ID
    public async Task CleanUpDormantAccountsAsync()
    {
        List<string> admins = await GetGroupMembersAsync();
        List<string> usersToBeDeleted = await GetDormantAccounts(admins);
        await DeleteUsersInBatchAsync(usersToBeDeleted);
    }

    private async Task DeleteUsersInBatchAsync(List<string> users)
    {
        try
        {
            // Create the batch request content
            var batchRequestContent = new BatchRequestContentCollection(_graphClient);

            // Add delete requests to the batch
            foreach (var userId in users)
            {
                // Build the DELETE request for each user
                var request = _graphClient.Users[userId].ToDeleteRequestInformation();

                await batchRequestContent.AddBatchRequestStepAsync(request, Guid.NewGuid().ToString());
                _logger.LogInformation($"The user {userId} will be deleted");

                if (batchRequestContent.BatchRequestSteps.Count == 20)
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

    public async Task<List<string>> GetDormantAccounts(List<string> admins)
    {
        _logger.LogInformation("Search for all dormant accounts in the directory...");
        int skippedUserCount = 0;
        List<string> usersToDelete = new List<string>();

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
                        if (admins.Contains(user.Id!))
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
        return usersToDelete;
    }

    public async Task<List<string>> GetGroupMembersAsync()
    {
        string adminGroupId = Environment.GetEnvironmentVariable("AdminGroupId");
        List<string> users = new List<string>();

        try
        {
            // Get the members of the group
            var members = await _graphClient.Groups[adminGroupId].Members
                .GetAsync(requestConfiguration =>
                    {
                        requestConfiguration.QueryParameters.Top = 999;
                        requestConfiguration.QueryParameters.Select = new string[] { "id, DisplayName" };
                    });

            foreach (var user in members.Value)
            {
                users.Add(user.Id);
            }
        }
        catch (ServiceException ex)
        {
            _logger.LogError($"Error getting group members: {ex.Message}");
        }

        return users;
    }
}