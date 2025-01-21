# Woodgrove Groceries clean up function

As users try the [Woodgrove Groceries](https://woodgrovedemo.com) live demo application, it’s possible that an excess number of accounts will be created in the Microsoft Entra External ID tenant over time. When users no longer access the demo application, these inactive (dormant) accounts will be deleted automatically from the directory. If your account has been deleted, you will have to sign-up again.

This Azure Function deletes dormant accounts. A dormant account is considered an account that has not signed in within the last 30 days, based on the `signInActivity/lastSignInDateTime` attribute. This function uses a `TimerTrigger`, which makes it easy to execute functions on a schedule. 

## How it works

1. The function [runs daily](https://learn.microsoft.com/en-us/azure/azure-functions/functions-bindings-timer) at 09:30 AM.
1. It uses [Microsoft Graph SDK](https://learn.microsoft.com/en-us/graph/sdks/sdks-overview) to query the directory with an [application registration](https://learn.microsoft.com/en-us/graph/tutorials/dotnet-app-only?tabs=aad&tutorial-step=1) that has the necessary permissions to [manage users](https://learn.microsoft.com/en-us/graph/api/user-delete). Authentication is performed using client credentials and a [private certificate](https://learn.microsoft.com/en-us/graph/sdks/choose-authentication-providers?tabs=csharp#using-a-client-certificate). 
1. To avoid accidentally deleting administrative accounts, it checks if the user is a member of a security group.
1. Then, it searches for dormant accounts and returns a list of users to be deleted.
1. Finally, it uses a batch operation to delete the users.

## Setup 

Application settings in this function app contain configuration options that affect the function app. These settings are accessed as environment variables.

- **AdminGroupID** - The security group ID containing admin accounts that must remain undeleted.
- **TenantId** - The Microsoft Entra external ID tenant ID.
- **ClientId** - The client ID.
- **CertificateThumbprint** - The certificate thumbprint is used for the client credentials flow. Additionally, you should upload the certificate to the Azure Function app and add the `WEBSITE_LOAD_CERTIFICATES` environment setting with the certificate thumbprint.


 
