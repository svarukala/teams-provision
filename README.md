# Overview
This sample code uses .NET 6 and Microsoft Graph SDK. It uses app-only auth (client-credential flow) to do the following:
1. Display app-only access token
2. Display user access token
3. List users (app-only)
4. Create New Team (app-only)
5. List groups w/o Team (app-only)
6. Teamify group (app-only)
7. List all Teams (app-only)
8. List sites (app-only) - TBF
9. Teamify site (app-only) - TBF
10. Create 1000 Teams (app-only)

Note: Option 8 and 9 are yet to be finished

# Register Azure AD App
1. Use the sample here https://docs.microsoft.com/en-us/graph/tutorials/dotnet?tabs=aad to create Azure AD App registration for a C# Console Application

# Minimal Steps to Awesome
1. Rename the appsettings.sample.json to appsettings.json. 
2. Replace the placeholders with your Azure AD App's app id, client secret, tenant id. For the teamOwnerUserId, provide the user id (guid) of the user who will be added as the team owner. You can get user id by navigating to https://graph.microsoft.com/v1.0/me/ in MS Graph Explorer (https://aka.ms/ge).

```json
{
    "settings": {
      "clientId": "aad-client-id",
      "clientSecret": "aad-client-secret",
      "tenantId": "tenant-id",
      "authTenant": "common",
      "graphUserScopes": [
        "user.read",
        "mail.read",
        "mail.send"
      ],
      "teamOwnerUserId": "user-id"
    }
}
```
3. To build the code: 
```
dotnet build
```
4. To run the code:
```
dotnet run
```
