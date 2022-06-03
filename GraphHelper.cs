using Azure.Core;
using Azure.Identity;
using Microsoft.Graph;

class GraphHelper
{
    // Settings object
    private static Settings? _settings;
    // User auth token credential
    private static DeviceCodeCredential? _deviceCodeCredential;
    // Client configured with user authentication
    private static GraphServiceClient? _userClient;

    // App-ony auth token credential
    private static ClientSecretCredential? _clientSecretCredential;
    // Client configured with app-only authentication
    private static GraphServiceClient? _appClient;

    private static void EnsureGraphForAppOnlyAuth()
    {
        // Ensure settings isn't null
        _ = _settings ??
            throw new System.NullReferenceException("Settings cannot be null");

        if (_clientSecretCredential == null)
        {
            _clientSecretCredential = new ClientSecretCredential(
                _settings.TenantId, _settings.ClientId, _settings.ClientSecret);
        }

        if (_appClient == null)
        {
            _appClient = new GraphServiceClient(_clientSecretCredential,
                // Use the default scope, which will request the scopes
                // configured on the app registration
                new[] { "https://graph.microsoft.com/.default" });
        }
    }
    private static void EnsureGraphForAppOnlyAuth(string[] scopes)
    {
        // Ensure settings isn't null
        _ = _settings ??
            throw new System.NullReferenceException("Settings cannot be null");

        if (_clientSecretCredential == null)
        {
            _clientSecretCredential = new ClientSecretCredential(
                _settings.TenantId, _settings.ClientId, _settings.ClientSecret);
        }

        if (_appClient == null)
        {
            _appClient = new GraphServiceClient(_clientSecretCredential,
                scopes);
        }
    }
    public static void InitializeGraphForAppOnlyAuth(Settings settings)
    {
        _settings = settings;
        EnsureGraphForAppOnlyAuth();
    }

    public static void InitializeGraphForUserAuth(Settings settings,
        Func<DeviceCodeInfo, CancellationToken, Task> deviceCodePrompt)
    {
        _settings = settings;

        _deviceCodeCredential = new DeviceCodeCredential(deviceCodePrompt,
            settings.AuthTenant, settings.ClientId);

        _userClient = new GraphServiceClient(_deviceCodeCredential, settings.GraphUserScopes);
    }

    public static async Task<string> GetUserTokenAsync()
    {
        // Ensure credential isn't null
        _ = _deviceCodeCredential ??
            throw new System.NullReferenceException("Graph has not been initialized for user auth");

        // Ensure scopes isn't null
        _ = _settings?.GraphUserScopes ?? throw new System.ArgumentNullException("Argument 'scopes' cannot be null");

        // Request token with given scopes
        var context = new TokenRequestContext(_settings.GraphUserScopes);
        var response = await _deviceCodeCredential.GetTokenAsync(context);
        return response.Token;
    }

    public static async Task<string> GetAppOnlyTokenAsync()
    {
        // Ensure credential isn't null
        _ = _clientSecretCredential ??
            throw new System.NullReferenceException("Graph has not been initialized for app-only auth");

        // Request token with given scopes
        var context = new TokenRequestContext(new[] { "https://graph.microsoft.com/.default" });
        var response = await _clientSecretCredential.GetTokenAsync(context);
        return response.Token;
    }

    public static Task<User> GetUserAsync()
    {
        // Ensure client isn't null
        _ = _userClient ??
            throw new System.NullReferenceException("Graph has not been initialized for user auth");

        return _userClient.Me
            .Request()
            .Select(u => new
            {
                // Only request specific properties
                u.DisplayName,
                u.Mail,
                u.UserPrincipalName
            })
            .GetAsync();
    }

    public static Task<IGraphServiceUsersCollectionPage> GetUsersAsync()
    {
        EnsureGraphForAppOnlyAuth();
        // Ensure client isn't null
        _ = _appClient ??
            throw new System.NullReferenceException("Graph has not been initialized for app-only auth");

        return _appClient.Users
            .Request()
            .Select(u => new
            {
                // Only request specific properties
                u.DisplayName,
                u.Id,
                u.Mail
            })
            // Get at most 25 results
            .Top(25)
            // Sort by display name
            .OrderBy("DisplayName")
            .GetAsync();
    }

    public async static Task<string> CreateNewTeamAsync()
    {
        EnsureGraphForAppOnlyAuth();
        //EnsureGraphForAppOnlyAuth(new[] { "https://graph.microsoft.com/Team.Create" });

        // Ensure client isn't null
        _ = _appClient ??
            throw new System.NullReferenceException("Graph has not been initialized for app-only auth");

        var team = new Team
        {
            DisplayName = "AutoGen-" + System.IO.Path.GetRandomFileName(),
            Description = "My Sample Team’s Description",
            Members = new TeamMembersCollectionPage()
            {
                new AadUserConversationMember
                {
                    Roles = new List<String>()
                    {
                        "owner"
                    },
                    AdditionalData = new Dictionary<string, object>()
                    {
                        {"user@odata.bind", "https://graph.microsoft.com/v1.0/users('4cb08dcb-b50e-4ee6-9712-03fd4c746a6c')"}
                    }
                }
            },
            AdditionalData = new Dictionary<string, object>()
            {
                {"template@odata.bind", "https://graph.microsoft.com/v1.0/teamsTemplates('standard')"}
            }
        };
        var result = await _appClient.Teams.Request().AddAsync(team);
        //Console.WriteLine("Created team: " + result.ToString());
        return team.DisplayName;
    }

    public async static Task CreateNewTeamBulkAsync()
    {
        EnsureGraphForAppOnlyAuth();
        //EnsureGraphForAppOnlyAuth(new[] { "https://graph.microsoft.com/Team.Create" });

        // Ensure client isn't null
        _ = _appClient ??
            throw new System.NullReferenceException("Graph has not been initialized for app-only auth");
        for (int i = 0; i < 500; i++)
        {
            var team = new Team
            {
                DisplayName = i + " -AutoGen-" + System.IO.Path.GetRandomFileName(),
                Description = "My Sample Team’s Description",
                Members = new TeamMembersCollectionPage()
            {
                new AadUserConversationMember
                {
                    Roles = new List<String>()
                    {
                        "owner"
                    },
                    AdditionalData = new Dictionary<string, object>()
                    {
                        {"user@odata.bind", "https://graph.microsoft.com/v1.0/users('19e76345-163a-4668-967f-5cbe4ca9f1b9')"}
                    }
                }
            },
                AdditionalData = new Dictionary<string, object>()
            {
                {"template@odata.bind", "https://graph.microsoft.com/v1.0/teamsTemplates('standard')"}
            }
            };
            await _appClient.Teams.Request().AddAsync(team);
            Console.WriteLine("Team index: " + i);
        }
    }
    public async static Task<IGraphServiceGroupsCollectionPage> ListGroupsWithoutTeamsAsync()
    {
        EnsureGraphForAppOnlyAuth();
        //EnsureGraphForAppOnlyAuth(new[] { "https://graph.microsoft.com/Group.Read.All" });

        // Ensure client isn't null
        _ = _appClient ??
            throw new System.NullReferenceException("Graph has not been initialized for app-only auth");

        var queryOptions = new List<QueryOption>()
        {
            new QueryOption("$count", "true")
        };
        var groups = await _appClient.Groups
        .Request(queryOptions)
        .Header("ConsistencyLevel", "eventual")
        .Filter("groupTypes/any(c:c eq 'Unified') and NOT(resourceProvisioningOptions/any(x:x eq 'Team'))")
        .Select("id, displayName") //, resourceProvisioningOptions")
        .GetAsync();

        return groups;
    }

    public async static Task<string> TeamifyGroupAsync(string groupId)
    {
        EnsureGraphForAppOnlyAuth();
        //EnsureGraphForAppOnlyAuth(new[] { "https://graph.microsoft.com/Group.Read.All" });

        // Ensure client isn't null
        _ = _appClient ??
            throw new System.NullReferenceException("Graph has not been initialized for app-only auth");

        var team = new Team
        {
            MemberSettings = new TeamMemberSettings
            {
                AllowCreatePrivateChannels = true,
                AllowCreateUpdateChannels = true
            },
            MessagingSettings = new TeamMessagingSettings
            {
                AllowUserEditMessages = true,
                AllowUserDeleteMessages = true
            },
            FunSettings = new TeamFunSettings
            {
                AllowGiphy = true,
                GiphyContentRating = GiphyRatingType.Strict
            }
        };

        var result = await _appClient.Groups[groupId].Team
            .Request()
            .PutAsync(team);
        return "Teamified group successfully: " + result.WebUrl;
        /*
        //Alternate way to teamify a group
        var team = new Team
        {
            AdditionalData = new Dictionary<string, object>()
            {
                {"template@odata.bind", "https://graph.microsoft.com/v1.0/teamsTemplates('standard')"},
                {"group@odata.bind", $"https://graph.microsoft.com/v1.0/groups('{groupId}')"}
            }
        };

        await _appClient.Teams
            .Request()
            .AddAsync(team);
        */
    }

    public async static Task<IGraphServiceGroupsCollectionPage> ListAllTeamsAsync()
    {
        EnsureGraphForAppOnlyAuth();
        //EnsureGraphForAppOnlyAuth(new[] { "https://graph.microsoft.com/Group.Read.All" });

        // Ensure client isn't null
        _ = _appClient ??
            throw new System.NullReferenceException("Graph has not been initialized for app-only auth");

        var queryOptions = new List<QueryOption>()
        {
            new QueryOption("$count", "true")
        };
        var groups = await _appClient.Groups
        .Request(queryOptions)
        .Header("ConsistencyLevel", "eventual")
        .Filter("resourceProvisioningOptions/any(x:x eq 'Team')")
        .Select("id, displayName") //, resourceProvisioningOptions")
        .GetAsync();

        return groups;
    }

    public static Task<IGraphServiceSitesCollectionPage> ListNonGroupSitesAsync()
    {
        EnsureGraphForAppOnlyAuth();
        // Ensure client isn't null
        _ = _appClient ??
            throw new System.NullReferenceException("Graph has not been initialized for app-only auth");

        return _appClient.Sites
                .Request()
                //.Filter("siteCollection/root ne null")
                .Select("siteCollection,webUrl")
                .GetAsync();
    }
}