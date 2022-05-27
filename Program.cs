Console.WriteLine("Teams Provision (Using .NET 6)\n");

var settings = Settings.LoadSettings();

// Initialize Graph
//InitializeGraph(settings);
InitializeGraphAppOnly(settings);

// Greet the user by name
//await GreetUserAsync();

int choice = -1;

while (choice != 0)
{
    Console.WriteLine("Please choose one of the following options:");
    Console.WriteLine("0. Exit");
    Console.WriteLine("1. Display app-only access token");
    Console.WriteLine("2. Display user access token");
    Console.WriteLine("3. List users (app-only)");
    Console.WriteLine("4. Create New Team (app-only)");
    Console.WriteLine("5. List groups w/o Team (app-only)");
    Console.WriteLine("6. Teamify group (app-only)");
    Console.WriteLine("7. List sites w/o group (app-only)");
    Console.WriteLine("8. Teamify site (app-only)");
    try
    {
        choice = int.Parse(Console.ReadLine() ?? string.Empty);
    }
    catch (System.FormatException)
    {
        // Set to invalid value
        choice = -1;
    }
    switch (choice)
    {
        case 0:
            // Exit the program
            Console.WriteLine("Goodbye...");
            break;
        case 1:
            // Display access token
            await DisplayAppOnlyAccessTokenAsync();
            break;
        case 2:
            // List emails from user's inbox
            await DisplayAccessTokenAsync();
            break;
        case 3:
            // List users
            await ListUsersAsync();
            break;
        case 4:
            // Run any Graph code
            await CreateNewTeamAsync();
            break;
        case 5:
            // Run any Graph code
            await ListGroupsWithoutTeamsAsync();
            break;            
        case 6:
            // Run any Graph code
            await TeamifyGroupAsync();
            break;            
        case 7:
            // Run any Graph code
            await ListNonGroupSitesAsync();
            break;              
        default:
            Console.WriteLine("Invalid choice! Please try again.");
            break;
    }
    /*
    switch(choice)
    {
        case 0:
            // Exit the program
            Console.WriteLine("Goodbye...");
            break;
        case 1:
            // Display access token
            await DisplayAccessTokenAsync();
            break;
        case 2:
            // List emails from user's inbox
            await ListInboxAsync();
            break;
        case 3:
            // Send an email message
            await SendMailAsync();
            break;
        case 4:
            // List users
            await ListUsersAsync();
            break;
        case 5:
            // Run any Graph code
            await MakeGraphCallAsync();
            break;
        default:
            Console.WriteLine("Invalid choice! Please try again.");
            break;
    }
    */
}

void InitializeGraphAppOnly(Settings settings)
{
    GraphHelper.InitializeGraphForAppOnlyAuth(settings);
}

void InitializeGraph(Settings settings)
{
    GraphHelper.InitializeGraphForUserAuth(settings,
        (info, cancel) =>
        {
            // Display the device code message to
            // the user. This tells them
            // where to go to sign in and provides the
            // code to use.
            Console.WriteLine(info.Message);
            return Task.FromResult(0);
        });
}

async Task GreetUserAsync()
{
    try
    {
        var user = await GraphHelper.GetUserAsync();
        Console.WriteLine($"Hello, {user?.DisplayName}!");
        // For Work/school accounts, email is in Mail property
        // Personal accounts, email is in UserPrincipalName
        Console.WriteLine($"Email: {user?.Mail ?? user?.UserPrincipalName ?? ""}");
    }
    catch (Exception ex)
    {
        Console.WriteLine($"Error getting user: {ex.Message}");
    }
}

async Task DisplayAccessTokenAsync()
{
    try
    {
        InitializeGraph(settings);
        var userToken = await GraphHelper.GetUserTokenAsync();
        Console.WriteLine($"User token: {userToken}");
    }
    catch (Exception ex)
    {
        Console.WriteLine($"Error getting user access token: {ex.Message}");
    }
}

async Task DisplayAppOnlyAccessTokenAsync()
{
    try
    {
        var appOnlyToken = await GraphHelper.GetAppOnlyTokenAsync();
        Console.WriteLine($"AppOnly token: {appOnlyToken}");
    }
    catch (Exception ex)
    {
        Console.WriteLine($"Error getting app-only access token: {ex.Message}");
    }
}

async Task ListInboxAsync()
{
    // TODO
}

async Task SendMailAsync()
{
    // TODO
}

async Task ListUsersAsync()
{
    try
    {
        var userPage = await GraphHelper.GetUsersAsync();

        // Output each users's details
        foreach (var user in userPage.CurrentPage)
        {
            Console.WriteLine($"User: {user.DisplayName ?? "NO NAME"}");
            Console.WriteLine($"  ID: {user.Id}");
            Console.WriteLine($"  Email: {user.Mail ?? "NO EMAIL"}");
        }

        // If NextPageRequest is not null, there are more users
        // available on the server
        // Access the next page like:
        // userPage.NextPageRequest.GetAsync();
        var moreAvailable = userPage.NextPageRequest != null;

        Console.WriteLine($"\nMore users available? {moreAvailable}");
    }
    catch (Exception ex)
    {
        Console.WriteLine($"Error getting users: {ex.Message}");
    }
}

async Task CreateNewTeamAsync()
{
    try
    {
        var teamName = await GraphHelper.CreateNewTeamAsync();
        Console.WriteLine($"Created new team: {teamName}");
    }
    catch (Exception ex)
    {
        Console.WriteLine($"Error creating team: {ex.Message}");
    }
}

async Task ListGroupsWithoutTeamsAsync()
{
    try
    {
        var groups = await GraphHelper.ListGroupsWithoutTeamsAsync();
        Console.WriteLine($"Total groups without teams: "+ groups.Count);
        
            
        foreach (var group in groups.CurrentPage)
        {
            Console.WriteLine($"Name: {group.DisplayName ?? "NO NAME"}, ID: {group.Id}");
        }

        // If NextPageRequest is not null, there are more users
        // available on the server
        // Access the next page like:
        // userPage.NextPageRequest.GetAsync();
        var moreAvailable = groups.NextPageRequest != null;

        Console.WriteLine($"\nMore groups available? {moreAvailable}");
    }
    catch (Exception ex)
    {
        Console.WriteLine($"Error getting groups: {ex.Message}");
    }
}

async Task TeamifyGroupAsync()
{
    try
    {
        Console.WriteLine("Enter group ID to teamify:");
        var groupId = Console.ReadLine();
        var result = await GraphHelper.TeamifyGroupAsync(groupId);
        Console.WriteLine(result);
    }
    catch (Exception ex)
    {
        Console.WriteLine($"Error getting groups: {ex.Message}");
    }
}

async Task ListNonGroupSitesAsync()
{
    try
    {
        var sites = await GraphHelper.ListNonGroupSitesAsync();

        // Output each users's details
        foreach (var site in sites.CurrentPage)
        {
            Console.WriteLine($"SiteUrl: {site.WebUrl ?? "NO NAME"}");
        }

        // If NextPageRequest is not null, there are more users
        // available on the server
        // Access the next page like:
        // userPage.NextPageRequest.GetAsync();
        var moreAvailable = sites.NextPageRequest != null;

        Console.WriteLine($"\nMore sites available? {moreAvailable}");
    }
    catch (Exception ex)
    {
        Console.WriteLine($"Error getting users: {ex.Message}");
    }
}