// See https://aka.ms/new-console-template for more information
Console.WriteLine("Graph API web server\n");

var settings = Settings.LoadSettings();

//Console.WriteLine(settings.GraphUserScopes[3]);

// Initialize Graph
InitializeGraph(settings);

await GetUserInfo();
//await GetTeamsChannelMessages();
await GetOutlookEmails();

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

async Task GetUserInfo()
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

async Task GetTeamsChannelMessages()
{
    try
    {
        var user = await GraphHelper.GetTeamsChannelMessagesAsync();
        Console.WriteLine(user);
    }
    catch (Exception ex)
    {
        Console.WriteLine($"Error getting user: {ex.Message}");
    }
}

async Task GetOutlookEmails()
{
    try
    {
        var messagePage = await GraphHelper.GetOutlookEmailsAsync();

        // Output each message's details
        foreach (var message in messagePage.CurrentPage)
        {
            Console.WriteLine($"Message: {message.Subject ?? "NO SUBJECT"}");
            Console.WriteLine($"  From: {message.From?.EmailAddress?.Name}");
            Console.WriteLine($"  Status: {(message.IsRead!.Value ? "Read" : "Unread")}");
            Console.WriteLine($"  Received: {message.ReceivedDateTime?.ToLocalTime().ToString()}");
            Console.WriteLine($"  Content: {message.Body.Content ?? "NO CONTENT"}");
        }

        // If NextPageRequest is not null, there are more messages
        // available on the server
        // Access the next page like:
        // messagePage.NextPageRequest.GetAsync();
        var moreAvailable = messagePage.NextPageRequest != null;

        Console.WriteLine($"\nMore messages available? {moreAvailable}");
    }
    catch (Exception ex)
    {
        Console.WriteLine($"Error getting user's inbox: {ex.Message}");
    }

}