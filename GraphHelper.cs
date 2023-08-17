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

	public static void InitializeGraphForUserAuth(Settings settings,
		Func<DeviceCodeInfo, CancellationToken, Task> deviceCodePrompt)
	{
		_settings = settings;

		_deviceCodeCredential = new DeviceCodeCredential(deviceCodePrompt,
			settings.TenantId, settings.ClientId);

		_userClient = new GraphServiceClient(_deviceCodeCredential, settings.GraphUserScopes);
        
        //Console.WriteLine(_settings);
        //Console.WriteLine(_deviceCodeCredential);
        //Console.WriteLine(_userClient);
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

	public static Task<IChannelMessagesCollectionPage> GetTeamsChannelMessagesAsync()
	{
		// Ensure client isn't null
		_ = _userClient ??
			throw new System.NullReferenceException("Graph has not been initialized for user auth");

		return _userClient.Teams["87ff2845-8127-4baa-b050-87f9847559cb"].Channels["19:57720ec1cfba45f99354552d6d97df02@thread.skype"].Messages.Request().GetAsync();
	}

	public static Task<IMailFolderMessagesCollectionPage> GetOutlookEmailsAsync()
	{
		// Ensure client isn't null
		_ = _userClient ??
			throw new System.NullReferenceException("Graph has not been initialized for user auth");

		return _userClient.Me
			.MailFolders["Inbox"]
			.Messages
			.Request()
			.Select(m => new
			{
				// Only request specific properties
				m.From,
				m.IsRead,
				m.ReceivedDateTime,
				m.Subject,
				m.Body
			})
			// Get at most 25 results
			.Top(25)
			// Sort by received time, newest first
			.OrderBy("ReceivedDateTime DESC")
			.GetAsync();
	}
}