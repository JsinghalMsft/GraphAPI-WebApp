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

}