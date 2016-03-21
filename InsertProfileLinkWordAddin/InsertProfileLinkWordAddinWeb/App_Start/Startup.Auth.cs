using System.Configuration;
using System.IdentityModel.Tokens;
using Microsoft.Owin.Security.ActiveDirectory;
using Owin;

namespace InsertProfileLinkWordAddinWeb
{
	public partial class Startup
	{
		public void ConfigureAuth(IAppBuilder app)
		{
			app.UseWindowsAzureActiveDirectoryBearerAuthentication(
				new WindowsAzureActiveDirectoryBearerAuthenticationOptions
				{
					Tenant = ConfigurationManager.AppSettings["Tenant"],
					TokenValidationParameters = new TokenValidationParameters
					{
						ValidateIssuer = false,
						ValidAudience = ConfigurationManager.AppSettings["ClientId"]
					}
				});
		}

	}
}
