using System;
using System.Collections.Generic;
using System.Threading.Tasks;
using System.Web.Http;
using InsertProfileLinkWordAddinWeb.Model;
using Microsoft.Office365.Discovery;

namespace InsertProfileLinkWordAddinWeb.Controllers.API
{
	[Authorize]
	[RoutePrefix("api/discovery")]
	public class DiscoveryController : ApiController
	{
		[Route("services")]
		[HttpPost]
		public async Task<IDictionary<string, CapabilityDiscoveryResult>> Services(AccessToken accessToken)
		{
			var discovery = new DiscoveryClient(new Uri("https://api.office.com/discovery/v1.0/me/"), () => accessToken.Token);
			return await discovery.DiscoverCapabilitiesAsync();
		}
	}
}
