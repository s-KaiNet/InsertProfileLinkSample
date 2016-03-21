using System.Web.Http;
using Newtonsoft.Json.Serialization;
using Owin;

namespace InsertProfileLinkWordAddinWeb
{
	public static class WebApiConfig
	{
		public static void Register(IAppBuilder app)
		{
			var config = new HttpConfiguration();
			config.Formatters.JsonFormatter.SerializerSettings.ContractResolver = new CamelCasePropertyNamesContractResolver();
			config.MapHttpAttributeRoutes();

			app.UseWebApi(config);
		}
	}
}
