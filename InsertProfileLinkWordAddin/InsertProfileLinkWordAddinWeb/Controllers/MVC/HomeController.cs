using System.Web.Mvc;

namespace InsertProfileLinkWordAddinWeb.Controllers.MVC
{
	public class HomeController : Controller
	{
		[Route]
		public ActionResult Index()
		{
			return View();
		}
	}
}