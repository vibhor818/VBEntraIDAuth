using Microsoft.AspNetCore.Mvc;
using Microsoft.Graph;

namespace VBEntraIDAuth.Controllers
{
    public class ProfileController : Controller
    {
        private readonly GraphServiceClient _graphClient;

        public ProfileController(GraphServiceClient graphClient)
        {
            _graphClient = graphClient;
        }

        public async Task<IActionResult> Index()
        {
            if (User?.Identity?.IsAuthenticated == true)
            {
                // Fetch user details from Microsoft Graph
                var me = await _graphClient.Me
                .Request()
                .Select("displayName,streetAddress,city,state")
                .GetAsync();

                ViewBag.Name = me.DisplayName;
                ViewBag.Address = $"{me.StreetAddress}, {me.City}, {me.State}, {me.PostalCode}, {me.Country}";

                return View();
            }
            return View();
        }
    }

}
