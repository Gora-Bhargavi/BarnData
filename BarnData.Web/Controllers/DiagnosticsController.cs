using Microsoft.AspNetCore.Mvc;

namespace BarnData.Web.Controllers
{
    // Hosts the Phase 1 smoke test at /Diagnostics/SmokeTest
    // Delete this controller + matching view once Phase 2 is in place.
    public class DiagnosticsController : Controller
    {
        [HttpGet]
        public IActionResult SmokeTest() => View();
    }
}
