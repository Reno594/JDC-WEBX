using Microsoft.AspNetCore.Mvc;
using MyWebApp.Models;

namespace MyWebApp.Controllers
{
    public class FormController : Controller
    {
        public IActionResult Index()
        {
            return View();
        }

        [HttpPost]
        public IActionResult Submit(FormModel model)
        {
            if (ModelState.IsValid)
            {
                TempData["SuccessMessage"] = "Formulario enviado correctamente!";
                return RedirectToAction("Index");
            }

            return View("Index", model);
        }
    }
}











