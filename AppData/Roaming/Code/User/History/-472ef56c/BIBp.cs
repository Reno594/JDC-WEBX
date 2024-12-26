using Microsoft.AspNetCore.Mvc;
using MyWebApp.Models;

namespace MyWebApp.Controllers
{
    public class FormController : Controller
    {
        // Método para mostrar el formulario
        public IActionResult Index()
        {
            return View();
        }

        // Método para manejar la sumisión del formulario
        [HttpPost]
        public IActionResult Submit(FormModel model)
        {
            if (ModelState.IsValid)
            {
                // Lógica para manejar los datos del formulario
                TempData["SuccessMessage"] = "Formulario enviado correctamente!";
                return RedirectToAction("Index");
            }

            return View("Index", model);
        }
    }
}










