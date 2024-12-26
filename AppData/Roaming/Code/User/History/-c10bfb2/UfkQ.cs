using Microsoft.AspNetCore.Mvc;

namespace MyWebApp.Controllers
{
    public class CalculationsController : Controller
    {
        // Método corregido para que el modificador public esté en el lugar adecuado
        public IActionResult Calculate(string? s)
        {
            if (string.IsNullOrEmpty(s))
            {
                return BadRequest("Input cannot be null or empty.");
            }

            int result = int.Parse(s);
            // Tu lógica aquí
            return View();
        }
    }
}







