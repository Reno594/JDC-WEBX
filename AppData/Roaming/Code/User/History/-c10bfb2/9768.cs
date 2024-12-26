using Microsoft.AspNetCore.Mvc;

namespace MyWebApp.Controllers
{
    public class CalculationsController : Controller
    {
        // Método para manejar cálculos
        public IActionResult Calculate(string? s)
        {
            if (string.IsNullOrEmpty(s))
            {
                return BadRequest("Input cannot be null or empty.");
            }

            if (int.TryParse(s, out int result))
            {
                int calculationResult = result * 2; // Ejemplo de cálculo
                ViewBag.CalculationResult = calculationResult;
                return View();
            }
            else
            {
                return BadRequest("Invalid input. Please enter a valid number.");
            }
        }
    }
}










