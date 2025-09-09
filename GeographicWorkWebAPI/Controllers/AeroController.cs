using GeographicDynamic_DAL.DTOs.Windbreak;
using GeographicDynamic_DAL.Interface;
using Microsoft.AspNetCore.Http;
using Microsoft.AspNetCore.Mvc;

namespace GeographicWorkWebAPI.Controllers
{
    [ApiController]
    public class AeroController : Controller
    {
        private readonly IAero _aero;
        //aleks
        public AeroController(IAero aero) 
        {
            _aero = aero;
        }

        [HttpPost("AeroGadageba")]
        public IActionResult ExcelisWakiTxvaAero()
        {
            var result = _aero.ExcelisWakiTxvaAero();
            if (result.Success) return Ok(result);
            return BadRequest(result);
        }
    }
}
