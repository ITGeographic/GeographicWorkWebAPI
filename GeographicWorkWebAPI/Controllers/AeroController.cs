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
        private readonly IAeroWithObjectId _aeroWithObjectId;
        public AeroController(IAero aero,IAeroWithObjectId aeroWithObjectId) 
        {
            _aero = aero;
            _aeroWithObjectId = aeroWithObjectId;
        }

        [HttpPost("AeroGadageba")]
        public IActionResult ExcelisWakiTxvaAero()
        {
            var result = _aero.ExcelisWakiTxvaAero();
            if (result.Success) return Ok(result);
            return BadRequest(result);
        }
        [HttpPost("AeroGadagebaWithObjectId")]
        public IActionResult ExcelisWakiTxvaAeroWithObjectId()
        {
            var result = _aeroWithObjectId.ExcelisWakiTxvaAeroWithObjectId();
            if (result.Success) return Ok(result);
            return BadRequest(result);
        }
    }
}
