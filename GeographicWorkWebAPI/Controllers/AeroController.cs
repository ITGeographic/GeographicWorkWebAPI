using GeographicDynamic_DAL.DTOs.Windbreak;
using GeographicDynamic_DAL.Interface;
using Microsoft.AspNetCore.Http;
using Microsoft.AspNetCore.Mvc;
using GeographicDynamic_DAL.Repository;

namespace GeographicWorkWebAPI.Controllers
{
    [ApiController]
    public class AeroController : Controller
    {
        private readonly IAero _aero;
        private readonly IAeroWithObjectId _aeroWithObjectId;
        private readonly IAeroWithoutFolders _aeroWithoutFolders;
        private readonly IAeroImagesFolderingRepository _aeroImagesFoldering;


        public AeroController(IAero aero,IAeroWithObjectId aeroWithObjectId, IAeroWithoutFolders aeroWithoutFolders, IAeroImagesFolderingRepository aeroImagesFolderingRepository) 
        {
            _aero = aero;
            _aeroWithObjectId = aeroWithObjectId;
            _aeroWithoutFolders = aeroWithoutFolders;
            _aeroImagesFoldering = aeroImagesFolderingRepository;
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
        [HttpPost("AeroGadagebaWithoutFolders")]
        public IActionResult ExcelisWakiTxvaAeroWithoutFolders()
        {
            var result = _aeroWithoutFolders.ExcelisWakiTxvaAeroWithoutFolders();
            if (result.Success) return Ok(result);
            return BadRequest(result);
        }

        [HttpPost("AeroGadagebisDafoldereba")]
        public IActionResult AeroImagesFolderingRepository()
        {
            // Call your repository instead of calling itself
            var result = _aeroImagesFoldering.DafolderebaPotoebis();

            if (result.Success)
                return Ok(result);

            return BadRequest(result);
        }
    }

}
