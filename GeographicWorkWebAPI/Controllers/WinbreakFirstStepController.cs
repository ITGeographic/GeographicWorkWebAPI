using GeographicDynamic_DAL.DTOs.Windbreak;
using GeographicDynamic_DAL.Interface;
using GeographicDynamic_DAL.Models;
using Microsoft.AspNetCore.Http;
using Microsoft.AspNetCore.Mvc;

namespace GeographicWorkWebAPI.Controllers
{
    
    [ApiController]
    public class WinbreakFirstStepController : ControllerBase
    {


        /*test*/
        private readonly IWindbreakFirstStep _windbreak;

        public WinbreakFirstStepController(IWindbreakFirstStep windbreak)
        {
            _windbreak = windbreak;
        }
        [HttpPost("RenamePhotosInFolderFirstStep")]
        public IActionResult RenamePhotosInFolder(RenamePhotoDTO renamePhotoDTO)
        {
            var result = _windbreak.RenamePhotosInFolder(renamePhotoDTO);
            if (result.Success) return Ok(result);
            return BadRequest(result);
        }

        [HttpPost("upload-access")]
        public async Task<IActionResult> UploadAccessFile([FromForm] IFormFile file)
        {
            if (file == null || file.Length == 0)
                return BadRequest("No file provided");

            string saveFolder = @"\\server\Programmers\Alex_Script\Windbreak\Calculations\uploads";

            if (!Directory.Exists(saveFolder))
                Directory.CreateDirectory(saveFolder);

            string savedPath = Path.Combine(saveFolder, file.FileName);

            using (var stream = new FileStream(savedPath, FileMode.Create))
            {
                await file.CopyToAsync(stream);
            }

            return Ok(new { savedPath });
        }
        [HttpPost("upload-excel")]
        public async Task<IActionResult> UploadExcelFile([FromForm] IFormFile file)
        {
            if (file == null || file.Length == 0)
                return BadRequest("No file provided");

            string saveFolder = @"\\server\Programmers\Alex_Script\Windbreak\Calculations\uploads";

            if (!Directory.Exists(saveFolder))
                Directory.CreateDirectory(saveFolder);

            string savedPath = Path.Combine(saveFolder, file.FileName);

            using (var stream = new FileStream(savedPath, FileMode.Create))
            {
                await file.CopyToAsync(stream);
            }

            return Ok(new { savedPath });
        }

        [HttpPost("ExcelCalculationsFirstStep")]
        public IActionResult ExcelCalculations(ExcelReadDTO excelReadDTO)
        {
            var result = _windbreak.ExcelCalculations(excelReadDTO);
            if (result.Success) return Ok(result);
            return BadRequest(result);
        }

        [HttpGet("GetProjectNamesListFirstStep")]
        public IActionResult GetProjectNamesList()
        {
            var result = _windbreak.GetProjectNames();
            if (result.Success) return Ok(result);
            return BadRequest(result);
        }

        [HttpGet("GetEtapiIDListFirstStep")]
        public IActionResult GetEtapiIDList()
        {
            var result = _windbreak.GetEtapiID();
            if (result.Success) return Ok(result);
            return BadRequest(result);
        }

        [HttpPost("PostPhotoSplitPathsFirstStep")]
        public IActionResult PostPhotoSplitPaths(PhotoSplitDTO photoSplitDTO)
        {
            var result = _windbreak.PhotoSplitKerdzoSaxelmwifo(photoSplitDTO.GadanomriliPhotoFolderPath, photoSplitDTO.DestinationFolderPath);
            if (result.Success) return Ok(result);
            return BadRequest(result);
        }

        [HttpGet("getSaxeobaListFirstStep")]
        public IActionResult getSaxeobaList()
        {
            var result = _windbreak.getSaxeobaList();
            if (result.Success) return Ok(result);
            return BadRequest(result);
        }
        [HttpGet("GetVarjisFartebiListFirstStep")]
        public IActionResult GetVarjisFartebiList(int AreaNameID)
        {
            var result = _windbreak.GetVarjisFartebi(AreaNameID);
            if (result.Success) return Ok(result);
            return BadRequest(result);
        }

        [HttpPost("GetCheckPhotoDateFirstStep")]
        public IActionResult GetCheckPhotoDate(CheckPhotoDateDTO checkPhotoDateDTO)
        {
            var result = _windbreak.GetCheckPhotoDate(checkPhotoDateDTO.folderPath, checkPhotoDateDTO.resultPath);
            if (result.Success) return Ok(result);
            return BadRequest(result);
        }

        [HttpPost("FotoebisGayofaFirstStep")]
        public IActionResult FotoebisGayofa()
        {
            //var result = _windbreak.GetCheckPhotoDate(checkPhotoDateDTO.folderPath, checkPhotoDateDTO.resultPath);
            //if (result.Success) return Ok(result);
            GeographicDynamicDbContext windBreakContext = new GeographicDynamicDbContext();
            try
            {

                GadanomriliFotoebi photo = new GadanomriliFotoebi();

                var directories = Directory.GetDirectories(@"D:\\Documents\\Desktop\\I_etapi\\Photoes").OrderBy(filePath => Convert.ToInt32(Path.GetFileNameWithoutExtension(filePath)));

                foreach (var folderPath in directories)
                {
                    var idxLiter = folderPath.LastIndexOf('\\');
                    string literIDstr = folderPath.Substring(idxLiter + 1);

                    double literID = Convert.ToDouble(literIDstr);

                    var directories1 = Directory.GetDirectories(folderPath).OrderBy(filePath => Convert.ToInt32(Path.GetFileNameWithoutExtension(filePath)));

                    var list = directories1.OrderBy(filePath => Convert.ToInt32(Path.GetFileNameWithoutExtension(filePath)));


                    foreach (var item in list)
                    {
                        DirectoryInfo d5 = new DirectoryInfo(item);
                        FileInfo[] infos1 = d5.GetFiles();

                        var idxUniqid = item.LastIndexOf('\\');

                        string uniqIDstr = item.Substring(idxUniqid + 1);

                        double uniqID = Convert.ToDouble(uniqIDstr);

                        string photoN = "";

                        var PhotoDate = "";

                        //photo.UniqId = uniqID;
                        bool ismoved = true;
                        foreach (FileInfo f6 in infos1)
                        {
                            if (!f6.Name.Contains(".db"))
                            {

                                QarsafariGrouped? qarsafaritest = windBreakContext.QarsafariGroupeds.FirstOrDefault(m => m.UniqId == uniqID);


                                if (qarsafaritest?.Sakutreba == "კერძო" || qarsafaritest?.Sakutreba == "იურიდიული პირი")
                                {
                                    if (ismoved)
                                    {

                                        photo.LiterId = literID;
                                        string destinationFolder = Path.Combine((string.Concat(@"D:\Documents\Desktop\I_etapi\Split" + "\\" + "photoSplit" + "\\" + "Kerdzo")), literID.ToString());
                                        if (!Directory.Exists(destinationFolder))
                                        {
                                            Directory.CreateDirectory(destinationFolder);
                                        }
                                        string destinationFile = Path.Combine(destinationFolder, uniqID.ToString());
                                        //File.Copy(item, destinationFile);
                                        Microsoft.VisualBasic.FileIO.FileSystem.CopyDirectory(item, destinationFile);

                                        ismoved = false;
                                    }

                                }
                                if (qarsafaritest?.Sakutreba != "კერძო" && qarsafaritest?.Sakutreba != "იურიდიული პირი")
                                {
                                    if (ismoved)
                                    {
                                        photo.LiterId = literID;


                                        string destinationFolder = Path.Combine((string.Concat(@"D:\Documents\Desktop\I_etapi\Split" + "\\" + "photoSplit" + "\\" + "Saxelmwifo")), literID.ToString());

                                        if (!Directory.Exists(destinationFolder))
                                        {
                                            Directory.CreateDirectory(destinationFolder);
                                        }
                                        string destinationFile = Path.Combine(destinationFolder, uniqID.ToString());
                                        //File.Copy(item, destinationFile);
                                        Microsoft.VisualBasic.FileIO.FileSystem.CopyDirectory(item, destinationFile);
                                        ismoved = false;
                                    }
                                }
                            }
                        }
                    }

                }

            }

            catch (Exception ex)
            {
                return BadRequest("false");
            }


            return BadRequest("false");
        }
    }
}
