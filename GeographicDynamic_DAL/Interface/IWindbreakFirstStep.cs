using GeographicDynamic_DAL.DTOs.Windbreak;
using GeographicDynamicWebAPI.Wrappers;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace GeographicDynamic_DAL.Interface
{
    public interface IWindbreakFirstStep
    {
        public Result<bool> RenamePhotosInFolder(RenamePhotoDTO renamePhotoDTO);
        public Result<bool> ExcelCalculations(ExcelReadDTO excelReadDTO);
        public Result<DictionaryDTO> GetProjectNames();
        public Result<DictionaryDTO> GetEtapiID();
        public Result<DictionaryDTO> getSaxeobaList();
        public Result<bool> PhotoSplitKerdzoSaxelmwifo(string GadanomriliPhotoFolderPath, string DestinationFolderPath);
        public Result<VarjisFartiDTO> GetVarjisFartebi(int AreaNameID);
        public Result<bool> GetCheckPhotoDate(string folderPath, string resultPath);
    }
}
