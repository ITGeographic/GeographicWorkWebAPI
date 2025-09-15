using GeographicDynamic_DAL.Repository;
using GeographicDynamicWebAPI.Wrappers;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace GeographicDynamic_DAL.Interface
{
    public interface IAero
    {
        Result<List<AeroRecord>> ExcelisWakiTxvaAero();
    }
}
