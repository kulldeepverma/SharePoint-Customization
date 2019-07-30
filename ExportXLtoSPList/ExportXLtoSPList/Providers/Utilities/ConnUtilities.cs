using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace ExportXLtoSPList.Providers.Utilities
{
    class ConnUtilities
    {
        internal static string CreateXlConnectionString(string excelPath, string excelExten)
        {
            String connString = string.Empty;
            if (excelExten == "xls")
                connString = @"Provider=Microsoft.Jet.OLEDB.4.0;Data Source= " + excelPath + ";Extended Properties=\"Excel 8.0;HDR=YES;\"";
            else if (excelExten == "xlsx")
                connString = @"Provider=Microsoft.ACE.OLEDB.12.0;Data Source= " + excelPath + ";Extended Properties=\"Excel 12.0;HDR=YES;\"";
            return connString;
        }
    }
}
