using Microsoft.SharePoint;
using Microsoft.SharePoint.Administration;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Runtime.InteropServices;
using System.Text;
using System.Web;
namespace ExportXLtoSPList.Providers.ExceptionManager
{
    public class SPErrorLogs
    {
        public static void LogExceptionToSPLog(string webpartname, string methodname, Exception ex)
        {
            string _errorTitle = "WebPart Name :" + webpartname + " - Method Name :" + methodname;
            SPDiagnosticsService diagSvc = SPDiagnosticsService.Local;
            diagSvc.WriteTrace(0,
                               new SPDiagnosticsCategory(_errorTitle,
                                                         TraceSeverity.Monitorable,
                                                         EventSeverity.Error),
                               TraceSeverity.Monitorable,
                               "An exception occurred: {0}",
                               new object[] { ex });
        }
    }
}
