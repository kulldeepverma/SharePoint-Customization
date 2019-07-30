using Microsoft.SharePoint;
using System;
using System.Collections;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using ExportXLtoSPList.Providers.Utilities;
using System.Data;

namespace ExportXLtoSPList.Providers.DataAccess
{
    class DataAccessProvider
    {
        internal static List<System.Collections.DictionaryEntry> getXlColnames(System.Data.OleDb.OleDbConnection xlConn, string sheetName,string excelPath, string excelExten)
        {
            System.Collections.Generic.List<DictionaryEntry> XlColumns;
            System.Data.DataSet ds = new System.Data.DataSet();
            ds = getXldata(xlConn, sheetName, excelPath, excelExten);
            XlColumns = new System.Collections.Generic.List<DictionaryEntry>();
            foreach (System.Data.DataColumn xlClnm in ds.Tables[0].Columns)
            {
                XlColumns.Add(new DictionaryEntry(xlClnm.Caption, xlClnm.Caption));
            }
            return XlColumns;
        }
        internal static System.Data.DataSet getXldata(System.Data.OleDb.OleDbConnection xlConn, string sheetName, string excelPath, string excelExten)
        {
            System.Data.DataSet ds = new System.Data.DataSet();
            SPSecurity.RunWithElevatedPrivileges(delegate()
            {
                xlConn = new global::System.Data.OleDb.OleDbConnection(Utilities.ConnUtilities.CreateXlConnectionString(excelPath, excelExten));
                string cmd = string.Format("select * from [{0}]", sheetName);
                System.Data.OleDb.OleDbCommand Comm = new System.Data.OleDb.OleDbCommand(cmd, xlConn);
                Comm.Connection = xlConn;
                System.Data.OleDb.OleDbDataAdapter adapter = new System.Data.OleDb.OleDbDataAdapter();
                xlConn.Open();
                Comm.CommandType = System.Data.CommandType.Text;
                adapter.SelectCommand = Comm;
                adapter.Fill(ds);
                xlConn.Close();
            });
            return ds;
        }
        internal static System.Collections.Generic.List<DictionaryEntry> getSPListColumns(string spListGuid)
        {
            SPList List = getSPList(spListGuid);
            System.Collections.Generic.List<DictionaryEntry> spColumns = new System.Collections.Generic.List<DictionaryEntry>();
            foreach (SPField field in List.Fields)
            {
                if (field.Hidden == false && field.ReadOnlyField == false)
                {
                    //if (field.Title != "Content Type" && field.Title != "Title" && field.Title != "Attachments")
                    if (field.Title != "Content Type" && field.Title != "Attachments")
                        spColumns.Add(new DictionaryEntry(field.Title, field.Title));
                }
            }
            return spColumns;
        }

        internal static SPList getSPList(string spListGuid)
        {
            SPSite Site = SPContext.Current.Site;
            SPList List = null;
            SPSecurity.RunWithElevatedPrivileges(delegate()
            {
                SPWeb web = Site.OpenWeb();
                web.AllowUnsafeUpdates = true;
                List = web.Lists.TryGetList(spListGuid);
            });
            return List;
        }
        internal static string[] GetExcelSheetNames(System.Data.OleDb.OleDbConnection xlConn)
        {
            xlConn.Open();
            DataTable dt = xlConn.GetOleDbSchemaTable(System.Data.OleDb.OleDbSchemaGuid.Tables, null);
            String[] excelSheets = new String[dt.Rows.Count];
            int i = 0;
            // Add the sheet name to the string array.
            foreach (System.Data.DataRow row in dt.Rows)
            {
                excelSheets[i] = row["TABLE_NAME"].ToString();
                i++;
            }
            xlConn.Close();
            return excelSheets;
        }
    }
}
