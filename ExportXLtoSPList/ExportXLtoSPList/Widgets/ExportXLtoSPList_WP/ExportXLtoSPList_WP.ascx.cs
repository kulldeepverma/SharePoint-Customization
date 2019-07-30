using System;
using System.Collections;
using System.ComponentModel;
using System.Web;
using System.Web.UI.WebControls;
using System.Web.UI.WebControls.WebParts;
using Microsoft.SharePoint;
using System.Web.UI;
using System.Collections.Generic;
using Microsoft.SharePoint.Administration;
using ExportXLtoSPList.Providers.ExceptionManager;
using ExportXLtoSPList.Providers.Utilities;
using ExportXLtoSPList.Providers.DataAccess;
namespace ExportXLtoSPList.Widgets.ExportXLtoSPList_WP
{
    [ToolboxItemAttribute(false)]
    public partial class ExportXLtoSPList_WP : WebPart
    {
        string oWebpartName = "ExportXLtoSPList_WP";
        static global::System.Data.OleDb.OleDbConnection xlConn;
        static String spListGuid;
        static String excelPath, excelExten;
        static System.Data.DataTable dt;
        string spColumnName, xlColumnName = string.Empty;

        public ExportXLtoSPList_WP()
        {
        }

        protected override void OnInit(EventArgs e)
        {
            base.OnInit(e);
            InitializeControl();
        }

        protected void Page_Load(object sender, EventArgs e)
        { 
            if (!Page.IsPostBack)
            {
                SPSecurity.RunWithElevatedPrivileges(delegate() {
                    SPWeb oWeb = SPContext.Current.Web;
                    SPList oList = oWeb.Lists.
                });
            }
            lblSuccessMessage.Text = string.Empty;
        }

        protected void btnUpload_Click(object sender, EventArgs e)
        {
            if (ddlSPList.SelectedValue != "0")
            {
                spListGuid = ddlSPList.SelectedValue.Trim().ToString();
                SPSecurity.RunWithElevatedPrivileges(delegate()
                {
                    try
                    {
                        CreateDirectory();
                        excelPath = HttpContext.Current.Server.MapPath("/_layouts/Temp/" + fuExcelSheet.FileName);
                        if (fuExcelSheet.FileName != string.Empty)
                        {
                            if (excelPath.EndsWith("xls", StringComparison.InvariantCultureIgnoreCase))
                            {
                                excelExten = "xls";
                                xlConn = new global::System.Data.OleDb.OleDbConnection(Providers.Utilities.ConnUtilities.CreateXlConnectionString(excelPath, excelExten));

                                ddlSheetName.DataSource = DataAccessProvider.GetExcelSheetNames(xlConn);
                                ddlSheetName.DataBind();

                                btnShowData.Visible = true;
                                ddlSheetName.Visible = true;
                                this.Page.ClientScript.RegisterStartupScript(GetType(), "key", "document.getElementById('light').style.display='block';", true);
                            }
                            else if (excelPath.EndsWith("xlsx", StringComparison.InvariantCultureIgnoreCase))
                            {
                                excelExten = "xlsx";
                                xlConn = new global::System.Data.OleDb.OleDbConnection(Providers.Utilities.ConnUtilities.CreateXlConnectionString(excelPath, excelExten));

                                ddlSheetName.DataSource = DataAccessProvider.GetExcelSheetNames(xlConn);
                                ddlSheetName.DataBind();

                                this.Page.ClientScript.RegisterStartupScript(GetType(), "key", "document.getElementById('light').style.display='block';", true);
                                litMessage.Text = string.Empty;
                                btnShowData.Visible = true;
                                ddlSheetName.Visible = true;
                            }
                            else
                            {
                                litMessage.Text = "Please Upload .xls or .xlsx file";
                            }
                        }
                        else
                        {
                            litMessage.Text = "Please provide a valid file path";
                        }
                    }
                    catch (Exception ex)
                    {
                        litMessage.Text = "Incorrect Format";
                        rptBindColumns.Visible = false;
                        SPErrorLogs.LogExceptionToSPLog(oWebpartName, "btnUpload_Click", ex);
                    }
                });
            }
        }
        protected void rptBindColumns_OnItemDataBound(object sender, RepeaterItemEventArgs e)
        {
            try
            {
                if (!(e.Item.ItemType == ListItemType.Item || e.Item.ItemType == ListItemType.AlternatingItem)) return;
                CheckBox chkSelection = (e.Item.FindControl("chkXlNames") as CheckBox);
                DropDownList drpSPList = (DropDownList)e.Item.FindControl("drpSPlistColumns");
                System.Collections.Generic.List<DictionaryEntry> XlColumns = DataAccessProvider.getXlColnames(xlConn, ddlSheetName.SelectedValue, excelPath, excelExten);
                drpSPList.DataSource = XlColumns;
                drpSPList.DataTextField = "Value";
                drpSPList.DataValueField = "Key";
                drpSPList.DataBind();
                if (XlColumns.Exists(o => o.Value.ToString().Trim().ToLower().Equals(chkSelection.Text.ToString().Trim().ToLower())))
                {
                    drpSPList.SelectedValue = chkSelection.Text;
                    chkSelection.Checked = true;
                }
            }
            catch (Exception ex)
            {
                SPErrorLogs.LogExceptionToSPLog(oWebpartName, "rptBindColumns_OnItemDataBound", ex);
            }
        }


        private void insertXlDataToSPList(String XlSPLists)
        {
            SPSecurity.RunWithElevatedPrivileges(delegate()
            {
                System.Data.DataSet dsXldata = new System.Data.DataSet();
                dsXldata = Providers.DataAccess.DataAccessProvider.getXldata(xlConn, ddlSheetName.SelectedValue.ToString(), excelPath, excelExten);
                String[] XlSpArray = XlSPLists.Split(';');
                SPList List = DataAccessProvider.getSPList(spListGuid);
                int xlrow = 0;
                bool isRowEmpty;
                foreach (System.Data.DataRow dr in dsXldata.Tables[0].Rows)
                {
                    xlrow += 1;
                    isRowEmpty = true;
                    try
                    {
                        SPListItem listitems = List.Items.Add();
                        foreach (string myString in XlSpArray)
                        {
                            string[] values = myString.Split('-');
                            if (listitems != null)
                            {
                                spColumnName = values[0];
                                xlColumnName = values[1];
                                if (dr[values[1]].ToString().Trim() != String.Empty)
                                {
                                    isRowEmpty = false;
                                    if (listitems.Fields[values[0]].GetType().Name == "SPFieldDateTime")
                                    {
                                        DateTime tempDateTime;
                                        if (DateTime.TryParse(dr[values[1]].ToString(), out tempDateTime))
                                            listitems[values[0]] = tempDateTime.ToShortDateString();
                                    }
                                    else
                                        listitems[values[0]] = dr[values[1]];  //values[1] --> Excelcolumn and values[0]---> SPfield name
                                }
                            }

                        }
                        if (!isRowEmpty) listitems.Update();
                    }
                    catch (Exception ex)
                    {
                        litValidationMessage.Text = " No. of Rows copied to SpList: " + (xlrow - 1) + "<br>" +
                                                    " Type mismatch between the columns:" + " " + spColumnName + " - " + xlColumnName + " at row " + (xlrow + 1);
                        this.Page.ClientScript.RegisterStartupScript(GetType(), "key", "document.getElementById('light').style.display='block';", true);
                        SPErrorLogs.LogExceptionToSPLog(oWebpartName, "insertXlDataToSPList", ex);
                    }
                }
            });
        }

        protected void btnImport_Click(object sender, EventArgs e)
        {
            try
            {
                SPSecurity.RunWithElevatedPrivileges(delegate()
                {
                    string message = string.Empty;
                    String XlSPLists = string.Empty;
                    litValidationMessage.Text = string.Empty;
                    litValidationErrorMessage.Text = string.Empty;
                    foreach (RepeaterItem item in this.rptBindColumns.Items)
                    {
                        DropDownList ddlSPlistColumn = item.FindControl("drpSPlistColumns") as DropDownList;
                        CheckBox chkXlColumn = item.FindControl("chkXlNames") as CheckBox;
                        if (chkXlColumn.Checked == true)
                        {

                            if (message.IndexOf(ddlSPlistColumn.SelectedValue) < 0)
                            {
                                message = message + ";" + ddlSPlistColumn.SelectedValue;
                                XlSPLists = XlSPLists + ";" + chkXlColumn.Text + "-" + ddlSPlistColumn.SelectedValue;
                            }
                            else
                            {
                                litValidationMessage.Text = "Repeated spreadsheet columns:" + " " + ddlSPlistColumn.SelectedValue;
                                this.Page.ClientScript.RegisterStartupScript(GetType(), "key", "document.getElementById('light').style.display='block';", true);
                                break;
                            }
                        }
                    }
                    if (litValidationMessage.Text == string.Empty && XlSPLists != String.Empty)
                    {
                        insertXlDataToSPList(XlSPLists.Substring(1));
                        DeleteAllFiles();
                        lblSuccessMessage.Visible = true;
                        lblSuccessMessage.Text = "Data Imported Successfully";
                        rptBindColumns.Visible = false;
                        ddlSheetName.Visible = false;
                        btnget.Visible = false;
                        litValidationMessage.Text = string.Empty;
                        litValidationErrorMessage.Text = string.Empty;
                    }
                    else
                    {
                        litValidationMessage.Text = (litValidationMessage.Text != string.Empty) ? litValidationMessage.Text : "Select atleast one column";
                        this.Page.ClientScript.RegisterStartupScript(GetType(), "key", "document.getElementById('light').style.display='block';", true);
                    }
                });
            }
            catch (Exception ex)
            {
                SPErrorLogs.LogExceptionToSPLog(oWebpartName, "btnImport_Click", ex);
            }
        }

        protected void btnShowData_Click(object sender, EventArgs e)
        {
            try
            {
                SPSecurity.RunWithElevatedPrivileges(delegate()
                {
                    xlConn = new global::System.Data.OleDb.OleDbConnection(ConnUtilities.CreateXlConnectionString(excelPath, excelExten));
                    rptBindColumns.DataSource = DataAccessProvider.getSPListColumns(ddlSPList.SelectedValue.ToString());
                    rptBindColumns.DataBind();
                    this.Page.ClientScript.RegisterStartupScript(GetType(), "key", "document.getElementById('light').style.display='block';", true);
                    btnget.Visible = true;
                    rptBindColumns.Visible = true;
                    litValidationMessage.Text = "";
                    litValidationMessage.Text = "";
                });
            }
            catch (Exception ex)
            {
                lblSuccessMessage.Text = "Selected worksheet is not correctly formatted. Please try again.";
                SPErrorLogs.LogExceptionToSPLog(oWebpartName, "btnShowData_Click", ex);
            }
        }

        private void CreateDirectory()
        {
            if (fuExcelSheet.HasFile)
                try
                {
                    SPSecurity.RunWithElevatedPrivileges(delegate()
                    {
                        String path = HttpContext.Current.Server.MapPath("/_layouts/Temp/");
                        if (System.IO.Directory.Exists(path))
                            UploadFiles(path);
                        else
                        {
                            System.IO.Directory.CreateDirectory(path);
                            UploadFiles(path);
                        }
                    });
                }
                catch (Exception ex)
                {
                    litMessage.Text = "ERROR: " + ex.Message.ToString();
                    SPErrorLogs.LogExceptionToSPLog(oWebpartName, "CreateDirectory", ex);
                }
        }
        private void UploadFiles(String Path)
        {
            SPSecurity.RunWithElevatedPrivileges(delegate()
            {
                fuExcelSheet.SaveAs(Path + "/" + fuExcelSheet.FileName);
            });
        }
        private void DeleteAllFiles()
        {
            SPSecurity.RunWithElevatedPrivileges(delegate()
            {
                string path = HttpContext.Current.Server.MapPath("/_layouts/Temp/");

                if (System.IO.Directory.Exists(path))
                {
                    System.IO.DirectoryInfo dir = new System.IO.DirectoryInfo(path);
                    System.IO.FileInfo[] files = dir.GetFiles();
                    if (files.Length > 0)
                    {
                        for (int i = 0; i < files.Length - 1; i++)
                        {
                            files[i].Delete();
                        }
                    }
                }
            });
        }
    }
}
