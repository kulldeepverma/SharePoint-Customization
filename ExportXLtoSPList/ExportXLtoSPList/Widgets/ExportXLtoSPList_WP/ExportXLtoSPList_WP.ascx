<%@ Assembly Name="$SharePoint.Project.AssemblyFullName$" %>
<%@ Assembly Name="Microsoft.Web.CommandUI, Version=15.0.0.0, Culture=neutral, PublicKeyToken=71e9bce111e9429c" %>
<%@ Register TagPrefix="SharePoint" Namespace="Microsoft.SharePoint.WebControls" Assembly="Microsoft.SharePoint, Version=15.0.0.0, Culture=neutral, PublicKeyToken=71e9bce111e9429c" %>
<%@ Register TagPrefix="Utilities" Namespace="Microsoft.SharePoint.Utilities" Assembly="Microsoft.SharePoint, Version=15.0.0.0, Culture=neutral, PublicKeyToken=71e9bce111e9429c" %>
<%@ Register TagPrefix="asp" Namespace="System.Web.UI" Assembly="System.Web.Extensions, Version=4.0.0.0, Culture=neutral, PublicKeyToken=31bf3856ad364e35" %>
<%@ Import Namespace="Microsoft.SharePoint" %>
<%@ Register TagPrefix="WebPartPages" Namespace="Microsoft.SharePoint.WebPartPages" Assembly="Microsoft.SharePoint, Version=15.0.0.0, Culture=neutral, PublicKeyToken=71e9bce111e9429c" %>
<%@ Control Language="C#" AutoEventWireup="true" CodeBehind="ExportXLtoSPList_WP.ascx.cs" Inherits="ExportXLtoSPList.Widgets.ExportXLtoSPList_WP.ExportXLtoSPList_WP" %>

<meta name="viewport" content="width=device-width, initial-scale=1">
<link rel="stylesheet" href="//maxcdn.bootstrapcdn.com/bootstrap/3.3.7/css/bootstrap.min.css">
<link href="../../_layouts/15/ExportXLtoSPList/css/main.css" rel="stylesheet" />
<script src="//ajax.googleapis.com/ajax/libs/jquery/3.2.1/jquery.min.js"></script>
<script src="//maxcdn.bootstrapcdn.com/bootstrap/3.3.7/js/bootstrap.min.js"></script>
<script type="text/javascript">
    function hide() {
        document.getElementById('light').style.display = 'none';
    }
</script>

<style>
    #light {
        display: none;
    }

    .btn-top-align {
        margin-top: 26px;
    }

    .btn-top-align2 {
        margin-top: 10px;
    }

    .btn span.glyphicon {
        opacity: 0;
    }

    .btn.active span.glyphicon {
        opacity: 1;
    }
</style>
<div class="container" style="width: 100% !important;">
    <div class="panel panel-default">
        <div class="panel-heading">Import Excel </div>
        <div class="panel-body">
            <div id="firstLoad">
                <div class="form-group col-sm-4 col-md-4 col-lg-4">
                    <label for="fuExcelSheet">Browse to excel file which you intend to upload </label>
                    <asp:FileUpload ID="fuExcelSheet" runat="server" CssClass="form-control" />
                    <asp:RequiredFieldValidator ID="rfFileUploadCtrl" runat="server" ControlToValidate="fuExcelSheet" ErrorMessage="Please select file" ValidationGroup="uploadclick" ForeColor="Red" Display="Dynamic"></asp:RequiredFieldValidator>
                    <asp:Label runat="server" ID="litMessage" CssClass="ErrorMessage"></asp:Label>
                </div>
                <div class="form-group col-sm-4 col-md-4 col-lg-4">
                    <label for="ddlSPList">Choose SharePoint List:</label>
                    <asp:DropDownList ID="ddlSPList" runat="server" class="form-control">
                        <asp:ListItem Text="Select..." Value="0" Selected="True"></asp:ListItem>
                        <asp:ListItem Text="ImportData" Value="ImportData"></asp:ListItem>
                        <asp:ListItem Text="ListName1" Value="ListName1"></asp:ListItem>
                        <asp:ListItem Text="ListName2" Value="ListName2"></asp:ListItem>
                        <asp:ListItem Text="ListName3" Value="ListName3"></asp:ListItem>
                        <asp:ListItem Text="ListName4" Value="ListName4"></asp:ListItem>
                    </asp:DropDownList>
                    <asp:RequiredFieldValidator ID="rfdSPList" runat="server" ControlToValidate="ddlSPList" InitialValue="0" ValidationGroup="uploadclick" ErrorMessage="*Please select the target List" ForeColor="Red" Display="Dynamic"></asp:RequiredFieldValidator>
                </div>
                <div class="form-group col-sm-4 col-md-4 col-lg-4">
                    <asp:Button ID="btnUpload" runat="server" OnClick="btnUpload_Click" Text="Load Excel" class="btn btn-default btn-top-align" ValidationGroup="uploadclick" />
                </div>
                <div class="clearfix"></div>
            </div>
            <div id="light">
                <div class="clearfix"></div>
                <div class="form-group col-sm-4 col-md-4 col-lg-4">
                    <label for="cno">Select your worksheet:</label>
                    <asp:DropDownList ID="ddlSheetName" runat="server" class="form-control" />
                </div>
                <div class="form-group col-sm-4 col-md-4 col-lg-4">
                </div>
                <div class="form-group col-sm-6 col-md-6 col-lg-6">
                    <asp:Button ID="btnShowData" runat="server" Visible="false" OnClick="btnShowData_Click" Text="Show Columns For Mapping" class="btn btn-warning btn-top-align2" />
                </div>
                <div class="clearfix"></div>
                <div class="form-group col-sm-12 col-md-12 col-lg-12">
                    <asp:Label ID="litValidationErrorMessage" runat="server"></asp:Label>
                </div>
                <div class="clearfix"></div>
                <div class="form-group col-sm-12 col-md-12 col-lg-12">
                    <asp:Repeater ID="rptBindColumns" runat="server" OnItemDataBound="rptBindColumns_OnItemDataBound">
                        <HeaderTemplate>
                            <table class="table table-bordered">
                                <thead>
                                    <tr>
                                        <th colspan="2" style="text-align: center">Map the columns</th>
                                    </tr>
                                    <tr>
                                        <th>SharePoint List Columns</th>
                                        <th>WorkSheet Columns</th>
                                    </tr>
                                </thead>
                                <tbody>
                        </HeaderTemplate>
                        <ItemTemplate>
                            <tr>
                                <td>
                                    <asp:CheckBox ID="chkXlNames" Text='<%#Eval("Value")%>' CssClass="checkbox checkbox-success" runat="server" />
                                </td>
                                <td>
                                    <asp:DropDownList ID="drpSPlistColumns" runat="server" CssClass="form-control" Width="50%" />
                                </td>
                            </tr>

                            </div>
                            </div>
                        </ItemTemplate>
                        <FooterTemplate>
                            </tbody>
                            </table>
                        </FooterTemplate>
                    </asp:Repeater>
                </div>
                <div class="clearfix"></div>
                <div class="form-group col-sm-6 col-md-6 col-lg-6">
                    <asp:Label ID="litValidationMessage" runat="server" CssClass="ErrorMessage"></asp:Label>
                </div>
                <div class="form-group col-sm-12 col-md-12 col-lg-12" style="text-align: center;">
                    <asp:Button ID="btnget" Text="Import" Visible="false" runat="Server" OnClick="btnImport_Click" CssClass="btn btn-success" />
                </div>
                <div class="clearfix"></div>

            </div>
            <div class="form-group col-sm-12 col-md-12 col-lg-12">
                <asp:Label ID="lblSuccessMessage" runat="server" CssClass="ErrorMessage"></asp:Label>
            </div>
        </div>
    </div>
</div>

