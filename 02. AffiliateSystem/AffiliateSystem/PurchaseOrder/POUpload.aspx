<%@ Page Title="" Language="vb" AutoEventWireup="false" MasterPageFile="~/SiteContent.Master" CodeBehind="POUpload.aspx.vb" Inherits="AffiliateSystem.POUpload" %>
<%@ Register Assembly="DevExpress.Web.v14.1, Version=14.1.6.0, Culture=neutral, PublicKeyToken=b88d1754d700e49a" namespace="DevExpress.Web.ASPxGridView.Export" TagPrefix="dx" %>
<%@ Register Assembly="DevExpress.Web.v14.1, Version=14.1.6.0, Culture=neutral, PublicKeyToken=b88d1754d700e49a" namespace="DevExpress.Web.ASPxGridView" TagPrefix="dx" %>
<%@ Register assembly="DevExpress.Web.v14.1, Version=14.1.6.0, Culture=neutral, PublicKeyToken=b88d1754d700e49a" namespace="DevExpress.Web.ASPxEditors" tagprefix="dx" %>
<%@ Register Assembly="DevExpress.Web.v14.1, Version=14.1.6.0, Culture=neutral, PublicKeyToken=b88d1754d700e49a" namespace="DevExpress.Web.ASPxGlobalEvents" TagPrefix="dx" %>
<%@ Register Assembly="DevExpress.Web.v14.1, Version=14.1.6.0, Culture=neutral, PublicKeyToken=b88d1754d700e49a" namespace="DevExpress.Web.ASPxPanel" TagPrefix="dx" %>
<%@ Register Assembly="DevExpress.Web.v14.1, Version=14.1.6.0, Culture=neutral, PublicKeyToken=b88d1754d700e49a" namespace="DevExpress.Web.ASPxRoundPanel" TagPrefix="dx" %>
<%@ Register Assembly="DevExpress.Web.v14.1, Version=14.1.6.0, Culture=neutral, PublicKeyToken=b88d1754d700e49a" Namespace="DevExpress.Web.ASPxCallback" TagPrefix="dx" %>
<%@ Register Assembly="DevExpress.Web.v14.1, Version=14.1.6.0, Culture=neutral, PublicKeyToken=b88d1754d700e49a" Namespace="DevExpress.Web.ASPxUploadControl" TagPrefix="dx" %>

<asp:Content ID="Content1" ContentPlaceHolderID="head" runat="server">
    <style type="text/css">
        .dxeHLC, .dxeHC, .dxeHFC
        {
            display: none;
        }
        .style1
        {
            width: 5px;
            height: 20px;
        }
        .style2
        {
            width: 100px;
            height: 20px;
        }
        .style3
        {
            height: 20px;
        }
    </style>
<script type="text/javascript">
    function OnInit(s, e) {
        AdjustSizeGrid();
    }
    function OnControlsInitializedGrid(s, e) {
        ASPxClientUtils.AttachEventToElement(window, "resize", function (evt) {
            AdjustSizeGrid();
        });
    }
    function AdjustSizeGrid() {

        var myWidth = 0, myHeight = 0;
        if (typeof (window.innerWidth) == 'number') {
            //Non-IE
            myWidth = window.innerWidth;
            myHeight = window.innerHeight;
        } else if (document.documentElement && (document.documentElement.clientWidth || document.documentElement.clientHeight)) {
            //IE 6+ in 'standards compliant mode'
            myWidth = document.documentElement.clientWidth;
            myHeight = document.documentElement.clientHeight;
        } else if (document.body && (document.body.clientWidth || document.body.clientHeight)) {
            //IE 4 compatible
            myWidth = document.body.clientWidth;
            myHeight = document.body.clientHeight;
        }

        var height = Math.max(0, myHeight);
        height = height - (height * 33 / 100)
        grid.SetHeight(height);
    }

    function memo_OnInit(s, e) {
        var input = memo.GetInputElement();
        if (ASPxClientUtils.opera)
            input.oncontextmenu = function () { return false; };
        else
            input.onpaste = CorrectTextWithDelay;
    }

    function CorrectTextWithDelay() {
        var maxLength = se.GetNumber();
        setTimeout(function () { memo.SetText(memo.GetText().substr(0, maxLength)); }, 0);
    }

    function Uploader_OnUploadStart() {
        btnUpload.SetEnabled(false);
    }

    function Uploader_OnFilesUploadComplete(args) {
        UpdateUploadButton();
    }

    function UpdateUploadButton() {
        btnUpload.SetEnabled(uploader.GetText(0) != "");
        var a = uploader.GetText();
        var b = filename.SetText(a);
    }

    var order;
    var pFieldName;

    function onSorting(s, e) {
        order = order == "ASC" ? "DESC" : "ASC";
        e.cancel = true;
        pFieldName = e.column.fieldName
        s.PerformCallback('sorting|' + order + '|' + pFieldName);
    }
</script>
</asp:Content>

<asp:Content ID="Content2" ContentPlaceHolderID="Content" runat="server">
    <table style="width:100%;">
        <tr>
            <td  width="100%" colspan="3">
                <table style="border-width: 1pt thin thin thin; border-style: ridge; border-color:#9598A1; width:100%; height: 50px;">
                    <tr>
                        <td colspan="8" height="20">
                            <table id="Table1">
                                <tr>
                                    <td style="width:5px;"></td>
                                    <td align="left" valign="middle" height="20px" width="60px">
                                        <dx:ASPxLabel ID="ASPxLabel1" runat="server" Text="FILE"
                                            Font-Names="Tahoma" Font-Size="8pt">
                                        </dx:ASPxLabel>
                                    </td>
                                    <td align="left" valign="middle" height="20px" width="380px">
                                        <dx:ASPxUploadControl ID="Uploader" runat="server"  
                                            Width="100%" Font-Names="Verdana" Font-Size="8pt"
                                            ClientInstanceName="Uploader"
                                            ShowClearFileSelectionButton="False"
                                            NullText="Click here to browse files..."
                                            OnFileUploadComplete="Uploader_FileUploadComplete">
                                            <ClientSideEvents FilesUploadComplete="function(s, e) { Uploader_OnFilesUploadComplete(e); }"
                                                FileUploadComplete="function(s, e) { Uploader_OnFileUploadComplete(e); }" 
                                                FileUploadStart="function(s, e) { Uploader_OnUploadStart(); }"
                                                TextChanged="function(s, e) { var test = uploader.GetText(); txtFileName.SetText(test); UpdateUploadButton(); }" />
                                            <ValidationSettings AllowedFileExtensions=".xls,.xlsx" MaxFileSize="4000000" />
                                            <BrowseButton Text="...">
                                            </BrowseButton> 
                                            <BrowseButtonStyle Paddings-Padding="3px" >
                                            </BrowseButtonStyle>
                                        </dx:ASPxUploadControl>
                                    </td>
                                    <td align="right" width="180px">
                                        <table style="width:100%;" align="right">
                                            <tr>
                                                <td align="right" valign="middle" style="height:25px; width:90px;">
                                                    <dx:ASPxButton ID="btnUpload" runat="server" Text="CHECK FILE" ClientInstanceName="btnUpload" 
                                                        Font-Names="Tahoma" Width="85px" AutoPostBack="False" Font-Size="8pt" >                                                        
                                                        <ClientSideEvents Click="function(s, e) {
                                                            grid.PerformCallback('checkItem');
                                                        }" />
                                                    </dx:ASPxButton>
                                                </td>
                                                <td align="right" valign="middle" style="height:25px; width:90px;">
                                                    <dx:ASPxButton ID="btnClear" runat="server" Text="CLEAR"
                                                        Font-Names="Tahoma" Width="85px" AutoPostBack="False" Font-Size="8pt">                                 
                                                    </dx:ASPxButton>
                                                </td>
                                            </tr>
                                        </table>
                                    </td>
                                </tr>
                            </table>
                        </td>
                    </tr>
                </table>
            </td>            
        </tr>

        <tr>
            <td colspan="3" height="15">
            <%--error message--%>
                <table id="tblMsg" style="border-width: thin; border-style: inset hidden ridge hidden; border-color:#9598A1; width:100%; height: 15px;">
                    <tr>
                        <td align="left" valign="middle" height="15px">
                            <dx:ASPxLabel ID="lblInfo" runat="server" Text="[lblinfo]" Font-Names="Tahoma" 
                                ClientInstanceName="lblInfo" Font-Bold="True" Font-Italic="True" 
                                Font-Size="8pt" ForeColor="Red">
                            </dx:ASPxLabel>
                        </td>
                    </tr>         
                </table>
            </td>            
        </tr>

        <tr>
            <td colspan="3" align="left" valign="top" height="220">
                <dx:ASPxGridView ID="grid" runat="server" Width="100%"
                    Font-Names="Tahoma" KeyFieldName="NoUrut;PartNo;PONo"
                    AutoGenerateColumns="False"
                    ClientInstanceName="grid" Font-Size="8pt">
                    <ClientSideEvents Init="OnInit" 
                        CallbackError="function(s, e) {e.handled = true;}" EndCallback="function(s, e) {
                        
                        var pMsg = s.cpMessage;
                        if (pMsg != '') {
                            if (pMsg.substring(1,5) == '1001' || pMsg.substring(1,5) == '1002' || pMsg.substring(1,5) == '1003' || pMsg.substring(1,5) == '7001') {
                                lblInfo.GetMainElement().style.color = 'Blue';
                            } else {
                                lblInfo.GetMainElement().style.color = 'Red';
                            }
                            lblInfo.SetText(pMsg);
                        } else {
                            lblInfo.SetText('');
                        }
                        delete s.cpMessage;
                    }" RowClick="function(s, e) {
	                    lblInfo.SetText('');
                    }" />                    
                    <Columns>
                        <dx:GridViewDataTextColumn VisibleIndex="1" Caption="NO." FieldName="NoUrut" Width="30px">
                            <HeaderStyle HorizontalAlign="Center"></HeaderStyle>
                            <CellStyle Font-Names="Tahoma" Font-Size="8pt">
                            </CellStyle>
                        </dx:GridViewDataTextColumn>
                        <dx:GridViewDataTextColumn VisibleIndex="2" Caption="PERIOD." FieldName="Period" Width="60px">
                            <PropertiesTextEdit DisplayFormatString="MMM yyyy">
                            </PropertiesTextEdit>
                            <HeaderStyle HorizontalAlign="Center"></HeaderStyle>
                            <CellStyle Font-Names="Tahoma" Font-Size="8pt">
                            </CellStyle>
                        </dx:GridViewDataTextColumn>
                        <dx:GridViewDataTextColumn VisibleIndex="3" Caption="PO NO." FieldName="PONo" Width="90px">
                            <HeaderStyle HorizontalAlign="Center"></HeaderStyle>
                            <CellStyle Font-Names="Tahoma" Font-Size="8pt">
                            </CellStyle>
                        </dx:GridViewDataTextColumn>
                        <dx:GridViewDataTextColumn VisibleIndex="4" Caption="SHIP BY" FieldName="ShipCls" Width="60px">
                            <HeaderStyle HorizontalAlign="Center"></HeaderStyle>
                            <CellStyle Font-Names="Tahoma" Font-Size="8pt">
                            </CellStyle>
                        </dx:GridViewDataTextColumn>
                        <dx:GridViewDataTextColumn VisibleIndex="5" Caption="PART NO." FieldName="PartNo" Width="90px">
                            <HeaderStyle HorizontalAlign="Center"></HeaderStyle>
                            <CellStyle Font-Names="Tahoma" Font-Size="8pt">
                            </CellStyle>
                        </dx:GridViewDataTextColumn>
                        <dx:GridViewDataTextColumn VisibleIndex="6" Caption="PART NAME" FieldName="PartName" Width="180px">
                            <HeaderStyle HorizontalAlign="Center"></HeaderStyle>
                            <CellStyle Font-Names="Tahoma" Font-Size="8pt">
                            </CellStyle>
                        </dx:GridViewDataTextColumn>
                        <dx:GridViewDataTextColumn VisibleIndex="7" Caption="UOM" FieldName="UnitDesc" Width="40px">
                            <HeaderStyle HorizontalAlign="Center"></HeaderStyle>
                            <CellStyle Font-Names="Tahoma" Font-Size="8pt">
                            </CellStyle>
                        </dx:GridViewDataTextColumn>
                        <dx:GridViewDataTextColumn VisibleIndex="8" Caption="MOQ" FieldName="MinOrderQty" Width="50px">
                            <PropertiesTextEdit DisplayFormatString="{0:n0}">
                            </PropertiesTextEdit>
                            <HeaderStyle HorizontalAlign="Center"></HeaderStyle>
                            <CellStyle Font-Names="Tahoma" Font-Size="8pt">
                            </CellStyle>
                        </dx:GridViewDataTextColumn>
                        <dx:GridViewDataTextColumn VisibleIndex="9" Caption="MAKER" FieldName="Maker" Width="80px">
                            <HeaderStyle HorizontalAlign="Center"></HeaderStyle>
                            <CellStyle Font-Names="Tahoma" Font-Size="8pt">
                            </CellStyle>
                        </dx:GridViewDataTextColumn>
                        <dx:GridViewDataTextColumn VisibleIndex="10" Caption="PROJECT" FieldName="Project" Width="80px">
                            <HeaderStyle HorizontalAlign="Center"></HeaderStyle>
                            <CellStyle Font-Names="Tahoma" Font-Size="8pt">
                            </CellStyle>
                        </dx:GridViewDataTextColumn>
                        <dx:GridViewDataTextColumn VisibleIndex="11" Caption="SUPPLIER" FieldName="SupplierID" Width="100px">
                            <HeaderStyle HorizontalAlign="Center"></HeaderStyle>
                            <CellStyle Font-Names="Tahoma" Font-Size="8pt">
                            </CellStyle>
                        </dx:GridViewDataTextColumn>
                        <dx:GridViewDataTextColumn VisibleIndex="12" Caption="TOTAL FIRM QTY" FieldName="POQty" Width="80px">
                            <PropertiesTextEdit DisplayFormatString="{0:n0}">
                            </PropertiesTextEdit>
                            <HeaderStyle HorizontalAlign="Center"></HeaderStyle>
                            <CellStyle Font-Names="Tahoma" Font-Size="8pt">
                            </CellStyle>
                        </dx:GridViewDataTextColumn>

                        <dx:GridViewDataTextColumn VisibleIndex="14" Caption="FORECAST N+1" FieldName="ForecastN1" Width="70px">
                            <PropertiesTextEdit DisplayFormatString="{0:n0}">
                            </PropertiesTextEdit>
                            <HeaderStyle HorizontalAlign="Center"></HeaderStyle>
                            <CellStyle Font-Names="Tahoma" Font-Size="8pt">
                            </CellStyle>
                        </dx:GridViewDataTextColumn>
                        <dx:GridViewDataTextColumn VisibleIndex="15" Caption="FORECAST N+2" FieldName="ForecastN2" Width="70px">
                            <PropertiesTextEdit DisplayFormatString="{0:n0}">
                            </PropertiesTextEdit>
                            <HeaderStyle HorizontalAlign="Center"></HeaderStyle>
                            <CellStyle Font-Names="Tahoma" Font-Size="8pt">
                            </CellStyle>
                        </dx:GridViewDataTextColumn>
                        <dx:GridViewDataTextColumn VisibleIndex="16" Caption="FORECAST N+3" FieldName="ForecastN3" Width="70px">
                            <PropertiesTextEdit DisplayFormatString="{0:n0}">
                            </PropertiesTextEdit>
                            <HeaderStyle HorizontalAlign="Center"></HeaderStyle>
                            <CellStyle Font-Names="Tahoma" Font-Size="8pt">
                            </CellStyle>
                        </dx:GridViewDataTextColumn>

                        <dx:GridViewBandColumn Caption="E.T.A SCHEDULE (BASED ON FIRM ORDER)" VisibleIndex="17">
                            <Columns>
                                <dx:GridViewDataTextColumn VisibleIndex="18" Caption="1" Width="60px" FieldName="DeliveryD1">
                                    <PropertiesTextEdit DisplayFormatString="{0:n0}">
                                    </PropertiesTextEdit>
                                    <HeaderStyle HorizontalAlign="Center"></HeaderStyle>
                                    <CellStyle Font-Names="Tahoma" Font-Size="8pt">
                                    </CellStyle>
                                </dx:GridViewDataTextColumn>
                                <dx:GridViewDataTextColumn VisibleIndex="19" Caption="2" Width="60px" FieldName="DeliveryD2">
                                    <PropertiesTextEdit DisplayFormatString="{0:n0}">
                                    </PropertiesTextEdit>
                                    <HeaderStyle HorizontalAlign="Center"></HeaderStyle>
                                    <CellStyle Font-Names="Tahoma" Font-Size="8pt">
                                    </CellStyle>
                                </dx:GridViewDataTextColumn>                               
                                <dx:GridViewDataTextColumn VisibleIndex="20" Caption="3" Width="60px" FieldName="DeliveryD3">
                                    <PropertiesTextEdit DisplayFormatString="{0:n0}">
                                    </PropertiesTextEdit>
                                    <HeaderStyle HorizontalAlign="Center"></HeaderStyle>
                                    <CellStyle Font-Names="Tahoma" Font-Size="8pt">
                                    </CellStyle>
                                </dx:GridViewDataTextColumn>                                
                                <dx:GridViewDataTextColumn VisibleIndex="21" Caption="4" Width="60px" FieldName="DeliveryD4">
                                    <PropertiesTextEdit DisplayFormatString="{0:n0}">
                                    </PropertiesTextEdit>
                                    <HeaderStyle HorizontalAlign="Center"></HeaderStyle>
                                    <CellStyle Font-Names="Tahoma" Font-Size="8pt">
                                    </CellStyle>
                                </dx:GridViewDataTextColumn>                                
                                <dx:GridViewDataTextColumn VisibleIndex="22" Caption="5" Width="60px" FieldName="DeliveryD5">
                                    <PropertiesTextEdit DisplayFormatString="{0:n0}">
                                    </PropertiesTextEdit>
                                    <HeaderStyle HorizontalAlign="Center"></HeaderStyle>
                                    <CellStyle Font-Names="Tahoma" Font-Size="8pt">
                                    </CellStyle>
                                </dx:GridViewDataTextColumn>                                
                                <dx:GridViewDataTextColumn VisibleIndex="23" Caption="6" Width="60px" FieldName="DeliveryD6">
                                    <PropertiesTextEdit DisplayFormatString="{0:n0}">                                        
                                    </PropertiesTextEdit>
                                    <HeaderStyle HorizontalAlign="Center"></HeaderStyle>
                                    <CellStyle Font-Names="Tahoma" Font-Size="8pt">
                                    </CellStyle>
                                </dx:GridViewDataTextColumn>                                
                                <dx:GridViewDataTextColumn VisibleIndex="24" Caption="7" Width="60px" FieldName="DeliveryD7">
                                    <PropertiesTextEdit DisplayFormatString="{0:n0}">                                        
                                    </PropertiesTextEdit>
                                    <HeaderStyle HorizontalAlign="Center"></HeaderStyle>
                                    <CellStyle Font-Names="Tahoma" Font-Size="8pt">
                                    </CellStyle>
                                </dx:GridViewDataTextColumn>                               
                                <dx:GridViewDataTextColumn VisibleIndex="25" Caption="8" Width="60px" FieldName="DeliveryD8">
                                    <PropertiesTextEdit DisplayFormatString="{0:n0}">                                        
                                    </PropertiesTextEdit>
                                    <HeaderStyle HorizontalAlign="Center"></HeaderStyle>
                                    <CellStyle Font-Names="Tahoma" Font-Size="8pt">
                                    </CellStyle>
                                </dx:GridViewDataTextColumn>                                
                                <dx:GridViewDataTextColumn VisibleIndex="26" Caption="9" Width="60px" FieldName="DeliveryD9">
                                    <PropertiesTextEdit DisplayFormatString="{0:n0}">                                        
                                    </PropertiesTextEdit>
                                    <HeaderStyle HorizontalAlign="Center"></HeaderStyle>
                                    <CellStyle Font-Names="Tahoma" Font-Size="8pt">
                                    </CellStyle>
                                </dx:GridViewDataTextColumn>                                
                                <dx:GridViewDataTextColumn VisibleIndex="27" Caption="10" Width="60px" FieldName="DeliveryD10">
                                    <PropertiesTextEdit DisplayFormatString="{0:n0}">                                        
                                    </PropertiesTextEdit>
                                    <HeaderStyle HorizontalAlign="Center"></HeaderStyle>
                                    <CellStyle Font-Names="Tahoma" Font-Size="8pt">
                                    </CellStyle>
                                </dx:GridViewDataTextColumn>                                
                                <dx:GridViewDataTextColumn VisibleIndex="28" Caption="11" Width="60px" FieldName="DeliveryD11">
                                    <PropertiesTextEdit DisplayFormatString="{0:n0}">                                        
                                    </PropertiesTextEdit>
                                    <HeaderStyle HorizontalAlign="Center"></HeaderStyle>
                                    <CellStyle Font-Names="Tahoma" Font-Size="8pt">
                                    </CellStyle>
                                </dx:GridViewDataTextColumn>                                
                                <dx:GridViewDataTextColumn VisibleIndex="29" Caption="12" Width="60px" FieldName="DeliveryD12">
                                    <PropertiesTextEdit DisplayFormatString="{0:n0}">                                        
                                    </PropertiesTextEdit>
                                    <HeaderStyle HorizontalAlign="Center"></HeaderStyle>
                                    <CellStyle Font-Names="Tahoma" Font-Size="8pt">
                                    </CellStyle>
                                </dx:GridViewDataTextColumn>                                
                                <dx:GridViewDataTextColumn VisibleIndex="30" Caption="13" Width="60px" FieldName="DeliveryD13">
                                    <PropertiesTextEdit DisplayFormatString="{0:n0}">                                        
                                    </PropertiesTextEdit>
                                    <HeaderStyle HorizontalAlign="Center"></HeaderStyle>
                                    <CellStyle Font-Names="Tahoma" Font-Size="8pt">
                                    </CellStyle>
                                </dx:GridViewDataTextColumn>                                
                                <dx:GridViewDataTextColumn VisibleIndex="31" Caption="14" Width="60px" FieldName="DeliveryD14">
                                    <PropertiesTextEdit DisplayFormatString="{0:n0}">                                        
                                    </PropertiesTextEdit>
                                    <HeaderStyle HorizontalAlign="Center"></HeaderStyle>
                                    <CellStyle Font-Names="Tahoma" Font-Size="8pt">
                                    </CellStyle>
                                </dx:GridViewDataTextColumn>                                
                                <dx:GridViewDataTextColumn VisibleIndex="32" Caption="15" Width="60px" FieldName="DeliveryD15">
                                    <PropertiesTextEdit DisplayFormatString="{0:n0}">                                        
                                    </PropertiesTextEdit>
                                    <HeaderStyle HorizontalAlign="Center"></HeaderStyle>
                                    <CellStyle Font-Names="Tahoma" Font-Size="8pt">
                                    </CellStyle>
                                </dx:GridViewDataTextColumn>                                
                                <dx:GridViewDataTextColumn VisibleIndex="33" Caption="16" Width="60px" FieldName="DeliveryD16">
                                    <PropertiesTextEdit DisplayFormatString="{0:n0}">                                        
                                    </PropertiesTextEdit>
                                    <HeaderStyle HorizontalAlign="Center"></HeaderStyle>
                                    <CellStyle Font-Names="Tahoma" Font-Size="8pt">
                                    </CellStyle>
                                </dx:GridViewDataTextColumn>                                
                                <dx:GridViewDataTextColumn VisibleIndex="34" Caption="17" Width="60px" FieldName="DeliveryD17">
                                    <PropertiesTextEdit DisplayFormatString="{0:n0}">                                        
                                    </PropertiesTextEdit>
                                    <HeaderStyle HorizontalAlign="Center"></HeaderStyle>
                                    <CellStyle Font-Names="Tahoma" Font-Size="8pt">
                                    </CellStyle>
                                </dx:GridViewDataTextColumn>                                
                                <dx:GridViewDataTextColumn VisibleIndex="35" Caption="18" Width="60px" FieldName="DeliveryD18">
                                    <PropertiesTextEdit DisplayFormatString="{0:n0}">                                        
                                    </PropertiesTextEdit>
                                    <HeaderStyle HorizontalAlign="Center"></HeaderStyle>
                                    <CellStyle Font-Names="Tahoma" Font-Size="8pt">
                                    </CellStyle>
                                </dx:GridViewDataTextColumn>                                
                                <dx:GridViewDataTextColumn VisibleIndex="36" Caption="19" Width="60px" FieldName="DeliveryD19">
                                    <PropertiesTextEdit DisplayFormatString="{0:n0}">                                        
                                    </PropertiesTextEdit>
                                    <HeaderStyle HorizontalAlign="Center"></HeaderStyle>
                                    <CellStyle Font-Names="Tahoma" Font-Size="8pt">
                                    </CellStyle>
                                </dx:GridViewDataTextColumn>                                
                                <dx:GridViewDataTextColumn VisibleIndex="37" Caption="20" Width="60px" FieldName="DeliveryD20">
                                    <PropertiesTextEdit DisplayFormatString="{0:n0}">                                        
                                    </PropertiesTextEdit>
                                    <HeaderStyle HorizontalAlign="Center"></HeaderStyle>
                                    <CellStyle Font-Names="Tahoma" Font-Size="8pt">
                                    </CellStyle>
                                </dx:GridViewDataTextColumn>                                
                                <dx:GridViewDataTextColumn VisibleIndex="38" Caption="21" Width="60px" FieldName="DeliveryD21">
                                    <PropertiesTextEdit DisplayFormatString="{0:n0}">                                        
                                    </PropertiesTextEdit>
                                    <HeaderStyle HorizontalAlign="Center"></HeaderStyle>
                                    <CellStyle Font-Names="Tahoma" Font-Size="8pt">
                                    </CellStyle>
                                </dx:GridViewDataTextColumn>                                
                                <dx:GridViewDataTextColumn VisibleIndex="39" Caption="22" Width="60px" FieldName="DeliveryD22">
                                    <PropertiesTextEdit DisplayFormatString="{0:n0}">                                        
                                    </PropertiesTextEdit>
                                    <HeaderStyle HorizontalAlign="Center"></HeaderStyle>
                                    <CellStyle Font-Names="Tahoma" Font-Size="8pt">
                                    </CellStyle>
                                </dx:GridViewDataTextColumn>                                
                                <dx:GridViewDataTextColumn VisibleIndex="40" Caption="23" Width="60px" FieldName="DeliveryD23">
                                    <PropertiesTextEdit DisplayFormatString="{0:n0}">                                        
                                    </PropertiesTextEdit>
                                    <HeaderStyle HorizontalAlign="Center"></HeaderStyle>
                                    <CellStyle Font-Names="Tahoma" Font-Size="8pt">
                                    </CellStyle>
                                </dx:GridViewDataTextColumn>                                
                                <dx:GridViewDataTextColumn VisibleIndex="41" Caption="24" Width="60px" FieldName="DeliveryD24">
                                    <PropertiesTextEdit DisplayFormatString="{0:n0}">                                        
                                    </PropertiesTextEdit>
                                    <HeaderStyle HorizontalAlign="Center"></HeaderStyle>
                                    <CellStyle Font-Names="Tahoma" Font-Size="8pt">
                                    </CellStyle>
                                </dx:GridViewDataTextColumn>                               
                                <dx:GridViewDataTextColumn VisibleIndex="42" Caption="25" Width="60px" FieldName="DeliveryD25">
                                    <PropertiesTextEdit DisplayFormatString="{0:n0}">                                        
                                    </PropertiesTextEdit>
                                    <HeaderStyle HorizontalAlign="Center"></HeaderStyle>
                                    <CellStyle Font-Names="Tahoma" Font-Size="8pt">
                                    </CellStyle>
                                </dx:GridViewDataTextColumn>                                
                                <dx:GridViewDataTextColumn VisibleIndex="43" Caption="26" Width="60px" FieldName="DeliveryD26">
                                    <PropertiesTextEdit DisplayFormatString="{0:n0}">                                        
                                    </PropertiesTextEdit>
                                    <HeaderStyle HorizontalAlign="Center"></HeaderStyle>
                                    <CellStyle Font-Names="Tahoma" Font-Size="8pt">
                                    </CellStyle>
                                </dx:GridViewDataTextColumn>                                
                                <dx:GridViewDataTextColumn VisibleIndex="44" Caption="27" Width="60px" FieldName="DeliveryD27">
                                    <PropertiesTextEdit DisplayFormatString="{0:n0}">                                        
                                    </PropertiesTextEdit>
                                    <HeaderStyle HorizontalAlign="Center"></HeaderStyle>
                                    <CellStyle Font-Names="Tahoma" Font-Size="8pt">
                                    </CellStyle>
                                </dx:GridViewDataTextColumn>                                
                                <dx:GridViewDataTextColumn VisibleIndex="45" Caption="28" Width="60px" FieldName="DeliveryD28">
                                    <PropertiesTextEdit DisplayFormatString="{0:n0}">                                        
                                    </PropertiesTextEdit>
                                    <HeaderStyle HorizontalAlign="Center"></HeaderStyle>
                                    <CellStyle Font-Names="Tahoma" Font-Size="8pt">
                                    </CellStyle>
                                </dx:GridViewDataTextColumn>                               
                                <dx:GridViewDataTextColumn VisibleIndex="46" Caption="29" Width="60px" FieldName="DeliveryD29">
                                    <PropertiesTextEdit DisplayFormatString="{0:n0}">                                        
                                    </PropertiesTextEdit>
                                    <HeaderStyle HorizontalAlign="Center"></HeaderStyle>
                                    <CellStyle Font-Names="Tahoma" Font-Size="8pt">
                                    </CellStyle>
                                </dx:GridViewDataTextColumn>                                
                                <dx:GridViewDataTextColumn VisibleIndex="47" Caption="30" Width="60px" FieldName="DeliveryD30">
                                    <PropertiesTextEdit DisplayFormatString="{0:n0}">                                        
                                    </PropertiesTextEdit>
                                    <HeaderStyle HorizontalAlign="Center"></HeaderStyle>
                                    <CellStyle Font-Names="Tahoma" Font-Size="8pt">
                                    </CellStyle>
                                </dx:GridViewDataTextColumn>                                
                                <dx:GridViewDataTextColumn VisibleIndex="48" Caption="31" Width="60px" FieldName="DeliveryD31">
                                    <PropertiesTextEdit DisplayFormatString="{0:n0}">                                        
                                    </PropertiesTextEdit>
                                    <HeaderStyle HorizontalAlign="Center"></HeaderStyle>
                                    <CellStyle Font-Names="Tahoma" Font-Size="8pt">
                                    </CellStyle>
                                </dx:GridViewDataTextColumn>
                            </Columns>
                            <HeaderStyle HorizontalAlign="Center" />
                        </dx:GridViewBandColumn>
                        <dx:GridViewDataTextColumn VisibleIndex="49" Caption="Remarks" FieldName="ErrorCls" Width="300px">
                            <HeaderStyle HorizontalAlign="Center"></HeaderStyle>
                            <CellStyle Font-Names="Tahoma" Font-Size="8pt">
                            </CellStyle>
                        </dx:GridViewDataTextColumn>
                        <dx:GridViewDataTextColumn VisibleIndex="50" Caption="KanbanCls" FieldName="KanbanCls" Width="0px">
                            <HeaderStyle HorizontalAlign="Center"></HeaderStyle>
                            <CellStyle Font-Names="Tahoma" Font-Size="8pt">
                            </CellStyle>
                        </dx:GridViewDataTextColumn>
                        <dx:GridViewDataTextColumn VisibleIndex="50" Caption="Period" FieldName="Period2" Width="0px">
                            <HeaderStyle HorizontalAlign="Center"></HeaderStyle>
                            <CellStyle Font-Names="Tahoma" Font-Size="8pt">
                            </CellStyle>
                        </dx:GridViewDataTextColumn>
                    </Columns>
                    <SettingsBehavior AllowSelectByRowClick="True" AllowSort="False"
                        ColumnResizeMode="Control" EnableRowHotTrack="True" />
                    <SettingsPager PageSize="10" Position="Top">
                        <Summary Text="Page {0} of {1} [{2} record(s)]"
                                  AllPagesText="Page {0} of {1} " />
                        <Summary AllPagesText="Page {0} of {1} " Text="Page {0} of {1} [{2} record(s)]"></Summary>
                    </SettingsPager>
                    <SettingsEditing Mode="Batch" NewItemRowPosition="Bottom">
                        <BatchEditSettings ShowConfirmOnLosingChanges="False" />
                    </SettingsEditing>
                    <Settings ShowHorizontalScrollBar="True" ShowVerticalScrollBar="True"
                        ShowGroupButtons="False" ShowStatusBar="Hidden"
                        VerticalScrollableHeight="220" />
                    <Styles>
                        <SelectedRow ForeColor="Black">
                        </SelectedRow>
                    </Styles>
                    <StylesEditors ButtonEditCellSpacing="0">
                        <ProgressBar Height="21px">
                        </ProgressBar>
                    </StylesEditors>
                </dx:ASPxGridView>
            </td>
        </tr>

        <tr>
            <td valign="top" align="left">
                <dx:ASPxButton ID="btnSubMenu" runat="server" Text="SUB MENU"
                    Font-Names="Tahoma" Width="85px" Font-Size="8pt">
                </dx:ASPxButton>
            </td>
            <td valign="top" align="right" style="width: 50px;">
                <dx:ASPxButton ID="btnDownload" runat="server" Text="DOWNLOAD TEMPLATE" ClientInstanceName="btnDownload"
                    Font-Names="Tahoma" Width="85px" AutoPostBack="False" Font-Size="8pt" >
                    <ClientSideEvents Click="function(s, e) {ASPxCallback1.PerformCallback();}" />
                </dx:ASPxButton>
            </td>       
            <td valign="top" align="right" style="width: 50px;">
                <dx:ASPxButton ID="btnSave" runat="server" Text="SAVE" ClientInstanceName="btnSave"
                    Font-Names="Tahoma" Width="85px" AutoPostBack="False" Font-Size="8pt" >
                    <ClientSideEvents Click="function(s, e) {grid.PerformCallback('save');}" />
                </dx:ASPxButton>
            </td>
        </tr>
    </table>

    <dx:ASPxCallback ID="ASPxCallback1" runat="server" ClientInstanceName="ASPxCallback1">
        <ClientSideEvents EndCallback="function(s, e) {
            var pMsg = s.cpMessage;
            if (pMsg != '') {
                if (pMsg.substring(1,5) == '9998') {
                    lblInfo.GetMainElement().style.color = 'Blue';
                } else {
                    lblInfo.GetMainElement().style.color = 'Red';
                }
                lblInfo.SetText(pMsg);
            } else {
                lblInfo.SetText('');
            }
            delete s.cpMessage;
        }" />       
    </dx:ASPxCallback>

    <dx:ASPxGlobalEvents ID="ge" runat="server" >
        <ClientSideEvents ControlsInitialized="function(s, e) {
	        OnControlsInitializedSplitter();
	        OnControlsInitializedGrid();
        }" />
    </dx:ASPxGlobalEvents>
</asp:Content>

