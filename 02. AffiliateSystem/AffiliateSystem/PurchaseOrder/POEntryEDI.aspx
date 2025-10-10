<%@ Page Title="" Language="vb" AutoEventWireup="false" MasterPageFile="~/SiteContent.Master" CodeBehind="POEntryEDI.aspx.vb" Inherits="AffiliateSystem.POEntryEDI" %>
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
        height = height - (height * 35 / 100)
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
            <td width="100%">
                <table style="border-width: 1pt thin thin thin; border-style: ridge; border-color:#9598A1; width:100%; height: 50px;">
                    <tr>
                        <td colspan="8" height="20">
                            <table id="Table1">
                                <tr>
                                    <td style="width:5px;"></td>
                                    <td align="left" valign="middle" height="20px" width="120px">
                                        <dx:ASPxLabel ID="ASPxLabel2" runat="server" Text="MONTH"
                                            Font-Names="Tahoma" Font-Size="8pt">
                                        </dx:ASPxLabel>
                                    </td>                                    
                                        <td align="left" valign="middle" height="25px" width="120px">
                                            <dx:ASPxTimeEdit ID="dtPeriodFrom" runat="server" ClientInstanceName="dtPeriodFrom" 
                                                DisplayFormatString="MMM yyyy" EditFormat="Custom" 
                                                EditFormatString="MMM yyyy" Font-Names="Tahoma" Font-Size="8pt" Width="120px">
                                                <ClientSideEvents ValueChanged="function(s, e) {
                                                    grid.PerformCallback('kosong');
	                                                lblInfo.SetText('');
                                                }" />
                                            </dx:ASPxTimeEdit>
                                        </td>                                    
                                    <td align="right" width="180px">                                        
                                    </td>
                                </tr>
                                <tr>
                                    <td style="width:5px;"></td>
                                    <td align="left" valign="middle" height="20px" width="120px">
                                        <dx:ASPxLabel ID="ASPxLabel1" runat="server" Text="OES FILE"
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
                                            <ValidationSettings AllowedFileExtensions=".txt" MaxFileSize="4000000" />
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
                                                    <dx:ASPxButton ID="btnUpload" runat="server" Text="IMPORT" ClientInstanceName="btnUpload" 
                                                        Font-Names="Tahoma" Width="85px" AutoPostBack="False" Font-Size="8pt" >                                                        
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
    </table>

    <div style="height:1px;"></div>

    <table style="width:100%; height: 15px;">
        <tr>
            <td colspan="8" height="15">
            <%--error message--%>
                <table id="tblMsg" style="border-width: thin; border-style: inset hidden ridge hidden; border-color:#9598A1; width:100%; height: 15px;">
                    <tr>
                        <td align="left" valign="middle" height="15px">
                            <dx:ASPxLabel ID="lblInfo" runat="server" Text="[lblinfo]" Font-Names="Tahoma" 
                                ClientInstanceName="lblInfo" Font-Bold="True" Font-Italic="True" 
                                Font-Size="8pt">
                            </dx:ASPxLabel>
                        </td>
                    </tr>         
                </table>
            </td>            
        </tr>
    </table>

    <div style="height:10px;"></div>
   
<%--    <table style="width:100%;">
        <tr>
            <td valign="top" align="left">
                <dx:ASPxProgressBar ID="ProgressBar" runat="server" 
                    Width="100%" ClientInstanceName="ProgressBar" Font-Names="verdana" 
                    Theme="Office2010Silver">
                </dx:ASPxProgressBar>
            </td>
        </tr>
    </table>--%>

<%--    <div style="height:10px;"></div>--%>

    <table style="width:100%;">
        <tr>
            <td align="left" valign="middle" style="height:20px; width:50px;">
                <dx:ASPxLabel ID="ASPxLabel5" runat="server" Text="PO NO. (REGULER)"
                    Font-Names="Tahoma" Font-Size="8pt">
                </dx:ASPxLabel>
            </td>
            <td align="left" valign="middle" height="20px" width="180px">
                <dx:ASPxTextBox ID="txtPONoReg" runat="server" Width="180px" 
                    ClientInstanceName="txtPONoReg" Font-Names="Tahoma" Font-Size="8pt"
                    MaxLength="20" onkeypress="return singlequote(event)" Height="20px">
                    <ClientSideEvents LostFocus="function(s, e) { 
	                    lblInfo.SetText('');
                    }" />
                </dx:ASPxTextBox>
            </td>
        </tr>
        <tr>
            <td align="left" valign="top" height="220" colspan="2">
                <dx:ASPxGridView ID="grid" runat="server" Width="100%"
                    Font-Names="Tahoma" KeyFieldName="PartNo;PONo"
                    AutoGenerateColumns="False"
                    ClientInstanceName="grid" Font-Size="8pt">
                    <ClientSideEvents Init="OnInit"/>                    
                    <Columns>
                        <dx:GridViewDataTextColumn VisibleIndex="0" Caption="NO." FieldName="NoUrut" Width="30px" HeaderStyle-HorizontalAlign="Center">                            
<HeaderStyle HorizontalAlign="Center"></HeaderStyle>

                            <CellStyle Font-Names="Tahoma" Font-Size="8pt">
                            </CellStyle>
                        </dx:GridViewDataTextColumn>
                        <dx:GridViewDataTextColumn VisibleIndex="1" Caption="ORDER NO." FieldName="OrderNo" Width="90px" HeaderStyle-HorizontalAlign="Center">                            
<HeaderStyle HorizontalAlign="Center"></HeaderStyle>

                            <CellStyle Font-Names="Tahoma" Font-Size="8pt">
                            </CellStyle>
                        </dx:GridViewDataTextColumn>
                        <dx:GridViewDataTextColumn VisibleIndex="2" Caption="PART NO." FieldName="PartNo" Width="90px" HeaderStyle-HorizontalAlign="Center">                            
<HeaderStyle HorizontalAlign="Center"></HeaderStyle>

                            <CellStyle Font-Names="Tahoma" Font-Size="8pt">
                            </CellStyle>
                        </dx:GridViewDataTextColumn>
                        <dx:GridViewDataTextColumn VisibleIndex="3" Caption="PART NAME" FieldName="PartName" Width="180px" HeaderStyle-HorizontalAlign="Center">                            
<HeaderStyle HorizontalAlign="Center"></HeaderStyle>

                            <CellStyle Font-Names="Tahoma" Font-Size="8pt">
                            </CellStyle>
                        </dx:GridViewDataTextColumn>
                        <dx:GridViewDataTextColumn VisibleIndex="4" Caption="KANBAN CLS" FieldName="KanbanCls" Width="60px" HeaderStyle-HorizontalAlign="Center" CellStyle-HorizontalAlign="Center">                            
<HeaderStyle HorizontalAlign="Center"></HeaderStyle>

                            <CellStyle Font-Names="Tahoma" Font-Size="8pt">
                            </CellStyle>
                        </dx:GridViewDataTextColumn>
                        <dx:GridViewDataTextColumn VisibleIndex="5" Caption="UOM" FieldName="UnitDesc" Width="40px" HeaderStyle-HorizontalAlign="Center" CellStyle-HorizontalAlign="Center">                            
<HeaderStyle HorizontalAlign="Center"></HeaderStyle>

                            <CellStyle Font-Names="Tahoma" Font-Size="8pt">
                            </CellStyle>
                        </dx:GridViewDataTextColumn>
                        <dx:GridViewDataTextColumn VisibleIndex="6" Caption="MOQ" FieldName="MinOrderQty" Width="70px" HeaderStyle-HorizontalAlign="Center">                            
<HeaderStyle HorizontalAlign="Center"></HeaderStyle>

                            <CellStyle Font-Names="Tahoma" Font-Size="8pt">
                            </CellStyle>
                        </dx:GridViewDataTextColumn>
                        
                        <dx:GridViewDataTextColumn VisibleIndex="7" Caption="QTY/BOX" FieldName="QtyBox" Width="70px" HeaderStyle-HorizontalAlign="Center">                            
<HeaderStyle HorizontalAlign="Center"></HeaderStyle>

                            <CellStyle Font-Names="Tahoma" Font-Size="8pt">
                            </CellStyle>
                        </dx:GridViewDataTextColumn>  
                        <dx:GridViewDataTextColumn VisibleIndex="8" Caption="MAKER" FieldName="Maker" Width="100px" HeaderStyle-HorizontalAlign="Center">                           
<HeaderStyle HorizontalAlign="Center"></HeaderStyle>

                            <CellStyle Font-Names="Tahoma" Font-Size="8pt">
                            </CellStyle>
                        </dx:GridViewDataTextColumn>

                        <dx:GridViewDataTextColumn VisibleIndex="9" Caption="PO NO." FieldName="PONo" Width="120px" HeaderStyle-HorizontalAlign="Center" Visible="False" >                            
<HeaderStyle HorizontalAlign="Center"></HeaderStyle>

                            <CellStyle Font-Names="Tahoma" Font-Size="8pt">
                            </CellStyle>
                        </dx:GridViewDataTextColumn>

                        <dx:GridViewDataTextColumn VisibleIndex="10" Caption="TOTAL FIRM QTY" FieldName="POQty" Width="0px" HeaderStyle-HorizontalAlign="Center">
                            <PropertiesTextEdit DisplayFormatString="{0:n0}">
                                <MaskSettings ErrorText="Please input valid value !"
                                Mask="<0..999999g>" IncludeLiterals="DecimalSymbol" />
                                <ValidationSettings ErrorDisplayMode="None"></ValidationSettings>
                            </PropertiesTextEdit>

<HeaderStyle HorizontalAlign="Center"></HeaderStyle>

                            <CellStyle Font-Names="Tahoma" Font-Size="8pt">
                            </CellStyle>
                        </dx:GridViewDataTextColumn>

                        <dx:GridViewDataTextColumn VisibleIndex="11" Caption="CURR" FieldName="CurrDesc" Width="0px" HeaderStyle-HorizontalAlign="Center" CellStyle-HorizontalAlign="Center" Visible="True">
                            <HeaderStyle HorizontalAlign="Center"></HeaderStyle>
                            <CellStyle Font-Names="Tahoma" Font-Size="8pt">
                            </CellStyle>
                        </dx:GridViewDataTextColumn>  
                        <dx:GridViewDataTextColumn VisibleIndex="12" Caption="PRICE" FieldName="Price" Width="0px" HeaderStyle-HorizontalAlign="Center" Visible="True">
                            <PropertiesTextEdit DisplayFormatString="{0:n2}">
                                <MaskSettings ErrorText="Please input valid value !"
                                Mask="<0..999999g>" IncludeLiterals="DecimalSymbol" />
                            </PropertiesTextEdit>                            

<HeaderStyle HorizontalAlign="Center"></HeaderStyle>

                            <CellStyle Font-Names="Tahoma" Font-Size="8pt">
                            </CellStyle>
                        </dx:GridViewDataTextColumn>  
                        <dx:GridViewDataTextColumn VisibleIndex="13" Caption="AMOUNT" FieldName="Amount" Width="0px" HeaderStyle-HorizontalAlign="Center">
                            <PropertiesTextEdit DisplayFormatString="{0:n2}">
                                <MaskSettings ErrorText="Please input valid value !"
                                Mask="<0..999999g>" IncludeLiterals="DecimalSymbol" />
                            </PropertiesTextEdit>                            

<HeaderStyle HorizontalAlign="Center"></HeaderStyle>

                            <CellStyle Font-Names="Tahoma" Font-Size="8pt">
                            </CellStyle>
                        </dx:GridViewDataTextColumn>
                        <dx:GridViewDataTextColumn VisibleIndex="14" Caption="FORECAST N+1" FieldName="ForecastN1" Width="80px" HeaderStyle-HorizontalAlign="Center">
<HeaderStyle HorizontalAlign="Center"></HeaderStyle>

                            <CellStyle Font-Names="Tahoma" Font-Size="8pt">
                            </CellStyle>
                        </dx:GridViewDataTextColumn>
                        <dx:GridViewDataTextColumn VisibleIndex="15" Caption="FORECAST N+2" FieldName="ForecastN2" Width="80px" HeaderStyle-HorizontalAlign="Center">
<HeaderStyle HorizontalAlign="Center"></HeaderStyle>

                            <CellStyle Font-Names="Tahoma" Font-Size="8pt">
                            </CellStyle>
                        </dx:GridViewDataTextColumn>
                        <dx:GridViewDataTextColumn VisibleIndex="16" Caption="FORECAST N+3" FieldName="ForecastN3" Width="80px" HeaderStyle-HorizontalAlign="Center">
<HeaderStyle HorizontalAlign="Center"></HeaderStyle>

                            <CellStyle Font-Names="Tahoma" Font-Size="8pt">
                            </CellStyle>
                        </dx:GridViewDataTextColumn>

                        <dx:GridViewBandColumn Caption="E.T.A SCHEDULE (BASED ON FIRM ORDER)" VisibleIndex="17" HeaderStyle-HorizontalAlign="Center">
                            <Columns>
                                <dx:GridViewDataTextColumn VisibleIndex="18" Caption="1" Width="70px" FieldName="DeliveryD1" HeaderStyle-HorizontalAlign="Center">                                    
                                    <PropertiesTextEdit DisplayFormatString="{0:n0}">
                                        <MaskSettings ErrorText="Please input valid value !"
                                        Mask="<0..999999g>" IncludeLiterals="DecimalSymbol" />
                                        <ValidationSettings ErrorDisplayMode="None"></ValidationSettings>
                                    </PropertiesTextEdit>

<HeaderStyle HorizontalAlign="Center"></HeaderStyle>

                                    <CellStyle Font-Names="Tahoma" Font-Size="8pt">
                                    </CellStyle>
                                </dx:GridViewDataTextColumn>
                                <dx:GridViewDataTextColumn VisibleIndex="19" Caption="2" Width="70px" FieldName="DeliveryD2" HeaderStyle-HorizontalAlign="Center">
                                    <PropertiesTextEdit DisplayFormatString="{0:n0}">
                                        <MaskSettings ErrorText="Please input valid value !"
                                        Mask="<0..999999g>" IncludeLiterals="DecimalSymbol" />
                                        <ValidationSettings ErrorDisplayMode="None"></ValidationSettings>
                                    </PropertiesTextEdit>

<HeaderStyle HorizontalAlign="Center"></HeaderStyle>

                                    <CellStyle Font-Names="Tahoma" Font-Size="8pt">
                                    </CellStyle>
                                </dx:GridViewDataTextColumn>                               
                                <dx:GridViewDataTextColumn VisibleIndex="20" Caption="3" Width="70px" FieldName="DeliveryD3" HeaderStyle-HorizontalAlign="Center">
                                    <PropertiesTextEdit DisplayFormatString="{0:n0}">
                                        <MaskSettings ErrorText="Please input valid value !"
                                        Mask="<0..999999g>" IncludeLiterals="DecimalSymbol" />
                                        <ValidationSettings ErrorDisplayMode="None"></ValidationSettings>
                                    </PropertiesTextEdit>

<HeaderStyle HorizontalAlign="Center"></HeaderStyle>

                                    <CellStyle Font-Names="Tahoma" Font-Size="8pt">
                                    </CellStyle>
                                </dx:GridViewDataTextColumn>                                
                                <dx:GridViewDataTextColumn VisibleIndex="21" Caption="4" Width="70px" FieldName="DeliveryD4" HeaderStyle-HorizontalAlign="Center">
                                    <PropertiesTextEdit DisplayFormatString="{0:n0}">
                                        <MaskSettings ErrorText="Please input valid value !"
                                        Mask="<0..999999g>" IncludeLiterals="DecimalSymbol" />
                                        <ValidationSettings ErrorDisplayMode="None"></ValidationSettings>
                                    </PropertiesTextEdit>

<HeaderStyle HorizontalAlign="Center"></HeaderStyle>

                                    <CellStyle Font-Names="Tahoma" Font-Size="8pt">
                                    </CellStyle>
                                </dx:GridViewDataTextColumn>                                
                                <dx:GridViewDataTextColumn VisibleIndex="22" Caption="5" Width="70px" FieldName="DeliveryD5" HeaderStyle-HorizontalAlign="Center">
                                    <PropertiesTextEdit DisplayFormatString="{0:n0}">
                                        <MaskSettings ErrorText="Please input valid value !"
                                        Mask="<0..999999g>" IncludeLiterals="DecimalSymbol" />
                                        <ValidationSettings ErrorDisplayMode="None"></ValidationSettings>
                                    </PropertiesTextEdit>

<HeaderStyle HorizontalAlign="Center"></HeaderStyle>

                                    <CellStyle Font-Names="Tahoma" Font-Size="8pt">
                                    </CellStyle>
                                </dx:GridViewDataTextColumn>                                
                                <dx:GridViewDataTextColumn VisibleIndex="23" Caption="6" Width="70px" FieldName="DeliveryD6" HeaderStyle-HorizontalAlign="Center">
                                    <PropertiesTextEdit DisplayFormatString="{0:n0}">
                                        <MaskSettings ErrorText="Please input valid value !"
                                        Mask="<0..999999g>" IncludeLiterals="DecimalSymbol" />
                                        <ValidationSettings ErrorDisplayMode="None"></ValidationSettings>
                                    </PropertiesTextEdit>

<HeaderStyle HorizontalAlign="Center"></HeaderStyle>

                                    <CellStyle Font-Names="Tahoma" Font-Size="8pt">
                                    </CellStyle>
                                </dx:GridViewDataTextColumn>                                
                                <dx:GridViewDataTextColumn VisibleIndex="24" Caption="7" Width="70px" FieldName="DeliveryD7" HeaderStyle-HorizontalAlign="Center">
                                    <PropertiesTextEdit DisplayFormatString="{0:n0}">
                                        <MaskSettings ErrorText="Please input valid value !"
                                        Mask="<0..999999g>" IncludeLiterals="DecimalSymbol" />
                                        <ValidationSettings ErrorDisplayMode="None"></ValidationSettings>
                                    </PropertiesTextEdit>

<HeaderStyle HorizontalAlign="Center"></HeaderStyle>

                                    <CellStyle Font-Names="Tahoma" Font-Size="8pt">
                                    </CellStyle>
                                </dx:GridViewDataTextColumn>                               
                                <dx:GridViewDataTextColumn VisibleIndex="25" Caption="8" Width="70px" FieldName="DeliveryD8" HeaderStyle-HorizontalAlign="Center">
                                    <PropertiesTextEdit DisplayFormatString="{0:n0}">
                                        <MaskSettings ErrorText="Please input valid value !"
                                        Mask="<0..999999g>" IncludeLiterals="DecimalSymbol" />
                                        <ValidationSettings ErrorDisplayMode="None"></ValidationSettings>
                                    </PropertiesTextEdit>

<HeaderStyle HorizontalAlign="Center"></HeaderStyle>

                                    <CellStyle Font-Names="Tahoma" Font-Size="8pt">
                                    </CellStyle>
                                </dx:GridViewDataTextColumn>                                
                                <dx:GridViewDataTextColumn VisibleIndex="26" Caption="9" Width="70px" FieldName="DeliveryD9" HeaderStyle-HorizontalAlign="Center">
                                    <PropertiesTextEdit DisplayFormatString="{0:n0}">
                                        <MaskSettings ErrorText="Please input valid value !"
                                        Mask="<0..999999g>" IncludeLiterals="DecimalSymbol" />
                                        <ValidationSettings ErrorDisplayMode="None"></ValidationSettings>
                                    </PropertiesTextEdit>

<HeaderStyle HorizontalAlign="Center"></HeaderStyle>

                                    <CellStyle Font-Names="Tahoma" Font-Size="8pt">
                                    </CellStyle>
                                </dx:GridViewDataTextColumn>                                
                                <dx:GridViewDataTextColumn VisibleIndex="27" Caption="10" Width="70px" FieldName="DeliveryD10" HeaderStyle-HorizontalAlign="Center">
                                    <PropertiesTextEdit DisplayFormatString="{0:n0}">
                                        <MaskSettings ErrorText="Please input valid value !"
                                        Mask="<0..999999g>" IncludeLiterals="DecimalSymbol" />
                                        <ValidationSettings ErrorDisplayMode="None"></ValidationSettings>
                                    </PropertiesTextEdit>

<HeaderStyle HorizontalAlign="Center"></HeaderStyle>

                                    <CellStyle Font-Names="Tahoma" Font-Size="8pt">
                                    </CellStyle>
                                </dx:GridViewDataTextColumn>                                
                                <dx:GridViewDataTextColumn VisibleIndex="28" Caption="11" Width="70px" FieldName="DeliveryD11" HeaderStyle-HorizontalAlign="Center">
                                    <PropertiesTextEdit DisplayFormatString="{0:n0}">
                                        <MaskSettings ErrorText="Please input valid value !"
                                        Mask="<0..999999g>" IncludeLiterals="DecimalSymbol" />
                                        <ValidationSettings ErrorDisplayMode="None"></ValidationSettings>
                                    </PropertiesTextEdit>

<HeaderStyle HorizontalAlign="Center"></HeaderStyle>

                                    <CellStyle Font-Names="Tahoma" Font-Size="8pt">
                                    </CellStyle>
                                </dx:GridViewDataTextColumn>                                
                                <dx:GridViewDataTextColumn VisibleIndex="29" Caption="12" Width="70px" FieldName="DeliveryD12" HeaderStyle-HorizontalAlign="Center">
                                    <PropertiesTextEdit DisplayFormatString="{0:n0}">
                                        <MaskSettings ErrorText="Please input valid value !"
                                        Mask="<0..999999g>" IncludeLiterals="DecimalSymbol" />
                                        <ValidationSettings ErrorDisplayMode="None"></ValidationSettings>
                                    </PropertiesTextEdit>

<HeaderStyle HorizontalAlign="Center"></HeaderStyle>

                                    <CellStyle Font-Names="Tahoma" Font-Size="8pt">
                                    </CellStyle>
                                </dx:GridViewDataTextColumn>                                
                                <dx:GridViewDataTextColumn VisibleIndex="30" Caption="13" Width="70px" FieldName="DeliveryD13" HeaderStyle-HorizontalAlign="Center">
                                    <PropertiesTextEdit DisplayFormatString="{0:n0}">
                                        <MaskSettings ErrorText="Please input valid value !"
                                        Mask="<0..999999g>" IncludeLiterals="DecimalSymbol" />
                                        <ValidationSettings ErrorDisplayMode="None"></ValidationSettings>
                                    </PropertiesTextEdit>

<HeaderStyle HorizontalAlign="Center"></HeaderStyle>

                                    <CellStyle Font-Names="Tahoma" Font-Size="8pt">
                                    </CellStyle>
                                </dx:GridViewDataTextColumn>                                
                                <dx:GridViewDataTextColumn VisibleIndex="31" Caption="14" Width="70px" FieldName="DeliveryD14" HeaderStyle-HorizontalAlign="Center">
                                    <PropertiesTextEdit DisplayFormatString="{0:n0}">
                                        <MaskSettings ErrorText="Please input valid value !"
                                        Mask="<0..999999g>" IncludeLiterals="DecimalSymbol" />
                                        <ValidationSettings ErrorDisplayMode="None"></ValidationSettings>
                                    </PropertiesTextEdit>

<HeaderStyle HorizontalAlign="Center"></HeaderStyle>

                                    <CellStyle Font-Names="Tahoma" Font-Size="8pt">
                                    </CellStyle>
                                </dx:GridViewDataTextColumn>                                
                                <dx:GridViewDataTextColumn VisibleIndex="32" Caption="15" Width="70px" FieldName="DeliveryD15" HeaderStyle-HorizontalAlign="Center">
                                    <PropertiesTextEdit DisplayFormatString="{0:n0}">
                                        <MaskSettings ErrorText="Please input valid value !"
                                        Mask="<0..999999g>" IncludeLiterals="DecimalSymbol" />
                                        <ValidationSettings ErrorDisplayMode="None"></ValidationSettings>
                                    </PropertiesTextEdit>

<HeaderStyle HorizontalAlign="Center"></HeaderStyle>

                                    <CellStyle Font-Names="Tahoma" Font-Size="8pt">
                                    </CellStyle>
                                </dx:GridViewDataTextColumn>                                
                                <dx:GridViewDataTextColumn VisibleIndex="33" Caption="16" Width="70px" FieldName="DeliveryD16" HeaderStyle-HorizontalAlign="Center">
                                    <PropertiesTextEdit DisplayFormatString="{0:n0}">
                                        <MaskSettings ErrorText="Please input valid value !"
                                        Mask="<0..999999g>" IncludeLiterals="DecimalSymbol" />
                                        <ValidationSettings ErrorDisplayMode="None"></ValidationSettings>
                                    </PropertiesTextEdit>

<HeaderStyle HorizontalAlign="Center"></HeaderStyle>

                                    <CellStyle Font-Names="Tahoma" Font-Size="8pt">
                                    </CellStyle>
                                </dx:GridViewDataTextColumn>                                
                                <dx:GridViewDataTextColumn VisibleIndex="34" Caption="17" Width="70px" FieldName="DeliveryD17" HeaderStyle-HorizontalAlign="Center">
                                    <PropertiesTextEdit DisplayFormatString="{0:n0}">
                                        <MaskSettings ErrorText="Please input valid value !"
                                        Mask="<0..999999g>" IncludeLiterals="DecimalSymbol" />
                                        <ValidationSettings ErrorDisplayMode="None"></ValidationSettings>
                                    </PropertiesTextEdit>

<HeaderStyle HorizontalAlign="Center"></HeaderStyle>

                                    <CellStyle Font-Names="Tahoma" Font-Size="8pt">
                                    </CellStyle>
                                </dx:GridViewDataTextColumn>                                
                                <dx:GridViewDataTextColumn VisibleIndex="35" Caption="18" Width="70px" FieldName="DeliveryD18" HeaderStyle-HorizontalAlign="Center">
                                    <PropertiesTextEdit DisplayFormatString="{0:n0}">
                                        <MaskSettings ErrorText="Please input valid value !"
                                        Mask="<0..999999g>" IncludeLiterals="DecimalSymbol" />
                                        <ValidationSettings ErrorDisplayMode="None"></ValidationSettings>
                                    </PropertiesTextEdit>

<HeaderStyle HorizontalAlign="Center"></HeaderStyle>

                                    <CellStyle Font-Names="Tahoma" Font-Size="8pt">
                                    </CellStyle>
                                </dx:GridViewDataTextColumn>                                
                                <dx:GridViewDataTextColumn VisibleIndex="36" Caption="19" Width="70px" FieldName="DeliveryD19" HeaderStyle-HorizontalAlign="Center">
                                    <PropertiesTextEdit DisplayFormatString="{0:n0}">
                                        <MaskSettings ErrorText="Please input valid value !"
                                        Mask="<0..999999g>" IncludeLiterals="DecimalSymbol" />
                                        <ValidationSettings ErrorDisplayMode="None"></ValidationSettings>
                                    </PropertiesTextEdit>

<HeaderStyle HorizontalAlign="Center"></HeaderStyle>

                                    <CellStyle Font-Names="Tahoma" Font-Size="8pt">
                                    </CellStyle>
                                </dx:GridViewDataTextColumn>                                
                                <dx:GridViewDataTextColumn VisibleIndex="37" Caption="20" Width="70px" FieldName="DeliveryD20" HeaderStyle-HorizontalAlign="Center">
                                    <PropertiesTextEdit DisplayFormatString="{0:n0}">
                                        <MaskSettings ErrorText="Please input valid value !"
                                        Mask="<0..999999g>" IncludeLiterals="DecimalSymbol" />
                                        <ValidationSettings ErrorDisplayMode="None"></ValidationSettings>
                                    </PropertiesTextEdit>

<HeaderStyle HorizontalAlign="Center"></HeaderStyle>

                                    <CellStyle Font-Names="Tahoma" Font-Size="8pt">
                                    </CellStyle>
                                </dx:GridViewDataTextColumn>                                
                                <dx:GridViewDataTextColumn VisibleIndex="38" Caption="21" Width="70px" FieldName="DeliveryD21" HeaderStyle-HorizontalAlign="Center">
                                    <PropertiesTextEdit DisplayFormatString="{0:n0}">
                                        <MaskSettings ErrorText="Please input valid value !"
                                        Mask="<0..999999g>" IncludeLiterals="DecimalSymbol" />
                                        <ValidationSettings ErrorDisplayMode="None"></ValidationSettings>
                                    </PropertiesTextEdit>

<HeaderStyle HorizontalAlign="Center"></HeaderStyle>

                                    <CellStyle Font-Names="Tahoma" Font-Size="8pt">
                                    </CellStyle>
                                </dx:GridViewDataTextColumn>                                
                                <dx:GridViewDataTextColumn VisibleIndex="39" Caption="22" Width="70px" FieldName="DeliveryD22" HeaderStyle-HorizontalAlign="Center">
                                    <PropertiesTextEdit DisplayFormatString="{0:n0}">
                                        <MaskSettings ErrorText="Please input valid value !"
                                        Mask="<0..999999g>" IncludeLiterals="DecimalSymbol" />
                                        <ValidationSettings ErrorDisplayMode="None"></ValidationSettings>
                                    </PropertiesTextEdit>

<HeaderStyle HorizontalAlign="Center"></HeaderStyle>

                                    <CellStyle Font-Names="Tahoma" Font-Size="8pt">
                                    </CellStyle>
                                </dx:GridViewDataTextColumn>                                
                                <dx:GridViewDataTextColumn VisibleIndex="40" Caption="23" Width="70px" FieldName="DeliveryD23" HeaderStyle-HorizontalAlign="Center">
                                    <PropertiesTextEdit DisplayFormatString="{0:n0}">
                                        <MaskSettings ErrorText="Please input valid value !"
                                        Mask="<0..999999g>" IncludeLiterals="DecimalSymbol" />
                                        <ValidationSettings ErrorDisplayMode="None"></ValidationSettings>
                                    </PropertiesTextEdit>

<HeaderStyle HorizontalAlign="Center"></HeaderStyle>

                                    <CellStyle Font-Names="Tahoma" Font-Size="8pt">
                                    </CellStyle>
                                </dx:GridViewDataTextColumn>                                
                                <dx:GridViewDataTextColumn VisibleIndex="41" Caption="24" Width="70px" FieldName="DeliveryD24" HeaderStyle-HorizontalAlign="Center">
                                    <PropertiesTextEdit DisplayFormatString="{0:n0}">
                                        <MaskSettings ErrorText="Please input valid value !"
                                        Mask="<0..999999g>" IncludeLiterals="DecimalSymbol" />
                                        <ValidationSettings ErrorDisplayMode="None"></ValidationSettings>
                                    </PropertiesTextEdit>

<HeaderStyle HorizontalAlign="Center"></HeaderStyle>

                                    <CellStyle Font-Names="Tahoma" Font-Size="8pt">
                                    </CellStyle>
                                </dx:GridViewDataTextColumn>                               
                                <dx:GridViewDataTextColumn VisibleIndex="42" Caption="25" Width="70px" FieldName="DeliveryD25" HeaderStyle-HorizontalAlign="Center">
                                    <PropertiesTextEdit DisplayFormatString="{0:n0}">
                                        <MaskSettings ErrorText="Please input valid value !"
                                        Mask="<0..999999g>" IncludeLiterals="DecimalSymbol" />
                                        <ValidationSettings ErrorDisplayMode="None"></ValidationSettings>
                                    </PropertiesTextEdit>

<HeaderStyle HorizontalAlign="Center"></HeaderStyle>

                                    <CellStyle Font-Names="Tahoma" Font-Size="8pt">
                                    </CellStyle>
                                </dx:GridViewDataTextColumn>                                
                                <dx:GridViewDataTextColumn VisibleIndex="43" Caption="26" Width="70px" FieldName="DeliveryD26" HeaderStyle-HorizontalAlign="Center">
                                    <PropertiesTextEdit DisplayFormatString="{0:n0}">
                                        <MaskSettings ErrorText="Please input valid value !"
                                        Mask="<0..999999g>" IncludeLiterals="DecimalSymbol" />
                                        <ValidationSettings ErrorDisplayMode="None"></ValidationSettings>
                                    </PropertiesTextEdit>

<HeaderStyle HorizontalAlign="Center"></HeaderStyle>

                                    <CellStyle Font-Names="Tahoma" Font-Size="8pt">
                                    </CellStyle>
                                </dx:GridViewDataTextColumn>                                
                                <dx:GridViewDataTextColumn VisibleIndex="44" Caption="27" Width="70px" FieldName="DeliveryD27" HeaderStyle-HorizontalAlign="Center">
                                    <PropertiesTextEdit DisplayFormatString="{0:n0}">
                                        <MaskSettings ErrorText="Please input valid value !"
                                        Mask="<0..999999g>" IncludeLiterals="DecimalSymbol" />
                                        <ValidationSettings ErrorDisplayMode="None"></ValidationSettings>
                                    </PropertiesTextEdit>

<HeaderStyle HorizontalAlign="Center"></HeaderStyle>

                                    <CellStyle Font-Names="Tahoma" Font-Size="8pt">
                                    </CellStyle>
                                </dx:GridViewDataTextColumn>                                
                                <dx:GridViewDataTextColumn VisibleIndex="45" Caption="28" Width="70px" FieldName="DeliveryD28" HeaderStyle-HorizontalAlign="Center">
                                    <PropertiesTextEdit DisplayFormatString="{0:n0}">
                                        <MaskSettings ErrorText="Please input valid value !"
                                        Mask="<0..999999g>" IncludeLiterals="DecimalSymbol" />
                                        <ValidationSettings ErrorDisplayMode="None"></ValidationSettings>
                                    </PropertiesTextEdit>

<HeaderStyle HorizontalAlign="Center"></HeaderStyle>

                                    <CellStyle Font-Names="Tahoma" Font-Size="8pt">
                                    </CellStyle>
                                </dx:GridViewDataTextColumn>                               
                                <dx:GridViewDataTextColumn VisibleIndex="46" Caption="29" Width="70px" FieldName="DeliveryD29" HeaderStyle-HorizontalAlign="Center">
                                    <PropertiesTextEdit DisplayFormatString="{0:n0}">
                                        <MaskSettings ErrorText="Please input valid value !"
                                        Mask="<0..999999g>" IncludeLiterals="DecimalSymbol" />
                                        <ValidationSettings ErrorDisplayMode="None"></ValidationSettings>
                                    </PropertiesTextEdit>

<HeaderStyle HorizontalAlign="Center"></HeaderStyle>

                                    <CellStyle Font-Names="Tahoma" Font-Size="8pt">
                                    </CellStyle>
                                </dx:GridViewDataTextColumn>                                
                                <dx:GridViewDataTextColumn VisibleIndex="47" Caption="30" Width="70px" FieldName="DeliveryD30" HeaderStyle-HorizontalAlign="Center">
                                    <PropertiesTextEdit DisplayFormatString="{0:n0}">
                                        <MaskSettings ErrorText="Please input valid value !"
                                        Mask="<0..999999g>" IncludeLiterals="DecimalSymbol" />
                                        <ValidationSettings ErrorDisplayMode="None"></ValidationSettings>
                                    </PropertiesTextEdit>
                                    <HeaderStyle HorizontalAlign="Center"></HeaderStyle>
                                    <CellStyle Font-Names="Tahoma" Font-Size="8pt">
                                    </CellStyle>
                                </dx:GridViewDataTextColumn>                                
                                <dx:GridViewDataTextColumn VisibleIndex="48" Caption="31" Width="70px" FieldName="DeliveryD31">
                                    <PropertiesTextEdit DisplayFormatString="{0:n0}">
                                        <MaskSettings ErrorText="Please input valid value !"
                                        Mask="<0..999999g>" IncludeLiterals="DecimalSymbol" />
                                        <ValidationSettings ErrorDisplayMode="None"></ValidationSettings>
                                    </PropertiesTextEdit>
                                    <HeaderStyle HorizontalAlign="Center"></HeaderStyle>
                                    <CellStyle Font-Names="Tahoma" Font-Size="8pt">
                                    </CellStyle>
                                </dx:GridViewDataTextColumn>
                            </Columns>
                            <HeaderStyle HorizontalAlign="Center" />
                        </dx:GridViewBandColumn>
                        <dx:GridViewDataTextColumn VisibleIndex="49" Caption="countPartNo" FieldName="countPartNo" Width="0px">
                        </dx:GridViewDataTextColumn>
                        <dx:GridViewDataTextColumn VisibleIndex="50" Caption="SupplierID" FieldName="SupplierID" Width="0px">
                        </dx:GridViewDataTextColumn>
                        <dx:GridViewDataTextColumn VisibleIndex="51" Caption="CurrCls" FieldName="CurrCls" Width="0px">
                        </dx:GridViewDataTextColumn>
                        <dx:GridViewDataTextColumn VisibleIndex="52" Caption="UnitCls" FieldName="UnitCls" Width="0px">
                        </dx:GridViewDataTextColumn>
                        <dx:GridViewDataTextColumn VisibleIndex="53" Caption="PODeliveryBy" FieldName="PODeliveryBy" Width="0px">
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
    </table>

    <div style="height:8px;"></div>

    <table style="width:100%;">
        <tr>
            <td align="left" valign="middle" style="height:20px; width:60px;">
                <dx:ASPxLabel ID="ASPxLabel3" runat="server" Text="PO NO. (ADDITIONAL)"
                    Font-Names="Tahoma" Font-Size="8pt">
                </dx:ASPxLabel>
            </td>
            <td align="left" valign="middle" height="20px" width="180px">
                <dx:ASPxTextBox ID="ASPxTextBox1" runat="server" Width="180px" 
                    ClientInstanceName="txtPONoReg" Font-Names="Tahoma" Font-Size="8pt"
                    MaxLength="20" onkeypress="return singlequote(event)" Height="20px">
                    <ClientSideEvents LostFocus="function(s, e) { 
	                    lblInfo.SetText('');
                    }" />
                </dx:ASPxTextBox>
            </td>
        </tr>
        <tr>
            <td align="left" valign="top" height="220" colspan="2">
                <dx:ASPxGridView ID="ASPxGridView1" runat="server" Width="100%"
                    Font-Names="Tahoma" KeyFieldName="PartNo;PONo"
                    AutoGenerateColumns="False"
                    ClientInstanceName="grid" Font-Size="8pt">
                    <ClientSideEvents Init="OnInit" />                    
                    <Columns>
                        <dx:GridViewDataTextColumn VisibleIndex="0" Caption="NO." FieldName="NoUrut" Width="30px" HeaderStyle-HorizontalAlign="Center">                            
<HeaderStyle HorizontalAlign="Center"></HeaderStyle>

                            <CellStyle Font-Names="Tahoma" Font-Size="8pt">
                            </CellStyle>
                        </dx:GridViewDataTextColumn>
                        <dx:GridViewDataTextColumn VisibleIndex="1" Caption="ORDER NO." FieldName="OrderNo" Width="90px" HeaderStyle-HorizontalAlign="Center">                            
<HeaderStyle HorizontalAlign="Center"></HeaderStyle>

                            <CellStyle Font-Names="Tahoma" Font-Size="8pt">
                            </CellStyle>
                        </dx:GridViewDataTextColumn>
                        <dx:GridViewDataTextColumn VisibleIndex="2" Caption="PART NO." FieldName="PartNo" Width="90px" HeaderStyle-HorizontalAlign="Center">                            
<HeaderStyle HorizontalAlign="Center"></HeaderStyle>

                            <CellStyle Font-Names="Tahoma" Font-Size="8pt">
                            </CellStyle>
                        </dx:GridViewDataTextColumn>
                        <dx:GridViewDataTextColumn VisibleIndex="3" Caption="PART NAME" FieldName="PartName" Width="180px" HeaderStyle-HorizontalAlign="Center">                            
<HeaderStyle HorizontalAlign="Center"></HeaderStyle>

                            <CellStyle Font-Names="Tahoma" Font-Size="8pt">
                            </CellStyle>
                        </dx:GridViewDataTextColumn>
                        <dx:GridViewDataTextColumn VisibleIndex="4" Caption="KANBAN CLS" FieldName="KanbanCls" Width="60px" HeaderStyle-HorizontalAlign="Center" CellStyle-HorizontalAlign="Center">                            
<HeaderStyle HorizontalAlign="Center"></HeaderStyle>

                            <CellStyle Font-Names="Tahoma" Font-Size="8pt">
                            </CellStyle>
                        </dx:GridViewDataTextColumn>
                        <dx:GridViewDataTextColumn VisibleIndex="5" Caption="UOM" FieldName="UnitDesc" Width="40px" HeaderStyle-HorizontalAlign="Center" CellStyle-HorizontalAlign="Center">                            
<HeaderStyle HorizontalAlign="Center"></HeaderStyle>

                            <CellStyle Font-Names="Tahoma" Font-Size="8pt">
                            </CellStyle>
                        </dx:GridViewDataTextColumn>
                        <dx:GridViewDataTextColumn VisibleIndex="6" Caption="MOQ" FieldName="MinOrderQty" Width="70px" HeaderStyle-HorizontalAlign="Center">                            
<HeaderStyle HorizontalAlign="Center"></HeaderStyle>

                            <CellStyle Font-Names="Tahoma" Font-Size="8pt">
                            </CellStyle>
                        </dx:GridViewDataTextColumn>
                        
                        <dx:GridViewDataTextColumn VisibleIndex="7" Caption="QTY/BOX" FieldName="QtyBox" Width="70px" HeaderStyle-HorizontalAlign="Center">                            
<HeaderStyle HorizontalAlign="Center"></HeaderStyle>

                            <CellStyle Font-Names="Tahoma" Font-Size="8pt">
                            </CellStyle>
                        </dx:GridViewDataTextColumn>  
                        <dx:GridViewDataTextColumn VisibleIndex="8" Caption="MAKER" FieldName="Maker" Width="100px" HeaderStyle-HorizontalAlign="Center">                           
<HeaderStyle HorizontalAlign="Center"></HeaderStyle>

                            <CellStyle Font-Names="Tahoma" Font-Size="8pt">
                            </CellStyle>
                        </dx:GridViewDataTextColumn>

                        <dx:GridViewDataTextColumn VisibleIndex="9" Caption="PO NO." FieldName="PONo" Width="120px" HeaderStyle-HorizontalAlign="Center" Visible="False" >                            
<HeaderStyle HorizontalAlign="Center"></HeaderStyle>

                            <CellStyle Font-Names="Tahoma" Font-Size="8pt">
                            </CellStyle>
                        </dx:GridViewDataTextColumn>

                        <dx:GridViewDataTextColumn VisibleIndex="10" Caption="TOTAL FIRM QTY" FieldName="POQty" Width="0px" HeaderStyle-HorizontalAlign="Center">
                            <PropertiesTextEdit DisplayFormatString="{0:n0}">
                                <MaskSettings ErrorText="Please input valid value !"
                                Mask="<0..999999g>" IncludeLiterals="DecimalSymbol" />
                                <ValidationSettings ErrorDisplayMode="None"></ValidationSettings>
                            </PropertiesTextEdit>

<HeaderStyle HorizontalAlign="Center"></HeaderStyle>

                            <CellStyle Font-Names="Tahoma" Font-Size="8pt">
                            </CellStyle>
                        </dx:GridViewDataTextColumn>

                        <dx:GridViewDataTextColumn VisibleIndex="11" Caption="CURR" FieldName="CurrDesc" Width="0px" HeaderStyle-HorizontalAlign="Center" CellStyle-HorizontalAlign="Center" Visible="True">
                            <HeaderStyle HorizontalAlign="Center"></HeaderStyle>
                            <CellStyle Font-Names="Tahoma" Font-Size="8pt">
                            </CellStyle>
                        </dx:GridViewDataTextColumn>  
                        <dx:GridViewDataTextColumn VisibleIndex="12" Caption="PRICE" FieldName="Price" Width="0px" HeaderStyle-HorizontalAlign="Center" Visible="True">
                            <PropertiesTextEdit DisplayFormatString="{0:n2}">
                                <MaskSettings ErrorText="Please input valid value !"
                                Mask="<0..999999g>" IncludeLiterals="DecimalSymbol" />
                            </PropertiesTextEdit>                            

<HeaderStyle HorizontalAlign="Center"></HeaderStyle>

                            <CellStyle Font-Names="Tahoma" Font-Size="8pt">
                            </CellStyle>
                        </dx:GridViewDataTextColumn>  
                        <dx:GridViewDataTextColumn VisibleIndex="13" Caption="AMOUNT" FieldName="Amount" Width="0px" HeaderStyle-HorizontalAlign="Center">
                            <PropertiesTextEdit DisplayFormatString="{0:n2}">
                                <MaskSettings ErrorText="Please input valid value !"
                                Mask="<0..999999g>" IncludeLiterals="DecimalSymbol" />
                            </PropertiesTextEdit>                            

<HeaderStyle HorizontalAlign="Center"></HeaderStyle>

                            <CellStyle Font-Names="Tahoma" Font-Size="8pt">
                            </CellStyle>
                        </dx:GridViewDataTextColumn>
                        <dx:GridViewDataTextColumn VisibleIndex="14" Caption="FORECAST N+1" FieldName="ForecastN1" Width="80px" HeaderStyle-HorizontalAlign="Center">
<HeaderStyle HorizontalAlign="Center"></HeaderStyle>

                            <CellStyle Font-Names="Tahoma" Font-Size="8pt">
                            </CellStyle>
                        </dx:GridViewDataTextColumn>
                        <dx:GridViewDataTextColumn VisibleIndex="15" Caption="FORECAST N+2" FieldName="ForecastN2" Width="80px" HeaderStyle-HorizontalAlign="Center">
<HeaderStyle HorizontalAlign="Center"></HeaderStyle>

                            <CellStyle Font-Names="Tahoma" Font-Size="8pt">
                            </CellStyle>
                        </dx:GridViewDataTextColumn>
                        <dx:GridViewDataTextColumn VisibleIndex="16" Caption="FORECAST N+3" FieldName="ForecastN3" Width="80px" HeaderStyle-HorizontalAlign="Center">
<HeaderStyle HorizontalAlign="Center"></HeaderStyle>

                            <CellStyle Font-Names="Tahoma" Font-Size="8pt">
                            </CellStyle>
                        </dx:GridViewDataTextColumn>

                        <dx:GridViewBandColumn Caption="E.T.A SCHEDULE (BASED ON FIRM ORDER)" VisibleIndex="17" HeaderStyle-HorizontalAlign="Center">
                            <Columns>
                                <dx:GridViewDataTextColumn VisibleIndex="18" Caption="1" Width="70px" FieldName="DeliveryD1" HeaderStyle-HorizontalAlign="Center">                                    
                                    <PropertiesTextEdit DisplayFormatString="{0:n0}">
                                        <MaskSettings ErrorText="Please input valid value !"
                                        Mask="<0..999999g>" IncludeLiterals="DecimalSymbol" />
                                        <ValidationSettings ErrorDisplayMode="None"></ValidationSettings>
                                    </PropertiesTextEdit>

<HeaderStyle HorizontalAlign="Center"></HeaderStyle>

                                    <CellStyle Font-Names="Tahoma" Font-Size="8pt">
                                    </CellStyle>
                                </dx:GridViewDataTextColumn>
                                <dx:GridViewDataTextColumn VisibleIndex="19" Caption="2" Width="70px" FieldName="DeliveryD2" HeaderStyle-HorizontalAlign="Center">
                                    <PropertiesTextEdit DisplayFormatString="{0:n0}">
                                        <MaskSettings ErrorText="Please input valid value !"
                                        Mask="<0..999999g>" IncludeLiterals="DecimalSymbol" />
                                        <ValidationSettings ErrorDisplayMode="None"></ValidationSettings>
                                    </PropertiesTextEdit>

<HeaderStyle HorizontalAlign="Center"></HeaderStyle>

                                    <CellStyle Font-Names="Tahoma" Font-Size="8pt">
                                    </CellStyle>
                                </dx:GridViewDataTextColumn>                               
                                <dx:GridViewDataTextColumn VisibleIndex="20" Caption="3" Width="70px" FieldName="DeliveryD3" HeaderStyle-HorizontalAlign="Center">
                                    <PropertiesTextEdit DisplayFormatString="{0:n0}">
                                        <MaskSettings ErrorText="Please input valid value !"
                                        Mask="<0..999999g>" IncludeLiterals="DecimalSymbol" />
                                        <ValidationSettings ErrorDisplayMode="None"></ValidationSettings>
                                    </PropertiesTextEdit>

<HeaderStyle HorizontalAlign="Center"></HeaderStyle>

                                    <CellStyle Font-Names="Tahoma" Font-Size="8pt">
                                    </CellStyle>
                                </dx:GridViewDataTextColumn>                                
                                <dx:GridViewDataTextColumn VisibleIndex="21" Caption="4" Width="70px" FieldName="DeliveryD4" HeaderStyle-HorizontalAlign="Center">
                                    <PropertiesTextEdit DisplayFormatString="{0:n0}">
                                        <MaskSettings ErrorText="Please input valid value !"
                                        Mask="<0..999999g>" IncludeLiterals="DecimalSymbol" />
                                        <ValidationSettings ErrorDisplayMode="None"></ValidationSettings>
                                    </PropertiesTextEdit>

<HeaderStyle HorizontalAlign="Center"></HeaderStyle>

                                    <CellStyle Font-Names="Tahoma" Font-Size="8pt">
                                    </CellStyle>
                                </dx:GridViewDataTextColumn>                                
                                <dx:GridViewDataTextColumn VisibleIndex="22" Caption="5" Width="70px" FieldName="DeliveryD5" HeaderStyle-HorizontalAlign="Center">
                                    <PropertiesTextEdit DisplayFormatString="{0:n0}">
                                        <MaskSettings ErrorText="Please input valid value !"
                                        Mask="<0..999999g>" IncludeLiterals="DecimalSymbol" />
                                        <ValidationSettings ErrorDisplayMode="None"></ValidationSettings>
                                    </PropertiesTextEdit>

<HeaderStyle HorizontalAlign="Center"></HeaderStyle>

                                    <CellStyle Font-Names="Tahoma" Font-Size="8pt">
                                    </CellStyle>
                                </dx:GridViewDataTextColumn>                                
                                <dx:GridViewDataTextColumn VisibleIndex="23" Caption="6" Width="70px" FieldName="DeliveryD6" HeaderStyle-HorizontalAlign="Center">
                                    <PropertiesTextEdit DisplayFormatString="{0:n0}">
                                        <MaskSettings ErrorText="Please input valid value !"
                                        Mask="<0..999999g>" IncludeLiterals="DecimalSymbol" />
                                        <ValidationSettings ErrorDisplayMode="None"></ValidationSettings>
                                    </PropertiesTextEdit>

<HeaderStyle HorizontalAlign="Center"></HeaderStyle>

                                    <CellStyle Font-Names="Tahoma" Font-Size="8pt">
                                    </CellStyle>
                                </dx:GridViewDataTextColumn>                                
                                <dx:GridViewDataTextColumn VisibleIndex="24" Caption="7" Width="70px" FieldName="DeliveryD7" HeaderStyle-HorizontalAlign="Center">
                                    <PropertiesTextEdit DisplayFormatString="{0:n0}">
                                        <MaskSettings ErrorText="Please input valid value !"
                                        Mask="<0..999999g>" IncludeLiterals="DecimalSymbol" />
                                        <ValidationSettings ErrorDisplayMode="None"></ValidationSettings>
                                    </PropertiesTextEdit>

<HeaderStyle HorizontalAlign="Center"></HeaderStyle>

                                    <CellStyle Font-Names="Tahoma" Font-Size="8pt">
                                    </CellStyle>
                                </dx:GridViewDataTextColumn>                               
                                <dx:GridViewDataTextColumn VisibleIndex="25" Caption="8" Width="70px" FieldName="DeliveryD8" HeaderStyle-HorizontalAlign="Center">
                                    <PropertiesTextEdit DisplayFormatString="{0:n0}">
                                        <MaskSettings ErrorText="Please input valid value !"
                                        Mask="<0..999999g>" IncludeLiterals="DecimalSymbol" />
                                        <ValidationSettings ErrorDisplayMode="None"></ValidationSettings>
                                    </PropertiesTextEdit>

<HeaderStyle HorizontalAlign="Center"></HeaderStyle>

                                    <CellStyle Font-Names="Tahoma" Font-Size="8pt">
                                    </CellStyle>
                                </dx:GridViewDataTextColumn>                                
                                <dx:GridViewDataTextColumn VisibleIndex="26" Caption="9" Width="70px" FieldName="DeliveryD9" HeaderStyle-HorizontalAlign="Center">
                                    <PropertiesTextEdit DisplayFormatString="{0:n0}">
                                        <MaskSettings ErrorText="Please input valid value !"
                                        Mask="<0..999999g>" IncludeLiterals="DecimalSymbol" />
                                        <ValidationSettings ErrorDisplayMode="None"></ValidationSettings>
                                    </PropertiesTextEdit>

<HeaderStyle HorizontalAlign="Center"></HeaderStyle>

                                    <CellStyle Font-Names="Tahoma" Font-Size="8pt">
                                    </CellStyle>
                                </dx:GridViewDataTextColumn>                                
                                <dx:GridViewDataTextColumn VisibleIndex="27" Caption="10" Width="70px" FieldName="DeliveryD10" HeaderStyle-HorizontalAlign="Center">
                                    <PropertiesTextEdit DisplayFormatString="{0:n0}">
                                        <MaskSettings ErrorText="Please input valid value !"
                                        Mask="<0..999999g>" IncludeLiterals="DecimalSymbol" />
                                        <ValidationSettings ErrorDisplayMode="None"></ValidationSettings>
                                    </PropertiesTextEdit>

<HeaderStyle HorizontalAlign="Center"></HeaderStyle>

                                    <CellStyle Font-Names="Tahoma" Font-Size="8pt">
                                    </CellStyle>
                                </dx:GridViewDataTextColumn>                                
                                <dx:GridViewDataTextColumn VisibleIndex="28" Caption="11" Width="70px" FieldName="DeliveryD11" HeaderStyle-HorizontalAlign="Center">
                                    <PropertiesTextEdit DisplayFormatString="{0:n0}">
                                        <MaskSettings ErrorText="Please input valid value !"
                                        Mask="<0..999999g>" IncludeLiterals="DecimalSymbol" />
                                        <ValidationSettings ErrorDisplayMode="None"></ValidationSettings>
                                    </PropertiesTextEdit>

<HeaderStyle HorizontalAlign="Center"></HeaderStyle>

                                    <CellStyle Font-Names="Tahoma" Font-Size="8pt">
                                    </CellStyle>
                                </dx:GridViewDataTextColumn>                                
                                <dx:GridViewDataTextColumn VisibleIndex="29" Caption="12" Width="70px" FieldName="DeliveryD12" HeaderStyle-HorizontalAlign="Center">
                                    <PropertiesTextEdit DisplayFormatString="{0:n0}">
                                        <MaskSettings ErrorText="Please input valid value !"
                                        Mask="<0..999999g>" IncludeLiterals="DecimalSymbol" />
                                        <ValidationSettings ErrorDisplayMode="None"></ValidationSettings>
                                    </PropertiesTextEdit>

<HeaderStyle HorizontalAlign="Center"></HeaderStyle>

                                    <CellStyle Font-Names="Tahoma" Font-Size="8pt">
                                    </CellStyle>
                                </dx:GridViewDataTextColumn>                                
                                <dx:GridViewDataTextColumn VisibleIndex="30" Caption="13" Width="70px" FieldName="DeliveryD13" HeaderStyle-HorizontalAlign="Center">
                                    <PropertiesTextEdit DisplayFormatString="{0:n0}">
                                        <MaskSettings ErrorText="Please input valid value !"
                                        Mask="<0..999999g>" IncludeLiterals="DecimalSymbol" />
                                        <ValidationSettings ErrorDisplayMode="None"></ValidationSettings>
                                    </PropertiesTextEdit>

<HeaderStyle HorizontalAlign="Center"></HeaderStyle>

                                    <CellStyle Font-Names="Tahoma" Font-Size="8pt">
                                    </CellStyle>
                                </dx:GridViewDataTextColumn>                                
                                <dx:GridViewDataTextColumn VisibleIndex="31" Caption="14" Width="70px" FieldName="DeliveryD14" HeaderStyle-HorizontalAlign="Center">
                                    <PropertiesTextEdit DisplayFormatString="{0:n0}">
                                        <MaskSettings ErrorText="Please input valid value !"
                                        Mask="<0..999999g>" IncludeLiterals="DecimalSymbol" />
                                        <ValidationSettings ErrorDisplayMode="None"></ValidationSettings>
                                    </PropertiesTextEdit>

<HeaderStyle HorizontalAlign="Center"></HeaderStyle>

                                    <CellStyle Font-Names="Tahoma" Font-Size="8pt">
                                    </CellStyle>
                                </dx:GridViewDataTextColumn>                                
                                <dx:GridViewDataTextColumn VisibleIndex="32" Caption="15" Width="70px" FieldName="DeliveryD15" HeaderStyle-HorizontalAlign="Center">
                                    <PropertiesTextEdit DisplayFormatString="{0:n0}">
                                        <MaskSettings ErrorText="Please input valid value !"
                                        Mask="<0..999999g>" IncludeLiterals="DecimalSymbol" />
                                        <ValidationSettings ErrorDisplayMode="None"></ValidationSettings>
                                    </PropertiesTextEdit>

<HeaderStyle HorizontalAlign="Center"></HeaderStyle>

                                    <CellStyle Font-Names="Tahoma" Font-Size="8pt">
                                    </CellStyle>
                                </dx:GridViewDataTextColumn>                                
                                <dx:GridViewDataTextColumn VisibleIndex="33" Caption="16" Width="70px" FieldName="DeliveryD16" HeaderStyle-HorizontalAlign="Center">
                                    <PropertiesTextEdit DisplayFormatString="{0:n0}">
                                        <MaskSettings ErrorText="Please input valid value !"
                                        Mask="<0..999999g>" IncludeLiterals="DecimalSymbol" />
                                        <ValidationSettings ErrorDisplayMode="None"></ValidationSettings>
                                    </PropertiesTextEdit>

<HeaderStyle HorizontalAlign="Center"></HeaderStyle>

                                    <CellStyle Font-Names="Tahoma" Font-Size="8pt">
                                    </CellStyle>
                                </dx:GridViewDataTextColumn>                                
                                <dx:GridViewDataTextColumn VisibleIndex="34" Caption="17" Width="70px" FieldName="DeliveryD17" HeaderStyle-HorizontalAlign="Center">
                                    <PropertiesTextEdit DisplayFormatString="{0:n0}">
                                        <MaskSettings ErrorText="Please input valid value !"
                                        Mask="<0..999999g>" IncludeLiterals="DecimalSymbol" />
                                        <ValidationSettings ErrorDisplayMode="None"></ValidationSettings>
                                    </PropertiesTextEdit>

<HeaderStyle HorizontalAlign="Center"></HeaderStyle>

                                    <CellStyle Font-Names="Tahoma" Font-Size="8pt">
                                    </CellStyle>
                                </dx:GridViewDataTextColumn>                                
                                <dx:GridViewDataTextColumn VisibleIndex="35" Caption="18" Width="70px" FieldName="DeliveryD18" HeaderStyle-HorizontalAlign="Center">
                                    <PropertiesTextEdit DisplayFormatString="{0:n0}">
                                        <MaskSettings ErrorText="Please input valid value !"
                                        Mask="<0..999999g>" IncludeLiterals="DecimalSymbol" />
                                        <ValidationSettings ErrorDisplayMode="None"></ValidationSettings>
                                    </PropertiesTextEdit>

<HeaderStyle HorizontalAlign="Center"></HeaderStyle>

                                    <CellStyle Font-Names="Tahoma" Font-Size="8pt">
                                    </CellStyle>
                                </dx:GridViewDataTextColumn>                                
                                <dx:GridViewDataTextColumn VisibleIndex="36" Caption="19" Width="70px" FieldName="DeliveryD19" HeaderStyle-HorizontalAlign="Center">
                                    <PropertiesTextEdit DisplayFormatString="{0:n0}">
                                        <MaskSettings ErrorText="Please input valid value !"
                                        Mask="<0..999999g>" IncludeLiterals="DecimalSymbol" />
                                        <ValidationSettings ErrorDisplayMode="None"></ValidationSettings>
                                    </PropertiesTextEdit>

<HeaderStyle HorizontalAlign="Center"></HeaderStyle>

                                    <CellStyle Font-Names="Tahoma" Font-Size="8pt">
                                    </CellStyle>
                                </dx:GridViewDataTextColumn>                                
                                <dx:GridViewDataTextColumn VisibleIndex="37" Caption="20" Width="70px" FieldName="DeliveryD20" HeaderStyle-HorizontalAlign="Center">
                                    <PropertiesTextEdit DisplayFormatString="{0:n0}">
                                        <MaskSettings ErrorText="Please input valid value !"
                                        Mask="<0..999999g>" IncludeLiterals="DecimalSymbol" />
                                        <ValidationSettings ErrorDisplayMode="None"></ValidationSettings>
                                    </PropertiesTextEdit>

<HeaderStyle HorizontalAlign="Center"></HeaderStyle>

                                    <CellStyle Font-Names="Tahoma" Font-Size="8pt">
                                    </CellStyle>
                                </dx:GridViewDataTextColumn>                                
                                <dx:GridViewDataTextColumn VisibleIndex="38" Caption="21" Width="70px" FieldName="DeliveryD21" HeaderStyle-HorizontalAlign="Center">
                                    <PropertiesTextEdit DisplayFormatString="{0:n0}">
                                        <MaskSettings ErrorText="Please input valid value !"
                                        Mask="<0..999999g>" IncludeLiterals="DecimalSymbol" />
                                        <ValidationSettings ErrorDisplayMode="None"></ValidationSettings>
                                    </PropertiesTextEdit>

<HeaderStyle HorizontalAlign="Center"></HeaderStyle>

                                    <CellStyle Font-Names="Tahoma" Font-Size="8pt">
                                    </CellStyle>
                                </dx:GridViewDataTextColumn>                                
                                <dx:GridViewDataTextColumn VisibleIndex="39" Caption="22" Width="70px" FieldName="DeliveryD22" HeaderStyle-HorizontalAlign="Center">
                                    <PropertiesTextEdit DisplayFormatString="{0:n0}">
                                        <MaskSettings ErrorText="Please input valid value !"
                                        Mask="<0..999999g>" IncludeLiterals="DecimalSymbol" />
                                        <ValidationSettings ErrorDisplayMode="None"></ValidationSettings>
                                    </PropertiesTextEdit>

<HeaderStyle HorizontalAlign="Center"></HeaderStyle>

                                    <CellStyle Font-Names="Tahoma" Font-Size="8pt">
                                    </CellStyle>
                                </dx:GridViewDataTextColumn>                                
                                <dx:GridViewDataTextColumn VisibleIndex="40" Caption="23" Width="70px" FieldName="DeliveryD23" HeaderStyle-HorizontalAlign="Center">
                                    <PropertiesTextEdit DisplayFormatString="{0:n0}">
                                        <MaskSettings ErrorText="Please input valid value !"
                                        Mask="<0..999999g>" IncludeLiterals="DecimalSymbol" />
                                        <ValidationSettings ErrorDisplayMode="None"></ValidationSettings>
                                    </PropertiesTextEdit>

<HeaderStyle HorizontalAlign="Center"></HeaderStyle>

                                    <CellStyle Font-Names="Tahoma" Font-Size="8pt">
                                    </CellStyle>
                                </dx:GridViewDataTextColumn>                                
                                <dx:GridViewDataTextColumn VisibleIndex="41" Caption="24" Width="70px" FieldName="DeliveryD24" HeaderStyle-HorizontalAlign="Center">
                                    <PropertiesTextEdit DisplayFormatString="{0:n0}">
                                        <MaskSettings ErrorText="Please input valid value !"
                                        Mask="<0..999999g>" IncludeLiterals="DecimalSymbol" />
                                        <ValidationSettings ErrorDisplayMode="None"></ValidationSettings>
                                    </PropertiesTextEdit>

<HeaderStyle HorizontalAlign="Center"></HeaderStyle>

                                    <CellStyle Font-Names="Tahoma" Font-Size="8pt">
                                    </CellStyle>
                                </dx:GridViewDataTextColumn>                               
                                <dx:GridViewDataTextColumn VisibleIndex="42" Caption="25" Width="70px" FieldName="DeliveryD25" HeaderStyle-HorizontalAlign="Center">
                                    <PropertiesTextEdit DisplayFormatString="{0:n0}">
                                        <MaskSettings ErrorText="Please input valid value !"
                                        Mask="<0..999999g>" IncludeLiterals="DecimalSymbol" />
                                        <ValidationSettings ErrorDisplayMode="None"></ValidationSettings>
                                    </PropertiesTextEdit>

<HeaderStyle HorizontalAlign="Center"></HeaderStyle>

                                    <CellStyle Font-Names="Tahoma" Font-Size="8pt">
                                    </CellStyle>
                                </dx:GridViewDataTextColumn>                                
                                <dx:GridViewDataTextColumn VisibleIndex="43" Caption="26" Width="70px" FieldName="DeliveryD26" HeaderStyle-HorizontalAlign="Center">
                                    <PropertiesTextEdit DisplayFormatString="{0:n0}">
                                        <MaskSettings ErrorText="Please input valid value !"
                                        Mask="<0..999999g>" IncludeLiterals="DecimalSymbol" />
                                        <ValidationSettings ErrorDisplayMode="None"></ValidationSettings>
                                    </PropertiesTextEdit>

<HeaderStyle HorizontalAlign="Center"></HeaderStyle>

                                    <CellStyle Font-Names="Tahoma" Font-Size="8pt">
                                    </CellStyle>
                                </dx:GridViewDataTextColumn>                                
                                <dx:GridViewDataTextColumn VisibleIndex="44" Caption="27" Width="70px" FieldName="DeliveryD27" HeaderStyle-HorizontalAlign="Center">
                                    <PropertiesTextEdit DisplayFormatString="{0:n0}">
                                        <MaskSettings ErrorText="Please input valid value !"
                                        Mask="<0..999999g>" IncludeLiterals="DecimalSymbol" />
                                        <ValidationSettings ErrorDisplayMode="None"></ValidationSettings>
                                    </PropertiesTextEdit>

<HeaderStyle HorizontalAlign="Center"></HeaderStyle>

                                    <CellStyle Font-Names="Tahoma" Font-Size="8pt">
                                    </CellStyle>
                                </dx:GridViewDataTextColumn>                                
                                <dx:GridViewDataTextColumn VisibleIndex="45" Caption="28" Width="70px" FieldName="DeliveryD28" HeaderStyle-HorizontalAlign="Center">
                                    <PropertiesTextEdit DisplayFormatString="{0:n0}">
                                        <MaskSettings ErrorText="Please input valid value !"
                                        Mask="<0..999999g>" IncludeLiterals="DecimalSymbol" />
                                        <ValidationSettings ErrorDisplayMode="None"></ValidationSettings>
                                    </PropertiesTextEdit>

<HeaderStyle HorizontalAlign="Center"></HeaderStyle>

                                    <CellStyle Font-Names="Tahoma" Font-Size="8pt">
                                    </CellStyle>
                                </dx:GridViewDataTextColumn>                               
                                <dx:GridViewDataTextColumn VisibleIndex="46" Caption="29" Width="70px" FieldName="DeliveryD29" HeaderStyle-HorizontalAlign="Center">
                                    <PropertiesTextEdit DisplayFormatString="{0:n0}">
                                        <MaskSettings ErrorText="Please input valid value !"
                                        Mask="<0..999999g>" IncludeLiterals="DecimalSymbol" />
                                        <ValidationSettings ErrorDisplayMode="None"></ValidationSettings>
                                    </PropertiesTextEdit>

<HeaderStyle HorizontalAlign="Center"></HeaderStyle>

                                    <CellStyle Font-Names="Tahoma" Font-Size="8pt">
                                    </CellStyle>
                                </dx:GridViewDataTextColumn>                                
                                <dx:GridViewDataTextColumn VisibleIndex="47" Caption="30" Width="70px" FieldName="DeliveryD30" HeaderStyle-HorizontalAlign="Center">
                                    <PropertiesTextEdit DisplayFormatString="{0:n0}">
                                        <MaskSettings ErrorText="Please input valid value !"
                                        Mask="<0..999999g>" IncludeLiterals="DecimalSymbol" />
                                        <ValidationSettings ErrorDisplayMode="None"></ValidationSettings>
                                    </PropertiesTextEdit>
                                    <HeaderStyle HorizontalAlign="Center"></HeaderStyle>
                                    <CellStyle Font-Names="Tahoma" Font-Size="8pt">
                                    </CellStyle>
                                </dx:GridViewDataTextColumn>                                
                                <dx:GridViewDataTextColumn VisibleIndex="48" Caption="31" Width="70px" FieldName="DeliveryD31">
                                    <PropertiesTextEdit DisplayFormatString="{0:n0}">
                                        <MaskSettings ErrorText="Please input valid value !"
                                        Mask="<0..999999g>" IncludeLiterals="DecimalSymbol" />
                                        <ValidationSettings ErrorDisplayMode="None"></ValidationSettings>
                                    </PropertiesTextEdit>
                                    <HeaderStyle HorizontalAlign="Center"></HeaderStyle>
                                    <CellStyle Font-Names="Tahoma" Font-Size="8pt">
                                    </CellStyle>
                                </dx:GridViewDataTextColumn>
                            </Columns>
                            <HeaderStyle HorizontalAlign="Center" />
                        </dx:GridViewBandColumn>
                        <dx:GridViewDataTextColumn VisibleIndex="49" Caption="countPartNo" FieldName="countPartNo" Width="0px">
                        </dx:GridViewDataTextColumn>
                        <dx:GridViewDataTextColumn VisibleIndex="50" Caption="SupplierID" FieldName="SupplierID" Width="0px">
                        </dx:GridViewDataTextColumn>
                        <dx:GridViewDataTextColumn VisibleIndex="51" Caption="CurrCls" FieldName="CurrCls" Width="0px">
                        </dx:GridViewDataTextColumn>
                        <dx:GridViewDataTextColumn VisibleIndex="52" Caption="UnitCls" FieldName="UnitCls" Width="0px">
                        </dx:GridViewDataTextColumn>
                        <dx:GridViewDataTextColumn VisibleIndex="53" Caption="PODeliveryBy" FieldName="PODeliveryBy" Width="0px">
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
    </table>

    <div style="height:8px;"></div>

    <table id="button" style=" width:100%;">
        <tr>
            <td valign="top" align="left">
                <dx:ASPxButton ID="btnSubMenu" runat="server" Text="SUB MENU"
                    Font-Names="Tahoma" Width="85px" Font-Size="8pt">
                </dx:ASPxButton>
            </td>
            <td valign="top" align="right" style="width: 50px;">
            </td>       
            <td valign="top" align="right" style="width: 50px;">
                <dx:ASPxButton ID="btnSave" runat="server" Text="UPLOAD"
                    Font-Names="Tahoma" Width="85px" Font-Size="8pt">
                </dx:ASPxButton>
            </td>
        </tr>
    </table>
                    
    <dx:ASPxGlobalEvents ID="ge" runat="server" >
        <ClientSideEvents ControlsInitialized="function(s, e) {
	        OnControlsInitializedSplitter();
	        OnControlsInitializedGrid();
        }" />
    </dx:ASPxGlobalEvents>
</asp:Content>

