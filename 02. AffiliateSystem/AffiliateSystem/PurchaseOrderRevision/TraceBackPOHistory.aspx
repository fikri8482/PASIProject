<%@ Page Title="" Language="vb" AutoEventWireup="false" MasterPageFile="~/SiteContent.Master" CodeBehind="TraceBackPOHistory.aspx.vb" Inherits="AffiliateSystem.TraceBackPOHistory" %>
<%@ Register Assembly="DevExpress.Web.v14.1, Version=14.1.6.0, Culture=neutral, PublicKeyToken=b88d1754d700e49a" namespace="DevExpress.Web.ASPxGridView.Export" TagPrefix="dx" %>
<%@ Register Assembly="DevExpress.Web.v14.1, Version=14.1.6.0, Culture=neutral, PublicKeyToken=b88d1754d700e49a" namespace="DevExpress.Web.ASPxGridView" TagPrefix="dx" %>
<%@ Register assembly="DevExpress.Web.v14.1, Version=14.1.6.0, Culture=neutral, PublicKeyToken=b88d1754d700e49a" namespace="DevExpress.Web.ASPxEditors" tagprefix="dx" %>
<%@ Register Assembly="DevExpress.Web.v14.1, Version=14.1.6.0, Culture=neutral, PublicKeyToken=b88d1754d700e49a" namespace="DevExpress.Web.ASPxGlobalEvents" TagPrefix="dx" %>
<%@ Register Assembly="DevExpress.Web.v14.1, Version=14.1.6.0, Culture=neutral, PublicKeyToken=b88d1754d700e49a" namespace="DevExpress.Web.ASPxPanel" TagPrefix="dx" %>
<%@ Register Assembly="DevExpress.Web.v14.1, Version=14.1.6.0, Culture=neutral, PublicKeyToken=b88d1754d700e49a" namespace="DevExpress.Web.ASPxRoundPanel" TagPrefix="dx" %>
<%@ Register Assembly="DevExpress.Web.v14.1, Version=14.1.6.0, Culture=neutral, PublicKeyToken=b88d1754d700e49a" Namespace="DevExpress.Web.ASPxCallback" TagPrefix="dx" %>

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

    function clear() {
        txtPartName.SetText('');
    }
</script>
</asp:Content>

<asp:Content ID="Content2" ContentPlaceHolderID="Content" runat="server">
    <table style="width:100%;">
        <tr>
            <td width="100%">
                <table style="border-width: 1pt thin thin thin; border-style: ridge; border-color:#9598A1; width:100%; height: 70px;">
                    <tr>
                        <td colspan="8" height="20">
                            <table id="Table1">
                                <tr>
                                    <td style="width:5px;"></td>
                                    <td align="left" valign="middle" height="20px" width="120px">
                                        <dx:ASPxLabel ID="ASPxLabel1" runat="server" Text="PERIOD"
                                            Font-Names="Tahoma" Font-Size="8pt">
                                        </dx:ASPxLabel>
                                    </td>
                                    <td align="left" valign="middle" height="20px" width="180px">
                                        <dx:ASPxTimeEdit ID="dtPeriodFrom" runat="server" ClientInstanceName="dtPeriodFrom" 
                                            DisplayFormatString="MMM yyyy" EditFormat="Custom" 
                                            EditFormatString="MMM yyyy" Font-Names="Tahoma" Font-Size="8pt" Width="180px">
                                            <ClientSideEvents ValueChanged="function(s, e) {	                                            
                                                grid.PerformCallback('kosong');
                                                cboPONo.PerformCallback(dtPeriodFrom.GetValue().toString());                                                
	                                            lblInfo.SetText('');                                                
                                            }" />
                                        </dx:ASPxTimeEdit>                                                                                   
                                    </td>
                                    <td align="left" valign="middle" width="250px" class="style3"></td>
                                    <td style="width:5px;"></td>
                                    <td align="left" valign="middle" height="20px" width="180px">                                       
                                    </td>
                                    <td style="width:5px;"></td> 
                                </tr>                               
                                <tr>
                                    <td style="width: 5px;"></td>
                                    <td align="left" valign="middle" height="20px" width="120px">
                                        <dx:ASPxLabel ID="ASPxLabel3" runat="server" Text="PARTS NO. / NAME"
                                            Font-Names="Tahoma" Font-Size="8pt">
                                        </dx:ASPxLabel>
                                    </td>
                                    <td align="left" valign="middle" height="20px" width="180px">
                                        <dx:ASPxComboBox ID="cboPartNo" runat="server" 
                                            ClientInstanceName="cboPartNo" Width="100%"
                                            Font-Size="8pt" 
                                            Font-Names="Verdana" TextFormatString="{0}">
                                            <ClientSideEvents SelectedIndexChanged="function(s, e) {
                                                grid.PerformCallback('kosong');
                                                txtPartName.SetText(cboPartNo.GetSelectedItem().GetColumnText(1))
	                                            lblInfo.SetText('');                                                
                                            }" />
                                            <LoadingPanelStyle ImageSpacing="5px">
                                            </LoadingPanelStyle>
                                        </dx:ASPxComboBox>
                                    </td>
                                    <td align="left" valign="middle" width="250px" class="style3">
                                        <dx:ASPxTextBox ID="txtPartName" runat="server" Width="250px" 
                                            ClientInstanceName="txtPartName" Font-Names="Tahoma" Font-Size="8pt"
                                            MaxLength="20" onkeypress="return singlequote(event)" Height="20px" ReadOnly="True">
                                            <readonlystyle backcolor="#CCCCCC">
                                            </readonlystyle>
                                        </dx:ASPxTextBox>
                                    </td>
                                    <td style="width:5px;"></td>  
                                    <td align="left" valign="middle" height="20px" width="180px">
                                        <table style="width:100%;" align="right">
                                            <tr>
                                                <td align="right" valign="middle" style="height:25px; width:90px;">
                                                    <dx:ASPxButton ID="btnSearch" runat="server" Text="SEARCH"
                                                        Font-Names="Tahoma" Width="85px" AutoPostBack="False" Font-Size="8pt" >
                                                        <ClientSideEvents Click="function(s, e) {
                                                            grid.PerformCallback('load');
                                                            lblInfo.SetText('');
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
                                    <td style="width:5px;"></td> 
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

    <div style="height:1px;"></div>

    <table style="width:100%;">
        <tr>
            <td colspan="8" align="left" valign="top" height="220">
                <%--Column : Grid--%>
                <dx:ASPxGridView ID="grid" runat="server" Width="100%"
                    Font-Names="Tahoma" KeyFieldName="PartNo;bln"
                    AutoGenerateColumns="False"
                    ClientInstanceName="grid" Font-Size="8pt">
                    <ClientSideEvents Init="OnInit" FocusedRowChanged="function(s, e) {
	                    OnGridFocusedRowChanged();
                    }" CallbackError="function(s, e) {e.handled = true;}" EndCallback="function(s, e) {
                        var pMsg = s.cpMessage;
                        if (pMsg != '') {
                            if (pMsg.substring(1,5) == '1001' || pMsg.substring(1,5) == '1002' || pMsg.substring(1,5) == '1003' || pMsg.substring(1,5) == '2001') {
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
                        <dx:GridViewDataTextColumn VisibleIndex="1" Caption="No" FieldName="NoUrut" Width="30px">
                            <HeaderStyle HorizontalAlign="Center"></HeaderStyle>
                        </dx:GridViewDataTextColumn>
                        <dx:GridViewDataTextColumn VisibleIndex="2" Caption="PART NO." FieldName="PartNo" Width="90px">
                            <HeaderStyle HorizontalAlign="Center"></HeaderStyle>
                            <CellStyle Font-Names="Tahoma" Font-Size="8pt">
                            </CellStyle>
                        </dx:GridViewDataTextColumn>
                        <dx:GridViewDataTextColumn VisibleIndex="3" Caption="PART NAME" FieldName="PartName" Width="180px">
                            <HeaderStyle HorizontalAlign="Center"></HeaderStyle>
                            <CellStyle Font-Names="Tahoma" Font-Size="8pt">
                            </CellStyle>
                        </dx:GridViewDataTextColumn>
                        <dx:GridViewDataTextColumn VisibleIndex="19" Caption=" " Width="0px"
                            FieldName="bln">                            
                            <HeaderStyle HorizontalAlign="Center"></HeaderStyle>
                            <CellStyle Font-Names="Tahoma" Font-Size="8pt">
                            </CellStyle>
                        </dx:GridViewDataTextColumn>
                        <dx:GridViewDataTextColumn VisibleIndex="4" Caption=" " Width="50px"
                            FieldName="BulanDesc">                            
                            <HeaderStyle HorizontalAlign="Center"></HeaderStyle>
                            <CellStyle Font-Names="Tahoma" Font-Size="8pt">
                            </CellStyle>
                        </dx:GridViewDataTextColumn>
                        <dx:GridViewDataTextColumn VisibleIndex="6" Caption="WK1" Width="80px"
                            FieldName="qty1">
                            <PropertiesTextEdit DisplayFormatString="{0:n0}">
                                <MaskSettings ErrorText="Please input valid value !"
                                Mask="<0..999999g>" IncludeLiterals="DecimalSymbol" />
                                <ValidationSettings ErrorDisplayMode="None"></ValidationSettings>
                            </PropertiesTextEdit>
                            <HeaderStyle HorizontalAlign="Center"></HeaderStyle>
                            <CellStyle Font-Names="Tahoma" Font-Size="8pt">
                            </CellStyle>
                        </dx:GridViewDataTextColumn>
                        <dx:GridViewDataTextColumn VisibleIndex="7" Caption="WK2" Width="80px"
                            FieldName="qty2">
                            <PropertiesTextEdit DisplayFormatString="{0:n0}">
                                <MaskSettings ErrorText="Please input valid value !"
                                Mask="<0..999999g>" IncludeLiterals="DecimalSymbol" />
                                <ValidationSettings ErrorDisplayMode="None"></ValidationSettings>
                            </PropertiesTextEdit>
                            <HeaderStyle HorizontalAlign="Center"></HeaderStyle>
                            <CellStyle Font-Names="Tahoma" Font-Size="8pt">
                            </CellStyle>
                        </dx:GridViewDataTextColumn>
                        <dx:GridViewDataTextColumn VisibleIndex="8" Caption="WK3" Width="80px"
                            FieldName="qty3">
                            <PropertiesTextEdit DisplayFormatString="{0:n0}">
                                <MaskSettings ErrorText="Please input valid value !"
                                Mask="<0..999999g>" IncludeLiterals="DecimalSymbol" />
                                <ValidationSettings ErrorDisplayMode="None"></ValidationSettings>
                            </PropertiesTextEdit>
                            <HeaderStyle HorizontalAlign="Center"></HeaderStyle>
                            <CellStyle Font-Names="Tahoma" Font-Size="8pt">
                            </CellStyle>
                        </dx:GridViewDataTextColumn>
                        <dx:GridViewDataTextColumn VisibleIndex="9" Caption="WK4" Width="80px"
                            FieldName="qty4">
                            <PropertiesTextEdit DisplayFormatString="{0:n0}">
                                <MaskSettings ErrorText="Please input valid value !"
                                Mask="<0..999999g>" IncludeLiterals="DecimalSymbol" />
                                <ValidationSettings ErrorDisplayMode="None"></ValidationSettings>
                            </PropertiesTextEdit>
                            <HeaderStyle HorizontalAlign="Center"></HeaderStyle>
                            <CellStyle Font-Names="Tahoma" Font-Size="8pt">
                            </CellStyle>
                        </dx:GridViewDataTextColumn>
                        <dx:GridViewDataTextColumn VisibleIndex="10" Caption="WK5" Width="80px"
                            FieldName="qty5">
                            <PropertiesTextEdit DisplayFormatString="{0:n0}">
                                <MaskSettings ErrorText="Please input valid value !"
                                Mask="<0..999999g>" IncludeLiterals="DecimalSymbol" />
                                <ValidationSettings ErrorDisplayMode="None"></ValidationSettings>
                            </PropertiesTextEdit>
                            <HeaderStyle HorizontalAlign="Center"></HeaderStyle>
                            <CellStyle Font-Names="Tahoma" Font-Size="8pt">
                            </CellStyle>
                        </dx:GridViewDataTextColumn>
                        <dx:GridViewDataTextColumn VisibleIndex="11" Caption="WK6" Width="80px"
                            FieldName="qty6">
                            <PropertiesTextEdit DisplayFormatString="{0:n0}">
                                <MaskSettings ErrorText="Please input valid value !"
                                Mask="<0..999999g>" IncludeLiterals="DecimalSymbol" />
                                <ValidationSettings ErrorDisplayMode="None"></ValidationSettings>
                            </PropertiesTextEdit>
                            <HeaderStyle HorizontalAlign="Center"></HeaderStyle>
                            <CellStyle Font-Names="Tahoma" Font-Size="8pt">
                            </CellStyle>
                        </dx:GridViewDataTextColumn>
                        <dx:GridViewDataTextColumn VisibleIndex="12" Caption="WK7" Width="80px"
                            FieldName="qty7">
                            <PropertiesTextEdit DisplayFormatString="{0:n0}">
                                <MaskSettings ErrorText="Please input valid value !"
                                Mask="<0..999999g>" IncludeLiterals="DecimalSymbol" />
                                <ValidationSettings ErrorDisplayMode="None"></ValidationSettings>
                            </PropertiesTextEdit>
                            <HeaderStyle HorizontalAlign="Center"></HeaderStyle>
                            <CellStyle Font-Names="Tahoma" Font-Size="8pt">
                            </CellStyle>
                        </dx:GridViewDataTextColumn>
                        <dx:GridViewDataTextColumn VisibleIndex="13" Caption="WK8" Width="80px"
                            FieldName="qty8">
                            <PropertiesTextEdit DisplayFormatString="{0:n0}">
                                <MaskSettings ErrorText="Please input valid value !"
                                Mask="<0..999999g>" IncludeLiterals="DecimalSymbol" />
                                <ValidationSettings ErrorDisplayMode="None"></ValidationSettings>
                            </PropertiesTextEdit>
                            <HeaderStyle HorizontalAlign="Center"></HeaderStyle>
                            <CellStyle Font-Names="Tahoma" Font-Size="8pt">
                            </CellStyle>
                        </dx:GridViewDataTextColumn>
                        <dx:GridViewDataTextColumn VisibleIndex="14" Caption="WK9" Width="80px"
                            FieldName="qty9">
                            <PropertiesTextEdit DisplayFormatString="{0:n0}">
                                <MaskSettings ErrorText="Please input valid value !"
                                Mask="<0..999999g>" IncludeLiterals="DecimalSymbol" />
                                <ValidationSettings ErrorDisplayMode="None"></ValidationSettings>
                            </PropertiesTextEdit>
                            <HeaderStyle HorizontalAlign="Center"></HeaderStyle>
                            <CellStyle Font-Names="Tahoma" Font-Size="8pt">
                            </CellStyle>
                        </dx:GridViewDataTextColumn>
                        <dx:GridViewDataTextColumn VisibleIndex="15" Caption="WK10" Width="80px"
                            FieldName="qty10">
                            <PropertiesTextEdit DisplayFormatString="{0:n0}">
                                <MaskSettings ErrorText="Please input valid value !"
                                Mask="<0..999999g>" IncludeLiterals="DecimalSymbol" />
                                <ValidationSettings ErrorDisplayMode="None"></ValidationSettings>
                            </PropertiesTextEdit>
                            <HeaderStyle HorizontalAlign="Center"></HeaderStyle>
                            <CellStyle Font-Names="Tahoma" Font-Size="8pt">
                            </CellStyle>
                        </dx:GridViewDataTextColumn>
                        <dx:GridViewDataTextColumn VisibleIndex="16" Caption="WK11" Width="80px"
                            FieldName="qty11">
                            <PropertiesTextEdit DisplayFormatString="{0:n0}">
                                <MaskSettings ErrorText="Please input valid value !"
                                Mask="<0..999999g>" IncludeLiterals="DecimalSymbol" />
                                <ValidationSettings ErrorDisplayMode="None"></ValidationSettings>
                            </PropertiesTextEdit>
                            <HeaderStyle HorizontalAlign="Center"></HeaderStyle>
                            <CellStyle Font-Names="Tahoma" Font-Size="8pt">
                            </CellStyle>
                        </dx:GridViewDataTextColumn>
                        <dx:GridViewDataTextColumn VisibleIndex="17" Caption="WK12" Width="80px"
                            FieldName="qty12">
                            <PropertiesTextEdit DisplayFormatString="{0:n0}">
                                <MaskSettings ErrorText="Please input valid value !"
                                Mask="<0..999999g>" IncludeLiterals="DecimalSymbol" />
                                <ValidationSettings ErrorDisplayMode="None"></ValidationSettings>
                            </PropertiesTextEdit>
                            <HeaderStyle HorizontalAlign="Center"></HeaderStyle>
                            <CellStyle Font-Names="Tahoma" Font-Size="8pt">
                            </CellStyle>
                        </dx:GridViewDataTextColumn>
                        <dx:GridViewDataTextColumn VisibleIndex="18" Caption="SeqNo" FieldName="SeqNo" Width="0px">
                            <HeaderStyle HorizontalAlign="Center"></HeaderStyle>
                            <CellStyle Font-Names="Tahoma" Font-Size="8pt">
                            </CellStyle>
                        </dx:GridViewDataTextColumn>
                    </Columns>
                    <SettingsBehavior AllowSelectByRowClick="True" AllowSort="False"
                        ColumnResizeMode="Control" EnableRowHotTrack="True" />
                    <SettingsPager Visible="False" PageSize="13" Position="Top" 
                        Mode="ShowAllRecords">
                        <Summary Text="Page {0} of {1} [{2} record(s)]"
                                  AllPagesText="Page {0} of {1} " />
                        <Summary AllPagesText="Page {0} of {1} " Text="Page {0} of {1} [{2} record(s)]"></Summary>
                    </SettingsPager>
                    <SettingsEditing Mode="Batch" NewItemRowPosition="Bottom">
                        <BatchEditSettings ShowConfirmOnLosingChanges="False" />
                    </SettingsEditing>
                    <Settings ShowHorizontalScrollBar="True" ShowVerticalScrollBar="True"
                        ShowGroupButtons="False" ShowStatusBar="Hidden"
                        VerticalScrollableHeight="190" />
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
                <dx:ASPxButton ID="btnExcel" runat="server" Text="DOWNLOAD"
                    Font-Names="Tahoma"
                    Width="90px" AutoPostBack="False" Font-Size="8pt">
                 <ClientSideEvents Click="function(s, e) {                    
                    lblInfo.SetText('');
                    grid.PerformCallback('excel');
                }" /> 
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

    <dx:ASPxCallback ID="ButtonApprove" runat="server" ClientInstanceName="ButtonApprove">
        <ClientSideEvents EndCallback="function(s, e) {
	        txtUser1.SetText(s.cpUser1);
            txtUser2.SetText(s.cpUser2);
            txtUser3.SetText(s.cpUser3);
            txtUser4.SetText(s.cpUser4);
            txtUser5.SetText(s.cpUser5);
            txtUser6.SetText(s.cpUser6);
            txtUser7.SetText(s.cpUser7);
            txtUser8.SetText(s.cpUser8);

            txtDate1.SetText(s.cpDate1);
            txtDate2.SetText(s.cpDate2);
            txtDate3.SetText(s.cpDate3);
            txtDate4.SetText(s.cpDate4);
            txtDate5.SetText(s.cpDate5);
            txtDate6.SetText(s.cpDate6);
            txtDate7.SetText(s.cpDate7);
            txtDate8.SetText(s.cpDate8);

            var pMsg = s.cpMessage;
            if (pMsg != '') {
                if (pMsg.substring(1,5) == '1007') {
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
    <dx:ASPxCallback ID="ButtonPartNo" runat="server" ClientInstanceName="ButtonPartNo">
        <ClientSideEvents EndCallback="function(s, e) {
	         if (s.cpDelivery != '') 
            {
                txtDelivery.SetText(s.cpDelivery);
                txtCommercial.SetText(s.cpCommercial);
                txtShip.SetText(s.cpShip);
                txtRemarks.SetText(s.cpRemarks);
            } else {
                lblInfo.SetText('');
            }
            delete s.cpDelivery;
        }" />
    </dx:ASPxCallback>
</asp:Content>

