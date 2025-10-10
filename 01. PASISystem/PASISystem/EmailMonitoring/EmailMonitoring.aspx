<%@ Page Title="" Language="vb" AutoEventWireup="false" MasterPageFile="~/SiteContent.Master" CodeBehind="EmailMonitoring.aspx.vb" Inherits="PASISystem.EmailMonitoring" %>
<%@ Register Assembly="DevExpress.Web.v14.1, Version=14.1.6.0, Culture=neutral, PublicKeyToken=b88d1754d700e49a" namespace="DevExpress.Web.ASPxGridView.Export" TagPrefix="dx" %>
<%@ Register Assembly="DevExpress.Web.v14.1, Version=14.1.6.0, Culture=neutral, PublicKeyToken=b88d1754d700e49a" namespace="DevExpress.Web.ASPxGridView" TagPrefix="dx" %>
<%@ Register assembly="DevExpress.Web.v14.1, Version=14.1.6.0, Culture=neutral, PublicKeyToken=b88d1754d700e49a" namespace="DevExpress.Web.ASPxEditors" tagprefix="dx" %>
<%@ Register Assembly="DevExpress.Web.v14.1, Version=14.1.6.0, Culture=neutral, PublicKeyToken=b88d1754d700e49a" namespace="DevExpress.Web.ASPxGlobalEvents" TagPrefix="dx" %>
<%@ Register Assembly="DevExpress.Web.v14.1, Version=14.1.6.0, Culture=neutral, PublicKeyToken=b88d1754d700e49a" namespace="DevExpress.Web.ASPxPanel" TagPrefix="dx" %>
<%@ Register Assembly="DevExpress.Web.v14.1, Version=14.1.6.0, Culture=neutral, PublicKeyToken=b88d1754d700e49a" namespace="DevExpress.Web.ASPxRoundPanel" TagPrefix="dx" %>

<asp:Content ID="Content1" ContentPlaceHolderID="head" runat="server">
    <style type="text/css">
        .dxeHLC, .dxeHC, .dxeHFC
        {
            display: none;
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
        height = height - (height * 50 / 100)
        grid.SetHeight(height);
    }

    function OnBatchEditStartEditing(s, e) {
        currentColumnName = e.focusedColumn.fieldName;

        if (currentColumnName == "Period" || currentColumnName == "AffiliateID" || currentColumnName == "SupplierID"
            || currentColumnName == "PONo" || currentColumnName == "SuratJalanNo") {
            e.cancel = true;
        }

        currentEditableVisibleIndex = e.visibleIndex;
    }
</script>
</asp:Content>

<asp:Content ID="Content2" ContentPlaceHolderID="Content" runat="server">
    <table style="width:100%;">
        <tr>
            <td colspan="2">
                <table style="border-width: 1pt thin thin thin; border-style: ridge; border-color:#9598A1; width:100%; height: 70px;">
                    <tr>
                        <td colspan="2" height="10">
                            <table id="Table1">
                                <tr>
                                    <td style="width:5px;"></td>
                                    <td align="left" valign="middle" height="25px" width="130px">
                                        <dx:ASPxLabel ID="ASPxLabel1" runat="server" Text="PERIOD"
                                            Font-Names="Tahoma" Font-Size="8pt">
                                        </dx:ASPxLabel>
                                    </td>
                                    <td align="left" valign="middle" height="25px" width="400px">
                                        <table>
                                            <tr>
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
                                                <td align="left" valign="middle" style="height:25px; width:280px;">
                                                </td>
                                            </tr>
                                        </table>                                        
                                    </td>
                                    <td align="left" valign="middle" height="25px" width="20px"></td>
                                    <td align="left" valign="middle" style="height:25px; width:180px;"></td>
                                    <td style="width:5px;"></td>
                                </tr>
                                <tr>
                                    <td style="width:5px;"></td>
                                    <td align="left" valign="middle" height="25px" width="130px">
                                        <dx:ASPxLabel ID="ASPxLabel3" runat="server" Text="AFFILIATE CODE"
                                            Font-Names="Tahoma" Font-Size="8pt">
                                        </dx:ASPxLabel>
                                    </td>
                                    <td align="left" valign="middle" height="25px" width="400px">
                                        <table>
                                            <tr>
                                                <td align="left" valign="middle" height="25px" width="120px">
                                                    <dx:ASPxComboBox ID="cboAffiliate" runat="server" 
                                                        ClientInstanceName="cboAffiliate" Width="100%"
                                                        Font-Size="8pt" 
                                                        Font-Names="Verdana" TextFormatString="{0}">
                                                        <ClientSideEvents SelectedIndexChanged="function(s, e) {
	                                                        txtAffiliate.SetText(cboAffiliate.GetSelectedItem().GetColumnText(1));
                                                            grid.PerformCallback('kosong');
	                                                        lblInfo.SetText('');
                                                        }" />
                                                        <LoadingPanelStyle ImageSpacing="5px">
                                                        </LoadingPanelStyle>
                                                    </dx:ASPxComboBox>
                                                </td>
                                                <td align="left" valign="middle" style="height:25px; width:280px;">
                                                    <dx:ASPxTextBox ID="txtAffiliate" runat="server" Width="100%" Height="20px"
                                                        ClientInstanceName="txtAffiliate" Font-Names="Verdana"
                                                        Font-Size="8pt" MaxLength="100" BackColor="#CCCCCC" ReadOnly="True">
                                                    </dx:ASPxTextBox>
                                                </td>
                                            </tr>
                                        </table> 
                                    </td>
                                    <td align="left" valign="middle" height="25px" width="20px"></td>
                                    <td align="left" valign="middle" style="height:25px; width:180px;">                                        
                                    </td>
                                    <td style="width:5px;"></td>
                                </tr>
                                <tr>
                                    <td style="width:5px;"></td>
                                    <td align="left" valign="middle" height="25px" width="130px">
                                        <dx:ASPxLabel ID="ASPxLabel4" runat="server" Text="SUPPLIER CODE"
                                            Font-Names="Tahoma" Font-Size="8pt">
                                        </dx:ASPxLabel>
                                    </td>
                                    <td align="left" valign="middle" height="25px" width="400px">
                                        <table>
                                            <tr>
                                                <td align="left" valign="middle" height="25px" width="120px">
                                                    <dx:ASPxComboBox ID="cboSupplier" runat="server" 
                                                        ClientInstanceName="cboSupplier" Width="100%"
                                                        Font-Size="8pt" 
                                                        Font-Names="Verdana" TextFormatString="{0}">
                                                        <ClientSideEvents SelectedIndexChanged="function(s, e) {
	                                                        txtSupplier.SetText(cboSupplier.GetSelectedItem().GetColumnText(1));
                                                            grid.PerformCallback('kosong');
	                                                        lblInfo.SetText('');
                                                        }" />
                                                        <LoadingPanelStyle ImageSpacing="5px">
                                                        </LoadingPanelStyle>
                                                    </dx:ASPxComboBox>
                                                </td>
                                                <td align="left" valign="middle" style="height:25px; width:280px;">
                                                    <dx:ASPxTextBox ID="txtSupplier" runat="server" Width="100%" Height="20px"
                                                        ClientInstanceName="txtSupplier" Font-Names="Verdana"
                                                        Font-Size="8pt" MaxLength="100" BackColor="#CCCCCC" ReadOnly="True">
                                                    </dx:ASPxTextBox>
                                                </td>
                                            </tr>
                                        </table> 
                                    </td>
                                    <td align="left" valign="middle" height="25px" width="20px"></td>
                                    <td align="left" valign="middle" style="height:25px; width:180px;">     
                                        <table style="width:100%;">
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
                                                        Font-Names="Tahoma" Width="85px" AutoPostBack="False" Font-Size="8pt" Visible="False">
                                                        <ClientSideEvents Click="function(s, e) {
                                                            cboAffiliate.SetText('== ALL ==');
                                                            txtAffiliate.SetText('== ALL ==');
                                                            cboSupplier.SetText('== ALL ==');
                                                            txtSupplier.SetText('== ALL ==');
                                                            lblInfo.SetText('');
                                                            grid.PerformCallback('kosong');
                                                        }" />
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
        
        <tr>
            <td colspan="2" height="15">
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
        
        <tr>
            <td colspan="2" align="left" valign="top" height="220">
                <dx:ASPxGridView ID="grid" runat="server" Width="100%"
                    Font-Names="Tahoma" KeyFieldName="Period;AffiliateID;SupplierID;PONo;SuratJalanNo;KanbanNo"
                    AutoGenerateColumns="False"
                    ClientInstanceName="grid" Font-Size="8pt">
                    <ClientSideEvents Init="OnInit" CallbackError="function(s, e) {e.handled = true;}" EndCallback="function(s, e) {
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
                    }" BatchEditStartEditing="OnBatchEditStartEditing" />
                    
                    <Columns>
                        <dx:GridViewDataTextColumn VisibleIndex="1" Caption="PERIOD" FieldName="Period" Width="80px" HeaderStyle-HorizontalAlign="Center">
                            <PropertiesTextEdit DisplayFormatString="MMM yyyy">
                            </PropertiesTextEdit>
                            <CellStyle Font-Names="Tahoma" Font-Size="8pt">
                            </CellStyle>
                        </dx:GridViewDataTextColumn>
                        <dx:GridViewDataTextColumn VisibleIndex="2" Caption="AFFILIATE" FieldName="AffiliateID" Width="90px" HeaderStyle-HorizontalAlign="Center">
                            <CellStyle Font-Names="Tahoma" Font-Size="8pt">
                            </CellStyle>
                        </dx:GridViewDataTextColumn>
                        <dx:GridViewDataTextColumn VisibleIndex="3" Caption="SUPPLIER" FieldName="SupplierID" Width="90px" HeaderStyle-HorizontalAlign="Center">
                            <CellStyle Font-Names="Tahoma" Font-Size="8pt">
                            </CellStyle>
                        </dx:GridViewDataTextColumn>
                        <dx:GridViewDataTextColumn VisibleIndex="4" Caption="PO NO." FieldName="PONo" Width="100px" HeaderStyle-HorizontalAlign="Center">
                            <CellStyle Font-Names="Tahoma" Font-Size="8pt">
                            </CellStyle>
                        </dx:GridViewDataTextColumn>
                        <dx:GridViewDataTextColumn VisibleIndex="4" Caption="KANBAN NO." FieldName="KanbanNo" Width="100px" HeaderStyle-HorizontalAlign="Center">
                            <CellStyle Font-Names="Tahoma" Font-Size="8pt">
                            </CellStyle>
                        </dx:GridViewDataTextColumn>
                        <dx:GridViewDataTextColumn VisibleIndex="5" Caption="SURAT JALAN NO." FieldName="SuratJalanNo" Width="120px" HeaderStyle-HorizontalAlign="Center">
                            <CellStyle Font-Names="Tahoma" Font-Size="8pt">
                            </CellStyle>
                        </dx:GridViewDataTextColumn>
                        <dx:GridViewDataCheckColumn Caption="PO" FieldName="POCls" 
                            Name="AllowAccess" VisibleIndex="6" Width="85px" >
                            <PropertiesCheckEdit ValueChecked="1" ValueType="System.Int32" 
                                ValueUnchecked="0">
                            </PropertiesCheckEdit>
                            <HeaderStyle HorizontalAlign="Center"></HeaderStyle>
                            <CellStyle Font-Names="Tahoma" Font-Size="8pt">
                            </CellStyle>
                        </dx:GridViewDataCheckColumn>
                        <dx:GridViewDataCheckColumn Caption="DELIVERY" FieldName="DeliveryCls" 
                            Name="AllowAccess" VisibleIndex="7" Width="85px" >
                            <PropertiesCheckEdit ValueChecked="1" ValueType="System.Int32" 
                                ValueUnchecked="0">
                            </PropertiesCheckEdit>
                            <HeaderStyle HorizontalAlign="Center"></HeaderStyle>
                            <CellStyle Font-Names="Tahoma" Font-Size="8pt">
                            </CellStyle>
                        </dx:GridViewDataCheckColumn>
                        <dx:GridViewDataCheckColumn Caption="RECEIVING" FieldName="SendReceivingTOSupplier" 
                            Name="AllowAccess" VisibleIndex="8" Width="85px" >
                            <PropertiesCheckEdit ValueChecked="1" ValueType="System.Int32" 
                                ValueUnchecked="0">
                            </PropertiesCheckEdit>
                            <HeaderStyle HorizontalAlign="Center"></HeaderStyle>
                            <CellStyle Font-Names="Tahoma" Font-Size="8pt">
                            </CellStyle>
                        </dx:GridViewDataCheckColumn>
                    </Columns>
                    <SettingsBehavior AllowSelectByRowClick="True" AllowSort="False"
                        ColumnResizeMode="Control" EnableRowHotTrack="True" />
                    <SettingsPager Visible="False" PageSize="50" Position="Top" 
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

        <tr>
            <td valign="top" align="left">
                <dx:ASPxButton ID="btnSubMenu" runat="server" Text="SUB MENU"
                    Font-Names="Tahoma" Width="90px" Font-Size="8pt">
                </dx:ASPxButton>
            </td>
            <td valign="top" align="right">
                <dx:ASPxButton ID="btnDownload" runat="server" Text="SAVE"
                    Font-Names="Tahoma" Width="85px" Font-Size="8pt" AutoPostBack="False">
                    <ClientSideEvents Click="function(s, e) {     
                        grid.UpdateEdit();
                        grid.PerformCallback('load');
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
</asp:Content>

