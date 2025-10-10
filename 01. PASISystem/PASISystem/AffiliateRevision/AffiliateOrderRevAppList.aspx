<%@ Page Title="" Language="vb" AutoEventWireup="false" MasterPageFile="~/SiteContent.Master" CodeBehind="AffiliateOrderRevAppList.aspx.vb" Inherits="PASISystem.AffiliateOrderRevAppList" %>
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
        height = height - (height * 53 / 100)
        grid.SetHeight(height);
    }

    function OnBatchEditStartEditing(s, e) {
        currentColumnName = e.focusedColumn.fieldName;

        if (currentColumnName == "DetailPage" || currentColumnName == "RevisePage" || currentColumnName == "Period" || currentColumnName == "PONo" || currentColumnName == "CommercialCls"
            || currentColumnName == "ShipCls" || currentColumnName == "CurrAff"
            || currentColumnName == "AmountAff" || currentColumnName == "EntryDate" || currentColumnName == "EntryUser" || currentColumnName == "POStatus1"
            || currentColumnName == "POStatus2" || currentColumnName == "POStatus3" || currentColumnName == "POStatus4" || currentColumnName == "POStatus5" || currentColumnName == "POStatus6" || currentColumnName == "POStatus7" || currentColumnName == "POStatus8") {
            e.cancel = true;
        }

        currentEditableVisibleIndex = e.visibleIndex;
    }
    function SelectedIndexChangedAff() {
        txtAffiliateName.SetText(cboAffiliateCode.GetSelectedItem().GetColumnText(1));
        lblInfo.SetText('');
    }
    function SelectedIndexChangedSupp() {
        txtSupplierName.SetText(cboSupplierCode.GetSelectedItem().GetColumnText(1));
        lblInfo.SetText('');
    }
</script>
</asp:Content>

<asp:Content ID="Content2" ContentPlaceHolderID="Content" runat="server">
    <table style="width:100%;">
        <tr>
            <td width="50%">
                <table style="border-width: 1pt thin thin thin; border-style: ridge; border-color:#9598A1; width:100%; height: 120px;">
                    <tr>
                        <td height="30">
                            <table id="Table1">
                                <tr>
                                    <td style="width:5px;"></td>
                                    <td align="left" valign="middle" height="25px" width="105px">
                                        <dx:ASPxLabel ID="ASPxLabel1" runat="server" Text="PERIOD"
                                            Font-Names="Tahoma" Font-Size="8pt">
                                        </dx:ASPxLabel>
                                    </td>
                                    <td align="left" valign="middle" height="25px" width="203px" colspan="5">
                                        <table>
                                            <tr>
                                                <td>
                                                    <dx:ASPxTimeEdit ID="dtPeriodFrom" runat="server" ClientInstanceName="dtPeriodFrom" 
                                                        DisplayFormatString="MMM yyyy" EditFormat="Custom" 
                                                        EditFormatString="MMM yyyy" Font-Names="Tahoma" Font-Size="8pt" 
                                                        Width="100px">
                                                        <ClientSideEvents ValueChanged="function(s, e) {
	                                                        grid.SetFocusedRowIndex(-1);
                                                            grid.PerformCallback('kosong');
	                                                        lblInfo.SetText('');
                                                        }" />
                                                    </dx:ASPxTimeEdit>
                                                </td>
                                                <td align="left" valign="middle" height="25px" width="3px"> ~
                                                </td>
                                                <td>
                                                    <dx:ASPxTimeEdit ID="dtPeriodTo" runat="server" ClientInstanceName="dtPeriodTo" 
                                                        DisplayFormatString="MMM yyyy" EditFormat="Custom" 
                                                        EditFormatString="MMM yyyy" Font-Names="Tahoma" Font-Size="8pt" 
                                                        Width="100px">
                                                        <ClientSideEvents ValueChanged="function(s, e) {
	                                                        grid.SetFocusedRowIndex(-1);
                                                            grid.PerformCallback('kosong');
	                                                        lblInfo.SetText('');
                                                        }" />
                                                    </dx:ASPxTimeEdit>
                                                </td>
                                            </tr>
                                        </table>                                        
                                    </td>
                                </tr>
                                <tr>
                                    <td style="width:5px;"></td>
                                    <td align="left" valign="middle" height="25px" width="105px">
                                        PO REVISION No.</td>
                                    <td align="left" valign="middle" height="25px" width="203px" colspan="5">
                                        <dx:ASPxComboBox ID="cboPONoRev" runat="server" 
                                            ClientInstanceName="cboPONoRev" Width="203px"
                                            Font-Size="8pt" 
                                            Font-Names="Verdana" TextFormatString="{0}">
                                            <ClientSideEvents SelectedIndexChanged="function(s, e) {
                                                var pDateFrom = dtPeriodFrom.GetText();
                                                var pDateTo = dtPeriodTo.GetText();
                                                var pPONoRev = cboPONoRev.GetText();                                                
                                                cboPONo.PerformCallback('loadCombo' + '|' + pDateFrom + '|' + pDateTo + '|' + pPONoRev);
                                                lblInfo.SetText('');
                                            }" />
                                            <LoadingPanelStyle ImageSpacing="5px">
                                            </LoadingPanelStyle>
                                        </dx:ASPxComboBox>
                                    </td>                                    
                                </tr>
                                <tr>
                                    <td style="width:5px;"></td>
                                    <td align="left" valign="middle" height="25px" width="105px">
                                        PO No.</td>
                                    <td align="left" valign="middle" height="25px" width="203px" colspan="5">
                                        <dx:ASPxComboBox ID="cboPONo" runat="server" 
                                            ClientInstanceName="cboPONo" Width="203px"
                                            Font-Size="8pt" 
                                            Font-Names="Verdana" TextFormatString="{0}">
                                            <ClientSideEvents SelectedIndexChanged="function(s, e) {
                                                var pDateFrom = dtPeriodFrom.GetText();
                                                var pDateTo = dtPeriodTo.GetText();
                                                var pPONoRev = cboPONoRev.GetText();         
                                                var pPONo = cboPONo.GetText();         
                                                cboAffiliateCode.PerformCallback('loadCombo' + '|' + pDateFrom + '|' + pDateTo + '|' + pPONoRev + '|' + pPONo);
                                                cboSupplierCode.PerformCallback('loadCombo' + '|' + pDateFrom + '|' + pDateTo + '|' + pPONoRev + '|' + pPONo);
                                                rblSendToSupp.SetValue('2');
                                                rblCommercial.SetValue('2');
	                                            lblInfo.SetText('');                            
                                            }" />
                                            <LoadingPanelStyle ImageSpacing="5px">
                                            </LoadingPanelStyle>
                                        </dx:ASPxComboBox>
                                    </td>
                                </tr>
                                <tr>
                                    <td style="width:5px;">&nbsp;</td>
                                     <td align="left" valign="middle" width="120px">
                                        <dx:ASPxLabel ID="ASPxLabel3" runat="server" Text="AFFILIATE CODE/NAME" Font-Names="Tahoma"
                                            Font-Size="8pt" Width="150px">
                                        </dx:ASPxLabel>
                                    </td>
                                    <td align="left" valign="middle" width="150px">
                                        <dx:ASPxComboBox ID="cboAffiliateCode" runat="server" TextFormatString="{0}" DropDownStyle="DropDown"
                                            Width="203px" ClientInstanceName="cboAffiliateCode" IncrementalFilteringMode="StartsWith"
                                            Font-Names="Tahoma" Font-Size="8pt" MaxLength="10" TabIndex="4">
                                            <ClientSideEvents SelectedIndexChanged="function(s, e) {SelectedIndexChangedAff();}"
                                                LostFocus="function(s, e) { lblInfo.SetText(''); }" />
                                        </dx:ASPxComboBox>
                                    </td>
                                    <td align="left" valign="middle" colspan="4" width="360px">
                                        <dx:ASPxTextBox ID="txtAffiliateName" runat="server" ClientInstanceName="txtAffiliateName"
                                            Width="370px" Height="20px" ReadOnly="True" TabIndex="5" 
                                            Font-Names="Tahoma" Font-Size="8pt">
                                            <ReadOnlyStyle BackColor="#CCCCCC">
                                            </ReadOnlyStyle>
                                        </dx:ASPxTextBox>
                                    </td>
                                </tr>
                                <tr>
                                    <td style="width: 5px;">
                                    </td>
                                    <td align="left" valign="middle" width="80px">
                                        <dx:ASPxLabel ID="ASPxLabel5" runat="server" Text="SUPPLIER CODE/NAME" Font-Names="Tahoma"
                                            Font-Size="8pt" Width="150px">
                                        </dx:ASPxLabel>
                                    </td>
                                    <td align="left" valign="middle" width="150px">
                                        <dx:ASPxComboBox ID="cboSupplierCode" runat="server" TextFormatString="{0}" DropDownStyle="DropDown"
                                            Height="20px" Width="203px" ClientInstanceName="cboSupplierCode" IncrementalFilteringMode="StartsWith"
                                            Font-Names="Tahoma" Font-Size="8pt" MaxLength="10" TabIndex="6">
                                            <ClientSideEvents SelectedIndexChanged="function(s, e) {SelectedIndexChangedSupp();}"
                                                LostFocus="function(s, e) { lblInfo.SetText(''); }" />
                                        </dx:ASPxComboBox>
                                    </td>
                                    <td align="left" valign="middle" width="360px" colspan="4">
                                        <dx:ASPxTextBox ID="txtSupplierName" runat="server" ClientInstanceName="txtSupplierName"
                                            Width="370px" Height="20px" ReadOnly="True" TabIndex="7" 
                                            Font-Names="Tahoma" Font-Size="8pt">
                                            <ReadOnlyStyle BackColor="#CCCCCC">
                                            </ReadOnlyStyle>
                                        </dx:ASPxTextBox>
                                    </td>
                                </tr>
                                <tr>
                                    <td style="width: 5px;">
                                        &nbsp;</td>
                                    <td align="left" valign="middle" width="80px">
                                        <dx:ASPxLabel ID="ASPxLabel2" runat="server" Text="SEND TO SUPPLIER"
                                            Font-Names="Tahoma" Font-Size="8pt">
                                        </dx:ASPxLabel>
                                    </td>
                                    <td align="left" valign="middle" width="150px">
                                        <dx:ASPxRadioButtonList ID="rblSendToSupp" runat="server" RepeatDirection="Horizontal"
                                            Width="200px" ClientInstanceName="rblSendToSupp" SelectedIndex="0" Font-Names="Tahoma"
                                            Font-Size="8pt" Height="5px">
                                            <Items>
                                                <dx:ListEditItem Text="ALL" Value="2" Selected="True" />
                                                <dx:ListEditItem Text="YES" Value="1" />
                                                <dx:ListEditItem Text="NO" Value="0" />
                                            </Items>
                                            <Border BorderStyle="None"></Border>
                                        </dx:ASPxRadioButtonList>
                                    </td>
                                    <td align="left" valign="middle" width="360px" style="width: -720">
                                        &nbsp;</td>
                                    <td align="left" valign="middle" width="360px" style="width: -180">
                                        &nbsp;</td>
                                    <td align="left" valign="middle" width="360px" style="width: 0">
                                        &nbsp;</td>
                                    <td align="left" valign="middle" width="360px" style="width: 90px">
                                        &nbsp;</td>
                                </tr>
                                <tr>
                                    <td style="width: 5px;">
                                        &nbsp;</td>
                                    <td align="left" valign="middle" width="80px">
                                        <dx:ASPxLabel ID="ASPxLabel4" runat="server" Text="COMMERCIAL"
                                            Font-Names="Tahoma" Font-Size="8pt">
                                        </dx:ASPxLabel>
                                    </td>
                                    <td align="left" valign="middle" width="150px">
                                        <dx:ASPxRadioButtonList ID="rblCommercial" runat="server" RepeatDirection="Horizontal"
                                            Width="200px" ClientInstanceName="rblCommercial" SelectedIndex="0" Font-Names="Tahoma"
                                            Font-Size="8pt" Height="5px">
                                            <Items>
                                                <dx:ListEditItem Text="ALL" Value="2" Selected="True" />
                                                <dx:ListEditItem Text="YES" Value="1" />
                                                <dx:ListEditItem Text="NO" Value="0" />
                                            </Items>
                                            <Border BorderStyle="None"></Border>
                                        </dx:ASPxRadioButtonList>
                                    </td>
                                    <td align="left" valign="middle" width="360px" style="width: -720">
                                        &nbsp;</td>
                                    <td align="left" valign="middle" width="360px" style="width: -180">
                                        &nbsp;</td>
                                    <td align="left" valign="middle" width="360px" style="width: 0">
                                        &nbsp;</td>
                                    <td align="left" valign="middle" width="360px" style="width: 90px">
                                        <table>
                                            <tr>
                                                <td align="right" valign="middle" width="85px">
                                                    <dx:ASPxButton ID="btnRefresh" runat="server" Text="SEARCH"
                                                        Font-Names="Tahoma" Width="85px" AutoPostBack="False" Font-Size="8pt" >
                                                         <ClientSideEvents Click="function(s, e) {
                                                            var pDateFrom = dtPeriodFrom.GetText();
                                                            var pDateTo = dtPeriodTo.GetText();
                                                            var pPONoRev = cboPONoRev.GetText();         
                                                            var pPONo = cboPONo.GetText();
                                                            var pAffCode= cboAffiliateCode.GetText();
                                                            var pSuppCode = cboSupplierCode.GetText();
                                                            var pSendTo = rblSendToSupp.GetValue();
                                                            var pComm = rblCommercial.GetValue();
                                                            
                                                            grid.PerformCallback('load' + '|' + pDateFrom + '|' + pDateTo+ '|' + pPONoRev + '|' + pPONo+ '|' + pAffCode + '|' + pSuppCode + '|' + pSendTo + '|' + pComm);
                                                            lblInfo.SetText('');
                                                        }" />
                                                    </dx:ASPxButton>
                                                </td>
                                                <td align="left" valign="middle" width="85px">
                                                    <dx:ASPxButton ID="btnClear" runat="server" Text="CLEAR"
                                                        Font-Names="Tahoma" Width="85px" AutoPostBack="False" Font-Size="8pt">
                                                        <ClientSideEvents Click="function(s, e) {
                                                            txtPartCode.SetText('');
                                                            txtPartName.SetText('');
                                                            lblInfo.SetText('');
                                                            grid.SetFocusedRowIndex(-1);
                                                            grid.PerformCallback('kosong');
                                                        }" />                                   
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
            <td valign="top" width="40%" align="left">
                <dx:ASPxRoundPanel ID="ASPxRoundPanel1" runat="server" HeaderText="PO STATUS" ShowCollapseButton="true"
                    View="GroupBox" Width="100%" Font-Size="8pt" Font-Names="Tahoma">
                    <PanelCollection>
                        <dx:PanelContent ID="PanelContent1" runat="server">
                            <table id="Table2">
                                <tr>
                                    <td align="left" height="20px" valign="middle" width="50%">
                                        <dx:ASPxLabel ID="ASPxLabel8" runat="server" Font-Names="Tahoma" Font-Size="8pt"
                                            ForeColor="#003366" Text="(1) AFFILIATE ENTRY" Width="100%">
                                        </dx:ASPxLabel>
                                    </td>
                                    <td align="left" height="20px" valign="middle" width="50%">
                                        <dx:ASPxLabel ID="ASPxLabel14" runat="server" Font-Names="Tahoma" Font-Size="8pt"
                                            ForeColor="#003366" Text="(5) SUPPLIER PENDING (PARTIAL)" Width="100%">
                                        </dx:ASPxLabel>
                                    </td>
                                </tr>
                                <tr>
                                    <td align="left" height="20px" valign="middle" width="50%">
                                        <dx:ASPxLabel ID="ASPxLabel11" runat="server" Font-Names="Tahoma" Font-Size="8pt"
                                            ForeColor="#003366" Text="(2) AFFILIATE APPROVAL" Width="100%">
                                        </dx:ASPxLabel>
                                    </td>
                                    <td align="left" height="20px" valign="middle" width="50%">
                                        <dx:ASPxLabel ID="ASPxLabel15" runat="server" Font-Names="Tahoma" Font-Size="8pt"
                                            ForeColor="#003366" Text="(6) SUPPLIER UNAPPROVE" Width="100%">
                                        </dx:ASPxLabel>
                                    </td>
                                </tr>
                                <tr>
                                    <td align="left" height="20px" valign="middle" width="50%">
                                        <dx:ASPxLabel ID="ASPxLabel12" runat="server" Font-Names="Tahoma" Font-Size="8pt"
                                            ForeColor="#003366" Text="(3) PASI SEND AFFILIATE PO REVISION TO SUPPLIER" 
                                            Width="100%">
                                        </dx:ASPxLabel>
                                    </td>
                                    <td align="left" height="20px" valign="middle" width="50%">
                                        <dx:ASPxLabel ID="ASPxLabel16" runat="server" Font-Names="Tahoma" Font-Size="8pt"
                                            ForeColor="#003366" Text="(7) PASI APPROVAL" Width="100%">
                                        </dx:ASPxLabel>
                                    </td>
                                </tr>
                                <tr>
                                    <td align="left" height="20px" valign="middle" width="50%">
                                        <dx:ASPxLabel ID="ASPxLabel13" runat="server" Font-Names="Tahoma" Font-Size="8pt"
                                            ForeColor="#003366" Text="(4) SUPPLIER APPROVAL" Width="100%">
                                        </dx:ASPxLabel>
                                    </td>
                                    <td align="left" height="20px" valign="middle" width="50%">
                                        <dx:ASPxLabel ID="ASPxLabel17" runat="server" Font-Names="Tahoma" Font-Size="8pt"
                                            ForeColor="#003366" Text="(8) AFFILIATE FINAL APPROVAL" Width="100%">
                                        </dx:ASPxLabel>
                                    </td>
                                </tr>
                            </table>
                        </dx:PanelContent>
                    </PanelCollection>
                </dx:ASPxRoundPanel>
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
        <tr>
            <td colspan="8" align="left" valign="top" height="220">
                <%--Column : Grid--%>
                <dx:ASPxGridView ID="grid" runat="server" Width="100%"
                    Font-Names="Tahoma" KeyFieldName="Period;AffiliateID;PORevNo;PONo;SupplierID"
                    AutoGenerateColumns="False"
                    ClientInstanceName="grid" Font-Size="8pt">
                    <ClientSideEvents Init="OnInit" FocusedRowChanged="function(s, e) {
	                    OnGridFocusedRowChanged();
                    }" EndCallback="function(s, e) {
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
                        <dx:GridViewDataTextColumn VisibleIndex="0" Caption=" " Width="50px" FieldName="DetailPage">
                            <DataItemTemplate>
                                <a id="clickElement" href="AffiliateOrderRevAppDetail.aspx?id=<%#GetRowValue(Container)%>&t1=<%#GetPeriod(Container)%>&t2=<%#GetPORevNo(Container)%>&t3=<%#GetPONo(Container)%>&t4=<%#GetCommercial(Container)%>&t5=<%#GetAffiliateID(Container)%>&t6=<%#GetAffiliateName(Container)%>&t7=<%#GetSupplierID(Container)%>&t8=<%#GetSupplierName(Container)%>&t9=<%#GetKanban(Container)%>&t10=<%#GetRemarks(Container)%>&Session=~/AffiliateRevision/AffiliateOrderRevAppList.aspx">
                                    <%# "DETAIL"%></a>
                            </DataItemTemplate>
                            <CellStyle Font-Names="Tahoma" Font-Size="8pt">
                            </CellStyle>
                        </dx:GridViewDataTextColumn>                       
                        <dx:GridViewDataTextColumn VisibleIndex="1" Caption="PERIOD" FieldName="Period" Width="70px"
                            HeaderStyle-HorizontalAlign="Center">
                            <PropertiesTextEdit DisplayFormatString="MMM yyyy">
                            </PropertiesTextEdit>
                            <HeaderStyle HorizontalAlign="Center"></HeaderStyle>
                            <CellStyle Font-Names="Tahoma" Font-Size="8pt">
                            </CellStyle>
                        </dx:GridViewDataTextColumn>
                        <dx:GridViewDataTextColumn VisibleIndex="4" Caption="PO REVISION NO." 
                            FieldName="PORevNo" Width="220px"
                            HeaderStyle-HorizontalAlign="Center">
                            <HeaderStyle HorizontalAlign="Center"></HeaderStyle>
                            <CellStyle Font-Names="Tahoma" Font-Size="8pt">
                            </CellStyle>
                        </dx:GridViewDataTextColumn>
                        <dx:GridViewDataTextColumn VisibleIndex="5" Caption="PO NO." FieldName="PONo" Width="220px"
                            HeaderStyle-HorizontalAlign="Center">
                            <HeaderStyle HorizontalAlign="Center"></HeaderStyle>
                            <CellStyle Font-Names="Tahoma" Font-Size="8pt">
                            </CellStyle>
                        </dx:GridViewDataTextColumn>
                        <dx:GridViewDataTextColumn VisibleIndex="5" Caption="PO MARKING" FieldName="POMarking" Width="220px"
                            HeaderStyle-HorizontalAlign="Center">
                            <HeaderStyle HorizontalAlign="Center"></HeaderStyle>
                            <CellStyle Font-Names="Tahoma" Font-Size="8pt">
                            </CellStyle>
                        </dx:GridViewDataTextColumn>
                        <dx:GridViewDataTextColumn VisibleIndex="6" Caption="COMMERCIAL" FieldName="CommercialCls"
                            Width="95px" HeaderStyle-HorizontalAlign="Center" 
                            CellStyle-HorizontalAlign="Center">
                            <HeaderStyle HorizontalAlign="Center"></HeaderStyle>
                            <CellStyle Font-Names="Tahoma" Font-Size="8pt">
                            </CellStyle>
                        </dx:GridViewDataTextColumn>
                        <dx:GridViewDataTextColumn VisibleIndex="9" Caption="SHIP BY" FieldName="ShipCls"
                            Width="100px" HeaderStyle-HorizontalAlign="Center">
                            <HeaderStyle HorizontalAlign="Center"></HeaderStyle>
                            <CellStyle Font-Names="Tahoma" Font-Size="8pt">
                            </CellStyle>
                        </dx:GridViewDataTextColumn>                        
                        <dx:GridViewDataTextColumn VisibleIndex="14" Caption="CREATED DATE" FieldName="EntryDate"
                            HeaderStyle-HorizontalAlign="Center" Width="130px">
                            <PropertiesTextEdit DisplayFormatString="yyyy-MM-dd HH:mm:ss">
                            </PropertiesTextEdit>
                            <CellStyle Font-Names="Tahoma" Font-Size="8pt">
                            </CellStyle>
                            <HeaderStyle HorizontalAlign="Center"></HeaderStyle>
                        </dx:GridViewDataTextColumn>
                        <dx:GridViewDataTextColumn VisibleIndex="15" Caption="CREATED BY" FieldName="EntryUser"
                            HeaderStyle-HorizontalAlign="Center" Width="145px">
                            <HeaderStyle HorizontalAlign="Center"></HeaderStyle>
                            <CellStyle Font-Names="Tahoma" Font-Size="8pt">
                            </CellStyle>
                        </dx:GridViewDataTextColumn>
                        <dx:GridViewBandColumn Caption="PO STATUS" VisibleIndex="16">
                            <Columns>
                                <dx:GridViewDataCheckColumn Caption="1" FieldName="POStatus1" ReadOnly="True" VisibleIndex="1" Width="30px">
                                    <PropertiesCheckEdit ValueChecked="1" ValueType="System.String" ValueUnchecked="0">
                                        <DisplayImageChecked IconID="support_feature_16x16">
                                        </DisplayImageChecked>
                                        <DisplayImageUnchecked Width="0px">
                                        </DisplayImageUnchecked>
                                        <CheckBoxStyle>
                                        <Border BorderStyle="None" />
                                        </CheckBoxStyle>
                                    </PropertiesCheckEdit>
                                    <HeaderStyle HorizontalAlign="Center"></HeaderStyle>
                                </dx:GridViewDataCheckColumn>
                                <dx:GridViewDataCheckColumn Caption="2" FieldName="POStatus2" ReadOnly="True" VisibleIndex="3" Width="30px">
                                    <PropertiesCheckEdit ValueChecked="1" ValueType="System.String" ValueUnchecked="0">
                                        <DisplayImageChecked IconID="support_feature_16x16">
                                        </DisplayImageChecked>
                                        <DisplayImageUnchecked Width="0px">
                                        </DisplayImageUnchecked>
                                    </PropertiesCheckEdit>
                                    <HeaderStyle HorizontalAlign="Center"></HeaderStyle>
                                </dx:GridViewDataCheckColumn>
                                <dx:GridViewDataCheckColumn Caption="3" FieldName="POStatus3" ReadOnly="True" VisibleIndex="5" Width="30px">
                                    <PropertiesCheckEdit ValueChecked="1" ValueType="System.String" ValueUnchecked="0">
                                        <DisplayImageChecked IconID="support_feature_16x16">
                                        </DisplayImageChecked>
                                        <DisplayImageUnchecked Width="0px">
                                        </DisplayImageUnchecked>
                                    </PropertiesCheckEdit>
                                    <HeaderStyle HorizontalAlign="Center"></HeaderStyle>
                                </dx:GridViewDataCheckColumn>
                                <dx:GridViewDataCheckColumn Caption="4" FieldName="POStatus4" ReadOnly="True" VisibleIndex="7" Width="30px">
                                    <PropertiesCheckEdit ValueChecked="1" ValueType="System.String" ValueUnchecked="0">
                                        <DisplayImageChecked IconID="support_feature_16x16">
                                        </DisplayImageChecked>
                                        <DisplayImageUnchecked Width="0px">
                                        </DisplayImageUnchecked>
                                    </PropertiesCheckEdit>
                                </dx:GridViewDataCheckColumn>
                                <dx:GridViewDataCheckColumn Caption="5" FieldName="POStatus5" ReadOnly="True" VisibleIndex="9" Width="30px">
                                    <PropertiesCheckEdit ValueChecked="1" ValueType="System.String" ValueUnchecked="0">
                                        <DisplayImageChecked IconID="support_feature_16x16">
                                        </DisplayImageChecked>
                                        <DisplayImageUnchecked Width="0px">
                                        </DisplayImageUnchecked>
                                    </PropertiesCheckEdit>
                                    <HeaderStyle HorizontalAlign="Center"></HeaderStyle>
                                </dx:GridViewDataCheckColumn>
                                <dx:GridViewDataCheckColumn Caption="6" FieldName="POStatus6" ReadOnly="True" VisibleIndex="11" Width="30px">
                                    <PropertiesCheckEdit ValueChecked="1" ValueType="System.String" ValueUnchecked="0">
                                        <DisplayImageChecked IconID="support_feature_16x16">
                                        </DisplayImageChecked>
                                        <DisplayImageUnchecked Width="0px">
                                        </DisplayImageUnchecked>
                                    </PropertiesCheckEdit>
                                    <HeaderStyle HorizontalAlign="Center"></HeaderStyle>
                                </dx:GridViewDataCheckColumn>
                                <dx:GridViewDataCheckColumn Caption="7" FieldName="POStatus7" ReadOnly="True" VisibleIndex="13" Width="30px">
                                    <PropertiesCheckEdit ValueChecked="1" ValueType="System.String" ValueUnchecked="0">
                                        <DisplayImageChecked IconID="support_feature_16x16">
                                        </DisplayImageChecked>
                                        <DisplayImageUnchecked Width="0px">
                                        </DisplayImageUnchecked>
                                    </PropertiesCheckEdit>
                                    <HeaderStyle HorizontalAlign="Center"></HeaderStyle>
                                </dx:GridViewDataCheckColumn>
                                <dx:GridViewDataCheckColumn Caption="8" FieldName="POStatus8" ReadOnly="True" VisibleIndex="14" Width="30px">
                                    <PropertiesCheckEdit ValueChecked="1" ValueType="System.String" ValueUnchecked="0">
                                        <DisplayImageChecked IconID="support_feature_16x16">
                                        </DisplayImageChecked>
                                        <DisplayImageUnchecked Width="0px">
                                        </DisplayImageUnchecked>
                                    </PropertiesCheckEdit>
                                    <HeaderStyle HorizontalAlign="Center"></HeaderStyle>
                                </dx:GridViewDataCheckColumn>
                            </Columns>
                            <HeaderStyle HorizontalAlign="Center"></HeaderStyle>
                        </dx:GridViewBandColumn>
                        <dx:GridViewDataTextColumn Caption="AFFILIATE CODE" FieldName="AffiliateID" 
                            VisibleIndex="2" Width="100px" HeaderStyle-HorizontalAlign="Center">
                            <HeaderStyle HorizontalAlign="Center"></HeaderStyle>
                            <CellStyle Font-Names="Tahoma" Font-Size="8pt">
                            </CellStyle>
                        </dx:GridViewDataTextColumn>
                        <dx:GridViewDataTextColumn Caption="AFFILIATE NAME" FieldName="AffiliateName" 
                            VisibleIndex="3" Width="100px" HeaderStyle-HorizontalAlign="Center">
                            <HeaderStyle HorizontalAlign="Center"></HeaderStyle>
                            <CellStyle Font-Names="Tahoma" Font-Size="8pt">
                            </CellStyle>
                        </dx:GridViewDataTextColumn>
                        <dx:GridViewDataTextColumn Caption="SUPPLIER CODE" FieldName="SupplierID" 
                            VisibleIndex="7" Width="100px" HeaderStyle-HorizontalAlign="Center">
                            <HeaderStyle HorizontalAlign="Center"></HeaderStyle>
                            <CellStyle Font-Names="Tahoma" Font-Size="8pt">
                            </CellStyle>
                        </dx:GridViewDataTextColumn>
                        <dx:GridViewDataTextColumn Caption="SUPPLIER NAME" FieldName="SupplierName" 
                            VisibleIndex="8" Width="100px" HeaderStyle-HorizontalAlign="Center">
                            <HeaderStyle HorizontalAlign="Center"></HeaderStyle>
                            <CellStyle Font-Names="Tahoma" Font-Size="8pt">
                            </CellStyle>
                        </dx:GridViewDataTextColumn>
                        <dx:GridViewDataTextColumn Caption="KanbanCls" FieldName="KanbanCls" 
                            VisibleIndex="12" Width="0px">
                        </dx:GridViewDataTextColumn>
                        <dx:GridViewDataTextColumn Caption="Remarks" FieldName="Remarks" 
                            VisibleIndex="13" Width="0px">
                        </dx:GridViewDataTextColumn>
                    </Columns>
                    <SettingsBehavior AllowSelectByRowClick="True" AllowSort="False"
                        ColumnResizeMode="Control" EnableRowHotTrack="True" />
                    <SettingsPager Visible="False" PageSize="13" Position="Top">
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
        </tr>
    </table>
    
    <dx:ASPxGlobalEvents ID="ge" runat="server" >
        <ClientSideEvents ControlsInitialized="function(s, e) {
	        OnControlsInitializedSplitter();
	        OnControlsInitializedGrid();
        }" />
    </dx:ASPxGlobalEvents>                    
</asp:Content>
