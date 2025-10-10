<%@ Page Title="" Language="vb" AutoEventWireup="false" MasterPageFile="~/SiteContent.Master" CodeBehind="POExportEntry.aspx.vb" Inherits="AffiliateSystem.POExportEntry" %>
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
            width: 157px;
        }
        .style2
        {
            height: 20px;
            width: 157px;
        }
        .style3
        {
            width: 100px;
        }
        .style4
        {
            width: 66px;
        }
        .style5
        {
            height: 20px;
            width: 66px;
        }
        .style6
        {
            width: 32px;
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
        height = height - (height * 40 / 100)
        grid.SetHeight(height);
    }

    function OnBatchEditStartEditing(s, e) {
        currentColumnName = e.focusedColumn.fieldName;

        if (currentColumnName == "NoUrut" || currentColumnName == "PartNo" || currentColumnName == "PartName" || currentColumnName == "UnitDesc"
            || currentColumnName == "MinOrderQty" || currentColumnName == "Maker" || currentColumnName == "KanbanCls" || currentColumnName == "PONo" || currentColumnName == "QtyBox"
            || currentColumnName == "CurrDesc" || currentColumnName == "Price" || currentColumnName == "Amount" || currentColumnName == "POQty"
            || currentColumnName == "ForecastN1" || currentColumnName == "ForecastN2" || currentColumnName == "ForecastN3") {
            e.cancel = true;
        }

        currentEditableVisibleIndex = e.visibleIndex;
    }

    function onHeijunka() {
        
    }

    function OnBatchEditEndEditing(s, e) {
        window.setTimeout(function () {
            //            var period = dtPeriod.GetValue();
            //            var nBulan = period.getMonth();
            //            var tahun = period.getFullYear();
            //            var kabisat;
            var hariIsi;
            hariIsi = 31;
            //            //Januari, Maret, May, July, Aug, Oct, Dec --> 31
            //            if (nBulan == "0" || nBulan == "2" || nBulan == "4" || nBulan == "6" || nBulan == "7" || nBulan == "9" || nBulan == "11") {
            //                hariIsi = 31;
            //            }
            //            //April, Jun, Sep, Nov --> 30
            //            if (nBulan == "3" || nBulan == "5" || nBulan == "8" || nBulan == "10") {
            //                hariIsi = 30;
            //            }
            //            //Februari --> ??
            //            if (nBulan == "1") {
            //                kabisat = tahun % 4;
            //                if (kabisat = 0) {
            //                    hariIsi = 29;
            //                } else {
            //                    hariIsi = 28;
            //                }
            //            }

            //s.batchEditApi.SetCellValue(e.visibleIndex, "AmountSupp", priceSupp * Qty);

            var total = 0;
            for (i = 1; i <= hariIsi; i++) {
                var day = "DeliveryD" + i;
                var qtyHarian = s.batchEditApi.GetCellValue(e.visibleIndex, day);
                total = total + parseInt(qtyHarian);
                s.batchEditApi.SetCellValue(e.visibleIndex, "POQty", total);
            }

            //            var priceAff = s.batchEditApi.GetCellValue(e.visibleIndex, "PriceAff");
            //            var Qty = s.batchEditApi.GetCellValue(e.visibleIndex, "POQty");
            //            var RemainingQty = s.batchEditApi.GetCellValue(e.visibleIndex - 1, "MonthlyProductionCapacity");
            //            //var priceSupp = s.batchEditApi.GetCellValue(e.visibleIndex, "PriceSupp");
            //            var qtyBox = s.batchEditApi.GetCellValue(e.visibleIndex - 1, "QtyBox");

            //            //s.batchEditApi.SetCellValue(e.visibleIndex, "AmountAff", priceAff * Qty);
            //            s.batchEditApi.SetCellValue(e.visibleIndex, "Flag", 1);
        }, 10);
    }

    function validasubmit() {
        lblInfo.GetMainElement().style.color = 'Red';
        if (txtPONo.GetText() == "") {
            lblInfo.SetText("[6011] Please Input PO No. first!");
            txtPONo.Focus();
            e.ProcessOnServer = false;
            return false;
        }

        if (txtShip.GetText() == "") {
            lblInfo.SetText("[6011] Please Input Ship By first!");
            txtShip.Focus();
            e.ProcessOnServer = false;
            return false;
        }
    }

    function clear() {
        txtUser1.SetText('');
        txtUser2.SetText('');
        txtUser3.SetText('');
        txtUser4.SetText('');
        txtUser5.SetText('');
        txtUser6.SetText('');
        txtUser7.SetText('');
        txtUser8.SetText('');

        txtDate1.SetText('');
        txtDate2.SetText('');
        txtDate3.SetText('');
        txtDate4.SetText('');
        txtDate5.SetText('');
        txtDate6.SetText('');
        txtDate7.SetText('');
        txtDate8.SetText('');
    }

    function up_delete() {
        if (txtPONo.GetText() == "") {
            lblInfo.GetMainElement().style.color = 'Red';
            lblInfo.SetText("[6011] Please input PO No first!");
            e.ProcessOnServer = false;
            return false;
        }

        if (txtDate2.GetText() != "") {
            lblInfo.GetMainElement().style.color = 'Red';
            lblInfo.SetText("[6011] Can't delete, because this PO already Approve!");
            e.ProcessOnServer = false;
            return false;
        }

        var msg = confirm('Are you sure want to delete this data ?');
        if (msg == false) {
            e.processOnServer = false;
            return;
        }

        var pGroupCode = txtPONo.GetText();
        ButtonDelete.PerformCallback('delete|' + pGroupCode);
    }
</script>
</asp:Content>

<asp:Content ID="Content2" ContentPlaceHolderID="Content" runat="server">
    <table style="width:100%;">
        <tr>
            <td colspan="16" width="100%">
                <table style="border-width: 1pt thin thin thin; border-style: ridge; border-color:#9598A1; width:100%; height: 70px;">
                    <tr>
                        <td colspan="8" height="20">
                            <table id="Table1">
                                <tr>
                                    <td style="width:5px;"></td>
                                    <td align="left" valign="middle" height="20px" class="style1">
                                        <dx:ASPxLabel ID="ASPxLabel1" runat="server" Text="PERIOD"
                                            Font-Names="Tahoma" Font-Size="8pt">
                                        </dx:ASPxLabel>
                                    </td>
                                    <td align="left" valign="middle" height="20px" width="180px">
                                        <dx:ASPxTimeEdit ID="dtPeriodFrom" runat="server" ClientInstanceName="dtPeriodFrom" 
                                            DisplayFormatString="MMM yyyy" EditFormat="Custom" 
                                            EditFormatString="MMM yyyy" Font-Names="Tahoma" Font-Size="8pt" Width="180px">
                                            <ClientSideEvents ValueChanged="function(s, e) {	                                            
                                                grid.PerformCallback('load');
	                                            lblInfo.SetText('');
                                            }" />
                                        </dx:ASPxTimeEdit>                                                                                   
                                    </td>
                                    <td style="width:5px;">&nbsp;</td>
                                                <td align="left">
                                                    <table style="width:100%;">
                                                        <tr>
                                                            <td>
                                        <dx:ASPxLabel ID="ASPxLabel4" runat="server" Text="COMMERCIAL"
                                            Font-Names="Tahoma" Font-Size="8pt">
                                        </dx:ASPxLabel>
                                                            </td>
                                                            <td>
                                                                &nbsp;</td>
                                                            <td>
                                                                &nbsp;</td>
                                                            <td>
                                                    <dx:ASPxRadioButton ID="rdrCom1" ClientInstanceName="rdrCom1" runat="server" Text="YES" GroupName="Commercial" Font-Names="Tahoma" Font-Size="8pt">
                                                        <ClientSideEvents LostFocus="function(s, e) {                                                        
                                                        lblInfo.SetText('');                                                         
                                                        }" />
                                                    </dx:ASPxRadioButton>
                                                            </td>
                                                            <td>
                                                    <dx:ASPxRadioButton ID="rdrCom2" ClientInstanceName="rdrCom2" runat="server" Text="NO" GroupName="Commercial" Font-Names="Tahoma" Font-Size="8pt">
                                                        <ClientSideEvents LostFocus="function(s, e) {
                                                        lblInfo.SetText('');                                                         
                                                        }" />
                                                    </dx:ASPxRadioButton>
                                                            </td>
                                                        </tr>
                                                    </table>
                                                </td>
                                    <td class="style6">&nbsp;</td>
                                    <td class="style4">&nbsp;</td>
                                    <td style="width:5px;">&nbsp;</td>
                                    <td align="left" valign="middle" style="height:20px; width:60px;">
                                        &nbsp;</td>
                                    <td align="left" valign="middle" height="20px" width="150px">
                                        <table>
                                            <tr>
                                                <td>
                                                    &nbsp;</td>
                                            </tr>
                                        </table> 
                                    </td>
                                </tr>
                                <tr>
                                    <td style="width:5px;"></td>
                                    <td align="left" valign="middle" class="style2">
                                        <dx:ASPxLabel ID="ASPxLabel5" runat="server" Text="PO NO."
                                            Font-Names="Tahoma" Font-Size="8pt">
                                        </dx:ASPxLabel>
                                    </td>
                                    <td align="left" valign="middle" height="20px" width="180px">
                                        <dx:ASPxTextBox ID="txtPONo" runat="server" Width="180px" 
                                            ClientInstanceName="txtPONo" Font-Names="Tahoma" Font-Size="8pt"
                                            MaxLength="20" onkeypress="return singlequote(event)" Height="20px">
                                            <ClientSideEvents LostFocus="function(s, e) { 

	                                            lblInfo.SetText('');
                                            }" />
                                        </dx:ASPxTextBox>
                                    </td>
                                    <td style="width:5px;">&nbsp;</td>
                                    <td style="width:5px;">
                                        <table style="width:3560%; margin-right: 20px;">
                                            <tr>
                                                <td class="style3">
                                        <dx:ASPxLabel ID="ASPxLabel2" runat="server" Text="PO EMERGENCY"
                                            Font-Names="Tahoma" Font-Size="8pt">
                                        </dx:ASPxLabel>
                                                </td>
                                                <td>
                                                    <dx:ASPxRadioButton ID="rdrEmergency2" ClientInstanceName="rdrKanban2" 
                                                        runat="server" Text="YES" GroupName="POKan" Font-Names="Tahoma" Font-Size="8pt">
                                                        <ClientSideEvents CheckedChanged="function(s, e) {
                                                        lblInfo.SetText('');
                                                        }" />
                                                    </dx:ASPxRadioButton>
                                                </td>
                                                <td>
                                                    <dx:ASPxRadioButton ID="rdrEmergency3" ClientInstanceName="rdrKanban3" 
                                                        runat="server" Text="NO" GroupName="POKan" Font-Names="Tahoma" Font-Size="8pt">
                                                        <ClientSideEvents CheckedChanged="function(s, e) {
                                                        lblInfo.SetText('');
                                                        }" />
                                                    </dx:ASPxRadioButton>
                                                </td>
                                            </tr>
                                        </table>
                                    </td>
                                    <td class="style6"></td>
                                    <td align="left" valign="middle" class="style5">
                                        &nbsp;</td>
                                    <td align="left" valign="middle" style="height:20px; width:60px;">
                                        &nbsp;</td>
                                    <td align="left" valign="middle" height="20px" width="180px">
                                        &nbsp;</td>

                                    <td align="left" valign="middle" height="20px" width="80px"></td>

                                    <td align="right" valign="middle" style="height:25px; width:200px;">
                                        <table style="width:100%;">
                                            <tr>
                                                <td align="right" valign="middle" style="height:25px; width:90px;">
                                                    <%--<dx:ASPxButton ID="btnCraete" runat="server" Text="CREATE"
                                                        Font-Names="Tahoma" Width="85px" AutoPostBack="False" Font-Size="8pt">
                                                        <ClientSideEvents Click="function(s, e) {
                                                            lblInfo.SetText('');
                                                            clear();
                                                            validasubmit();
                                                            grid.PerformCallback('load');
                                                        }" />
                                                    </dx:ASPxButton>--%>
                                                </td>
                                                <td align="right" valign="middle" style="height:25px; width:90px;">
                                                    &nbsp;</td>
                                            </tr>
                                        </table>
                                    </td>
                                </tr>                               
                                <tr>
                                    <td style="width:5px;">&nbsp;</td>
                                    <td align="left" valign="middle" class="style2">
                                        <dx:ASPxLabel ID="ASPxLabel20" runat="server" Text="AFFILIATE CODE/NAME"
                                            Font-Names="Tahoma" Font-Size="8pt">
                                        </dx:ASPxLabel>
                                    </td>
                                    <td align="left" valign="middle" height="20px" width="180px">
                                        <dx:ASPxComboBox ID="cboAffiliate" width="180px" runat="server" 
                                            ValueType="System.String" ClientInstanceName="cboAffiliate">
                                        </dx:ASPxComboBox>
                                    </td>
                                    <td style="width:5px;">&nbsp;</td>
                                    <td style="width:5px;">
                                        <dx:ASPxTextBox ID="txtAffiliate" runat="server" Width="180px" 
                                            ClientInstanceName="txtAffiliate" Font-Names="Tahoma" Font-Size="8pt"
                                            MaxLength="20" onkeypress="return singlequote(event)" Height="20px">
                                            <readonlystyle backcolor="#CCCCCC">
                                            </readonlystyle>
                                        </dx:ASPxTextBox>
                                    </td>
                                    <td class="style6">&nbsp;</td>
                                    <td align="left" valign="middle" class="style5">
                                        &nbsp;</td>
                                    <td align="left" valign="middle" style="height:20px; width:60px;">
                                        &nbsp;</td>
                                    <td align="left" valign="middle" height="20px" width="180px">
                                        &nbsp;</td>

                                    <td align="left" valign="middle" height="20px" width="80px">&nbsp;</td>

                                    <td align="right" valign="middle" style="height:25px; width:200px;">
                <dx:ASPxButton ID="btnClear" runat="server" Text="CLEAR"
                Font-Names="Tahoma" Width="85px" AutoPostBack="False" Font-Size="8pt">                                
            </dx:ASPxButton>
                                    </td>
                                </tr>                               
                                <tr>
                                    <td style="width:5px;">&nbsp;</td>
                                    <td align="left" valign="middle" class="style2">
                                        <dx:ASPxLabel ID="ASPxLabel19" runat="server" Text="SUPPLIER CODE/NAME"
                                            Font-Names="Tahoma" Font-Size="8pt">
                                        </dx:ASPxLabel>
                                    </td>
                                    <td align="left" valign="middle" height="20px" width="180px">
                                        <dx:ASPxComboBox ID="cboSupplier" width="180px" runat="server" 
                                            ValueType="System.String" ClientInstanceName="cboSupplier">
                                        </dx:ASPxComboBox>
                                    </td>
                                    <td style="width:5px;">&nbsp;</td>
                                    <td style="width:5px;">
                                        <dx:ASPxTextBox ID="txtSupplier" runat="server" Width="180px" 
                                            ClientInstanceName="txtSupplier" Font-Names="Tahoma" Font-Size="8pt"
                                            MaxLength="20" onkeypress="return singlequote(event)" Height="20px">
                                            <readonlystyle backcolor="#CCCCCC">
                                            </readonlystyle>
                                        </dx:ASPxTextBox>
                                    </td>
                                    <td class="style6">&nbsp;</td>
                                    <td align="left" valign="middle" class="style5">
                                        <dx:ASPxLabel ID="ASPxLabel3" runat="server" Text="SHIP BY"
                                            Font-Names="Tahoma" Font-Size="8pt">
                                        </dx:ASPxLabel>
                                    </td>
                                    <td align="left" valign="middle" style="height:20px; width:60px;">
                                        &nbsp;</td>
                                    <td align="left" valign="middle" height="20px" width="180px">
                                        <dx:ASPxTextBox ID="txtShip" runat="server" Width="180px" 
                                            ClientInstanceName="txtShip" Font-Names="Tahoma" Font-Size="8pt"
                                            MaxLength="25" onkeypress="return singlequote(event)" Height="20px">
                                            <ClientSideEvents LostFocus="function(s, e) { lblInfo.SetText(&quot;&quot;); }" />
                                        </dx:ASPxTextBox>
                                    </td>

                                    <td align="left" valign="middle" height="20px" width="80px">&nbsp;</td>

                                    <td align="right" valign="middle" style="height:25px; width:200px;">
                                        &nbsp;</td>
                                </tr>                               
                            </table>
                        </td>
                    </tr>
                </table>
            </td>            
        </tr>

        <tr>
            <td colspan="16" height="15">
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
            <td valign="top" colspan="16"  align="left">
                <dx:ASPxRoundPanel ID="ASPxRoundPanel1" runat="server" HeaderText="PO STATUS" ShowCollapseButton="true"
                    View="GroupBox" Width="100%" Font-Names="Tahoma">
                    <PanelCollection>
                        <dx:PanelContent ID="PanelContent1" runat="server">
                            <table id="Table2">
                                <tr>
                                    <td align="left" valign="middle">
                                        <dx:ASPxLabel ID="ASPxLabel18" runat="server" Font-Names="Tahoma" Font-Size="8pt"
                                            ForeColor="#003366" Text="(1) AFFILIATE ENTRY">
                                        </dx:ASPxLabel>
                                    </td>
                                    <td align="left" valign="middle">
                                        <dx:ASPxTextBox ID="txtDate1" runat="server" ClientInstanceName="txtDate1" 
                                            Font-Names="Tahoma" Font-Size="8pt" Height="20px" MaxLength="10" 
                                            ReadOnly="True" Width="130px">
                                            <readonlystyle backcolor="#CCCCCC">
                                            </readonlystyle>
                                        </dx:ASPxTextBox>
                                    </td>
                                    <td align="left" valign="middle">
                                        <dx:ASPxTextBox ID="txtUser1" runat="server" ClientInstanceName="txtUser1" 
                                            Font-Names="Tahoma" Font-Size="8pt" Height="20px" MaxLength="10" 
                                            ReadOnly="True" Width="130px">
                                            <readonlystyle backcolor="#CCCCCC">
                                            </readonlystyle>
                                        </dx:ASPxTextBox>
                                    </td>
                                    <td align="left" valign="middle">
                                        &nbsp;</td>
                                    <td align="left" valign="middle">
                                        &nbsp;</td>
                                    <td align="left" valign="middle">
                                        <dx:ASPxLabel ID="ASPxLabel14" runat="server" Font-Names="Tahoma" 
                                            Font-Size="8pt" ForeColor="#003366" Text="(5) SUPPLIER PENDING (PARTIAL)">
                                        </dx:ASPxLabel>
                                    </td>
                                    <td align="left" valign="middle">
                                        <dx:ASPxTextBox ID="txtDate5" runat="server" ClientInstanceName="txtDate5" 
                                            Font-Names="Tahoma" Font-Size="8pt" Height="20px" MaxLength="10" 
                                            ReadOnly="True" Width="130px">
                                            <readonlystyle backcolor="#CCCCCC">
                                            </readonlystyle>
                                        </dx:ASPxTextBox>
                                    </td>
                                    <td align="left" valign="middle">
                                        <dx:ASPxTextBox ID="txtUser5" runat="server" ClientInstanceName="txtUser5" 
                                            Font-Names="Tahoma" Font-Size="8pt" Height="20px" MaxLength="10" 
                                            ReadOnly="True" Width="130px">
                                            <readonlystyle backcolor="#CCCCCC">
                                            </readonlystyle>
                                        </dx:ASPxTextBox>
                                    </td>
                                </tr>
                                <tr>
                                    <td align="left" valign="middle">
                                        <dx:ASPxLabel ID="ASPxLabel11" runat="server" Font-Names="Tahoma" Font-Size="8pt"
                                            ForeColor="#003366" Text="(2) AFFILIATE APPROVAL">
                                        </dx:ASPxLabel>
                                    </td>
                                    <td align="left" valign="middle">
                                        <dx:ASPxTextBox ID="txtDate2" runat="server" ClientInstanceName="txtDate2"
                                            Font-Names="Tahoma" Font-Size="8pt" Height="20px" MaxLength="10"
                                            ReadOnly="True" Width="130px">
                                            <readonlystyle backcolor="#CCCCCC">
                                            </readonlystyle>
                                        </dx:ASPxTextBox>
                                    </td>
                                    <td align="left" valign="middle">
                                        <dx:ASPxTextBox ID="txtUser2" runat="server" ClientInstanceName="txtUser2"
                                            Font-Names="Tahoma" Font-Size="8pt" Height="20px" MaxLength="10"
                                            ReadOnly="True" Width="130px">
                                            <readonlystyle backcolor="#CCCCCC">
                                            </readonlystyle>
                                        </dx:ASPxTextBox>
                                    </td>
                                    <td align="left" valign="middle">
                                        &nbsp;</td>
                                    <td align="left" valign="middle">
                                        &nbsp;</td>
                                    <td align="left" valign="middle">
                                        <dx:ASPxLabel ID="ASPxLabel15" runat="server" Font-Names="Tahoma" 
                                            Font-Size="8pt" ForeColor="#003366" Text="(6) SUPPLIER UNAPPROVE">
                                        </dx:ASPxLabel>
                                    </td>
                                    <td align="left" valign="middle">
                                        <dx:ASPxTextBox ID="txtDate6" runat="server" ClientInstanceName="txtDate6"
                                            Font-Names="Tahoma" Font-Size="8pt" Height="20px" MaxLength="10"
                                            ReadOnly="True" Width="130px">
                                            <readonlystyle backcolor="#CCCCCC">
                                            </readonlystyle>
                                        </dx:ASPxTextBox>
                                    </td>
                                    <td align="left" valign="middle">
                                        <dx:ASPxTextBox ID="txtUser6" runat="server" ClientInstanceName="txtUser6"
                                            Font-Names="Tahoma" Font-Size="8pt" Height="20px" MaxLength="10"
                                            ReadOnly="True" Width="130px">
                                            <readonlystyle backcolor="#CCCCCC">
                                            </readonlystyle>
                                        </dx:ASPxTextBox>
                                    </td>
                                </tr>
                                <tr>
                                    <td align="left" valign="middle">
                                        <dx:ASPxLabel ID="ASPxLabel12" runat="server" Font-Names="Tahoma" Font-Size="8pt"
                                            ForeColor="#003366" Text="(3) PASI SEND AFFILIATE PO TO SUPPLIER">
                                        </dx:ASPxLabel>
                                    </td>
                                    <td align="left" valign="middle">
                                        <dx:ASPxTextBox ID="txtDate3" runat="server" ClientInstanceName="txtDate3"
                                            Font-Names="Tahoma" Font-Size="8pt" Height="20px" MaxLength="10"
                                            ReadOnly="True" Width="130px">
                                            <readonlystyle backcolor="#CCCCCC">
                                            </readonlystyle>
                                        </dx:ASPxTextBox>
                                    </td>
                                    <td align="left" valign="middle">
                                        <dx:ASPxTextBox ID="txtUser3" runat="server" ClientInstanceName="txtUser3"
                                            Font-Names="Tahoma" Font-Size="8pt" Height="20px" MaxLength="10"
                                            ReadOnly="True" Width="130px">
                                            <readonlystyle backcolor="#CCCCCC">
                                            </readonlystyle>
                                        </dx:ASPxTextBox>
                                    </td>
                                    <td align="left" valign="middle">
                                        &nbsp;</td>
                                    <td align="left" valign="middle">
                                        &nbsp;</td>
                                    <td align="left" valign="middle">
                                        <dx:ASPxLabel ID="ASPxLabel16" runat="server" Font-Names="Tahoma" 
                                            Font-Size="8pt" ForeColor="#003366" Text="(7) PASI APPROVAL">
                                        </dx:ASPxLabel>
                                    </td>
                                    <td align="left" valign="middle">
                                        <dx:ASPxTextBox ID="txtDate7" runat="server" ClientInstanceName="txtDate7"
                                            Font-Names="Tahoma" Font-Size="8pt" Height="20px" MaxLength="10"
                                            ReadOnly="True" Width="130px">
                                            <readonlystyle backcolor="#CCCCCC">
                                            </readonlystyle>
                                        </dx:ASPxTextBox>
                                    </td>
                                    <td align="left" valign="middle">
                                        <dx:ASPxTextBox ID="txtUser7" runat="server" ClientInstanceName="txtUser7"
                                            Font-Names="Tahoma" Font-Size="8pt" Height="20px" MaxLength="10"
                                            ReadOnly="True" Width="130px">
                                            <readonlystyle backcolor="#CCCCCC">
                                            </readonlystyle>
                                        </dx:ASPxTextBox>
                                    </td>
                                </tr>
                                <tr>
                                    <td align="left" valign="middle">
                                        <dx:ASPxLabel ID="ASPxLabel13" runat="server" Font-Names="Tahoma" Font-Size="8pt"
                                            ForeColor="#003366" Text="(4) SUPPLIER APPROVAL (FULL)">
                                        </dx:ASPxLabel>
                                    </td>
                                    <td align="left" valign="middle">
                                        <dx:ASPxTextBox ID="txtDate4" runat="server" ClientInstanceName="txtDate4"
                                            Font-Names="Tahoma" Font-Size="8pt" Height="20px" MaxLength="10"
                                            ReadOnly="True" Width="130px">
                                            <readonlystyle backcolor="#CCCCCC">
                                            </readonlystyle>
                                        </dx:ASPxTextBox>
                                    </td>
                                    <td align="left" valign="middle">
                                        <dx:ASPxTextBox ID="txtUser4" runat="server" ClientInstanceName="txtUser4"
                                            Font-Names="Tahoma" Font-Size="8pt" Height="20px" MaxLength="10"
                                            ReadOnly="True" Width="130px">
                                            <readonlystyle backcolor="#CCCCCC">
                                            </readonlystyle>
                                        </dx:ASPxTextBox>
                                    </td>
                                    <td align="left" valign="middle">
                                        &nbsp;</td>
                                    <td align="left" valign="middle">
                                        &nbsp;</td>
                                    <td align="left" valign="middle">
                                        <dx:ASPxLabel ID="ASPxLabel17" runat="server" Font-Names="Tahoma"
                                            Font-Size="8pt" ForeColor="#003366" Text="(8) AFFILIATE FINAL APPROVAL">
                                        </dx:ASPxLabel>
                                    </td>
                                    <td align="left" valign="middle">
                                        <dx:ASPxTextBox ID="txtDate8" runat="server" ClientInstanceName="txtDate8"
                                            Font-Names="Tahoma" Font-Size="8pt" Height="20px" MaxLength="10"
                                            ReadOnly="True" Width="130px">
                                            <readonlystyle backcolor="#CCCCCC">
                                            </readonlystyle>
                                        </dx:ASPxTextBox>
                                    </td>
                                    <td align="left" valign="middle">
                                        <dx:ASPxTextBox ID="txtUser8" runat="server" ClientInstanceName="txtUser8"
                                            Font-Names="Tahoma" Font-Size="8pt" Height="20px" MaxLength="10"
                                            ReadOnly="True" Width="130px">
                                            <readonlystyle backcolor="#CCCCCC">
                                            </readonlystyle>
                                        </dx:ASPxTextBox>
                                    </td>
                                </tr>
                            </table>
                        </dx:PanelContent>
                    </PanelCollection>
                </dx:ASPxRoundPanel>
            </td>
        </tr>
            <table style="width: 100%;">
                <tr>
                    <td>
                        <dx:ASPxGridView ID="ASPxGridView1" runat="server" AutoGenerateColumns="False">
                            <Columns>
                                <dx:GridViewDataTextColumn Caption="WEEK 1" VisibleIndex="0">
                                </dx:GridViewDataTextColumn>
                                <dx:GridViewDataTextColumn Caption="WEEK 2" VisibleIndex="1">
                                </dx:GridViewDataTextColumn>
                                <dx:GridViewDataTextColumn Caption="WEEK 3" VisibleIndex="2">
                                </dx:GridViewDataTextColumn>
                                <dx:GridViewDataTextColumn Caption="WEEK 4" VisibleIndex="3">
                                </dx:GridViewDataTextColumn>
                                <dx:GridViewDataTextColumn Caption="WEEK 5" VisibleIndex="4">
                                </dx:GridViewDataTextColumn>
                            </Columns>
                        </dx:ASPxGridView>
                        &nbsp;
                    </td>
                </tr>
                </table>
        </tr>
        <tr>
            <td colspan="16" align="left" valign="top" height="220">
    <%--Column : Grid--%>
                <dx:ASPxGridView ID="grid" runat="server" Width="100%"
                    Font-Names="Tahoma" KeyFieldName="PartNo;PONo"
                    AutoGenerateColumns="False"
                    ClientInstanceName="grid" Font-Size="8pt">
                    <ClientSideEvents Init="OnInit" FocusedRowChanged="function(s, e) {
	                    OnGridFocusedRowChanged();
                    }" CallbackError="function(s, e) {e.handled = true;}" BatchEditEndEditing="OnBatchEditEndEditing" EndCallback="function(s, e) {
                        var pMsg = s.cpMessage;
                        if (pMsg != '') {
                            if (pMsg.substring(1,5) == '1001' || pMsg.substring(1,5) == '1002' || pMsg.substring(1,5) == '1003' || pMsg.substring(1,5) == '2001') {
                                lblInfo.GetMainElement().style.color = 'Blue';
                                txtUser1.SetText(s.cpUser1);
                                txtDate1.SetText(s.cpDate1);     
                                txtPONo.SetText(s.cpPONo);  
                            } else {
                                lblInfo.GetMainElement().style.color = 'Red';
                            }
                            lblInfo.SetText(pMsg);
                        } else {
                            lblInfo.SetText('');
                        }
                        delete s.cpMessage;
                        delete s.cpPONo;
                    }" RowClick="function(s, e) {
	                    lblInfo.SetText('');
                    }" BatchEditStartEditing="OnBatchEditStartEditing" />                    
                    <Columns>
                        <dx:GridViewDataCheckColumn Caption=" " FieldName="AllowAccess" 
                            Name="AllowAccess" VisibleIndex="0" Width="30px" >
                            <PropertiesCheckEdit ValueChecked="1" ValueType="System.Int32" 
                                ValueUnchecked="0">
                            </PropertiesCheckEdit>
                            <HeaderStyle HorizontalAlign="Center"></HeaderStyle>
                            <CellStyle Font-Names="Tahoma" Font-Size="8pt">
                            </CellStyle>
                        </dx:GridViewDataCheckColumn>
                        <dx:GridViewDataTextColumn VisibleIndex="1" Caption="NO." FieldName="NoUrut" Width="30px" HeaderStyle-HorizontalAlign="Center">                            
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
<dx:GridViewDataTextColumn FieldName="PONo" Width="120px" Caption="PO NO." Visible="False" 
                            VisibleIndex="15">
<HeaderStyle HorizontalAlign="Center"></HeaderStyle>

<CellStyle Font-Names="Tahoma" Font-Size="8pt"></CellStyle>
</dx:GridViewDataTextColumn>
<dx:GridViewDataTextColumn FieldName="POQty" Width="100px" Caption="TOTAL FIRM QTY" VisibleIndex="14">
<PropertiesTextEdit DisplayFormatString="{0:n0}">
<MaskSettings Mask="&lt;0..999999g&gt;" IncludeLiterals="DecimalSymbol" 
        ErrorText="Please input valid value !"></MaskSettings>

<ValidationSettings ErrorDisplayMode="None"></ValidationSettings>
</PropertiesTextEdit>

<HeaderStyle HorizontalAlign="Center"></HeaderStyle>

<CellStyle Font-Names="Tahoma" Font-Size="8pt"></CellStyle>
</dx:GridViewDataTextColumn>
                        <dx:GridViewDataTextColumn VisibleIndex="18" Caption="VARIANCE" Width="0px" 
                            HeaderStyle-HorizontalAlign="Center" CellStyle-HorizontalAlign="Center">                            
<HeaderStyle HorizontalAlign="Center"></HeaderStyle>

<CellStyle HorizontalAlign="Center"></CellStyle>
                        </dx:GridViewDataTextColumn>
                        <dx:GridViewDataTextColumn VisibleIndex="19" Caption="VARIANCE %" Width="0px">
                        </dx:GridViewDataTextColumn>
                        <dx:GridViewDataTextColumn VisibleIndex="17" Caption="FORECAST N" 
                            FieldName="Amount" Width="0px">
                        </dx:GridViewDataTextColumn>
                        <dx:GridViewDataTextColumn VisibleIndex="20" Caption="FORECAST N+1" 
                            FieldName="ForecastN1" Width="80px" HeaderStyle-HorizontalAlign="Center">
<HeaderStyle HorizontalAlign="Center"></HeaderStyle>

                            <CellStyle Font-Names="Tahoma" Font-Size="8pt">
                            </CellStyle>
                        </dx:GridViewDataTextColumn>
                        <dx:GridViewDataTextColumn VisibleIndex="21" Caption="FORECAST N+2" 
                            FieldName="ForecastN2" Width="80px" HeaderStyle-HorizontalAlign="Center">
<HeaderStyle HorizontalAlign="Center"></HeaderStyle>

                            <CellStyle Font-Names="Tahoma" Font-Size="8pt">
                            </CellStyle>
                        </dx:GridViewDataTextColumn>
                        <dx:GridViewDataTextColumn VisibleIndex="22" Caption="FORECAST N+3" 
                            FieldName="ForecastN3" Width="80px" HeaderStyle-HorizontalAlign="Center">
<HeaderStyle HorizontalAlign="Center"></HeaderStyle>

                            <CellStyle Font-Names="Tahoma" Font-Size="8pt">
                            </CellStyle>
                        </dx:GridViewDataTextColumn>

                        <dx:GridViewDataTextColumn VisibleIndex="9" Caption="WEEK 1 FIRM QTY" 
                            HeaderStyle-HorizontalAlign="Center" >                            
<HeaderStyle HorizontalAlign="Center"></HeaderStyle>
                        </dx:GridViewDataTextColumn>

                        <dx:GridViewDataTextColumn VisibleIndex="10" Caption="WEEK 2 FIRM QTY" 
                            HeaderStyle-HorizontalAlign="Center">
<HeaderStyle HorizontalAlign="Center"></HeaderStyle>
                        </dx:GridViewDataTextColumn>

                        <dx:GridViewDataTextColumn VisibleIndex="11" Caption="WEEK 3 FIRM QTY" 
                            HeaderStyle-HorizontalAlign="Center" CellStyle-HorizontalAlign="Center" 
                            Visible="True">
                            <HeaderStyle HorizontalAlign="Center"></HeaderStyle>
                            <CellStyle Font-Names="Tahoma" Font-Size="8pt">
                            </CellStyle>
                        </dx:GridViewDataTextColumn>  
                        <dx:GridViewDataTextColumn VisibleIndex="12" Caption="WEEK 4 FIRM QTY" 
                            HeaderStyle-HorizontalAlign="Center" Visible="True">
                            <PropertiesTextEdit DisplayFormatString="{0:n2}">
                                <MaskSettings ErrorText="Please input valid value !"
                                Mask="<0..999999g>" IncludeLiterals="DecimalSymbol" />
                            </PropertiesTextEdit>                            

<HeaderStyle HorizontalAlign="Center"></HeaderStyle>

                            <CellStyle Font-Names="Tahoma" Font-Size="8pt">
                            </CellStyle>
                        </dx:GridViewDataTextColumn>  
                        <dx:GridViewDataTextColumn VisibleIndex="13" Caption="WEEK 5 FIRM QTY" 
                            HeaderStyle-HorizontalAlign="Center">
                            <PropertiesTextEdit DisplayFormatString="{0:n2}">
                                <MaskSettings ErrorText="Please input valid value !"
                                Mask="<0..999999g>" IncludeLiterals="DecimalSymbol" />
                            </PropertiesTextEdit>                            

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
        <table style="width: 100%;">
            <tr>
                <td align="left">
                <dx:ASPxButton ID="btnSubMenu" runat="server" Text="SUB MENU"
                    Font-Names="Tahoma" Width="85px" Font-Size="8pt" 
        HorizontalAlign="Left">
                </dx:ASPxButton>
                    &nbsp;
                </td>
                <td>
                <dx:ASPxTextBox ID="txtHeijunka" runat="server" ClientInstanceName="txtHeijunka" 
                    Width="0px" BackColor="White" ForeColor="White">
                    <Border BorderColor="White" />
                </dx:ASPxTextBox>                          
                </td>
                <td>
                <dx:ASPxTextBox ID="txtMode" runat="server" ClientInstanceName="txtMode" 
                    Width="0px" BackColor="White" ForeColor="White">
                    <Border BorderColor="White" />
                </dx:ASPxTextBox>                          
                </td>
                <td align="right">
                <dx:ASPxButton ID="btnApprove" runat="server" Text="APPROVE"
                    Font-Names="Tahoma"
                    Width="90px" AutoPostBack="False" Font-Size="8pt" 
                    ClientInstanceName="btnApprove" Enabled="False">
                 <ClientSideEvents Click="function(s, e) {                    
                    lblInfo.SetText('');
                    var pFlag = btnApprove.GetText();
                    ButtonApprove.PerformCallback(pFlag);
                   
                }" /> 
                </dx:ASPxButton>
                <dx:ASPxButton ID="btnRevise" runat="server" Text="UPLOAD"
                    Font-Names="Tahoma"
                    Width="90px" AutoPostBack="False" Font-Size="8pt">
                </dx:ASPxButton>
                                                    <dx:ASPxButton ID="btnHeijunka" 
                    runat="server" Text="DOWNLOAD UPLOAD ERROR LIST"
                                                        Font-Names="Tahoma" Width="85px" 
                    AutoPostBack="False" Font-Size="8pt" 
                                                        Visible="False">
                                                        <ClientSideEvents Click="function(s, e) {
                                                            txtHeijunka.SetText('Heijunka')
                                                            grid.UpdateEdit();                                                            
                                                            grid.PerformCallback('loadHeijunka');
                                                            lblInfo.SetText('');
                                                        }" />
                                                    </dx:ASPxButton>
                <dx:ASPxButton ID="btnDelete" runat="server" Text="DELETE"                              
                    Font-Names="Tahoma" Width="80px" AutoPostBack="False" Font-Size="8pt" 
                    ClientInstanceName="btnDelete" Enabled="False">
                    <ClientSideEvents Click="function(s, e) {
                        up_delete();
                    }" />
                </dx:ASPxButton>
                <dx:ASPxButton ID="btnSubmit" runat="server" Text="SAVE"
                    Font-Names="Tahoma"
                    Width="90px" AutoPostBack="False" Font-Size="8pt" ClientInstanceName="btnSubmit">
                    <ClientSideEvents Click="function(s, e) {
                        validasubmit();
                        grid.UpdateEdit();
                        grid.PerformCallback('load');
                        txtHeijunka.SetText('');
                    }" />
                </dx:ASPxButton>
                    &nbsp;
                </td>
            </tr>
            </table>
        <tr>
            <td valign="top" align="left">
            </td>
            <td valign="top" align="right" style="width: 50px;">
            </td>
            <td valign="top" align="right" style="width: 50px;">      
            </td>
            <td valign="top" align="right" style="width: 50px;">      
            </td>
            
            <td valign="top" align="right" style="width: 50px;">
            </td>
            <td valign="top" align="right" style="width: 50px;">
            </td>
            <td valign="top" align="right" style="width: 50px;">
            </td>            
            <td align="right" style="width:80px;">                                   
            </td>
            <td valign="top" align="right" style="width: 50px;">
            </td>
        </tr>
    </table>
                    
    <dx:ASPxGlobalEvents ID="ge" runat="server" >
        <ClientSideEvents ControlsInitialized="function(s, e) {
	        OnControlsInitializedSplitter();
	        OnControlsInitializedGrid();
        }" />
    </dx:ASPxGlobalEvents>

    <dx:ASPxCallback ID="ButtonDelete" runat="server" ClientInstanceName = "ButtonDelete">
        <ClientSideEvents EndCallback="function(s, e) {
            var pMsg = s.cpMessage;
            if (pMsg != '') {
                if (pMsg.substring(1,5) == '1003') {
                    lblInfo.GetMainElement().style.color = 'Blue';
                    clear();
                    grid.PerformCallback('load');
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
                if (pMsg.substring(1,5) == '1007' || pMsg.substring(1,5) == '1011') {
                    lblInfo.GetMainElement().style.color = 'Blue';
                } else {
                    lblInfo.GetMainElement().style.color = 'Red';
                }
                lblInfo.SetText(pMsg);
            } else {
                lblInfo.SetText('');
            }
            btnApprove.SetText(s.cpButton);
            
            
            if(s.cpButton == 'UNAPPROVE'){            
                btnSubmit.SetEnabled(false);
                btnDelete.SetEnabled(false);                
            }else{                
                btnSubmit.SetEnabled(true); 
                btnDelete.SetEnabled(true);                                 
            }

            delete s.cpButton;
            delete s.cpMessage;
        }" />
    </dx:ASPxCallback>
</asp:Content>

