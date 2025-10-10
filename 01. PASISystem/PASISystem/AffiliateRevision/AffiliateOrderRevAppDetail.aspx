<%@ Page Title="" Language="vb" AutoEventWireup="false" MasterPageFile="~/SiteContent.Master" CodeBehind="AffiliateOrderRevAppDetail.aspx.vb" Inherits="PASISystem.AffiliateOrderRevAppDetail" %>
<%@ Register Assembly="DevExpress.Web.v14.1, Version=14.1.6.0, Culture=neutral, PublicKeyToken=b88d1754d700e49a"
    Namespace="DevExpress.Web.ASPxGridView.Export" TagPrefix="dx" %>
<%@ Register Assembly="DevExpress.Web.v14.1, Version=14.1.6.0, Culture=neutral, PublicKeyToken=b88d1754d700e49a"
    Namespace="DevExpress.Web.ASPxGridView" TagPrefix="dx" %>
<%@ Register Assembly="DevExpress.Web.v14.1, Version=14.1.6.0, Culture=neutral, PublicKeyToken=b88d1754d700e49a"
    Namespace="DevExpress.Web.ASPxEditors" TagPrefix="dx" %>
<%@ Register Assembly="DevExpress.Web.v14.1, Version=14.1.6.0, Culture=neutral, PublicKeyToken=b88d1754d700e49a"
    Namespace="DevExpress.Web.ASPxGlobalEvents" TagPrefix="dx" %>
<%@ Register Assembly="DevExpress.Web.v14.1, Version=14.1.6.0, Culture=neutral, PublicKeyToken=b88d1754d700e49a"
    Namespace="DevExpress.Web.ASPxPanel" TagPrefix="dx" %>
<%@ Register Assembly="DevExpress.Web.v14.1, Version=14.1.6.0, Culture=neutral, PublicKeyToken=b88d1754d700e49a"
    Namespace="DevExpress.Web.ASPxRoundPanel" TagPrefix="dx" %>
<%@ Register Assembly="DevExpress.Web.v14.1, Version=14.1.6.0, Culture=neutral, PublicKeyToken=b88d1754d700e49a"
    Namespace="DevExpress.Web.ASPxCallback" TagPrefix="dx" %>
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
            height = height - (height * 60 / 100)
            grid.SetHeight(height);
        }

        function OnBatchEditStartEditing(s, e) {
            currentColumnName = e.focusedColumn.fieldName;

            if (currentColumnName == "NoUrut" || currentColumnName == "PartNo" || currentColumnName == "PartName" || currentColumnName == "KanbanCls"
            || currentColumnName == "Description" || currentColumnName == "MOQ" || currentColumnName == "QtyBox" || currentColumnName == "Maker"
            || currentColumnName == "MonthlyProductionCapacity" || currentColumnName == "BYWHAT"
            || currentColumnName == "CurrAff" || currentColumnName == "PriceAff" || currentColumnName == "AmountAff"
            || currentColumnName == "CurrSupp" || currentColumnName == "PriceSupp" || currentColumnName == "AmountSupp"
            || currentColumnName == "ForecastN1" || currentColumnName == "ForecastN2" || currentColumnName == "ForecastN3"
            || currentColumnName == "Sort") {
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
        var selection = [];
        var cells = [];
        function SetCellColor(color) {
            var str = "";
            for (var i = 0; i < selection.length; i++) {
                str += "|" + selection[i][0] + "," + selection[i][1];
            }
            grid.PerformCallback("C|" + color + str);
            cells = [];
            selection = [];
        }
        function OnBatchEditEndEditing(s, e) {
            window.setTimeout(function () {
                var period = txtPeriod.GetText();
                var nBulan = period.substring(0, 3);
                var tahun = period.substring(4, 8);

                var kabisat;
                var hariIsi;

                //Januari, Maret, May, July, Aug, Oct, Dec --> 31
                if (nBulan == "Jan" || nBulan == "Mar" || nBulan == "May" || nBulan == "Jul" || nBulan == "Aug" || nBulan == "Oct" || nBulan == "Dec") {
                    hariIsi = 31;
                }
                //April, Jun, Sep, Nov --> 30
                if (nBulan == "Apr" || nBulan == "Jun" || nBulan == "Sep" || nBulan == "Nov") {
                    hariIsi = 30;
                }
                //Februari --> ??
                if (nBulan == "1") {
                    kabisat = tahun % 4;
                    if (kabisat = 0) {
                        hariIsi = 29;
                    } else {
                        hariIsi = 28;
                    }
                }

                //s.batchEditApi.SetCellValue(e.visibleIndex, "AmountSupp", priceSupp * Qty);
                var total = 0;
                for (i = 1; i <= hariIsi; i++) {
                    var day = "DeliveryD" + i;
                    var qtyHarian = s.batchEditApi.GetCellValue(e.visibleIndex, day);
                    total = total + parseInt(qtyHarian);
                    s.batchEditApi.SetCellValue(e.visibleIndex, "POQty", total);
                }
                var priceAff = s.batchEditApi.GetCellValue(e.visibleIndex, "PriceAff");
                var Qty = s.batchEditApi.GetCellValue(e.visibleIndex, "POQty");
                var RemainingQty = s.batchEditApi.GetCellValue(e.visibleIndex - 1, "MonthlyProductionCapacity");
                var priceSupp = s.batchEditApi.GetCellValue(e.visibleIndex, "PriceSupp");
                var qtyBox = s.batchEditApi.GetCellValue(e.visibleIndex - 1, "QtyBox");

                s.batchEditApi.SetCellValue(e.visibleIndex, "AmountAff", priceAff * Qty);
                s.batchEditApi.SetCellValue(e.visibleIndex, "Flag", 1);
            }, 10);
        } 
    </script>
</asp:Content>
<asp:Content ID="Content2" ContentPlaceHolderID="Content" runat="server">
    <table style="width: 100%;">
        <tr>
            <td width="50%">
                <table style="width: 100%;">
                    <tr>
                        <td valign="top" width="60%">
                            <table style="border-width: 1pt thin thin thin; border-style: ridge; border-color: #9598A1;
                                width: 100%;">
                                <tr>
                                    <td height="30">
                                        <table id="Table1">
                                            <tr>
                                                <td style="width: 5px;">
                                                    &nbsp;
                                                </td>
                                                <td align="left" valign="middle" width="120px">
                                                    <dx:aspxlabel id="ASPxLabel3" runat="server" text="PERIOD" font-names="Tahoma" font-size="8pt"
                                                        width="100%">
                                                    </dx:aspxlabel>
                                                </td>
                                                <td align="left" valign="middle" style="width: 0">
                                                    <dx:aspxtextbox id="txtPeriod" runat="server" width="150px" height="20px" clientinstancename="txtPeriod"
                                                        font-names="Tahoma" font-size="8pt" maxlength="10" readonly="True">
                                                        <ClientSideEvents TextChanged="function(s, e) {
                                                            grid.SetFocusedRowIndex(-1);
                                                            grid.PerformCallback('kosong');
	                                                        lblErrMsg.SetText('');
                                                        }" />
                                                        <ReadOnlyStyle BackColor="#CCCCCC">
                                                        </ReadOnlyStyle>
                                                    </dx:aspxtextbox>
                                                </td>
                                                <td align="left" valign="middle" width="150px" style="width: 75px">
                                                    &nbsp;
                                                </td>
                                                <td align="left" valign="middle" width="150px">
                                                    &nbsp;
                                                </td>
                                                <td align="left" valign="middle" width="75px">
                                                    &nbsp;
                                                </td>
                                                <td align="left" valign="middle" width="165px">
                                                    &nbsp;
                                                </td>
                                                <td align="left" valign="middle" width="165px">
                                                    &nbsp;
                                                </td>
                                            </tr>
                                            <tr>
                                                <td style="width: 5px;">
                                                </td>
                                                <td align="left" valign="middle" width="120px">
                                                    <dx:aspxlabel id="ASPxLabel11" runat="server" text="PO REVISION NO." font-names="Tahoma"
                                                        font-size="8pt" width="75px">
                                                    </dx:aspxlabel>
                                                </td>
                                                <td align="left" valign="middle" width="250px" colspan="2">
                                                    <dx:aspxtextbox id="txtPORev" runat="server" width="220px" height="20px" clientinstancename="txtPORev"
                                                        font-names="Tahoma" font-size="8pt" maxlength="10" readonly="True">
                                                        <ClientSideEvents TextChanged="function(s, e) {
                                                            grid.SetFocusedRowIndex(-1);
                                                            grid.PerformCallback('kosong');
	                                                        lblErrMsg.SetText('');
                                                        }" />
                                                        <ReadOnlyStyle BackColor="#CCCCCC">
                                                        </ReadOnlyStyle>
                                                    </dx:aspxtextbox>
                                                </td>
                                                <td align="left" valign="middle">
                                                    &nbsp;
                                                </td>
                                                <td align="left" valign="middle">
                                                    &nbsp;
                                                </td>
                                                <td align="left" valign="middle">
                                                    &nbsp;
                                                </td>
                                                <td align="left" valign="middle">
                                                    &nbsp;
                                                </td>
                                            </tr>
                                            <tr>
                                                <td style="width: 5px;">
                                                    &nbsp;
                                                </td>
                                                <td align="left" valign="middle" width="120px">
                                                    <dx:aspxlabel id="ASPxLabel5" runat="server" text="PO NO." font-names="Tahoma" font-size="8pt"
                                                        width="75px">
                                                    </dx:aspxlabel>
                                                </td>
                                                <td align="left" valign="middle" width="150px">
                                                    <dx:aspxtextbox id="txtPONo" runat="server" width="150px" height="20px" clientinstancename="txtPONo"
                                                        font-names="Tahoma" font-size="8pt" maxlength="10" readonly="True">
                                                        <ClientSideEvents TextChanged="function(s, e) {
                                                            grid.SetFocusedRowIndex(-1);
                                                            grid.PerformCallback('kosong');
	                                                        lblErrMsg.SetText('');
                                                        }" />
                                                        <ReadOnlyStyle BackColor="#CCCCCC">
                                                        </ReadOnlyStyle>
                                                    </dx:aspxtextbox>
                                                </td>
                                                <td align="left" valign="middle">
                                                    <dx:aspxlabel id="ASPxLabel8" runat="server" text="COMMERCIAL" font-names="Tahoma"
                                                        font-size="8pt">
                                                    </dx:aspxlabel>
                                                </td>
                                                <td align="left" valign="middle">
                                                    <dx:aspxtextbox id="txtCommercial" runat="server" width="165px" height="20px" clientinstancename="txtCommercial"
                                                        font-names="Tahoma" font-size="8pt" maxlength="10" readonly="True">
                                                        <ClientSideEvents TextChanged="function(s, e) {
                                                            grid.SetFocusedRowIndex(-1);
                                                            grid.PerformCallback('kosong');
	                                                        lblErrMsg.SetText('');
                                                        }" />
                                                        <ReadOnlyStyle BackColor="#CCCCCC">
                                                        </ReadOnlyStyle>
                                                    </dx:aspxtextbox>
                                                </td>
                                                <td align="left" valign="middle">
                                                    &nbsp;
                                                </td>
                                                <td align="left" valign="middle">
                                                    &nbsp;
                                                </td>
                                                <td align="left" valign="middle">
                                                    &nbsp;
                                                </td>
                                            </tr>
                                            <tr>
                                                <td style="width: 5px;">
                                                </td>
                                                <td align="left" valign="middle" width="80px">
                                                    <dx:aspxlabel id="ASPxLabel1" runat="server" text="AFFILIATE" font-names="Tahoma"
                                                        font-size="8pt" width="100%">
                                                    </dx:aspxlabel>
                                                </td>
                                                <td align="left" valign="middle" width="150px">
                                                    <dx:aspxtextbox id="txtAffiliateID" runat="server" width="150px" height="20px" clientinstancename="txtAffiliateID"
                                                        font-names="Tahoma" font-size="8pt" maxlength="10" readonly="True">
                                                        <ClientSideEvents TextChanged="function(s, e) {
                                                            grid.SetFocusedRowIndex(-1);
                                                            grid.PerformCallback('kosong');
	                                                        lblErrMsg.SetText('');
                                                        }" />
                                                        <ReadOnlyStyle BackColor="#CCCCCC">
                                                        </ReadOnlyStyle>
                                                    </dx:aspxtextbox>
                                                </td>
                                                <td align="left" valign="middle" colspan="2">
                                                    <dx:aspxtextbox id="txtAffiliateName" runat="server" clientinstancename="txtAffiliateName"
                                                        width="100%" height="20px" readonly="True" font-names="Tahoma" font-size="8pt">
                                                        <ReadOnlyStyle BackColor="#CCCCCC">
                                                        </ReadOnlyStyle>
                                                    </dx:aspxtextbox>
                                                </td>
                                                <td align="left" valign="middle" width="80px">
                                                    <dx:aspxlabel id="ASPxLabel10" runat="server" text="PO KANBAN" font-names="Tahoma"
                                                        font-size="8pt" width="80px">
                                                    </dx:aspxlabel>
                                                </td>
                                                <td align="left" valign="top" width="200px">
                                                    <dx:aspxradiobuttonlist id="rblPOKanban" runat="server" repeatdirection="Horizontal"
                                                        width="200px" clientinstancename="rblPOKanban" font-names="Tahoma"
                                                        font-size="8pt" height="5px" ReadOnly="True">
                                                        <Items>
                                                            <dx:ListEditItem Text="YES" Value="1" />
                                                            <dx:ListEditItem Text="NO" Value="0" />
                                                        </Items>
                                                        <Border BorderStyle="None"></Border>
                                                    </dx:aspxradiobuttonlist>
                                                </td>
                                                <td align="left" valign="top" width="300px">
                                                    &nbsp;
                                                </td>
                                            </tr>
                                            <tr>
                                                <td style="width: 5px;">
                                                    &nbsp;
                                                </td>
                                                <td align="left" valign="middle" width="80px">
                                                    <dx:aspxlabel id="ASPxLabel2" runat="server" text="SUPPLIER CODE/NAME" font-names="Tahoma"
                                                        font-size="8pt" width="150px">
                                                    </dx:aspxlabel>
                                                </td>
                                                <td align="left" valign="middle" width="150px">
                                                    <dx:aspxtextbox id="txtSupplierCode" runat="server" width="150px" height="20px" clientinstancename="txtSupplierCode"
                                                        font-names="Tahoma" font-size="8pt" maxlength="10" readonly="True">
                                                        <ClientSideEvents TextChanged="function(s, e) {
                                                            grid.SetFocusedRowIndex(-1);
                                                            grid.PerformCallback('kosong');
	                                                        lblErrMsg.SetText('');
                                                        }" />
                                                        <ReadOnlyStyle BackColor="#CCCCCC">
                                                        </ReadOnlyStyle>
                                                    </dx:aspxtextbox>
                                                </td>
                                                <td align="left" valign="middle" colspan="2">
                                                    <dx:aspxtextbox id="txtSupplierName" runat="server" clientinstancename="txtSupplierName"
                                                        width="100%" height="20px" readonly="True" font-names="Tahoma" font-size="8pt">
                                                        <ReadOnlyStyle BackColor="#CCCCCC">
                                                        </ReadOnlyStyle>
                                                    </dx:aspxtextbox>
                                                </td>
                                                <td align="left" valign="middle" width="80px">
                                                    <dx:aspxlabel id="ASPxLabel7" runat="server" text="SUPPLIER REMARKS" 
                                                        font-names="Tahoma" font-size="8pt"
                                                        width="100%">
                                                    </dx:aspxlabel>
                                                </td>
                                                <td align="left" valign="top" width="200px">
                                        <dx:ASPxMemo ID="txtRemarks" runat="server" ClientInstanceName="txtRemarks" 
                                            Font-Names="Tahoma" Font-Size="8" Height="50px" MaxLength="200" ReadOnly="True" 
                                            Width="200px">
                                            <readonlystyle backcolor="#CCCCCC">
                                            </readonlystyle>
                                        </dx:ASPxMemo>
                                                </td>
                                                <td align="left" valign="bottom" width="300px">
                                                    <table>
                                                        <tr>
                                                            <td align="right" valign="middle">
                                                                <dx:aspxbutton id="btnRefresh" runat="server" text="SEARCH" font-names="Tahoma" width="85px"
                                                                    autopostback="False" font-size="8pt" clientinstancename="btnRefresh">
                                                                    <ClientSideEvents Click="function(s, e) {
                                                                        var pDate = txtPeriod.GetText();
                                                                        var pPORevNo = txtPORev.GetText();
                                                                        var pPONo = txtPONo.GetText();
                                                                        var pAffCode= txtAffiliateID.GetText();
                                                                        var pSuppCode = txtSupplierCode.GetText();
                                                                        var pComm = txtCommercial.GetText();
                                                                        var pKanban = rblPOKanban.GetValue();
                                                                                                                                                
                                                                        grid.PerformCallback('load' + '|' + pDate + '|' + pPORevNo + '|' + pPONo + '|' + pAffCode + '|' + pSuppCode + '|' + pKanban);
                                                                        lblInfo.SetText('');
                                                                    }" />
                                                                </dx:aspxbutton>
                                                            </td>
                                                            <td align="right" valign="middle">
                                                                <dx:aspxbutton id="btnClear" runat="server" text="CLEAR" font-names="Tahoma" width="85px"
                                                                    autopostback="False" font-size="8pt" clientinstancename="btnClear">
                                                                    <ClientSideEvents Click="function(s, e) {
                                                                        txtPeriod.SetText('');
                                                                        txtPORev.SetText('');
                                                                        txtAffiliateID.SetText('');
                                                                        txtAffiliateName.SetText('');
                                                                        txtPONo.SetText('');
                                                                        txtCommercial.SetText('');
                                                                        txtSupplierCode.SetText('');
                                                                        txtSupplierName.SetText('');
                                                                        rblPOKanban.SetValue(2);
                                                                        txtRemarks.SetText('');
                                                                        txtEntryDate.SetText('');
                                                                        txtEntryUser.SetText('');
                                                                        txtAffAppDate.SetText('');
                                                                        txtAffAppUser.SetText('');
                                                                        txtSuppPendDate.SetText('');
                                                                        txtSuppPendUser.SetText('');
                                                                        txtSuppUnpDate.SetText('');
                                                                        txtSuppUnpUser.SetText('');
                                                                        txtSendDate.SetText('');
                                                                        txtSendUser.SetText('');
                                                                        txtPASIAppDate.SetText('');
                                                                        txtPASIAppUser.SetText('');
                                                                        txtSuppAppDate.SetText('');
                                                                        txtSuppAppUser.SetText('');
                                                                        txtAffFinalAppDate.SetText('');
                                                                        txtAffFinalAppUser.SetText('');
                                                                        lblInfo.SetText('');
                                                                        grid.SetFocusedRowIndex(-1);
                                                                        grid.PerformCallback('kosong');
                                                                    }" />
                                                                </dx:aspxbutton>
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
            </td>
        </tr>
    </table>
    <div style="height: 1px;">
    </div>
    <table style="width: 100%; height: 15px;">
        <tr>
            <td colspan="8" height="15">
                <%--error message--%>
                <table id="tblMsg" style="border-width: thin; border-style: inset hidden ridge hidden;
                    border-color: #9598A1; width: 100%; height: 15px;">
                    <tr>
                        <td align="left" valign="middle" height="15px">
                            <dx:aspxlabel id="lblInfo" runat="server" text="[lblinfo]" font-names="Tahoma" clientinstancename="lblInfo"
                                font-bold="True" font-italic="True" font-size="8pt">
                            </dx:aspxlabel>
                        </td>
                    </tr>
                </table>
            </td>
        </tr>
    </table>
    <div style="height: 1px;">
    </div>
    <table style="width: 100%;">
        <tr>
            <td valign="top" colspan="6">
                <dx:aspxroundpanel id="ASPxRoundPanel1" runat="server" headertext="PO STATUS" showcollapsebutton="true"
                    view="GroupBox" width="100%" font-names="Tahoma" font-size="8pt">
                    <PanelCollection>
                        <dx:PanelContent ID="PanelContent1" runat="server">
                            <table id="Table2">
                                <tr>
                                    <td align="left" valign="middle">
                                        <dx:ASPxLabel ID="ASPxLabel18" runat="server" Font-Names="Tahoma" Font-Size="8pt"
                                            Text="(1) AFFILIATE ENTRY">
                                        </dx:ASPxLabel>
                                    </td>
                                    <td align="left" valign="middle">
                                        <dx:ASPxTextBox ID="txtEntryDate" runat="server" ClientInstanceName="txtEntryDate"
                                            Font-Names="Tahoma" Font-Size="8pt" Height="20px" MaxLength="10" ReadOnly="True"
                                            Width="150px">
                                            <ClientSideEvents TextChanged="function(s, e) {
                                                grid.SetFocusedRowIndex(-1);
                                                grid.PerformCallback('kosong');
	                                            lblErrMsg.SetText('');
                                            }" />
                                            <ReadOnlyStyle BackColor="#CCCCCC">
                                            </ReadOnlyStyle>
                                        </dx:ASPxTextBox>
                                    </td>
                                    <td align="left" valign="middle">
                                        <dx:ASPxTextBox ID="txtEntryUser" runat="server" ClientInstanceName="txtEntryUser"
                                            Font-Names="Tahoma" Font-Size="8pt" Height="20px" MaxLength="10" ReadOnly="True"
                                            Width="180px">
                                            <ClientSideEvents TextChanged="function(s, e) {
                                                grid.SetFocusedRowIndex(-1);
                                                grid.PerformCallback('kosong');
	                                            lblErrMsg.SetText('');
                                            }" />
                                            <ReadOnlyStyle BackColor="#CCCCCC">
                                            </ReadOnlyStyle>
                                        </dx:ASPxTextBox>
                                    </td>
                                    <td align="left" valign="middle">
                                        &nbsp;
                                    </td>
                                    <td align="left" valign="middle">
                                        &nbsp;
                                    </td>
                                    <td align="left" valign="middle">
                                        <dx:ASPxLabel ID="ASPxLabel14" runat="server" Font-Names="Tahoma" Font-Size="8pt"
                                            Text="(5) SUPPLIER PENDING (PARTIAL)">
                                        </dx:ASPxLabel>
                                    </td>
                                    <td align="left" valign="middle">
                                        <dx:ASPxTextBox ID="txtSuppPendDate" runat="server" ClientInstanceName="txtSuppPendDate"
                                            Font-Names="Tahoma" Font-Size="8pt" Height="20px" MaxLength="10" ReadOnly="True"
                                            Width="150px">
                                            <ClientSideEvents TextChanged="function(s, e) {
                                                grid.SetFocusedRowIndex(-1);
                                                grid.PerformCallback('kosong');
	                                            lblErrMsg.SetText('');
                                            }" />
                                            <ReadOnlyStyle BackColor="#CCCCCC">
                                            </ReadOnlyStyle>
                                        </dx:ASPxTextBox>
                                    </td>
                                    <td align="left" valign="middle">
                                        <dx:ASPxTextBox ID="txtSuppPendUser" runat="server" ClientInstanceName="txtSuppPendUser"
                                            Font-Names="Tahoma" Font-Size="8pt" Height="20px" MaxLength="10" ReadOnly="True"
                                            Width="180px">
                                            <ClientSideEvents TextChanged="function(s, e) {
                                                grid.SetFocusedRowIndex(-1);
                                                grid.PerformCallback('kosong');
	                                            lblErrMsg.SetText('');
                                            }" />
                                            <ReadOnlyStyle BackColor="#CCCCCC">
                                            </ReadOnlyStyle>
                                        </dx:ASPxTextBox>
                                    </td>
                                </tr>
                                <tr>
                                    <td align="left" valign="middle">
                                        <dx:ASPxLabel ID="ASPxLabel4" runat="server" Font-Names="Tahoma" Font-Size="8pt"
                                            Text="(2) AFFILIATE APPROVAL">
                                        </dx:ASPxLabel>
                                    </td>
                                    <td align="left" valign="middle">
                                        <dx:ASPxTextBox ID="txtAffAppDate" runat="server" ClientInstanceName="txtAffAppDate"
                                            Font-Names="Tahoma" Font-Size="8pt" Height="20px" MaxLength="10" ReadOnly="True"
                                            Width="150px">
                                            <ClientSideEvents TextChanged="function(s, e) {
                                                grid.SetFocusedRowIndex(-1);
                                                grid.PerformCallback('kosong');
	                                            lblErrMsg.SetText('');
                                            }" />
                                            <ReadOnlyStyle BackColor="#CCCCCC">
                                            </ReadOnlyStyle>
                                        </dx:ASPxTextBox>
                                    </td>
                                    <td align="left" valign="middle">
                                        <dx:ASPxTextBox ID="txtAffAppUser" runat="server" ClientInstanceName="txtAffAppUser"
                                            Font-Names="Tahoma" Font-Size="8pt" Height="20px" MaxLength="10" ReadOnly="True"
                                            Width="180px">
                                            <ClientSideEvents TextChanged="function(s, e) {
                                                grid.SetFocusedRowIndex(-1);
                                                grid.PerformCallback('kosong');
	                                            lblErrMsg.SetText('');
                                            }" />
                                            <ReadOnlyStyle BackColor="#CCCCCC">
                                            </ReadOnlyStyle>
                                        </dx:ASPxTextBox>
                                    </td>
                                    <td align="left" valign="middle">
                                        &nbsp;
                                    </td>
                                    <td align="left" valign="middle">
                                        &nbsp;
                                    </td>
                                    <td align="left" valign="middle">
                                        <dx:ASPxLabel ID="ASPxLabel15" runat="server" Font-Names="Tahoma" Font-Size="8pt"
                                            Text="(6) SUPPLIER UNAPPROVE">
                                        </dx:ASPxLabel>
                                    </td>
                                    <td align="left" valign="middle">
                                        <dx:ASPxTextBox ID="txtSuppUnpDate" runat="server" ClientInstanceName="txtSuppUnpDate"
                                            Font-Names="Tahoma" Font-Size="8pt" Height="20px" MaxLength="10" ReadOnly="True"
                                            Width="150px">
                                            <ClientSideEvents TextChanged="function(s, e) {
                                                grid.SetFocusedRowIndex(-1);
                                                grid.PerformCallback('kosong');
	                                            lblErrMsg.SetText('');
                                            }" />
                                            <ReadOnlyStyle BackColor="#CCCCCC">
                                            </ReadOnlyStyle>
                                        </dx:ASPxTextBox>
                                    </td>
                                    <td align="left" valign="middle">
                                        <dx:ASPxTextBox ID="txtSuppUnpUser" runat="server" ClientInstanceName="txtSuppUnpUser"
                                            Font-Names="Tahoma" Font-Size="8pt" Height="20px" MaxLength="10" ReadOnly="True"
                                            Width="180px">
                                            <ClientSideEvents TextChanged="function(s, e) {
                                                grid.SetFocusedRowIndex(-1);
                                                grid.PerformCallback('kosong');
	                                            lblErrMsg.SetText('');
                                            }" />
                                            <ReadOnlyStyle BackColor="#CCCCCC">
                                            </ReadOnlyStyle>
                                        </dx:ASPxTextBox>
                                    </td>
                                </tr>
                                <tr>
                                    <td align="left" valign="middle">
                                        <dx:ASPxLabel ID="ASPxLabel12" runat="server" Font-Names="Tahoma" Font-Size="8pt"
                                            Text="(3) PASI SEND AFFILIATE PO REVISION TO SUPPLIER">
                                        </dx:ASPxLabel>
                                    </td>
                                    <td align="left" valign="middle">
                                        <dx:ASPxTextBox ID="txtSendDate" runat="server" ClientInstanceName="txtSendDate"
                                            Font-Names="Tahoma" Font-Size="8pt" Height="20px" MaxLength="10" ReadOnly="True"
                                            Width="150px">
                                            <ClientSideEvents TextChanged="function(s, e) {
                                                grid.SetFocusedRowIndex(-1);
                                                grid.PerformCallback('kosong');
	                                            lblErrMsg.SetText('');
                                            }" />
                                            <ReadOnlyStyle BackColor="#CCCCCC">
                                            </ReadOnlyStyle>
                                        </dx:ASPxTextBox>
                                    </td>
                                    <td align="left" valign="middle">
                                        <dx:ASPxTextBox ID="txtSendUser" runat="server" ClientInstanceName="txtSendUser"
                                            Font-Names="Tahoma" Font-Size="8pt" Height="20px" MaxLength="10" ReadOnly="True"
                                            Width="180px">
                                            <ClientSideEvents TextChanged="function(s, e) {
                                                grid.SetFocusedRowIndex(-1);
                                                grid.PerformCallback('kosong');
	                                            lblErrMsg.SetText('');
                                            }" />
                                            <ReadOnlyStyle BackColor="#CCCCCC">
                                            </ReadOnlyStyle>
                                        </dx:ASPxTextBox>
                                    </td>
                                    <td align="left" valign="middle">
                                        &nbsp;
                                    </td>
                                    <td align="left" valign="middle">
                                        &nbsp;
                                    </td>
                                    <td align="left" valign="middle">
                                        <dx:ASPxLabel ID="ASPxLabel16" runat="server" Font-Names="Tahoma" Font-Size="8pt"
                                            Text="(7) PASI APPROVAL">
                                        </dx:ASPxLabel>
                                    </td>
                                    <td align="left" valign="middle">
                                        <dx:ASPxTextBox ID="txtPASIAppDate" runat="server" ClientInstanceName="txtPASIAppDate"
                                            Font-Names="Tahoma" Font-Size="8pt" Height="20px" MaxLength="10" ReadOnly="True"
                                            Width="150px">
                                            <ClientSideEvents TextChanged="function(s, e) {
                                                grid.SetFocusedRowIndex(-1);
                                                grid.PerformCallback('kosong');
	                                            lblErrMsg.SetText('');
                                            }" />
                                            <ReadOnlyStyle BackColor="#CCCCCC">
                                            </ReadOnlyStyle>
                                        </dx:ASPxTextBox>
                                    </td>
                                    <td align="left" valign="middle">
                                        <dx:ASPxTextBox ID="txtPASIAppUser" runat="server" ClientInstanceName="txtPASIAppUser"
                                            Font-Names="Tahoma" Font-Size="8pt" Height="20px" MaxLength="10" ReadOnly="True"
                                            Width="180px">
                                            <ClientSideEvents TextChanged="function(s, e) {
                                                grid.SetFocusedRowIndex(-1);
                                                grid.PerformCallback('kosong');
	                                            lblErrMsg.SetText('');
                                            }" />
                                            <ReadOnlyStyle BackColor="#CCCCCC">
                                            </ReadOnlyStyle>
                                        </dx:ASPxTextBox>
                                    </td>
                                </tr>
                                <tr>
                                    <td align="left" valign="middle">
                                        <dx:ASPxLabel ID="ASPxLabel13" runat="server" Font-Names="Tahoma" Font-Size="8pt"
                                            Text="(4) SUPPLIER APPROVAL (FULL)">
                                        </dx:ASPxLabel>
                                    </td>
                                    <td align="left" valign="middle">
                                        <dx:ASPxTextBox ID="txtSuppAppDate" runat="server" ClientInstanceName="txtSuppAppDate"
                                            Font-Names="Tahoma" Font-Size="8pt" Height="20px" MaxLength="10" ReadOnly="True"
                                            Width="150px">
                                            <ClientSideEvents TextChanged="function(s, e) {
                                                grid.SetFocusedRowIndex(-1);
                                                grid.PerformCallback('kosong');
	                                            lblErrMsg.SetText('');
                                            }" />
                                            <ReadOnlyStyle BackColor="#CCCCCC">
                                            </ReadOnlyStyle>
                                        </dx:ASPxTextBox>
                                    </td>
                                    <td align="left" valign="middle">
                                        <dx:ASPxTextBox ID="txtSuppAppUser" runat="server" ClientInstanceName="txtSuppAppUser"
                                            Font-Names="Tahoma" Font-Size="8pt" Height="20px" MaxLength="10" ReadOnly="True"
                                            Width="180px">
                                            <ClientSideEvents TextChanged="function(s, e) {
                                                grid.SetFocusedRowIndex(-1);
                                                grid.PerformCallback('kosong');
	                                            lblErrMsg.SetText('');
                                            }" />
                                            <ReadOnlyStyle BackColor="#CCCCCC">
                                            </ReadOnlyStyle>
                                        </dx:ASPxTextBox>
                                    </td>
                                    <td align="left" valign="middle">
                                        &nbsp;
                                    </td>
                                    <td align="left" valign="middle">
                                        &nbsp;
                                    </td>
                                    <td align="left" valign="middle">
                                        <dx:ASPxLabel ID="ASPxLabel17" runat="server" Font-Names="Tahoma" Font-Size="8pt"
                                            Text="(8) AFFILIATE FINAL APPROVAL">
                                        </dx:ASPxLabel>
                                    </td>
                                    <td align="left" valign="middle">
                                        <dx:ASPxTextBox ID="txtAffFinalAppDate" runat="server" ClientInstanceName="txtAffFinalAppDate"
                                            Font-Names="Tahoma" Font-Size="8pt" Height="20px" MaxLength="10" ReadOnly="True"
                                            Width="150px">
                                            <ClientSideEvents TextChanged="function(s, e) {
                                                grid.SetFocusedRowIndex(-1);
                                                grid.PerformCallback('kosong');
	                                            lblErrMsg.SetText('');
                                            }" />
                                            <ReadOnlyStyle BackColor="#CCCCCC">
                                            </ReadOnlyStyle>
                                        </dx:ASPxTextBox>
                                    </td>
                                    <td align="left" valign="middle">
                                        <dx:ASPxTextBox ID="txtAffFinalAppUser" runat="server" ClientInstanceName="txtAffFinalAppUser"
                                            Font-Names="Tahoma" Font-Size="8pt" Height="20px" MaxLength="10" ReadOnly="True"
                                            Width="180px">
                                            <ClientSideEvents TextChanged="function(s, e) {
                                                grid.SetFocusedRowIndex(-1);
                                                grid.PerformCallback('kosong');
	                                            lblErrMsg.SetText('');
                                            }" />
                                            <ReadOnlyStyle BackColor="#CCCCCC">
                                            </ReadOnlyStyle>
                                        </dx:ASPxTextBox>
                                    </td>
                                </tr>
                            </table>
                        </dx:PanelContent>
                    </PanelCollection>
                </dx:aspxroundpanel>
            </td>
        </tr>
        <tr>
            <td valign="top" align="left">
                &nbsp;
            </td>
            <td valign="top" align="left" width="30px">
                &nbsp;
            </td>
            <td valign="top" align="left" width="200px">
                &nbsp;
            </td>
            <td valign="top" align="left" width="200px">
                &nbsp;
            </td>
            <td valign="top" align="left" width="200px">
                &nbsp;
            </td>
            <td valign="top" align="right" width="30px">
                <table style="width: 100%;">
                    <tr>
                        <td align="right" valign="middle" width="30px">
                            <asp:TextBox ID="lSuuplier" runat="server" BackColor="Yellow" 
                                BorderStyle="None" ReadOnly="True" Width="30px"></asp:TextBox>
                        </td>
                        <td align="right" valign="middle" width="110PX">
                            <dx:ASPxLabel ID="ASPxLabel21" runat="server" Font-Names="Tahoma" 
                                Font-Size="8pt" Text=": EDITED BY PASI" Width="110px">
                            </dx:ASPxLabel>
                        </td>
                        <td width="3px">
                            &nbsp;</td>
                        <td align="right" valign="middle" width="30px">
                            <asp:TextBox ID="lSuuplier0" runat="server" BackColor="GreenYellow" 
                                BorderStyle="None" ReadOnly="True" Width="30px"></asp:TextBox>
                        </td>
                        <td align="right" valign="middle" width="140px">
                            <dx:ASPxLabel ID="ASPxLabel24" runat="server" Font-Names="Tahoma" 
                                Font-Size="8pt" Text=": EDITED BY SUPPLIER" Width="140px">
                            </dx:ASPxLabel>
                        </td>
                    </tr>
                </table>
            </td>
        </tr>
    </table>
    <div style="height: 1px;">
    </div>
    <table style="width: 100%;">
        <tr>
            <td colspan="8" align="left" valign="top" height="220">
                <%--Column : Grid--%>
                <dx:ASPxGridView ID="grid" runat="server" AutoGenerateColumns="False" 
                    ClientInstanceName="grid" Font-Names="Tahoma" Font-Size="8pt" 
                    KeyFieldName="PartNo1;AffiliateName" Width="100%">
                    <ClientSideEvents CallbackError="function(s, e) {e.handled = true;}" EndCallback="function(s, e) {
                        if (s.cpSearch != '') {
                            rblPOKanban.SetValue(s.cpKanban);
                            txtEntryDate.SetText(s.cpEntryDate);
                            txtEntryUser.SetText(s.cpEntryUser);
                            txtSuppPendDate.SetText(s.cpSuppAppPendingDate);
                            txtSuppPendUser.SetText(s.cpSuppAppPendingUser);
                            txtAffAppDate.SetText(s.cpAffAppDate);
                            txtAffAppUser.SetText(s.cpAffAppUser);
                            txtSuppUnpDate.SetText(s.cpSuppUnApproveDate);
                            txtSuppUnpUser.SetText(s.cpSuppUnApproveUser);
                            txtSendDate.SetText(s.cpSendDate);
                            txtSendUser.SetText(s.cpSendUser);
                            txtPASIAppDate.SetText(s.cpPASIAppDate);
                            txtPASIAppUser.SetText(s.cpPASIAppUser);
                            txtSuppAppDate.SetText(s.cpSuppAppDate);
                            txtSuppAppUser.SetText(s.cpSuppAppUser);
                            txtAffFinalAppDate.SetText(s.cpFinalAppDate);
                            txtAffFinalAppUser.SetText(s.cpFinalAppUser);
                        }
                        var pMsg = s.cpMessage;
                        if (pMsg != '') {
                            if (pMsg.substring(1,5) == '1001' || pMsg.substring(1,5) == '1002' || pMsg.substring(1,5) == '1003' || pMsg.substring(1,5) == '2001' || pMsg.substring(1,5) == '1008' || pMsg.substring(1,5) == '1009') {
                                lblInfo.GetMainElement().style.color = 'Blue';
                            } else {
                                lblInfo.GetMainElement().style.color = 'Red';
                            }
                            lblInfo.SetText(pMsg);
                        } else {
                            lblInfo.SetText('');
                        }
                        delete s.cpMessage;
                        delete s.cpSearch;
                    }" FocusedRowChanged="function(s, e) {
	                    OnGridFocusedRowChanged();
                    }" Init="OnInit" RowClick="function(s, e) {
	                    lblInfo.SetText('');
                    }" />
                    <Columns>
                        <dx:GridViewDataTextColumn Caption="NO." FieldName="NoUrut" 
                            HeaderStyle-HorizontalAlign="Center" VisibleIndex="1" Width="30px">
                            <CellStyle Font-Names="Tahoma" Font-Size="8pt">
                            </CellStyle>
                        </dx:GridViewDataTextColumn>
                        <dx:GridViewDataTextColumn Caption="PART NO." FieldName="PartNo" 
                            HeaderStyle-HorizontalAlign="Center" VisibleIndex="2" Width="90px">
                            <CellStyle Font-Names="Tahoma" Font-Size="8pt">
                            </CellStyle>
                        </dx:GridViewDataTextColumn>
                        <dx:GridViewDataTextColumn Caption="PART NAME" FieldName="PartName" 
                            HeaderStyle-HorizontalAlign="Center" VisibleIndex="3" Width="180px">
                            <CellStyle Font-Names="Tahoma" Font-Size="8pt">
                            </CellStyle>
                        </dx:GridViewDataTextColumn>
                        <dx:GridViewDataTextColumn Caption="KANBAN CLS" 
                            CellStyle-HorizontalAlign="Center" FieldName="KanbanCls" 
                            HeaderStyle-HorizontalAlign="Center" VisibleIndex="4" Width="60px">
                            <CellStyle Font-Names="Tahoma" Font-Size="8pt">
                            </CellStyle>
                        </dx:GridViewDataTextColumn>
                        <dx:GridViewDataTextColumn Caption="UOM" CellStyle-HorizontalAlign="Center" 
                            FieldName="UnitDesc" HeaderStyle-HorizontalAlign="Center" VisibleIndex="5" 
                            Width="40px">
                            <CellStyle Font-Names="Tahoma" Font-Size="8pt">
                            </CellStyle>
                        </dx:GridViewDataTextColumn>
                        <dx:GridViewDataTextColumn Caption="MOQ" FieldName="MOQ" 
                            HeaderStyle-HorizontalAlign="Center" VisibleIndex="6" Width="70px">
                            <PropertiesTextEdit DisplayFormatString="{0:n0}">
                            </PropertiesTextEdit>
                            <CellStyle Font-Names="Tahoma" Font-Size="8pt">
                            </CellStyle>
                        </dx:GridViewDataTextColumn>
                        <dx:GridViewDataTextColumn Caption="QTY/BOX" FieldName="QtyBox" 
                            HeaderStyle-HorizontalAlign="Center" VisibleIndex="7" Width="70px">
                            <PropertiesTextEdit DisplayFormatString="{0:n0}">
                            </PropertiesTextEdit>
                            <CellStyle Font-Names="Tahoma" Font-Size="8pt">
                            </CellStyle>
                        </dx:GridViewDataTextColumn>
                        <dx:GridViewDataTextColumn Caption="MAKER" FieldName="Maker" 
                            HeaderStyle-HorizontalAlign="Center" VisibleIndex="8" Width="100px">
                            <CellStyle Font-Names="Tahoma" Font-Size="8pt">
                            </CellStyle>
                        </dx:GridViewDataTextColumn>
                        <dx:GridViewDataTextColumn Caption=" " FieldName="AffiliateName" 
                            HeaderStyle-HorizontalAlign="Center" VisibleIndex="9" Width="80px">
                            <CellStyle Font-Names="Tahoma" Font-Size="8pt">
                            </CellStyle>
                        </dx:GridViewDataTextColumn>
                        <dx:GridViewDataTextColumn Caption="TOTAL FIRM QTY" FieldName="POQty" 
                            HeaderStyle-HorizontalAlign="Center" VisibleIndex="10" Width="100px">
                            <PropertiesTextEdit DisplayFormatString="{0:n0}">
                            </PropertiesTextEdit>
                            <CellStyle Font-Names="Tahoma" Font-Size="8pt">
                            </CellStyle>
                        </dx:GridViewDataTextColumn>
                        <dx:GridViewDataTextColumn Caption="TOTAL FIRM QTY" FieldName="POQtyOld" 
                            HeaderStyle-HorizontalAlign="Center" VisibleIndex="11" Width="0px">
                            <PropertiesTextEdit DisplayFormatString="{0:n0}">
                            </PropertiesTextEdit>
                            <CellStyle Font-Names="Tahoma" Font-Size="8pt">
                            </CellStyle>
                        </dx:GridViewDataTextColumn>                                               
                        <dx:GridViewDataTextColumn Caption="FORECAST N+1" FieldName="ForecastN1" 
                            HeaderStyle-HorizontalAlign="Center" VisibleIndex="15" Width="80px">
                            <PropertiesTextEdit DisplayFormatString="{0:n0}">
                            </PropertiesTextEdit>
                            <CellStyle Font-Names="Tahoma" Font-Size="8pt">
                            </CellStyle>
                        </dx:GridViewDataTextColumn>
                        <dx:GridViewDataTextColumn Caption="FORECAST N+2" FieldName="ForecastN2" 
                            HeaderStyle-HorizontalAlign="Center" VisibleIndex="16" Width="80px">
                            <PropertiesTextEdit DisplayFormatString="{0:n0}">
                            </PropertiesTextEdit>
                            <CellStyle Font-Names="Tahoma" Font-Size="8pt">
                            </CellStyle>
                        </dx:GridViewDataTextColumn>
                        <dx:GridViewDataTextColumn Caption="FORECAST N+3" FieldName="ForecastN3" 
                            HeaderStyle-HorizontalAlign="Center" VisibleIndex="17" Width="80px">
                            <PropertiesTextEdit DisplayFormatString="{0:n0}">
                            </PropertiesTextEdit>
                            <CellStyle Font-Names="Tahoma" Font-Size="8pt">
                            </CellStyle>
                        </dx:GridViewDataTextColumn>
                        <dx:GridViewBandColumn Caption="E.T.A SCHEDULE (BASED ON FIRM ORDER)" 
                            HeaderStyle-HorizontalAlign="Center" VisibleIndex="17">
                            <Columns>
                                <dx:GridViewDataTextColumn Caption="1" FieldName="DeliveryD1" 
                                    HeaderStyle-HorizontalAlign="Center" VisibleIndex="18" Width="60px">
                                    <PropertiesTextEdit DisplayFormatString="{0:n0}">
                                    </PropertiesTextEdit>
                                    <CellStyle Font-Names="Tahoma" Font-Size="8pt">
                                    </CellStyle>
                                </dx:GridViewDataTextColumn>
                                <dx:GridViewDataTextColumn Caption="1" FieldName="DeliveryD1Old" 
                                    HeaderStyle-HorizontalAlign="Center" VisibleIndex="19" Width="0px">
                                    <PropertiesTextEdit DisplayFormatString="{0:n0}">
                                    </PropertiesTextEdit>
                                    <CellStyle Font-Names="Tahoma" Font-Size="8pt">
                                    </CellStyle>
                                </dx:GridViewDataTextColumn>
                                <dx:GridViewDataTextColumn Caption="2" FieldName="DeliveryD2" 
                                    HeaderStyle-HorizontalAlign="Center" VisibleIndex="20" Width="60px">
                                    <PropertiesTextEdit DisplayFormatString="{0:n0}">
                                    </PropertiesTextEdit>
                                    <CellStyle Font-Names="Tahoma" Font-Size="8pt">
                                    </CellStyle>
                                </dx:GridViewDataTextColumn>
                                <dx:GridViewDataTextColumn Caption="2" FieldName="DeliveryD2Old" 
                                    HeaderStyle-HorizontalAlign="Center" VisibleIndex="21" Width="0px">
                                    <PropertiesTextEdit DisplayFormatString="{0:n0}">
                                    </PropertiesTextEdit>
                                    <CellStyle Font-Names="Tahoma" Font-Size="8pt">
                                    </CellStyle>
                                </dx:GridViewDataTextColumn>
                                <dx:GridViewDataTextColumn Caption="3" FieldName="DeliveryD3" 
                                    HeaderStyle-HorizontalAlign="Center" VisibleIndex="22" Width="60px">
                                    <PropertiesTextEdit DisplayFormatString="{0:n0}">
                                    </PropertiesTextEdit>
                                    <CellStyle Font-Names="Tahoma" Font-Size="8pt">
                                    </CellStyle>
                                </dx:GridViewDataTextColumn>
                                <dx:GridViewDataTextColumn Caption="3" FieldName="DeliveryD3Old" 
                                    HeaderStyle-HorizontalAlign="Center" VisibleIndex="23" Width="0px">
                                    <PropertiesTextEdit DisplayFormatString="{0:n0}">
                                    </PropertiesTextEdit>
                                    <CellStyle Font-Names="Tahoma" Font-Size="8pt">
                                    </CellStyle>
                                </dx:GridViewDataTextColumn>
                                <dx:GridViewDataTextColumn Caption="4" FieldName="DeliveryD4" 
                                    HeaderStyle-HorizontalAlign="Center" VisibleIndex="24" Width="60px">
                                    <PropertiesTextEdit DisplayFormatString="{0:n0}">
                                    </PropertiesTextEdit>
                                    <CellStyle Font-Names="Tahoma" Font-Size="8pt">
                                    </CellStyle>
                                </dx:GridViewDataTextColumn>
                                <dx:GridViewDataTextColumn Caption="4" FieldName="DeliveryD4Old" 
                                    HeaderStyle-HorizontalAlign="Center" VisibleIndex="25" Width="0px">
                                    <PropertiesTextEdit DisplayFormatString="{0:n0}">
                                    </PropertiesTextEdit>
                                    <CellStyle Font-Names="Tahoma" Font-Size="8pt">
                                    </CellStyle>
                                </dx:GridViewDataTextColumn>
                                <dx:GridViewDataTextColumn Caption="5" FieldName="DeliveryD5" 
                                    HeaderStyle-HorizontalAlign="Center" VisibleIndex="26" Width="60px">
                                    <PropertiesTextEdit DisplayFormatString="{0:n0}">
                                    </PropertiesTextEdit>
                                    <CellStyle Font-Names="Tahoma" Font-Size="8pt">
                                    </CellStyle>
                                </dx:GridViewDataTextColumn>
                                <dx:GridViewDataTextColumn Caption="5" FieldName="DeliveryD5Old" 
                                    HeaderStyle-HorizontalAlign="Center" VisibleIndex="27" Width="0px">
                                    <PropertiesTextEdit DisplayFormatString="{0:n0}">
                                    </PropertiesTextEdit>
                                    <CellStyle Font-Names="Tahoma" Font-Size="8pt">
                                    </CellStyle>
                                </dx:GridViewDataTextColumn>
                                <dx:GridViewDataTextColumn Caption="6" FieldName="DeliveryD6" 
                                    HeaderStyle-HorizontalAlign="Center" VisibleIndex="28" Width="60px">
                                    <PropertiesTextEdit DisplayFormatString="{0:n0}">
                                    </PropertiesTextEdit>
                                    <CellStyle Font-Names="Tahoma" Font-Size="8pt">
                                    </CellStyle>
                                </dx:GridViewDataTextColumn>
                                <dx:GridViewDataTextColumn Caption="6" FieldName="DeliveryD6Old" 
                                    HeaderStyle-HorizontalAlign="Center" VisibleIndex="29" Width="0px">
                                    <PropertiesTextEdit DisplayFormatString="{0:n0}">
                                    </PropertiesTextEdit>
                                    <CellStyle Font-Names="Tahoma" Font-Size="8pt">
                                    </CellStyle>
                                </dx:GridViewDataTextColumn>
                                <dx:GridViewDataTextColumn Caption="7" FieldName="DeliveryD7" 
                                    HeaderStyle-HorizontalAlign="Center" VisibleIndex="30" Width="60px">
                                    <PropertiesTextEdit DisplayFormatString="{0:n0}">
                                    </PropertiesTextEdit>
                                    <CellStyle Font-Names="Tahoma" Font-Size="8pt">
                                    </CellStyle>
                                </dx:GridViewDataTextColumn>
                                <dx:GridViewDataTextColumn Caption="7" FieldName="DeliveryD7Old" 
                                    HeaderStyle-HorizontalAlign="Center" VisibleIndex="31" Width="0px">
                                    <PropertiesTextEdit DisplayFormatString="{0:n0}">
                                    </PropertiesTextEdit>
                                    <CellStyle Font-Names="Tahoma" Font-Size="8pt">
                                    </CellStyle>
                                </dx:GridViewDataTextColumn>
                                <dx:GridViewDataTextColumn Caption="8" FieldName="DeliveryD8" 
                                    HeaderStyle-HorizontalAlign="Center" VisibleIndex="32" Width="60px">
                                    <PropertiesTextEdit DisplayFormatString="{0:n0}">
                                    </PropertiesTextEdit>
                                    <CellStyle Font-Names="Tahoma" Font-Size="8pt">
                                    </CellStyle>
                                </dx:GridViewDataTextColumn>
                                <dx:GridViewDataTextColumn Caption="8" FieldName="DeliveryD8Old" 
                                    HeaderStyle-HorizontalAlign="Center" VisibleIndex="33" Width="0px">
                                    <PropertiesTextEdit DisplayFormatString="{0:n0}">
                                    </PropertiesTextEdit>
                                    <CellStyle Font-Names="Tahoma" Font-Size="8pt">
                                    </CellStyle>
                                </dx:GridViewDataTextColumn>
                                <dx:GridViewDataTextColumn Caption="9" FieldName="DeliveryD9" 
                                    HeaderStyle-HorizontalAlign="Center" VisibleIndex="34" Width="60px">
                                    <PropertiesTextEdit DisplayFormatString="{0:n0}">
                                    </PropertiesTextEdit>
                                    <CellStyle Font-Names="Tahoma" Font-Size="8pt">
                                    </CellStyle>
                                </dx:GridViewDataTextColumn>
                                <dx:GridViewDataTextColumn Caption="9" FieldName="DeliveryD9Old" 
                                    HeaderStyle-HorizontalAlign="Center" VisibleIndex="35" Width="0px">
                                    <PropertiesTextEdit DisplayFormatString="{0:n0}">
                                    </PropertiesTextEdit>
                                    <CellStyle Font-Names="Tahoma" Font-Size="8pt">
                                    </CellStyle>
                                </dx:GridViewDataTextColumn>
                                <dx:GridViewDataTextColumn Caption="10" FieldName="DeliveryD10" 
                                    HeaderStyle-HorizontalAlign="Center" VisibleIndex="36" Width="60px">
                                    <PropertiesTextEdit DisplayFormatString="{0:n0}">
                                    </PropertiesTextEdit>
                                    <CellStyle Font-Names="Tahoma" Font-Size="8pt">
                                    </CellStyle>
                                </dx:GridViewDataTextColumn>
                                <dx:GridViewDataTextColumn Caption="10" FieldName="DeliveryD10Old" 
                                    HeaderStyle-HorizontalAlign="Center" VisibleIndex="37" Width="0px">
                                    <PropertiesTextEdit DisplayFormatString="{0:n0}">
                                    </PropertiesTextEdit>
                                    <CellStyle Font-Names="Tahoma" Font-Size="8pt">
                                    </CellStyle>
                                </dx:GridViewDataTextColumn>
                                <dx:GridViewDataTextColumn Caption="11" FieldName="DeliveryD11" 
                                    HeaderStyle-HorizontalAlign="Center" VisibleIndex="38" Width="60px">
                                    <PropertiesTextEdit DisplayFormatString="{0:n0}">
                                    </PropertiesTextEdit>
                                    <CellStyle Font-Names="Tahoma" Font-Size="8pt">
                                    </CellStyle>
                                </dx:GridViewDataTextColumn>
                                <dx:GridViewDataTextColumn Caption="11" FieldName="DeliveryD11Old" 
                                    HeaderStyle-HorizontalAlign="Center" VisibleIndex="39" Width="0px">
                                    <PropertiesTextEdit DisplayFormatString="{0:n0}">
                                    </PropertiesTextEdit>
                                    <CellStyle Font-Names="Tahoma" Font-Size="8pt">
                                    </CellStyle>
                                </dx:GridViewDataTextColumn>
                                <dx:GridViewDataTextColumn Caption="12" FieldName="DeliveryD12" 
                                    HeaderStyle-HorizontalAlign="Center" VisibleIndex="40" Width="60px">
                                    <PropertiesTextEdit DisplayFormatString="{0:n0}">
                                    </PropertiesTextEdit>
                                    <CellStyle Font-Names="Tahoma" Font-Size="8pt">
                                    </CellStyle>
                                </dx:GridViewDataTextColumn>
                                <dx:GridViewDataTextColumn Caption="12" FieldName="DeliveryD12Old" 
                                    HeaderStyle-HorizontalAlign="Center" VisibleIndex="41" Width="0px">
                                    <PropertiesTextEdit DisplayFormatString="{0:n0}">
                                    </PropertiesTextEdit>
                                    <CellStyle Font-Names="Tahoma" Font-Size="8pt">
                                    </CellStyle>
                                </dx:GridViewDataTextColumn>
                                <dx:GridViewDataTextColumn Caption="13" FieldName="DeliveryD13" 
                                    HeaderStyle-HorizontalAlign="Center" VisibleIndex="42" Width="60px">
                                    <PropertiesTextEdit DisplayFormatString="{0:n0}">
                                    </PropertiesTextEdit>
                                    <CellStyle Font-Names="Tahoma" Font-Size="8pt">
                                    </CellStyle>
                                </dx:GridViewDataTextColumn>
                                <dx:GridViewDataTextColumn Caption="13" FieldName="DeliveryD13Old" 
                                    HeaderStyle-HorizontalAlign="Center" VisibleIndex="43" Width="0px">
                                    <PropertiesTextEdit DisplayFormatString="{0:n0}">
                                    </PropertiesTextEdit>
                                    <CellStyle Font-Names="Tahoma" Font-Size="8pt">
                                    </CellStyle>
                                </dx:GridViewDataTextColumn>
                                <dx:GridViewDataTextColumn Caption="14" FieldName="DeliveryD14" 
                                    HeaderStyle-HorizontalAlign="Center" VisibleIndex="44" Width="60px">
                                    <PropertiesTextEdit DisplayFormatString="{0:n0}">
                                    </PropertiesTextEdit>
                                    <CellStyle Font-Names="Tahoma" Font-Size="8pt">
                                    </CellStyle>
                                </dx:GridViewDataTextColumn>
                                <dx:GridViewDataTextColumn Caption="14" FieldName="DeliveryD14Old" 
                                    HeaderStyle-HorizontalAlign="Center" VisibleIndex="45" Width="0px">
                                    <PropertiesTextEdit DisplayFormatString="{0:n0}">
                                    </PropertiesTextEdit>
                                    <CellStyle Font-Names="Tahoma" Font-Size="8pt">
                                    </CellStyle>
                                </dx:GridViewDataTextColumn>
                                <dx:GridViewDataTextColumn Caption="15" FieldName="DeliveryD15" 
                                    HeaderStyle-HorizontalAlign="Center" VisibleIndex="46" Width="60px">
                                    <PropertiesTextEdit DisplayFormatString="{0:n0}">
                                    </PropertiesTextEdit>
                                    <CellStyle Font-Names="Tahoma" Font-Size="8pt">
                                    </CellStyle>
                                </dx:GridViewDataTextColumn>
                                <dx:GridViewDataTextColumn Caption="15" FieldName="DeliveryD15Old" 
                                    HeaderStyle-HorizontalAlign="Center" VisibleIndex="47" Width="0px">
                                    <PropertiesTextEdit DisplayFormatString="{0:n0}">
                                    </PropertiesTextEdit>
                                    <CellStyle Font-Names="Tahoma" Font-Size="8pt">
                                    </CellStyle>
                                </dx:GridViewDataTextColumn>
                                <dx:GridViewDataTextColumn Caption="16" FieldName="DeliveryD16" 
                                    HeaderStyle-HorizontalAlign="Center" VisibleIndex="48" Width="60px">
                                    <PropertiesTextEdit DisplayFormatString="{0:n0}">
                                    </PropertiesTextEdit>
                                    <CellStyle Font-Names="Tahoma" Font-Size="8pt">
                                    </CellStyle>
                                </dx:GridViewDataTextColumn>
                                <dx:GridViewDataTextColumn Caption="16" FieldName="DeliveryD16Old" 
                                    HeaderStyle-HorizontalAlign="Center" VisibleIndex="49" Width="0px">
                                    <PropertiesTextEdit DisplayFormatString="{0:n0}">
                                    </PropertiesTextEdit>
                                    <CellStyle Font-Names="Tahoma" Font-Size="8pt">
                                    </CellStyle>
                                </dx:GridViewDataTextColumn>
                                <dx:GridViewDataTextColumn Caption="17" FieldName="DeliveryD17" 
                                    HeaderStyle-HorizontalAlign="Center" VisibleIndex="50" Width="60px">
                                    <PropertiesTextEdit DisplayFormatString="{0:n0}">
                                    </PropertiesTextEdit>
                                    <CellStyle Font-Names="Tahoma" Font-Size="8pt">
                                    </CellStyle>
                                </dx:GridViewDataTextColumn>
                                <dx:GridViewDataTextColumn Caption="17" FieldName="DeliveryD17Old" 
                                    HeaderStyle-HorizontalAlign="Center" VisibleIndex="51" Width="0px">
                                    <PropertiesTextEdit DisplayFormatString="{0:n0}">
                                    </PropertiesTextEdit>
                                    <CellStyle Font-Names="Tahoma" Font-Size="8pt">
                                    </CellStyle>
                                </dx:GridViewDataTextColumn>
                                <dx:GridViewDataTextColumn Caption="18" FieldName="DeliveryD18" 
                                    HeaderStyle-HorizontalAlign="Center" VisibleIndex="52" Width="60px">
                                    <PropertiesTextEdit DisplayFormatString="{0:n0}">
                                    </PropertiesTextEdit>
                                    <CellStyle Font-Names="Tahoma" Font-Size="8pt">
                                    </CellStyle>
                                </dx:GridViewDataTextColumn>
                                <dx:GridViewDataTextColumn Caption="18" FieldName="DeliveryD18Old" 
                                    HeaderStyle-HorizontalAlign="Center" VisibleIndex="53" Width="0px">
                                    <PropertiesTextEdit DisplayFormatString="{0:n0}">
                                    </PropertiesTextEdit>
                                    <CellStyle Font-Names="Tahoma" Font-Size="8pt">
                                    </CellStyle>
                                </dx:GridViewDataTextColumn>
                                <dx:GridViewDataTextColumn Caption="19" FieldName="DeliveryD19" 
                                    HeaderStyle-HorizontalAlign="Center" VisibleIndex="54" Width="60px">
                                    <PropertiesTextEdit DisplayFormatString="{0:n0}">
                                    </PropertiesTextEdit>
                                    <CellStyle Font-Names="Tahoma" Font-Size="8pt">
                                    </CellStyle>
                                </dx:GridViewDataTextColumn>
                                <dx:GridViewDataTextColumn Caption="19" FieldName="DeliveryD19Old" 
                                    HeaderStyle-HorizontalAlign="Center" VisibleIndex="55" Width="0px">
                                    <PropertiesTextEdit DisplayFormatString="{0:n0}">
                                    </PropertiesTextEdit>
                                    <CellStyle Font-Names="Tahoma" Font-Size="8pt">
                                    </CellStyle>
                                </dx:GridViewDataTextColumn>
                                <dx:GridViewDataTextColumn Caption="20" FieldName="DeliveryD20" 
                                    HeaderStyle-HorizontalAlign="Center" VisibleIndex="56" Width="60px">
                                    <PropertiesTextEdit DisplayFormatString="{0:n0}">
                                    </PropertiesTextEdit>
                                    <CellStyle Font-Names="Tahoma" Font-Size="8pt">
                                    </CellStyle>
                                </dx:GridViewDataTextColumn>
                                <dx:GridViewDataTextColumn Caption="20" FieldName="DeliveryD20Old" 
                                    HeaderStyle-HorizontalAlign="Center" VisibleIndex="57" Width="0px">
                                    <PropertiesTextEdit DisplayFormatString="{0:n0}">
                                    </PropertiesTextEdit>
                                    <CellStyle Font-Names="Tahoma" Font-Size="8pt">
                                    </CellStyle>
                                </dx:GridViewDataTextColumn>
                                <dx:GridViewDataTextColumn Caption="21" FieldName="DeliveryD21" 
                                    HeaderStyle-HorizontalAlign="Center" VisibleIndex="58" Width="60px">
                                    <PropertiesTextEdit DisplayFormatString="{0:n0}">
                                    </PropertiesTextEdit>
                                    <CellStyle Font-Names="Tahoma" Font-Size="8pt">
                                    </CellStyle>
                                </dx:GridViewDataTextColumn>
                                <dx:GridViewDataTextColumn Caption="21" FieldName="DeliveryD21Old" 
                                    HeaderStyle-HorizontalAlign="Center" VisibleIndex="59" Width="0px">
                                    <PropertiesTextEdit DisplayFormatString="{0:n0}">
                                    </PropertiesTextEdit>
                                    <CellStyle Font-Names="Tahoma" Font-Size="8pt">
                                    </CellStyle>
                                </dx:GridViewDataTextColumn>
                                <dx:GridViewDataTextColumn Caption="22" FieldName="DeliveryD22" 
                                    HeaderStyle-HorizontalAlign="Center" VisibleIndex="60" Width="60px">
                                    <PropertiesTextEdit DisplayFormatString="{0:n0}">
                                    </PropertiesTextEdit>
                                    <CellStyle Font-Names="Tahoma" Font-Size="8pt">
                                    </CellStyle>
                                </dx:GridViewDataTextColumn>
                                <dx:GridViewDataTextColumn Caption="22" FieldName="DeliveryD22Old" 
                                    HeaderStyle-HorizontalAlign="Center" VisibleIndex="61" Width="0px">
                                    <PropertiesTextEdit DisplayFormatString="{0:n0}">
                                    </PropertiesTextEdit>
                                    <CellStyle Font-Names="Tahoma" Font-Size="8pt">
                                    </CellStyle>
                                </dx:GridViewDataTextColumn>
                                <dx:GridViewDataTextColumn Caption="23" FieldName="DeliveryD23" 
                                    HeaderStyle-HorizontalAlign="Center" VisibleIndex="62" Width="60px">
                                    <PropertiesTextEdit DisplayFormatString="{0:n0}">
                                    </PropertiesTextEdit>
                                    <CellStyle Font-Names="Tahoma" Font-Size="8pt">
                                    </CellStyle>
                                </dx:GridViewDataTextColumn>
                                <dx:GridViewDataTextColumn Caption="23" FieldName="DeliveryD23Old" 
                                    HeaderStyle-HorizontalAlign="Center" VisibleIndex="63" Width="0px">
                                    <PropertiesTextEdit DisplayFormatString="{0:n0}">
                                    </PropertiesTextEdit>
                                    <CellStyle Font-Names="Tahoma" Font-Size="8pt">
                                    </CellStyle>
                                </dx:GridViewDataTextColumn>
                                <dx:GridViewDataTextColumn Caption="24" FieldName="DeliveryD24" 
                                    HeaderStyle-HorizontalAlign="Center" VisibleIndex="64" Width="60px">
                                    <PropertiesTextEdit DisplayFormatString="{0:n0}">
                                    </PropertiesTextEdit>
                                    <CellStyle Font-Names="Tahoma" Font-Size="8pt">
                                    </CellStyle>
                                </dx:GridViewDataTextColumn>
                                <dx:GridViewDataTextColumn Caption="24" FieldName="DeliveryD24Old" 
                                    HeaderStyle-HorizontalAlign="Center" VisibleIndex="65" Width="0px">
                                    <PropertiesTextEdit DisplayFormatString="{0:n0}">
                                    </PropertiesTextEdit>
                                    <CellStyle Font-Names="Tahoma" Font-Size="8pt">
                                    </CellStyle>
                                </dx:GridViewDataTextColumn>
                                <dx:GridViewDataTextColumn Caption="25" FieldName="DeliveryD25" 
                                    HeaderStyle-HorizontalAlign="Center" VisibleIndex="66" Width="60px">
                                    <PropertiesTextEdit DisplayFormatString="{0:n0}">
                                    </PropertiesTextEdit>
                                    <CellStyle Font-Names="Tahoma" Font-Size="8pt">
                                    </CellStyle>
                                </dx:GridViewDataTextColumn>
                                <dx:GridViewDataTextColumn Caption="25" FieldName="DeliveryD25Old" 
                                    HeaderStyle-HorizontalAlign="Center" VisibleIndex="67" Width="0px">
                                    <PropertiesTextEdit DisplayFormatString="{0:n0}">
                                    </PropertiesTextEdit>
                                    <CellStyle Font-Names="Tahoma" Font-Size="8pt">
                                    </CellStyle>
                                </dx:GridViewDataTextColumn>
                                <dx:GridViewDataTextColumn Caption="26" FieldName="DeliveryD26" 
                                    HeaderStyle-HorizontalAlign="Center" VisibleIndex="68" Width="60px">
                                    <PropertiesTextEdit DisplayFormatString="{0:n0}">
                                    </PropertiesTextEdit>
                                    <CellStyle Font-Names="Tahoma" Font-Size="8pt">
                                    </CellStyle>
                                </dx:GridViewDataTextColumn>
                                <dx:GridViewDataTextColumn Caption="26" FieldName="DeliveryD26Old" 
                                    HeaderStyle-HorizontalAlign="Center" VisibleIndex="69" Width="0px">
                                    <PropertiesTextEdit DisplayFormatString="{0:n0}">
                                    </PropertiesTextEdit>
                                    <CellStyle Font-Names="Tahoma" Font-Size="8pt">
                                    </CellStyle>
                                </dx:GridViewDataTextColumn>
                                <dx:GridViewDataTextColumn Caption="27" FieldName="DeliveryD27" 
                                    HeaderStyle-HorizontalAlign="Center" VisibleIndex="70" Width="60px">
                                    <PropertiesTextEdit DisplayFormatString="{0:n0}">
                                    </PropertiesTextEdit>
                                    <CellStyle Font-Names="Tahoma" Font-Size="8pt">
                                    </CellStyle>
                                </dx:GridViewDataTextColumn>
                                <dx:GridViewDataTextColumn Caption="27" FieldName="DeliveryD27Old" 
                                    HeaderStyle-HorizontalAlign="Center" VisibleIndex="71" Width="0px">
                                    <PropertiesTextEdit DisplayFormatString="{0:n0}">
                                    </PropertiesTextEdit>
                                    <CellStyle Font-Names="Tahoma" Font-Size="8pt">
                                    </CellStyle>
                                </dx:GridViewDataTextColumn>
                                <dx:GridViewDataTextColumn Caption="28" FieldName="DeliveryD28" 
                                    HeaderStyle-HorizontalAlign="Center" VisibleIndex="72" Width="60px">
                                    <PropertiesTextEdit DisplayFormatString="{0:n0}">
                                    </PropertiesTextEdit>
                                    <CellStyle Font-Names="Tahoma" Font-Size="8pt">
                                    </CellStyle>
                                </dx:GridViewDataTextColumn>
                                <dx:GridViewDataTextColumn Caption="28" FieldName="DeliveryD28Old" 
                                    HeaderStyle-HorizontalAlign="Center" VisibleIndex="73" Width="0px">
                                    <PropertiesTextEdit DisplayFormatString="{0:n0}">
                                    </PropertiesTextEdit>
                                    <CellStyle Font-Names="Tahoma" Font-Size="8pt">
                                    </CellStyle>
                                </dx:GridViewDataTextColumn>
                                <dx:GridViewDataTextColumn Caption="29" FieldName="DeliveryD29" 
                                    HeaderStyle-HorizontalAlign="Center" VisibleIndex="74" Width="60px">
                                    <PropertiesTextEdit DisplayFormatString="{0:n0}">
                                    </PropertiesTextEdit>
                                    <CellStyle Font-Names="Tahoma" Font-Size="8pt">
                                    </CellStyle>
                                </dx:GridViewDataTextColumn>
                                <dx:GridViewDataTextColumn Caption="29" FieldName="DeliveryD29Old" 
                                    HeaderStyle-HorizontalAlign="Center" VisibleIndex="75" Width="0px">
                                    <PropertiesTextEdit DisplayFormatString="{0:n0}">
                                    </PropertiesTextEdit>
                                    <CellStyle Font-Names="Tahoma" Font-Size="8pt">
                                    </CellStyle>
                                </dx:GridViewDataTextColumn>
                                <dx:GridViewDataTextColumn Caption="30" FieldName="DeliveryD30" 
                                    HeaderStyle-HorizontalAlign="Center" VisibleIndex="76" Width="60px">
                                    <PropertiesTextEdit DisplayFormatString="{0:n0}">
                                    </PropertiesTextEdit>
                                    <CellStyle Font-Names="Tahoma" Font-Size="8pt">
                                    </CellStyle>
                                </dx:GridViewDataTextColumn>
                                <dx:GridViewDataTextColumn Caption="30" FieldName="DeliveryD30Old" 
                                    HeaderStyle-HorizontalAlign="Center" VisibleIndex="77" Width="0px">
                                    <PropertiesTextEdit DisplayFormatString="{0:n0}">
                                    </PropertiesTextEdit>
                                    <CellStyle Font-Names="Tahoma" Font-Size="8pt">
                                    </CellStyle>
                                </dx:GridViewDataTextColumn>
                                <dx:GridViewDataTextColumn Caption="31" FieldName="DeliveryD31" 
                                    HeaderStyle-HorizontalAlign="Center" VisibleIndex="78" Width="60px">
                                    <PropertiesTextEdit DisplayFormatString="{0:n0}">
                                    </PropertiesTextEdit>
                                    <CellStyle Font-Names="Tahoma" Font-Size="8pt">
                                    </CellStyle>
                                </dx:GridViewDataTextColumn>
                                <dx:GridViewDataTextColumn Caption="31" FieldName="DeliveryD31Old" 
                                    HeaderStyle-HorizontalAlign="Center" VisibleIndex="79" Width="0px">
                                    <PropertiesTextEdit DisplayFormatString="{0:n0}">
                                    </PropertiesTextEdit>
                                    <CellStyle Font-Names="Tahoma" Font-Size="8pt">
                                    </CellStyle>
                                </dx:GridViewDataTextColumn>
                                <dx:GridViewDataTextColumn Caption="PartNo1" FieldName="PartNo1" 
                                    HeaderStyle-HorizontalAlign="Center" VisibleIndex="80" Width="0px">
                                    <CellStyle Font-Names="Tahoma" Font-Size="8pt">
                                    </CellStyle>
                                </dx:GridViewDataTextColumn>
                            </Columns>
                            <HeaderStyle HorizontalAlign="Center" />
                        </dx:GridViewBandColumn>
                    </Columns>
                    <SettingsBehavior AllowSelectByRowClick="True" AllowSort="False" 
                        ColumnResizeMode="Control" EnableRowHotTrack="True" />
                    <SettingsPager Visible="False" Mode="ShowAllRecords">
                        <Summary AllPagesText="Page {0} of {1} " 
                            Text="Page {0} of {1} [{2} record(s)]" />
                        <Summary AllPagesText="Page {0} of {1} " 
                            Text="Page {0} of {1} [{2} record(s)]" />
                    </SettingsPager>
                    <SettingsEditing Mode="Batch" NewItemRowPosition="Bottom">
                        <BatchEditSettings ShowConfirmOnLosingChanges="False" />
                    </SettingsEditing>
                    <Settings ShowGroupButtons="False" ShowHorizontalScrollBar="True" 
                        ShowStatusBar="Hidden" ShowVerticalScrollBar="True" 
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
                <dx:ASPxButton ID="btnSubMenu" runat="server" Text="SUB MENU" Font-Names="Tahoma"
                    Width="90px" Font-Size="8pt">
                </dx:ASPxButton>
            </td>
            <td valign="top" align="right" style="width: 50px;">
            </td>
            <td valign="top" align="right" style="width: 50px;">
            </td>
            <td align="right" style="width: 80px;">
                &nbsp;</td>
            <td valign="top" align="right" style="width: 50px;">
                <dx:ASPxButton ID="btnApprove" runat="server" AutoPostBack="False" 
                    Font-Names="Tahoma" Font-Size="8pt" Text="APPROVE" Width="90px" ClientInstanceName="btnApprove">
                    <ClientSideEvents Click="function(s, e) {                    
                    lblInfo.SetText('');                    
                    ButtonApprove.PerformCallback();
                }" />
                </dx:ASPxButton>
            </td>
        </tr>
    </table>
    <dx:aspxglobalevents id="ge" runat="server">
        <ClientSideEvents ControlsInitialized="function(s, e) {
	        OnControlsInitializedSplitter();
	        OnControlsInitializedGrid();
        }" />
    </dx:aspxglobalevents>
    <dx:ASPxCallback ID="ButtonApprove" runat="server" 
        ClientInstanceName="ButtonApprove">
        <ClientSideEvents EndCallback="function(s, e) {
	        rblPOKanban.SetValue(s.cpKanban);
            txtEntryDate.SetText(s.cpEntryDate);
            txtEntryUser.SetText(s.cpEntryUser);
            txtSuppPendDate.SetText(s.cpSuppAppPendingDate);
            txtSuppPendUser.SetText(s.cpSuppAppPendingUser);
            txtAffAppDate.SetText(s.cpAffAppDate);
            txtAffAppUser.SetText(s.cpAffAppUser);
            txtSuppUnpDate.SetText(s.cpSuppUnApproveDate);
            txtSuppUnpUser.SetText(s.cpSuppUnApproveUser);
            txtSendDate.SetText(s.cpSendDate);
            txtSendUser.SetText(s.cpSendUser);
            txtPASIAppDate.SetText(s.cpPASIAppDate);
            txtPASIAppUser.SetText(s.cpPASIAppUser);
            txtSuppAppDate.SetText(s.cpSuppAppDate);
            txtSuppAppUser.SetText(s.cpSuppAppUser);
            txtAffFinalAppDate.SetText(s.cpFinalAppDate);
            txtAffFinalAppUser.SetText(s.cpFinalAppUser);

            if (txtPASIAppDate.GetText() != '' && txtPASIAppDate.GetText() != '-') {
                btnApprove.SetEnabled(false);
            }
            if (txtAffFinalAppDate.GetText() != '' && txtAffFinalAppDate.GetText() != '-') {
                btnApprove.SetEnabled(false);
            }

            var pMsg = s.cpMessage;
            if (pMsg != '') {
                if (pMsg.substring(1,5) == '1008' || pMsg.substring(1,5) == '1009') {
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
    </asp:Content>

