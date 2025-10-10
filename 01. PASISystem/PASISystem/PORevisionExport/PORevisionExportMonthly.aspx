<%@ Page Title="" Language="vb" AutoEventWireup="false" MasterPageFile="~/SiteContent.Master"
    CodeBehind="PORevisionExportMonthly.aspx.vb" Inherits="PASISystem.PORevisionExportMonthly" %>

<%@ Register Assembly="DevExpress.Web.v14.1, Version=14.1.6.0, Culture=neutral, PublicKeyToken=b88d1754d700e49a"
    Namespace="DevExpress.Web.ASPxGridView.Export" TagPrefix="dx" %>
<%@ Register Assembly="DevExpress.Web.v14.1, Version=14.1.6.0, Culture=neutral, PublicKeyToken=b88d1754d700e49a"
    Namespace="DevExpress.Web.ASPxGridView" TagPrefix="dx" %>
<%@ Register Assembly="DevExpress.Web.v14.1, Version=14.1.6.0, Culture=neutral, PublicKeyToken=b88d1754d700e49a"
    Namespace="DevExpress.Web.ASPxEditors" TagPrefix="dx" %>
<%@ Register Assembly="DevExpress.Web.v14.1, Version=14.1.6.0, Culture=neutral, PublicKeyToken=b88d1754d700e49a"
    Namespace="DevExpress.Web.ASPxGlobalEvents" TagPrefix="dx" %>
<%@ Register Assembly="DevExpress.Web.v14.1, Version=14.1.6.0, Culture=neutral, PublicKeyToken=b88d1754d700e49a"
    Namespace="DevExpress.Web.ASPxRoundPanel" TagPrefix="dx" %>
<%@ Register Assembly="DevExpress.Web.v14.1, Version=14.1.6.0, Culture=neutral, PublicKeyToken=b88d1754d700e49a"
    Namespace="DevExpress.Web.ASPxPanel" TagPrefix="dx" %>
<%@ Register Assembly="DevExpress.Web.v14.1, Version=14.1.6.0, Culture=neutral, PublicKeyToken=b88d1754d700e49a"
    Namespace="DevExpress.Web.ASPxCallback" TagPrefix="dx" %>
<asp:Content ID="Content1" ContentPlaceHolderID="head" runat="server">
    <style type="text/css">
        .dxeHLC, .dxeHC, .dxeHFC
        {
            display: none;
        }
        .style1
        {
            height: 16px;
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
            height = height - (height * 70 / 100)
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
                var period = dtPeriod.GetValue();
                var nBulan = period.getMonth();
                var tahun = period.getFullYear();
                var kabisat;
                var hariIsi;

                //Januari, Maret, May, July, Aug, Oct, Dec --> 31
                if (nBulan == "0" || nBulan == "2" || nBulan == "4" || nBulan == "6" || nBulan == "7" || nBulan == "9" || nBulan == "11") {
                    hariIsi = 31;
                }
                //April, Jun, Sep, Nov --> 30
                if (nBulan == "3" || nBulan == "5" || nBulan == "8" || nBulan == "10") {
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
            }, 10);
        } 
    </script>
</asp:Content>
<asp:Content ID="Content2" ContentPlaceHolderID="Content" runat="server">
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
                                        <dx:ASPxLabel ID="ASPxLabel3" runat="server" Text="PERIOD" Font-Names="Tahoma" Font-Size="8pt"
                                            Width="100%">
                                        </dx:ASPxLabel>
                                    </td>
                                    <td align="left" valign="middle" width="150px">
                                        <dx:ASPxTimeEdit ID="dtPeriodFrom" runat="server" ClientInstanceName="dtPeriodFrom"
                                            DisplayFormatString="dd MMM yyyy" EditFormat="Custom" EditFormatString="dd MMM yyyy"
                                            Width="150px" Font-Names="Tahoma" Font-Size="8pt">
                                        </dx:ASPxTimeEdit>
                                    </td>
                                    <td align="left" valign="middle" width="10px">
                                        <dx:ASPxLabel ID="ASPxLabel7" runat="server" Text="COMMERCIAL" Font-Names="Tahoma"
                                            Font-Size="8pt" Width="100px">
                                        </dx:ASPxLabel>
                                    </td>
                                    <td align="left" valign="middle" width="150px">
                                        <dx:ASPxRadioButtonList ID="rblCommercial" runat="server" RepeatDirection="Horizontal"
                                            Width="150px" ClientInstanceName="rblCommercial" SelectedIndex="0" TabIndex="9"
                                            Font-Names="Tahoma" Font-Size="8pt">
                                            <RadioButtonStyle HorizontalAlign="Left">
                                            </RadioButtonStyle>
                                            <Items>
                                                <dx:ListEditItem Text="ALL" Value="2" Selected="True" />
                                                <dx:ListEditItem Text="YES" Value="1" />
                                                <dx:ListEditItem Text="NO" Value="0" />
                                            </Items>
                                            <Border BorderStyle="None"></Border>
                                        </dx:ASPxRadioButtonList>
                                    </td>
                                    <td align="left" valign="middle" width="50px">
                                        <dx:ASPxLabel ID="ASPxLabel28" runat="server" Text="DELIVERY LOCATION" Font-Names="Tahoma"
                                            Font-Size="8pt" Width="150px">
                                        </dx:ASPxLabel>
                                    </td>
                                    <td align="left" valign="middle" width="165px">
                                        <dx:ASPxComboBox ID="cboDeliveryLoc" runat="server" TextFormatString="{0}" DropDownStyle="DropDown"
                                            Width="150px" ClientInstanceName="cboDeliveryLoc" IncrementalFilteringMode="StartsWith"
                                            Font-Names="Tahoma" Font-Size="8pt" MaxLength="10" TabIndex="4">
                                            <ClientSideEvents SelectedIndexChanged="function(s, e) {SelectedIndexChangedAff();}"
                                                LostFocus="function(s, e) { lblInfo.SetText(&quot;&quot;); }" />
                                        </dx:ASPxComboBox>
                                    </td>
                                    <td align="left" valign="middle" width="165px">
                                        <dx:ASPxTextBox ID="txtDeliveryLoc" runat="server" ClientInstanceName="txtDeliveryLoc"
                                            Width="200px" Height="20px" ReadOnly="True" TabIndex="5" Font-Names="Tahoma"
                                            Font-Size="8pt">
                                            <ReadOnlyStyle BackColor="#CCCCCC">
                                            </ReadOnlyStyle>
                                        </dx:ASPxTextBox>
                                    </td>
                                </tr>
                                <tr>
                                    <td style="width: 5px;">
                                        &nbsp;
                                    </td>
                                    <td align="left" valign="middle" width="120px">
                                        <dx:ASPxLabel ID="ASPxLabel2" runat="server" Text="PO MONTHLY / EMERGENCY" Font-Names="Tahoma"
                                            Font-Size="8pt" Width="150px">
                                        </dx:ASPxLabel>
                                    </td>
                                    <td align="left" valign="middle" width="150px">
                                        <dx:ASPxTextBox ID="txtPOEmergency" runat="server" ClientInstanceName="txtPOEmergency"
                                            Width="150px" Height="20px" ReadOnly="True" TabIndex="5" Font-Names="Tahoma"
                                            Font-Size="8pt">
                                            <ReadOnlyStyle BackColor="#CCCCCC">
                                            </ReadOnlyStyle>
                                        </dx:ASPxTextBox>
                                    </td>
                                    <td align="left" valign="middle" width="10px" style="width: -720">
                                        <dx:ASPxLabel ID="ASPxLabel27" runat="server" Text="SHIP BY" Font-Names="Tahoma"
                                            Font-Size="8pt" Width="100px">
                                        </dx:ASPxLabel>
                                    </td>
                                    <td align="left" valign="middle" width="360px" style="width: -180">
                                        <dx:ASPxTextBox ID="txtShipBy" runat="server" ClientInstanceName="txtShipBy" Width="150px"
                                            Height="20px" ReadOnly="True" TabIndex="5" Font-Names="Tahoma" Font-Size="8pt">
                                            <ReadOnlyStyle BackColor="#CCCCCC">
                                            </ReadOnlyStyle>
                                        </dx:ASPxTextBox>
                                    </td>
                                    <td align="left" valign="middle" width="360px" style="width: 0">
                                        &nbsp;
                                    </td>
                                    <td align="left" valign="middle" width="360px" style="width: 90px">
                                        &nbsp;
                                    </td>
                                    <td align="left" valign="middle" width="360px" style="width: 90px">
                                        &nbsp;
                                    </td>
                                </tr>
                                <tr>
                                    <td style="width: 5px;">
                                    </td>
                                    <td align="left" valign="middle" width="120px">
                                        <dx:ASPxLabel ID="ASPxLabel1" runat="server" Text="AFFILIATE CODE/NAME" Font-Names="Tahoma"
                                            Font-Size="8pt" Width="150px">
                                        </dx:ASPxLabel>
                                    </td>
                                    <td align="left" valign="middle" width="150px">
                                        <dx:ASPxComboBox ID="cboAffiliateCode" runat="server" TextFormatString="{0}" DropDownStyle="DropDown"
                                            Width="150px" ClientInstanceName="cboAffiliateCode" IncrementalFilteringMode="StartsWith"
                                            Font-Names="Tahoma" Font-Size="8pt" MaxLength="10" TabIndex="4">
                                            <ClientSideEvents SelectedIndexChanged="function(s, e) {SelectedIndexChangedAff();}"
                                                LostFocus="function(s, e) { lblInfo.SetText(&quot;&quot;); }" />
                                        </dx:ASPxComboBox>
                                    </td>
                                    <td align="left" valign="middle" colspan="4" width="360px">
                                        <dx:ASPxTextBox ID="txtAffiliateName" runat="server" ClientInstanceName="txtAffiliateName"
                                            Width="250px" Height="20px" ReadOnly="True" TabIndex="5" Font-Names="Tahoma"
                                            Font-Size="8pt">
                                            <ReadOnlyStyle BackColor="#CCCCCC">
                                            </ReadOnlyStyle>
                                        </dx:ASPxTextBox>
                                    </td>
                                    <td align="left" valign="middle" width="360px">
                                        &nbsp;
                                    </td>
                                </tr>
                                <tr>
                                    <td style="width: 5px;">
                                    </td>
                                    <td align="left" valign="middle" width="120px">
                                        <dx:ASPxLabel ID="ASPxLabel26" runat="server" Text="SUPPLIER CODE/NAME" Font-Names="Tahoma"
                                            Font-Size="8pt" Width="150px">
                                        </dx:ASPxLabel>
                                    </td>
                                    <td align="left" valign="middle" width="150px">
                                        <dx:ASPxComboBox ID="cboSupplierCode" runat="server" TextFormatString="{0}" DropDownStyle="DropDown"
                                            Width="150px" ClientInstanceName="cboSupplierCode" IncrementalFilteringMode="StartsWith"
                                            Font-Names="Tahoma" Font-Size="8pt" MaxLength="10" TabIndex="4">
                                            <ClientSideEvents SelectedIndexChanged="function(s, e) {SelectedIndexChangedAff();}"
                                                LostFocus="function(s, e) { lblInfo.SetText(&quot;&quot;); }" />
                                        </dx:ASPxComboBox>
                                    </td>
                                    <td align="left" valign="middle" colspan="4" width="360px">
                                        <dx:ASPxTextBox ID="txtSupplierName" runat="server" ClientInstanceName="txtSupplierName"
                                            Width="250px" Height="20px" ReadOnly="True" TabIndex="5" Font-Names="Tahoma"
                                            Font-Size="8pt">
                                            <ReadOnlyStyle BackColor="#CCCCCC">
                                            </ReadOnlyStyle>
                                        </dx:ASPxTextBox>
                                    </td>
                                    <td align="left" valign="middle" width="360px">
                                        <table style="width: 100%;">
                                            <tr>
                                                <td align="right" valign="middle">
                                                    <dx:ASPxButton ID="btnRefresh" runat="server" Text="SEARCH" Font-Names="Tahoma" Width="85px"
                                                        AutoPostBack="False" Font-Size="8pt" ClientInstanceName="btnRefresh" TabIndex="10"
                                                        Visible="False">
                                                        <ClientSideEvents Click="function(s, e) {
                                                            var pDateFrom = dtPeriodFrom.GetText();
                                                            var pDateTo = dtPeriodTo.GetText();
                                                            var pAffCode= cboAffiliateCode.GetText();
                                                            var pSendTo = rblMonthlyEme.GetValue();
                                                            var pMonthly = rblSendTo.GetValue();
                                                            var pComm = rblCommercial.GetValue();
                                                            
                                                            grid.PerformCallback('load' + '|' + pDateFrom + '|' + pDateTo + '|' + pAffCode + '|' + pSendTo + '|' + pMonthly + '|' + pComm);
                                                            lblInfo.SetText('');
                                                        }" />
                                                    </dx:ASPxButton>
                                                </td>
                                                <td align="right" valign="middle">
                                                    <dx:ASPxButton ID="btnClear" runat="server" Text="CLEAR" Font-Names="Tahoma" Width="85px"
                                                        AutoPostBack="False" Font-Size="8pt" ClientInstanceName="btnClear" TabIndex="11">
                                                        <ClientSideEvents Click="function(s, e) {
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
                                <tr>
                                    <td style="width: 5px;">
                                    </td>
                                    <td align="left" valign="middle" width="80px">
                                        <dx:ASPxLabel ID="ASPxLabel6" runat="server" Text="REVISION NO." Font-Names="Tahoma"
                                            Font-Size="8pt" Width="150px">
                                        </dx:ASPxLabel>
                                    </td>
                                    <td align="left" valign="middle" width="150px">
                                        <dx:ASPxTextBox ID="txtRevisionNo" runat="server" ClientInstanceName="txtAffiliateName"
                                            Width="150px" Height="20px" TabIndex="5" Font-Names="Tahoma" Font-Size="8pt">
                                            <ReadOnlyStyle BackColor="#CCCCCC">
                                            </ReadOnlyStyle>
                                        </dx:ASPxTextBox>
                                    </td>
                                    <td align="left" valign="middle" width="360px" colspan="4">
                                        &nbsp;
                                    </td>
                                    <td align="left" valign="middle" width="360px">
                                        &nbsp;
                                    </td>
                                </tr>
                                <tr>
                                    <td style="width: 5px;">
                                        &nbsp;
                                    </td>
                                    <td align="left" valign="middle" width="80px">
                                        &nbsp;
                                    </td>
                                    <td align="left" valign="middle" width="150px">
                                        &nbsp;
                                    </td>
                                    <td align="left" valign="middle" width="80px" colspan="4">
                                        &nbsp;
                                    </td>
                                    <td align="left" valign="middle" width="80px">
                                        &nbsp;
                                    </td>
                                </tr>
                                <tr>
                                    <td style="width: 5px;">
                                        &nbsp;
                                    </td>
                                    <td align="left" valign="middle" width="80px">
                                        &nbsp;
                                    </td>
                                    <td align="left" valign="middle" width="150px">
                                        &nbsp;
                                    </td>
                                    <td align="left" valign="middle" width="80px" colspan="2">
                                        &nbsp;
                                    </td>
                                    <td align="left" valign="middle" width="120px" colspan="2">
                                        &nbsp;
                                    </td>
                                    <td align="left" valign="middle" width="120px">
                                        &nbsp;
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
                            <dx:ASPxLabel ID="lblInfo" runat="server" Text="[lblinfo]" Font-Names="Tahoma" ClientInstanceName="lblInfo"
                                Font-Bold="True" Font-Italic="True" Font-Size="8pt">
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
            <td valign="top" colspan="6">
                &nbsp;
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
                            <asp:TextBox ID="lSuuplier" runat="server" BackColor="Yellow" BorderStyle="None"
                                ReadOnly="True" Width="30px" Visible="False"></asp:TextBox>
                        </td>
                        <td align="right" valign="middle" width="110PX">
                            <dx:ASPxLabel ID="ASPxLabel21" runat="server" Font-Names="Tahoma" Font-Size="8pt"
                                Text=": EDITED BY PASI" Width="110px" Visible="False">
                            </dx:ASPxLabel>
                        </td>
                        <td width="3px">
                            &nbsp;
                        </td>
                        <td align="right" valign="middle" width="30px">
                            <asp:TextBox ID="lSuuplier0" runat="server" BackColor="GreenYellow" BorderStyle="None"
                                ReadOnly="True" Width="30px"></asp:TextBox>
                        </td>
                        <td align="right" valign="middle" width="140px">
                            <dx:ASPxLabel ID="ASPxLabel24" runat="server" Font-Names="Tahoma" Font-Size="8pt"
                                Text=": REVISION" Width="140px">
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
            <td colspan="8" align="center" valign="top">
                <table style="width: 100%;">
                    <tr>
                        <td width="150px">
                            &nbsp;
                        </td>
                        <td width="150px">
                            &nbsp;
                        </td>
                        <td rowspan="5">
                            <table style="width: 100%;" border="0" cellpadding="0" cellspacing="0" width="100%">
                                <tr>
                                    <td height="16px" rowspan="0">
                                    </td>
                                    <td width="150px" height="16px" style="border-style: solid none none solid; border-width: 0.1px;
                                        border-color: #808080;">
                                        <dx:ASPxTextBox runat="server" Width="150px" MaxLength="10" Height="16px" Font-Names="Tahoma"
                                            Font-Size="8pt" ID="ASPxLabel25" BackColor="#FFD2A6" Text="WEEK 1" ReadOnly="True"
                                            HorizontalAlign="Center" Font-Bold="True">
                                            <Border BorderStyle="Groove" BorderColor="Silver" BorderWidth="1px" />
                                            <BorderLeft BorderColor="#CCCCCC" BorderStyle="Groove" BorderWidth="1px" />
                                            <BorderTop BorderColor="#CCCCCC" BorderStyle="Groove" BorderWidth="1px" />
                                            <BorderRight BorderColor="#CCCCCC" BorderStyle="Groove" BorderWidth="1px" />
                                            <BorderBottom BorderColor="#CCCCCC" BorderStyle="Groove" BorderWidth="1px" />
                                        </dx:ASPxTextBox>
                                    </td>
                                    <td width="150px" height="16px" style="border-style: solid none none solid; border-width: 0.1px;
                                        border-color: #808080;">
                                        <dx:ASPxTextBox runat="server" Width="150px" MaxLength="10" Height="16px" Font-Names="Tahoma"
                                            Font-Size="8pt" ID="ASPxTextBox1" BackColor="#FFD2A6" Text="WEEK 2" ReadOnly="True"
                                            HorizontalAlign="Center" Font-Bold="True">
                                            <Border BorderStyle="Groove" BorderColor="Silver" BorderWidth="1px" />
                                            <BorderLeft BorderColor="#CCCCCC" BorderStyle="Groove" BorderWidth="1px" />
                                            <BorderTop BorderColor="#CCCCCC" BorderStyle="Groove" BorderWidth="1px" />
                                            <BorderRight BorderColor="#CCCCCC" BorderStyle="Groove" BorderWidth="1px" />
                                            <BorderBottom BorderColor="#CCCCCC" BorderStyle="Groove" BorderWidth="1px" />
                                        </dx:ASPxTextBox>
                                    </td>
                                    <td width="150px" height="16px" style="border-style: solid none none solid; border-width: 0.1px;
                                        border-color: #808080;">
                                        <dx:ASPxTextBox runat="server" Width="150px" MaxLength="10" Height="16px" Font-Names="Tahoma"
                                            Font-Size="8pt" ID="ASPxTextBox2" BackColor="#FFD2A6" Text="WEEK 3" ReadOnly="True"
                                            HorizontalAlign="Center" Font-Bold="True">
                                            <Border BorderStyle="Groove" BorderColor="Silver" BorderWidth="1px" />
                                            <BorderLeft BorderColor="#CCCCCC" BorderStyle="Groove" BorderWidth="1px" />
                                            <BorderTop BorderColor="#CCCCCC" BorderStyle="Groove" BorderWidth="1px" />
                                            <BorderRight BorderColor="#CCCCCC" BorderStyle="Groove" BorderWidth="1px" />
                                            <BorderBottom BorderColor="#CCCCCC" BorderStyle="Groove" BorderWidth="1px" />
                                        </dx:ASPxTextBox>
                                    </td>
                                    <td width="150px" height="16px" style="border-style: solid none none solid; border-width: 0.1px;
                                        border-color: #808080;">
                                        <dx:ASPxTextBox runat="server" Width="150px" MaxLength="10" Height="16px" Font-Names="Tahoma"
                                            Font-Size="8pt" ID="ASPxTextBox3" BackColor="#FFD2A6" Text="WEEK 4" ReadOnly="True"
                                            HorizontalAlign="Center" Font-Bold="True">
                                            <Border BorderStyle="Groove" BorderColor="Silver" BorderWidth="1px" />
                                            <BorderLeft BorderColor="#CCCCCC" BorderStyle="Groove" BorderWidth="1px" />
                                            <BorderTop BorderColor="#CCCCCC" BorderStyle="Groove" BorderWidth="1px" />
                                            <BorderRight BorderColor="#CCCCCC" BorderStyle="Groove" BorderWidth="1px" />
                                            <BorderBottom BorderColor="#CCCCCC" BorderStyle="Groove" BorderWidth="1px" />
                                        </dx:ASPxTextBox>
                                    </td>
                                    <td width="150px" height="16px" style="border-style: solid none none solid; border-width: 0.1px;
                                        border-color: #808080;">
                                        <dx:ASPxTextBox runat="server" Width="150px" MaxLength="10" Height="16px" Font-Names="Tahoma"
                                            Font-Size="8pt" ID="ASPxTextBox4" BackColor="#FFD2A6" Text="WEEK 5" ReadOnly="True"
                                            HorizontalAlign="Center" Font-Bold="True">
                                            <Border BorderStyle="Groove" BorderColor="Silver" BorderWidth="1px" />
                                            <BorderLeft BorderColor="#CCCCCC" BorderStyle="Groove" BorderWidth="1px" />
                                            <BorderTop BorderColor="#CCCCCC" BorderStyle="Groove" BorderWidth="1px" />
                                            <BorderRight BorderColor="#CCCCCC" BorderStyle="Groove" BorderWidth="1px" />
                                            <BorderBottom BorderColor="#CCCCCC" BorderStyle="Groove" BorderWidth="1px" />
                                        </dx:ASPxTextBox>
                                    </td>
                                </tr>
                                <tr>
                                    <td width="150px" align="left" height="16px" style="border-style: solid none none solid;
                                        border-width: 0.1px; border-color: #808080;">
                                        <dx:ASPxTextBox runat="server" Width="175px" MaxLength="10" Height="16px" Font-Names="Tahoma"
                                            Font-Size="8pt" ID="ASPxTextBox5" BackColor="#FFD2A6" Text="ORDER NO" ReadOnly="True"
                                            Font-Bold="True">
                                            <Border BorderStyle="Groove" BorderColor="Silver" BorderWidth="1px" />
                                            <BorderLeft BorderColor="#CCCCCC" BorderStyle="Groove" BorderWidth="1px" />
                                            <BorderTop BorderColor="#CCCCCC" BorderStyle="Groove" BorderWidth="1px" />
                                            <BorderRight BorderColor="#CCCCCC" BorderStyle="Groove" BorderWidth="1px" />
                                            <BorderBottom BorderColor="#CCCCCC" BorderStyle="Groove" BorderWidth="1px" />
                                        </dx:ASPxTextBox>
                                    </td>
                                    <td width="150px" height="16px" style="border-style: solid none none solid; border-width: 0.1px;
                                        border-color: #808080;">
                                        <dx:ASPxTextBox runat="server" Width="150px" MaxLength="10" Height="16px" ClientInstanceName="txtOrderNoWeek1"
                                            Font-Names="Tahoma" Font-Size="8pt" ID="txtOrderNoWeek1" BackColor="#FFFFE1"
                                            ReadOnly="True">
                                            <ReadOnlyStyle BackColor="#CCCCCC">
                                            </ReadOnlyStyle>
                                            <Border BorderStyle="Groove" BorderColor="Silver" BorderWidth="1px" />
                                            <BorderLeft BorderColor="#CCCCCC" BorderStyle="Groove" BorderWidth="1px" />
                                            <BorderTop BorderColor="#CCCCCC" BorderStyle="Groove" BorderWidth="1px" />
                                            <BorderRight BorderColor="#CCCCCC" BorderStyle="Groove" BorderWidth="1px" />
                                            <BorderBottom BorderColor="#CCCCCC" BorderStyle="Groove" BorderWidth="1px" />
                                        </dx:ASPxTextBox>
                                    </td>
                                    <td width="150px" height="16px" style="border-style: solid none none solid; border-width: 0.1px;
                                        border-color: #808080;">
                                        <dx:ASPxTextBox runat="server" Width="150px" MaxLength="10" Height="16px" ClientInstanceName="txtOrderNoWeek2"
                                            Font-Names="Tahoma" Font-Size="8pt" ID="txtOrderNoWeek2" BackColor="#FFFFE1"
                                            ReadOnly="True">
                                            <ReadOnlyStyle BackColor="#CCCCCC">
                                            </ReadOnlyStyle>
                                            <Border BorderStyle="Groove" BorderColor="Silver" BorderWidth="1px" />
                                            <BorderLeft BorderColor="#CCCCCC" BorderStyle="Groove" BorderWidth="1px" />
                                            <BorderTop BorderColor="#CCCCCC" BorderStyle="Groove" BorderWidth="1px" />
                                            <BorderRight BorderColor="#CCCCCC" BorderStyle="Groove" BorderWidth="1px" />
                                            <BorderBottom BorderColor="#CCCCCC" BorderStyle="Groove" BorderWidth="1px" />
                                        </dx:ASPxTextBox>
                                    </td>
                                    <td width="150px" height="16px" style="border-style: solid none none solid; border-width: 0.1px;
                                        border-color: #808080;">
                                        <dx:ASPxTextBox runat="server" Width="150px" MaxLength="10" Height="16px" ClientInstanceName="txtOrderNoWeek3"
                                            Font-Names="Tahoma" Font-Size="8pt" ID="txtOrderNoWeek3" BackColor="#FFFFE1"
                                            ReadOnly="True">
                                            <ReadOnlyStyle BackColor="#CCCCCC">
                                            </ReadOnlyStyle>
                                            <Border BorderStyle="Groove" BorderColor="Silver" BorderWidth="1px" />
                                            <BorderLeft BorderColor="#CCCCCC" BorderStyle="Groove" BorderWidth="1px" />
                                            <BorderTop BorderColor="#CCCCCC" BorderStyle="Groove" BorderWidth="1px" />
                                            <BorderRight BorderColor="#CCCCCC" BorderStyle="Groove" BorderWidth="1px" />
                                            <BorderBottom BorderColor="#CCCCCC" BorderStyle="Groove" BorderWidth="1px" />
                                        </dx:ASPxTextBox>
                                    </td>
                                    <td width="150px" height="16px" style="border-style: solid none none solid; border-width: 0.1px;
                                        border-color: #808080;">
                                        <dx:ASPxTextBox runat="server" Width="150px" MaxLength="10" Height="16px" ClientInstanceName="txtOrderNoWeek4"
                                            Font-Names="Tahoma" Font-Size="8pt" ID="txtOrderNoWeek4" BackColor="#FFFFE1"
                                            ReadOnly="True">
                                            <ReadOnlyStyle BackColor="#CCCCCC">
                                            </ReadOnlyStyle>
                                            <Border BorderStyle="Groove" BorderColor="Silver" BorderWidth="1px" />
                                            <BorderLeft BorderColor="#CCCCCC" BorderStyle="Groove" BorderWidth="1px" />
                                            <BorderTop BorderColor="#CCCCCC" BorderStyle="Groove" BorderWidth="1px" />
                                            <BorderRight BorderColor="#CCCCCC" BorderStyle="Groove" BorderWidth="1px" />
                                            <BorderBottom BorderColor="#CCCCCC" BorderStyle="Groove" BorderWidth="1px" />
                                        </dx:ASPxTextBox>
                                    </td>
                                    <td width="150px" height="16px" style="border-style: solid none none solid; border-width: 0.1px;
                                        border-color: #808080;">
                                        <dx:ASPxTextBox runat="server" Width="150px" MaxLength="10" Height="16px" ClientInstanceName="txtOrderNoWeek5"
                                            Font-Names="Tahoma" Font-Size="8pt" ID="txtOrderNoWeek5" BackColor="#FFFFE1"
                                            ReadOnly="True">
                                            <ReadOnlyStyle BackColor="#CCCCCC">
                                            </ReadOnlyStyle>
                                            <Border BorderStyle="Groove" BorderColor="Silver" BorderWidth="1px" />
                                            <BorderLeft BorderColor="#CCCCCC" BorderStyle="Groove" BorderWidth="1px" />
                                            <BorderTop BorderColor="#CCCCCC" BorderStyle="Groove" BorderWidth="1px" />
                                            <BorderRight BorderColor="#CCCCCC" BorderStyle="Groove" BorderWidth="1px" />
                                            <BorderBottom BorderColor="#CCCCCC" BorderStyle="Groove" BorderWidth="1px" />
                                        </dx:ASPxTextBox>
                                    </td>
                                </tr>
                                <tr>
                                    <td width="150px" align="left" height="16px" style="border-style: solid none none solid;
                                        border-width: 0.1px; border-color: #808080;">
                                        <dx:ASPxTextBox runat="server" Width="175px" MaxLength="10" Height="21px" Font-Names="Tahoma"
                                            Font-Size="8pt" ID="ASPxTextBox6" BackColor="#FFD2A6" Text="ETD VENDOR" ReadOnly="True"
                                            Font-Bold="True">
                                            <Border BorderStyle="Groove" BorderColor="Silver" BorderWidth="1px" />
                                            <BorderLeft BorderColor="#CCCCCC" BorderStyle="Groove" BorderWidth="1px" />
                                            <BorderTop BorderColor="#CCCCCC" BorderStyle="Groove" BorderWidth="1px" />
                                            <BorderRight BorderColor="#CCCCCC" BorderStyle="Groove" BorderWidth="1px" />
                                            <BorderBottom BorderColor="#CCCCCC" BorderStyle="Groove" BorderWidth="1px" />
                                        </dx:ASPxTextBox>
                                    </td>
                                    <td width="150px" height="16px" style="border-style: solid none none solid; border-width: 0.1px;
                                        border-color: #808080;">
                                        <dx:ASPxTimeEdit ID="dtWeekETDVendorOld1" runat="server" ClientInstanceName="dtWeekETDVendorOld1"
                                            DisplayFormatString="dd MMM yyyy" EditFormat="Custom" EditFormatString="dd MMM yyyy"
                                            Font-Names="Tahoma" Font-Size="8pt" BackColor="#FFFFE1" Width="150px" ReadOnly="True">
                                            <ReadOnlyStyle BackColor="#CCCCCC">
                                            </ReadOnlyStyle>
                                            <Border BorderStyle="Groove" BorderColor="Silver" BorderWidth="1px" />
                                            <BorderLeft BorderColor="#CCCCCC" BorderStyle="Groove" BorderWidth="1px" />
                                            <BorderTop BorderColor="#CCCCCC" BorderStyle="Groove" BorderWidth="1px" />
                                            <BorderRight BorderColor="#CCCCCC" BorderStyle="Groove" BorderWidth="1px" />
                                            <BorderBottom BorderColor="#CCCCCC" BorderStyle="Groove" BorderWidth="1px" />
                                        </dx:ASPxTimeEdit>
                                    </td>
                                    <td width="150px" height="16px" style="border-style: solid none none solid; border-width: 0.1px;
                                        border-color: #808080;">
                                        <dx:ASPxTimeEdit ID="dtWeekETDVendorOld2" runat="server" ClientInstanceName="dtWeekETDVendorOld2"
                                            DisplayFormatString="dd MMM yyyy" EditFormat="Custom" EditFormatString="dd MMM yyyy"
                                            Font-Names="Tahoma" Font-Size="8pt" BackColor="#FFFFE1" Width="150px" ReadOnly="True">
                                            <ReadOnlyStyle BackColor="#CCCCCC">
                                            </ReadOnlyStyle>
                                            <Border BorderStyle="Groove" BorderColor="Silver" BorderWidth="1px" />
                                            <BorderLeft BorderColor="#CCCCCC" BorderStyle="Groove" BorderWidth="1px" />
                                            <BorderTop BorderColor="#CCCCCC" BorderStyle="Groove" BorderWidth="1px" />
                                            <BorderRight BorderColor="#CCCCCC" BorderStyle="Groove" BorderWidth="1px" />
                                            <BorderBottom BorderColor="#CCCCCC" BorderStyle="Groove" BorderWidth="1px" />
                                        </dx:ASPxTimeEdit>
                                    </td>
                                    <td width="150px" height="16px" style="border-style: solid none none solid; border-width: 0.1px;
                                        border-color: #808080;">
                                        <dx:ASPxTimeEdit ID="dtWeekETDVendorOld3" runat="server" ClientInstanceName="dtWeekETDVendorOld3"
                                            DisplayFormatString="dd MMM yyyy" EditFormat="Custom" EditFormatString="dd MMM yyyy"
                                            Font-Names="Tahoma" Font-Size="8pt" BackColor="#FFFFE1" Width="150px" ReadOnly="True">
                                            <ReadOnlyStyle BackColor="#CCCCCC">
                                            </ReadOnlyStyle>
                                            <Border BorderStyle="Groove" BorderColor="Silver" BorderWidth="1px" />
                                            <BorderLeft BorderColor="#CCCCCC" BorderStyle="Groove" BorderWidth="1px" />
                                            <BorderTop BorderColor="#CCCCCC" BorderStyle="Groove" BorderWidth="1px" />
                                            <BorderRight BorderColor="#CCCCCC" BorderStyle="Groove" BorderWidth="1px" />
                                            <BorderBottom BorderColor="#CCCCCC" BorderStyle="Groove" BorderWidth="1px" />
                                        </dx:ASPxTimeEdit>
                                    </td>
                                    <td width="150px" height="16px" style="border-style: solid none none solid; border-width: 0.1px;
                                        border-color: #808080;">
                                        <dx:ASPxTimeEdit ID="dtWeekETDVendorOld4" runat="server" ClientInstanceName="dtWeekETDVendorOld4"
                                            DisplayFormatString="dd MMM yyyy" EditFormat="Custom" EditFormatString="dd MMM yyyy"
                                            Font-Names="Tahoma" Font-Size="8pt" BackColor="#FFFFE1" Width="150px" ReadOnly="True">
                                            <ReadOnlyStyle BackColor="#CCCCCC">
                                            </ReadOnlyStyle>
                                            <Border BorderStyle="Groove" BorderColor="Silver" BorderWidth="1px" />
                                            <BorderLeft BorderColor="#CCCCCC" BorderStyle="Groove" BorderWidth="1px" />
                                            <BorderTop BorderColor="#CCCCCC" BorderStyle="Groove" BorderWidth="1px" />
                                            <BorderRight BorderColor="#CCCCCC" BorderStyle="Groove" BorderWidth="1px" />
                                            <BorderBottom BorderColor="#CCCCCC" BorderStyle="Groove" BorderWidth="1px" />
                                        </dx:ASPxTimeEdit>
                                    </td>
                                    <td width="150px" height="16px" style="border-style: solid none none solid; border-width: 0.1px;
                                        border-color: #808080;">
                                        <dx:ASPxTimeEdit ID="dtWeekETDVendorOld5" runat="server" ClientInstanceName="dtWeekETDVendorOld5"
                                            DisplayFormatString="dd MMM yyyy" EditFormat="Custom" EditFormatString="dd MMM yyyy"
                                            Font-Names="Tahoma" Font-Size="8pt" BackColor="#FFFFE1" Width="150px" ReadOnly="True">
                                            <ReadOnlyStyle BackColor="#CCCCCC">
                                            </ReadOnlyStyle>
                                            <Border BorderStyle="Groove" BorderColor="Silver" BorderWidth="1px" />
                                            <BorderLeft BorderColor="#CCCCCC" BorderStyle="Groove" BorderWidth="1px" />
                                            <BorderTop BorderColor="#CCCCCC" BorderStyle="Groove" BorderWidth="1px" />
                                            <BorderRight BorderColor="#CCCCCC" BorderStyle="Groove" BorderWidth="1px" />
                                            <BorderBottom BorderColor="#CCCCCC" BorderStyle="Groove" BorderWidth="1px" />
                                        </dx:ASPxTimeEdit>
                                    </td>
                                </tr>
                                <tr>
                                    <td width="150px" align="left" height="16px" style="border-style: solid none none solid;
                                        border-width: 0.1px; border-color: #808080;">
                                        <dx:ASPxTextBox runat="server" Width="175px" MaxLength="10" Height="21px" Font-Names="Tahoma"
                                            Font-Size="8pt" ID="ASPxTextBox10" BackColor="GreenYellow" Text="ETD VENDOR* (REVISION)"
                                            ReadOnly="True" ForeColor="Red" Font-Bold="True">
                                            <Border BorderStyle="Groove" BorderColor="Silver" BorderWidth="1px" />
                                            <BorderLeft BorderColor="#CCCCCC" BorderStyle="Groove" BorderWidth="1px" />
                                            <BorderTop BorderColor="#CCCCCC" BorderStyle="Groove" BorderWidth="1px" />
                                            <BorderRight BorderColor="#CCCCCC" BorderStyle="Groove" BorderWidth="1px" />
                                            <BorderBottom BorderColor="#CCCCCC" BorderStyle="Groove" BorderWidth="1px" />
                                        </dx:ASPxTextBox>
                                    </td>
                                    <td width="150px" height="16px" style="border-style: solid none none solid; border-width: 0.1px;
                                        border-color: #808080;">
                                        <dx:ASPxTimeEdit ID="dtWeekETDVendor1" runat="server" ClientInstanceName="dtWeekETDVendor1"
                                            DisplayFormatString="dd MMM yyyy" EditFormat="Custom" EditFormatString="dd MMM yyyy"
                                            Font-Names="Tahoma" Font-Size="8pt" BackColor="GreenYellow" Width="150px">
                                            <Border BorderStyle="Groove" BorderColor="Silver" BorderWidth="1px" />
                                            <BorderLeft BorderColor="#CCCCCC" BorderStyle="Groove" BorderWidth="1px" />
                                            <BorderTop BorderColor="#CCCCCC" BorderStyle="Groove" BorderWidth="1px" />
                                            <BorderRight BorderColor="#CCCCCC" BorderStyle="Groove" BorderWidth="1px" />
                                            <BorderBottom BorderColor="#CCCCCC" BorderStyle="Groove" BorderWidth="1px" />
                                        </dx:ASPxTimeEdit>
                                    </td>
                                    <td width="150px" height="16px" style="border-style: solid none none solid; border-width: 0.1px;
                                        border-color: #808080;">
                                        <dx:ASPxTimeEdit ID="dtWeekETDVendor2" runat="server" ClientInstanceName="dtWeekETDVendor2"
                                            DisplayFormatString="dd MMM yyyy" EditFormat="Custom" EditFormatString="dd MMM yyyy"
                                            Font-Names="Tahoma" Font-Size="8pt" BackColor="GreenYellow" Width="150px">
                                            <Border BorderStyle="Groove" BorderColor="Silver" BorderWidth="1px" />
                                            <BorderLeft BorderColor="#CCCCCC" BorderStyle="Groove" BorderWidth="1px" />
                                            <BorderTop BorderColor="#CCCCCC" BorderStyle="Groove" BorderWidth="1px" />
                                            <BorderRight BorderColor="#CCCCCC" BorderStyle="Groove" BorderWidth="1px" />
                                            <BorderBottom BorderColor="#CCCCCC" BorderStyle="Groove" BorderWidth="1px" />
                                        </dx:ASPxTimeEdit>
                                    </td>
                                    <td width="150px" height="16px" style="border-style: solid none none solid; border-width: 0.1px;
                                        border-color: #808080;">
                                        <dx:ASPxTimeEdit ID="dtWeekETDVendor3" runat="server" ClientInstanceName="dtWeekETDVendor3"
                                            DisplayFormatString="dd MMM yyyy" EditFormat="Custom" EditFormatString="dd MMM yyyy"
                                            Font-Names="Tahoma" Font-Size="8pt" BackColor="GreenYellow" Width="150px">
                                            <Border BorderStyle="Groove" BorderColor="Silver" BorderWidth="1px" />
                                            <BorderLeft BorderColor="#CCCCCC" BorderStyle="Groove" BorderWidth="1px" />
                                            <BorderTop BorderColor="#CCCCCC" BorderStyle="Groove" BorderWidth="1px" />
                                            <BorderRight BorderColor="#CCCCCC" BorderStyle="Groove" BorderWidth="1px" />
                                            <BorderBottom BorderColor="#CCCCCC" BorderStyle="Groove" BorderWidth="1px" />
                                        </dx:ASPxTimeEdit>
                                    </td>
                                    <td width="150px" height="16px" style="border-style: solid none none solid; border-width: 0.1px;
                                        border-color: #808080;">
                                        <dx:ASPxTimeEdit ID="dtWeekETDVendor4" runat="server" ClientInstanceName="dtWeekETDVendor4"
                                            DisplayFormatString="dd MMM yyyy" EditFormat="Custom" EditFormatString="dd MMM yyyy"
                                            Font-Names="Tahoma" Font-Size="8pt" BackColor="GreenYellow" Width="150px">
                                            <Border BorderStyle="Groove" BorderColor="Silver" BorderWidth="1px" />
                                            <BorderLeft BorderColor="#CCCCCC" BorderStyle="Groove" BorderWidth="1px" />
                                            <BorderTop BorderColor="#CCCCCC" BorderStyle="Groove" BorderWidth="1px" />
                                            <BorderRight BorderColor="#CCCCCC" BorderStyle="Groove" BorderWidth="1px" />
                                            <BorderBottom BorderColor="#CCCCCC" BorderStyle="Groove" BorderWidth="1px" />
                                        </dx:ASPxTimeEdit>
                                    </td>
                                    <td width="150px" height="16px" style="border-style: solid none none solid; border-width: 0.1px;
                                        border-color: #808080;">
                                        <dx:ASPxTimeEdit ID="dtWeekETDVendor5" runat="server" ClientInstanceName="dtWeekETDVendor5"
                                            DisplayFormatString="dd MMM yyyy" EditFormat="Custom" EditFormatString="dd MMM yyyy"
                                            Font-Names="Tahoma" Font-Size="8pt" BackColor="GreenYellow" Width="150px">
                                            <Border BorderStyle="Groove" BorderColor="Silver" BorderWidth="1px" />
                                            <BorderLeft BorderColor="#CCCCCC" BorderStyle="Groove" BorderWidth="1px" />
                                            <BorderTop BorderColor="#CCCCCC" BorderStyle="Groove" BorderWidth="1px" />
                                            <BorderRight BorderColor="#CCCCCC" BorderStyle="Groove" BorderWidth="1px" />
                                            <BorderBottom BorderColor="#CCCCCC" BorderStyle="Groove" BorderWidth="1px" />
                                        </dx:ASPxTimeEdit>
                                    </td>
                                </tr>
                                <tr>
                                    <td width="150px" align="left" height="16px" style="border-style: solid none none solid;
                                        border-width: 0.1px; border-color: #808080;">
                                        <dx:ASPxTextBox runat="server" Width="175px" MaxLength="10" Height="21px" Font-Names="Tahoma"
                                            Font-Size="8pt" ID="ASPxTextBox7" BackColor="#FFD2A6" Text="ETD PORT" ReadOnly="True"
                                            Font-Bold="True">
                                            <Border BorderStyle="Groove" BorderColor="Silver" BorderWidth="1px" />
                                            <BorderLeft BorderColor="#CCCCCC" BorderStyle="Groove" BorderWidth="1px" />
                                            <BorderTop BorderColor="#CCCCCC" BorderStyle="Groove" BorderWidth="1px" />
                                            <BorderRight BorderColor="#CCCCCC" BorderStyle="Groove" BorderWidth="1px" />
                                            <BorderBottom BorderColor="#CCCCCC" BorderStyle="Groove" BorderWidth="1px" />
                                        </dx:ASPxTextBox>
                                    </td>
                                    <td width="150px" align="left" height="16px" style="border-style: solid none none solid;
                                        border-width: 0.1px; border-color: #808080;">
                                        <dx:ASPxTimeEdit ID="dtWeekETDPortOld1" runat="server" ClientInstanceName="dtWeekETDPortOld1"
                                            DisplayFormatString="dd MMM yyyy" EditFormat="Custom" EditFormatString="dd MMM yyyy"
                                            Font-Names="Tahoma" Font-Size="8pt" BackColor="#FFFFE1" Width="150px" ReadOnly="True">
                                            <ReadOnlyStyle BackColor="#CCCCCC">
                                            </ReadOnlyStyle>
                                            <Border BorderStyle="Groove" BorderColor="Silver" BorderWidth="1px" />
                                            <BorderLeft BorderColor="#CCCCCC" BorderStyle="Groove" BorderWidth="1px" />
                                            <BorderTop BorderColor="#CCCCCC" BorderStyle="Groove" BorderWidth="1px" />
                                            <BorderRight BorderColor="#CCCCCC" BorderStyle="Groove" BorderWidth="1px" />
                                            <BorderBottom BorderColor="#CCCCCC" BorderStyle="Groove" BorderWidth="1px" />
                                        </dx:ASPxTimeEdit>
                                    </td>
                                    <td width="150px" align="left" height="16px" style="border-style: solid none none solid;
                                        border-width: 0.1px; border-color: #808080;">
                                        <dx:ASPxTimeEdit ID="dtWeekETDPortOld2" runat="server" ClientInstanceName="dtWeekETDPortOld2"
                                            DisplayFormatString="dd MMM yyyy" EditFormat="Custom" EditFormatString="dd MMM yyyy"
                                            Font-Names="Tahoma" Font-Size="8pt" BackColor="#FFFFE1" Width="150px" ReadOnly="True">
                                            <ReadOnlyStyle BackColor="#CCCCCC">
                                            </ReadOnlyStyle>
                                            <Border BorderStyle="Groove" BorderColor="Silver" BorderWidth="1px" />
                                            <BorderLeft BorderColor="#CCCCCC" BorderStyle="Groove" BorderWidth="1px" />
                                            <BorderTop BorderColor="#CCCCCC" BorderStyle="Groove" BorderWidth="1px" />
                                            <BorderRight BorderColor="#CCCCCC" BorderStyle="Groove" BorderWidth="1px" />
                                            <BorderBottom BorderColor="#CCCCCC" BorderStyle="Groove" BorderWidth="1px" />
                                        </dx:ASPxTimeEdit>
                                    </td>
                                    <td width="150px" align="left" height="16px" style="border-style: solid none none solid;
                                        border-width: 0.1px; border-color: #808080;">
                                        <dx:ASPxTimeEdit ID="dtWeekETDPortOld3" runat="server" ClientInstanceName="dtWeekETDPortOld3"
                                            DisplayFormatString="dd MMM yyyy" EditFormat="Custom" EditFormatString="dd MMM yyyy"
                                            Font-Names="Tahoma" Font-Size="8pt" BackColor="#FFFFE1" Width="150px" ReadOnly="True">
                                            <ReadOnlyStyle BackColor="#CCCCCC">
                                            </ReadOnlyStyle>
                                            <Border BorderStyle="Groove" BorderColor="Silver" BorderWidth="1px" />
                                            <BorderLeft BorderColor="#CCCCCC" BorderStyle="Groove" BorderWidth="1px" />
                                            <BorderTop BorderColor="#CCCCCC" BorderStyle="Groove" BorderWidth="1px" />
                                            <BorderRight BorderColor="#CCCCCC" BorderStyle="Groove" BorderWidth="1px" />
                                            <BorderBottom BorderColor="#CCCCCC" BorderStyle="Groove" BorderWidth="1px" />
                                        </dx:ASPxTimeEdit>
                                    </td>
                                    <td width="150px" align="left" height="16px" style="border-style: solid none none solid;
                                        border-width: 0.1px; border-color: #808080;">
                                        <dx:ASPxTimeEdit ID="dtWeekETDPortOld4" runat="server" ClientInstanceName="dtWeekETDPortOld4"
                                            DisplayFormatString="dd MMM yyyy" EditFormat="Custom" EditFormatString="dd MMM yyyy"
                                            Font-Names="Tahoma" Font-Size="8pt" BackColor="#FFFFE1" Width="150px" ReadOnly="True">
                                            <ReadOnlyStyle BackColor="#CCCCCC">
                                            </ReadOnlyStyle>
                                            <Border BorderStyle="Groove" BorderColor="Silver" BorderWidth="1px" />
                                            <BorderLeft BorderColor="#CCCCCC" BorderStyle="Groove" BorderWidth="1px" />
                                            <BorderTop BorderColor="#CCCCCC" BorderStyle="Groove" BorderWidth="1px" />
                                            <BorderRight BorderColor="#CCCCCC" BorderStyle="Groove" BorderWidth="1px" />
                                            <BorderBottom BorderColor="#CCCCCC" BorderStyle="Groove" BorderWidth="1px" />
                                        </dx:ASPxTimeEdit>
                                    </td>
                                    <td width="150px" align="left" height="16px" style="border-style: solid none none solid;
                                        border-width: 0.1px; border-color: #808080;">
                                        <dx:ASPxTimeEdit ID="dtWeekETDPortOld5" runat="server" ClientInstanceName="dtWeekETDPortOld5"
                                            DisplayFormatString="dd MMM yyyy" EditFormat="Custom" EditFormatString="dd MMM yyyy"
                                            Font-Names="Tahoma" Font-Size="8pt" BackColor="#FFFFE1" Width="150px" ReadOnly="True">
                                            <ReadOnlyStyle BackColor="#CCCCCC">
                                            </ReadOnlyStyle>
                                            <Border BorderStyle="Groove" BorderColor="Silver" BorderWidth="1px" />
                                            <BorderLeft BorderColor="#CCCCCC" BorderStyle="Groove" BorderWidth="1px" />
                                            <BorderTop BorderColor="#CCCCCC" BorderStyle="Groove" BorderWidth="1px" />
                                            <BorderRight BorderColor="#CCCCCC" BorderStyle="Groove" BorderWidth="1px" />
                                            <BorderBottom BorderColor="#CCCCCC" BorderStyle="Groove" BorderWidth="1px" />
                                        </dx:ASPxTimeEdit>
                                    </td>
                                </tr>
                                <tr>
                                    <td width="150px" align="left" height="16px" style="border-style: solid none none solid;
                                        border-width: 0.1px; border-color: #808080;">
                                        <dx:ASPxTextBox runat="server" Width="175px" MaxLength="10" Height="21px" Font-Names="Tahoma"
                                            Font-Size="8pt" ID="ASPxTextBox11" BackColor="GreenYellow" Text="ETD PORT* (REVISION)"
                                            ReadOnly="True" ForeColor="Red" Font-Bold="True">
                                            <Border BorderStyle="Groove" BorderColor="Silver" BorderWidth="1px" />
                                            <BorderLeft BorderColor="#CCCCCC" BorderStyle="Groove" BorderWidth="1px" />
                                            <BorderTop BorderColor="#CCCCCC" BorderStyle="Groove" BorderWidth="1px" />
                                            <BorderRight BorderColor="#CCCCCC" BorderStyle="Groove" BorderWidth="1px" />
                                            <BorderBottom BorderColor="#CCCCCC" BorderStyle="Groove" BorderWidth="1px" />
                                        </dx:ASPxTextBox>
                                    </td>
                                    <td width="150px" align="left" height="16px" style="border-style: solid none none solid;
                                        border-width: 0.1px; border-color: #808080;">
                                        <dx:ASPxTimeEdit ID="dtWeekETDPort1" runat="server" ClientInstanceName="dtWeekETDPort1"
                                            DisplayFormatString="dd MMM yyyy" EditFormat="Custom" EditFormatString="dd MMM yyyy"
                                            Font-Names="Tahoma" Font-Size="8pt" BackColor="GreenYellow" Width="150px">
                                            <Border BorderStyle="Groove" BorderColor="Silver" BorderWidth="1px" />
                                            <BorderLeft BorderColor="#CCCCCC" BorderStyle="Groove" BorderWidth="1px" />
                                            <BorderTop BorderColor="#CCCCCC" BorderStyle="Groove" BorderWidth="1px" />
                                            <BorderRight BorderColor="#CCCCCC" BorderStyle="Groove" BorderWidth="1px" />
                                            <BorderBottom BorderColor="#CCCCCC" BorderStyle="Groove" BorderWidth="1px" />
                                        </dx:ASPxTimeEdit>
                                    </td>
                                    <td width="150px" align="left" height="16px" style="border-style: solid none none solid;
                                        border-width: 0.1px; border-color: #808080;">
                                        <dx:ASPxTimeEdit ID="dtWeekETDPort2" runat="server" ClientInstanceName="dtWeekETDPort2"
                                            DisplayFormatString="dd MMM yyyy" EditFormat="Custom" EditFormatString="dd MMM yyyy"
                                            Font-Names="Tahoma" Font-Size="8pt" BackColor="GreenYellow" Width="150px">
                                            <Border BorderStyle="Groove" BorderColor="Silver" BorderWidth="1px" />
                                            <BorderLeft BorderColor="#CCCCCC" BorderStyle="Groove" BorderWidth="1px" />
                                            <BorderTop BorderColor="#CCCCCC" BorderStyle="Groove" BorderWidth="1px" />
                                            <BorderRight BorderColor="#CCCCCC" BorderStyle="Groove" BorderWidth="1px" />
                                            <BorderBottom BorderColor="#CCCCCC" BorderStyle="Groove" BorderWidth="1px" />
                                        </dx:ASPxTimeEdit>
                                    </td>
                                    <td width="150px" align="left" height="16px" style="border-style: solid none none solid;
                                        border-width: 0.1px; border-color: #808080;">
                                        <dx:ASPxTimeEdit ID="dtWeekETDPort3" runat="server" ClientInstanceName="dtWeekETDPort3"
                                            DisplayFormatString="dd MMM yyyy" EditFormat="Custom" EditFormatString="dd MMM yyyy"
                                            Font-Names="Tahoma" Font-Size="8pt" BackColor="GreenYellow" Width="150px">
                                            <Border BorderStyle="Groove" BorderColor="Silver" BorderWidth="1px" />
                                            <BorderLeft BorderColor="#CCCCCC" BorderStyle="Groove" BorderWidth="1px" />
                                            <BorderTop BorderColor="#CCCCCC" BorderStyle="Groove" BorderWidth="1px" />
                                            <BorderRight BorderColor="#CCCCCC" BorderStyle="Groove" BorderWidth="1px" />
                                            <BorderBottom BorderColor="#CCCCCC" BorderStyle="Groove" BorderWidth="1px" />
                                        </dx:ASPxTimeEdit>
                                    </td>
                                    <td width="150px" align="left" height="16px" style="border-style: solid none none solid;
                                        border-width: 0.1px; border-color: #808080;">
                                        <dx:ASPxTimeEdit ID="dtWeekETDPort4" runat="server" ClientInstanceName="dtWeekETDPort4"
                                            DisplayFormatString="dd MMM yyyy" EditFormat="Custom" EditFormatString="dd MMM yyyy"
                                            Font-Names="Tahoma" Font-Size="8pt" BackColor="GreenYellow" Width="150px">
                                            <Border BorderStyle="Groove" BorderColor="Silver" BorderWidth="1px" />
                                            <BorderLeft BorderColor="#CCCCCC" BorderStyle="Groove" BorderWidth="1px" />
                                            <BorderTop BorderColor="#CCCCCC" BorderStyle="Groove" BorderWidth="1px" />
                                            <BorderRight BorderColor="#CCCCCC" BorderStyle="Groove" BorderWidth="1px" />
                                            <BorderBottom BorderColor="#CCCCCC" BorderStyle="Groove" BorderWidth="1px" />
                                        </dx:ASPxTimeEdit>
                                    </td>
                                    <td width="150px" align="left" height="16px" style="border-style: solid none none solid;
                                        border-width: 0.1px; border-color: #808080;">
                                        <dx:ASPxTimeEdit ID="dtWeekETDPort5" runat="server" ClientInstanceName="dtWeekETDPort5"
                                            DisplayFormatString="dd MMM yyyy" EditFormat="Custom" EditFormatString="dd MMM yyyy"
                                            Font-Names="Tahoma" Font-Size="8pt" BackColor="GreenYellow" Width="150px">
                                            <Border BorderStyle="Groove" BorderColor="Silver" BorderWidth="1px" />
                                            <BorderLeft BorderColor="#CCCCCC" BorderStyle="Groove" BorderWidth="1px" />
                                            <BorderTop BorderColor="#CCCCCC" BorderStyle="Groove" BorderWidth="1px" />
                                            <BorderRight BorderColor="#CCCCCC" BorderStyle="Groove" BorderWidth="1px" />
                                            <BorderBottom BorderColor="#CCCCCC" BorderStyle="Groove" BorderWidth="1px" />
                                        </dx:ASPxTimeEdit>
                                    </td>
                                </tr>
                                <tr>
                                    <td width="150px" align="left" style="border-style: solid none none solid; border-width: 0.1px;
                                        border-color: #808080;" class="style1">
                                        <dx:ASPxTextBox runat="server" Width="175px" MaxLength="10" Height="21px" Font-Names="Tahoma"
                                            Font-Size="8pt" ID="ASPxTextBox8" BackColor="#FFD2A6" Text="ETA PORT" ReadOnly="True"
                                            Font-Bold="True">
                                            <Border BorderStyle="Groove" BorderColor="Silver" BorderWidth="1px" />
                                            <BorderLeft BorderColor="#CCCCCC" BorderStyle="Groove" BorderWidth="1px" />
                                            <BorderTop BorderColor="#CCCCCC" BorderStyle="Groove" BorderWidth="1px" />
                                            <BorderRight BorderColor="#CCCCCC" BorderStyle="Groove" BorderWidth="1px" />
                                            <BorderBottom BorderColor="#CCCCCC" BorderStyle="Groove" BorderWidth="1px" />
                                        </dx:ASPxTextBox>
                                    </td>
                                    <td width="150px" align="left" style="border-style: solid none none solid; border-width: 0.1px;
                                        border-color: #808080;" class="style1">
                                        <dx:ASPxTimeEdit ID="dtWeekETAPortOld1" runat="server" ClientInstanceName="dtWeekETAPortOld1"
                                            DisplayFormatString="dd MMM yyyy" EditFormat="Custom" EditFormatString="dd MMM yyyy"
                                            Font-Names="Tahoma" Font-Size="8pt" BackColor="#FFFFE1" Width="150px" ReadOnly="True">
                                            <ReadOnlyStyle BackColor="#CCCCCC">
                                            </ReadOnlyStyle>
                                            <Border BorderStyle="Groove" BorderColor="Silver" BorderWidth="1px" />
                                            <BorderLeft BorderColor="#CCCCCC" BorderStyle="Groove" BorderWidth="1px" />
                                            <BorderTop BorderColor="#CCCCCC" BorderStyle="Groove" BorderWidth="1px" />
                                            <BorderRight BorderColor="#CCCCCC" BorderStyle="Groove" BorderWidth="1px" />
                                            <BorderBottom BorderColor="#CCCCCC" BorderStyle="Groove" BorderWidth="1px" />
                                        </dx:ASPxTimeEdit>
                                    </td>
                                    <td width="150px" align="left" style="border-style: solid none none solid; border-width: 0.1px;
                                        border-color: #808080;" class="style1">
                                        <dx:ASPxTimeEdit ID="dtWeekETAPortOld2" runat="server" ClientInstanceName="dtWeekETAPortOld2"
                                            DisplayFormatString="dd MMM yyyy" EditFormat="Custom" EditFormatString="dd MMM yyyy"
                                            Font-Names="Tahoma" Font-Size="8pt" BackColor="#FFFFE1" Width="150px" ReadOnly="True">
                                            <ReadOnlyStyle BackColor="#CCCCCC">
                                            </ReadOnlyStyle>
                                            <Border BorderStyle="Groove" BorderColor="Silver" BorderWidth="1px" />
                                            <BorderLeft BorderColor="#CCCCCC" BorderStyle="Groove" BorderWidth="1px" />
                                            <BorderTop BorderColor="#CCCCCC" BorderStyle="Groove" BorderWidth="1px" />
                                            <BorderRight BorderColor="#CCCCCC" BorderStyle="Groove" BorderWidth="1px" />
                                            <BorderBottom BorderColor="#CCCCCC" BorderStyle="Groove" BorderWidth="1px" />
                                        </dx:ASPxTimeEdit>
                                    </td>
                                    <td width="150px" align="left" style="border-style: solid none none solid; border-width: 0.1px;
                                        border-color: #808080;" class="style1">
                                        <dx:ASPxTimeEdit ID="dtWeekETAPortOld3" runat="server" ClientInstanceName="dtWeekETAPortOld3"
                                            DisplayFormatString="dd MMM yyyy" EditFormat="Custom" EditFormatString="dd MMM yyyy"
                                            Font-Names="Tahoma" Font-Size="8pt" BackColor="#FFFFE1" Width="150px" ReadOnly="True">
                                            <ReadOnlyStyle BackColor="#CCCCCC">
                                            </ReadOnlyStyle>
                                            <Border BorderStyle="Groove" BorderColor="Silver" BorderWidth="1px" />
                                            <BorderLeft BorderColor="#CCCCCC" BorderStyle="Groove" BorderWidth="1px" />
                                            <BorderTop BorderColor="#CCCCCC" BorderStyle="Groove" BorderWidth="1px" />
                                            <BorderRight BorderColor="#CCCCCC" BorderStyle="Groove" BorderWidth="1px" />
                                            <BorderBottom BorderColor="#CCCCCC" BorderStyle="Groove" BorderWidth="1px" />
                                        </dx:ASPxTimeEdit>
                                    </td>
                                    <td width="150px" align="left" style="border-style: solid none none solid; border-width: 0.1px;
                                        border-color: #808080;" class="style1">
                                        <dx:ASPxTimeEdit ID="dtWeekETAPortOld4" runat="server" ClientInstanceName="dtWeekETAPortOld4"
                                            DisplayFormatString="dd MMM yyyy" EditFormat="Custom" EditFormatString="dd MMM yyyy"
                                            Font-Names="Tahoma" Font-Size="8pt" BackColor="#FFFFE1" Width="150px" ReadOnly="True">
                                            <ReadOnlyStyle BackColor="#CCCCCC">
                                            </ReadOnlyStyle>
                                            <Border BorderStyle="Groove" BorderColor="Silver" BorderWidth="1px" />
                                            <BorderLeft BorderColor="#CCCCCC" BorderStyle="Groove" BorderWidth="1px" />
                                            <BorderTop BorderColor="#CCCCCC" BorderStyle="Groove" BorderWidth="1px" />
                                            <BorderRight BorderColor="#CCCCCC" BorderStyle="Groove" BorderWidth="1px" />
                                            <BorderBottom BorderColor="#CCCCCC" BorderStyle="Groove" BorderWidth="1px" />
                                        </dx:ASPxTimeEdit>
                                    </td>
                                    <td width="150px" align="left" style="border-style: solid none none solid; border-width: 0.1px;
                                        border-color: #808080;" class="style1">
                                        <dx:ASPxTimeEdit ID="dtWeekETAPortOld5" runat="server" ClientInstanceName="dtWeekETAPortOld5"
                                            DisplayFormatString="dd MMM yyyy" EditFormat="Custom" EditFormatString="dd MMM yyyy"
                                            Font-Names="Tahoma" Font-Size="8pt" BackColor="#FFFFE1" Width="150px" ReadOnly="True">
                                            <ReadOnlyStyle BackColor="#CCCCCC">
                                            </ReadOnlyStyle>
                                            <Border BorderStyle="Groove" BorderColor="Silver" BorderWidth="1px" />
                                            <BorderLeft BorderColor="#CCCCCC" BorderStyle="Groove" BorderWidth="1px" />
                                            <BorderTop BorderColor="#CCCCCC" BorderStyle="Groove" BorderWidth="1px" />
                                            <BorderRight BorderColor="#CCCCCC" BorderStyle="Groove" BorderWidth="1px" />
                                            <BorderBottom BorderColor="#CCCCCC" BorderStyle="Groove" BorderWidth="1px" />
                                        </dx:ASPxTimeEdit>
                                    </td>
                                </tr>
                                <tr>
                                    <td width="150px" align="left" style="border-style: solid none none solid; border-width: 0.1px;
                                        border-color: #808080;" class="style1">
                                        <dx:ASPxTextBox runat="server" Width="175px" MaxLength="10" Height="21px" Font-Names="Tahoma"
                                            Font-Size="8pt" ID="ASPxTextBox12" BackColor="GreenYellow" Text="ETA PORT* (REVISION)"
                                            ReadOnly="True" ForeColor="Red" Font-Bold="True">
                                            <Border BorderStyle="Groove" BorderColor="Silver" BorderWidth="1px" />
                                            <BorderLeft BorderColor="#CCCCCC" BorderStyle="Groove" BorderWidth="1px" />
                                            <BorderTop BorderColor="#CCCCCC" BorderStyle="Groove" BorderWidth="1px" />
                                            <BorderRight BorderColor="#CCCCCC" BorderStyle="Groove" BorderWidth="1px" />
                                            <BorderBottom BorderColor="#CCCCCC" BorderStyle="Groove" BorderWidth="1px" />
                                        </dx:ASPxTextBox>
                                    </td>
                                    <td width="150px" align="left" style="border-style: solid none none solid; border-width: 0.1px;
                                        border-color: #808080;" class="style1">
                                        <dx:ASPxTimeEdit ID="dtWeekETAPort1" runat="server" ClientInstanceName="dtWeekETAPort1"
                                            DisplayFormatString="dd MMM yyyy" EditFormat="Custom" EditFormatString="dd MMM yyyy"
                                            Font-Names="Tahoma" Font-Size="8pt" BackColor="GreenYellow" Width="150px">
                                            <Border BorderStyle="Groove" BorderColor="Silver" BorderWidth="1px" />
                                            <BorderLeft BorderColor="#CCCCCC" BorderStyle="Groove" BorderWidth="1px" />
                                            <BorderTop BorderColor="#CCCCCC" BorderStyle="Groove" BorderWidth="1px" />
                                            <BorderRight BorderColor="#CCCCCC" BorderStyle="Groove" BorderWidth="1px" />
                                            <BorderBottom BorderColor="#CCCCCC" BorderStyle="Groove" BorderWidth="1px" />
                                        </dx:ASPxTimeEdit>
                                    </td>
                                    <td width="150px" align="left" style="border-style: solid none none solid; border-width: 0.1px;
                                        border-color: #808080;" class="style1">
                                        <dx:ASPxTimeEdit ID="dtWeekETAPort2" runat="server" ClientInstanceName="dtWeekETAPort2"
                                            DisplayFormatString="dd MMM yyyy" EditFormat="Custom" EditFormatString="dd MMM yyyy"
                                            Font-Names="Tahoma" Font-Size="8pt" BackColor="GreenYellow" Width="150px">
                                            <Border BorderStyle="Groove" BorderColor="Silver" BorderWidth="1px" />
                                            <BorderLeft BorderColor="#CCCCCC" BorderStyle="Groove" BorderWidth="1px" />
                                            <BorderTop BorderColor="#CCCCCC" BorderStyle="Groove" BorderWidth="1px" />
                                            <BorderRight BorderColor="#CCCCCC" BorderStyle="Groove" BorderWidth="1px" />
                                            <BorderBottom BorderColor="#CCCCCC" BorderStyle="Groove" BorderWidth="1px" />
                                        </dx:ASPxTimeEdit>
                                    </td>
                                    <td width="150px" align="left" style="border-style: solid none none solid; border-width: 0.1px;
                                        border-color: #808080;" class="style1">
                                        <dx:ASPxTimeEdit ID="dtWeekETAPort3" runat="server" ClientInstanceName="dtWeekETAPort3"
                                            DisplayFormatString="dd MMM yyyy" EditFormat="Custom" EditFormatString="dd MMM yyyy"
                                            Font-Names="Tahoma" Font-Size="8pt" BackColor="GreenYellow" Width="150px">
                                            <Border BorderStyle="Groove" BorderColor="Silver" BorderWidth="1px" />
                                            <BorderLeft BorderColor="#CCCCCC" BorderStyle="Groove" BorderWidth="1px" />
                                            <BorderTop BorderColor="#CCCCCC" BorderStyle="Groove" BorderWidth="1px" />
                                            <BorderRight BorderColor="#CCCCCC" BorderStyle="Groove" BorderWidth="1px" />
                                            <BorderBottom BorderColor="#CCCCCC" BorderStyle="Groove" BorderWidth="1px" />
                                        </dx:ASPxTimeEdit>
                                    </td>
                                    <td width="150px" align="left" style="border-style: solid none none solid; border-width: 0.1px;
                                        border-color: #808080;" class="style1">
                                        <dx:ASPxTimeEdit ID="dtWeekETAPort4" runat="server" ClientInstanceName="dtWeekETAPort4"
                                            DisplayFormatString="dd MMM yyyy" EditFormat="Custom" EditFormatString="dd MMM yyyy"
                                            Font-Names="Tahoma" Font-Size="8pt" BackColor="GreenYellow" Width="150px">
                                            <Border BorderStyle="Groove" BorderColor="Silver" BorderWidth="1px" />
                                            <BorderLeft BorderColor="#CCCCCC" BorderStyle="Groove" BorderWidth="1px" />
                                            <BorderTop BorderColor="#CCCCCC" BorderStyle="Groove" BorderWidth="1px" />
                                            <BorderRight BorderColor="#CCCCCC" BorderStyle="Groove" BorderWidth="1px" />
                                            <BorderBottom BorderColor="#CCCCCC" BorderStyle="Groove" BorderWidth="1px" />
                                        </dx:ASPxTimeEdit>
                                    </td>
                                    <td width="150px" align="left" style="border-style: solid none none solid; border-width: 0.1px;
                                        border-color: #808080;" class="style1">
                                        <dx:ASPxTimeEdit ID="dtWeekETAPort5" runat="server" ClientInstanceName="dtWeekETAPort5"
                                            DisplayFormatString="dd MMM yyyy" EditFormat="Custom" EditFormatString="dd MMM yyyy"
                                            Font-Names="Tahoma" Font-Size="8pt" BackColor="GreenYellow" Width="150px">
                                            <Border BorderStyle="Groove" BorderColor="Silver" BorderWidth="1px" />
                                            <BorderLeft BorderColor="#CCCCCC" BorderStyle="Groove" BorderWidth="1px" />
                                            <BorderTop BorderColor="#CCCCCC" BorderStyle="Groove" BorderWidth="1px" />
                                            <BorderRight BorderColor="#CCCCCC" BorderStyle="Groove" BorderWidth="1px" />
                                            <BorderBottom BorderColor="#CCCCCC" BorderStyle="Groove" BorderWidth="1px" />
                                        </dx:ASPxTimeEdit>
                                    </td>
                                </tr>
                                <tr>
                                    <td width="150px" align="left" height="16px" style="border-style: solid none none solid;
                                        border-width: 0.1px; border-color: #808080;">
                                        <dx:ASPxTextBox runat="server" Width="175px" MaxLength="10" Height="21px" Font-Names="Tahoma"
                                            Font-Size="8pt" ID="ASPxTextBox9" BackColor="#FFD2A6" Text="ETA FACTORY" ReadOnly="True"
                                            Font-Bold="True">
                                            <Border BorderStyle="Groove" BorderColor="Silver" BorderWidth="1px" />
                                            <BorderLeft BorderColor="#CCCCCC" BorderStyle="Groove" BorderWidth="1px" />
                                            <BorderTop BorderColor="#CCCCCC" BorderStyle="Groove" BorderWidth="1px" />
                                            <BorderRight BorderColor="#CCCCCC" BorderStyle="Groove" BorderWidth="1px" />
                                            <BorderBottom BorderColor="#CCCCCC" BorderStyle="Groove" BorderWidth="1px" />
                                        </dx:ASPxTextBox>
                                    </td>
                                    <td width="150px" align="left" height="16px" style="border-style: solid none none solid;
                                        border-width: 0.1px; border-color: #808080;">
                                        <dx:ASPxTimeEdit ID="dtETAFactWeekOld1" runat="server" ClientInstanceName="dtETAFactWeekOld1"
                                            DisplayFormatString="dd MMM yyyy" EditFormat="Custom" EditFormatString="dd MMM yyyy"
                                            Font-Names="Tahoma" Font-Size="8pt" BackColor="#FFFFE1" Width="150px" ReadOnly="True">
                                            <ReadOnlyStyle BackColor="#CCCCCC">
                                            </ReadOnlyStyle>
                                            <Border BorderStyle="Groove" BorderColor="Silver" BorderWidth="1px" />
                                            <BorderLeft BorderColor="#CCCCCC" BorderStyle="Groove" BorderWidth="1px" />
                                            <BorderTop BorderColor="#CCCCCC" BorderStyle="Groove" BorderWidth="1px" />
                                            <BorderRight BorderColor="#CCCCCC" BorderStyle="Groove" BorderWidth="1px" />
                                            <BorderBottom BorderColor="#CCCCCC" BorderStyle="Groove" BorderWidth="1px" />
                                        </dx:ASPxTimeEdit>
                                    </td>
                                    <td width="150px" align="left" height="16px" style="border-style: solid none none solid;
                                        border-width: 0.1px; border-color: #808080;">
                                        <dx:ASPxTimeEdit ID="dtETAFactWeekOld2" runat="server" ClientInstanceName="dtETAFactWeekOld2"
                                            DisplayFormatString="dd MMM yyyy" EditFormat="Custom" EditFormatString="dd MMM yyyy"
                                            Font-Names="Tahoma" Font-Size="8pt" BackColor="#FFFFE1" Width="150px" ReadOnly="True">
                                            <ReadOnlyStyle BackColor="#CCCCCC">
                                            </ReadOnlyStyle>
                                            <Border BorderStyle="Groove" BorderColor="Silver" BorderWidth="1px" />
                                            <BorderLeft BorderColor="#CCCCCC" BorderStyle="Groove" BorderWidth="1px" />
                                            <BorderTop BorderColor="#CCCCCC" BorderStyle="Groove" BorderWidth="1px" />
                                            <BorderRight BorderColor="#CCCCCC" BorderStyle="Groove" BorderWidth="1px" />
                                            <BorderBottom BorderColor="#CCCCCC" BorderStyle="Groove" BorderWidth="1px" />
                                        </dx:ASPxTimeEdit>
                                    </td>
                                    <td width="150px" align="left" height="16px" style="border-style: solid none none solid;
                                        border-width: 0.1px; border-color: #808080;">
                                        <dx:ASPxTimeEdit ID="dtETAFactWeekOld3" runat="server" ClientInstanceName="dtETAFactWeekOld3"
                                            DisplayFormatString="dd MMM yyyy" EditFormat="Custom" EditFormatString="dd MMM yyyy"
                                            Font-Names="Tahoma" Font-Size="8pt" BackColor="#FFFFE1" Width="150px" ReadOnly="True">
                                            <ReadOnlyStyle BackColor="#CCCCCC">
                                            </ReadOnlyStyle>
                                            <Border BorderStyle="Groove" BorderColor="Silver" BorderWidth="1px" />
                                            <BorderLeft BorderColor="#CCCCCC" BorderStyle="Groove" BorderWidth="1px" />
                                            <BorderTop BorderColor="#CCCCCC" BorderStyle="Groove" BorderWidth="1px" />
                                            <BorderRight BorderColor="#CCCCCC" BorderStyle="Groove" BorderWidth="1px" />
                                            <BorderBottom BorderColor="#CCCCCC" BorderStyle="Groove" BorderWidth="1px" />
                                        </dx:ASPxTimeEdit>
                                    </td>
                                    <td width="150px" align="left" height="16px" style="border-style: solid none none solid;
                                        border-width: 0.1px; border-color: #808080;">
                                        <dx:ASPxTimeEdit ID="dtETAFactWeekOld4" runat="server" ClientInstanceName="dtETAFactWeekOld4"
                                            DisplayFormatString="dd MMM yyyy" EditFormat="Custom" EditFormatString="dd MMM yyyy"
                                            Font-Names="Tahoma" Font-Size="8pt" BackColor="#FFFFE1" Width="150px" ReadOnly="True">
                                            <ReadOnlyStyle BackColor="#CCCCCC">
                                            </ReadOnlyStyle>
                                            <Border BorderStyle="Groove" BorderColor="Silver" BorderWidth="1px" />
                                            <BorderLeft BorderColor="#CCCCCC" BorderStyle="Groove" BorderWidth="1px" />
                                            <BorderTop BorderColor="#CCCCCC" BorderStyle="Groove" BorderWidth="1px" />
                                            <BorderRight BorderColor="#CCCCCC" BorderStyle="Groove" BorderWidth="1px" />
                                            <BorderBottom BorderColor="#CCCCCC" BorderStyle="Groove" BorderWidth="1px" />
                                        </dx:ASPxTimeEdit>
                                    </td>
                                    <td width="150px" align="left" height="16px" style="border-style: solid none none solid;
                                        border-width: 0.1px; border-color: #808080;">
                                        <dx:ASPxTimeEdit ID="dtETAFactWeekOld5" runat="server" ClientInstanceName="dtETAFactWeekOld5"
                                            DisplayFormatString="dd MMM yyyy" EditFormat="Custom" EditFormatString="dd MMM yyyy"
                                            Font-Names="Tahoma" Font-Size="8pt" BackColor="#FFFFE1" Width="150px" ReadOnly="True">
                                            <ReadOnlyStyle BackColor="#CCCCCC">
                                            </ReadOnlyStyle>
                                            <Border BorderStyle="Groove" BorderColor="Silver" BorderWidth="1px" />
                                            <BorderLeft BorderColor="#CCCCCC" BorderStyle="Groove" BorderWidth="1px" />
                                            <BorderTop BorderColor="#CCCCCC" BorderStyle="Groove" BorderWidth="1px" />
                                            <BorderRight BorderColor="#CCCCCC" BorderStyle="Groove" BorderWidth="1px" />
                                            <BorderBottom BorderColor="#CCCCCC" BorderStyle="Groove" BorderWidth="1px" />
                                        </dx:ASPxTimeEdit>
                                    </td>
                                </tr>
                                <tr>
                                    <td width="150px" align="left" height="16px" style="border-style: solid none none solid;
                                        border-width: 0.1px; border-color: #808080;">
                                        <dx:ASPxTextBox runat="server" Width="175px" MaxLength="10" Height="21px" Font-Names="Tahoma"
                                            Font-Size="8pt" ID="ASPxTextBox13" BackColor="GreenYellow" Text="ETA FACTORY*(REVISION)"
                                            ReadOnly="True" ForeColor="Red" Font-Bold="True">
                                            <Border BorderStyle="Groove" BorderColor="Silver" BorderWidth="1px" />
                                            <BorderLeft BorderColor="#CCCCCC" BorderStyle="Groove" BorderWidth="1px" />
                                            <BorderTop BorderColor="#CCCCCC" BorderStyle="Groove" BorderWidth="1px" />
                                            <BorderRight BorderColor="#CCCCCC" BorderStyle="Groove" BorderWidth="1px" />
                                            <BorderBottom BorderColor="#CCCCCC" BorderStyle="Groove" BorderWidth="1px" />
                                        </dx:ASPxTextBox>
                                    </td>
                                    <td width="150px" align="left" height="16px" style="border-style: solid none none solid;
                                        border-width: 0.1px; border-color: #808080;">
                                        <dx:ASPxTimeEdit ID="dtETAFactWeek1" runat="server" ClientInstanceName="dtETAFactWeek1"
                                            DisplayFormatString="dd MMM yyyy" EditFormat="Custom" EditFormatString="dd MMM yyyy"
                                            Font-Names="Tahoma" Font-Size="8pt" BackColor="GreenYellow" Width="150px">
                                            <Border BorderStyle="Groove" BorderColor="Silver" BorderWidth="1px" />
                                            <BorderLeft BorderColor="#CCCCCC" BorderStyle="Groove" BorderWidth="1px" />
                                            <BorderTop BorderColor="#CCCCCC" BorderStyle="Groove" BorderWidth="1px" />
                                            <BorderRight BorderColor="#CCCCCC" BorderStyle="Groove" BorderWidth="1px" />
                                            <BorderBottom BorderColor="#CCCCCC" BorderStyle="Groove" BorderWidth="1px" />
                                        </dx:ASPxTimeEdit>
                                    </td>
                                    <td width="150px" align="left" height="16px" style="border-style: solid none none solid;
                                        border-width: 0.1px; border-color: #808080;">
                                        <dx:ASPxTimeEdit ID="dtETAFactWeek2" runat="server" ClientInstanceName="dtETAFactWeek2"
                                            DisplayFormatString="dd MMM yyyy" EditFormat="Custom" EditFormatString="dd MMM yyyy"
                                            Font-Names="Tahoma" Font-Size="8pt" BackColor="GreenYellow" Width="150px">
                                            <Border BorderStyle="Groove" BorderColor="Silver" BorderWidth="1px" />
                                            <BorderLeft BorderColor="#CCCCCC" BorderStyle="Groove" BorderWidth="1px" />
                                            <BorderTop BorderColor="#CCCCCC" BorderStyle="Groove" BorderWidth="1px" />
                                            <BorderRight BorderColor="#CCCCCC" BorderStyle="Groove" BorderWidth="1px" />
                                            <BorderBottom BorderColor="#CCCCCC" BorderStyle="Groove" BorderWidth="1px" />
                                        </dx:ASPxTimeEdit>
                                    </td>
                                    <td width="150px" align="left" height="16px" style="border-style: solid none none solid;
                                        border-width: 0.1px; border-color: #808080;">
                                        <dx:ASPxTimeEdit ID="dtETAFactWeek3" runat="server" ClientInstanceName="dtETAFactWeek3"
                                            DisplayFormatString="dd MMM yyyy" EditFormat="Custom" EditFormatString="dd MMM yyyy"
                                            Font-Names="Tahoma" Font-Size="8pt" BackColor="GreenYellow" Width="150px">
                                            <Border BorderStyle="Groove" BorderColor="Silver" BorderWidth="1px" />
                                            <BorderLeft BorderColor="#CCCCCC" BorderStyle="Groove" BorderWidth="1px" />
                                            <BorderTop BorderColor="#CCCCCC" BorderStyle="Groove" BorderWidth="1px" />
                                            <BorderRight BorderColor="#CCCCCC" BorderStyle="Groove" BorderWidth="1px" />
                                            <BorderBottom BorderColor="#CCCCCC" BorderStyle="Groove" BorderWidth="1px" />
                                        </dx:ASPxTimeEdit>
                                    </td>
                                    <td width="150px" align="left" height="16px" style="border-style: solid none none solid;
                                        border-width: 0.1px; border-color: #808080;">
                                        <dx:ASPxTimeEdit ID="dtETAFactWeek4" runat="server" ClientInstanceName="dtETAFactWeek4"
                                            DisplayFormatString="dd MMM yyyy" EditFormat="Custom" EditFormatString="dd MMM yyyy"
                                            Font-Names="Tahoma" Font-Size="8pt" BackColor="GreenYellow" Width="150px">
                                            <Border BorderStyle="Groove" BorderColor="Silver" BorderWidth="1px" />
                                            <BorderLeft BorderColor="#CCCCCC" BorderStyle="Groove" BorderWidth="1px" />
                                            <BorderTop BorderColor="#CCCCCC" BorderStyle="Groove" BorderWidth="1px" />
                                            <BorderRight BorderColor="#CCCCCC" BorderStyle="Groove" BorderWidth="1px" />
                                            <BorderBottom BorderColor="#CCCCCC" BorderStyle="Groove" BorderWidth="1px" />
                                        </dx:ASPxTimeEdit>
                                    </td>
                                    <td width="150px" align="left" height="16px" style="border-style: solid none none solid;
                                        border-width: 0.1px; border-color: #808080;">
                                        <dx:ASPxTimeEdit ID="dtETAFactWeek5" runat="server" ClientInstanceName="dtETAFactWeek5"
                                            DisplayFormatString="dd MMM yyyy" EditFormat="Custom" EditFormatString="dd MMM yyyy"
                                            Font-Names="Tahoma" Font-Size="8pt" BackColor="GreenYellow" Width="150px">
                                            <Border BorderStyle="Groove" BorderColor="Silver" BorderWidth="1px" />
                                            <BorderLeft BorderColor="#CCCCCC" BorderStyle="Groove" BorderWidth="1px" />
                                            <BorderTop BorderColor="#CCCCCC" BorderStyle="Groove" BorderWidth="1px" />
                                            <BorderRight BorderColor="#CCCCCC" BorderStyle="Groove" BorderWidth="1px" />
                                            <BorderBottom BorderColor="#CCCCCC" BorderStyle="Groove" BorderWidth="1px" />
                                        </dx:ASPxTimeEdit>
                                    </td>
                                </tr>
                                <tr>
                                    <td width="150px" align="left" height="16px" style="border-style: solid none none solid;
                                        border-width: 0.1px; border-color: #808080;">
                                        &nbsp;
                                    </td>
                                    <td width="150px" align="left" height="16px" style="border-style: solid none none solid;
                                        border-width: 0.1px; border-color: #808080;">
                                        &nbsp;
                                    </td>
                                    <td width="150px" align="left" height="16px" style="border-style: solid none none solid;
                                        border-width: 0.1px; border-color: #808080;">
                                        &nbsp;
                                    </td>
                                    <td width="150px" align="left" height="16px" style="border-style: solid none none solid;
                                        border-width: 0.1px; border-color: #808080;">
                                        &nbsp;
                                    </td>
                                    <td width="150px" align="left" height="16px" style="border-style: solid none none solid;
                                        border-width: 0.1px; border-color: #808080;">
                                        &nbsp;
                                    </td>
                                    <td width="150px" align="left" height="16px" style="border-style: solid none none solid;
                                        border-width: 0.1px; border-color: #808080;">
                                        &nbsp;
                                    </td>
                                </tr>
                            </table>
                        </td>
                        <td width="150px">
                            &nbsp;
                        </td>
                        <td width="150px">
                            &nbsp;
                        </td>
                    </tr>
                    <tr>
                        <td width="150px">
                            &nbsp;
                        </td>
                        <td width="150px">
                            &nbsp;
                        </td>
                        <td width="150px">
                            &nbsp;
                        </td>
                        <td width="150px">
                            &nbsp;
                        </td>
                    </tr>
                    <tr>
                        <td width="150px">
                            &nbsp;
                        </td>
                        <td width="150px">
                            &nbsp;
                        </td>
                        <td width="150px">
                            &nbsp;
                        </td>
                        <td width="150px">
                            &nbsp;
                        </td>
                    </tr>
                    <tr>
                        <td width="150px">
                            &nbsp;
                        </td>
                        <td width="150px">
                            &nbsp;
                        </td>
                        <td width="150px">
                            &nbsp;
                        </td>
                        <td width="150px">
                            &nbsp;
                        </td>
                    </tr>
                    <tr>
                        <td width="150px">
                            &nbsp;
                        </td>
                        <td width="150px">
                            &nbsp;
                        </td>
                        <td width="150px">
                            &nbsp;
                        </td>
                        <td width="150px">
                            &nbsp;
                        </td>
                    </tr>
                </table>
            </td>
        </tr>
        <div style="height: 1px;">
        </div>
        <tr>
            <td colspan="8" align="left" valign="top" height="220">
                <%--Column : Grid--%>
                <dx:ASPxGridView ID="grid" runat="server" AutoGenerateColumns="False" ClientInstanceName="grid"
                    Font-Names="Tahoma" Font-Size="8pt" KeyFieldName="PartNo1;AffiliateName" Width="100%">
                    <ClientSideEvents CallbackError="function(s, e) {e.handled = true;}" EndCallback="function(s, e) {
                        if (s.cpSearch != '') {
                            txtOrderNoWeek1.SetText(s.cpOrder1);
                            txtOrderNoWeek2.SetText(s.cpOrder2);
                            txtOrderNoWeek3.SetText(s.cpOrder3);
                            txtOrderNoWeek4.SetText(s.cpOrder4);
                            txtOrderNoWeek5.SetText(s.cpOrder5);

                            dtWeekETDVendorOld1.SetText(s.cpETDVendorOld1);
                            dtWeekETDVendorOld2.SetText(s.cpETDVendorOld2);
                            dtWeekETDVendorOld3.SetText(s.cpETDVendorOld3);
                            dtWeekETDVendorOld4.SetText(s.cpETDVendorOld4);
                            dtWeekETDVendorOld5.SetText(s.cpETDVendorOld5);

                            dtWeekETDVendor1.SetText(s.cpETDVendor1);
                            dtWeekETDVendor2.SetText(s.cpETDVendor2);
                            dtWeekETDVendor3.SetText(s.cpETDVendor3);
                            dtWeekETDVendor4.SetText(s.cpETDVendor4);
                            dtWeekETDVendor5.SetText(s.cpETDVendor5);

                            dtWeekETDPortOld1.SetText(s.cpETDPortOld1);
                            dtWeekETDPortOld2.SetText(s.cpETDPortOld2);
                            dtWeekETDPortOld3.SetText(s.cpETDPortOld3);
                            dtWeekETDPortOld4.SetText(s.cpETDPortOld4);
                            dtWeekETDPortOld5.SetText(s.cpETDPortOld5);

                            dtWeekETDPort1.SetText(s.cpETDPort1);
                            dtWeekETDPort2.SetText(s.cpETDPort2);
                            dtWeekETDPort3.SetText(s.cpETDPort3);
                            dtWeekETDPort4.SetText(s.cpETDPort4);
                            dtWeekETDPort5.SetText(s.cpETDPort5);

                            dtWeekETAPortOld1.SetText(s.cpETAPortOld1);
                            dtWeekETAPortOld2.SetText(s.cpETAPortOld2);
                            dtWeekETAPortOld3.SetText(s.cpETAPortOld3);
                            dtWeekETAPortOld4.SetText(s.cpETAPortOld4);
                            dtWeekETAPortOld5.SetText(s.cpETAPortOld5);

                            dtWeekETAPort1.SetText(s.cpETAPort1);
                            dtWeekETAPort2.SetText(s.cpETAPort2);
                            dtWeekETAPort3.SetText(s.cpETAPort3);
                            dtWeekETAPort4.SetText(s.cpETAPort4);
                            dtWeekETAPort5.SetText(s.cpETAPort5);

                            dtETAFactWeekOld1.SetText(s.cpETAFactoryOld1);
                            dtETAFactWeekOld2.SetText(s.cpETAFactoryOld2);
                            dtETAFactWeekOld3.SetText(s.cpETAFactoryOld3);
                            dtETAFactWeekOld4.SetText(s.cpETAFactoryOld4);
                            dtETAFactWeekOld5.SetText(s.cpETAFactoryOld5);

                            dtETAFactWeek1.SetText(s.cpETAFactory1);
                            dtETAFactWeek2.SetText(s.cpETAFactory2);
                            dtETAFactWeek3.SetText(s.cpETAFactory3);
                            dtETAFactWeek4.SetText(s.cpETAFactory4);
                            dtETAFactWeek5.SetText(s.cpETAFactory5);
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
                        <dx:GridViewDataTextColumn Caption="NO." FieldName="NoUrut" HeaderStyle-HorizontalAlign="Center"
                            VisibleIndex="1" Width="30px">
                            <HeaderStyle HorizontalAlign="Center"></HeaderStyle>
                            <CellStyle Font-Names="Tahoma" Font-Size="8pt">
                            </CellStyle>
                        </dx:GridViewDataTextColumn>
                        <dx:GridViewDataTextColumn Caption="PART NO." FieldName="PartNo" HeaderStyle-HorizontalAlign="Center"
                            VisibleIndex="2" Width="90px">
                            <HeaderStyle HorizontalAlign="Center"></HeaderStyle>
                            <CellStyle Font-Names="Tahoma" Font-Size="8pt">
                            </CellStyle>
                        </dx:GridViewDataTextColumn>
                        <dx:GridViewDataTextColumn Caption="PART NAME" FieldName="PartName" HeaderStyle-HorizontalAlign="Center"
                            VisibleIndex="3" Width="180px">
                            <HeaderStyle HorizontalAlign="Center"></HeaderStyle>
                            <CellStyle Font-Names="Tahoma" Font-Size="8pt">
                            </CellStyle>
                        </dx:GridViewDataTextColumn>
                        <dx:GridViewDataTextColumn Caption="KANBAN CLS" CellStyle-HorizontalAlign="Center"
                            FieldName="KanbanCls" HeaderStyle-HorizontalAlign="Center" VisibleIndex="4" Width="0px">
                            <HeaderStyle HorizontalAlign="Center"></HeaderStyle>
                            <CellStyle Font-Names="Tahoma" Font-Size="8pt">
                            </CellStyle>
                        </dx:GridViewDataTextColumn>
                        <dx:GridViewDataTextColumn Caption="UOM" CellStyle-HorizontalAlign="Center" FieldName="UnitDesc"
                            HeaderStyle-HorizontalAlign="Center" VisibleIndex="5" Width="40px">
                            <HeaderStyle HorizontalAlign="Center"></HeaderStyle>
                            <CellStyle Font-Names="Tahoma" Font-Size="8pt">
                            </CellStyle>
                        </dx:GridViewDataTextColumn>
                        <dx:GridViewDataTextColumn Caption="MOQ" FieldName="MOQ" HeaderStyle-HorizontalAlign="Center"
                            VisibleIndex="6" Width="70px">
                            <PropertiesTextEdit DisplayFormatString="{0:n0}">
                            </PropertiesTextEdit>
                            <HeaderStyle HorizontalAlign="Center"></HeaderStyle>
                            <CellStyle Font-Names="Tahoma" Font-Size="8pt">
                            </CellStyle>
                        </dx:GridViewDataTextColumn>
                        <dx:GridViewDataTextColumn Caption="QTY/BOX" FieldName="QtyBox" HeaderStyle-HorizontalAlign="Center"
                            VisibleIndex="7" Width="70px">
                            <PropertiesTextEdit DisplayFormatString="{0:n0}">
                            </PropertiesTextEdit>
                            <HeaderStyle HorizontalAlign="Center"></HeaderStyle>
                            <CellStyle Font-Names="Tahoma" Font-Size="8pt">
                            </CellStyle>
                        </dx:GridViewDataTextColumn>
                        <dx:GridViewDataTextColumn Caption="MAKER" FieldName="Maker" HeaderStyle-HorizontalAlign="Center"
                            VisibleIndex="9" Width="0px">
                            <HeaderStyle HorizontalAlign="Center"></HeaderStyle>
                            <CellStyle Font-Names="Tahoma" Font-Size="8pt">
                            </CellStyle>
                        </dx:GridViewDataTextColumn>
                        <dx:GridViewDataTextColumn Caption=" " FieldName="AffiliateName" HeaderStyle-HorizontalAlign="Center"
                            VisibleIndex="8" Width="80px">
                            <HeaderStyle HorizontalAlign="Center"></HeaderStyle>
                            <CellStyle Font-Names="Tahoma" Font-Size="8pt">
                            </CellStyle>
                        </dx:GridViewDataTextColumn>
                        <dx:GridViewDataTextColumn FieldName="POQty" ShowInCustomizationForm="True" Width="100px"
                            Caption="TOTAL FIRM QTY" VisibleIndex="15">
                            <PropertiesTextEdit DisplayFormatString="{0:n0}">
                            </PropertiesTextEdit>
                            <HeaderStyle HorizontalAlign="Center"></HeaderStyle>
                            <CellStyle Font-Names="Tahoma" Font-Size="8pt">
                            </CellStyle>
                        </dx:GridViewDataTextColumn>
                        <dx:GridViewDataTextColumn FieldName="POQtyOld" ShowInCustomizationForm="True" Width="0px"
                            Caption="TOTAL FIRM QTY" VisibleIndex="16">
                            <PropertiesTextEdit DisplayFormatString="{0:n0}">
                            </PropertiesTextEdit>
                            <HeaderStyle HorizontalAlign="Center"></HeaderStyle>
                            <CellStyle Font-Names="Tahoma" Font-Size="8pt">
                            </CellStyle>
                        </dx:GridViewDataTextColumn>
                        <dx:GridViewDataTextColumn Caption="FORECAST N+1" FieldName="ForecastN1" HeaderStyle-HorizontalAlign="Center"
                            VisibleIndex="20" Width="80px">
                            <PropertiesTextEdit DisplayFormatString="{0:n0}">
                            </PropertiesTextEdit>
                            <HeaderStyle HorizontalAlign="Center"></HeaderStyle>
                        </dx:GridViewDataTextColumn>
                        <dx:GridViewDataTextColumn Caption="FORECAST N+2" FieldName="ForecastN2" HeaderStyle-HorizontalAlign="Center"
                            VisibleIndex="21" Width="80px">
                            <HeaderStyle HorizontalAlign="Center"></HeaderStyle>
                        </dx:GridViewDataTextColumn>
                        <dx:GridViewDataTextColumn Caption="FORECAST N+3" FieldName="ForecastN3" HeaderStyle-HorizontalAlign="Center"
                            VisibleIndex="22" Width="80px">
                            <HeaderStyle HorizontalAlign="Center"></HeaderStyle>
                        </dx:GridViewDataTextColumn>
                        <dx:GridViewDataTextColumn Caption="WEEK 5 FIRM QTY" HeaderStyle-HorizontalAlign="Center"
                            VisibleIndex="14">
                            <PropertiesTextEdit DisplayFormatString="{0:n0}">
                            </PropertiesTextEdit>
                            <HeaderStyle HorizontalAlign="Center"></HeaderStyle>
                            <CellStyle Font-Names="Tahoma" Font-Size="8pt">
                            </CellStyle>
                        </dx:GridViewDataTextColumn>
                        <dx:GridViewDataTextColumn Caption="WEEK 4 FIRM QTY" HeaderStyle-HorizontalAlign="Center"
                            VisibleIndex="13">
                            <PropertiesTextEdit DisplayFormatString="{0:n0}">
                            </PropertiesTextEdit>
                            <HeaderStyle HorizontalAlign="Center"></HeaderStyle>
                            <CellStyle Font-Names="Tahoma" Font-Size="8pt">
                            </CellStyle>
                        </dx:GridViewDataTextColumn>
                        <dx:GridViewDataTextColumn Caption="WEEK 3 FIRM QTY" CellStyle-HorizontalAlign="Center"
                            HeaderStyle-HorizontalAlign="Center" VisibleIndex="12">
                            <HeaderStyle HorizontalAlign="Center"></HeaderStyle>
                            <CellStyle Font-Names="Tahoma" Font-Size="8pt">
                            </CellStyle>
                        </dx:GridViewDataTextColumn>
                        <dx:GridViewDataTextColumn Caption="WEEK 2 FIRM QTY" HeaderStyle-HorizontalAlign="Center"
                            VisibleIndex="11">
                        </dx:GridViewDataTextColumn>
                        <dx:GridViewDataTextColumn Caption="WEEK 1 FIRM QTY" HeaderStyle-HorizontalAlign="Center"
                            VisibleIndex="10">
                        </dx:GridViewDataTextColumn>
                        <dx:GridViewDataTextColumn Caption="FORECAST N" VisibleIndex="17" HeaderStyle-HorizontalAlign="Center">
                        </dx:GridViewDataTextColumn>
                        <dx:GridViewDataTextColumn Caption="VARIANCE %" VisibleIndex="19" HeaderStyle-HorizontalAlign="Center">
                        </dx:GridViewDataTextColumn>
                        <dx:GridViewDataTextColumn Caption="VARIANCE" VisibleIndex="18" HeaderStyle-HorizontalAlign="Center">
                        </dx:GridViewDataTextColumn>
                    </Columns>
                    <SettingsBehavior AllowSelectByRowClick="True" AllowSort="False" ColumnResizeMode="Control"
                        EnableRowHotTrack="True" />
                    <SettingsPager Mode="ShowAllRecords">
                        <Summary AllPagesText="Page {0} of {1} " Text="Page {0} of {1} [{2} record(s)]" />
                        <Summary AllPagesText="Page {0} of {1} " Text="Page {0} of {1} [{2} record(s)]" />
                    </SettingsPager>
                    <SettingsEditing Mode="Batch" NewItemRowPosition="Bottom">
                        <BatchEditSettings ShowConfirmOnLosingChanges="False" />
                    </SettingsEditing>
                    <Settings ShowGroupButtons="False" ShowHorizontalScrollBar="True" ShowStatusBar="Hidden"
                        ShowVerticalScrollBar="True" VerticalScrollableHeight="190" />
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
                <dx:ASPxButton ID="btnSendSupplier" runat="server" Text="SEND TO SUPPLIER" Font-Names="Tahoma"
                    ClientInstanceName="btnSendSupplier" Width="80px" AutoPostBack="False" Font-Size="8pt">
                    <ClientSideEvents Click="function(s, e) {
                        grid.UpdateEdit();
                        var pDate = dtPeriod.GetText();
                        var pPONo = cboPONo.GetText();
                        var pAffCode= cboAffiliateCode.GetText();
                        var pSuppCode = txtSupplierCode.GetText();
                        var pComm = txtCommercial.GetText();
                        var pDelCode = cboDeliveryLocation.GetText();
                        var pShipBy = txtShipBy.GetText();
                        lblInfo.SetText('');
                        grid.PerformCallback('send' + '|' + pDate + '|' + pAffCode + '|' + pPONo + '|' + pSuppCode + '|' + pComm + '|' + pDelCode + '|' + pShipBy);
                        grid.PerformCallback('load' + '|' + pDate + '|' + pAffCode + '|' + pPONo + '|' + pSuppCode + '|' + pComm + '|' + pDelCode + '|' + pShipBy);
                    }" />
                </dx:ASPxButton>
            </td>
            <td valign="top" align="right" style="width: 50px;">
                <dx:ASPxButton ID="btnSubmit" runat="server" Text="SAVE" Font-Names="Tahoma" Width="90px"
                    ClientInstanceName="btnSubmit" AutoPostBack="False" Font-Size="8pt">
                    <ClientSideEvents Click="function(s, e) {
                        grid.UpdateEdit();
                        var pDate = dtPeriod.GetText();
                        var pPONo = cboPONo.GetText();
                        var pAffCode= cboAffiliateCode.GetText();
                        var pSuppCode = txtSupplierCode.GetText();
                        var pComm = txtCommercial.GetText();
                        var pDelCode = cboDeliveryLocation.GetText();
                        var pShipBy = txtShipBy.GetText();
                        lblInfo.SetText('');
                        grid.PerformCallback('save' + '|' + pDate + '|' + pAffCode + '|' + pPONo + '|' + pSuppCode + '|' + pComm + '|' + pDelCode + pShipBy);
                    }" />
                </dx:ASPxButton>
            </td>
        </tr>
    </table>
    <dx:ASPxGlobalEvents ID="ge" runat="server">
        <ClientSideEvents ControlsInitialized="function(s, e) {
	        OnControlsInitializedSplitter();
	        OnControlsInitializedGrid();
        }" />
    </dx:ASPxGlobalEvents>
    <dx:ASPxCallback ID="ButtonApprove" runat="server" ClientInstanceName="ButtonApprove">
        <ClientSideEvents EndCallback="function(s, e) {
            txtOrderNoWeek1.SetText(s.cpOrder1);
            txtOrderNoWeek2.SetText(s.cpOrder2);
            txtOrderNoWeek3.SetText(s.cpOrder3);
            txtOrderNoWeek4.SetText(s.cpOrder4);
            txtOrderNoWeek5.SetText(s.cpOrder5);

            dtWeekETDVendorOld1.SetText(s.cpETDVendorOld1);
            dtWeekETDVendorOld2.SetText(s.cpETDVendorOld2);
            dtWeekETDVendorOld3.SetText(s.cpETDVendorOld3);
            dtWeekETDVendorOld4.SetText(s.cpETDVendorOld4);
            dtWeekETDVendorOld5.SetText(s.cpETDVendorOld5);

            dtWeekETDVendor1.SetText(s.cpETDVendor1);
            dtWeekETDVendor2.SetText(s.cpETDVendor2);
            dtWeekETDVendor3.SetText(s.cpETDVendor3);
            dtWeekETDVendor4.SetText(s.cpETDVendor4);
            dtWeekETDVendor5.SetText(s.cpETDVendor5);

            dtWeekETDPortOld1.SetText(s.cpETDPortOld1);
            dtWeekETDPortOld2.SetText(s.cpETDPortOld2);
            dtWeekETDPortOld3.SetText(s.cpETDPortOld3);
            dtWeekETDPortOld4.SetText(s.cpETDPortOld4);
            dtWeekETDPortOld5.SetText(s.cpETDPortOld5);

            dtWeekETDPort1.SetText(s.cpETDPort1);
            dtWeekETDPort2.SetText(s.cpETDPort2);
            dtWeekETDPort3.SetText(s.cpETDPort3);
            dtWeekETDPort4.SetText(s.cpETDPort4);
            dtWeekETDPort5.SetText(s.cpETDPort5);

            dtWeekETAPortOld1.SetText(s.cpETAPortOld1);
            dtWeekETAPortOld2.SetText(s.cpETAPortOld2);
            dtWeekETAPortOld3.SetText(s.cpETAPortOld3);
            dtWeekETAPortOld4.SetText(s.cpETAPortOld4);
            dtWeekETAPortOld5.SetText(s.cpETAPortOld5);

            dtWeekETAPort1.SetText(s.cpETAPort1);
            dtWeekETAPort2.SetText(s.cpETAPort2);
            dtWeekETAPort3.SetText(s.cpETAPort3);
            dtWeekETAPort4.SetText(s.cpETAPort4);
            dtWeekETAPort5.SetText(s.cpETAPort5);

            dtETAFactWeekOld1.SetText(s.cpETAFactoryOld1);
            dtETAFactWeekOld2.SetText(s.cpETAFactoryOld2);
            dtETAFactWeekOld3.SetText(s.cpETAFactoryOld3);
            dtETAFactWeekOld4.SetText(s.cpETAFactoryOld4);
            dtETAFactWeekOld5.SetText(s.cpETAFactoryOld5);

            dtETAFactWeek1.SetText(s.cpETAFactory1);
            dtETAFactWeek2.SetText(s.cpETAFactory2);
            dtETAFactWeek3.SetText(s.cpETAFactory3);
            dtETAFactWeek4.SetText(s.cpETAFactory4);
            dtETAFactWeek5.SetText(s.cpETAFactory5);

            delete s.cpMessage;
        }" />
    </dx:ASPxCallback>
    <dx:ASPxCallback ID="cbSupplier" runat="server" ClientInstanceName="cbSupplier">
        <ClientSideEvents EndCallback="function(s, e) {
	        txtOrderNoWeek1.SetText(s.cpOrder1);
            txtOrderNoWeek2.SetText(s.cpOrder2);
            txtOrderNoWeek3.SetText(s.cpOrder3);
            txtOrderNoWeek4.SetText(s.cpOrder4);
            txtOrderNoWeek5.SetText(s.cpOrder5);

            dtWeekETDVendorOld1.SetText(s.cpETDVendorOld1);
            dtWeekETDVendorOld2.SetText(s.cpETDVendorOld2);
            dtWeekETDVendorOld3.SetText(s.cpETDVendorOld3);
            dtWeekETDVendorOld4.SetText(s.cpETDVendorOld4);
            dtWeekETDVendorOld5.SetText(s.cpETDVendorOld5);

            dtWeekETDVendor1.SetText(s.cpETDVendor1);
            dtWeekETDVendor2.SetText(s.cpETDVendor2);
            dtWeekETDVendor3.SetText(s.cpETDVendor3);
            dtWeekETDVendor4.SetText(s.cpETDVendor4);
            dtWeekETDVendor5.SetText(s.cpETDVendor5);

            dtWeekETDPortOld1.SetText(s.cpETDPortOld1);
            dtWeekETDPortOld2.SetText(s.cpETDPortOld2);
            dtWeekETDPortOld3.SetText(s.cpETDPortOld3);
            dtWeekETDPortOld4.SetText(s.cpETDPortOld4);
            dtWeekETDPortOld5.SetText(s.cpETDPortOld5);

            dtWeekETDPort1.SetText(s.cpETDPort1);
            dtWeekETDPort2.SetText(s.cpETDPort2);
            dtWeekETDPort3.SetText(s.cpETDPort3);
            dtWeekETDPort4.SetText(s.cpETDPort4);
            dtWeekETDPort5.SetText(s.cpETDPort5);

            dtWeekETAPortOld1.SetText(s.cpETAPortOld1);
            dtWeekETAPortOld2.SetText(s.cpETAPortOld2);
            dtWeekETAPortOld3.SetText(s.cpETAPortOld3);
            dtWeekETAPortOld4.SetText(s.cpETAPortOld4);
            dtWeekETAPortOld5.SetText(s.cpETAPortOld5);

            dtWeekETAPort1.SetText(s.cpETAPort1);
            dtWeekETAPort2.SetText(s.cpETAPort2);
            dtWeekETAPort3.SetText(s.cpETAPort3);
            dtWeekETAPort4.SetText(s.cpETAPort4);
            dtWeekETAPort5.SetText(s.cpETAPort5);

            dtETAFactWeekOld1.SetText(s.cpETAFactoryOld1);
            dtETAFactWeekOld2.SetText(s.cpETAFactoryOld2);
            dtETAFactWeekOld3.SetText(s.cpETAFactoryOld3);
            dtETAFactWeekOld4.SetText(s.cpETAFactoryOld4);
            dtETAFactWeekOld5.SetText(s.cpETAFactoryOld5);

            dtETAFactWeek1.SetText(s.cpETAFactory1);
            dtETAFactWeek2.SetText(s.cpETAFactory2);
            dtETAFactWeek3.SetText(s.cpETAFactory3);
            dtETAFactWeek4.SetText(s.cpETAFactory4);
            dtETAFactWeek5.SetText(s.cpETAFactory5);

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
