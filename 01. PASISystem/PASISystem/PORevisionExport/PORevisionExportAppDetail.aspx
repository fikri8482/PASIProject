<%@ Page Title="" Language="vb" AutoEventWireup="false" MasterPageFile="~/SiteContent.Master"
    CodeBehind="PORevisionExportAppDetail.aspx.vb" Inherits="PASISystem.PORevisionExportAppDetail" %>

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

            if (currentColumnName == "NoUrut" || currentColumnName == "PartNo" || currentColumnName == "PartName" 
            || currentColumnName == "Description" || currentColumnName == "MOQ" || currentColumnName == "QtyBox" 
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
        function SelectedIndexChangedDel() {
            txtDeliveryLocation.SetText(cboDeliveryLocation.GetSelectedItem().GetColumnText(1));
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
                var minggu = 5;

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
                for (i = 1; i <= minggu; i++) {
                    var day = "Week" + i;
                    var qtyHarian = s.batchEditApi.GetCellValue(e.visibleIndex, day);
                    total = total + parseInt(qtyHarian);
                    s.batchEditApi.SetCellValue(e.visibleIndex, "POQty", total);
                }

                var Qty = s.batchEditApi.GetCellValue(e.visibleIndex, "POQty");
                var qtyBox = s.batchEditApi.GetCellValue(e.visibleIndex - 1, "QtyBox");
                //s.batchEditApi.SetCellValue(e.visibleIndex, "Flag", 1);
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
                                        <dx:ASPxTimeEdit ID="dtPeriod" runat="server" ClientInstanceName="dtPeriod" DisplayFormatString="MMM yyyy"
                                            EditFormat="Custom" EditFormatString="MMM yyyy" Font-Names="Tahoma" Font-Size="8pt"
                                            Width="150px">
                                            <ReadOnlyStyle BackColor="#CCCCCC">
                                            </ReadOnlyStyle>
                                        </dx:ASPxTimeEdit>
                                    </td>
                                    <td align="left" valign="middle" width="10px">
                                        &nbsp;
                                    </td>
                                    <td align="left" valign="middle" width="150px">
                                        &nbsp;
                                    </td>
                                    <td align="left" valign="middle" width="75px">
                                        <dx:ASPxLabel ID="ASPxLabel9" runat="server" Text="DELIVERY LOCATION" Font-Names="Tahoma"
                                            Font-Size="8pt" Width="80px">
                                        </dx:ASPxLabel>
                                    </td>
                                    <td align="left" valign="middle" width="165px">
                                        &nbsp;<dx:ASPxComboBox ID="cboDeliveryLocation" runat="server" TextFormatString="{0}"
                                            DropDownStyle="DropDown" Height="20px" Width="100%" ClientInstanceName="cboDeliveryLocation"
                                            IncrementalFilteringMode="StartsWith" Font-Names="Tahoma" Font-Size="8pt" 
                                            EnableIncrementalFiltering="True">
                                            <ClientSideEvents SelectedIndexChanged="function(s, e) {SelectedIndexChangedDel();cboPONo.PerformCallback(cboDeliveryLocation.GetValue().toString());}"
                                                LostFocus="function(s, e) { lblInfo.SetText(''); }" />
                                            <ReadOnlyStyle BackColor="#CCCCCC">
                                            </ReadOnlyStyle>
                                        </dx:ASPxComboBox>
                                    </td>
                                    <td align="left" valign="middle" width="165px">
                                        &nbsp;<dx:ASPxTextBox ID="txtDeliveryLocation" runat="server" ClientInstanceName="txtDeliveryLocation"
                                            Width="100%" Height="20px" ReadOnly="True" Font-Names="Tahoma" Font-Size="8pt">
                                            <ReadOnlyStyle BackColor="#CCCCCC">
                                            </ReadOnlyStyle>
                                        </dx:ASPxTextBox>
                                    </td>
                                </tr>
                                <tr>
                                    <td style="width: 5px;">
                                    </td>
                                    <td align="left" valign="middle" width="120px">
                                        <dx:ASPxLabel ID="ASPxLabel1" runat="server" Text="AFFILIATE" Font-Names="Tahoma"
                                            Font-Size="8pt" Width="100%">
                                        </dx:ASPxLabel>
                                    </td>
                                    <td align="left" valign="middle" width="150px">
                                        <dx:ASPxComboBox ID="cboAffiliateCode" runat="server" TextFormatString="{0}" DropDownStyle="DropDown"
                                            Height="20px" Width="100%" ClientInstanceName="cboAffiliateCode" IncrementalFilteringMode="StartsWith"
                                            Font-Names="Tahoma" Font-Size="8pt" EnableIncrementalFiltering="True">
                                            <ClientSideEvents SelectedIndexChanged="function(s, e) {SelectedIndexChangedAff();cboPONo.PerformCallback(cboAffiliateCode.GetValue().toString());}"
                                                LostFocus="function(s, e) { lblInfo.SetText(&quot;&quot;); }" />
                                            <ReadOnlyStyle BackColor="#CCCCCC">
                                            </ReadOnlyStyle>
                                        </dx:ASPxComboBox>
                                    </td>
                                    <td align="left" valign="middle" colspan="2">
                                        <dx:ASPxTextBox ID="txtAffiliateName" runat="server" ClientInstanceName="txtAffiliateName"
                                            Width="100%" Height="20px" ReadOnly="True" Font-Names="Tahoma" Font-Size="8pt">
                                            <ReadOnlyStyle BackColor="#CCCCCC">
                                            </ReadOnlyStyle>
                                        </dx:ASPxTextBox>
                                    </td>
                                    <td align="left" valign="middle">
                                        <dx:ASPxLabel ID="ASPxLabel10" runat="server" Text="PO EMERGENCY" Font-Names="Tahoma"
                                            Font-Size="8pt" Width="80px">
                                        </dx:ASPxLabel>
                                    </td>
                                    <td align="left" valign="middle">
                                        <dx:ASPxTextBox ID="txtPOEmergency" runat="server" Width="150px" Height="20px" ClientInstanceName="txtPOEmergency"
                                            Font-Names="Tahoma" Font-Size="8pt" MaxLength="10" ReadOnly="True">
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
                                </tr>
                                <tr>
                                    <td style="width: 5px;">
                                    </td>
                                    <td align="left" valign="middle" width="80px">
                                        <dx:ASPxLabel ID="ASPxLabel5" runat="server" Text="PO NO." Font-Names="Tahoma" Font-Size="8pt"
                                            Width="75px">
                                        </dx:ASPxLabel>
                                    </td>
                                    <td align="left" valign="middle" width="150px">
                                        <dx:ASPxComboBox ID="cboPONo" runat="server" TextFormatString="{0}" DropDownStyle="DropDown"
                                            Height="20px" Width="100%" ClientInstanceName="cboPONo" IncrementalFilteringMode="StartsWith"
                                            Font-Names="Tahoma" Font-Size="8pt">
                                            <ClientSideEvents SelectedIndexChanged="function(s, e) {
                                                var pDate = dtPeriod.GetText();
                                                var pPONo = cboPONo.GetText();
                                                var pAffCode= cboAffiliateCode.GetText();
                                                
                                                cbPONo.PerformCallback('loadCombo' + '|' + pDate + '|' + pPONo + '|' + pAffCode);}"
                                                LostFocus="function(s, e) { lblInfo.SetText(&quot;&quot;); }" />
                                            <ReadOnlyStyle BackColor="#CCCCCC">
                                            </ReadOnlyStyle>
                                        </dx:ASPxComboBox>
                                    </td>
                                    <td align="left" valign="middle" width="80px" style="width: -160">
                                        <dx:ASPxLabel ID="ASPxLabel8" runat="server" Text="COMMERCIAL" Font-Names="Tahoma"
                                            Font-Size="8pt">
                                        </dx:ASPxLabel>
                                    </td>
                                    <td align="left" valign="middle" width="80px" style="width: -40">
                                        <dx:ASPxTextBox ID="txtCommercial" runat="server" Width="165px" Height="20px" ClientInstanceName="txtCommercial"
                                            Font-Names="Tahoma" Font-Size="8pt" MaxLength="10" ReadOnly="True">
                                            <ClientSideEvents TextChanged="function(s, e) {
                                                grid.SetFocusedRowIndex(-1);
                                                grid.PerformCallback('kosong');
	                                            lblErrMsg.SetText('');
                                            }" />
                                            <ReadOnlyStyle BackColor="#CCCCCC">
                                            </ReadOnlyStyle>
                                        </dx:ASPxTextBox>
                                    </td>
                                    <td align="left" valign="middle" width="80px">
                                        &nbsp;
                                    </td>
                                    <td align="left" valign="top" width="300px">
                                        &nbsp;
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
                                        <dx:ASPxLabel ID="ASPxLabel2" runat="server" Text="SUPPLIER CODE/NAME" Font-Names="Tahoma"
                                            Font-Size="8pt" Width="150px">
                                        </dx:ASPxLabel>
                                    </td>
                                    <td align="left" valign="middle" width="150px">
                                        <dx:ASPxTextBox ID="txtSupplierCode" runat="server" Width="100%" Height="20px" ClientInstanceName="txtSupplierCode"
                                            Font-Names="Tahoma" Font-Size="8pt" MaxLength="10" ReadOnly="True">
                                            <ClientSideEvents TextChanged="function(s, e) {
                                                grid.SetFocusedRowIndex(-1);
                                                grid.PerformCallback('kosong');
	                                            lblErrMsg.SetText('');
                                            }" />
                                            <ReadOnlyStyle BackColor="#CCCCCC">
                                            </ReadOnlyStyle>
                                        </dx:ASPxTextBox>
                                    </td>
                                    <td align="left" valign="middle" colspan="2">
                                        <dx:ASPxTextBox ID="txtSupplierName" runat="server" ClientInstanceName="txtSupplierName"
                                            Width="100%" Height="20px" ReadOnly="True" Font-Names="Tahoma" Font-Size="8pt">
                                            <ReadOnlyStyle BackColor="#CCCCCC">
                                            </ReadOnlyStyle>
                                        </dx:ASPxTextBox>
                                    </td>
                                    <td align="left" valign="middle" width="80px">
                                        &nbsp;
                                    </td>
                                    <td align="left" valign="top" width="300px">
                                        &nbsp;
                                    </td>
                                    <td align="left" valign="top" width="300px">
                                        <table>
                                            <tr>
                                                <td align="right" valign="middle">
                                                    <dx:ASPxButton ID="btnRefresh" runat="server" Text="SEARCH" Font-Names="Tahoma" Width="85px"
                                                        AutoPostBack="False" Font-Size="8pt" ClientInstanceName="btnRefresh">
                                                        <ClientSideEvents Click="function(s, e) {
                                                            var pDate = dtPeriod.GetText();
                                                            var pAffCode= cboAffiliateCode.GetText();
                                                            var pPONo = cboPONo.GetText();
                                                            var pComm = txtCommercial.GetText();
                                                            var pSuppCode = txtSupplierCode.GetText();
                                                            var pShipBy = txtShipBy.GetText();
                                                            var pDeliveryLoc = cboDeliveryLocation.GetText();
                                                            var pPOEmergency = txtPOEmergency.GetText();
                                                            
                                                            grid.PerformCallback('load' + '|' + pDate + '|' + pAffCode + '|' + pPONo + '|' + pComm + '|' + pSuppCode + '|' + pShipBy + '|' + pDeliveryLoc + '|' + pPOEmergency);
                                                            lblInfo.SetText('');
                                                        }" />
                                                    </dx:ASPxButton>
                                                </td>
                                                <td align="right" valign="middle">
                                                    <dx:ASPxButton ID="btnClear" runat="server" Text="CLEAR" Font-Names="Tahoma" Width="85px"
                                                        AutoPostBack="False" Font-Size="8pt" ClientInstanceName="btnClear">
                                                        <ClientSideEvents Click="function(s, e) {
                                                            cboAffiliateCode.SetText('');
                                                            txtAffiliateName.SetText('');
                                                            cboPONo.SetText('');
                                                            txtCommercial.SetText('');
                                                            txtSupplierCode.SetText('');
                                                            txtSupplierName.SetText('');
                                                            txtShipBy.SetText('');
                                                            cboDeliveryLocation.SetText('');
                                                            txtDeliveryLocation.SetText('');
                                                            txtPOEmergency.SetText('');
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
                                                    </dx:ASPxButton>
                                                </td>
                                            </tr>
                                        </table>
                                    </td>
                                </tr>
                                <tr>
                                    <td style="width: 5px;">
                                        &nbsp;
                                    </td>
                                    <td align="left" valign="middle" width="80px">
                                        <dx:ASPxLabel ID="ASPxLabel7" runat="server" Text="SHIP BY" Font-Names="Tahoma" Font-Size="8pt"
                                            Width="100%">
                                        </dx:ASPxLabel>
                                    </td>
                                    <td align="left" valign="middle" width="150px">
                                        <dx:ASPxTextBox ID="txtShipBy" runat="server" Width="150px" Height="20px" ClientInstanceName="txtShipBy"
                                            Font-Names="Tahoma" Font-Size="8pt" MaxLength="10" ReadOnly="True">
                                            <ClientSideEvents TextChanged="function(s, e) {
                                                grid.SetFocusedRowIndex(-1);
                                                grid.PerformCallback('kosong');
	                                            lblErrMsg.SetText('');
                                            }" />
                                            <ReadOnlyStyle BackColor="#CCCCCC">
                                            </ReadOnlyStyle>
                                        </dx:ASPxTextBox>
                                    </td>
                                    <td align="left" valign="middle" colspan="2">
                                        &nbsp;
                                    </td>
                                    <td align="left" valign="middle" width="80px">
                                        &nbsp;
                                    </td>
                                    <td align="left" valign="top" width="300px">
                                        &nbsp;
                                    </td>
                                    <td align="left" valign="top" width="300px">
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
                        <td rowspan="3">
                            <table style="width: 100%;" border="0" cellpadding="0" cellspacing="0" width="100%">
                                <tr>
                                    <td height="16px" rowspan="0" >
                                    </td>
                                    <td width="150px" height="16px" style="border-style: solid none none solid; border-width: 0.1px;
                                        border-color: #808080;">
                                        <dx:ASPxTextBox runat="server" width="150px" MaxLength="10" Height="16px" 
                                            Font-Names="Tahoma" Font-Size="8pt" ID="ASPxLabel24" BackColor="#FFD2A6" 
                                            Text="WEEK 1" ReadOnly="True" HorizontalAlign="Center">
                                            <Border BorderStyle="Groove" BorderColor="Silver" BorderWidth="1px" />
                                            <BorderLeft BorderColor="#CCCCCC" BorderStyle="Groove" BorderWidth="1px" />
                                            <BorderTop BorderColor="#CCCCCC" BorderStyle="Groove" BorderWidth="1px" />
                                            <BorderRight BorderColor="#CCCCCC" BorderStyle="Groove" BorderWidth="1px" />
                                            <BorderBottom BorderColor="#CCCCCC" BorderStyle="Groove" BorderWidth="1px" />
                                        </dx:ASPxTextBox>
                                    </td>
                                    <td width="150px" height="16px" style="border-style: solid none none solid; border-width: 0.1px;
                                        border-color: #808080;">
                                        <dx:ASPxTextBox runat="server" width="150px" MaxLength="10" Height="16px" 
                                            Font-Names="Tahoma" Font-Size="8pt" ID="ASPxTextBox1" BackColor="#FFD2A6" 
                                            Text="WEEK 2" ReadOnly="True" HorizontalAlign="Center">
                                            <Border BorderStyle="Groove" BorderColor="Silver" BorderWidth="1px" />
                                            <BorderLeft BorderColor="#CCCCCC" BorderStyle="Groove" BorderWidth="1px" />
                                            <BorderTop BorderColor="#CCCCCC" BorderStyle="Groove" BorderWidth="1px" />
                                            <BorderRight BorderColor="#CCCCCC" BorderStyle="Groove" BorderWidth="1px" />
                                            <BorderBottom BorderColor="#CCCCCC" BorderStyle="Groove" BorderWidth="1px" />
                                        </dx:ASPxTextBox>
                                    </td>
                                    <td width="150px" height="16px" style="border-style: solid none none solid; border-width: 0.1px;
                                        border-color: #808080;">
                                        <dx:ASPxTextBox runat="server" width="150px" MaxLength="10" Height="16px" 
                                            Font-Names="Tahoma" Font-Size="8pt" ID="ASPxTextBox2" BackColor="#FFD2A6" 
                                            Text="WEEK 3" ReadOnly="True" HorizontalAlign="Center">
                                            <Border BorderStyle="Groove" BorderColor="Silver" BorderWidth="1px" />
                                            <BorderLeft BorderColor="#CCCCCC" BorderStyle="Groove" BorderWidth="1px" />
                                            <BorderTop BorderColor="#CCCCCC" BorderStyle="Groove" BorderWidth="1px" />
                                            <BorderRight BorderColor="#CCCCCC" BorderStyle="Groove" BorderWidth="1px" />
                                            <BorderBottom BorderColor="#CCCCCC" BorderStyle="Groove" BorderWidth="1px" />
                                        </dx:ASPxTextBox>
                                    </td>
                                    <td width="150px" height="16px" style="border-style: solid none none solid; border-width: 0.1px;
                                        border-color: #808080;">
                                        <dx:ASPxTextBox runat="server" width="150px" MaxLength="10" Height="16px" 
                                            Font-Names="Tahoma" Font-Size="8pt" ID="ASPxTextBox3" BackColor="#FFD2A6" 
                                            Text="WEEK 4" ReadOnly="True" HorizontalAlign="Center">
                                            <Border BorderStyle="Groove" BorderColor="Silver" BorderWidth="1px" />
                                            <BorderLeft BorderColor="#CCCCCC" BorderStyle="Groove" BorderWidth="1px" />
                                            <BorderTop BorderColor="#CCCCCC" BorderStyle="Groove" BorderWidth="1px" />
                                            <BorderRight BorderColor="#CCCCCC" BorderStyle="Groove" BorderWidth="1px" />
                                            <BorderBottom BorderColor="#CCCCCC" BorderStyle="Groove" BorderWidth="1px" />
                                        </dx:ASPxTextBox>
                                    </td>
                                    <td width="150px" height="16px" style="border-style: solid none none solid; border-width: 0.1px;
                                        border-color: #808080;">
                                        <dx:ASPxTextBox runat="server" width="150px" MaxLength="10" Height="16px" 
                                            Font-Names="Tahoma" Font-Size="8pt" ID="ASPxTextBox4" BackColor="#FFD2A6" 
                                            Text="WEEK 5" ReadOnly="True" HorizontalAlign="Center">
                                            <Border BorderStyle="Groove" BorderColor="Silver" BorderWidth="1px" />
                                            <BorderLeft BorderColor="#CCCCCC" BorderStyle="Groove" BorderWidth="1px" />
                                            <BorderTop BorderColor="#CCCCCC" BorderStyle="Groove" BorderWidth="1px" />
                                            <BorderRight BorderColor="#CCCCCC" BorderStyle="Groove" BorderWidth="1px" />
                                            <BorderBottom BorderColor="#CCCCCC" BorderStyle="Groove" BorderWidth="1px" />
                                        </dx:ASPxTextBox>
                                    </td>
                                </tr>
                                <tr>
                                    <td width="150px" align="left" height="16px" style="border-style: solid none none solid; border-width: 0.1px;
                                        border-color: #808080;">
                                        <dx:ASPxTextBox runat="server" width="150px" MaxLength="10" Height="16px" 
                                            Font-Names="Tahoma" Font-Size="8pt" ID="ASPxTextBox5" BackColor="#FFD2A6" 
                                            Text="ORDER NO" ReadOnly="True">
                                            <Border BorderStyle="Groove" BorderColor="Silver" BorderWidth="1px" />
                                            <BorderLeft BorderColor="#CCCCCC" BorderStyle="Groove" BorderWidth="1px" />
                                            <BorderTop BorderColor="#CCCCCC" BorderStyle="Groove" BorderWidth="1px" />
                                            <BorderRight BorderColor="#CCCCCC" BorderStyle="Groove" BorderWidth="1px" />
                                            <BorderBottom BorderColor="#CCCCCC" BorderStyle="Groove" BorderWidth="1px" />
                                        </dx:ASPxTextBox>
                                    </td>
                                    <td width="150px" height="16px" style="border-style: solid none none solid; border-width: 0.1px;
                                        border-color: #808080;">
                                        <dx:ASPxTextBox runat="server" width="150px" MaxLength="10" Height="16px" ClientInstanceName="txtOrderNoWeek1"
                                            Font-Names="Tahoma" Font-Size="8pt" ID="txtOrderNoWeek1" BackColor="#FFFFE1">
                                            <Border BorderStyle="Groove" BorderColor="Silver" BorderWidth="1px" />
                                            <BorderLeft BorderColor="#CCCCCC" BorderStyle="Groove" BorderWidth="1px" />
                                            <BorderTop BorderColor="#CCCCCC" BorderStyle="Groove" BorderWidth="1px" />
                                            <BorderRight BorderColor="#CCCCCC" BorderStyle="Groove" BorderWidth="1px" />
                                            <BorderBottom BorderColor="#CCCCCC" BorderStyle="Groove" BorderWidth="1px" />
                                        </dx:ASPxTextBox>
                                    </td>
                                    <td width="150px" height="16px" style="border-style: solid none none solid; border-width: 0.1px;
                                        border-color: #808080;">
                                        <dx:ASPxTextBox runat="server" width="150px" MaxLength="10" Height="16px" ClientInstanceName="txtOrderNoWeek2"
                                            Font-Names="Tahoma" Font-Size="8pt" ID="txtOrderNoWeek2" BackColor="#FFFFE1">
                                            <Border BorderStyle="Groove" BorderColor="Silver" BorderWidth="1px" />
                                            <BorderLeft BorderColor="#CCCCCC" BorderStyle="Groove" BorderWidth="1px" />
                                            <BorderTop BorderColor="#CCCCCC" BorderStyle="Groove" BorderWidth="1px" />
                                            <BorderRight BorderColor="#CCCCCC" BorderStyle="Groove" BorderWidth="1px" />
                                            <BorderBottom BorderColor="#CCCCCC" BorderStyle="Groove" BorderWidth="1px" />
                                        </dx:ASPxTextBox>
                                    </td>
                                    <td width="150px" height="16px" style="border-style: solid none none solid; border-width: 0.1px;
                                        border-color: #808080;">
                                        <dx:ASPxTextBox runat="server" width="150px" MaxLength="10" Height="16px" ClientInstanceName="txtOrderNoWeek3"
                                            Font-Names="Tahoma" Font-Size="8pt" ID="txtOrderNoWeek3" BackColor="#FFFFE1">
                                            <Border BorderStyle="Groove" BorderColor="Silver" BorderWidth="1px" />
                                            <BorderLeft BorderColor="#CCCCCC" BorderStyle="Groove" BorderWidth="1px" />
                                            <BorderTop BorderColor="#CCCCCC" BorderStyle="Groove" BorderWidth="1px" />
                                            <BorderRight BorderColor="#CCCCCC" BorderStyle="Groove" BorderWidth="1px" />
                                            <BorderBottom BorderColor="#CCCCCC" BorderStyle="Groove" BorderWidth="1px" />
                                        </dx:ASPxTextBox>
                                    </td>
                                    <td width="150px" height="16px" style="border-style: solid none none solid; border-width: 0.1px;
                                        border-color: #808080;">
                                        <dx:ASPxTextBox runat="server" width="150px" MaxLength="10" Height="16px" ClientInstanceName="txtOrderNoWeek4"
                                            Font-Names="Tahoma" Font-Size="8pt" ID="txtOrderNoWeek4" BackColor="#FFFFE1">
                                            <Border BorderStyle="Groove" BorderColor="Silver" BorderWidth="1px" />
                                            <BorderLeft BorderColor="#CCCCCC" BorderStyle="Groove" BorderWidth="1px" />
                                            <BorderTop BorderColor="#CCCCCC" BorderStyle="Groove" BorderWidth="1px" />
                                            <BorderRight BorderColor="#CCCCCC" BorderStyle="Groove" BorderWidth="1px" />
                                            <BorderBottom BorderColor="#CCCCCC" BorderStyle="Groove" BorderWidth="1px" />
                                        </dx:ASPxTextBox>
                                    </td>
                                    <td width="150px" height="16px" style="border-style: solid none none solid; border-width: 0.1px;
                                        border-color: #808080;">
                                        <dx:ASPxTextBox runat="server" width="150px" MaxLength="10" Height="16px" ClientInstanceName="txtOrderNoWeek5"
                                            Font-Names="Tahoma" Font-Size="8pt" ID="txtOrderNoWeek5" BackColor="#FFFFE1">
                                            <Border BorderStyle="Groove" BorderColor="Silver" BorderWidth="1px" />
                                            <BorderLeft BorderColor="#CCCCCC" BorderStyle="Groove" BorderWidth="1px" />
                                            <BorderTop BorderColor="#CCCCCC" BorderStyle="Groove" BorderWidth="1px" />
                                            <BorderRight BorderColor="#CCCCCC" BorderStyle="Groove" BorderWidth="1px" />
                                            <BorderBottom BorderColor="#CCCCCC" BorderStyle="Groove" BorderWidth="1px" />
                                        </dx:ASPxTextBox>
                                    </td>
                                </tr>
                                <tr>
                                    <td width="150px" align="left" height="16px" style="border-style: solid none none solid; border-width: 0.1px;
                                        border-color: #808080;">
                                        <dx:ASPxTextBox runat="server" width="150px" MaxLength="10" Height="21px" 
                                            Font-Names="Tahoma" Font-Size="8pt" ID="ASPxTextBox6" BackColor="#FFD2A6" 
                                            Text="ETD VENDOR" ReadOnly="True">
                                            <Border BorderStyle="Groove" BorderColor="Silver" BorderWidth="1px" />
                                            <BorderLeft BorderColor="#CCCCCC" BorderStyle="Groove" BorderWidth="1px" />
                                            <BorderTop BorderColor="#CCCCCC" BorderStyle="Groove" BorderWidth="1px" />
                                            <BorderRight BorderColor="#CCCCCC" BorderStyle="Groove" BorderWidth="1px" />
                                            <BorderBottom BorderColor="#CCCCCC" BorderStyle="Groove" BorderWidth="1px" />
                                        </dx:ASPxTextBox>
                                    </td>
                                    <td width="150px" height="16px" style="border-style: solid none none solid; border-width: 0.1px;
                                        border-color: #808080;">
                                        <dx:ASPxTimeEdit ID="dtWeekETDVendor1" runat="server" ClientInstanceName="dtWeekETDVendor1" DisplayFormatString="dd MMM yyyy"
                                            EditFormat="Custom" EditFormatString="dd MMM yyyy" Font-Names="Tahoma" Font-Size="8pt" BackColor="#FFFFE1" 
                                            Width="150px">
                                            <Border BorderStyle="Groove" BorderColor="Silver" BorderWidth="1px" />
                                            <BorderLeft BorderColor="#CCCCCC" BorderStyle="Groove" BorderWidth="1px" />
                                            <BorderTop BorderColor="#CCCCCC" BorderStyle="Groove" BorderWidth="1px" />
                                            <BorderRight BorderColor="#CCCCCC" BorderStyle="Groove" BorderWidth="1px" />
                                            <BorderBottom BorderColor="#CCCCCC" BorderStyle="Groove" BorderWidth="1px" />
                                        </dx:ASPxTimeEdit>
                                    </td>
                                    <td width="150px" height="16px" style="border-style: solid none none solid; border-width: 0.1px;
                                        border-color: #808080;">
                                        <dx:ASPxTimeEdit ID="dtWeekETDVendor2" runat="server" 
                                            ClientInstanceName="dtWeekETDVendor2" DisplayFormatString="dd MMM yyyy"
                                            EditFormat="Custom" EditFormatString="dd MMM yyyy" Font-Names="Tahoma" 
                                            Font-Size="8pt" BackColor="#FFFFE1" 
                                            Width="150px">
                                            <Border BorderStyle="Groove" BorderColor="Silver" BorderWidth="1px" />
                                            <BorderLeft BorderColor="#CCCCCC" BorderStyle="Groove" BorderWidth="1px" />
                                            <BorderTop BorderColor="#CCCCCC" BorderStyle="Groove" BorderWidth="1px" />
                                            <BorderRight BorderColor="#CCCCCC" BorderStyle="Groove" BorderWidth="1px" />
                                            <BorderBottom BorderColor="#CCCCCC" BorderStyle="Groove" BorderWidth="1px" />
                                        </dx:ASPxTimeEdit>
                                    </td>
                                    <td width="150px" height="16px" style="border-style: solid none none solid; border-width: 0.1px;
                                        border-color: #808080;">
                                        <dx:ASPxTimeEdit ID="dtWeekETDVendor3" runat="server" 
                                            ClientInstanceName="dtWeekETDVendor3" DisplayFormatString="dd MMM yyyy"
                                            EditFormat="Custom" EditFormatString="dd MMM yyyy" Font-Names="Tahoma" 
                                            Font-Size="8pt" BackColor="#FFFFE1" 
                                            Width="150px">
                                            <Border BorderStyle="Groove" BorderColor="Silver" BorderWidth="1px" />
                                            <BorderLeft BorderColor="#CCCCCC" BorderStyle="Groove" BorderWidth="1px" />
                                            <BorderTop BorderColor="#CCCCCC" BorderStyle="Groove" BorderWidth="1px" />
                                            <BorderRight BorderColor="#CCCCCC" BorderStyle="Groove" BorderWidth="1px" />
                                            <BorderBottom BorderColor="#CCCCCC" BorderStyle="Groove" BorderWidth="1px" />
                                        </dx:ASPxTimeEdit>
                                    </td>
                                    <td width="150px" height="16px" style="border-style: solid none none solid; border-width: 0.1px;
                                        border-color: #808080;">
                                        <dx:ASPxTimeEdit ID="dtWeekETDVendor4" runat="server" 
                                            ClientInstanceName="dtWeekETDVendor4" DisplayFormatString="dd MMM yyyy"
                                            EditFormat="Custom" EditFormatString="dd MMM yyyy" Font-Names="Tahoma" 
                                            Font-Size="8pt" BackColor="#FFFFE1" 
                                            Width="150px">
                                            <Border BorderStyle="Groove" BorderColor="Silver" BorderWidth="1px" />
                                            <BorderLeft BorderColor="#CCCCCC" BorderStyle="Groove" BorderWidth="1px" />
                                            <BorderTop BorderColor="#CCCCCC" BorderStyle="Groove" BorderWidth="1px" />
                                            <BorderRight BorderColor="#CCCCCC" BorderStyle="Groove" BorderWidth="1px" />
                                            <BorderBottom BorderColor="#CCCCCC" BorderStyle="Groove" BorderWidth="1px" />
                                        </dx:ASPxTimeEdit>
                                    </td>
                                    <td width="150px" height="16px" style="border-style: solid none none solid; border-width: 0.1px;
                                        border-color: #808080;">
                                        <dx:ASPxTimeEdit ID="dtWeekETDVendor5" runat="server" 
                                            ClientInstanceName="dtWeekETDVendor5" DisplayFormatString="dd MMM yyyy"
                                            EditFormat="Custom" EditFormatString="dd MMM yyyy" Font-Names="Tahoma" 
                                            Font-Size="8pt" BackColor="#FFFFE1" 
                                            Width="150px">
                                            <Border BorderStyle="Groove" BorderColor="Silver" BorderWidth="1px" />
                                            <BorderLeft BorderColor="#CCCCCC" BorderStyle="Groove" BorderWidth="1px" />
                                            <BorderTop BorderColor="#CCCCCC" BorderStyle="Groove" BorderWidth="1px" />
                                            <BorderRight BorderColor="#CCCCCC" BorderStyle="Groove" BorderWidth="1px" />
                                            <BorderBottom BorderColor="#CCCCCC" BorderStyle="Groove" BorderWidth="1px" />
                                        </dx:ASPxTimeEdit>
                                    </td>
                                </tr>
                                <tr>
                                    <td width="150px" align="left" height="16px" style="border-style: solid none none solid; border-width: 0.1px;
                                        border-color: #808080;">
                                        <dx:ASPxTextBox runat="server" width="150px" MaxLength="10" Height="21px" 
                                            Font-Names="Tahoma" Font-Size="8pt" ID="ASPxTextBox7" BackColor="#FFD2A6" 
                                            Text="ETD PORT" ReadOnly="True">
                                            <Border BorderStyle="Groove" BorderColor="Silver" BorderWidth="1px" />
                                            <BorderLeft BorderColor="#CCCCCC" BorderStyle="Groove" BorderWidth="1px" />
                                            <BorderTop BorderColor="#CCCCCC" BorderStyle="Groove" BorderWidth="1px" />
                                            <BorderRight BorderColor="#CCCCCC" BorderStyle="Groove" BorderWidth="1px" />
                                            <BorderBottom BorderColor="#CCCCCC" BorderStyle="Groove" BorderWidth="1px" />
                                        </dx:ASPxTextBox>
                                    </td>
                                    <td width="150px" align="left" height="16px" style="border-style: solid none none solid; border-width: 0.1px;
                                        border-color: #808080;">
                                        <dx:ASPxTimeEdit ID="dtWeekETDPort1" runat="server" 
                                            ClientInstanceName="dtWeekETDPort1" DisplayFormatString="dd MMM yyyy"
                                            EditFormat="Custom" EditFormatString="dd MMM yyyy" Font-Names="Tahoma" 
                                            Font-Size="8pt" BackColor="#FFFFE1" 
                                            Width="150px">
                                            <Border BorderStyle="Groove" BorderColor="Silver" BorderWidth="1px" />
                                            <BorderLeft BorderColor="#CCCCCC" BorderStyle="Groove" BorderWidth="1px" />
                                            <BorderTop BorderColor="#CCCCCC" BorderStyle="Groove" BorderWidth="1px" />
                                            <BorderRight BorderColor="#CCCCCC" BorderStyle="Groove" BorderWidth="1px" />
                                            <BorderBottom BorderColor="#CCCCCC" BorderStyle="Groove" BorderWidth="1px" />
                                        </dx:ASPxTimeEdit>
                                    </td>
                                    <td width="150px" align="left" height="16px" style="border-style: solid none none solid; border-width: 0.1px;
                                        border-color: #808080;">
                                        <dx:ASPxTimeEdit ID="dtWeekETDPort2" runat="server" 
                                            ClientInstanceName="dtWeekETDPort2" DisplayFormatString="dd MMM yyyy"
                                            EditFormat="Custom" EditFormatString="dd MMM yyyy" Font-Names="Tahoma" 
                                            Font-Size="8pt" BackColor="#FFFFE1" 
                                            Width="150px">
                                            <Border BorderStyle="Groove" BorderColor="Silver" BorderWidth="1px" />
                                            <BorderLeft BorderColor="#CCCCCC" BorderStyle="Groove" BorderWidth="1px" />
                                            <BorderTop BorderColor="#CCCCCC" BorderStyle="Groove" BorderWidth="1px" />
                                            <BorderRight BorderColor="#CCCCCC" BorderStyle="Groove" BorderWidth="1px" />
                                            <BorderBottom BorderColor="#CCCCCC" BorderStyle="Groove" BorderWidth="1px" />
                                        </dx:ASPxTimeEdit>
                                    </td>
                                    <td width="150px" align="left" height="16px" style="border-style: solid none none solid; border-width: 0.1px;
                                        border-color: #808080;">
                                        <dx:ASPxTimeEdit ID="dtWeekETDPort3" runat="server" 
                                            ClientInstanceName="dtWeekETDPort3" DisplayFormatString="dd MMM yyyy"
                                            EditFormat="Custom" EditFormatString="dd MMM yyyy" Font-Names="Tahoma" 
                                            Font-Size="8pt" BackColor="#FFFFE1" 
                                            Width="150px">
                                            <Border BorderStyle="Groove" BorderColor="Silver" BorderWidth="1px" />
                                            <BorderLeft BorderColor="#CCCCCC" BorderStyle="Groove" BorderWidth="1px" />
                                            <BorderTop BorderColor="#CCCCCC" BorderStyle="Groove" BorderWidth="1px" />
                                            <BorderRight BorderColor="#CCCCCC" BorderStyle="Groove" BorderWidth="1px" />
                                            <BorderBottom BorderColor="#CCCCCC" BorderStyle="Groove" BorderWidth="1px" />
                                        </dx:ASPxTimeEdit>
                                    </td>
                                    <td width="150px" align="left" height="16px" style="border-style: solid none none solid; border-width: 0.1px;
                                        border-color: #808080;">
                                        <dx:ASPxTimeEdit ID="dtWeekETDPort4" runat="server" 
                                            ClientInstanceName="dtWeekETDPort4" DisplayFormatString="dd MMM yyyy"
                                            EditFormat="Custom" EditFormatString="dd MMM yyyy" Font-Names="Tahoma" 
                                            Font-Size="8pt" BackColor="#FFFFE1" 
                                            Width="150px">
                                            <Border BorderStyle="Groove" BorderColor="Silver" BorderWidth="1px" />
                                            <BorderLeft BorderColor="#CCCCCC" BorderStyle="Groove" BorderWidth="1px" />
                                            <BorderTop BorderColor="#CCCCCC" BorderStyle="Groove" BorderWidth="1px" />
                                            <BorderRight BorderColor="#CCCCCC" BorderStyle="Groove" BorderWidth="1px" />
                                            <BorderBottom BorderColor="#CCCCCC" BorderStyle="Groove" BorderWidth="1px" />
                                        </dx:ASPxTimeEdit>
                                    </td>
                                    <td width="150px"align="left" height="16px" style="border-style: solid none none solid; border-width: 0.1px;
                                        border-color: #808080;">
                                        <dx:ASPxTimeEdit ID="dtWeekETDPort5" runat="server" 
                                            ClientInstanceName="dtWeekETDPort5" DisplayFormatString="dd MMM yyyy"
                                            EditFormat="Custom" EditFormatString="dd MMM yyyy" Font-Names="Tahoma" 
                                            Font-Size="8pt" BackColor="#FFFFE1" 
                                            Width="150px">
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
                                        <dx:ASPxTextBox runat="server" width="150px" MaxLength="10" Height="21px" 
                                            Font-Names="Tahoma" Font-Size="8pt" ID="ASPxTextBox8" BackColor="#FFD2A6" 
                                            Text="ETA PORT" ReadOnly="True">
                                            <Border BorderStyle="Groove" BorderColor="Silver" BorderWidth="1px" />
                                            <BorderLeft BorderColor="#CCCCCC" BorderStyle="Groove" BorderWidth="1px" />
                                            <BorderTop BorderColor="#CCCCCC" BorderStyle="Groove" BorderWidth="1px" />
                                            <BorderRight BorderColor="#CCCCCC" BorderStyle="Groove" BorderWidth="1px" />
                                            <BorderBottom BorderColor="#CCCCCC" BorderStyle="Groove" BorderWidth="1px" />
                                        </dx:ASPxTextBox>
                                    </td>
                                    <td width="150px" align="left" style="border-style: solid none none solid; border-width: 0.1px;
                                        border-color: #808080;" class="style1">
                                        <dx:ASPxTimeEdit ID="dtETAPortweek1" runat="server" 
                                            ClientInstanceName="dtETAPortweek1" DisplayFormatString="dd MMM yyyy"
                                            EditFormat="Custom" EditFormatString="dd MMM yyyy" Font-Names="Tahoma" 
                                            Font-Size="8pt" BackColor="#FFFFE1" 
                                            Width="150px">
                                            <Border BorderStyle="Groove" BorderColor="Silver" BorderWidth="1px" />
                                            <BorderLeft BorderColor="#CCCCCC" BorderStyle="Groove" BorderWidth="1px" />
                                            <BorderTop BorderColor="#CCCCCC" BorderStyle="Groove" BorderWidth="1px" />
                                            <BorderRight BorderColor="#CCCCCC" BorderStyle="Groove" BorderWidth="1px" />
                                            <BorderBottom BorderColor="#CCCCCC" BorderStyle="Groove" BorderWidth="1px" />
                                        </dx:ASPxTimeEdit>
                                    </td>
                                    <td width="150px" align="left" style="border-style: solid none none solid; border-width: 0.1px;
                                        border-color: #808080;" class="style1">
                                        <dx:ASPxTimeEdit ID="dtETAPortweek2" runat="server" 
                                            ClientInstanceName="dtETAPortweek2" DisplayFormatString="dd MMM yyyy"
                                            EditFormat="Custom" EditFormatString="dd MMM yyyy" Font-Names="Tahoma" 
                                            Font-Size="8pt" BackColor="#FFFFE1" 
                                            Width="150px">
                                            <Border BorderStyle="Groove" BorderColor="Silver" BorderWidth="1px" />
                                            <BorderLeft BorderColor="#CCCCCC" BorderStyle="Groove" BorderWidth="1px" />
                                            <BorderTop BorderColor="#CCCCCC" BorderStyle="Groove" BorderWidth="1px" />
                                            <BorderRight BorderColor="#CCCCCC" BorderStyle="Groove" BorderWidth="1px" />
                                            <BorderBottom BorderColor="#CCCCCC" BorderStyle="Groove" BorderWidth="1px" />
                                        </dx:ASPxTimeEdit>
                                    </td>
                                    <td width="150px" align="left" style="border-style: solid none none solid; border-width: 0.1px;
                                        border-color: #808080;" class="style1">
                                        <dx:ASPxTimeEdit ID="dtETAPortweek3" runat="server" 
                                            ClientInstanceName="dtETAPortweek3" DisplayFormatString="dd MMM yyyy"
                                            EditFormat="Custom" EditFormatString="dd MMM yyyy" Font-Names="Tahoma" 
                                            Font-Size="8pt" BackColor="#FFFFE1" 
                                            Width="150px">
                                            <Border BorderStyle="Groove" BorderColor="Silver" BorderWidth="1px" />
                                            <BorderLeft BorderColor="#CCCCCC" BorderStyle="Groove" BorderWidth="1px" />
                                            <BorderTop BorderColor="#CCCCCC" BorderStyle="Groove" BorderWidth="1px" />
                                            <BorderRight BorderColor="#CCCCCC" BorderStyle="Groove" BorderWidth="1px" />
                                            <BorderBottom BorderColor="#CCCCCC" BorderStyle="Groove" BorderWidth="1px" />
                                        </dx:ASPxTimeEdit>
                                    </td>
                                    <td width="150px" align="left" style="border-style: solid none none solid; border-width: 0.1px;
                                        border-color: #808080;" class="style1">
                                        <dx:ASPxTimeEdit ID="dtETAPortweek4" runat="server" 
                                            ClientInstanceName="dtETAPortweek4" DisplayFormatString="dd MMM yyyy"
                                            EditFormat="Custom" EditFormatString="dd MMM yyyy" Font-Names="Tahoma" 
                                            Font-Size="8pt" BackColor="#FFFFE1" 
                                            Width="150px">
                                            <Border BorderStyle="Groove" BorderColor="Silver" BorderWidth="1px" />
                                            <BorderLeft BorderColor="#CCCCCC" BorderStyle="Groove" BorderWidth="1px" />
                                            <BorderTop BorderColor="#CCCCCC" BorderStyle="Groove" BorderWidth="1px" />
                                            <BorderRight BorderColor="#CCCCCC" BorderStyle="Groove" BorderWidth="1px" />
                                            <BorderBottom BorderColor="#CCCCCC" BorderStyle="Groove" BorderWidth="1px" />
                                        </dx:ASPxTimeEdit>
                                    </td>
                                    <td width="150px" align="left" style="border-style: solid none none solid; border-width: 0.1px;
                                        border-color: #808080;" class="style1">
                                        <dx:ASPxTimeEdit ID="dtETAPortweek5" runat="server" 
                                            ClientInstanceName="dtETAPortweek5" DisplayFormatString="dd MMM yyyy"
                                            EditFormat="Custom" EditFormatString="dd MMM yyyy" Font-Names="Tahoma" 
                                            Font-Size="8pt" BackColor="#FFFFE1" 
                                            Width="150px">
                                            <Border BorderStyle="Groove" BorderColor="Silver" BorderWidth="1px" />
                                            <BorderLeft BorderColor="#CCCCCC" BorderStyle="Groove" BorderWidth="1px" />
                                            <BorderTop BorderColor="#CCCCCC" BorderStyle="Groove" BorderWidth="1px" />
                                            <BorderRight BorderColor="#CCCCCC" BorderStyle="Groove" BorderWidth="1px" />
                                            <BorderBottom BorderColor="#CCCCCC" BorderStyle="Groove" BorderWidth="1px" />
                                        </dx:ASPxTimeEdit>
                                    </td>
                                </tr>
                                <tr>
                                    <td width="150px" align="left" height="16px" style="border-style: solid none none solid; border-width: 0.1px;
                                        border-color: #808080;">
                                        <dx:ASPxTextBox runat="server" width="150px" MaxLength="10" Height="21px" 
                                            Font-Names="Tahoma" Font-Size="8pt" ID="ASPxTextBox9" BackColor="#FFD2A6" 
                                            Text="ETD FACTORY" ReadOnly="True">
                                            <Border BorderStyle="Groove" BorderColor="Silver" BorderWidth="1px" />
                                            <BorderLeft BorderColor="#CCCCCC" BorderStyle="Groove" BorderWidth="1px" />
                                            <BorderTop BorderColor="#CCCCCC" BorderStyle="Groove" BorderWidth="1px" />
                                            <BorderRight BorderColor="#CCCCCC" BorderStyle="Groove" BorderWidth="1px" />
                                            <BorderBottom BorderColor="#CCCCCC" BorderStyle="Groove" BorderWidth="1px" />
                                        </dx:ASPxTextBox>
                                    </td>
                                    <td width="150px" align="left" height="16px" style="border-style: solid none none solid; border-width: 0.1px;
                                        border-color: #808080;" >
                                        <dx:ASPxTimeEdit ID="dtETAFactWeek1" runat="server" 
                                            ClientInstanceName="dtETAFactWeek1" DisplayFormatString="dd MMM yyyy"
                                            EditFormat="Custom" EditFormatString="dd MMM yyyy" Font-Names="Tahoma" 
                                            Font-Size="8pt" BackColor="#FFFFE1" 
                                            Width="150px">
                                            <Border BorderStyle="Groove" BorderColor="Silver" BorderWidth="1px" />
                                            <BorderLeft BorderColor="#CCCCCC" BorderStyle="Groove" BorderWidth="1px" />
                                            <BorderTop BorderColor="#CCCCCC" BorderStyle="Groove" BorderWidth="1px" />
                                            <BorderRight BorderColor="#CCCCCC" BorderStyle="Groove" BorderWidth="1px" />
                                            <BorderBottom BorderColor="#CCCCCC" BorderStyle="Groove" BorderWidth="1px" />
                                        </dx:ASPxTimeEdit>
                                    </td>
                                    <td width="150px" align="left" height="16px" style="border-style: solid none none solid; border-width: 0.1px;
                                        border-color: #808080;">
                                        <dx:ASPxTimeEdit ID="dtETAFactWeek2" runat="server" 
                                            ClientInstanceName="dtETAFactWeek1" DisplayFormatString="dd MMM yyyy"
                                            EditFormat="Custom" EditFormatString="dd MMM yyyy" Font-Names="Tahoma" 
                                            Font-Size="8pt" BackColor="#FFFFE1" 
                                            Width="150px">
                                            <Border BorderStyle="Groove" BorderColor="Silver" BorderWidth="1px" />
                                            <BorderLeft BorderColor="#CCCCCC" BorderStyle="Groove" BorderWidth="1px" />
                                            <BorderTop BorderColor="#CCCCCC" BorderStyle="Groove" BorderWidth="1px" />
                                            <BorderRight BorderColor="#CCCCCC" BorderStyle="Groove" BorderWidth="1px" />
                                            <BorderBottom BorderColor="#CCCCCC" BorderStyle="Groove" BorderWidth="1px" />
                                        </dx:ASPxTimeEdit>
                                    </td>
                                    <td width="150px" align="left" height="16px" style="border-style: solid none none solid; border-width: 0.1px;
                                        border-color: #808080;">
                                        <dx:ASPxTimeEdit ID="dtETAFactWeek3" runat="server" 
                                            ClientInstanceName="dtETAFactWeek1" DisplayFormatString="dd MMM yyyy"
                                            EditFormat="Custom" EditFormatString="dd MMM yyyy" Font-Names="Tahoma" 
                                            Font-Size="8pt" BackColor="#FFFFE1" 
                                            Width="150px">
                                            <Border BorderStyle="Groove" BorderColor="Silver" BorderWidth="1px" />
                                            <BorderLeft BorderColor="#CCCCCC" BorderStyle="Groove" BorderWidth="1px" />
                                            <BorderTop BorderColor="#CCCCCC" BorderStyle="Groove" BorderWidth="1px" />
                                            <BorderRight BorderColor="#CCCCCC" BorderStyle="Groove" BorderWidth="1px" />
                                            <BorderBottom BorderColor="#CCCCCC" BorderStyle="Groove" BorderWidth="1px" />
                                        </dx:ASPxTimeEdit>
                                    </td>
                                    <td width="150px" align="left" height="16px" style="border-style: solid none none solid; border-width: 0.1px;
                                        border-color: #808080;">
                                        <dx:ASPxTimeEdit ID="dtETAFactWeek4" runat="server" 
                                            ClientInstanceName="dtETAFactWeek1" DisplayFormatString="dd MMM yyyy"
                                            EditFormat="Custom" EditFormatString="dd MMM yyyy" Font-Names="Tahoma" 
                                            Font-Size="8pt" BackColor="#FFFFE1" 
                                            Width="150px">
                                            <Border BorderStyle="Groove" BorderColor="Silver" BorderWidth="1px" />
                                            <BorderLeft BorderColor="#CCCCCC" BorderStyle="Groove" BorderWidth="1px" />
                                            <BorderTop BorderColor="#CCCCCC" BorderStyle="Groove" BorderWidth="1px" />
                                            <BorderRight BorderColor="#CCCCCC" BorderStyle="Groove" BorderWidth="1px" />
                                            <BorderBottom BorderColor="#CCCCCC" BorderStyle="Groove" BorderWidth="1px" />
                                        </dx:ASPxTimeEdit>
                                    </td>
                                    <td width="150px" align="left" height="16px" style="border-style: solid none none solid; border-width: 0.1px;
                                        border-color: #808080;">
                                        <dx:ASPxTimeEdit ID="dtETAFactWeek5" runat="server" 
                                            ClientInstanceName="dtETAFactWeek1" DisplayFormatString="dd MMM yyyy"
                                            EditFormat="Custom" EditFormatString="dd MMM yyyy" Font-Names="Tahoma" 
                                            Font-Size="8pt" BackColor="#FFFFE1" 
                                            Width="150px">
                                            <Border BorderStyle="Groove" BorderColor="Silver" BorderWidth="1px" />
                                            <BorderLeft BorderColor="#CCCCCC" BorderStyle="Groove" BorderWidth="1px" />
                                            <BorderTop BorderColor="#CCCCCC" BorderStyle="Groove" BorderWidth="1px" />
                                            <BorderRight BorderColor="#CCCCCC" BorderStyle="Groove" BorderWidth="1px" />
                                            <BorderBottom BorderColor="#CCCCCC" BorderStyle="Groove" BorderWidth="1px" />
                                        </dx:ASPxTimeEdit>
                                    </td>
                                </tr>
                                <tr>
                                    <td width="150px" align="left" height="16px" style="border-style: solid none none solid; border-width: 0.1px;
                                        border-color: #808080;">
                                        &nbsp;</td>
                                    <td width="150px" align="left" height="16px" style="border-style: solid none none solid; border-width: 0.1px;
                                        border-color: #808080;" >
                                        &nbsp;</td>
                                    <td width="150px" align="left" height="16px" style="border-style: solid none none solid; border-width: 0.1px;
                                        border-color: #808080;">
                                        &nbsp;</td>
                                    <td width="150px" align="left" height="16px" style="border-style: solid none none solid; border-width: 0.1px;
                                        border-color: #808080;">
                                        &nbsp;</td>
                                    <td width="150px" align="left" height="16px" style="border-style: solid none none solid; border-width: 0.1px;
                                        border-color: #808080;">
                                        &nbsp;</td>
                                    <td width="150px" align="left" height="16px" style="border-style: solid none none solid; border-width: 0.1px;
                                        border-color: #808080;">
                                        &nbsp;</td>
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
                </table>
            </td>
        </tr>
        <div style="height: 1px;">
        </div>
        <tr>
            <td colspan="8" align="left" valign="top" height="220">
                <%--Column : Grid--%>
                <dx:ASPxGridView ID="grid" runat="server" Width="100%" Font-Names="Tahoma" KeyFieldName="Sort;PartNo"
                    AutoGenerateColumns="False" ClientInstanceName="grid" Font-Size="8pt">
                    <ClientSideEvents Init="OnInit" FocusedRowChanged="function(s, e) {OnGridFocusedRowChanged();}"
                        EndCallback="function(s, e) {
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

                        if (txtSuppAppDate.GetText() != '' || txtSuppPendDate.GetText() != '' || txtSuppUnpDate.GetText() != '' || txtPASIAppDate.GetText() != '' || txtAffFinalAppDate.GetText() != '') {
                            btnSubmit.SetEnabled(false);
                            btnSendSupplier.SetEnabled(false);
                        }
                                                
                        var pMsg = s.cpMessage;
                       
                        if (pMsg != '') {
                            if (pMsg.substring(1,5) == '1001' || pMsg.substring(1,5) == '1002' || pMsg.substring(1,5) == '1003' || pMsg.substring(1,5) == '2001'|| pMsg.substring(1,5) == '1008') {
                                lblInfo.GetMainElement().style.color = 'Blue';
                            } else {
                                lblInfo.GetMainElement().style.color = 'Red';
                            }
                            lblInfo.SetText(pMsg);
                        } else {
                            lblInfo.SetText('');
                        }
                        delete s.cpMessage;
                        grid.CancelEdit();
                        
                        }" RowClick="function(s, e) {
	                    lblInfo.SetText('');}" BatchEditStartEditing="OnBatchEditStartEditing" BatchEditEndEditing="OnBatchEditEndEditing"
                        CallbackError="function(s, e) {e.handled=true;}" />
                    <Columns>
                        <dx:GridViewDataTextColumn VisibleIndex="0" Caption="NO" FieldName="NoUrut" Width="30px"
                            HeaderStyle-HorizontalAlign="Center">
                            <HeaderStyle HorizontalAlign="Center"></HeaderStyle>
                        </dx:GridViewDataTextColumn>
                        <dx:GridViewDataTextColumn VisibleIndex="1" Caption="PARTNO" FieldName="PartNo" Width="100px"
                            HeaderStyle-HorizontalAlign="Center">
                            <PropertiesTextEdit DisplayFormatString="MMM yyyy">
                            </PropertiesTextEdit>
                            <HeaderStyle HorizontalAlign="Center"></HeaderStyle>
                        </dx:GridViewDataTextColumn>
                        <dx:GridViewDataTextColumn Caption="PartNos" FieldName="PartNos" VisibleIndex="2"
                            Width="0px">
                        </dx:GridViewDataTextColumn>
                        <dx:GridViewDataTextColumn VisibleIndex="3" Caption="PARTNAME" FieldName="PartName"
                            Width="180px" HeaderStyle-HorizontalAlign="Center">
                            <HeaderStyle HorizontalAlign="Center"></HeaderStyle>
                        </dx:GridViewDataTextColumn>
                        <dx:GridViewDataTextColumn VisibleIndex="4" Caption="UOM" FieldName="Description"
                            Width="50px" HeaderStyle-HorizontalAlign="Center">
                            <HeaderStyle HorizontalAlign="Center"></HeaderStyle>
                        </dx:GridViewDataTextColumn>
                        <dx:GridViewDataTextColumn VisibleIndex="5" Caption="MOQ" FieldName="MOQ"
                            Width="50px" HeaderStyle-HorizontalAlign="Center">
                            <HeaderStyle HorizontalAlign="Center"></HeaderStyle>
                        </dx:GridViewDataTextColumn>
                        <dx:GridViewDataTextColumn VisibleIndex="6" Caption="QTY/BOX" FieldName="QtyBox"
                            Width="70px" HeaderStyle-HorizontalAlign="Center">
                            <PropertiesTextEdit>
                                <MaskSettings Mask="<0..9999999999999g>" IncludeLiterals="DecimalSymbol" />
                            </PropertiesTextEdit>
                            <HeaderStyle HorizontalAlign="Center"></HeaderStyle>
                            <CellStyle HorizontalAlign="Right">
                            </CellStyle>
                        </dx:GridViewDataTextColumn>
                        <dx:GridViewDataTextColumn Caption="WEEK 1 FIRM QTY" FieldName="Week1" VisibleIndex="8"
                            Width="70px" HeaderStyle-HorizontalAlign="Center">
                            <PropertiesTextEdit>
                                <MaskSettings Mask="<0..9999999999999g>" IncludeLiterals="DecimalSymbol" />
                            </PropertiesTextEdit>
                            <HeaderStyle HorizontalAlign="Center"></HeaderStyle>
                            <CellStyle HorizontalAlign="Right">
                            </CellStyle>
                        </dx:GridViewDataTextColumn>
                        <dx:GridViewDataTextColumn Caption="WEEK 2 FIRM QTY" VisibleIndex="9" FieldName="Week2"
                            Width="70px" HeaderStyle-HorizontalAlign="Center">
                            <PropertiesTextEdit>
                                <MaskSettings Mask="<0..9999999999999g>" IncludeLiterals="DecimalSymbol" />
                            </PropertiesTextEdit>
                            <HeaderStyle HorizontalAlign="Center"></HeaderStyle>
                            <CellStyle HorizontalAlign="Right">
                            </CellStyle>
                        </dx:GridViewDataTextColumn>    
                        <dx:GridViewDataTextColumn VisibleIndex="10" Caption="WEEK 3 FIRM QTY" FieldName="Week3" 
                            Width="70px" HeaderStyle-HorizontalAlign="Center">
                            <PropertiesTextEdit>
                                <MaskSettings Mask="<0..9999999999999g>" IncludeLiterals="DecimalSymbol" />
                            </PropertiesTextEdit>
                            <HeaderStyle HorizontalAlign="Center"></HeaderStyle>
                            <CellStyle HorizontalAlign="Right">
                            </CellStyle>
                        </dx:GridViewDataTextColumn>        
                        <dx:GridViewDataTextColumn VisibleIndex="11" Caption="WEEK 4 FIRM QTY" FieldName="Week4" 
                            Width="70px" HeaderStyle-HorizontalAlign="Center">
                            <PropertiesTextEdit>
                                <MaskSettings Mask="<0..9999999999999g>" IncludeLiterals="DecimalSymbol" />
                            </PropertiesTextEdit>
                            <HeaderStyle HorizontalAlign="Center"></HeaderStyle>
                            <CellStyle HorizontalAlign="Right">
                            </CellStyle>
                        </dx:GridViewDataTextColumn>
                        <dx:GridViewDataTextColumn VisibleIndex="12" Caption="WEEK 5 FIRM QTY" FieldName="Week5" 
                            Width="70px" HeaderStyle-HorizontalAlign="Center">
                            <PropertiesTextEdit>
                                <MaskSettings Mask="<0..9999999999999g>" IncludeLiterals="DecimalSymbol" />
                            </PropertiesTextEdit>
                            <HeaderStyle HorizontalAlign="Center"></HeaderStyle>
                            <CellStyle HorizontalAlign="Right">
                            </CellStyle>
                        </dx:GridViewDataTextColumn>
                        <dx:GridViewDataTextColumn FieldName="POQty" ShowInCustomizationForm="True" Width="80px"
                            Caption="TOTAL FIRM QTY" VisibleIndex="13">
                            <PropertiesTextEdit DisplayFormatString="{0:n0}">
                                <MaskSettings Mask="<0..9999999999999g>" IncludeLiterals="DecimalSymbol"></MaskSettings>
                                <ValidationSettings ErrorDisplayMode="None">
                                </ValidationSettings>
                            </PropertiesTextEdit>
                            <HeaderStyle HorizontalAlign="Center"></HeaderStyle>
                        </dx:GridViewDataTextColumn>
                        <dx:GridViewDataTextColumn Caption="FORECAST N" VisibleIndex="14" FieldName="ForecastN" 
                            Width="70px" HeaderStyle-HorizontalAlign="Center">
                            <PropertiesTextEdit DisplayFormatString="{0:n0}">
                                <MaskSettings Mask="<0..9999999999999g>" IncludeLiterals="DecimalSymbol"></MaskSettings>
                                <ValidationSettings ErrorDisplayMode="None">
                                </ValidationSettings>
                            </PropertiesTextEdit>
                            <HeaderStyle HorizontalAlign="Center"></HeaderStyle>
                        </dx:GridViewDataTextColumn>
                        <dx:GridViewDataTextColumn VisibleIndex="15" Caption="VARIANCE" FieldName="Variance" 
                            Width="70px" HeaderStyle-HorizontalAlign="Center">
                            <PropertiesTextEdit DisplayFormatString="{0:n0}">
                                <MaskSettings Mask="<0..9999999999999g>" IncludeLiterals="DecimalSymbol"></MaskSettings>
                                <ValidationSettings ErrorDisplayMode="None">
                                </ValidationSettings>
                            </PropertiesTextEdit>
                            <HeaderStyle HorizontalAlign="Center"></HeaderStyle>
                        </dx:GridViewDataTextColumn>
                        <dx:GridViewDataTextColumn VisibleIndex="16" Caption="VARIANCE (%)" FieldName="VariancePersen" 
                            Width="70px" HeaderStyle-HorizontalAlign="Center">
                            <PropertiesTextEdit DisplayFormatString="{0:n0}">
                                <MaskSettings Mask="<0..9999999999999g>" IncludeLiterals="DecimalSymbol"></MaskSettings>
                                <ValidationSettings ErrorDisplayMode="None">
                                </ValidationSettings>
                            </PropertiesTextEdit>
                            <HeaderStyle HorizontalAlign="Center"></HeaderStyle>
                        </dx:GridViewDataTextColumn>
                        <dx:GridViewDataTextColumn FieldName="ForecastN1" ShowInCustomizationForm="True"
                            Width="80px" Caption="FORECAST N+1" VisibleIndex="21">
                            <PropertiesTextEdit DisplayFormatString="{0:n0}">
                                <MaskSettings Mask="&lt;0..9999999999999g&gt;" IncludeLiterals="DecimalSymbol"></MaskSettings>
                            </PropertiesTextEdit>
                            <HeaderStyle HorizontalAlign="Center"></HeaderStyle>
                        </dx:GridViewDataTextColumn>
                        <dx:GridViewDataTextColumn FieldName="ForecastN2" ShowInCustomizationForm="True"
                            Width="80px" Caption="FORECAST N+2" VisibleIndex="22">
                            <PropertiesTextEdit DisplayFormatString="{0:n0}">
                                <MaskSettings Mask="&lt;0..9999999999999g&gt;" IncludeLiterals="DecimalSymbol"></MaskSettings>
                            </PropertiesTextEdit>
                            <HeaderStyle HorizontalAlign="Center"></HeaderStyle>
                        </dx:GridViewDataTextColumn>
                        <dx:GridViewDataTextColumn FieldName="ForecastN3" Caption="FORECAST N+3" HeaderStyle-HorizontalAlign="Center"
                            VisibleIndex="23" Width="80px">
                            <PropertiesTextEdit DisplayFormatString="{0:n0}">
                                <MaskSettings Mask="<0..9999999999999g>" IncludeLiterals="DecimalSymbol" />
                            </PropertiesTextEdit>
                            <HeaderStyle HorizontalAlign="Center"></HeaderStyle>
                        </dx:GridViewDataTextColumn>
                        <dx:GridViewDataTextColumn FieldName="Sort" Caption="Sort" HeaderStyle-HorizontalAlign="Center"
                            VisibleIndex="24" Width="80px">
                            <HeaderStyle HorizontalAlign="Center"></HeaderStyle>
                        </dx:GridViewDataTextColumn>
                    </Columns>
                    <SettingsBehavior AllowSelectByRowClick="True" AllowSort="False" ColumnResizeMode="Control"
                        EnableRowHotTrack="True" />
                    <SettingsBehavior AllowSort="False" AllowSelectByRowClick="True" ColumnResizeMode="Control"
                        EnableRowHotTrack="True"></SettingsBehavior>
                    <SettingsPager Mode="ShowAllRecords">
                    </SettingsPager>
                    <SettingsEditing Mode="Batch" NewItemRowPosition="Bottom">
                        <BatchEditSettings ShowConfirmOnLosingChanges="False" />
                        <BatchEditSettings ShowConfirmOnLosingChanges="False"></BatchEditSettings>
                    </SettingsEditing>
                    <Settings ShowHorizontalScrollBar="True" ShowVerticalScrollBar="True" ShowGroupButtons="False"
                        ShowStatusBar="Hidden" VerticalScrollableHeight="190" />
                    <Settings ShowGroupButtons="False" ShowVerticalScrollBar="True" ShowHorizontalScrollBar="True"
                        VerticalScrollableHeight="190" ShowStatusBar="Hidden"></Settings>
                    <SettingsCommandButton EditButton-ButtonType="Link">
                        <EditButton Text="Detail">
                        </EditButton>
                    </SettingsCommandButton>
                    <Styles>
                        <SelectedRow ForeColor="Black">
                        </SelectedRow>
                        <BatchEditCell BackColor="Yellow">
                        </BatchEditCell>
                        <BatchEditModifiedCell BackColor="Yellow">
                        </BatchEditModifiedCell>
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
                &nbsp;
            </td>
            <td valign="top" align="right" style="width: 50px;">
                <dx:ASPxButton ID="btnApprove" runat="server" AutoPostBack="False" Font-Names="Tahoma"
                    Font-Size="8pt" Text="FINAL APPROVE" Width="90px" 
                    ClientInstanceName="btnApprove">
                    <ClientSideEvents Click="function(s, e) {                    
                    lblInfo.SetText('');                    
                    ButtonApprove.PerformCallback();
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
    <dx:ASPxCallback ID="cbPONo" runat="server" ClientInstanceName="cbPONo">
        <ClientSideEvents EndCallback="function(s, e) {
	        if (s.cpCommercialCls) {
	            txtCommercial.SetText(s.cpCommercialCls);
	        }

            if (s.cpSupplierID) {
	            txtSupplierCode.SetText(s.cpSupplierID);
	        }

            if (s.cpSupplierName) {
	            txtSupplierName.SetText(s.cpSupplierName);
	        }

            if (s.cpShipCls) {
	            txtShipBy.SetText(s.cpShipCls);
	        }

            if (s.cpPODeliveryBy) {
	            txtDeliveryLocation.SetText(s.cpPODeliveryBy);
	        }

        }" />
    </dx:ASPxCallback>
</asp:Content>
