<%@ Page Title="" Language="vb" AutoEventWireup="false" MasterPageFile="~/SiteContent.Master" CodeBehind="POExportEntryEmergency.aspx.vb" Inherits="PASISystem.POExportEntryEmergency" %>
<%@ Register Assembly="DevExpress.Web.v14.1, Version=14.1.6.0, Culture=neutral, PublicKeyToken=b88d1754d700e49a" namespace="DevExpress.Web.ASPxGridView.Export" TagPrefix="dx" %>
<%@ Register Assembly="DevExpress.Web.v14.1, Version=14.1.6.0, Culture=neutral, PublicKeyToken=b88d1754d700e49a" namespace="DevExpress.Web.ASPxGridView" TagPrefix="dx" %>
<%@ Register assembly="DevExpress.Web.v14.1, Version=14.1.6.0, Culture=neutral, PublicKeyToken=b88d1754d700e49a" namespace="DevExpress.Web.ASPxEditors" tagprefix="dx" %>
<%@ Register Assembly="DevExpress.Web.v14.1, Version=14.1.6.0, Culture=neutral, PublicKeyToken=b88d1754d700e49a" namespace="DevExpress.Web.ASPxGlobalEvents" TagPrefix="dx" %>
<%@ Register Assembly="DevExpress.Web.v14.1, Version=14.1.6.0, Culture=neutral, PublicKeyToken=b88d1754d700e49a" namespace="DevExpress.Web.ASPxPanel" TagPrefix="dx" %>
<%@ Register Assembly="DevExpress.Web.v14.1, Version=14.1.6.0, Culture=neutral, PublicKeyToken=b88d1754d700e49a" namespace="DevExpress.Web.ASPxRoundPanel" TagPrefix="dx" %>
<%@ Register Assembly="DevExpress.Web.v14.1, Version=14.1.6.0, Culture=neutral, PublicKeyToken=b88d1754d700e49a" Namespace="DevExpress.Web.ASPxCallback" TagPrefix="dx" %>

<%@ Register assembly="DevExpress.Web.v14.1, Version=14.1.6.0, Culture=neutral, PublicKeyToken=b88d1754d700e49a" namespace="DevExpress.Web.ASPxUploadControl" tagprefix="dx" %>

<asp:Content ID="Content1" ContentPlaceHolderID="head" runat="server">
    <style type="text/css">
        .dxeHLC, .dxeHC, .dxeHFC
        {
            display: none;
        }
        .style2
        {
            height: 20px;
            width: 157px;
        }
        .style7
        {
            width: 140px;
        }
        .style8
        {
            height: 20px;
            width: 107px;
        }
        .style9
        {
            height: 20px;
            width: 130px;
        }
        .style10
        {
            width: 115px;
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
       
        if (currentColumnName == "NoUrut" || currentColumnName == "PartNo" || currentColumnName == "PartName" || currentColumnName == "Description"
            || currentColumnName == "MOQ" || currentColumnName == "PONo" || currentColumnName == "QtyBox"
            || currentColumnName == "AffiliateID" || currentColumnName == "SupplierID" || currentColumnName == "Period") {
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
            //var hariIsi;
            //hariIsi = 31;
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

            var qtyMonthly1 = s.batchEditApi.GetCellValue(e.visibleIndex, "Week1");
            var qtyMonthly2 = s.batchEditApi.GetCellValue(e.visibleIndex, "Week2");
            var qtyMonthly3 = s.batchEditApi.GetCellValue(e.visibleIndex, "Week3");
            var qtyMonthly4 = s.batchEditApi.GetCellValue(e.visibleIndex, "Week4");
            var qtyMonthly5 = s.batchEditApi.GetCellValue(e.visibleIndex, "Week5");

            total = parseInt(qtyMonthly1) + parseInt(qtyMonthly2) + parseInt(qtyMonthly3) + parseInt(qtyMonthly4) + parseInt(qtyMonthly5);
            s.batchEditApi.SetCellValue(e.visibleIndex, "TotalPOQty", total);

        }, 10);
    }

    function validasubmit() {
        lblInfo.GetMainElement().style.color = 'Red';
        if (txtPONo.GetText() == "") {
            lblInfo.SetText("[6011] Please Input PO No. First!");
            txtPONo.Focus();
            e.ProcessOnServer = false;
            return false;
        }

        if (cboAffiliate.GetText() == "") {
            lblInfo.SetText("[6011] Please Select Affiliate Code First!");
            cboAffiliate.Focus();
            e.ProcessOnServer = false;
            return false;
        }

       

        if (cboSupplier.GetText() == "") {
            lblInfo.SetText("[6011] Please Select Supplier Code First!");
            cboSupplier.Focus();
            e.ProcessOnServer = false;
            return false;
        }

    }

    function up_Insert() {
        var pIsUpdate = '';
        var pStartDate = dtPeriodFrom.GetValue();
        var vStartDate = pStartDate.getMonth() + '/' + pStartDate.getDate() + '/' + pStartDate.getFullYear();
        var pPONo = txtPONo.GetValue();
        var pAffiliateID = cboAffiliate.GetSelectedItem().GetColumnText(0);
        var pSupplierID = cboSupplier.GetSelectedItem().GetColumnText(0);
        var pComercial = rdrCom1.GetChecked();
        var pEmergency = rdrEmergency2.GetChecked();
        var pShip = rdrShip2.GetChecked();
        var pOrder1 = txtOrder1.GetText();
        var pOrder2 = txtOrder2.GetText();
        var pOrder3 = txtOrder3.GetText();
        var pOrder4 = txtOrder4.GetText();
        var pOrder5 = txtOrder5.GetText();
        var pVendor1 = dt1.GetValue();
        var pVendor2 = dt2.GetValue();
        var pVendor3 = dt3.GetValue();
        var pVendor4 = dt4.GetValue();
        var pVendor5 = dt5.GetValue();
        var pETDPort1 = dt6.GetValue();
        var pETDPort2 = dt7.GetValue();
        var pETDPort3 = dt8.GetValue();
        var pETDPort4 = dt9.GetValue();
        var pETDPort5 = dt10.GetValue();
        var pETAPort1 = dt11.GetValue();
        var pETAPort2 = dt12.GetValue();
        var pETAPort3 = dt13.GetValue();
        var pETAPort4 = dt14.GetValue();
        var pETAPort5 = dt15.GetValue();
        var pETAFactory1 = dt16.GetValue();
        var pETAFactory2 = dt17.GetValue();
        var pETAFactory3 = dt18.GetValue();
        var pETAFactory4 = dt19.GetValue();
        var pETAFactory5 = dt20.GetValue();     

        var vPeriod = pPeriod.getMonth() + '/' + pPeriod.getDate() + '/' + pPeriod.getFullYear();
        var vVendor1 = pVendor1.getMonth() + '/' + pVendor1.getDate() + '/' + pVendor1.getFullYear();
        var vVendor2 = pVendor2.getMonth() + '/' + pVendor2.getDate() + '/' + pVendor2.getFullYear();
        var vVendor3 = pVendor3.getMonth() + '/' + pVendor3.getDate() + '/' + pVendor3.getFullYear();
        var vVendor4 = pVendor4.getMonth() + '/' + pVendor4.getDate() + '/' + pVendor4.getFullYear();
        var vVendor5 = pVendor5.getMonth() + '/' + pVendor5.getDate() + '/' + pVendor5.getFullYear();
        var vETDPort1 = pETDPort1.getMonth() + '/' + pETDPort1.getDate() + '/' + pETDPort1.getFullYear();
        var vETDPort2 = pETDPort2.getMonth() + '/' + pETDPort2.getDate() + '/' + pETDPort2.getFullYear();
        var vETDPort3 = pETDPort3.getMonth() + '/' + pETDPort3.getDate() + '/' + pETDPort3.getFullYear();
        var vETDPort4 = pETDPort4.getMonth() + '/' + pETDPort4.getDate() + '/' + pETDPort4.getFullYear();
        var vETDPort5 = pETDPort5.getMonth() + '/' + pETDPort5.getDate() + '/' + pETDPort5.getFullYear();
        var vETAPort1 = pETAPort1.getMonth() + '/' + pETAPort1.getDate() + '/' + pETAPort1.getFullYear();
        var vETAPort2 = pETAPort2.getMonth() + '/' + pETAPort2.getDate() + '/' + pETAPort2.getFullYear();
        var vETAPort3 = pETAPort3.getMonth() + '/' + pETAPort3.getDate() + '/' + pETAPort3.getFullYear();
        var vETAPort4 = pETAPort4.getMonth() + '/' + pETAPort4.getDate() + '/' + pETAPort4.getFullYear();
        var vETAPort5 = pETAPort5.getMonth() + '/' + pETAPort5.getDate() + '/' + pETAPort5.getFullYear();
        var vETAFactory1 = pETAFactory1.getMonth() + '/' + pETAFactory1.getDate() + '/' + pETAFactory1.getFullYear();
        var vETAFactory2 = pETAFactory2.getMonth() + '/' + pETAFactory2.getDate() + '/' + pETAFactory2.getFullYear();
        var vETAFactory3 = pETAFactory3.getMonth() + '/' + pETAFactory3.getDate() + '/' + pETAFactory3.getFullYear();
        var vETAFactory4 = pETAFactory4.getMonth() + '/' + pETAFactory4.getDate() + '/' + pETAFactory4.getFullYear();
        var vETAFactory5 = pETAFactory5.getMonth() + '/' + pETAFactory5.getDate() + '/' + pETAFactory5.getFullYear();


        grid.PerformCallback('save|' + pIsUpdate + '|' + vPeriod + '|' + pPONo + '|' + pAffiliateID + '|' + pSupplierID + '|' + pComercial + '|' + pEmergency + '|' + pShip + '|' + pOrder1 + '|' + pOrder2 + '|' + pOrder3 + '|' + pOrder4 + '|' + pOrder5 + '|' + vVendor1 + '|' + vVendor2 + '|' + vVendor3 + '|' + vVendor4 + '|' + vVendor5 + '|' + vETDPort1 + '|' + vETDPort2 + '|' + vETDPort3 + '|' + vETDPort4 + '|' + vETDPort5 + '|' + vETAPort1 + '|' + vETAPort2 + '|' + vETAPort3 + '|' + vETAPort4 + '|' + vETAPort5 + '|' + vETAFactory1 + '|' + vETAFactory2 + '|' + vETAFactory3 + '|' + vETAFactory4 + '|' + vETAFactory5);

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
            <td colspan="16" width="100%">
                <table style="border-width: 1pt thin thin thin; border-style: ridge; border-color:#9598A1; width:100%; height: 70px;">
                    <tr>
                        <td colspan="8" height="20">
                            <table id="Table1">
                                <tr>
                                    <td style="width:5px;"></td>
                                    <td align="left" valign="middle" height="20px" width="110px">
                                        <dx:ASPxLabel ID="ASPxLabel1" runat="server" Text="PERIOD"
                                            Font-Names="Tahoma" Font-Size="8pt">
                                        </dx:ASPxLabel>
                                    </td>
                                    <td align="left" valign="middle" height="20px" width="110px">
                                        <dx:ASPxTimeEdit ID="dtPeriodFrom" runat="server" ClientInstanceName="dtPeriodFrom"
                                            DisplayFormatString="yyyy-MM" EditFormat="Custom" EditFormatString="yyyy-MM"
                                            Width="110px" Font-Names="Tahoma" Font-Size="8pt"   
                                            height="15px">
                                        </dx:ASPxTimeEdit>                                                                                
                                    </td>
                                    <td style="width:5px;">&nbsp;</td>
                                                <td align="left">
                                                    <table style="width:90%;">
                                                        <tr>
                                                            <td>
                                        <dx:ASPxLabel ID="ASPxLabel4" runat="server" Text="COMMERCIAL"
                                            Font-Names="Tahoma" Font-Size="8pt">
                                        </dx:ASPxLabel>
                                                            </td>
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
                                    <td class="style7">
                                        <dx:ASPxLabel ID="ASPxLabel32" runat="server" Text="DELIVERY LOCATION"
                                            Font-Names="Tahoma" Font-Size="8pt" width="100%">
                                        </dx:ASPxLabel>
                                    </td>
                                    <td align="left" valign="middle" class="style9">
                                        <dx:ASPxComboBox ID="cboDelLoc" width="130px" runat="server" 
                                            Font-Size="8pt" Font-Names="Tahoma" TextFormatString="{0}" 
                                            ClientInstanceName="cboDelLoc" TabIndex="3">
                                            <ClientSideEvents SelectedIndexChanged="function(s, e) {
	                                           txtDelLoc.SetText(cboDelLoc.GetSelectedItem().GetColumnText(1));                                   
	                                            grid.PerformCallback('kosong');	
                                                lblInfo.SetText('');	
                                            }" BeginCallback="function(s, e) {
	txtDelLoc.SetText(cboDelLoc.GetSelectedItem().GetColumnText(1)); 
}" CallbackError="function(s, e) {
	txtDelLoc.SetText(cboDelLoc.GetSelectedItem().GetColumnText(1)); 
}" EndCallback="function(s, e) {
	txtDelLoc.SetText(cboDelLoc.GetSelectedItem().GetColumnText(1)); 
}" Init="function(s, e) {
	txtDelLoc.SetText(cboDelLoc.GetSelectedItem().GetColumnText(1)); 
}" LostFocus="function(s, e) {
	txtDelLoc.SetText(cboDelLoc.GetSelectedItem().GetColumnText(1)); 
}" TextChanged="function(s, e) {
	txtDelLoc.SetText(cboDelLoc.GetSelectedItem().GetColumnText(1)); 
}" ValueChanged="function(s, e) {
	txtDelLoc.SetText(cboDelLoc.GetSelectedItem().GetColumnText(1)); 
}" />
                                            <LoadingPanelStyle ImageSpacing="5px">
                                            </LoadingPanelStyle>
                                        </dx:ASPxComboBox>
                                    </td>
                                    <td align="left" valign="middle" style="height:20px; width:60px;">
                                        <dx:ASPxTextBox ID="txtDelLoc" runat="server" Width="250px" 
                                            ClientInstanceName="txtDelLoc" Font-Names="Tahoma" Font-Size="8pt"
                                            ReadOnly="True" MaxLength="100" Height="20px">
                                            <readonlystyle backcolor="#CCCCCC">
                                            </readonlystyle>
                                        </dx:ASPxTextBox>
                                    </td>
                                </tr>
                                <tr>
                                    <td style="width:5px;"></td>
                                    <td align="left" valign="middle" class="style2">
                                        <dx:ASPxLabel ID="ASPxLabel5" runat="server" Text="PO MONTHLY /EMERGENCY"
                                            Font-Names="Tahoma" Font-Size="8pt">
                                        </dx:ASPxLabel>
                                    </td>
                                    <td align="left" valign="middle" height="20px" width="110px">
                                        <dx:ASPxTextBox ID="txtPOEmergency" runat="server" Width="110px" 
                                            ClientInstanceName="txtPOEmergency" Font-Names="Tahoma" Font-Size="8pt"
                                            MaxLength="1" ReadOnly="True" Height="20px" 
                                            TabIndex="2" Text="E">
                                            <ClientSideEvents LostFocus="function(s, e) { 

	                                            lblInfo.SetText('');
                                            }" />
                                            <readonlystyle backcolor="#CCCCCC">
                                            </readonlystyle>
                                        </dx:ASPxTextBox>
                                    </td>
                                     <td style="width:5px;">&nbsp;</td>
                                                <td align="left">
                                                    <table style="width:89%;">
                                            <tr>
                                                <td>
                                        <dx:ASPxLabel ID="ASPxLabel2" runat="server" Text="SHIP BY"
                                            Font-Names="Tahoma" Font-Size="8pt">
                                        </dx:ASPxLabel>
                                                </td>
 
                                                <td>
                                                    &nbsp;</td>
                                                <td>
                                                    &nbsp;</td>
                                                <td>
                                                    &nbsp;</td>
                                                <td>
                                                    &nbsp;</td>
                                                    <td>
                                                    &nbsp;</td>
                                                <td>
                                                    &nbsp;</td>
                                                            <td>
                                                    <dx:ASPxRadioButton ID="rdrShipBy2" ClientInstanceName="rdrShipBy2" runat="server" 
                                                                    Text="BOAT" GroupName="POEm" Font-Names="Tahoma" Font-Size="8pt">
                                                        <ClientSideEvents LostFocus="function(s, e) {                                                        
                                                        lblInfo.SetText('');                                                         
                                                        }" />
                                                    </dx:ASPxRadioButton>
                                                            </td>
                                                            <td>
                                                    <dx:ASPxRadioButton ID="rdrShipBy3" ClientInstanceName="rdrShipBy3" runat="server" 
                                                                    Text="AIR" GroupName="POEm" Font-Names="Tahoma" Font-Size="8pt">
                                                        <ClientSideEvents LostFocus="function(s, e) {
                                                        lblInfo.SetText('');                                                         
                                                        }" />
                                                    </dx:ASPxRadioButton>
                                                            </td>
                                            </tr>
                                        </table>
                                    </td>
                                    <td align="left" valign="middle" class="style8">
                                        &nbsp;</td>
                                    <td align="left" valign="middle" height="20px" class="style10">
                                        &nbsp;</td>

                                    <td align="left" valign="middle" height="20px" width="180px">
                                        &nbsp;</td>

                                </tr>                               
                                <tr>
                                    <td style="width:5px;">&nbsp;</td>
                                    <td align="left" valign="middle" class="style2">
                                        <dx:ASPxLabel ID="ASPxLabel20" runat="server" Text="AFFILIATE CODE/NAME"
                                            Font-Names="Tahoma" Font-Size="8pt">
                                        </dx:ASPxLabel>
                                    </td>
                                    <td align="left" valign="middle" height="20px" width="110px">
                                        <dx:ASPxComboBox ID="cboAffiliate" width="110px" runat="server" Font-Size="8pt" 
                                            Font-Names="Tahoma" TextFormatString="{0}" ClientInstanceName="cboAffiliate" 
                                            TabIndex="3">
                                            <ClientSideEvents SelectedIndexChanged="function(s, e) {
	                                            txtAffiliate.SetText(cboAffiliate.GetSelectedItem().GetColumnText(1));                               
	                                            grid.PerformCallback('load');
	                                            lblInfo.SetText('');	
                                            }" />
                                            <LoadingPanelStyle ImageSpacing="5px">
                                            </LoadingPanelStyle>
                                        </dx:ASPxComboBox>
                                    </td>
                                    <td style="width:5px;">&nbsp;</td>
                                    <td style="width:5px;">
                                        <dx:ASPxTextBox ID="txtAffiliate" runat="server" Width="250px" 
                                            ClientInstanceName="txtAffiliate" Font-Names="Tahoma" Font-Size="8pt"
                                            ReadOnly="True" MaxLength="100" Height="20px">
                                            <readonlystyle backcolor="#CCCCCC">
                                            </readonlystyle>
                                        </dx:ASPxTextBox>
                                    </td>
                                    <td align="left" valign="middle" class="style8">
                                        &nbsp;</td>
                                    <td align="left" valign="middle" height="20px" class="style10">
                                        &nbsp;</td>

                                    <td align="right" valign="middle" height="20px" width="180px">
                <dx:ASPxButton ID="btnClear" runat="server" Text="CLEAR"
                Font-Names="Tahoma" Width="85px" AutoPostBack="False" Font-Size="8pt" ImagePosition="Right">
                    <ClientSideEvents Click="function(s, e) {
    
    dtPeriodFrom.SetDate(new Date());
	rdrCom1.SetChecked(true);
    rdrCom2.SetChecked(false);
    rdrShipBy2.SetChecked(true);
    rdrShipBy3.SetChecked(false);
	txtPOEmergency.SetText('E');
	cboAffiliate.SetText('');
    txtAffiliate.SetText('');
    txtOrder1.SetText('');
    dt1.SetDate(new Date());
    dt6.SetDate(new Date());
    dt11.SetDate(new Date());
    dt16.SetDate(new Date());
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

            <table style="width: 100%;">
        <tr>
            <td colspan="8" align="center" valign="top">      
                <table style="width:100%;">
                    <tr>
                        <td width="120px">
                            &nbsp;</td>
                        <td width="120px">
                            &nbsp;</td>
                        <td rowspan="3">
                            <table border="0" cellpadding="0" cellspacing="0" width="300px">                                
                                <tr>
                                    <td width="150px" height="16px">
                                        <dx:ASPxTextBox runat="server" Width="150px" MaxLength="10" Height="16px" Font-Names="Tahoma"
                                            Font-Size="8pt" ID="ASPxTextBox5" BackColor="#FFD2A6" Text="ORDER NO" ReadOnly="True"
                                            Font-Bold="True">
                                        </dx:ASPxTextBox>
                                    </td>
                                    <td width="150px" height="16px">
                                        <dx:ASPxTextBox runat="server" Width="150px" MaxLength="10" Height="16px" ClientInstanceName="txtOrder1"
                                            Font-Names="Tahoma" Font-Size="8pt" ID="txtOrder1"
                                            >
                                        </dx:ASPxTextBox>
                                    </td>
                                </tr>
                                <tr>
                                    <td width="150px" height="16px">
                                        <dx:ASPxTextBox runat="server" Width="150px" MaxLength="10" Height="16px" Font-Names="Tahoma"
                                            Font-Size="8pt" ID="ASPxTextBox6" BackColor="#FFD2A6" Text="ETD VENDOR" 
                                            Font-Bold="True">
                                        </dx:ASPxTextBox>
                                    </td>
                                    <td width="150px" height="16px">
                                        <dx:ASPxDateEdit ID="dt1" runat="server" ClientInstanceName="dt1"
                                            DisplayFormatString="yyyy-MM-dd" EditFormat="Custom" EditFormatString="yyyy-MM-dd"
                                            Font-Names="Tahoma" Font-Size="8pt" Width="150px">                                           
                                        </dx:ASPxDateEdit>
                                    </td>
                                </tr>                                
                                <tr>
                                    <td width="150px" height="16px">
                                        <dx:ASPxTextBox runat="server" Width="150px" MaxLength="10" Height="16px" Font-Names="Tahoma"
                                            Font-Size="8pt" ID="ASPxTextBox7" BackColor="#FFD2A6" Text="ETD PORT" 
                                            Font-Bold="True">
                                        </dx:ASPxTextBox>
                                    </td>
                                    <td width="150px" height="16px">
                                        <dx:ASPxDateEdit ID="dt6" runat="server" ClientInstanceName="dt6"
                                            DisplayFormatString="yyyy-MM-dd" EditFormat="Custom" EditFormatString="yyyy-MM-dd"
                                            Font-Names="Tahoma" Font-Size="8pt" Width="150px">
                                        </dx:ASPxDateEdit>
                                    </td>
                                </tr>                                
                                <tr>
                                    <td width="150px" height="16px">
                                        <dx:ASPxTextBox runat="server" Width="150px" MaxLength="10" Height="16px" Font-Names="Tahoma"
                                            Font-Size="8pt" ID="ASPxTextBox8" BackColor="#FFD2A6" Text="ETA PORT"
                                            Font-Bold="True">
                                        </dx:ASPxTextBox>
                                    </td>
                                    <td width="150px" height="16px">
                                        <dx:ASPxDateEdit ID="dt11" runat="server" ClientInstanceName="dt11"
                                            DisplayFormatString="yyyy-MM-dd" EditFormat="Custom" EditFormatString="yyyy-MM-dd"
                                            Font-Names="Tahoma" Font-Size="8pt" Width="150px">
                                        </dx:ASPxDateEdit>
                                    </td>
                                </tr>                                
                                <tr>
                                    <td width="150px" height="16px">
                                        <dx:ASPxTextBox runat="server" Width="150px" MaxLength="10" Height="16px" Font-Names="Tahoma"
                                            Font-Size="8pt" ID="ASPxTextBox9" BackColor="#FFD2A6" Text="ETA FACTORY" 
                                            Font-Bold="True">
                                        </dx:ASPxTextBox>
                                    </td>
                                    <td width="150px" height="16px">
                                        <dx:ASPxDateEdit ID="dt16" runat="server" ClientInstanceName="dt16"
                                            DisplayFormatString="yyyy-MM-dd" EditFormat="Custom" EditFormatString="yyyy-MM-dd"
                                            Font-Names="Tahoma" Font-Size="8pt" Width="150px" >
                                        </dx:ASPxDateEdit>
                                    </td>                                    
                                </tr>                                
                             </table>
                        </td>
                        <td width="120px">
                            &nbsp;</td>
                        <td width="120px">
                            &nbsp;</td>
                        <td width="120px" valign="top">
                            <asp:TextBox ID="lSuuplier0" runat="server" BackColor="Red" BorderStyle="None"
                                ReadOnly="True" Width="30px"></asp:TextBox>
                            <dx:ASPxLabel runat="server" Text=": ERROR" Font-Names="Tahoma" Font-Size="8pt"
                                ID="ASPxLabel3" Width="80px" RightToLeft="False">
                            </dx:ASPxLabel>
                        </td>
                    </tr>
                    <tr>
                        <td width="120px">
                            &nbsp;</td>
                        <td width="120px">
                            &nbsp;</td>
                        <td width="120px">
                            &nbsp;</td>
                        <td width="120px">
                            &nbsp;</td>
                        <td width="180px">
                         <dx:ASPxRoundPanel ID="ASPxRoundPanel1" runat="server" HeaderText="ERROR CHECK" ShowCollapseButton="true"
                    View="GroupBox" Width="100%" Font-Size="8pt" Font-Names="Tahoma" BackColor="#FFD2A6" 
                                ForeColor="Red" >
                    <PanelCollection>
                        <dx:PanelContent ID="PanelContent1" runat="server">
                            <table id="Table2">
                                <tr>
                                    <td align="left" height="50px" valign="middle" width="100%">
                                        <dx:ASPxLabel ID="ASPxLabel8" runat="server" Font-Names="Tahoma" Font-Size="8pt"
                                            ForeColor="Red" Text="NUMBER OF ROW : " Width="100%" bgcolor="#FFD2A6">
                                        </dx:ASPxLabel>
                                    </td>
                                </tr>
                            </table>
                        </dx:PanelContent>
                    </PanelCollection>
                </dx:ASPxRoundPanel>
                            &nbsp;</td>
                    </tr>
                    <tr>
                        <td width="120px">
                            &nbsp;</td>
                        <td width="120px">
                            &nbsp;</td>
                        <td width="120px">
                            &nbsp;</td>
                        <td width="120px">
                            &nbsp;</td>
                        <td width="120px">
                            &nbsp;</td>
                    </tr>
                </table>
            </td>
        </tr>
        </table>
        
        <div style="height: 1px;"></div>
                
      <table style="width: 100%;">
        <tr>
            <td colspan="16" align="left" valign="top" height="220">
                <%--Column : Grid--%>
                <dx:ASPxGridView ID="grid" runat="server" Width="100%"
                    Font-Names="Tahoma" KeyFieldName="PONo;PartNo;AffiliateID;SupplierID"
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
<ClientSideEvents FocusedRowChanged="function(s, e) {
	                    OnGridFocusedRowChanged();
                    }" RowClick="function(s, e) {
	                    lblInfo.SetText(&#39;&#39;);
                    }" BatchEditStartEditing="OnBatchEditStartEditing" 
                        BatchEditEndEditing="OnBatchEditEndEditing" EndCallback="function(s, e) {
                        var pMsg = s.cpMessage;
                        if (pMsg != &#39;&#39;) {
                            if (pMsg.substring(1,5) == &#39;1001&#39; || pMsg.substring(1,5) == &#39;1002&#39; || pMsg.substring(1,5) == &#39;1003&#39; || pMsg.substring(1,5) == &#39;2001&#39;) {
                                lblInfo.GetMainElement().style.color = &#39;Blue&#39;;
                                txtUser1.SetText(s.cpUser1);
                                txtDate1.SetText(s.cpDate1);     
                                txtPONo.SetText(s.cpPONo);  
                            } else {
                                lblInfo.GetMainElement().style.color = &#39;Red&#39;;
                            }
                            lblInfo.SetText(pMsg);
                        } else {
                            lblInfo.SetText(&#39;&#39;);
                        }
                        delete s.cpMessage;
                        delete s.cpPONo;
                    }" CallbackError="function(s, e) {e.handled = true;}" Init="OnInit"></ClientSideEvents>
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
                        <dx:GridViewDataTextColumn VisibleIndex="1" Caption="NO." FieldName="NoUrut" 
                            Width="30px" HeaderStyle-HorizontalAlign="Center">                            
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
                        <dx:GridViewDataTextColumn VisibleIndex="4" Caption="UOM" Width="40px" 
                            HeaderStyle-HorizontalAlign="Center" CellStyle-HorizontalAlign="Center" 
                            FieldName="Description">                            
<HeaderStyle HorizontalAlign="Center"></HeaderStyle>

                            <CellStyle Font-Names="Tahoma" Font-Size="8pt">
                            </CellStyle>
                        </dx:GridViewDataTextColumn>
                        <dx:GridViewDataTextColumn VisibleIndex="5" Caption="MOQ" FieldName="MOQ" 
                            Width="70px" HeaderStyle-HorizontalAlign="Center">                            
                            <propertiestextedit displayformatstring="{0:n0}"></propertiestextedit>
<HeaderStyle HorizontalAlign="Center"></HeaderStyle>

                            <CellStyle Font-Names="Tahoma" Font-Size="8pt">
                            </CellStyle>
                        </dx:GridViewDataTextColumn>
                        
                        <dx:GridViewDataTextColumn VisibleIndex="6" Caption="QTY/BOX" 
                            FieldName="QtyBox" Width="70px" HeaderStyle-HorizontalAlign="Center">                            
                            <propertiestextedit displayformatstring="{0:n0}"></propertiestextedit>
<HeaderStyle HorizontalAlign="Center"></HeaderStyle>

                            <CellStyle Font-Names="Tahoma" Font-Size="8pt">
                            </CellStyle>
                        </dx:GridViewDataTextColumn>  
<dx:GridViewDataTextColumn VisibleIndex="12" Caption="TOTAL FIRM QTY" 
                            FieldName="TotalPOQty" Width="90px" HeaderStyle-HorizontalAlign="Center">                            
    <propertiestextedit displayformatstring="{0:n0}"></propertiestextedit>
<HeaderStyle HorizontalAlign="Center"></HeaderStyle>

                            <CellStyle Font-Names="Tahoma" Font-Size="8pt">
                            </CellStyle>
                        </dx:GridViewDataTextColumn>

                        <dx:GridViewDataTextColumn VisibleIndex="7" Caption="WEEK 1 FIRM QTY" 
                            HeaderStyle-HorizontalAlign="Center" FieldName="Week1" Width="0px" 
                            Visible="False">                            
<HeaderStyle HorizontalAlign="Center"></HeaderStyle>
<CellStyle Font-Names="Tahoma" Font-Size="8pt" BackColor="White">
                            </CellStyle>
                            <PropertiesTextEdit DisplayFormatString="{0:n2}">
                                


<MaskSettings ErrorText="Please input valid value !"
                                Mask="<0..999999g>" IncludeLiterals="DecimalSymbol" />
                            


</PropertiesTextEdit>
                        </dx:GridViewDataTextColumn>

                        <dx:GridViewDataTextColumn VisibleIndex="8" Caption="WEEK 2 FIRM QTY" 
                            HeaderStyle-HorizontalAlign="Center" FieldName="Week2" Width="0px" 
                            Visible="False">
<HeaderStyle HorizontalAlign="Center"></HeaderStyle>
<CellStyle Font-Names="Tahoma" Font-Size="8pt" BackColor="White">
                            </CellStyle>
                            <PropertiesTextEdit DisplayFormatString="{0:n2}">
                                


<MaskSettings ErrorText="Please input valid value !"
                                Mask="<0..999999g>" IncludeLiterals="DecimalSymbol" />
                            


</PropertiesTextEdit>  
                        </dx:GridViewDataTextColumn>

                        <dx:GridViewDataTextColumn VisibleIndex="9" Caption="WEEK 3 FIRM QTY" 
                            HeaderStyle-HorizontalAlign="Center" CellStyle-HorizontalAlign="Center" 
                            Visible="False" FieldName="Week3" Width="0px">
                            <HeaderStyle HorizontalAlign="Center"></HeaderStyle>
                            <CellStyle Font-Names="Tahoma" Font-Size="8pt" BackColor="White">
                            </CellStyle>
                            <PropertiesTextEdit DisplayFormatString="{0:n2}">
                                


<MaskSettings ErrorText="Please input valid value !"
                                Mask="<0..999999g>" IncludeLiterals="DecimalSymbol" />
                            


</PropertiesTextEdit>  
                        </dx:GridViewDataTextColumn>  
                        <dx:GridViewDataTextColumn VisibleIndex="10" Caption="WEEK 4 FIRM QTY" 
                            HeaderStyle-HorizontalAlign="Center" Visible="False" FieldName="Week4" 
                            Width="0px">
                            <PropertiesTextEdit DisplayFormatString="{0:n2}">
                                


<MaskSettings ErrorText="Please input valid value !"
                                Mask="<0..999999g>" IncludeLiterals="DecimalSymbol" />
                            


</PropertiesTextEdit>                            

<HeaderStyle HorizontalAlign="Center"></HeaderStyle>

                            <CellStyle Font-Names="Tahoma" Font-Size="8pt" BackColor="White">
                            </CellStyle>
                        </dx:GridViewDataTextColumn>  
                        <dx:GridViewDataTextColumn VisibleIndex="11" Caption="WEEK 5 FIRM QTY" 
                            HeaderStyle-HorizontalAlign="Center" FieldName="Week5" Width="0px" 
                            Visible="False">
                            <PropertiesTextEdit DisplayFormatString="{0:n2}">
                                


<MaskSettings ErrorText="Please input valid value !"
                                Mask="<0..999999g>" IncludeLiterals="DecimalSymbol" />
                            


</PropertiesTextEdit>                            

<HeaderStyle HorizontalAlign="Center"></HeaderStyle>

                            <CellStyle Font-Names="Tahoma" Font-Size="8pt" BackColor="White">
                            </CellStyle>
                        </dx:GridViewDataTextColumn>
                        <dx:GridViewDataTextColumn Caption="AFFILIATE CODE" FieldName="AffiliateID" 
                            VisibleIndex="19" Width="0px">
                        </dx:GridViewDataTextColumn>
                        <dx:GridViewDataTextColumn Caption="SUPPLIER CODE" FieldName="SupplierID" 
                            VisibleIndex="20" Width="0px">
                        </dx:GridViewDataTextColumn>
                        <dx:GridViewDataTextColumn FieldName="ShipCls" Width="0px" Caption="SHIP CLS" 
                            VisibleIndex="30">
<HeaderStyle HorizontalAlign="Center"></HeaderStyle>

<CellStyle Font-Names="Tahoma" Font-Size="8pt"></CellStyle>
</dx:GridViewDataTextColumn>
<dx:GridViewDataTextColumn FieldName="CommercialCls" Width="0px" Caption="COMMERCIAL" 
                            VisibleIndex="31">
<HeaderStyle HorizontalAlign="Center"></HeaderStyle>

<CellStyle Font-Names="Tahoma" Font-Size="8pt"></CellStyle>
</dx:GridViewDataTextColumn>
<dx:GridViewDataTextColumn FieldName="PONo" Width="0px" Caption="PO NO." 
                            VisibleIndex="32">
<HeaderStyle HorizontalAlign="Center"></HeaderStyle>

<CellStyle Font-Names="Tahoma" Font-Size="8pt"></CellStyle>
</dx:GridViewDataTextColumn>
<dx:GridViewDataTextColumn FieldName="ForwarderID" Width="0px" Caption="DELIVERY LOCATION" 
                            VisibleIndex="33">
<HeaderStyle HorizontalAlign="Center"></HeaderStyle>

<CellStyle Font-Names="Tahoma" Font-Size="8pt"></CellStyle>
</dx:GridViewDataTextColumn>
                        <dx:GridViewDataDateColumn Caption="Period" FieldName="Period" 
                            VisibleIndex="34" Width="0px">
                           <PropertiesDateEdit EditFormatString="dd MMM yyyy" DisplayFormatString="dd MMM yyyy">
                            </PropertiesDateEdit>
                        </dx:GridViewDataDateColumn>
                    </Columns>
                    <SettingsBehavior AllowSelectByRowClick="True" AllowSort="False"
                        ColumnResizeMode="Control" EnableRowHotTrack="True" />

<SettingsBehavior AllowSort="False" AllowSelectByRowClick="True" ColumnResizeMode="Control" 
                        EnableRowHotTrack="True"></SettingsBehavior>

                    <SettingsPager PageSize="10" Position="Top">
                        <Summary Text="Page {0} of {1} [{2} record(s)]"
                                  AllPagesText="Page {0} of {1} " />
                        <Summary AllPagesText="Page {0} of {1} " Text="Page {0} of {1} [{2} record(s)]"></Summary>
                    </SettingsPager>
                    <SettingsEditing Mode="Batch" NewItemRowPosition="Bottom">
                        <BatchEditSettings ShowConfirmOnLosingChanges="False" />
<BatchEditSettings ShowConfirmOnLosingChanges="False"></BatchEditSettings>
                    </SettingsEditing>
                    <Settings ShowHorizontalScrollBar="True" ShowVerticalScrollBar="True"
                        ShowGroupButtons="False" ShowStatusBar="Hidden"
                        VerticalScrollableHeight="220" />

<Settings ShowGroupButtons="False" ShowVerticalScrollBar="True" ShowHorizontalScrollBar="True" 
                        VerticalScrollableHeight="220" ShowStatusBar="Hidden"></Settings>

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
        
        <table style="width: 100%;">
            <tr>
                <td align="left">
                <dx:ASPxButton ID="btnSubMenu" runat="server" Text="BACK"
                    Font-Names="Tahoma" Width="85px" Font-Size="8pt" 
        HorizontalAlign="Center">
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
                <dx:ASPxTextBox ID="tampung" runat="server" ClientInstanceName="tampung" 
                    Width="0px" BackColor="White" ForeColor="White">
                    <Border BorderColor="White" />
                </dx:ASPxTextBox>                          
                </td>
                <td align="right">
                <dx:ASPxButton ID="btnCheck" runat="server" Text="CHECK DATA"
                    Font-Names="Tahoma"
                    Width="80px" AutoPostBack="False" Font-Size="8pt" 
                    ClientInstanceName="btnCheck" VerticalAlign="Bottom">
                 <ClientSideEvents Click="function(s, e) {
                if (HF.Set('HTcls') == '1') {
        lblerrmessage.SetText('[6011] Data Already Deliver by HT!');
		lblerrmessage.GetMainElement().style.color = 'Red';
        return false;
    }
    if (txtsuratjalanno.GetText() == '') {
        lblerrmessage.SetText('[6011] Please Input Surat Jalan No first!');
		lblerrmessage.GetMainElement().style.color = 'Red';
        return false;
    }
    if (txtdrivername.GetText() == '') {
        lblerrmessage.SetText('[6011] Please Input Driver Name first!');
		lblerrmessage.GetMainElement().style.color = 'Red';
        return false;
    }
    if (txtdrivercontact.GetText() == '') {
        lblerrmessage.SetText('[6011] Please Input Driver Contact first!');
		lblerrmessage.GetMainElement().style.color = 'Red';
        return false;
    }
    if (txtnopol.GetText() == '') {
        lblerrmessage.SetText('[6011] Please Input NO Pol first!');
		lblerrmessage.GetMainElement().style.color = 'Red';
        return false;
    }
    if (txtjenisarmada.GetText() == '') {
        lblerrmessage.SetText('[6011] Please Input Jenis Armada first!');
		lblerrmessage.GetMainElement().style.color = 'Red';
        return false;
    }
	grid.UpdateEdit();
    var millisecondsToWait = 50;
            setTimeout(function() {
                grid.PerformCallback('gridload');
            }, millisecondsToWait);	
	grid.CancelEdit();
                                                          
                                                        }" />
                </dx:ASPxButton>
                </td>
                <td align="right" width="85px">
                <dx:ASPxButton ID="btnApprove" runat="server" Text="SEND TO SUPPLIER"
                    Font-Names="Tahoma"
                    Width="80px" AutoPostBack="False" Font-Size="8pt" 
                    ClientInstanceName="btnApprove" VerticalAlign="Bottom">
                 <ClientSideEvents Click="function(s, e) {                    
                    
                    grid.UpdateEdit();
                    grid.PerformCallback('load');
                    
                }" /> 
                </dx:ASPxButton>
                </td>
                <td align="right" width="85px">
                                                    <dx:ASPxButton ID="btnUpload" 
                    runat="server" Text="UPLOAD ERROR LIST"
                                                        Font-Names="Tahoma" Width="85px" 
                    AutoPostBack="False" Font-Size="8pt" VerticalAlign="Bottom">
                                                        <ClientSideEvents Click="function(s, e) {grid.PerformCallback('save');}" />
                                                    </dx:ASPxButton>
                                                    </td>
                <td align="right" width ="85px">
                                                    <dx:ASPxButton ID="btnDownload" 
                    runat="server" Text="DOWNLOAD ERROR LIST"
                                                        Font-Names="Tahoma" Width="85px" 
                    AutoPostBack="False" Font-Size="8pt" ClientInstanceName="btnDownload" 
                        VerticalAlign="Bottom">
                                                        <ClientSideEvents Click="function(s, e) {grid.PerformCallback('downloadSummary');}" />
                                                    </dx:ASPxButton>
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

