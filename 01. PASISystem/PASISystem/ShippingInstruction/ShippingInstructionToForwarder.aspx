<%@ Page Title="" Language="vb" AutoEventWireup="false" MasterPageFile="~/SiteContent.Master"
    CodeBehind="ShippingInstructionToForwarder.aspx.vb" Inherits="PASISystem.ShippingInstructionToForwarder" %>

<%@ Register Assembly="DevExpress.Web.v14.1, Version=14.1.6.0, Culture=neutral, PublicKeyToken=b88d1754d700e49a"
    Namespace="DevExpress.Web.ASPxEditors" TagPrefix="dx" %>
<%@ Register Assembly="DevExpress.Web.v14.1, Version=14.1.6.0, Culture=neutral, PublicKeyToken=b88d1754d700e49a"
    Namespace="DevExpress.Web.ASPxGridView" TagPrefix="dx" %>
<%@ Register Assembly="DevExpress.Web.v14.1, Version=14.1.6.0, Culture=neutral, PublicKeyToken=b88d1754d700e49a"
    Namespace="DevExpress.Web.ASPxGridView.Export" TagPrefix="dx2" %>
<%@ Register Assembly="DevExpress.Web.v14.1, Version=14.1.6.0, Culture=neutral, PublicKeyToken=b88d1754d700e49a"
    Namespace="DevExpress.Web.ASPxGlobalEvents" TagPrefix="dx" %>
<%@ Register Assembly="DevExpress.Web.v14.1, Version=14.1.6.0, Culture=neutral, PublicKeyToken=b88d1754d700e49a"
    Namespace="DevExpress.Web.ASPxHiddenField" TagPrefix="dx" %>
<%@ Register Assembly="DevExpress.Web.v14.1, Version=14.1.6.0, Culture=neutral, PublicKeyToken=b88d1754d700e49a"
    Namespace="DevExpress.Web.ASPxRoundPanel" TagPrefix="dx" %>
<%@ Register Assembly="DevExpress.Web.v14.1, Version=14.1.6.0, Culture=neutral, PublicKeyToken=b88d1754d700e49a"
    Namespace="DevExpress.Web.ASPxPanel" TagPrefix="dx" %>
<asp:Content ID="Content1" ContentPlaceHolderID="head" runat="server">
    <style type="text/css">
        .dxeBase
        {
            font: 12px Tahoma, Geneva, sans-serif;
        }
        
        .dxeBase
        {
            font: 12px Tahoma, Geneva, sans-serif;
        }
        
        .style3
        {
            height: 200px;
        }
        .style2
        {
            height: 30px;
        }
        .style4
        {
            height: 29px;
        }
        .style5
        {
            width: 94px;
        }
    </style>
</asp:Content>
<asp:Content ID="Content2" ContentPlaceHolderID="Content" runat="server">
    <script type="text/javascript">

        function OnAllCheckedChanged(s, e) {
            if (s.GetValue() == -1) s.SetValue(1);
            for (var i = 0; i < grid.GetVisibleRowsOnPage(); i++) {
                grid.batchEditApi.SetCellValue(i, "Act", s.GetValue());
            }
        }
        function numbersonly(e) {
            var unicode = e.charCode ? e.charCode : e.keyCode
            if (unicode != 8) { //if the key isn't the backspace key (which we should allow)
                if (unicode < 45 || unicode > 57) //if not a number
                    return false //disable key press
            }
        }
        function OnUpdateClick(s, e) {
            Grid.PerformCallback("Update");
        }

        function OnCancelClick(s, e) {
            Grid.PerformCallback("Cancel");
        }

        function OnInit(s, e) {
            AdjustSizeGrid();
        }
        function OnControlsInitializedGrid(s, e) {
            ASPxClientUtils.AttachEventToElement(window, "resize", function (evt) {
                AdjustSizeGrid();
            });
        }
        function AdjustSizeGrid() {
            debugger
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
            height = height - (height * 59 / 100)
            grid.SetHeight(height);
        }     
    </script>
    <table style="width: 100%; height: 190px;">
        <tr>
            <td align="left" style="width: 70%;">
                <dx:ASPxRoundPanel ID="ASPxRoundPanel1" runat="server" HeaderText="SHIPPING INSTRUCTION"
                    ShowCollapseButton="true" View="GroupBox" Width="100%" Height="200px" BackColor="White">
                    <ContentPaddings PaddingLeft="5px" PaddingRight="5px" />
                    <ContentPaddings PaddingLeft="5px" PaddingRight="5px"></ContentPaddings>
                    <PanelCollection>
                        <dx:PanelContent ID="PanelContent1" runat="server">
                            <table id="Table1">
                                <tr>
                                    <td align="left" height="20px" valign="middle" width="40%">
                                        <dx:ASPxComboBox ID="cboCreate" runat="server" Width="100px" ClientInstanceName="cboCreate">
                                            <ClientSideEvents SelectedIndexChanged="function(s, e) {
	                    if (cboCreate.GetText() == 'CREATE') {
	                    cboShippingNo.SetText('');
                    }
txtSend.SetText('');
lblErrMsg.SetText('');
}" Init="function(s, e) {
    if (btnsubmenu.GetText() == 'BACK') {
	    if (cboCreate.GetText() == 'CREATE') {
		    cboAffiliateCode.SetEnabled(false);
            cboForwarder.SetEnabled(false);
	    } 
    } else {
            cboAffiliateCode.SetEnabled(true);
            cboForwarder.SetEnabled(true);
           }
}" />
                                        </dx:ASPxComboBox>
                                    </td>
                                    <td align="left" height="20px" valign="middle" width="5%">
                                        &nbsp;
                                    </td>
                                    <td align="left" height="20px" valign="middle" width="5%">
                                        &nbsp;
                                    </td>
                                    <td align="left" height="20px" valign="middle" width="40%">
                                        &nbsp;
                                    </td>
                                </tr>                               
                                <tr>
                                    <td align="left" height="20px" valign="middle" width="40%">
                                        <dx:ASPxLabel ID="ASPxLabel14" runat="server" Text="AFFILIATE CODE/NAME*">
                                        </dx:ASPxLabel>
                                    </td>
                                    <td align="left" height="20px" valign="middle" width="5%">
                                        &nbsp;
                                    </td>
                                    <td align="left" height="20px" valign="middle" width="5%">
                                        <dx:ASPxComboBox ID="cboAffiliateCode" runat="server" Width="100px" ClientInstanceName="cboAffiliateCode"
                                            TextFormatString="{0}">
                                            <ClientSideEvents SelectedIndexChanged="function(s, e) {
	                                            txtAffiliateName.SetText(cboAffiliateCode.GetSelectedItem().GetColumnText(1));
	                                            lblErrMsg.SetText('');
                                                var pAffiliate = cboAffiliateCode.GetText();
                                                var pForwarder = cboForwarder.GetText();
                                                cboShippingNo.PerformCallback(pAffiliate + '|' + pForwarder);
                                            }" />
                                        </dx:ASPxComboBox>
                                    </td>
                                    <td align="left" height="20px" valign="middle" width="20%">
                                        <dx:ASPxTextBox ID="txtAffiliateName" runat="server" Width="100%" ClientInstanceName="txtAffiliateName"
                                            BackColor="#CCCCCC" ReadOnly="True">
                                        </dx:ASPxTextBox>
                                    </td>
                                </tr>
                                <tr>
                                    <td height="20px">
                                        <dx:ASPxLabel ID="ASPxLabel15" runat="server" Text="FORWARDER*">
                                        </dx:ASPxLabel>
                                    </td>
                                    <td height="20px">
                                        &nbsp;
                                    </td>
                                    <td height="20px">
                                        <dx:ASPxComboBox ID="cboForwarder" runat="server" Width="100px" ClientInstanceName="cboForwarder"
                                            TextFormatString="{0}">
                                            <ClientSideEvents SelectedIndexChanged="function(s, e) {
	                                            txtForwarder.SetText(cboForwarder.GetSelectedItem().GetColumnText(1));
	                                            lblErrMsg.SetText('');
                                                var pAffiliate = cboAffiliateCode.GetText();
                                                var pForwarder = cboForwarder.GetText();
                                                cboShippingNo.PerformCallback(pAffiliate + '|' + pForwarder);
                                            }" />
                                        </dx:ASPxComboBox>
                                    </td>
                                    <td>
                                        <dx:ASPxTextBox ID="txtForwarder" runat="server" Width="100%" ClientInstanceName="txtForwarder"
                                            BackColor="#CCCCCC" ReadOnly="True">
                                        </dx:ASPxTextBox>
                                    </td>
                                </tr>
                                <tr>
                                    <td height="20px">
                                        <dx:ASPxLabel ID="ASPxLabel16" runat="server" Text="SHIPPING INSTRUCTION NO.(INVOICE NO.)*">
                                        </dx:ASPxLabel>
                                    </td>
                                    <td height="20px">
                                        &nbsp;
                                    </td>
                                    <td height="20px">
                                        <dx:ASPxComboBox ID="cboShippingNo" runat="server" Width="100px" ClientInstanceName="cboShippingNo"
                                            TextFormatString="{0}" DropDownStyle="DropDown">
                                        </dx:ASPxComboBox>
                                    </td>
                                    <td>
                                        <dx:ASPxTextBox ID="txtExcel" runat="server" ClientInstanceName="txtExcel" ForeColor="White"
                                            Width="100px">
                                            <Border BorderStyle="None" />
                                            <Border BorderStyle="None"></Border>
                                        </dx:ASPxTextBox>
                                    </td>
                                </tr>
                                <tr>
                                    <td height="20px">
                                        <dx:ASPxLabel ID="ASPxLabel17" runat="server" Text="SHIPPING INSTRUCTION DATE*">
                                        </dx:ASPxLabel>
                                    </td>
                                    <td height="20px">
                                        &nbsp;
                                    </td>
                                    <td height="20px">
                                        <dx:ASPxDateEdit ID="dtShippingDate" runat="server" Width="100px" ClientInstanceName="dtShippingDate"
                                            EditFormatString="dd MMM yyyy" DisplayFormatString="dd MMM yyyy">
                                        </dx:ASPxDateEdit>
                                    </td>
                                    <td height="20px">
                                        <dx:ASPxTextBox ID="txtTally" runat="server" ClientInstanceName="txtTally" ForeColor="White"
                                            Width="100px">
                                            <Border BorderStyle="None" />
                                            <Border BorderStyle="None"></Border>
                                        </dx:ASPxTextBox>
                                    </td>
                                </tr>
                                <tr>
                                    <td class="style2">
                                        <dx:ASPxLabel ID="ASPxLabel1" runat="server" Text="AWB, B/L NO" Width="250px">
                                        </dx:ASPxLabel>
                                    </td>
                                    <td>
                                        &nbsp;
                                    </td>
                                    <td class="style2">
                                        <dx:ASPxTextBox ID="txtBLNo" runat="server" Width="100px" ClientInstanceName="txtBLNo">
                                        </dx:ASPxTextBox>
                                    </td>
                                    <td class="style2">
                                        <table>
                                            <tr>
                                                <td align="center">
                                                    <dx:ASPxLabel ID="ASPxLabel2" runat="server" Text="B/L DATE" Width="80px">
                                                    </dx:ASPxLabel>
                                                </td>
                                                <td width="130">
                                                    <dx:ASPxDateEdit ID="dtBLDate" runat="server" Width="100px" ClientInstanceName="dtBLDate"
                                                        EditFormatString="dd MMM yyyy" DisplayFormatString="dd MMM yyyy">
                                                    </dx:ASPxDateEdit>
                                                </td>
                                                <td>
                                                    <dx:ASPxTextBox ID="txtterm" runat="server" ClientInstanceName="txtterm" Width="100px"
                                                        ForeColor="White">
                                                        <Border BorderColor="White" />
<Border BorderColor="White"></Border>
                                                    </dx:ASPxTextBox>
                                                </td>
                                                <td>
                                                    &nbsp;
                                                </td>
                                            </tr>
                                        </table>
                                    </td>
                                </tr>
                                <tr>                                    
                                    <td class="style2" colspan="4">
                                        <table style="width: 100%;">
                                            <tr>
                                                <td>
                                                    <dx:ASPxLabel ID="ASPxLabel3" runat="server" Text="VIA" Width="100px">
                                                    </dx:ASPxLabel>
                                                </td>
                                                <td>
                                                    <dx:ASPxTextBox ID="txtVia" runat="server" Width="100px" 
                                                        ClientInstanceName="txtVia" MaxLength="20">
                                                    </dx:ASPxTextBox>
                                                </td>    
                                                <td>
                                                    <dx:ASPxLabel ID="ASPxLabel4" runat="server" Text="MEASUREMENT" Width="100px">
                                                    </dx:ASPxLabel>
                                                </td>
                                                <td>
                                                    <dx:ASPxTextBox ID="txtmeasurement" runat="server" Width="100px" ClientInstanceName="txtmeasurement" 
                                                        DisplayFormatString="{0:n3}" onkeypress="return numbersonly(event)" HorizontalAlign="Right" >
                                                    </dx:ASPxTextBox>
                                                </td> 
                                                <td>
                                                    <dx:ASPxLabel ID="ASPxLabel5" runat="server" Text="TOTAL PALLET" Width="100px">
                                                    </dx:ASPxLabel>
                                                </td>
                                                <td>
                                                    <dx:ASPxTextBox ID="txtTotalPallet" runat="server" Width="100px" ClientInstanceName="txtTotalPallet" 
                                                        onkeypress="return numbersonly(event)" HorizontalAlign="Right" >
                                                    </dx:ASPxTextBox>
                                                </td> 
                                                <td>
                                                    <dx:ASPxLabel ID="ASPxLabel7" runat="server" Text="GROSS WEIGHT" Width="100px">
                                                    </dx:ASPxLabel>
                                                </td>
                                                <td>
                                                    <dx:ASPxTextBox ID="txtGrossWeight" runat="server" Width="100px" ClientInstanceName="txtGrossWeight" 
                                                        DisplayFormatString="{0:n2}" onkeypress="return numbersonly(event)" HorizontalAlign="Right" >
                                                    </dx:ASPxTextBox>
                                                </td>                                             
                                            </tr>
                                        </table>
                                    </td>
                                </tr>
                                <tr>
                                    <td class="style2" colspan="4">
                                        <table style="width: 100%;">
                                            <tr>
                                                <td>
                                                    <dx:ASPxLabel ID="ASPxLabel25" runat="server" Text="ETD VENDOR" Width="100px">
                                                    </dx:ASPxLabel>
                                                </td>
                                                <td>
                                                    <dx:ASPxDateEdit ID="etdvendor" runat="server" ClientInstanceName="etdvendor" DisplayFormatString="dd MMM yyyy"
                                                        EditFormat="Custom" EditFormatString="dd MMM yyyy" Width="100px">
                                                    </dx:ASPxDateEdit>
                                                </td>
                                                <td>
                                                    <dx:ASPxLabel ID="ASPxLabel29" runat="server" Text="ETD PORT" Width="100px">
                                                    </dx:ASPxLabel>
                                                </td>
                                                <td>
                                                    <dx:ASPxDateEdit ID="etdport" runat="server" ClientInstanceName="etdport" DisplayFormatString="dd MMM yyyy"
                                                        EditFormat="Custom" EditFormatString="dd MMM yyyy" Width="100px">
                                                    </dx:ASPxDateEdit>
                                                </td>
                                                <td>
                                                    <dx:ASPxLabel ID="ASPxLabel28" runat="server" Text="ETA PORT" Width="100px">
                                                    </dx:ASPxLabel>
                                                </td>
                                                <td>
                                                    <dx:ASPxDateEdit ID="etaport" runat="server" ClientInstanceName="etaport" DisplayFormatString="dd MMM yyyy"
                                                        EditFormat="Custom" EditFormatString="dd MMM yyyy" Width="100px">
                                                    </dx:ASPxDateEdit>
                                                </td>
                                                <td>
                                                    <dx:ASPxLabel ID="ASPxLabel26" runat="server" Text="ETA FACTORY" Width="100px">
                                                    </dx:ASPxLabel>
                                                </td>
                                                <td>
                                                    <dx:ASPxDateEdit ID="etafactory" runat="server" ClientInstanceName="etafactory" DisplayFormatString="dd MMM yyyy"
                                                        EditFormat="Custom" EditFormatString="dd MMM yyyy" Width="100px">
                                                    </dx:ASPxDateEdit>
                                                </td>
                                            </tr>
                                        </table>
                                    </td>
                                </tr>
                                <tr>                                                                        
                                    <td class="style2" colspan="4">
                                        <table style="width: 100%;">
                                            <tr>
                                                <td>
                                                    <dx:ASPxLabel ID="ASPxLabel32" runat="server" Text="TERM OF DELIVERY" Width="120px">
                                                    </dx:ASPxLabel>
                                                </td>
                                                <td>
                                                    <dx:ASPxComboBox ID="cboterm" runat="server" ClientInstanceName="cboterm" DropDownStyle="DropDown"
                                                        TextFormatString="{0}" Width="120px">
                                                        <ClientSideEvents SelectedIndexChanged="function(s, e) {
                                                            lblErrMsg.SetText('');
	                                                        txtterm.SetText(cboterm.GetSelectedItem().GetColumnText(1));
                                                            grid.PerformCallback('load');
                                                        }" />
                                                    </dx:ASPxComboBox>
                                                </td>
                                                <td>
                                                    <dx:ASPxLabel ID="ASPxLabel6" runat="server" Text="TERM OF SERVICE" Width="120px">
                                                    </dx:ASPxLabel>
                                                </td>
                                                <td>
                                                    <dx:ASPxComboBox ID="cboService" runat="server" ClientInstanceName="cboService" DropDownStyle="DropDown"
                                                        TextFormatString="{0}" Width="120px">
                                                        <ClientSideEvents SelectedIndexChanged="function(s, e) {
                                                            lblErrMsg.SetText('');
                                                            grid.PerformCallback('load');
                                                        }" />
                                                    </dx:ASPxComboBox>
                                                </td>
                                                <td>
                                                    <dx:ASPxLabel ID="ASPxLabel31" runat="server" Text="FREIGHT" Width="120px">
                                                    </dx:ASPxLabel>
                                                </td>
                                                <td>
                                                    <dx:ASPxComboBox ID="cbofreight" runat="server" ClientInstanceName="cbofreight" DropDownStyle="DropDown"
                                                        TextFormatString="{0}" Width="120px">
                                                    </dx:ASPxComboBox>
                                                </td>
                                                <td>
                                                    &nbsp;</td>
                                                <td>
                                                    &nbsp;</td>
                                            </tr>
                                        </table>
                                    </td>
                                </tr>
                                <tr>                                                                        
                                    <td class="style2" colspan="4">
                                        <table style="width: 100%;">
                                            <tr>
                                                <td>
                                                    <dx:ASPxLabel ID="ASPxLabel8" runat="server" Text="SHIPPING LINE" Width="120px">
                                                    </dx:ASPxLabel>
                                                </td>
                                                <td>
                                                    <dx:ASPxTextBox ID="txtShippingLine" runat="server" Width="100px" 
                                                        ClientInstanceName="txtShippingLine" MaxLength="20">
                                                    </dx:ASPxTextBox>
                                                </td>
                                                <td>
                                                    <dx:ASPxLabel ID="ASPxLabel9" runat="server" Text="VESSEL" Width="120px">
                                                    </dx:ASPxLabel>
                                                </td>
                                                <td>
                                                    <dx:ASPxTextBox ID="txtVessel" runat="server" Width="100px" 
                                                        ClientInstanceName="txtVessel" MaxLength="20">
                                                    </dx:ASPxTextBox>
                                                </td>
                                                <td>
                                                    <dx:ASPxLabel ID="ASPxLabel10" runat="server" Text="VOYAGE" Width="120px">
                                                    </dx:ASPxLabel>
                                                </td>
                                                <td>
                                                    <dx:ASPxTextBox ID="txtVoyage" runat="server" Width="100px" 
                                                        ClientInstanceName="txtVoyage" MaxLength="20">
                                                    </dx:ASPxTextBox>
                                                </td>
                                                <td>
                                                    &nbsp;</td>
                                                <td>
                                                    &nbsp;</td>
                                            </tr>
                                            <tr>
                                                <td>
                                                    <dx:ASPxLabel ID="ASPxLabel11" runat="server" Text="Remarks" Width="120px">
                                                    </dx:ASPxLabel>
                                                </td>
                                                <td colspan = "7">
                                                    <dx:ASPxTextBox ID="txtRemarks" runat="server" Width="1000px" 
                                                        ClientInstanceName="txtRemarks" MaxLength="1100">
                                                    </dx:ASPxTextBox>
                                                </td>
                                            </tr>
                                        </table>
                                    </td>
                                </tr>
                            </table>
                        </dx:PanelContent>
                    </PanelCollection>
                </dx:ASPxRoundPanel>
            </td>
            <td align="left" style="width: 70%">
                <dx:ASPxRoundPanel ID="ASPxRoundPanel2" runat="server" HeaderText="FILTER FORWARDER RECEIVING"
                    ShowCollapseButton="true" View="GroupBox" Width="100%" Height="200px">
                    <ContentPaddings PaddingLeft="5px" PaddingRight="5px" />
                    <ContentPaddings PaddingLeft="5px" PaddingRight="5px"></ContentPaddings>
                    <PanelCollection>
                        <dx:PanelContent ID="PanelContent2" runat="server">
                            <table id="Table2">
                                <tr>
                                    <td height="20px">
                                        <dx:ASPxLabel ID="ASPxLabel19" runat="server" Text="SUPPLIER CODE/NAME" Width="125px">
                                        </dx:ASPxLabel>
                                    </td>
                                    <td height="20px">
                                        <dx:ASPxComboBox ID="cboSupplierCode" Height="20px" runat="server" Width="100px"
                                            ClientInstanceName="cboSupplierCode" TextFormatString="{0}">
                                            <ClientSideEvents SelectedIndexChanged="function(s, e) {
	txtSupplierName.SetText(cboSupplierCode.GetSelectedItem().GetColumnText(1));
	lblErrMsg.SetText('');
}" />
                                        </dx:ASPxComboBox>
                                    </td>
                                    <td height="20px">
                                        <dx:ASPxTextBox ID="txtSupplierName" runat="server" Width="160px" ClientInstanceName="txtSupplierName"
                                            BackColor="#CCCCCC" ReadOnly="True">
                                        </dx:ASPxTextBox>
                                    </td>
                                </tr>
                                <tr>
                                    <td height="20px">
                                        <dx:ASPxLabel ID="ASPxLabel20" runat="server" Text="PART CODE/NAME">
                                        </dx:ASPxLabel>
                                    </td>
                                    <td height="20px">
                                        <dx:ASPxComboBox ID="cboPartNo" runat="server" Width="100px" ClientInstanceName="cboPartNo"
                                            TextFormatString="{0}">
                                            <ClientSideEvents SelectedIndexChanged="function(s, e) {
	txtPartName.SetText(cboPartNo.GetSelectedItem().GetColumnText(1));
	lblErrMsg.SetText('');
}" />
                                        </dx:ASPxComboBox>
                                    </td>
                                    <td height="20px">
                                        <dx:ASPxTextBox ID="txtPartName" runat="server" Width="160px" ClientInstanceName="txtPartName"
                                            BackColor="#CCCCCC" ReadOnly="True">
                                        </dx:ASPxTextBox>
                                    </td>
                                </tr>
                                <tr>
                                    <td height="20px">
                                        <dx:ASPxLabel ID="ASPxLabel21" runat="server" Text="ORDER NO">
                                        </dx:ASPxLabel>
                                    </td>
                                    <td height="20px">
                                        <dx:ASPxTextBox ID="txtOrderNo" runat="server" Width="100px" ClientInstanceName="txtOrderNo">
                                        </dx:ASPxTextBox>
                                    </td>
                                    <td height="20px">
                                        &nbsp;
                                    </td>
                                </tr>
                                <tr>
                                    <td height="20px">
                                        <dx:ASPxLabel ID="ASPxLabel24" runat="server" Text="STATUS SHIPPING">
                                        </dx:ASPxLabel>
                                    </td>
                                    <td height="20px">
                                        <dx:ASPxTextBox ID="txtSend" runat="server" BackColor="#CCCCCC" ClientInstanceName="txtSend"
                                            Width="175px" Font-Bold="True" ReadOnly="True">
                                        </dx:ASPxTextBox>
                                    </td>
                                    <td height="20px">
                                        &nbsp;
                                    </td>
                                </tr>
                                <tr>
                                    <td height="20px">
                                        &nbsp;
                                    </td>
                                    <td height="20px">
                                        &nbsp;
                                    </td>
                                    <td height="20px">
                                        &nbsp;
                                    </td>
                                </tr>
                                <tr>
                                    <td height="20px">
                                        &nbsp;
                                    </td>
                                    <td height="20px">
                                        &nbsp;
                                    </td>
                                    <td height="20px">
                                        &nbsp;
                                    </td>
                                </tr>
                                <tr>
                                    <td class="style4">
                                    </td>
                                    <td class="style4">
                                    </td>
                                    <td class="style4">
                                        <table style="width: 100%;">
                                            <tr>
                                                <td>
                                                    <dx:ASPxButton ID="btnClearAll" runat="server" AutoPostBack="False" ClientInstanceName="btnClearAll"
                                                        Text="CLEAR" Width="70px">
                                                        <ClientSideEvents Click="function(s, e) {
cboAffiliateCode.SetText('');
txtAffiliateName.SetText('');
cboForwarder.SetText('');
txtForwarder.SetText('');
cboShippingNo.SetText('');
dtShippingDate.SetDate(new Date());
txtBLNo.SetText('');
dtBLDate.SetDate(new Date());
cboSupplierCode.SetText('==ALL==');
txtSupplierName.SetText('==ALL==');
cboPartNo.SetText('==ALL==');
txtPartName.SetText('==ALL==');
txtOrderNo.SetText('');
txtSend.SetText('');
lblErrMsg.SetText('');
cboCreate.SetText('CREATE');
}" />
                                                    </dx:ASPxButton>
                                                </td>
                                                <td>
                                                    <dx:ASPxButton ID="btnSearch" runat="server" AutoPostBack="False" ClientInstanceName="btnSearch"
                                                        Text="SEARCH" Width="70px">                                                        
                                                    </dx:ASPxButton>
                                                </td>
                                            </tr>
                                        </table>
                                    </td>
                                </tr>
                            </table>
                        </dx:PanelContent>
                    </PanelCollection>
                </dx:ASPxRoundPanel>
            </td>
        </tr>
    </table>
    <div style="height: 8px;">
    </div>
    <table align="left" width="100%">
        <tr align="left">
            <td width="100%" height="16px" style="border-top-style: solid; border-top-width: thin;
                border-top-color: #808080; border-bottom-style: solid; border-bottom-width: thin;
                border-bottom-color: #808080" align="left">
                <dx:ASPxLabel ID="lblErrMsg" runat="server" Font-Names="Tahoma" Font-Size="8pt" Text="ERROR MESSAGE"
                    Width="100%" ClientInstanceName="lblErrMsg" Height="16px">
                </dx:ASPxLabel>
            </td>
        </tr>
    </table>
    <br />
    <br />
    <table style="width: 100%;">
        <tr>
            <td align="left" class="style3" colspan="12">
                <dx:ASPxGridView ID="grid" runat="server" AutoGenerateColumns="False" Width="100%"
                    KeyFieldName="RowNo;AffiliateID;OrderNo;PartNo;SupplierID;SuratJalanNo;LabelNo"
                    ClientInstanceName="grid">
                    <ClientSideEvents EndCallback="function(s, e) {                  
                                        
						var pMsg = s.cpMessage;
                        if (pMsg != '') {
                            if (pMsg.substring(1,5) == '1001' || pMsg.substring(1,5) == '1002' || pMsg.substring(1,5) == '1003' || pMsg.substring(1,5) == '2001' || pMsg.substring(1,5) == '2005') {
                                lblErrMsg.GetMainElement().style.color = 'Blue';
                            } else {
                                lblErrMsg.GetMainElement().style.color = 'Red';
                            }
                            lblErrMsg.SetText(pMsg);
                        } else {
                            lblErrMsg.SetText('');
                        }
                        
                        if(s.cpButton == '0'){            
                            btnSendEDI.SetEnabled(true);
                        }else if (s.cpButton == '1'){                
                            btnSendEDI.SetEnabled(false);
                        }else {
                            btnSendEDI.SetEnabled(true);
                        }

                        delete s.cpButton;

                        if (pMsg.substring(1,5) == '1001') {
	                        cboCreate.SetText('UPDATE');
                        }
                        
                    var pSending = s.cpSending;
                    if (pSending == 'ALREADY SEND') {
	                    txtSending.SetText('ALREADY SEND');
                    }
                                            delete s.cpMessage;
                    }" RowClick="function(s, e) {
	                    lblErrMsg.SetText('');
                    }" CallbackError="function(s, e) {
	                    e.handled = true;
                    }" Init="OnInit" />
                    <Columns>
                        <dx:GridViewDataTextColumn Caption=" " FieldName="RowNo" Name="RowNo" VisibleIndex="0"
                            Width="30px">
                            <HeaderStyle Font-Names="Tahoma" HorizontalAlign="Center" VerticalAlign="Middle"
                                Wrap="True" Font-Size="8pt" />
                            <CellStyle Font-Names="Tahoma" HorizontalAlign="Center" VerticalAlign="Middle" Font-Size="8pt">
                            </CellStyle>
                        </dx:GridViewDataTextColumn>
                        <dx:GridViewDataCheckColumn FieldName="Act" Name="Act" VisibleIndex="1" Width="40px"
                            Caption=" ">
                            <HeaderStyle HorizontalAlign="Center" VerticalAlign="Middle" />
                            <HeaderCaptionTemplate>
                                <dx:ASPxCheckBox ID="chkAll" runat="server" ClientInstanceName="chkAll" ClientSideEvents-CheckedChanged="OnAllCheckedChanged"
                                    ValueType="System.String" ValueChecked="1" ValueUnchecked="0">
                                </dx:ASPxCheckBox>
                            </HeaderCaptionTemplate>
                            <PropertiesCheckEdit ValueChecked="1" ValueType="System.Int32" ValueUnchecked="0">
                            </PropertiesCheckEdit>
                        </dx:GridViewDataCheckColumn>
                        <dx:GridViewDataDateColumn Caption="ETD PORT" FieldName="ETDPort" Name="ETDPort"
                            VisibleIndex="2">
                            <PropertiesDateEdit DisplayFormatString="dd MMM yyyy" Spacing="0">
                            </PropertiesDateEdit>
                            <HeaderStyle Font-Names="Tahoma" Font-Size="8pt" Font-Underline="False" Wrap="True"
                                HorizontalAlign="Center" />
                            <CellStyle Font-Names="Verdana" Font-Size="8pt">
                            </CellStyle>
                        </dx:GridViewDataDateColumn>
                        <dx:GridViewDataTextColumn Caption="ORDER NO" FieldName="OrderNo" Name="OrderNo"
                            VisibleIndex="3" Width="70px">
                            <HeaderStyle Font-Names="Tahoma" Font-Size="8pt" Font-Underline="False" Wrap="True"
                                HorizontalAlign="Center" />
                            <CellStyle Font-Names="Tahoma" Font-Size="8pt">
                            </CellStyle>
                        </dx:GridViewDataTextColumn>
                        <dx:GridViewDataTextColumn Caption="PART NO" FieldName="PartNo" Name="PartNo" VisibleIndex="4"
                            Width="120px">
                            <HeaderStyle Font-Names="Tahoma" Font-Size="8pt" Font-Underline="False" Wrap="True"
                                HorizontalAlign="Center" />
                            <CellStyle Font-Names="Tahoma" Font-Size="8pt">
                            </CellStyle>
                        </dx:GridViewDataTextColumn>
                        <dx:GridViewDataTextColumn Caption="PART NAME" FieldName="PartName" Name="PartName"
                            VisibleIndex="5" Width="150px">
                            <HeaderStyle Font-Names="Tahoma" Font-Size="8pt" Font-Underline="False" Wrap="True"
                                HorizontalAlign="Center" />
                            <CellStyle Font-Names="Tahoma" Font-Size="8pt">
                            </CellStyle>
                        </dx:GridViewDataTextColumn>
                        <dx:GridViewDataTextColumn Caption="SJ NO" FieldName="SuratJalanNo" Name="SuratJalanNo"
                            VisibleIndex="7" Width="120px">
                            <HeaderStyle Font-Names="Tahoma" Font-Size="8pt" Font-Underline="False" Wrap="True"
                                HorizontalAlign="Center" />
                            <CellStyle Font-Names="Tahoma" Font-Size="8pt">
                            </CellStyle>
                        </dx:GridViewDataTextColumn>
                        <dx:GridViewDataTextColumn Caption="UOM" FieldName="UnitClsDesc" Name="UnitClsDesc"
                            VisibleIndex="8" Width="55px">
                            <HeaderStyle Font-Names="Tahoma" HorizontalAlign="Center" VerticalAlign="Middle"
                                Wrap="True" Font-Size="8pt" />
                            <CellStyle Font-Names="Tahoma" Font-Size="8pt" HorizontalAlign="Center">
                            </CellStyle>
                        </dx:GridViewDataTextColumn>
                         <dx:GridViewDataTextColumn Caption="PRICE" FieldName="Price" Name="Price"
                            VisibleIndex="8" Width="70px">
                            <PropertiesTextEdit DisplayFormatString="{0:n4}">
                                <MaskSettings IncludeLiterals="DecimalSymbol" />
                            </PropertiesTextEdit>
                            <HeaderStyle Font-Names="Tahoma" HorizontalAlign="Center" VerticalAlign="Middle"
                                Wrap="True" Font-Size="8pt" />
                            <CellStyle Font-Names="Tahoma" Font-Size="8pt" HorizontalAlign="Center">
                            </CellStyle>
                        </dx:GridViewDataTextColumn>
                        <dx:GridViewDataTextColumn Caption="QTY/ BOX" FieldName="QtyBox" Name="QtyBox" VisibleIndex="9"
                            Width="50px">
                            <PropertiesTextEdit DisplayFormatString="{0:n0}">
                                <MaskSettings IncludeLiterals="DecimalSymbol" />
                            </PropertiesTextEdit>
                            <HeaderStyle Font-Names="Tahoma" HorizontalAlign="Center" VerticalAlign="Middle"
                                Wrap="True" Font-Size="8pt" />
                            <CellStyle Font-Names="Tahoma" Font-Size="8pt" HorizontalAlign="Right">
                            </CellStyle>
                        </dx:GridViewDataTextColumn>
                        <dx:GridViewDataTextColumn Caption="GOOD RECEIVING QTY" FieldName="GoodRecQty" Name="GoodRecQty"
                            VisibleIndex="10" Width="70px">
                            <PropertiesTextEdit DisplayFormatString="{0:n0}">
                                <MaskSettings IncludeLiterals="DecimalSymbol" />
                            </PropertiesTextEdit>
                            <HeaderStyle Font-Names="Tahoma" Font-Size="8pt" Font-Underline="False" Wrap="True"
                                HorizontalAlign="Center" />
                            <CellStyle Font-Names="Tahoma" Font-Size="8pt" HorizontalAlign="Right">
                            </CellStyle>
                        </dx:GridViewDataTextColumn>
                        <dx:GridViewDataTextColumn Caption="SHIPPING QTY" FieldName="ShippingQty" Name="ShippingQty"
                            VisibleIndex="11" Width="70px">
                            <PropertiesTextEdit DisplayFormatString="{0:n0}">
                                <MaskSettings IncludeLiterals="DecimalSymbol" />
                            </PropertiesTextEdit>
                            <HeaderStyle Font-Names="Tahoma" Font-Size="8pt" Font-Underline="False" Wrap="True"
                                HorizontalAlign="Center" />
                            <CellStyle Font-Names="Tahoma" Font-Size="8pt" HorizontalAlign="Right">
                            </CellStyle>
                        </dx:GridViewDataTextColumn>
                        <dx:GridViewDataTextColumn Caption="BOX QTY" FieldName="BoxQty" Name="BoxQty" VisibleIndex="12"
                            Width="50px">
                            <PropertiesTextEdit DisplayFormatString="{0:n0}">
                                <MaskSettings IncludeLiterals="DecimalSymbol" />
                            </PropertiesTextEdit>
                            <HeaderStyle Font-Names="Tahoma" Font-Size="8pt" Font-Underline="False" Wrap="True"
                                HorizontalAlign="Center" />
                            <CellStyle Font-Names="Tahoma" Font-Size="8pt" HorizontalAlign="Right">
                            </CellStyle>
                        </dx:GridViewDataTextColumn>
                        <dx:GridViewDataTextColumn Caption="SUPPLIER CODE" FieldName="SupplierID" Name="SupplierID"
                            VisibleIndex="13" Width="80px">
                            <HeaderStyle Font-Names="Tahoma" Font-Size="8pt" Font-Underline="False" Wrap="True"
                                HorizontalAlign="Center" />
                            <CellStyle Font-Names="Tahoma" Font-Size="8pt">
                            </CellStyle>
                        </dx:GridViewDataTextColumn>
                        <%--<dx:GridViewDataTextColumn Caption="SUPPLIER NAME" FieldName="SupplierName" Name="SupplierName"
                            VisibleIndex="14" Width="170px">
                            <HeaderStyle Font-Names="Tahoma" Font-Size="8pt" Font-Underline="False" Wrap="True"
                                HorizontalAlign="Center" />
                            <CellStyle Font-Names="Tahoma" Font-Size="8pt">
                            </CellStyle>
                        </dx:GridViewDataTextColumn>--%>
                        <dx:GridViewDataTextColumn Caption="AdaData" FieldName="AdaData" Name="AdaData" VisibleIndex="15"
                            Width="0px">
                            <HeaderStyle Font-Names="Tahoma" Font-Size="8pt" Font-Underline="False" Wrap="True"
                                HorizontalAlign="Center" />
                            <CellStyle Font-Names="Tahoma" Font-Size="8pt">
                            </CellStyle>
                        </dx:GridViewDataTextColumn>
                        <dx:GridViewDataTextColumn Caption="UOM" FieldName="UnitCls" Name="UnitCls" VisibleIndex="16"
                            Width="0px">
                            <HeaderStyle Font-Names="Tahoma" HorizontalAlign="Center" VerticalAlign="Middle"
                                Wrap="True" Font-Size="8pt" />
                            <CellStyle Font-Names="Tahoma" Font-Size="8pt" HorizontalAlign="Center">
                            </CellStyle>
                        </dx:GridViewDataTextColumn>
                        <dx:GridViewDataTextColumn Caption="BOX NO." FieldName="LabelNo" VisibleIndex="6">
                            <HeaderStyle HorizontalAlign="Center" VerticalAlign="Middle" />
                            <CellStyle Font-Names="Tahoma" Font-Size="8pt">
                            </CellStyle>
                        </dx:GridViewDataTextColumn>
                    </Columns>
                    <SettingsPager Visible="False" Position="Top">
                        <Summary Text="Page {0} of {1} [{2} record(s)]" AllPagesText="Page {0} of {1} " />
                        <Summary AllPagesText="Page {0} of {1} " Text="Page {0} of {1} [{2} record(s)]">
                        </Summary>
                    </SettingsPager>
                    <SettingsEditing Mode="Batch">
                    </SettingsEditing>
                    <Settings ShowHorizontalScrollBar="True" ShowVerticalScrollBar="True" ShowGroupButtons="False"
                        ShowStatusBar="Hidden" VerticalScrollableHeight="250" />
                    <Settings ShowGroupButtons="False" ShowVerticalScrollBar="True" ShowHorizontalScrollBar="True"
                        VerticalScrollableHeight="250" ShowStatusBar="Hidden"></Settings>
                    <Styles>
                        <Header BackColor="#FFD2A6" Font-Names="Verdana" Font-Size="8pt">
                        </Header>
                        <Row BackColor="#FFFFE1" Font-Names="Verdana" Font-Size="8pt" Wrap="False">
                        </Row>
                        <RowHotTrack BackColor="#E8EFFD" Font-Names="Verdana" Font-Size="8pt" Wrap="False">
                        </RowHotTrack>
                        <SelectedRow Wrap="False">
                        </SelectedRow>
                    </Styles>
                    <StylesEditors ButtonEditCellSpacing="0">
                        <ProgressBar Height="21px">
                        </ProgressBar>
                    </StylesEditors>
                </dx:ASPxGridView>
                <br />
            </td>
        </tr>
        <tr>
            <td align="left">
                <dx:ASPxButton ID="btnsubmenu" runat="server" Text="SUB MENU" Width="90px" Font-Names="Tahoma"
                    Font-Size="8pt" ClientInstanceName="btnsubmenu">
                </dx:ASPxButton>
            </td>
            <td style="width: 90px;">
                <dx:aspxbutton id="btnImportEDI" runat="server" width="90px" font-names="Tahoma" font-size="8pt"
                    text="IMPORT ASN" clientinstancename="btnImportEDI" autopostback="False">
                    <ClientSideEvents Click="function(s, e) {                    
                    lblErrMsg.SetText('');                                
                    grid.PerformCallback('ImportEDI');                   
                    }" /> 
                 </dx:aspxbutton>
            </td>
            <td style="width: 90px;">
                <dx:aspxbutton id="btnSendEDI" runat="server" width="90px" font-names="Tahoma" font-size="8pt"
                    text="SEND E.D.I" clientinstancename="btnSendEDI" autopostback="False">
                    <ClientSideEvents Click="function(s, e) {                    
                    lblErrMsg.SetText('');                                
                    grid.PerformCallback('SendEDI');                   
                    }" /> 
                 </dx:aspxbutton>
            </td>
            <td align="left" width="90">
                <dx:ASPxButton ID="btnSave" runat="server" Text="SAVE" Width="90px" Font-Names="Tahoma"
                    Font-Size="8pt" AutoPostBack="False" ClientInstanceName="btnSave">
                    <ClientSideEvents Click="function(s, e) {
                    if (cboAffiliateCode.GetText() == '') {
                        lblErrMsg.GetMainElement().style.color = 'Red';
                        lblErrMsg.SetText('[6011] Please Select Affiliate ID first!');
                        cboAffiliateCode.Focus();
                        e.ProcessOnServer = false;
                        return false;
                    }
                    if (cboForwarder.GetText() == '') {
				        lblErrMsg.GetMainElement().style.color = 'Red';
                        lblErrMsg.SetText('[6011] Please Select Forwarder first!');
                        cboForwarder.Focus();
                        e.ProcessOnServer = false;
                        return false;
                    }
			        if (cboShippingNo.GetText() =='') {
				        lblErrMsg.GetMainElement().style.color = 'Red';
                        lblErrMsg.SetText('[6011] Please Input Shipping No first!');
                        cboShippingNo.Focus();
                        e.ProcessOnServer = false;
                        return false;
                    }
			        if (grid.GetVisibleRowsOnPage() == 0){
        		        lblErrMsg.GetMainElement().style.color = 'Red';
	    		        lblErrMsg.SetText('[6013] No data to submit!');
        		        e.processOnServer = false;
        		        return;
			        }
	            grid.UpdateEdit();

                var millisecondsToWait = 1000;

                setTimeout(function() {grid.PerformCallback('loadaftersubmit');
                    }, millisecondsToWait);	
                               
                    }" />
                </dx:ASPxButton>
            </td>
            <td align="left" width="90">
                <dx:ASPxButton ID="btnprintinvoice" runat="server" Text="PRINT INVOICE" Width="90px"
                    ClientInstanceName=" btnprintinvoice">
                </dx:ASPxButton>
            </td>
            <td align="left" width="90">
                <dx:ASPxButton ID="btnsend" runat="server" Text="SEND INVOICE" Width="90px" ClientInstanceName="btnsend"
                    AutoPostBack="False">
                    <ClientSideEvents Click="function(s, e) {
	if (cboAffiliateCode.GetText() == '') {
                lblErrMsg.GetMainElement().style.color = 'Red';
                lblErrMsg.SetText('[6011] Please Select Affiliate ID first!');
                cboAffiliateCode.Focus();
                e.ProcessOnServer = false;
                return false;
            }

            if (cboForwarder.GetText() == '') {
				lblErrMsg.GetMainElement().style.color = 'Red';
                lblErrMsg.SetText('[6011] Please Select Forwarder first!');
                cboForwarder.Focus();
                e.ProcessOnServer = false;
                return false;
            }

			if (cboShippingNo.GetText() == '') {
				lblErrMsg.GetMainElement().style.color = 'Red';
                lblErrMsg.SetText('[6011] Please Input Shipping No first!');
                cboShippingNo.Focus();
                e.ProcessOnServer = false;
                return false;
            }

			if (grid.GetVisibleRowsOnPage() == 0){
        		lblErrMsg.GetMainElement().style.color = 'Red';
	    		lblErrMsg.SetText('[6013] No data to submit!');
        		e.processOnServer = false;
        		return;
			}
	
    grid.UpdateEdit();
	grid.PerformCallback('SENDINV');
	cboCreate.SetText('UPDATE');
}" />
                </dx:ASPxButton>
            </td>
            <td align="left" width="90">
                <dx:ASPxButton ID="btnSendTally" runat="server" Text="SEND TALLY DATA (EXCEL)" Width="90px"
                    ClientInstanceName="btnSendTally" AutoPostBack="False">
                    <ClientSideEvents Click="function(s, e) {
	if (cboAffiliateCode.GetText() == '') {
                lblErrMsg.GetMainElement().style.color = 'Red';
                lblErrMsg.SetText('[6011] Please Select Affiliate ID first!');
                cboAffiliateCode.Focus();
                e.ProcessOnServer = false;
                return false;
            }

            if (cboForwarder.GetText() == '') {
				lblErrMsg.GetMainElement().style.color = 'Red';
                lblErrMsg.SetText('[6011] Please Select Forwarder first!');
                cboForwarder.Focus();
                e.ProcessOnServer = false;
                return false;
            }

			if (cboShippingNo.GetText() == '') {
				lblErrMsg.GetMainElement().style.color = 'Red';
                lblErrMsg.SetText('[6011] Please Input Shipping No first!');
                cboShippingNo.Focus();
                e.ProcessOnServer = false;
                return false;
            }

			if (grid.GetVisibleRowsOnPage() == 0){
        		lblErrMsg.GetMainElement().style.color = 'Red';
	    		lblErrMsg.SetText('[6013] No data to submit!');
        		e.processOnServer = false;
        		return;
			}
	
    grid.UpdateEdit();
	grid.PerformCallback('SENDTALLY');
	cboCreate.SetText('UPDATE');
}" />
                </dx:ASPxButton>
            </td>            
            <td align="right" width="50px">
                <dx:ASPxButton ID="btnPrintSI" runat="server" Text="PRINT SI" Width="90px" ClientInstanceName="btnPrintSI">
                </dx:ASPxButton>
            </td>
            <td align="right" width="50px">
                <dx:ASPxButton ID="btnPrintTally" runat="server" Text="PRINT TALLY" Width="90px"
                    ClientInstanceName="btnPrintTally">
                </dx:ASPxButton>
            </td>
            <td align="left" width="90">
                <dx:ASPxButton ID="btnDelete" runat="server" Text="DELETE" Width="90px" ClientInstanceName="btnDelete"
                    AutoPostBack="False" Visible="True">
                    <ClientSideEvents Click="function(s, e) {
                    if (cboShippingNo.GetText() == '') {                                   
                            lblerrmessage.SetText('[6011] Please Input Surat Jalan No first!');
		                    lblerrmessage.GetMainElement().style.color = 'Red';
                            return false;
                        }
                        var msg = confirm('Are you sure want to delete this data ?');
                        if (msg == false) {
                            e.processOnServer = false;
                            return;
                        }
	                    grid.PerformCallback('delete');
}" />
                </dx:ASPxButton>
            </td>
            <td align="left" width="90">
                <dx:ASPxButton ID="btnExcelDW" runat="server" Text="TO EXCEL" Width="90px" ClientInstanceName="btnExcelDW"
                    AutoPostBack="False" Visible="True">
                    <ClientSideEvents Click="function(s, e) {
	                    grid.PerformCallback('excel');
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
    <dx:ASPxHiddenField ID="HF" runat="server" ClientInstanceName="HF">
    </dx:ASPxHiddenField>
</asp:Content>
