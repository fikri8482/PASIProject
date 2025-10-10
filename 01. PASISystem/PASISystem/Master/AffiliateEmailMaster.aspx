<%@ Page Title="" Language="vb" AutoEventWireup="false" MasterPageFile="~/SiteContent.Master" CodeBehind="AffiliateEmailMaster.aspx.vb" Inherits="PASISystem.AffiliateEmailMaster" %>

<%@ Register Assembly="DevExpress.Web.v14.1, Version=14.1.6.0, Culture=neutral, PublicKeyToken=b88d1754d700e49a" Namespace="DevExpress.Web.ASPxCallback" TagPrefix="dx" %>
<%@ Register Assembly="DevExpress.Web.v14.1, Version=14.1.6.0, Culture=neutral, PublicKeyToken=b88d1754d700e49a" Namespace="DevExpress.Web.ASPxGridView.Export" TagPrefix="dx" %>
<%@ Register Assembly="DevExpress.Web.v14.1, Version=14.1.6.0, Culture=neutral, PublicKeyToken=b88d1754d700e49a" Namespace="DevExpress.Web.ASPxGridView" TagPrefix="dx" %>
<%@ Register assembly="DevExpress.Web.v14.1, Version=14.1.6.0, Culture=neutral, PublicKeyToken=b88d1754d700e49a" namespace="DevExpress.Web.ASPxEditors" tagprefix="dx" %>
<%@ Register Assembly="DevExpress.Web.v14.1, Version=14.1.6.0, Culture=neutral, PublicKeyToken=b88d1754d700e49a" Namespace="DevExpress.Web.ASPxPanel" TagPrefix="dx" %>
<%@ Register Assembly="DevExpress.Web.v14.1, Version=14.1.6.0, Culture=neutral, PublicKeyToken=b88d1754d700e49a" Namespace="DevExpress.Web.ASPxPopupControl" TagPrefix="dx" %>
<%@ Register Assembly="DevExpress.Web.v14.1, Version=14.1.6.0, Culture=neutral, PublicKeyToken=b88d1754d700e49a" namespace="DevExpress.Web.ASPxGlobalEvents" TagPrefix="dx" %>
<%@ Register assembly="DevExpress.Web.v14.1, Version=14.1.6.0, Culture=neutral, PublicKeyToken=b88d1754d700e49a" namespace="DevExpress.Web.ASPxHiddenField" tagprefix="dx1" %>
<%@ Register assembly="DevExpress.Web.v14.1, Version=14.1.6.0, Culture=neutral, PublicKeyToken=b88d1754d700e49a" namespace="DevExpress.Web.ASPxGridView.Export" tagprefix="dx2" %>
<asp:Content ID="Content1" ContentPlaceHolderID="head" runat="server">
<style type="text/css">
.dxeHLC, .dxeHC, .dxeHFC
{
display: none;
}
    .style1
    {
        width: 188px;
    }
    .style3
    {
        width: 56px;
    }
    .style4
    {
        width: 148px;
    }
    .style5
    {
        width: 200px;
    }
    .style6
    {
        width: 70px;
    }
</style> 

<script language="javascript" type="text/javascript">


    function singlequote(e) {
        var unicode = e.charCode ? e.charCode : e.keyCode
        if (unicode == 39) {
            return false //disable key press
        }
    }

    function numbersonly(e) {
        var unicode = e.charCode ? e.charCode : e.keyCode
        if (unicode != 8) { //if the key isn't the backspace key (which we should allow)
            if (unicode < 45 || unicode > 57) //if not a number
                return false //disable key press
        }
    }

    function clear() {
        cboAffiliateCode.SetText('');
        txtAffiliateName.SetText('');
        txtEmailToAffiliatePO.SetText('');
        txtEmailCCAffiliatePO.SetText('');
        txtEmailToAffiliatePORevision.SetText('');
        txtEmailCCAffiliatePORevision.SetText('');
        txtEmailToKanban.SetText('');
        txtEmailCCKanban.SetText('');
        txtEmailToSupplierDelivery.SetText('');
        txtEmailCCSupplierDelivery.SetText('');
        txtEmailToPASIReceiving.SetText('');
        txtEmailCCPASIReceiving.SetText('');
        txtEmailToAffiliateReceiving.SetText('');
        txtEmailCCAffiliateReceiving.SetText('');
        txtEmailToGoodReceive.SetText('');
        txtEmailCCGoodReceive.SetText('');
        txtEmailToInvoice.SetText('');
        txtEmailCCInvoice.SetText('');
        txtEmailToSummaryOutstanding.SetText('');
        txtEmailCCSummaryOutstanding.SetText('');

        //txtAffiliateID.GetInputElement().setAttribute('style', 'background-color:#FFFFFF;');
        txtAffiliateName.GetInputElement().setAttribute('style', 'background:#FFFFFF;foreground:#FFFFFF;');
        txtAffiliateName.GetInputElement().readOnly = false;
    }

//    function clear2() {
//        txtSupplierCode.GetInputElement().setAttribute('style', 'background:#CCCCCC;foreground:#CCCCCC;');
//        txtSupplierCode.GetInputElement().readOnly = false;
//    }

    function validasubmit() {
        lblInfo.GetMainElement().style.color = 'Red';
        if (cboAffiliateCode.GetText() == "") {
            lblInfo.SetText("[6011] Please Input Affiliate Code first!");
            cboAffiliateCode.Focus();
            e.ProcessOnServer = false;
            return false;
        }
             
    }

    function up_Insert() {
        var pIsUpdate = '';
        var pAffiliateID = cboAffiliateCode.GetText();
        var pAffiliatePOTO = txtEmailToAffiliatePO.GetText();
        var pAffiliatePOCC = txtEmailCCAffiliatePO.GetText();
        var pAffiliatePORevisionTO = txtEmailToAffiliatePORevision.GetText();
        var pAffiliatePORevisionCC = txtEmailCCAffiliatePORevision.GetText();
        var pKanbanTO = txtEmailToKanban.GetText();
        var pKanbanCC = txtEmailCCKanban.GetText();
        var pSupplierDeliveryTO = txtEmailToSupplierDelivery.GetText();
        var pSupplierDeliveryCC = txtEmailCCSupplierDelivery.GetText();
        var pPASIReceivingTO = txtEmailToPASIReceiving.GetText();
        var pPASIReceivingCC = txtEmailCCPASIReceiving.GetText();
        var pAffiliateReceivingTO = txtEmailToAffiliateReceiving.GetText();
        var pAffiliateReceivingCC = txtEmailCCAffiliateReceiving.GetText();
        var pGoodReceiveTO = txtEmailToGoodReceive.GetText();
        var pGoodReceiveCC = txtEmailCCGoodReceive.GetText();
        var pInvoiceTO = txtEmailToInvoice.GetText();
        var pInvoiceCC = txtEmailCCInvoice.GetText();
        var pSummaryOutstandingTO = txtEmailToSummaryOutstanding.GetText();
        var pSummaryOutstandingCC = txtEmailCCSummaryOutstanding.GetText();

        AffiliateSubmit.PerformCallback('save|' + pIsUpdate + '|' + pAffiliateID + '|' + pAffiliatePOTO + '|' + pAffiliatePOCC + '|' + pAffiliatePORevisionTO + '|' + pAffiliatePORevisionCC + '|' + pKanbanTO + '|' + pKanbanCC + '|' + pSupplierDeliveryTO + '|' + pSupplierDeliveryCC + '|' + pPASIReceivingTO + '|' + pPASIReceivingCC + '|' + pAffiliateReceivingTO + '|' + pAffiliateReceivingCC + '|' + pGoodReceiveTO + '|' + pGoodReceiveCC + '|' + pInvoiceTO + '|' + pInvoiceCC + '|' + pSummaryOutstandingTO + '|' + pSummaryOutstandingCC);
        
    }

    function readonly() {
        txtAffiliateName.GetInputElement().setAttribute('style', 'background:#CCCCCC;foreground:#CCCCCC;');
        txtAffiliateName.GetInputElement().readOnly = true;
        lblInfo.SetText('');
    }

    </script>
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
            height = height - (height * 52 / 100)
            grid.SetHeight(height);
        }
    </script>

</asp:Content>

<asp:Content ID="Content2" ContentPlaceHolderID="Content" runat="server">
    <table style="border-width: 1pt thin thin thin; border-style: ridge;  border-color:#9598A1; width:100%; height: 27px;">
        <tr>
            <td>
    <table style="width:100%;">
            <tr>
                <td align="left" class="style5">
                            <dx:ASPxLabel ID="ASPxLabel20" runat="server" Text="AFFILIATE CODE/NAME" 
                                Font-Names="Tahoma" font-size="8pt" >
                            </dx:ASPxLabel>
                        </td>
                <td align="left" class="style6">
                            <dx:ASPxComboBox ID="cboAffiliateCode" runat="server" Height="16px"  
                                            ClientInstanceName="cboAffiliateCode" Width="100px"
                                            Font-Size="8pt" 
                                            Font-Names="Tahoma" TextFormatString="{0}" TabIndex="1">
                                            <ClientSideEvents SelectedIndexChanged="function(s, e) {
                                                txtAffiliateName.SetText(cboAffiliateCode.GetSelectedItem().GetColumnText(1));
                                                cbSetData.PerformCallback(cboAffiliateCode.GetSelectedItem().GetColumnText(0)); 
	                                            lblInfo.SetText('');	
                                            }" />
                                            <LoadingPanelStyle ImageSpacing="5px">
                                            </LoadingPanelStyle>
                                     
                            </dx:ASPxComboBox>
                        </td>
                <td align="left" width="201px">
                            <dx:ASPxTextBox ID="txtAffiliateName" runat="server" Width="300px" 
                                Font-Names="Tahoma" font-size="8pt" ClientInstanceName="txtAffiliateName" 
                                MaxLength="2000" BackColor="#CCCCCC" style="margin-left: 0px" 
                                TabIndex="2" >
                                <ClientSideEvents KeyPress="function(s, e) {
	readonly();
}" />
                            </dx:ASPxTextBox>
                        </td>
                <td align="left" width="201px">
                            &nbsp;</td>
                <td align="left" width="201px">
                            &nbsp;</td>
                <td align="left" width="180px">
                            &nbsp;</td>
            </tr>
        </table>
        </td>            
        </tr>
    </table>
    
    <div style="height:5px;"></div> 

    <table style="border-width: thin; border-style: inset hidden ridge hidden; border-color:#9598A1; width:100%; height: 15px;">
        <tr>
            <td colspan="8" height="15">
                <%--error message--%>
                <table id="info" style="width:100%;">
                    <tr>
                        <td align="left" valign="middle" style="height:15px;">
                            <dx:ASPxLabel ID="lblInfo" runat="server" Font-Names="Tahoma"
                                ClientInstanceName="lblInfo" Font-Bold="True" Font-Italic="True" 
                                Font-Size="8pt" >
                            </dx:ASPxLabel>
                        </td>
                    </tr>         
                </table>
            </td>            
        </tr>
    </table>
    
    <div style="height:5px;"></div> 

    <table style="border-width: 1pt thin thin thin; border-style: ridge;  border-color:#9598A1; width:100%; height: 480px;">
        <tr>
            <td>
                <table style="width:100%; height: 450px;">
                    <tr>
                        <td align="left" class="style4" valign="top">
                            <dx:ASPxLabel ID="ASPxLabel39" runat="server" Font-Names="Tahoma" 
                                Font-Size="8pt" Text="PO TYPE">
                            </dx:ASPxLabel>
                        </td>
                        <td align="left" class="style3" valign="top">
                            &nbsp;</td>
                        <td align="left" class="style1" valign="top">
                            <dx:ASPxComboBox ID="cbotype" runat="server" ClientInstanceName="cbotype" 
                                Height="25px">
                                <ClientSideEvents SelectedIndexChanged="function(s, e) {
if (cbotype.GetValue() == 'DOMESTIC')
		{ txtEmailToKanban.SetEnabled(true);
		  txtEmailCCKanban.SetEnabled(true);
		  txtEmailToPASIReceiving.SetEnabled(true);
		  txtEmailCCPASIReceiving.SetEnabled(true);
	} else {
		 txtEmailToKanban.SetEnabled(false);
		 txtEmailCCKanban.SetEnabled(false);
		 txtEmailToPASIReceiving.SetEnabled(false);
		 txtEmailCCPASIReceiving.SetEnabled(false);
	} 
	                                
                                    cbSetData.PerformCallback(cboAffiliateCode.GetSelectedItem().GetColumnText(0)); 

}" />
                            </dx:ASPxComboBox>
                        </td>
                        <td align="right" valign="top">
                            &nbsp;</td>
                        <td align="left" valign="top">     
                            &nbsp;</td>
                    </tr>
                    <tr>
                        <td align="left" class="style4" valign="top">
                            <dx:ASPxLabel ID="ASPxLabel10" runat="server" Font-Names="Tahoma" 
                                Font-Size="8pt" Text="AFFILIATE PO">
                            </dx:ASPxLabel>
                        </td>
                        <td align="left" class="style3" valign="top">
                            <dx:ASPxLabel ID="ASPxLabel21" runat="server" Font-Names="Tahoma" 
                                Font-Size="8pt" Text="EMAIL TO">
                            </dx:ASPxLabel>
                        </td>
                        <td align="left" class="style1" valign="top">
                            <dx:ASPxMemo ID="txtEmailToAffiliatePO" runat="server" Height="42px" Width="335px"
                            ClientInstanceName="txtEmailToAffiliatePO" Font-Names="Tahoma" TabIndex="3">
                            </dx:ASPxMemo>
                        </td>
                        <td align="right" valign="top">
                            <dx:ASPxLabel ID="ASPxLabel30" runat="server" Font-Names="Tahoma" 
                                Font-Size="8pt" Text="EMAIL CC">
                            </dx:ASPxLabel>
                        </td>
                        <td align="left" valign="top">     
                        <dx:ASPxMemo ID="txtEmailCCAffiliatePO" runat="server" Height="42px" Width="335px"
                        ClientInstanceName="txtEmailCCAffiliatePO" Font-Names="Tahoma" MaxLength="2000" 
                                TabIndex="4">
                        </dx:ASPxMemo>
                        </td>
                    </tr>
                    <tr>
                        <td align="left" class="style4" valign="top">
                            <dx:ASPxLabel ID="ASPxLabel11" runat="server" Font-Names="Tahoma" 
                                Font-Size="8pt" Text="AFFILIATE PO REVISION">
                            </dx:ASPxLabel>
                        </td>
                        <td align="left" class="style3" valign="top">
                            <dx:ASPxLabel ID="ASPxLabel22" runat="server" Font-Names="Tahoma" 
                                Font-Size="8pt" Text="EMAIL TO">
                            </dx:ASPxLabel>
                        </td>
                        <td align="left" class="style1" valign="top">
                            <dx:ASPxMemo ID="txtEmailToAffiliatePORevision" runat="server" Height="42px" Width="335px"
                            ClientInstanceName="txtEmailToAffiliatePORevision" Font-Names="Tahoma" 
                                MaxLength="2000" TabIndex="5">
                            </dx:ASPxMemo>
                        </td>
                        <td align="right" valign="top">
                            <dx:ASPxLabel ID="ASPxLabel31" runat="server" Font-Names="Tahoma" 
                                Font-Size="8pt" Text="EMAIL CC">
                            </dx:ASPxLabel>
                        </td>
                        <td align="left" valign="top">
                        <dx:ASPxMemo ID="txtEmailCCAffiliatePORevision" runat="server" Height="42px" Width="335px"
                            ClientInstanceName="txtEmailCCAffiliatePORevision" Font-Names="Tahoma" 
                                MaxLength="2000" TabIndex="6">
                            </dx:ASPxMemo>
                        </td>
                    </tr>
                    <tr>
                        <td align="left" class="style4" valign="top">
                            <dx:ASPxLabel ID="ASPxLabel12" runat="server" Font-Names="Tahoma" 
                                Font-Size="8pt" Text="KANBAN">
                            </dx:ASPxLabel>
                        </td>
                        <td align="left" class="style3" valign="top">
                            <dx:ASPxLabel ID="ASPxLabel23" runat="server" Font-Names="Tahoma" 
                                Font-Size="8pt" Text="EMAIL TO">
                            </dx:ASPxLabel>
                        </td>
                        <td align="left" class="style1" valign="top">
                        <dx:ASPxMemo ID="txtEmailToKanban" runat="server" Height="42px" Width="335px"
                            ClientInstanceName="txtEmailToKanban" Font-Names="Tahoma" MaxLength="2000" 
                                TabIndex="7">
                            </dx:ASPxMemo>
                        </td>
                        <td align="right" valign="top">
                            <dx:ASPxLabel ID="ASPxLabel32" runat="server" Font-Names="Tahoma" 
                                Font-Size="8pt" Text="EMAIL CC">
                            </dx:ASPxLabel>
                        </td>
                        <td align="left" valign="top">
                        <dx:ASPxMemo ID="txtEmailCCKanban" runat="server" Height="42px" Width="335px"
                            ClientInstanceName="txtEmailCCKanban" Font-Names="Tahoma" MaxLength="2000" 
                                TabIndex="8">
                            </dx:ASPxMemo>
                        </td>
                    </tr>
                    <tr>
                        <td align="left" class="style4" valign="top">
                            <dx:ASPxLabel ID="ASPxLabel13" runat="server" Font-Names="Tahoma" 
                                Font-Size="8pt" Text="SUPPLIER DELIVERY">
                            </dx:ASPxLabel>
                        </td>
                        <td align="left" class="style3" valign="top">
                            <dx:ASPxLabel ID="ASPxLabel24" runat="server" Font-Names="Tahoma" 
                                Font-Size="8pt" Text="EMAIL TO">
                            </dx:ASPxLabel>
                        </td>
                        <td align="left" class="style1" valign="top">
                        <dx:ASPxMemo ID="txtEmailToSupplierDelivery" runat="server" Height="42px" Width="335px"
                            ClientInstanceName="txtEmailToSupplierDelivery" Font-Names="Tahoma" 
                                MaxLength="2000" TabIndex="9">
                            </dx:ASPxMemo>
                        </td>
                        <td align="right" valign="top">
                            <dx:ASPxLabel ID="ASPxLabel33" runat="server" Font-Names="Tahoma" 
                                Font-Size="8pt" Text="EMAIL CC">
                            </dx:ASPxLabel>
                        </td>
                        <td align="left" valign="top">
                        <dx:ASPxMemo ID="txtEmailCCSupplierDelivery" runat="server" Height="42px" Width="335px"
                            ClientInstanceName="txtEmailCCSupplierDelivery" Font-Names="Tahoma" 
                                MaxLength="2000" TabIndex="10">
                            </dx:ASPxMemo>
                        </td>
                    </tr>
                    <tr>
                        <td align="left" class="style4" valign="top">
                            <dx:ASPxLabel ID="ASPxLabel14" runat="server" Font-Names="Tahoma" 
                                Font-Size="8pt" Text="PASI RECEIVING">
                            </dx:ASPxLabel>
                        </td>
                        <td align="left" class="style3" valign="top">
                            <dx:ASPxLabel ID="ASPxLabel25" runat="server" Font-Names="Tahoma" 
                                Font-Size="8pt" Text="EMAIL TO">
                            </dx:ASPxLabel>
                        </td>
                        <td align="left" class="style1" valign="top">
                        <dx:ASPxMemo ID="txtEmailToPASIReceiving" runat="server" Height="42px" Width="335px"
                            ClientInstanceName="txtEmailToPASIReceiving" Font-Names="Tahoma" 
                                MaxLength="2000" TabIndex="11">
                            </dx:ASPxMemo>
                        </td>
                        <td align="right" valign="top">
                            <dx:ASPxLabel ID="ASPxLabel34" runat="server" Font-Names="Tahoma" 
                                Font-Size="8pt" Text="EMAIL CC">
                            </dx:ASPxLabel>
                        </td>
                        <td align="left" valign="top">
                        <dx:ASPxMemo ID="txtEmailCCPASIReceiving" runat="server" Height="42px" Width="335px"
                            ClientInstanceName="txtEmailCCPASIReceiving" Font-Names="Tahoma" 
                                MaxLength="2000" TabIndex="12">
                            </dx:ASPxMemo>
                        </td>
                    </tr>
                    <tr>
                        <td align="left" class="style4" valign="top">
                            <dx:ASPxLabel ID="ASPxLabel15" runat="server" Font-Names="Tahoma" 
                                Font-Size="8pt" Text="AFFILIATE RECEIVING">
                            </dx:ASPxLabel>
                        </td>
                        <td align="left" class="style3" valign="top">
                            <dx:ASPxLabel ID="ASPxLabel26" runat="server" Font-Names="Tahoma" 
                                Font-Size="8pt" Text="EMAIL TO">
                            </dx:ASPxLabel>
                        </td>
                        <td align="left" class="style1" valign="top">
                        <dx:ASPxMemo ID="txtEmailToAffiliateReceiving" runat="server" Height="42px" Width="335px"
                            ClientInstanceName="txtEmailToAffiliateReceiving" Font-Names="Tahoma" 
                                MaxLength="2000" TabIndex="13">
                            </dx:ASPxMemo>
                        </td>
                        <td align="right" valign="top">
                            <dx:ASPxLabel ID="ASPxLabel35" runat="server" Font-Names="Tahoma" 
                                Font-Size="8pt" Text="EMAIL CC">
                            </dx:ASPxLabel>
                        </td>
                        <td align="left" valign="top">
                        <dx:ASPxMemo ID="txtEmailCCAffiliateReceiving" runat="server" Height="42px" Width="335px"
                            ClientInstanceName="txtEmailCCAffiliateReceiving" Font-Names="Tahoma" 
                                MaxLength="2000" TabIndex="14">
                            </dx:ASPxMemo>
                        </td>
                    </tr>
                    <tr>
                        <td align="left" class="style4" valign="top">
                            <dx:ASPxLabel ID="ASPxLabel16" runat="server" Font-Names="Tahoma" 
                                Font-Size="8pt" Text="GOOD RECEIVE">
                            </dx:ASPxLabel>
                        </td>
                        <td align="left" class="style3" valign="top">
                            <dx:ASPxLabel ID="ASPxLabel27" runat="server" Font-Names="Tahoma" 
                                Font-Size="8pt" Text="EMAIL TO">
                            </dx:ASPxLabel>
                        </td>
                        <td align="left" class="style1" valign="top">
                        <dx:ASPxMemo ID="txtEmailToGoodReceive" runat="server" Height="42px" Width="335px"
                            ClientInstanceName="txtEmailToGoodReceive" Font-Names="Tahoma" 
                                MaxLength="2000" TabIndex="15">
                            </dx:ASPxMemo>
                        </td>
                        <td align="right" valign="top">
                            <dx:ASPxLabel ID="ASPxLabel36" runat="server" Font-Names="Tahoma" 
                                Font-Size="8pt" Text="EMAIL CC">
                            </dx:ASPxLabel>
                        </td>
                        <td align="left" valign="top">
                        <dx:ASPxMemo ID="txtEmailCCGoodReceive" runat="server" Height="42px" Width="335px"
                            ClientInstanceName="txtEmailCCGoodReceive" Font-Names="Tahoma" 
                                MaxLength="2000" TabIndex="16">
                            </dx:ASPxMemo>
                        </td>
                    </tr>
                    <tr>
                        <td align="left" class="style4" valign="top">
                            <dx:ASPxLabel ID="ASPxLabel17" runat="server" Font-Names="Tahoma" 
                                Font-Size="8pt" Text="INVOICE">
                            </dx:ASPxLabel>
                        </td>
                        <td align="left" class="style3" valign="top">
                            <dx:ASPxLabel ID="ASPxLabel28" runat="server" Font-Names="Tahoma" 
                                Font-Size="8pt" Text="EMAIL TO">
                            </dx:ASPxLabel>
                        </td>
                        <td align="left" class="style1" valign="top">
                        <dx:ASPxMemo ID="txtEmailToInvoice" runat="server" Height="42px" Width="335px"
                            ClientInstanceName="txtEmailToInvoice" Font-Names="Tahoma" MaxLength="2000" 
                                TabIndex="17">
                            </dx:ASPxMemo>
                        </td>
                        <td align="right" valign="top">
                            <dx:ASPxLabel ID="ASPxLabel37" runat="server" Font-Names="Tahoma" 
                                Font-Size="8pt" Text="EMAIL CC">
                            </dx:ASPxLabel>
                        </td>
                        <td align="left" valign="top">
                        <dx:ASPxMemo ID="txtEmailCCInvoice" runat="server" Height="42px" Width="335px"
                            ClientInstanceName="txtEmailCCInvoice" Font-Names="Tahoma" MaxLength="2000" 
                                TabIndex="18">
                            </dx:ASPxMemo>
                        </td>
                    </tr>
                    <tr>
                        <td align="left" class="style4" valign="top">
                            <dx:ASPxLabel ID="ASPxLabel18" runat="server" Font-Names="Tahoma" 
                                Font-Size="8pt" Text="SUMMARY OUTSTANDING">
                            </dx:ASPxLabel>
                        </td>
                        <td align="left" class="style3" valign="top">
                            <dx:ASPxLabel ID="ASPxLabel29" runat="server" Font-Names="Tahoma" 
                                Font-Size="8pt" Text="EMAIL TO">
                            </dx:ASPxLabel>
                        </td>
                        <td align="left" class="style1" valign="top">
                        <dx:ASPxMemo ID="txtEmailToSummaryOutstanding" runat="server" Height="42px" Width="335px"
                            ClientInstanceName="txtEmailToSummaryOutstanding" Font-Names="Tahoma" 
                                MaxLength="2000" TabIndex="19">
                            </dx:ASPxMemo>
                        </td>
                        <td align="right" valign="top">
                            <dx:ASPxLabel ID="ASPxLabel38" runat="server" Font-Names="Tahoma" 
                                Font-Size="8pt" Text="EMAIL CC">
                            </dx:ASPxLabel>
                        </td>
                        <td align="left" valign="top">
                        <dx:ASPxMemo ID="txtEmailCCSummaryOutstanding" runat="server" Height="42px" Width="335px"
                            ClientInstanceName="txtEmailCCSummaryOutstanding" Font-Names="Tahoma" 
                                MaxLength="2000" TabIndex="20">
                            </dx:ASPxMemo>
                        </td>
                    </tr>
                    </table>
            </td>
        </tr>
    </table> 

    <div style="height:8px;"></div>      

    <%--Button--%> 
    <table id="button" style=" width:100%;">
        <tr>                        
            <td valign="top" align="left">
                <dx:ASPxButton ID="btnSubMenu" runat="server" Text="SUB MENU"                     
                    Font-Names="Tahoma"
                    Width="90px" Font-Size="8pt" ClientInstanceName="btnSubMenu" TabIndex="23">
                </dx:ASPxButton>   
                </td>                     
            
            <td valign="top" align="right" style="width: 50px;">                                  
            </td>
            <td valign="top" align="right" style="width: 50px;">                                   
                &nbsp;</td>
            <td align="right" style="width:80px;">                                   
                <dx:ASPxButton ID="btnClear" runat="server" Text="CLEAR"                    
                    Font-Names="Tahoma"
                    Width="90px" AutoPostBack="False" Font-Size="8pt" 
                    ClientInstanceName="btnClear" TabIndex="22">                   
                </dx:ASPxButton>
            </td>
            <td valign="top" align="right" style="width: 50px;">                                   
                <dx:ASPxButton ID="btnSubmit" runat="server" Text="SAVE"                    
                    Font-Names="Tahoma"
                    Width="90px" AutoPostBack="False" Font-Size="8pt" 
                    ClientInstanceName="btnSubmit" TabIndex="21">
                    <ClientSideEvents Click="function(s, e) {
                        validasubmit();
                        up_Insert();
                        
                       
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

    <dx:ASPxCallback ID="AffiliateSubmit" runat="server" ClientInstanceName = "AffiliateSubmit">
        <ClientSideEvents EndCallback="function(s, e) {
            var pMsg = s.cpMessage;        
            if (pMsg != '') {
                if (s.cpType == 'error'){
                    lblInfo.GetMainElement().style.color = 'Red';
                }
                else if (s.cpType == 'info'){
                    lblInfo.GetMainElement().style.color = 'Blue';
                }
                else {
                    lblInfo.GetMainElement().style.color = 'Red';
                }
        
                lblInfo.SetText(pMsg);
               if(s.cpFunction == 'insert'){
                }
            } else {
                lblInfo.SetText('');
            }  
        }" />
    </dx:ASPxCallback>
    <dx:ASPxCallback ID="cbSetData" runat="server" ClientInstanceName = "cbSetData">
        <ClientSideEvents CallbackComplete="function(s, e) {
                txtEmailToAffiliatePO.SetText(s.cpAffiliatePOTO);
                txtEmailCCAffiliatePO.SetText(s.cpAffiliatePOCC);
                txtEmailToAffiliatePORevision.SetText(s.cpAffiliatePORevisionTO);
                txtEmailCCAffiliatePORevision.SetText(s.cpAffiliatePORevisionCC);
                txtEmailToKanban.SetText(s.cpKanbanTO);
                txtEmailCCKanban.SetText(s.cpKanbanCC);
                txtEmailToSupplierDelivery.SetText(s.cpSupplierDeliveryTO);
                txtEmailCCSupplierDelivery.SetText(s.cpSupplierDeliveryCC);
                txtEmailToPASIReceiving.SetText(s.cpPASIReceivingTO);
                txtEmailCCPASIReceiving.SetText(s.cpPASIReceivingCC);
                txtEmailToAffiliateReceiving.SetText(s.cpAffiliateReceivingTO);
                txtEmailCCAffiliateReceiving.SetText(s.cpAffiliateReceivingCC);
                txtEmailToGoodReceive.SetText(s.cpGoodReceiveTO);
                txtEmailCCGoodReceive.SetText(s.cpGoodReceiveCC);
                txtEmailToInvoice.SetText(s.cpInvoiceTO);
                txtEmailCCInvoice.SetText(s.cpInvoiceCC);
                txtEmailToSummaryOutstanding.SetText(s.cpSummaryOutstandingTO);
                txtEmailCCSummaryOutstanding.SetText(s.cpSummaryOutstandingCC);

          }" />
    </dx:ASPxCallback>
</asp:Content>
