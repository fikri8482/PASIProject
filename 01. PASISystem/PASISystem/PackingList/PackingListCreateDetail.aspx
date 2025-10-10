<%@ Page Title="" Language="vb" AutoEventWireup="false" MasterPageFile="~/SiteContent.Master"
    CodeBehind="PackingListCreateDetail.aspx.vb" Inherits="PASISystem.PackingListCreateDetail" %>

<%@ Register Assembly="DevExpress.Web.v14.1, Version=14.1.6.0, Culture=neutral, PublicKeyToken=b88d1754d700e49a"
    Namespace="DevExpress.Web.ASPxGridView.Export" TagPrefix="dx" %>
<%@ Register Assembly="DevExpress.Web.v14.1, Version=14.1.6.0, Culture=neutral, PublicKeyToken=b88d1754d700e49a"
    Namespace="DevExpress.Web.ASPxEditors" TagPrefix="dx1" %>
<%@ Register Assembly="DevExpress.Web.v14.1, Version=14.1.6.0, Culture=neutral, PublicKeyToken=b88d1754d700e49a"
    Namespace="DevExpress.Web.ASPxGridView" TagPrefix="dx" %>
<%@ Register Assembly="DevExpress.Web.v14.1, Version=14.1.6.0, Culture=neutral, PublicKeyToken=b88d1754d700e49a"
    Namespace="DevExpress.Web.ASPxHiddenField" TagPrefix="dx" %>
<%@ Register Assembly="DevExpress.Web.v14.1, Version=14.1.6.0, Culture=neutral, PublicKeyToken=b88d1754d700e49a" Namespace="DevExpress.Web.ASPxCallback" TagPrefix="dx" %>
<asp:Content ID="Content1" ContentPlaceHolderID="head" runat="server">
    <style type="text/css">
        .dxeBase
        {
            font: 12px Tahoma, Geneva, sans-serif;
            margin-left: 6px;
        }
        
        .style25
        {
            width: 1001px;
            height: 20px;
        }
    </style>
</asp:Content>
<asp:Content ID="Content2" ContentPlaceHolderID="Content" runat="server">
    <br />
    <script language="javascript" type="text/javascript">
        function OnUpdateClick(s, e) {
            Grid.PerformCallback("Update");
        }

        function OnCancelClick(s, e) {
            Grid.PerformCallback("Cancel");
        }

        function OnBatchEditStartEditing(s, e) {
            currentColumnName = e.focusedColumn.fieldName;

            if (currentColumnName == "colno" || currentColumnName == "colpono" || currentColumnName == "colpokanban" || currentColumnName == "colkanbanno"
            || currentColumnName == "colpartno" || currentColumnName == "colpartname" || currentColumnName == "coluom" || currentColumnName == "colQtyBox"
            || currentColumnName == "colpasiremaining" || currentColumnName == "colsuppdelqty" || currentColumnName == "colpasigoodrec" || currentColumnName == "colpasidefectrec"
            || currentColumnName == "colremainingdelqty" || currentColumnName == "coldelqtybox" || currentColumnName == "colpasidefectrec") {
                e.cancel = true;
            }

            currentEditableVisibleIndex = e.visibleIndex;
        }
        function OnAllCheckedChanged(s, e) {
            if (s.GetValue() == -1) s.SetValue(1);
            for (var i = 0; i < Grid.GetVisibleRowsOnPage(); i++) {
                Grid.batchEditApi.SetCellValue(i, "AllowAccess", s.GetValue());
            }
        }
    </script>
    <table style="border: thin groove #808080; width: 100%;" width="100%">
        <tr>
            <td>
                <table style="width: 100%;">
                    <tr>
                        <td align="left">
                            <dx1:ASPxLabel ID="ASPxLabel1" runat="server" Font-Names="Tahoma" Font-Size="8pt"
                                Text="DELIVERY DATE">
                            </dx1:ASPxLabel>
                        </td>
                        <td align="left">
                            <dx1:ASPxTextBox ID="txtdeliverydate" runat="server" Width="165px" Font-Names="Tahoma"
                                Font-Size="8pt" BackColor="#CCCCCC" ReadOnly="True" Height="16px" ClientInstanceName="txtdeliverydate">
                                <ClientSideEvents TextChanged="function(s, e) {
	                                lblerrmessage.SetText('');
                                }" />
                            </dx1:ASPxTextBox>
                        </td>
                        <td align="left">
                            &nbsp;
                        </td>
                        <td align="left">
                            <dx1:ASPxLabel ID="ASPxLabel7" runat="server" Font-Names="Tahoma" Font-Size="8pt"
                                Text="SUPPLIER CODE" Visible="False">
                            </dx1:ASPxLabel>
                        </td>
                        <td align="left">
                            <dx1:ASPxTextBox ID="txtSupplierCode" runat="server" Width="165px" Font-Names="Tahoma"
                                Font-Size="8pt" BackColor="#CCCCCC" ReadOnly="True" Height="16px" 
                                ClientInstanceName="txtSupplierCode" Visible="False">
                                <ClientSideEvents TextChanged="function(s, e) {
	                                lblerrmessage.SetText('');
                                }" />
                            </dx1:ASPxTextBox>
                        </td>
                        <td align="left">
                            <dx1:ASPxTextBox ID="txtSupplierName" runat="server" Width="165px" Font-Names="Tahoma"
                                Font-Size="8pt" BackColor="#CCCCCC" ReadOnly="True" Height="16px" 
                                ClientInstanceName="txtSupplierName" Visible="False">
                                <ClientSideEvents TextChanged="function(s, e) {
	                                lblerrmessage.SetText('');
                                }" />
                            </dx1:ASPxTextBox>
                        </td>
                        <td align="left">
                            &nbsp;
                        </td>
                    </tr>
                    <tr>
                        <td align="left">
                            <dx1:ASPxLabel ID="ASPxLabel2" runat="server" Font-Names="Tahoma" Font-Size="8pt"
                                Text="AFFILIATE CODE/NAME">
                            </dx1:ASPxLabel>
                        </td>
                        <td align="left">
                            <dx1:ASPxTextBox ID="txtaffiliatecode" runat="server" Width="165px" Font-Names="Tahoma"
                                Font-Size="8pt" BackColor="#CCCCCC" ReadOnly="True" Height="16px" ClientInstanceName="txtaffiliatecode">
                                <ClientSideEvents TextChanged="function(s, e) {
	                                lblerrmessage.SetText('');
                                }" />
                            </dx1:ASPxTextBox>
                        </td>
                        <td align="left">
                            <dx1:ASPxTextBox ID="txtaffiliatename" runat="server" Width="165px" Font-Names="Tahoma"
                                Font-Size="8pt" BackColor="#CCCCCC" ReadOnly="True" Height="16px" ClientInstanceName="txtaffiliatename">
                                <ClientSideEvents TextChanged="function(s, e) {
	                                lblerrmessage.SetText('');
                                }" />
                            </dx1:ASPxTextBox>
                        </td>
                        <td align="left">
                            <dx1:ASPxLabel ID="ASPxLabel5" runat="server" Font-Names="Tahoma" Font-Size="8pt"
                                Text="DELIVERY LOCATION">
                            </dx1:ASPxLabel>
                        </td>
                        <td align="left">
                            <dx1:ASPxTextBox ID="txtdeliverylocationCode" runat="server" Width="165px" Font-Names="Tahoma"
                                Font-Size="8pt" BackColor="#CCCCCC" ReadOnly="True" Height="16px" ClientInstanceName="txtdeliverylocationCode">
                                <ClientSideEvents TextChanged="function(s, e) {
	                                lblerrmessage.SetText('');
                                }" />
                            </dx1:ASPxTextBox>
                        </td>
                        <td align="left">
                            <dx1:ASPxTextBox ID="txtdeliverylocationName" runat="server" Width="165px" Font-Names="Tahoma"
                                Font-Size="8pt" BackColor="#CCCCCC" ReadOnly="True" Height="16px" ClientInstanceName="txtdeliverylocationName">
                                <ClientSideEvents TextChanged="function(s, e) {
	                                lblerrmessage.SetText('');
                                }" />
                            </dx1:ASPxTextBox>
                        </td>
                        <td align="left">
                        </td>
                    </tr>
                    <tr>
                        <td align="left">
                            <dx1:ASPxLabel ID="ASPxLabel4" runat="server" Font-Names="Tahoma" Font-Size="8pt"
                                Text="PASI SURAT JALAN NO.*">
                            </dx1:ASPxLabel>
                        </td>
                        <td align="left">
                            <dx1:ASPxTextBox ID="txtsuratjalanno" runat="server" Width="165px" Font-Names="Tahoma"
                                Font-Size="8pt" Height="16px" ClientInstanceName="txtsuratjalanno" BackColor="#CCCCCC"
                                ReadOnly="True">
                                <ClientSideEvents TextChanged="function(s, e) {
	                                lblerrmessage.SetText('');
                                }" />
                            </dx1:ASPxTextBox>
                        </td>
                        <td align="left">
                            <table style="width: 100%;">
                                <tr>
                                    <td style="width: 150px;">
                                        <dx1:ASPxLabel ID="ASPxLabel14" runat="server" Font-Names="Tahoma" Font-Size="8pt"
                                            Text="PACKING LIST NO">
                                        </dx1:ASPxLabel>
                                    </td>
                                    <td>
                                        <dx1:ASPxTextBox ID="txtInvoiceNo" runat="server" Width="100px" Font-Names="Tahoma"
                                            Font-Size="8pt" BackColor="#CCCCCC" Height="16px" MaxLength="20" ClientInstanceName="txtInvoiceNo"
                                            ReadOnly="True">
                                            <ClientSideEvents TextChanged="function(s, e) {lblerrmessage.SetText('');}" />
                                        </dx1:ASPxTextBox>
                                    </td>
                                </tr>
                            </table>
                        </td>
                        <td align="left">
                            <table style="width: 100%;">
                                <tr>
                                    <td style="width: 150px;">
                                        <dx1:ASPxLabel ID="ASPxLabel3" runat="server" Font-Names="Tahoma" Font-Size="8pt"
                                            Text="TYPE DOCUMENT">
                                        </dx1:ASPxLabel>
                                    </td>
                                    <td>
                                        <dx1:ASPxComboBox ID="cboCommercial" runat="server" ClientInstanceName="cboCommercial"
                                            Font-Names="Tahoma" TextFormatString="{0}" Font-Size="8pt" Width="100px" DropDownStyle="DropDownList">
                                            <ClientSideEvents SelectedIndexChanged="function(s, e) {lblerrmessage.SetText('');}" />
                                        </dx1:ASPxComboBox>
                                    </td>
                                </tr>
                            </table>
                        </td>
                        <td align="left">
                            &nbsp;
                        </td>
                        <td align="left">
                            &nbsp;
                        </td>
                        <td align="left">
                            &nbsp;
                        </td>
                    </tr>
                    <tr>
                        <td align="left">
                            <dx1:ASPxLabel ID="ASPxLabel15" runat="server" Font-Names="Tahoma" Font-Size="8pt"
                                Text="DRIVER NAME">
                            </dx1:ASPxLabel>
                        </td>
                        <td align="left">
                            <dx1:ASPxTextBox ID="txtdrivername" runat="server" Width="165px" Font-Names="Tahoma"
                                Font-Size="8pt" BackColor="#CCCCCC" Height="16px" ClientInstanceName="txtdrivername"
                                ReadOnly="True">
                                <ClientSideEvents TextChanged="function(s, e) {
	                                lblerrmessage.SetText('');
                                }" />
                            </dx1:ASPxTextBox>
                        </td>
                        <td align="left">
                            <table style="width: 100%;">
                                <tr>
                                    <td style="width: 150px;">
                                        <dx1:ASPxLabel ID="ASPxLabel24" runat="server" Font-Names="Tahoma" Font-Size="8pt"
                                            Text="DRIVER CONTACT">
                                        </dx1:ASPxLabel>
                                    </td>
                                    <td>
                                        <dx1:ASPxTextBox ID="txtdrivercontact" runat="server" Width="100px" Font-Names="Tahoma"
                                            Font-Size="8pt" BackColor="#CCCCCC" Height="16px" MaxLength="20" ClientInstanceName="txtdrivercontact"
                                            ReadOnly="True">
                                            <ClientSideEvents TextChanged="function(s, e) {lblerrmessage.SetText('');}" />
                                        </dx1:ASPxTextBox>
                                    </td>
                                </tr>
                            </table>
                        </td>
                        <td align="left">
                            <table style="width: 100%;">
                                <tr>
                                    <td style="width: 150px;">
                                        <dx1:ASPxLabel ID="ASPxLabel25" runat="server" Font-Names="Tahoma" Font-Size="8pt"
                                            Text="NO. POL">
                                        </dx1:ASPxLabel>
                                    </td>
                                    <td>
                                        <dx1:ASPxTextBox ID="txtnopol" runat="server" Width="100px" Font-Names="Tahoma" Font-Size="8pt"
                                            BackColor="#CCCCCC" Height="16px" MaxLength="20" ClientInstanceName="txtnopol"
                                            ReadOnly="True">
                                            <ClientSideEvents TextChanged="function(s, e) {lblerrmessage.SetText('');}" />
                                        </dx1:ASPxTextBox>
                                    </td>
                                </tr>
                            </table>
                        </td>
                        <td align="left">
                            <table style="width: 100%;">
                                <tr>
                                    <td style="width: 150px;">
                                        <dx1:ASPxLabel ID="ASPxLabel26" runat="server" Font-Names="Tahoma" Font-Size="8pt"
                                            Text="JENIS ARMADA">
                                        </dx1:ASPxLabel>
                                    </td>
                                    <td>
                                        <dx1:ASPxTextBox ID="txtjenisarmada" runat="server" Width="100px" Font-Names="Tahoma"
                                            Font-Size="8pt" BackColor="#CCCCCC" Height="16px" MaxLength="20" ClientInstanceName="txtjenisarmada"
                                            ReadOnly="True">
                                            <ClientSideEvents TextChanged="function(s, e) {lblerrmessage.SetText('');}" />
                                        </dx1:ASPxTextBox>
                                    </td>
                                </tr>
                            </table>
                        </td>
                        <td align="left">
                            <table style="width: 100%;">
                                <tr>
                                    <td style="width: 150px;">
                                        <dx1:ASPxLabel ID="ASPxLabel27" runat="server" Font-Names="Tahoma" Font-Size="8pt"
                                            Text="TOTAL BOX">
                                        </dx1:ASPxLabel>
                                    </td>
                                    <td>
                                        <dx1:ASPxTextBox ID="txttotalbox" runat="server" Width="100px" Font-Names="Tahoma"
                                            Font-Size="8pt" BackColor="#CCCCCC" ReadOnly="True" Height="16px" MaxLength="20"
                                            ClientInstanceName="txttotalbox" DisplayFormatString="{0}" HorizontalAlign="Right">
                                            <ClientSideEvents TextChanged="function(s, e) {lblerrmessage.SetText('');}" />
                                        </dx1:ASPxTextBox>
                                    </td>
                                </tr>
                            </table>
                        </td>
                        <td align="left">
                            &nbsp;
                        </td>
                    </tr>
                    <tr>
                        <td align="left">
                            <dx1:ASPxLabel ID="ASPxLabel19" runat="server" Font-Names="Tahoma" Font-Size="8pt"
                                Text="FROM">
                            </dx1:ASPxLabel>
                        </td>
                        <td align="left">
                            <%--<dx1:ASPxTextBox ID="txtFrom" runat="server" Width="165px" Font-Names="Tahoma" Font-Size="8pt"
                                BackColor="White" Height="16px" ClientInstanceName="txtFrom">
                                <ClientSideEvents TextChanged="function(s, e) {lblerrmessage.SetText('');}" />
                            </dx1:ASPxTextBox>--%>
                            <dx1:ASPxComboBox ID="cboFrom" runat="server" ClientInstanceName="cboFrom"
                                Font-Names="Tahoma" TextFormatString="{0}" Font-Size="8pt" Width="165px" DropDownStyle="DropDownList">
                                <ClientSideEvents SelectedIndexChanged="function(s, e) {lblerrmessage.SetText('');}" />
                            </dx1:ASPxComboBox>
                        </td>
                        <td align="left">
                            <table style="width: 100%;">
cb                                <tr>
                                    <td style="width: 150px;">
                                        <dx1:ASPxLabel ID="ASPxLabel6" runat="server" Font-Names="Tahoma" Font-Size="8pt"
                                            Text="VIA">
                                        </dx1:ASPxLabel>
                                    </td>
                                    <td>
                                        <dx1:ASPxTextBox ID="txtVia" runat="server" Width="100px" Font-Names="Tahoma" Font-Size="8pt"
                                            BackColor="White" Height="16px" MaxLength="20" ClientInstanceName="txtVia">
                                            <ClientSideEvents TextChanged="function(s, e) {lblerrmessage.SetText('');}" />
                                        </dx1:ASPxTextBox>
                                    </td>
                                </tr>
                            </table>
                        </td>
                        <td align="left">
                            <table style="width: 100%;">
                                <tr>
                                    <td style="width: 150px;">
                                        <dx1:ASPxLabel ID="ASPxLabel16" runat="server" Font-Names="Tahoma" Font-Size="8pt"
                                            Text="VESSEL">
                                        </dx1:ASPxLabel>
                                    </td>
                                    <td>
                                        <dx1:ASPxTextBox ID="txtVessel" runat="server" Width="100px" Font-Names="Tahoma"
                                            Font-Size="8pt" BackColor="White" Height="16px" MaxLength="20" ClientInstanceName="txtVessel">
                                            <ClientSideEvents TextChanged="function(s, e) {lblerrmessage.SetText('');}" />
                                        </dx1:ASPxTextBox>
                                    </td>
                                </tr>
                            </table>
                        </td>
                        <td align="left">
                            <table style="width: 100%;">
                                <tr>
                                    <td style="width: 150px;">
                                        <dx1:ASPxLabel ID="ASPxLabel17" runat="server" Font-Names="Tahoma" Font-Size="8pt"
                                            Text="ON/ABOUT">
                                        </dx1:ASPxLabel>
                                    </td>
                                    <td>
                                        <dx1:ASPxTextBox ID="txtOn" runat="server" Width="100px" Font-Names="Tahoma" Font-Size="8pt"
                                            BackColor="White" Height="16px" MaxLength="20" ClientInstanceName="txtOn">
                                            <ClientSideEvents TextChanged="function(s, e) {lblerrmessage.SetText('');}" />
                                        </dx1:ASPxTextBox>
                                    </td>
                                </tr>
                            </table>
                        </td>
                        <td align="left" rowspan="3" valign="top">
                            <table style="width: 100%;">
                                <tr>
                                    <td style="width: 150px;" valign="top">
                                        <dx1:ASPxLabel ID="ASPxLabel28" runat="server" Font-Names="Tahoma" Font-Size="8pt"
                                            Text="REMARKS">
                                        </dx1:ASPxLabel>
                                    </td>
                                    <td>
                                        <dx1:ASPxMemo ID="txtRemarks" runat="server" ClientInstanceName="txtRemarks" Font-Names="Tahoma"
                                            Font-Size="8" Height="50px" MaxLength="200" Width="150px">
                                            <ReadOnlyStyle BackColor="#CCCCCC">
                                            </ReadOnlyStyle>
                                        </dx1:ASPxMemo>
                                    </td>
                                </tr>
                            </table>
                        </td>
                        <td align="left">
                            &nbsp;
                        </td>
                    </tr>
                    <tr>
                        <td align="left">
                            <dx1:ASPxLabel ID="ASPxLabel8" runat="server" Font-Names="Tahoma" Font-Size="8pt"
                                Text="TO">
                            </dx1:ASPxLabel>
                        </td>
                        <td align="left">
                            <dx1:ASPxTextBox ID="txtTo" runat="server" Width="165px" Font-Names="Tahoma" Font-Size="8pt"
                                BackColor="White" Height="16px" MaxLength="20" ClientInstanceName="txtTo">
                                <ClientSideEvents TextChanged="function(s, e) {lblerrmessage.SetText('');}" />
                            </dx1:ASPxTextBox>
                        </td>
                        <td align="left">
                            <table style="width: 100%;">
                                <tr>
                                    <td style="width: 150px;">
                                        <dx1:ASPxLabel ID="ASPxLabel9" runat="server" Font-Names="Tahoma" Font-Size="8pt"
                                            Text="ABOUT">
                                        </dx1:ASPxLabel>
                                    </td>
                                    <td>
                                        <dx1:ASPxTextBox ID="txtAbout" runat="server" Width="100px" Font-Names="Tahoma" Font-Size="8pt"
                                            BackColor="White" Height="16px" MaxLength="20" ClientInstanceName="txtAbout">
                                            <ClientSideEvents TextChanged="function(s, e) {lblerrmessage.SetText('');}" />
                                        </dx1:ASPxTextBox>
                                    </td>
                                </tr>
                            </table>
                        </td>
                        <td align="left">
                            <table style="width: 100%;">
                                <tr>
                                    <td style="width: 150px;">
                                        <dx1:ASPxLabel ID="ASPxLabel10" runat="server" Font-Names="Tahoma" Font-Size="8pt"
                                            Text="AWB, B/L NO.">
                                        </dx1:ASPxLabel>
                                    </td>
                                    <td>
                                        <dx1:ASPxTextBox ID="txtAwb" runat="server" Width="100px" Font-Names="Tahoma" Font-Size="8pt"
                                            Height="16px" MaxLength="20" ClientInstanceName="txtAwb">
                                            <ClientSideEvents TextChanged="function(s, e) {
	                                lblerrmessage.SetText('');
                                }" />
                                        </dx1:ASPxTextBox>
                                    </td>
                                </tr>
                            </table>
                        </td>
                        <td align="left">
                            <table style="width: 100%;">
                                <tr>
                                    <td style="width: 150px;">
                                        <dx1:ASPxLabel ID="ASPxLabel20" runat="server" Font-Names="Tahoma" Font-Size="8pt"
                                            Text="CONTAINER NO.">
                                        </dx1:ASPxLabel>
                                    </td>
                                    <td>
                                        <dx1:ASPxTextBox ID="txtContainerNo" runat="server" Width="100px" Font-Names="Tahoma"
                                            Font-Size="8pt" Height="16px" MaxLength="20" ClientInstanceName="txtContainerNo">
                                            <ClientSideEvents TextChanged="function(s, e) {
	                                lblerrmessage.SetText('');
                                }" />
                                        </dx1:ASPxTextBox>
                                    </td>
                                </tr>
                            </table>
                        </td>
                        <td align="left">
                            &nbsp;
                        </td>
                    </tr>
                    <tr>
                        <td align="left">
                            <dx1:ASPxLabel ID="ASPxLabel22" runat="server" Font-Names="Tahoma" Font-Size="8pt"
                                Text="INSURANCE POLICY">
                            </dx1:ASPxLabel>
                        </td>
                        <td align="left">
                            <dx1:ASPxTextBox ID="txtInsurance" runat="server" Width="165px" Font-Names="Tahoma"
                                Font-Size="8pt" Height="16px" MaxLength="20" ClientInstanceName="txtInsurance">
                                <ClientSideEvents TextChanged="function(s, e) {
	                                lblerrmessage.SetText('');
                                }" />
                            </dx1:ASPxTextBox>
                        </td>
                        <td align="left">
                            <table style="width: 100%;">
                                <tr>
                                    <td style="width: 150px;">
                                        <dx1:ASPxLabel ID="ASPxLabel18" runat="server" Font-Names="Tahoma" Font-Size="8pt"
                                            Text="PRIVILEGE">
                                        </dx1:ASPxLabel>
                                    </td>
                                    <td>
                                        <dx1:ASPxTextBox ID="txtPrivilege" runat="server" Width="100px" Font-Names="Tahoma"
                                            Font-Size="8pt" BackColor="White" Height="16px" MaxLength="20" ClientInstanceName="txtPrivilege">
                                            <ClientSideEvents TextChanged="function(s, e) {lblerrmessage.SetText('');}" />
                                        </dx1:ASPxTextBox>
                                    </td>
                                </tr>
                            </table>
                        </td>
                        <td align="left">
                            <table style="width: 100%;">
                                <tr>
                                    <td style="width: 150px;">
                                        <dx1:ASPxLabel ID="ASPxLabel23" runat="server" Font-Names="Tahoma" Font-Size="8pt"
                                            Text="PAYMENT TERMS">
                                        </dx1:ASPxLabel>
                                    </td>
                                    <td>
                                        <dx1:ASPxTextBox ID="txtPaymentTerms" runat="server" Width="100px" Font-Names="Tahoma"
                                            Font-Size="8pt" Height="16px" MaxLength="20" ClientInstanceName="txtPaymentTerms">
                                            <ClientSideEvents TextChanged="function(s, e) {
	                                lblerrmessage.SetText('');
                                }" />
                                        </dx1:ASPxTextBox>
                                    </td>
                                </tr>
                            </table>
                        </td>
                        <td align="left">
                            <table style="width: 100%;">
                                <tr>
                                    <td style="width: 150px;">
                                        <dx1:ASPxLabel ID="ASPxLabel29" runat="server" Font-Names="Tahoma" Font-Size="8pt"
                                            Text="PLACE">
                                        </dx1:ASPxLabel>
                                    </td>
                                    <td>
                                        <dx1:ASPxTextBox ID="txtPlace" runat="server" Width="100px" Font-Names="Tahoma" Font-Size="8pt"
                                            Height="16px" MaxLength="20" ClientInstanceName="txtPlace">
                                            <ClientSideEvents TextChanged="function(s, e) {
	                                lblerrmessage.SetText('');
                                }" />
                                        </dx1:ASPxTextBox>
                                    </td>
                                </tr>
                            </table>
                        </td>
                        <td align="left">
                            <dx1:ASPxButton ID="btnaddrow" runat="server" Text="ADD CARTON " Width="85px" Font-Names="Tahoma"
                                Font-Size="8pt" ClientInstanceName="btnaddrow" AutoPostBack="False">
                                <ClientSideEvents Click="function(s, e) {
                                            lblStatus.SetText('addCarton');
                                            Grid.UpdateEdit();
                                            var millisecondsToWait = 4000;
                                                    setTimeout(function() {
                                                        Grid.PerformCallback('addrow');
                                                    }, millisecondsToWait);	
	                                        Grid.CancelEdit();
                                            lblerrmessage.SetText('');

                                            var pMsg = s.cpMessage;
                                            if (pMsg != '') {
                                                if (pMsg.substring(1,5) == '1001' || pMsg.substring(1,5) == '1002' || pMsg.substring(1,5) == '1003') {
                                                    lblerrmessage.GetMainElement().style.color = 'Blue';  
                                                } else {
                                                    lblerrmessage.GetMainElement().style.color = 'Red';
                                                }
                                                    lblerrmessage.SetText(pMsg);
                                                } else {
                                                    lblerrmessage.SetText('');
                                             }
                                        }" />
                            </dx1:ASPxButton>
                        </td>
                    </tr>
                </table>
            </td>
        </tr>
    </table>
    <table style="width: 100%;" width="100%">
        <tr>
            <td align="left">
                <table width="100%">
                    <tr>
                        <td align="left" bgcolor="White" class="style25" style="border-width: thin; border-style: inset hidden ridge hidden;"
                            width="100%" height="16px">
                            <table style="width: 100%;" width="100%">
                                <tr>
                                    <td width="100%">
                                        <dx1:ASPxLabel ID="lblerrmessage" runat="server" Font-Names="Verdana" Font-Size="8pt"
                                            Text="ERROR MESSAGE" Width="100%" ClientInstanceName="lblerrmessage">
                                        </dx1:ASPxLabel>
                                    </td>
                                </tr>
                            </table>
                        </td>
                    </tr>
                </table>
            </td>
        </tr>
        <tr>
            <td align="right">
                <table width="100%">
                    <tr>
                        <td align="right" width="83%">
                            <dx1:ASPxTextBox ID="ASPxTextBox1" runat="server" Width="20px" Font-Names="Verdana"
                                Font-Size="8pt" BackColor="Yellow" Height="16px">
                                <ClientSideEvents TextChanged="function(s, e) {
	                                lblerrmessage.SetText('');
                                }" />
                            </dx1:ASPxTextBox>
                        </td>
                        <td align="right">
                            <dx1:ASPxLabel ID="ASPxLabel13" runat="server" Font-Names="Tahoma" Font-Size="8pt"
                                Text=": NOT SAVE">
                            </dx1:ASPxLabel>
                        </td>
                        <td align="right">
                            <dx1:ASPxTextBox ID="txtsupplier8" runat="server" Width="20px" Font-Names="Verdana"
                                Font-Size="8pt" BackColor="#FF66CC" Height="16px">
                                <ClientSideEvents TextChanged="function(s, e) {
	                                lblerrmessage.SetText('');
                                }" />
                            </dx1:ASPxTextBox>
                        </td>
                        <td align="right">
                            <dx1:ASPxLabel ID="ASPxLabel11" runat="server" Font-Names="Tahoma" Font-Size="8pt"
                                Text=": DIFFERENCE">
                            </dx1:ASPxLabel>
                        </td>
                    </tr>
                </table>
            </td>
        </tr>
        <tr>
            <td align="left">
                <dx:ASPxGridView ID="Grid" runat="server" AutoGenerateColumns="False" KeyFieldName="NoUrut"
                    Width="100%" ClientInstanceName="Grid">
                    <ClientSideEvents BatchEditStartEditing="OnBatchEditStartEditing" CallbackError="function(s, e) {e.handled = true;}"
                        EndCallback="function(s, e) {
                        var pMsg = s.cpMessage;

                        if (pMsg != '') {
                            if (pMsg.substring(1,5) == '1001' || pMsg.substring(1,5) == '1002' || pMsg.substring(1,5) == '1003') {                           
                                lblerrmessage.GetMainElement().style.color = 'Blue';  
                            } else {
                                lblerrmessage.GetMainElement().style.color = 'Red';
                            }
                            lblerrmessage.SetText(pMsg);
                        } else {
                            lblerrmessage.SetText('');
                        }

                        if(s.cpButton == '0'){            
                            btnSendEDI.SetEnabled(true);
                        }else if (s.cpButton == '1'){                
                            btnSendEDI.SetEnabled(false);
                        }else {
                            btnSendEDI.SetEnabled(true);
                        }

                        delete s.cpButton;
                        delete s.cpMessage;

                        cboFrom.SetText(s.cpFromDelivery);
                        txtTo.SetText(s.cpToDelivery);
                        txtInsurance.SetText(s.cpInsurancePolicy);
                        txtVia.SetText(s.cpViaDelivery);
                        txtAbout.SetText(s.cpAboutDelivery);
                        txtPrivilege.SetText(s.cpPrivilege);
                        txtVessel.SetText(s.cpVessel);
                        txtAwb.SetText(s.cpAWBBLNo);
                        txtPaymentTerms.SetText(s.cpPaymentTerms);
                        txtOn.SetText(s.cpOnAbout);
                        txtContainerNo.SetText(s.cpContainerNo);
                        txtRemarks.SetText(s.cpRemarks);
                        txtPlace.SetText(s.cpPlace);
                        }" />
                    <Settings ShowHorizontalScrollBar="True" ShowVerticalScrollBar="True" ShowGroupButtons="False"
                        ShowStatusBar="Hidden" VerticalScrollableHeight="135" />
                    <Columns>
                        <dx:GridViewDataCheckColumn Caption=" " FieldName="AllowAccess" Name="AllowAccess"
                            VisibleIndex="0" Width="40px">
                            <PropertiesCheckEdit ValueChecked="1" ValueType="System.Int32" ValueUnchecked="0">
                            </PropertiesCheckEdit>
                            <HeaderCaptionTemplate>
                                <dx1:ASPxCheckBox ID="chkAll" runat="server" ClientInstanceName="chkAll" ClientSideEvents-CheckedChanged="OnAllCheckedChanged"
                                    ValueType="System.Int32" ValueChecked="1" ValueUnchecked="0">
                                </dx1:ASPxCheckBox>
                            </HeaderCaptionTemplate>
                            <CellStyle Font-Names="Tahoma" Font-Size="8pt">
                            </CellStyle>
                        </dx:GridViewDataCheckColumn>
                        <dx:GridViewDataTextColumn Caption="NO." FieldName="colno" Name="colno" VisibleIndex="1"
                            Width="30px">
                            <HeaderStyle Font-Names="Tahoma" Font-Size="8pt" Wrap="True" HorizontalAlign="Center" />
                            <CellStyle Font-Names="Tahoma" Font-Size="8pt">
                            </CellStyle>
                        </dx:GridViewDataTextColumn>
                        <dx:GridViewDataTextColumn Caption="PO NO." FieldName="colpono" Name="colpono" VisibleIndex="3"
                            Width="100px">
                            <HeaderStyle Font-Names="Verdana" Font-Size="8pt" Wrap="True" HorizontalAlign="Center" />
                            <CellStyle Font-Names="Verdana" Font-Size="8pt">
                            </CellStyle>
                        </dx:GridViewDataTextColumn>
                        <dx:GridViewDataTextColumn Caption="PO KANBAN" FieldName="colpokanban" Name="colpokanban"
                            VisibleIndex="4" Width="100px">
                            <HeaderStyle Font-Names="Verdana" Font-Size="8pt" Wrap="True" HorizontalAlign="Center" />
                            <CellStyle Font-Names="Verdana" Font-Size="8pt">
                            </CellStyle>
                        </dx:GridViewDataTextColumn>
                        <dx:GridViewDataTextColumn Caption="KANBAN NO." FieldName="colkanbanno" Name="colkanbanno"
                            VisibleIndex="5" Width="100px">
                            <HeaderStyle Font-Names="Verdana" Font-Size="8pt" Wrap="True" HorizontalAlign="Center" />
                            <CellStyle Font-Names="Verdana" Font-Size="8pt">
                            </CellStyle>
                        </dx:GridViewDataTextColumn>
                        <dx:GridViewDataTextColumn Caption="PART NO." FieldName="colpartno" Name="colpartno"
                            VisibleIndex="6" Width="100px">
                            <HeaderStyle Font-Names="Verdana" Font-Size="8pt" Wrap="True" HorizontalAlign="Center" />
                            <CellStyle Font-Names="Verdana" Font-Size="8pt">
                            </CellStyle>
                        </dx:GridViewDataTextColumn>
                        <dx:GridViewDataTextColumn Caption="PART NAME" FieldName="colpartname" Name="colpartname"
                            VisibleIndex="7" Width="150px">
                            <HeaderStyle Font-Names="Verdana" Font-Size="8pt" Wrap="True" HorizontalAlign="Center" />
                            <CellStyle Font-Names="Verdana" Font-Size="8pt">
                            </CellStyle>
                        </dx:GridViewDataTextColumn>
                        <dx:GridViewDataTextColumn Caption="UOM" FieldName="coluom" Name="coluom" VisibleIndex="8"
                            Width="60px">
                            <HeaderStyle Font-Names="Verdana" Font-Size="8pt" Wrap="True" HorizontalAlign="Center" />
                            <CellStyle Font-Names="Verdana" Font-Size="8pt">
                            </CellStyle>
                        </dx:GridViewDataTextColumn>
                        <dx:GridViewDataTextColumn Caption="QTY/BOX" FieldName="colQtyBox" Name="colQtyBox"
                            VisibleIndex="9" Width="70px">
                            <HeaderStyle Font-Names="Verdana" Font-Size="8pt" Wrap="True" HorizontalAlign="Center" />
                            <CellStyle Font-Names="Verdana" Font-Size="8pt">
                            </CellStyle>
                        </dx:GridViewDataTextColumn>
                        <dx:GridViewDataTextColumn Caption="SUPPLIER DELIVERY QTY" FieldName="colsuppdelqty"
                            Name="colsuppdelqty" VisibleIndex="10" Width="0px">
                            <HeaderStyle Font-Names="Verdana" Font-Size="8pt" Wrap="True" HorizontalAlign="Center" />
                            <CellStyle Font-Names="Verdana" Font-Size="8pt">
                            </CellStyle>
                        </dx:GridViewDataTextColumn>
                        <dx:GridViewDataTextColumn Caption="PASI GOOD RECEIVING QTY" FieldName="colpasigoodrec"
                            Name="colpasigoodrec" VisibleIndex="11" Width="0px">
                            <PropertiesTextEdit>
                                <MaskSettings ErrorText="Please input valid value !" Mask="<0..9999999999999g>.<00..99>"
                                    IncludeLiterals="DecimalSymbol" />
                            </PropertiesTextEdit>
                            <HeaderStyle Font-Names="Verdana" Font-Size="8pt" Wrap="True" HorizontalAlign="Center" />
                            <CellStyle Font-Names="Verdana" Font-Size="8pt" HorizontalAlign="Right">
                            </CellStyle>
                        </dx:GridViewDataTextColumn>
                        <dx:GridViewDataTextColumn Caption="PASI DEFECT RECEIVING" FieldName="colpasidefectrec"
                            Name="colpasidefectrec" VisibleIndex="12" Width="0px">
                            <HeaderStyle Font-Names="Verdana" Font-Size="8pt" Wrap="True" HorizontalAlign="Center" />
                            <CellStyle Font-Names="Verdana" Font-Size="8pt" HorizontalAlign="Right">
                            </CellStyle>
                        </dx:GridViewDataTextColumn>
                        <dx:GridViewDataTextColumn Caption="PASI REMAINING RECEIVING" FieldName="colpasiremaining"
                            Name="colpasiremaining" VisibleIndex="13" Width="0px">
                            <HeaderStyle Font-Names="Verdana" Font-Size="8pt" Wrap="True" HorizontalAlign="Center" />
                            <CellStyle Font-Names="Verdana" Font-Size="8pt" HorizontalAlign="Right">
                            </CellStyle>
                        </dx:GridViewDataTextColumn>
                        <dx:GridViewDataTextColumn Caption="PASI DELIVERY QTY" FieldName="colpasideliveryqty"
                            Name="colpasideliveryqty" VisibleIndex="14" Width="70px">
                            <HeaderStyle Font-Names="Tahoma" Font-Size="8pt" HorizontalAlign="Center" VerticalAlign="Middle"
                                Wrap="True" />
                            <CellStyle Font-Names="Tahoma" Font-Size="8pt" HorizontalAlign="Right">
                            </CellStyle>
                        </dx:GridViewDataTextColumn>
                        <dx:GridViewDataTextColumn Caption="REMAINING DELIVERY QTY" FieldName="colremainingdelqty"
                            Name="colremainingdelqty" VisibleIndex="15" Width="0px">
                            <PropertiesTextEdit>
                                <MaskSettings ErrorText="Please input valid value !" Mask="<0..9999999999999g>.<00..99>"
                                    IncludeLiterals="DecimalSymbol" />
                            </PropertiesTextEdit>
                            <HeaderStyle Font-Names="Tahoma" Font-Size="8pt" HorizontalAlign="Center" VerticalAlign="Middle"
                                Wrap="True" />
                            <CellStyle Font-Names="Tahoma" Font-Size="8pt">
                            </CellStyle>
                        </dx:GridViewDataTextColumn>
                        <dx:GridViewDataTextColumn Caption="DELIVERY QTY (BOX)" FieldName="coldelqtybox"
                            Name="coldelqtybox" VisibleIndex="16" Width="0px">
                            <HeaderStyle Font-Names="Tahoma" Font-Size="8pt" HorizontalAlign="Center" VerticalAlign="Middle"
                                Wrap="True" />
                            <CellStyle Font-Names="Tahoma" Font-Size="8pt">
                            </CellStyle>
                        </dx:GridViewDataTextColumn>
                        <dx:GridViewDataTextColumn FieldName="colCls" Name="colCls" VisibleIndex="19" 
                            Width="0px">
                        </dx:GridViewDataTextColumn>
                        <dx:GridViewDataTextColumn FieldName="colstsDO" Name="colstsDO" VisibleIndex="20"
                            Width="0px">
                        </dx:GridViewDataTextColumn>
                        <dx:GridViewDataTextColumn Caption="CARTON NO." FieldName="colcartonno" VisibleIndex="17"
                            Width="100px">
                            <HeaderStyle Font-Names="Tahoma" Font-Size="8pt" HorizontalAlign="Center" VerticalAlign="Middle"
                                Wrap="True" />
                            <CellStyle Font-Names="Tahoma" Font-Size="8pt">
                            </CellStyle>
                        </dx:GridViewDataTextColumn>
                        <dx:GridViewDataTextColumn Caption="CARTON QTY" FieldName="colcartonqty" VisibleIndex="18"
                            Width="70px">
                            <HeaderStyle Font-Names="Tahoma" Font-Size="8pt" HorizontalAlign="Center" VerticalAlign="Middle"
                                Wrap="True" />
                            <CellStyle Font-Names="Tahoma" Font-Size="8pt" HorizontalAlign="Right">
                            </CellStyle>
                        </dx:GridViewDataTextColumn>
                        <dx:GridViewDataTextColumn Caption="colponos" FieldName="colponos" VisibleIndex="21"
                            Width="0px">
                        </dx:GridViewDataTextColumn>
                        <dx:GridViewDataTextColumn Caption="colkanbannos" FieldName="colkanbannos" VisibleIndex="22"
                            Width="0px">
                        </dx:GridViewDataTextColumn>
                        <dx:GridViewDataTextColumn FieldName="colsupp" Name="colsupp" VisibleIndex="2" 
                            Width="0px" Caption="SUPPLIER">
                        </dx:GridViewDataTextColumn>
                        <dx:GridViewDataTextColumn Caption="SURAT JALAN SUPPLIER" FieldName="colSJSupp" Name="colSJSupp" VisibleIndex="23" 
                            Width="100px">
                            <HeaderStyle Font-Names="Tahoma" Font-Size="8pt" HorizontalAlign="Center" VerticalAlign="Middle"
                                Wrap="True" />
                            <CellStyle Font-Names="Tahoma" Font-Size="8pt" HorizontalAlign="Right">
                            </CellStyle>
                        </dx:GridViewDataTextColumn>
                    </Columns>
                    <SettingsPager Visible="False" Position="Top">
                        <Summary Text="Page {0} of {1} [{2} record(s)]" AllPagesText="Page {0} of {1} " />
                        <Summary AllPagesText="Page {0} of {1} " Text="Page {0} of {1} [{2} record(s)]">
                        </Summary>
                    </SettingsPager>
                    <SettingsEditing Mode="Batch" NewItemRowPosition="Bottom">
                    </SettingsEditing>
                    <Settings ShowGroupButtons="False" ShowVerticalScrollBar="True" ShowHorizontalScrollBar="True"
                        VerticalScrollableHeight="150" ShowStatusBar="Hidden"></Settings>
                    <StylesEditors ButtonEditCellSpacing="0">
                        <ProgressBar Height="21px">
                        </ProgressBar>
                    </StylesEditors>
                </dx:ASPxGridView>
            </td>
        </tr>
        <tr>
            <td align="left">
               <dx1:ASPxTextBox ID="lblStatus" runat="server" Width="0px" Font-Names="Verdana"
                                Font-Size="8pt" ForeColor="White" Height="16px" 
                    ClientInstanceName="lblStatus" BackColor="White" ReadOnly="True">
                   <Border BorderColor="White" />
               </dx1:ASPxTextBox>
            </td>
        </tr>
        <tr>
            <td align="left">
                <table style="width: 100%;">
                    <tr>
                        <td align="left">
                            <dx1:ASPxButton ID="btnsubmenu" runat="server" Width="90px" Font-Names="Tahoma" Font-Size="8pt"
                                Text="SUB MENU">
                                <ClientSideEvents Click="function(s, e) {}" />
                            </dx1:ASPxButton>
                        </td>                        
                        <td style="width: 90px;">
                            <dx1:ASPxButton ID="btnSendEDI" runat="server" Width="90px" Font-Names="Tahoma" Font-Size="8pt"
                                Text="SEND E.D.I TO AFFILIATE" ClientInstanceName="btnSendEDI" AutoPostBack="False">
                                <ClientSideEvents Click="function(s, e) {                    
                                lblerrmessage.SetText('');                                
                                Grid.PerformCallback('SendEDI');                   
                            }" /> 
                            </dx1:ASPxButton>
                        </td>
                        <td style="width: 90px;">
                            <dx1:ASPxButton ID="btnPrint" runat="server" Width="90px" Font-Names="Tahoma" Font-Size="8pt"
                                Text="PRINT" ClientInstanceName="btnPrint">
                            </dx1:ASPxButton>
                        </td>
                        <td style="width: 90px;">
                            <dx1:ASPxButton ID="btndeletePL" runat="server" Width="90px" Font-Names="Tahoma" Font-Size="8pt"
                                Text="DELETE PL" ClientInstanceName="btndeletePL" AutoPostBack="False">
                                <ClientSideEvents Click="function(s, e) {
                                    if (txtsuratjalanno.GetText() == '') {                                   
                                        lblerrmessage.SetText('[6011] Please Input Surat Jalan No first!');
		                                lblerrmessage.GetMainElement().style.color = 'Red';
                                        return false;
                                    }
                                    var msg = confirm('Are you sure want to delete this data ?');
                                    if (msg == false) {
                                        e.processOnServer = false;
                                        return;
                                    }
                                    lblStatus.SetText('deleteDataPL');
                                    Grid.UpdateEdit();
                                    var millisecondsToWait = 50;
                                            setTimeout(function() {
                                                Grid.PerformCallback('DeletePL');
                                            }, millisecondsToWait);	
	                                Grid.CancelEdit();
                                }" />
                            </dx1:ASPxButton>
                        </td>
                        <td style="width: 90px;">
                            <dx1:ASPxButton ID="btndelete" runat="server" Width="90px" Font-Names="Tahoma" Font-Size="8pt"
                                Text="DELETE CARTON" ClientInstanceName="btndelete" AutoPostBack="False">
                                <ClientSideEvents Click="function(s, e) {
                                    if (txtsuratjalanno.GetText() == '') {                                   
                                        lblerrmessage.SetText('[6011] Please Input Surat Jalan No first!');
		                                lblerrmessage.GetMainElement().style.color = 'Red';
                                        return false;
                                    }
                                    var msg = confirm('Are you sure want to delete this data ?');
                                    if (msg == false) {
                                        e.processOnServer = false;
                                        return;
                                    }
                                    lblStatus.SetText('deleteData');
                                    Grid.UpdateEdit();
                                    var millisecondsToWait = 50;
                                            setTimeout(function() {
                                                Grid.PerformCallback('Delete');
                                            }, millisecondsToWait);	
	                                Grid.CancelEdit();
                                }" />
                            </dx1:ASPxButton>
                        </td>
                        <td style="width: 90px;">
                            <dx1:ASPxButton ID="btnsubmit" runat="server" Width="90px" Font-Names="Tahoma" Font-Size="8pt"
                                Text="SAVE" ClientInstanceName="btnsubmit" AutoPostBack="False">
                                <ClientSideEvents Click="function(s, e) {
                                    lblStatus.SetText('saveData');
	                                Grid.UpdateEdit();
                                    var millisecondsToWait = 4000;
                                            setTimeout(function() {
                                               Grid.PerformCallback('saveDataMaster|' + cboFrom.GetText() + '|' + txtTo.GetText() + '|' + txtInsurance.GetText() + '|' + txtVia.GetText() + '|' + txtAbout.GetText() + '|' + txtPrivilege.GetText() + '|' + txtVessel.GetText() + '|' + txtAwb.GetText() + '|' + txtPaymentTerms.GetText() + '|' + txtOn.GetText() + '|' + txtContainerNo.GetText() + '|' + txtRemarks.GetText() + '|' + txtPlace.GetText() + '|' + cboCommercial.GetText());
                                            }, millisecondsToWait);	
	                                Grid.CancelEdit();
                                }" />
                            </dx1:ASPxButton>
                        </td>
                    </tr>
                </table>
            </td>
        </tr>
    </table>
    <dx:ASPxHiddenField ID="HF" runat="server" ClientInstanceName="HF">
    </dx:ASPxHiddenField>
</asp:Content>
