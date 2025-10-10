<%@ Page Title="" Language="vb" AutoEventWireup="false" MasterPageFile="~/SiteContent.Master" CodeBehind="InvToAffExport.aspx.vb" Inherits="PASISystem.InvToAffExport" %>
<%@ Register Assembly="DevExpress.Web.v14.1, Version=14.1.6.0, Culture=neutral, PublicKeyToken=b88d1754d700e49a" Namespace="DevExpress.Web.ASPxGridView.Export" TagPrefix="dx" %>
<%@ Register Assembly="DevExpress.Web.v14.1, Version=14.1.6.0, Culture=neutral, PublicKeyToken=b88d1754d700e49a" Namespace="DevExpress.Web.ASPxEditors" TagPrefix="dx1" %>
<%@ Register Assembly="DevExpress.Web.v14.1, Version=14.1.6.0, Culture=neutral, PublicKeyToken=b88d1754d700e49a" Namespace="DevExpress.Web.ASPxGridView" TagPrefix="dx" %>
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
        .style26
        {
            width: 713px;
        }
        </style>
</asp:Content>
<asp:Content ID="Content2" ContentPlaceHolderID="Content" runat="server">
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
            || currentColumnName == "colpartno" || currentColumnName == "colpartname" || currentColumnName == "coluom" || currentColumnName == "colqtybox"
            || currentColumnName == "colremaining" || currentColumnName == "colsupplierqty" || currentColumnName == "colboxqty") {
                e.cancel = true;
            }

            currentEditableVisibleIndex = e.visibleIndex;
        }
//        BatchEditStartEditing = "OnBatchEditStartEditing"
    </script>
    <table style="border: thin groove #808080; width: 100%;" width="100%">
        <tr>
            <td>
                <table style="width: 100%;">
                    <tr>
                        <td align="left">
                            <dx1:ASPxLabel ID="ASPxLabel1" runat="server" Font-Names="Tahoma" Font-Size="8pt"
                                Text="INVOICE DATE">
                            </dx1:ASPxLabel>
                        </td>
                        <td align="left">
                            <dx1:ASPxTextBox ID="txtInvoiceDate" runat="server" Width="165px" Font-Names="Tahoma"
                                Font-Size="8pt" BackColor="#CCCCCC" ReadOnly="True" Height="16px" 
                                ClientInstanceName="txtInvoiceDate">
                                <ClientSideEvents TextChanged="function(s, e) {
	                                lblerrmessage.SetText('');
                                }" />
                            </dx1:ASPxTextBox>
                        </td>
                        <td align="left">
                            &nbsp;
                        </td>
                        <td align="left">
                            <table>
                                <tr>
                                    <td style="width: 200px;">
                                        &nbsp;</td>
                                    <td  >
                                        &nbsp;</td>
                                </tr>
                            </table>
                        </td>
                        <td align="left">
                            <table>
                                <tr>
                                    <td style="width: 150px;">
                                        &nbsp;</td>
                                    <td  >
                                        &nbsp;</td>
                                </tr>
                            </table>
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
                        <td align="left" colspan="2">
                            <dx1:ASPxTextBox ID="txtaffiliatename" runat="server" Width="300px" Font-Names="Tahoma"
                                Font-Size="8pt" BackColor="#CCCCCC" ReadOnly="True" Height="16px" 
                                ClientInstanceName="txtaffiliatename">
                                <ClientSideEvents TextChanged="function(s, e) {
	                                lblerrmessage.SetText('');
                                }" />
                            </dx1:ASPxTextBox>
                        </td>
                        <td align="left">
                            <table>
                                <tr>
                                    <td style="width: 150px;">
                                        &nbsp;</td>
                                    <td>
                                        &nbsp;</td>
                                </tr>
                            </table>
                        </td>
                        <td align="left">
                            
                        </td>
                    </tr>
                    <tr>
                        <td align="left">
                            <dx1:ASPxLabel ID="ASPxLabel4" runat="server" Font-Names="Tahoma" Font-Size="8pt"
                                Text="PASI INVOICE NO.*">
                            </dx1:ASPxLabel>
                        </td>
                        <td align="left">
                            <dx1:ASPxTextBox ID="txtPasiInvoiceNo" runat="server" Width="165px" Font-Names="Tahoma"
                                Font-Size="8pt" BackColor="White" Height="16px" 
                                ClientInstanceName="txtPasiInvoiceNo">
                                <ClientSideEvents TextChanged="function(s, e) {
	                                lblerrmessage.SetText('');
                                }" />
                            </dx1:ASPxTextBox>
                        </td>
                        <td align="left">
                            <table>
                                <tr>
                                    <td style="width: 100px;">
                                        <dx1:ASPxLabel ID="ASPxLabel7" runat="server" Font-Names="Tahoma" Font-Size="8pt"
                                            Text="PAYMENT TERM">
                                        </dx1:ASPxLabel>
                                    </td>
                                    <td>
                                        <dx1:ASPxTextBox ID="txtPaymentTerm" runat="server" Width="150px" Font-Names="Tahoma"
                                            Font-Size="8pt" BackColor="White" Height="16px" MaxLength="20" 
                                            ClientInstanceName="txtPaymentTerm">
                                            <ClientSideEvents TextChanged="function(s, e) {
	                                lblerrmessage.SetText('');
                                }" />
                                        </dx1:ASPxTextBox>
                                    </td>
                                </tr>
                            </table>
                        </td>
                        <td align="left">
                            <table>
                                <tr>
                                    <td style="width: 100px;">
                                        <dx1:ASPxLabel ID="ASPxLabel8" runat="server" Font-Names="Tahoma" Font-Size="8pt"
                                            Text="DUE DATE">
                                        </dx1:ASPxLabel>
                                    </td>
                                    <td align="left">
                                        <dx1:ASPxDateEdit ID="dtDueDate" runat="server" ClientInstanceName="dtDueDate" 
                                            Font-Size="8pt" EditFormat="Custom" EditFormatString="dd MMM yyyy" Width="100px">
                                        </dx1:ASPxDateEdit>
                                    </td>
                                </tr>
                            </table>
                        </td>
                        <td align="left">
                            <table>
                                <tr>
                                    <td style="width: 150px;">
                                        <dx1:ASPxLabel ID="ASPxLabel10" runat="server" Font-Names="Tahoma" Font-Size="8pt"
                                            Text="TOTAL AMOUNT">
                                        </dx1:ASPxLabel>
                                    </td>
                                    <td align="left">
                                        <dx1:ASPxTextBox ID="txttotalamount" runat="server" Width="100px" Font-Names="Tahoma"
                                            Font-Size="8pt" BackColor="#CCCCCC" ReadOnly="True" Height="16px" MaxLength="20"
                                            ClientInstanceName="txttotalamount" DisplayFormatString="{0:n0}">
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
                            <dx1:ASPxLabel ID="ASPxLabel5" runat="server" Font-Names="Tahoma" Font-Size="8pt"
                                Text="NOTES">
                            </dx1:ASPxLabel>
                        </td>
                        <td align="left" colspan = "2">
                            <dx1:ASPxMemo ID="MmNotes" ClientInstanceName = "MmNotes" runat="server" Height="30px" Width="300px">
                            </dx1:ASPxMemo>
                        </td>
                        <td align="left">
                            <table>
                                <tr>
                                    <td style="width: 100px;">
                                        <dx1:ASPxLabel ID="ASPxLabel12" runat="server" Font-Names="Tahoma" Font-Size="8pt"
                                            Text="CURR">
                                        </dx1:ASPxLabel>
                                    </td>
                                    <td>
                <dx1:ASPxComboBox ID="cbocurr" runat="server" ClientInstanceName="cbocurr"
                                Font-Names="Tahoma" TextFormatString="{0}" Font-Size="8pt" Width="100px">
                                <ClientSideEvents SelectedIndexChanged="function(s, e) {
	                                            txtcurr.SetText(cbocurr.GetSelectedItem().GetColumnText(1));
                                                }" />
                            </dx1:ASPxComboBox>
                                    </td>
                                </tr>
                            </table>
                        </td>
                        <td align="left">
                                        <dx1:ASPxTextBox ID="txtcurr" runat="server" Width="100px" Font-Names="Tahoma"
                                            Font-Size="8pt" BackColor="#CCCCCC" ReadOnly="True" 
                                Height="16px" MaxLength="20"
                                            ClientInstanceName="txtcurr" DisplayFormatString="{0:n0}">
                                            <ClientSideEvents TextChanged="function(s, e) {
	                                lblerrmessage.SetText('');
                                }" />
                                        </dx1:ASPxTextBox>
                                    </td>                        
                        <td align="left">
                            &nbsp;
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
                <table style="width: 100%;" width="100%">
                    <tr>
                        <td align="right" class="style26" width="100%">
                            &nbsp;
                        </td>
                        <td align="right" class="style26" width="100%">
                            <dx1:ASPxTextBox ID="txtsupplier8" runat="server" Width="20px" Font-Names="Verdana"
                                Font-Size="8pt" BackColor="#FF66CC" Height="16px">
                                <ClientSideEvents TextChanged="function(s, e) {
	                                lblerrmessage.SetText('');
                                }" />
                            </dx1:ASPxTextBox>
                        </td>
                        <td width="100%">
                            <dx1:ASPxLabel ID="ASPxLabel11" runat="server" Font-Names="Tahoma" Font-Size="8pt"
                                Text="DIFFERENCE" Width="200px">
                            </dx1:ASPxLabel>
                        </td>
                    </tr>
                </table>
            </td>
        </tr>
        <tr>
            <td align="left">
                <dx:ASPxGridView ID="Grid" runat="server" AutoGenerateColumns="False" KeyFieldName="colorderno;colpartno;shippingno"
                    Width="100%" ClientInstanceName="Grid">
                    <Settings ShowGroupButtons="False" ShowVerticalScrollBar="True" ShowHorizontalScrollBar="True"
                        VerticalScrollableHeight="150" ShowStatusBar="Hidden"></Settings>
                    <ClientSideEvents CallbackError="function(s, e) {
e.handled = true;
}" EndCallback="function(s, e) {

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

                    <Columns>
                        <dx:GridViewDataTextColumn Caption="NO." FieldName="colno" Name="colno" VisibleIndex="0"
                            Width="30px">
                            <HeaderStyle Font-Names="Tahoma" Font-Size="8pt" Wrap="True" HorizontalAlign="Center" />
                            <CellStyle Font-Names="Tahoma" Font-Size="8pt">
                            </CellStyle>
                        </dx:GridViewDataTextColumn>
                        <dx:GridViewDataTextColumn Caption="ORDER NO." FieldName="colorderno" VisibleIndex="1"
                            Width="100px">
                            <HeaderStyle Font-Names="Verdana" Font-Size="8pt" Wrap="True" HorizontalAlign="Center" />
                            <CellStyle Font-Names="Verdana" Font-Size="8pt">
                            </CellStyle>
                        </dx:GridViewDataTextColumn>
                        <dx:GridViewDataTextColumn Caption="SUPPLIER SURAT JALAN" FieldName="colsj"
                            VisibleIndex="2" Width="100px">
                            <HeaderStyle Font-Names="Verdana" Font-Size="8pt" Wrap="True" HorizontalAlign="Center" />
                            <CellStyle Font-Names="Verdana" Font-Size="8pt">
                            </CellStyle>
                        </dx:GridViewDataTextColumn>
                        <dx:GridViewDataTextColumn Caption="PART NO." FieldName="colpartno" Name="colpartno"
                            VisibleIndex="3" Width="100px">
                            <HeaderStyle Font-Names="Verdana" Font-Size="8pt" Wrap="True" HorizontalAlign="Center" />
                            <CellStyle Font-Names="Verdana" Font-Size="8pt">
                            </CellStyle>
                        </dx:GridViewDataTextColumn>
                        <dx:GridViewDataTextColumn Caption="PART NAME" FieldName="colpartname" Name="colpartname"
                            VisibleIndex="4" Width="150px">
                            <HeaderStyle Font-Names="Verdana" Font-Size="8pt" Wrap="True" HorizontalAlign="Center" />
                            <CellStyle Font-Names="Verdana" Font-Size="8pt">
                            </CellStyle>
                        </dx:GridViewDataTextColumn>
                        <dx:GridViewDataTextColumn Caption="UOM" FieldName="coluom" Name="coluom" VisibleIndex="5"
                            Width="60px">
                            <HeaderStyle Font-Names="Verdana" Font-Size="8pt" Wrap="True" HorizontalAlign="Center" />
                            <CellStyle Font-Names="Verdana" Font-Size="8pt">
                            </CellStyle>
                        </dx:GridViewDataTextColumn>
                        <dx:GridViewDataTextColumn Caption="QTY/BOX" FieldName="colqtybox"
                            VisibleIndex="6" Width="70px">
                            <HeaderStyle Font-Names="Verdana" Font-Size="8pt" Wrap="True" HorizontalAlign="Center" />
                            <CellStyle Font-Names="Verdana" Font-Size="8pt">
                            </CellStyle>
                        </dx:GridViewDataTextColumn>
                        <dx:GridViewDataTextColumn Caption="SUPPLIER DELIVERY QTY" FieldName="coldeliveryqty"
                            Name="coldeliveryqty" VisibleIndex="7" Width="100px">
                            <HeaderStyle Font-Names="Verdana" Font-Size="8pt" Wrap="True" HorizontalAlign="Center" />
                            <CellStyle Font-Names="Verdana" Font-Size="8pt">
                            </CellStyle>
                        </dx:GridViewDataTextColumn>
                        <dx:GridViewDataTextColumn Caption="PASI INVOICE QTY" FieldName="colinvqty"
                            Name="colinvqty" VisibleIndex="8" Width="100px">
                            <HeaderStyle Font-Names="Verdana" Font-Size="8pt" Wrap="True" HorizontalAlign="Center" />
                            <CellStyle Font-Names="Verdana" Font-Size="8pt">
                            </CellStyle>
                        </dx:GridViewDataTextColumn>
                        <dx:GridViewDataTextColumn Caption="DELIVERY QTY (BOX)" FieldName="coldelqtybox"
                            Name="coldelqtybox" VisibleIndex="9" Width="100px">
                            <HeaderStyle Font-Names="Verdana" Font-Size="8pt" Wrap="True" HorizontalAlign="Center" />
                            <CellStyle Font-Names="Verdana" Font-Size="8pt">
                            </CellStyle>
                        </dx:GridViewDataTextColumn>

                        <dx:GridViewBandColumn Caption="INVOICE TO AFFILIATE" VisibleIndex="11">
                            <Columns>
                            <dx:GridViewDataTextColumn Caption="CURR" FieldName="colInvCurr" Name="colInvCurr"
                                VisibleIndex="0" Width="70px">
                                <HeaderStyle Font-Names="Tahoma" Font-Size="8pt" HorizontalAlign="Center" VerticalAlign="Middle"
                                    Wrap="True" />
                                <CellStyle Font-Names="Tahoma" Font-Size="8pt">
                                </CellStyle>
                            </dx:GridViewDataTextColumn>
                            <dx:GridViewDataTextColumn Caption="PRICE" FieldName="colInvPrice" Name="colInvPrice"
                                VisibleIndex="1" Width="100px">
                                <PropertiesTextEdit DisplayFormatString="{0:n0}">
                                    <MaskSettings ErrorText="Please input valid value !" Mask="<0..9999999999999g>.<00..99>"
                                        IncludeLiterals="DecimalSymbol" />
                                </PropertiesTextEdit>
                                <HeaderStyle Font-Names="Tahoma" Font-Size="8pt" HorizontalAlign="Center" VerticalAlign="Middle"
                                    Wrap="True" />
                                <CellStyle Font-Names="Tahoma" Font-Size="8pt">
                                </CellStyle>
                            </dx:GridViewDataTextColumn>
                            <dx:GridViewDataTextColumn Caption="AMOUNT" FieldName="colInvAmount" Name="colInvAmount"
                                VisibleIndex="2" Width="100px">
                                <PropertiesTextEdit DisplayFormatString="{0:n0}">
                                </PropertiesTextEdit>
                                <HeaderStyle Font-Names="Tahoma" Font-Size="8pt" HorizontalAlign="Center" VerticalAlign="Middle"
                                    Wrap="True" />
                                <CellStyle Font-Names="Tahoma" Font-Size="8pt">
                                </CellStyle>
                            </dx:GridViewDataTextColumn>
                            
                            </Columns>
                            <HeaderStyle HorizontalAlign="Center" />
                        </dx:GridViewBandColumn>
                        <dx:GridViewDataTextColumn FieldName="shippingno" VisibleIndex="10" Width="0px">
                        </dx:GridViewDataTextColumn>
                    </Columns>

                    <SettingsPager Visible="False" Position="Top">
                        <Summary Text="Page {0} of {1} [{2} record(s)]" AllPagesText="Page {0} of {1} " />
                        <Summary AllPagesText="Page {0} of {1} " Text="Page {0} of {1} [{2} record(s)]">
                        </Summary>
                    </SettingsPager>
                    <SettingsEditing Mode="Batch" NewItemRowPosition="Bottom">
                    </SettingsEditing>
                    <Settings ShowHorizontalScrollBar="True" ShowVerticalScrollBar="True" ShowGroupButtons="False"
                        ShowStatusBar="Hidden" VerticalScrollableHeight="135" />
                    <StylesEditors ButtonEditCellSpacing="0">
                        <ProgressBar Height="21px">
                        </ProgressBar>
                    </StylesEditors>
                </dx:ASPxGridView>
            </td>
        </tr>
        <tr>
            <td align="left">
                &nbsp;
            </td>
        </tr>
        <tr>
            <td align="left">
                <table style="width: 100%;">
                    <tr>
                        <td align="left">
                            <dx1:ASPxButton ID="btnsubmenu" runat="server" Width="90px" Font-Names="Tahoma" Font-Size="8pt"
                                Text="SUB MENU">
                                <ClientSideEvents Click="function(s, e) {

}" />
                            </dx1:ASPxButton>
                        </td>     
                        <td style="width: 90px;">
                            &nbsp;</td>  
                        <td style="width: 90px;">
                            &nbsp;</td>                
                         <td style="width: 90px;">
                             &nbsp;</td>
                        <td style="width: 90px;">
                            <dx1:ASPxButton ID="SendEDI" runat="server" Width="90px" Font-Names="Tahoma" Font-Size="8pt"
                                Text="SEND E.D.I" ClientInstanceName="SendEDI" AutoPostBack="False">
                                <ClientSideEvents Click="function(s, e) {
											var pInvoiceno = txtPasiInvoiceNo.GetText();
                                            var pAffiliateID = txtaffiliatecode.GetText();
                                                            
	                                        Grid.PerformCallback('EDI' + '|' + pInvoiceno  + '|' + pAffiliateID );
	                                        lblerrmessage.SetText('');
}" />
                            </dx1:ASPxButton>
                        </td>
                    </tr>
                </table>
            </td>
        </tr>
    </table>
</asp:Content>
