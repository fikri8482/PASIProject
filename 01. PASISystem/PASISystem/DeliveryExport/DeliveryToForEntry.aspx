<%@ Page Title="" Language="vb" AutoEventWireup="false" MasterPageFile="~/SiteContent.Master" CodeBehind="DeliveryToForEntry.aspx.vb" Inherits="PASISystem.DeliveryToForEntry" %>
<%@ Register Assembly="DevExpress.Web.v14.1, Version=14.1.6.0, Culture=neutral, PublicKeyToken=b88d1754d700e49a" Namespace="DevExpress.Web.ASPxGridView.Export" TagPrefix="dx" %>
<%@ Register Assembly="DevExpress.Web.v14.1, Version=14.1.6.0, Culture=neutral, PublicKeyToken=b88d1754d700e49a" Namespace="DevExpress.Web.ASPxEditors" TagPrefix="dx1" %>
<%@ Register Assembly="DevExpress.Web.v14.1, Version=14.1.6.0, Culture=neutral, PublicKeyToken=b88d1754d700e49a" Namespace="DevExpress.Web.ASPxGridView" TagPrefix="dx" %>
<%@ Register Assembly="DevExpress.Web.v14.1, Version=14.1.6.0, Culture=neutral, PublicKeyToken=b88d1754d700e49a" Namespace="DevExpress.Web.ASPxHiddenField" TagPrefix="dx" %>
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
    <script language="javascript" type="text/javascript">
        function OnAllCheckedChanged(s, e) {
            if (s.GetValue() == -1) s.SetValue(1);
            for (var i = 0; i < Grid.GetVisibleRowsOnPage(); i++) {
                Grid.batchEditApi.SetCellValue(i, "colno", s.GetValue());
            }
        }

        
        function OnUpdateClick(s, e) {
            Grid.PerformCallback("Update");
        }

        function OnCancelClick(s, e) {
            Grid.PerformCallback("Cancel");
        }

        function OnBatchEditStartEditing(s, e) {
            currentColumnName = e.focusedColumn.fieldName;

            if (currentColumnName == "colpono" || currentColumnName == "colpokanban" || currentColumnName == "colkanbanno"
            || currentColumnName == "colpartno" || currentColumnName == "colpartname" || currentColumnName == "coluom" || currentColumnName == "colqtybox"
            || currentColumnName == "colremaining" || currentColumnName == "colsupplierqty" || currentColumnName == "colboxqty") {
                e.cancel = true;
            }

            currentEditableVisibleIndex = e.visibleIndex;
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
                <dx1:ASPxDateEdit ID="dt1" runat="server" ClientInstanceName="dt1" Font-Names="Tahoma"
                    Font-Size="8pt" EditFormat="Custom" EditFormatString="dd MMM yyyy">
                </dx1:ASPxDateEdit>
                        </td>
                        <td align="left">
                            &nbsp;
                            <dx1:ASPxTextBox ID="txtdeliverydate" runat="server" Width="165px" Font-Names="Tahoma"
                                Font-Size="8pt" BackColor="#CCCCCC" ReadOnly="True" Height="16px" 
                                ClientInstanceName="txtdeliverydate" Visible="False">
                                <ClientSideEvents TextChanged="function(s, e) {
	                                lblerrmessage.SetText('');
                                }" />
                            </dx1:ASPxTextBox>
                        </td>
                        <td align="left">
                            <dx1:ASPxLabel ID="ASPxLabel7" runat="server" Font-Names="Tahoma" Font-Size="8pt"
                                Text="SUPPLIER CODE">
                            </dx1:ASPxLabel>
                        </td>
                        <td align="left">
                            <dx1:ASPxTextBox ID="txtSupplierCode" runat="server" Width="165px" Font-Names="Tahoma"
                                Font-Size="8pt" BackColor="#CCCCCC" ReadOnly="True" Height="16px" 
                                ClientInstanceName="txtSupplierCode">
                                <ClientSideEvents TextChanged="function(s, e) {
	                                lblerrmessage.SetText('');
                                }" />
                            </dx1:ASPxTextBox>
                        </td>
                        <td align="left">
                            <dx1:ASPxTextBox ID="txtSupplierName" runat="server" Width="165px" Font-Names="Tahoma"
                                Font-Size="8pt" BackColor="#CCCCCC" ReadOnly="True" Height="16px" 
                                ClientInstanceName="txtSupplierName">
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
                                Text="DELIVERY LOCATION" Visible="False">
                            </dx1:ASPxLabel>
                        </td>
                        <td align="left">
                            <dx1:ASPxTextBox ID="txtdeliverylocationCode" runat="server" Width="165px" Font-Names="Tahoma"
                                Font-Size="8pt" BackColor="#CCCCCC" ReadOnly="True" Height="16px" 
                                ClientInstanceName="txtdeliverylocationCode" Visible="False">
                                <ClientSideEvents TextChanged="function(s, e) {
	                                lblerrmessage.SetText('');
                                }" />
                            </dx1:ASPxTextBox>
                        </td>
                        <td align="left">
                            <dx1:ASPxTextBox ID="txtdeliverylocationName" runat="server" Width="165px" Font-Names="Tahoma"
                                Font-Size="8pt" BackColor="#CCCCCC" ReadOnly="True" Height="16px" 
                                ClientInstanceName="txtdeliverylocationName" Visible="False">
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
                            <dx1:ASPxLabel ID="ASPxLabel15" runat="server" Font-Names="Tahoma" Font-Size="8pt"
                                Text="FORWARDER CODE/NAME">
                            </dx1:ASPxLabel>
                        </td>
                        <td align="left">
                            <dx1:ASPxTextBox ID="txtForwarderCode" runat="server" Width="165px" Font-Names="Tahoma"
                                Font-Size="8pt" BackColor="#CCCCCC" ReadOnly="True" Height="16px" 
                                ClientInstanceName="txtForwarderCode">
                                <ClientSideEvents TextChanged="function(s, e) {
	                                lblerrmessage.SetText('');
                                }" />
                            </dx1:ASPxTextBox>
                        </td>
                        <td align="left">
                            <dx1:ASPxTextBox ID="txtForwarderName" runat="server" Width="165px" Font-Names="Tahoma"
                                Font-Size="8pt" BackColor="#CCCCCC" ReadOnly="True" Height="16px" 
                                ClientInstanceName="txtForwarderName">
                                <ClientSideEvents TextChanged="function(s, e) {
	                                lblerrmessage.SetText('');
                                }" />
                            </dx1:ASPxTextBox>
                        </td>
                        <td align="left">
                            &nbsp;</td>
                        <td align="left">
                            &nbsp;</td>
                        <td align="left">
                            &nbsp;</td>
                        <td align="left">
                            &nbsp;</td>
                    </tr>
                    <tr>
                        <td align="left">
                            <dx1:ASPxLabel ID="ASPxLabel3" runat="server" Font-Names="Tahoma" Font-Size="8pt"
                                Text="PASI SURAT JALAN NO.*">
                            </dx1:ASPxLabel>
                        </td>
                        <td align="left">
                            <dx1:ASPxTextBox ID="txtsuratjalanno" runat="server" Width="165px" Font-Names="Tahoma"
                                Font-Size="8pt" Height="16px" ClientInstanceName="txtsuratjalanno">
                                <ClientSideEvents TextChanged="function(s, e) {
	                                lblerrmessage.SetText('');
                                }" />
                            </dx1:ASPxTextBox>
                        </td>
                        <td align="left">
                            <table style="width: 100%;">
                                <tr>
                                    <td style="width: 150px;">
                                        <dx1:ASPxLabel ID="ASPxLabel12" runat="server" Font-Names="Tahoma" Font-Size="8pt"
                                            Text="INVOICE NO">
                                        </dx1:ASPxLabel>
                                    </td>
                                    <td>
                                        <dx1:ASPxTextBox ID="txtInvoiceNo" runat="server" Width="100px" Font-Names="Tahoma"
                                            Font-Size="8pt" BackColor="White" Height="16px" MaxLength="20" ClientInstanceName="txtInvoiceNo">
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
                        <td align="left">
                            &nbsp;
                        </td>
                        <td align="left">
                            <table style="width: 100%;">
                                <tr>
                                    <td width="80px">
                                        <dx1:ASPxLabel ID="ASPxLabel14" runat="server" Font-Names="Tahoma" Font-Size="8pt"
                                            Text="TOTAL PALLET">
                                        </dx1:ASPxLabel>
                                    </td>
                                    <td>
                                        <dx1:ASPxTextBox ID="txttotalpalet" runat="server" Width="80px" Font-Names="Tahoma"
                                            Font-Size="8pt" BackColor="White" Height="16px" MaxLength="20" 
                                            ClientInstanceName="txttotalpalet" DisplayFormatString="{0}" 
                                            HorizontalAlign = "Right">
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
                            <dx1:ASPxLabel ID="ASPxLabel4" runat="server" Font-Names="Tahoma" Font-Size="8pt"
                                Text="DRIVER NAME">
                            </dx1:ASPxLabel>
                        </td>
                        <td align="left">
                            <dx1:ASPxTextBox ID="txtdrivername" runat="server" Width="165px" Font-Names="Tahoma"
                                Font-Size="8pt" BackColor="White" Height="16px" 
                                ClientInstanceName="txtdrivername" MaxLength="15">
                                <ClientSideEvents TextChanged="function(s, e) {
	                                lblerrmessage.SetText('');
                                }" />
                            </dx1:ASPxTextBox>
                        </td>
                        <td align="left">
                            <table style="width: 100%;">
                                <tr>
                                    <td style="width: 150px;">
                                        <dx1:ASPxLabel ID="ASPxLabel6" runat="server" Font-Names="Tahoma" Font-Size="8pt"
                                            Text="DRIVER CONTACT">
                                        </dx1:ASPxLabel>
                                    </td>
                                    <td>
                                        <dx1:ASPxTextBox ID="txtdrivercontact" runat="server" Width="100px" Font-Names="Tahoma"
                                            Font-Size="8pt" BackColor="White" Height="16px" MaxLength="15" 
                                            ClientInstanceName="txtdrivercontact">
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
                                    <td>
                                        <dx1:ASPxLabel ID="ASPxLabel8" runat="server" Font-Names="Tahoma" Font-Size="8pt"
                                            Text="NO. POL">
                                        </dx1:ASPxLabel>
                                    </td>
                                    <td>
                                        <dx1:ASPxTextBox ID="txtnopol" runat="server" Width="100px" Font-Names="Tahoma" Font-Size="8pt"
                                            BackColor="White" Height="16px" MaxLength="10" 
                                            ClientInstanceName="txtnopol">
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
                                    <td>
                                        <dx1:ASPxLabel ID="ASPxLabel9" runat="server" Font-Names="Tahoma" Font-Size="8pt"
                                            Text="JENIS ARMADA">
                                        </dx1:ASPxLabel>
                                    </td>
                                    <td>
                                        <dx1:ASPxTextBox ID="txtjenisarmada" runat="server" Width="100px" Font-Names="Tahoma"
                                            Font-Size="8pt" BackColor="White" Height="16px" MaxLength="15" 
                                            ClientInstanceName="txtjenisarmada">
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
                                    <td width="80px">
                                        <dx1:ASPxLabel ID="ASPxLabel10" runat="server" Font-Names="Tahoma" Font-Size="8pt"
                                            Text="TOTAL BOX        ">
                                        </dx1:ASPxLabel>
                                    </td>
                                    <td>
                                        <dx1:ASPxTextBox ID="txttotalbox" runat="server" Width="80px" Font-Names="Tahoma"
                                            Font-Size="8pt" BackColor="#CCCCCC" ReadOnly="True" Height="16px" MaxLength="20" 
                                            ClientInstanceName="txttotalbox" DisplayFormatString="{0}" 
                                            HorizontalAlign = "Right">
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
                                Text=": NOT SAVE" >
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
                                Text=": DIFFERENCE" >
                            </dx1:ASPxLabel>
                        </td>
                    </tr>
                </table>
            </td>
        </tr>
        <tr>
            <td align="left">
                <dx:ASPxGridView ID="Grid" runat="server" AutoGenerateColumns="False" KeyFieldName="colno;colpono;colpartno;suppsj"
                    Width="100%" ClientInstanceName="Grid">
                    <ClientSideEvents BatchEditStartEditing="OnBatchEditStartEditing" CallbackError="function(s, e) {
                        e.handled = true;
                        }" EndCallback="function(s, e) {

                        txttotalbox.SetText(s.cptotalbox);

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

                    <Settings ShowHorizontalScrollBar="True" ShowVerticalScrollBar="True" ShowGroupButtons="False"
                        ShowStatusBar="Hidden" VerticalScrollableHeight="135" />
                    
                    <Columns>
                        <dx:GridViewDataCheckColumn Caption=" " FieldName="colno" Name="colno" 
                            VisibleIndex="1" Width="30px">
                            <HeaderStyle Font-Names="Tahoma" Font-Size="8pt" HorizontalAlign="Center" 
                                Wrap="True" />
                            <HeaderCaptionTemplate>
                                <dx1:ASPxCheckBox ID="chkAll" runat="server" ClientInstanceName="chkAll" ClientSideEvents-CheckedChanged="OnAllCheckedChanged"
                                ValueType="System.Int32" ValueChecked="1" ValueUnchecked="0">
                                </dx1:ASPxCheckBox>
                            </HeaderCaptionTemplate>

                            <PropertiesCheckEdit ValueChecked="1" ValueType="System.Int32" ValueUnchecked="0">
                            </PropertiesCheckEdit>
                            <CellStyle Font-Names="Tahoma" Font-Size="8pt">
                            </CellStyle>
                        </dx:GridViewDataCheckColumn>
                        <dx:GridViewDataTextColumn Caption="PO NO." FieldName="colpono" Name="colpono" VisibleIndex="2"
                            Width="100px">
                            <HeaderStyle Font-Names="Verdana" Font-Size="8pt" Wrap="True" HorizontalAlign="Center" />
                            <CellStyle Font-Names="Verdana" Font-Size="8pt">
                            </CellStyle>
                        </dx:GridViewDataTextColumn>
                        <dx:GridViewDataTextColumn Caption="PART NO." FieldName="colpartno" Name="colpartno"
                            VisibleIndex="5" Width="100px">
                            <HeaderStyle Font-Names="Verdana" Font-Size="8pt" Wrap="True" HorizontalAlign="Center" />
                            <CellStyle Font-Names="Verdana" Font-Size="8pt">
                            </CellStyle>
                        </dx:GridViewDataTextColumn>
                        <dx:GridViewDataTextColumn Caption="PART NAME" FieldName="colpartname" 
                            Name="colpartname" VisibleIndex="6"
                            Width="150px">
                            <HeaderStyle Font-Names="Verdana" Font-Size="8pt" Wrap="True" 
                                HorizontalAlign="Center" />
                            <CellStyle Font-Names="Verdana" Font-Size="8pt">
                            </CellStyle>
                        </dx:GridViewDataTextColumn>
                        <dx:GridViewDataTextColumn Caption="UOM" FieldName="coluom" Name="coluom" VisibleIndex="7"
                            Width="60px">
                            <HeaderStyle Font-Names="Verdana" Font-Size="8pt" Wrap="True" HorizontalAlign="Center" />
                            <CellStyle Font-Names="Verdana" Font-Size="8pt">
                            </CellStyle>
                        </dx:GridViewDataTextColumn>
                        <dx:GridViewDataTextColumn Caption="QTY/BOX" FieldName="colQtyBox" Name="colQtyBox"
                            VisibleIndex="8" Width="70px">
                            <HeaderStyle Font-Names="Verdana" Font-Size="8pt" Wrap="True" HorizontalAlign="Center" />
                            <CellStyle Font-Names="Verdana" Font-Size="8pt">
                            </CellStyle>
                        </dx:GridViewDataTextColumn>
                        <dx:GridViewDataTextColumn Caption="SUPPLIER DELIVERY QTY" FieldName="colsuppdelqty"
                            Name="colsuppdelqty" VisibleIndex="9" Width="100px">
                            <HeaderStyle Font-Names="Verdana" Font-Size="8pt" Wrap="True" HorizontalAlign="Center" />
                            <CellStyle Font-Names="Verdana" Font-Size="8pt" HorizontalAlign="Right">
                            </CellStyle>
                        </dx:GridViewDataTextColumn>
                        <dx:GridViewDataTextColumn Caption="PASI GOOD RECEIVING QTY" FieldName="colpasigoodrec"
                            Name="colpasigoodrec" VisibleIndex="10" Width="100px">
                            <PropertiesTextEdit >
                                
<MaskSettings ErrorText="Please input valid value !" Mask="<0..9999999999999g>.<00..99>"
                                    IncludeLiterals="DecimalSymbol" />
                            
</PropertiesTextEdit>
                            <HeaderStyle Font-Names="Verdana" Font-Size="8pt" Wrap="True" HorizontalAlign="Center" />
                            <CellStyle Font-Names="Verdana" Font-Size="8pt" HorizontalAlign="Right" >
                            </CellStyle>
                        </dx:GridViewDataTextColumn>
                        <dx:GridViewDataTextColumn Caption="PASI DEFECT RECEIVING" FieldName="colpasidefectrec"
                            Name="colpasidefectrec" VisibleIndex="11" Width="100px">
                            <HeaderStyle Font-Names="Verdana" Font-Size="8pt" Wrap="True" HorizontalAlign="Center" />
                            <CellStyle Font-Names="Verdana" Font-Size="8pt" HorizontalAlign="Right">
                            </CellStyle>
                        </dx:GridViewDataTextColumn>
                        <dx:GridViewDataTextColumn Caption="PASI REMAINING RECEIVING" 
                            FieldName="colpasiremaining" Name="colpasiremaining"
                            VisibleIndex="12" Width="100px">
                            <HeaderStyle Font-Names="Tahoma" Font-Size="8pt" Wrap="True" 
                                HorizontalAlign="Center" VerticalAlign="Middle" />
                            <CellStyle Font-Names="Tahoma" Font-Size="8pt" HorizontalAlign="Right">
                            </CellStyle>
                        </dx:GridViewDataTextColumn>
                        <dx:GridViewDataTextColumn Caption="PASI DELIVERY QTY*" 
                            FieldName="colpasideliveryqty" Name="colpasideliveryqty"
                            VisibleIndex="13" Width="70px">
                            <PropertiesTextEdit >
                                <MaskSettings ErrorText="Please input valid value !" Mask="<0..9999999999999g>.<00..99>"
                                           IncludeLiterals="DecimalSymbol" />
                            </PropertiesTextEdit>
                            <HeaderStyle Font-Names="Verdana" Font-Size="8pt" Wrap="True" HorizontalAlign="Center" />
                            <CellStyle Font-Names="Verdana" Font-Size="8pt">
                            </CellStyle>
                        </dx:GridViewDataTextColumn>
                        <dx:GridViewDataTextColumn Caption="REMAINING DELIVERY QTY" 
                            FieldName="colremainingdelqty" Name="colremainingdelqty"
                            VisibleIndex="14" Width="100px">
                            <PropertiesTextEdit >
                                
<MaskSettings ErrorText="Please input valid value !" Mask="<0..9999999999999g>.<00..99>"
                                    IncludeLiterals="DecimalSymbol" />
                            
</PropertiesTextEdit>
                            <HeaderStyle Font-Names="Tahoma" Font-Size="8pt" HorizontalAlign="Center" VerticalAlign="Middle"
                                Wrap="True" />
                            <CellStyle Font-Names="Tahoma" Font-Size="8pt">
                            </CellStyle>
                        </dx:GridViewDataTextColumn>
                        <dx:GridViewDataTextColumn Caption="DELIVERY QTY (BOX)" 
                            FieldName="coldelqtybox" Name="coldelqtybox"
                            VisibleIndex="15" Width="100px">
                        </dx:GridViewDataTextColumn>
                        <dx:GridViewDataTextColumn FieldName="colCls" Name="colCls"
                            VisibleIndex="17" Width="0px">
                            <HeaderStyle Font-Names="Tahoma" Font-Size="8pt" HorizontalAlign="Center" VerticalAlign="Middle"
                                Wrap="True"/>
                            <CellStyle Font-Names="Tahoma" Font-Size="8pt">
                            </CellStyle>
                        </dx:GridViewDataTextColumn>
                        <dx:GridViewDataTextColumn FieldName="colstsDO" Name="colstsDO"
                            VisibleIndex="18" Width="0px">
                        </dx:GridViewDataTextColumn>
                        <dx:GridViewDataTextColumn FieldName="suppsj" Name="suppsj"
                            VisibleIndex="19" Width="0px" Caption="SupSJ">
                        </dx:GridViewDataTextColumn>                        
                        <dx:GridViewDataTextColumn FieldName="SupplierID" Name="SupplierID" 
                            VisibleIndex="20" Width="0px">
                        </dx:GridViewDataTextColumn>
                        <dx:GridViewDataTextColumn Caption="colForwarder" FieldName="colForwarder" 
                            Name="colForwarder" VisibleIndex="16" Width="0px">
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
                            <dx1:ASPxButton ID="btnPrint" runat="server" Width="90px" Font-Names="Tahoma" Font-Size="8pt"
                                Text="PRINT" ClientInstanceName="btnPrint">
                                <ClientSideEvents Click="function(s, e) {
	if (txtsuratjalanno.GetText() == '') {
        lblerrmessage.SetText('[6011] No Data To Print!');
		lblerrmessage.GetMainElement().style.color = 'Red';
        return false;
    }
}" />
                            </dx1:ASPxButton>
                        </td>                
                         <td style="width: 90px;">
                            <dx1:ASPxButton ID="btndelete" runat="server" Width="90px" Font-Names="Tahoma" Font-Size="8pt"
                                Text="DELETE" ClientInstanceName="btndelete" AutoPostBack="False" 
                                 Visible="False">
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
	Grid.PerformCallback('Delete');
    txtsuratjalanno.SetText('');
    txtdrivername.SetText('');
    txtdrivercontact.SetText('');
    txtnopol.SetText('');
    txtjenisarmada.SetText('');
    txtInvoiceNo.SetText('');
    txttotalbox.SetText('');
    txttotalpalet.SetText('');
}" />
                            </dx1:ASPxButton>
                        </td>
                        <td style="width: 90px;">
                            <dx1:ASPxButton ID="btnsubmit" runat="server" Width="90px" Font-Names="Tahoma" Font-Size="8pt"
                                Text="SAVE" ClientInstanceName="btnsubmit" AutoPostBack="False">
                                <ClientSideEvents Click="function(s, e) {
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
	Grid.UpdateEdit();
    var millisecondsToWait = 50;
            setTimeout(function() {
                Grid.PerformCallback('gridload');
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
