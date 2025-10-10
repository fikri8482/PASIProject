<%@ Page Title="" Language="vb" AutoEventWireup="false" MasterPageFile="~/SiteContent.Master"
    CodeBehind="InvFromPASIDetailExport.aspx.vb" Inherits="PASISystem.InvFromPASIDetailExport" %>

<%@ Register Assembly="DevExpress.Web.v14.1, Version=14.1.6.0, Culture=neutral, PublicKeyToken=b88d1754d700e49a"
    Namespace="DevExpress.Web.ASPxGridView.Export" TagPrefix="dx" %>
<%@ Register Assembly="DevExpress.Web.v14.1, Version=14.1.6.0, Culture=neutral, PublicKeyToken=b88d1754d700e49a"
    Namespace="DevExpress.Web.ASPxEditors" TagPrefix="dx1" %>
<%@ Register Assembly="DevExpress.Web.v14.1, Version=14.1.6.0, Culture=neutral, PublicKeyToken=b88d1754d700e49a"
    Namespace="DevExpress.Web.ASPxGridView" TagPrefix="dx" %>
<asp:Content ID="Content1" ContentPlaceHolderID="head" runat="server">
    <style type="text/css">
        .dxeBase
        {
            font: 12px Tahoma, Geneva, sans-serif;
            margin-left: 0px;
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
        .style27
        {
            width: 704px;
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

            currentEditableVisibleIndex = e.visibleIndex;
        }

    </script>
    <table style="border: thin groove #808080; width: 100%;" width="100%">
        <tr>
            <td width="100%">
                <table style="width: 100%;" width="100%">
                    <tr>
                        <td align="left">
                            <dx1:ASPxLabel ID="ASPxLabel1" runat="server" Font-Names="Tahoma" Font-Size="8pt"
                                Text="INVOICE DATE">
                            </dx1:ASPxLabel>
                        </td>
                        <td align="left">
                            <table style="width:100%;">
                                <tr>
                                    <td>
                            <dx1:ASPxTextBox ID="txtinvdate" runat="server" Width="165px" Font-Names="Tahoma"
                                Font-Size="8pt" BackColor="#CCCCCC" ReadOnly="True" Height="16px" 
                                ClientInstanceName="txtinvdate">
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
                            &nbsp;<dx1:ASPxTextBox ID="txtkanbanno" runat="server" Width="20px" Font-Names="Tahoma"
                                Font-Size="8pt" BackColor="#CCCCCC" ReadOnly="True" Height="16px" 
                                ClientInstanceName="txtkanbanno" Visible="False">
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
                            <table style="width:100%;">
                                <tr>
                                    <td>
                            <dx1:ASPxTextBox ID="txtaffiliatecode" runat="server" Width="165px" Font-Names="Tahoma"
                                Font-Size="8pt" BackColor="#CCCCCC" ReadOnly="True" Height="16px" 
                                ClientInstanceName="txtaffiliatename">
                                <ClientSideEvents TextChanged="function(s, e) {
	                                lblerrmessage.SetText('');
                                }" />
                            </dx1:ASPxTextBox>
                                    </td>
                                    <td>
                            <dx1:ASPxTextBox ID="txtaffiliatename" runat="server" Width="165px" Font-Names="Tahoma"
                                Font-Size="8pt" BackColor="#CCCCCC" ReadOnly="True" Height="16px" 
                                ClientInstanceName="txtaffiliatename">
                                <ClientSideEvents TextChanged="function(s, e) {
	                                lblerrmessage.SetText('');
                                }" />
                            </dx1:ASPxTextBox>
                                    </td>
                                </tr>
                            </table>
                        </td>
                        <td align="left">
                            &nbsp;</td>
                        <td align="left">
                            &nbsp;</td>
                        <td align="left">
                            <dx1:ASPxTextBox ID="txtsupplier" runat="server" Width="20px" Font-Names="Tahoma"
                                Font-Size="8pt" BackColor="#CCCCCC" ReadOnly="True" Height="16px" 
                                ClientInstanceName="txtsupplier" Visible="False">
                                <ClientSideEvents TextChanged="function(s, e) {
	                                lblerrmessage.SetText('');
                                }" />
                            </dx1:ASPxTextBox>
                        </td>
                        <td align="left">
                            &nbsp;</td>
                    </tr>
                    <tr>
                        <td align="left">
                            <dx1:ASPxLabel ID="ASPxLabel3" runat="server" Font-Names="Tahoma" Font-Size="8pt"
                                Text="PASI SURAT JALAN NO.">
                            </dx1:ASPxLabel>
                        </td>
                        <td align="left">
                            <table style="width:100%;">
                                <tr>
                                    <td>
                            <dx1:ASPxTextBox ID="txtsuratjalanno" runat="server" Width="165px" Font-Names="Tahoma"
                                Font-Size="8pt" BackColor="#CCCCCC" ReadOnly="True" Height="16px" ClientInstanceName="txtsuratjalanno">
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
                            <dx1:ASPxTextBox ID="txtstatus" runat="server" Width="20px" Font-Names="Tahoma"
                                Font-Size="8pt" BackColor="#CCCCCC" ReadOnly="True" Height="16px" 
                                ClientInstanceName="txtstatus" Visible="False">
                                <ClientSideEvents TextChanged="function(s, e) {
	                                lblerrmessage.SetText('');
                                }" />
                            </dx1:ASPxTextBox>
                            </td>
                        <td align="left">
                            &nbsp;
                            </td>
                        <td align="left">
                            &nbsp;
                            <dx1:ASPxTextBox ID="txtpono" runat="server" Width="20px" Font-Names="Tahoma"
                                Font-Size="8pt" BackColor="#CCCCCC" ReadOnly="True" Height="16px" 
                                ClientInstanceName="txtpono" Visible="False">
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
                            <dx1:ASPxLabel ID="ASPxLabel4" runat="server" Font-Names="Tahoma" Font-Size="8pt"
                                Text="PASI INVOICE NO">
                            </dx1:ASPxLabel>
                        </td>
                        <td align="left">
                            <table style="width:100%;">
                                <tr>
                                    <td>
                            <dx1:ASPxTextBox ID="txtinv" runat="server" Width="165px" Font-Names="Tahoma"
                                Font-Size="8pt" BackColor="Silver" Height="16px" 
                                ClientInstanceName="txtinv" ReadOnly="True">
                                <ClientSideEvents TextChanged="function(s, e) {
	                                lblerrmessage.SetText('');
                                }" />
                            </dx1:ASPxTextBox>
                                    </td>
                                    <td>
                                        <dx1:ASPxLabel ID="ASPxLabel6" runat="server" Font-Names="Tahoma" Font-Size="8pt"
                                            Text="PAYMENT TERM">
                                        </dx1:ASPxLabel>
                                    </td>
                                    <td>
                                        <dx1:ASPxTextBox ID="txtpayment" runat="server" Width="100px" Font-Names="Tahoma"
                                            Font-Size="8pt" BackColor="Silver" Height="16px" MaxLength="20" 
                                            ClientInstanceName="txtpayment" ReadOnly="True">
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
                                            Text="DUE DATE">
                                        </dx1:ASPxLabel>
                                    </td>
                                    <td>
                            <dx1:ASPxTextBox ID="txtduedate" runat="server" Width="165px" Font-Names="Tahoma"
                                Font-Size="8pt" BackColor="#CCCCCC" ReadOnly="True" Height="16px" 
                                ClientInstanceName="txtduedate">
                                <ClientSideEvents TextChanged="function(s, e) {
	                                lblerrmessage.SetText('');
                                }" />
                            </dx1:ASPxTextBox>
                                    </td>
                                </tr>
                            </table>
                        </td>
                        <td align="left">
                            &nbsp;</td>
                        <td align="left">
                            <table style="width: 100%;">
                                <tr>
                                    <td>
                                        <dx1:ASPxLabel ID="ASPxLabel10" runat="server" Font-Names="Tahoma" Font-Size="8pt"
                                            Text="TOTAL AMOUNT">
                                        </dx1:ASPxLabel>
                                    </td>
                                    <td>
                                        <dx1:ASPxTextBox ID="txttotalamount" runat="server" Width="100px" Font-Names="Tahoma"
                                            Font-Size="8pt" BackColor="#CCCCCC" ReadOnly="True" Height="16px" MaxLength="20"
                                            ClientInstanceName="txttotalamount">
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
                            <dx1:ASPxLabel ID="ASPxLabel12" runat="server" Font-Names="Tahoma" Font-Size="8pt"
                                Text="NOTES ">
                            </dx1:ASPxLabel>
                        </td>
                        <td align="left">
                            <table style="width:100%;">
                                <tr>
                                    <td>
                            <dx1:ASPxTextBox ID="txtnotes" runat="server" Width="350px" Font-Names="Tahoma"
                                Font-Size="8pt" BackColor="Silver" Height="32px" 
                                ClientInstanceName="txtnotes" ReadOnly="True">
                                <ClientSideEvents TextChanged="function(s, e) {
	                                lblerrmessage.SetText('');
                                }" />
                            </dx1:ASPxTextBox>
                                    </td>
                                </tr>
                            </table>
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
                </table>
            </td>
        </tr>
    </table>
    <table style="width: 100%;" width="100%">
        <tr>
            <td align="left">
                <table width="100%" align="left">
                    <tr>
                        <td align="left" bgcolor="White" class="style25" style="border-width: thin; border-style: inset hidden ridge hidden;"
                            width="100%" height="16px">
                            <table style="width: 100%;" width="100%">
                                <tr>
                                    <td width="100%" align="left">
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
                <table style="width: 100%;" width="100%" align="right">
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
                        <td align="right">
                            <dx1:ASPxLabel ID="ASPxLabel11" runat="server" Font-Names="Tahoma" Font-Size="8pt"
                                Text="DIFFERENCE">
                            </dx1:ASPxLabel>
                        </td>
                    </tr>
                </table>
            </td>
        </tr>
        <tr>
            <td align="left">
                <dx:ASPxGridView ID="Grid" runat="server" AutoGenerateColumns="False" KeyFieldName="no;pono;kanbanno;partno"
                    Width="100%" ClientInstanceName="Grid">
                    <ClientSideEvents BatchEditStartEditing="OnBatchEditStartEditing" CallbackError="function(s, e) {
e.handled = true;
}" EndCallback="function(s, e) {
									Grid.UpdateEdit();
                                    Grid.CancelEdit();
                                    txtinvdate.SetValue(s.cpDate);
                                    txtinv.SetText(s.cpInvoiceNo);
                                    txtaffiliatecode.SetText(s.cpaffcode);
                                    txtaffiliatename.SetText(s.cpaffname);
                                    txtsuratjalanno.SetText(s.cpsj);
                                    txtpayment.SetText(s.cppayment);
                                    dt2.SetValue(s.cpduedate);
                                    txtkanbanno.SetText(s.cpKanbanno);
                                    txtpono.SetText(s.cppono);
                                    txttotalamount.SetText(s.cptotalamount);

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
<ClientSideEvents BatchEditStartEditing="OnBatchEditStartEditing" EndCallback="function(s, e) {
									Grid.UpdateEdit();
                                    Grid.CancelEdit();
                                    txtinvdate.SetValue(s.cpDate);
                                    txtinv.SetText(s.cpInvoiceNo);
                                    txtaffiliatecode.SetText(s.cpaffcode);
                                    txtaffiliatename.SetText(s.cpaffname);
                                    txtsuratjalanno.SetText(s.cpsj);
                                    txtpayment.SetText(s.cppayment);
                                    dt2.SetValue(s.cpduedate);
                                    txtkanbanno.SetText(s.cpKanbanno);
                                    txtpono.SetText(s.cppono);
                                    txttotalamount.SetText(s.cptotalamount);

                                      var pMsg = s.cpMessage;
                                        if (pMsg != &#39;&#39;) {
                                            if (pMsg.substring(1,5) == &#39;1001&#39; || pMsg.substring(1,5) == &#39;1002&#39; || pMsg.substring(1,5) == &#39;1003&#39;) {
                                                lblerrmessage.GetMainElement().style.color = &#39;Blue&#39;;  
                                            } else {
                                                lblerrmessage.GetMainElement().style.color = &#39;Red&#39;;
                                            }
                                                lblerrmessage.SetText(pMsg);
                                            } else {
                                                lblerrmessage.SetText(&#39;&#39;);
                                            }
}" CallbackError="function(s, e) {
e.handled = true;
}"></ClientSideEvents>

                    <Columns>
                        <dx:GridViewDataTextColumn Caption="NO." FieldName="no" Name="no" VisibleIndex="0"
                            Width="30px">
                            <HeaderStyle Font-Names="Tahoma" Font-Size="8pt" Wrap="True" HorizontalAlign="Center" />
                            <CellStyle Font-Names="Tahoma" Font-Size="8pt">
                            </CellStyle>
                        </dx:GridViewDataTextColumn>
                        <dx:GridViewDataTextColumn Caption="PO NO." FieldName="pono" Name="pono" VisibleIndex="1"
                            Width="100px">
                            <HeaderStyle Font-Names="Verdana" Font-Size="8pt" Wrap="True" HorizontalAlign="Center" />
                            <CellStyle Font-Names="Verdana" Font-Size="8pt">
                            </CellStyle>
                        </dx:GridViewDataTextColumn>
                        <dx:GridViewDataTextColumn Caption="KANBAN NO." FieldName="kanbanno" Name="kanbanno"
                            VisibleIndex="3" Width="100px">
                            <HeaderStyle Font-Names="Verdana" Font-Size="8pt" Wrap="True" HorizontalAlign="Center" />
                            <CellStyle Font-Names="Verdana" Font-Size="8pt">
                            </CellStyle>
                        </dx:GridViewDataTextColumn>
                        <dx:GridViewDataTextColumn Caption="PART NO." FieldName="partno" Name="partno"
                            VisibleIndex="4" Width="100px">
                            <HeaderStyle Font-Names="Verdana" Font-Size="8pt" Wrap="True" HorizontalAlign="Center" />
                            <CellStyle Font-Names="Verdana" Font-Size="8pt">
                            </CellStyle>
                        </dx:GridViewDataTextColumn>
                        <dx:GridViewDataTextColumn Caption="PART NAME" FieldName="partname" Name="partname"
                            VisibleIndex="5" Width="150px">
                            <HeaderStyle Font-Names="Verdana" Font-Size="8pt" Wrap="True" HorizontalAlign="Center" />
                            <CellStyle Font-Names="Verdana" Font-Size="8pt">
                            </CellStyle>
                        </dx:GridViewDataTextColumn>
                        <dx:GridViewDataTextColumn Caption="UOM" FieldName="uom" Name="uom" VisibleIndex="6"
                            Width="60px">
                            <HeaderStyle Font-Names="Verdana" Font-Size="8pt" Wrap="True" HorizontalAlign="Center" />
                            <CellStyle Font-Names="Verdana" Font-Size="8pt">
                            </CellStyle>
                        </dx:GridViewDataTextColumn>
                        <dx:GridViewDataTextColumn Caption="QTY/BOX" FieldName="qtybox" Name="qtybox"
                            VisibleIndex="7" Width="70px">
                            <PropertiesTextEdit DisplayFormatString="{0:n0}">
                                
<MaskSettings ErrorText="Please input valid value !" Mask="<0..9999999999999g>"
                                    IncludeLiterals="DecimalSymbol" />
                            
</PropertiesTextEdit>
                            <HeaderStyle Font-Names="Verdana" Font-Size="8pt" Wrap="True" HorizontalAlign="Center" />
                            <CellStyle Font-Names="Verdana" Font-Size="8pt">
                            </CellStyle>
                        </dx:GridViewDataTextColumn>
                        <dx:GridViewDataTextColumn Caption="PASI DELIVERY QTY" FieldName="pasidelqty"
                            Name="pasidelqty" VisibleIndex="8" Width="100px">
                            <PropertiesTextEdit DisplayFormatString="{0:n0}">
                                
<MaskSettings ErrorText="Please input valid value !" Mask="<0..9999999999999g>"
                                    IncludeLiterals="DecimalSymbol" />
                            
</PropertiesTextEdit>
                            <HeaderStyle Font-Names="Verdana" Font-Size="8pt" Wrap="True" HorizontalAlign="Center" />
                            <CellStyle Font-Names="Verdana" Font-Size="8pt">
                            </CellStyle>
                        </dx:GridViewDataTextColumn>
                        <dx:GridViewDataTextColumn Caption="AFFILIATE RECEIVING QTY" FieldName="recqty"
                            Name="recqty" VisibleIndex="9" Width="100px">
                            <PropertiesTextEdit DisplayFormatString="{0:n0}">
                                
<MaskSettings ErrorText="Please input valid value !" Mask="<0..9999999999999g>"
                                    IncludeLiterals="DecimalSymbol" />
                            
</PropertiesTextEdit>
                            <HeaderStyle Font-Names="Verdana" Font-Size="8pt" Wrap="True" HorizontalAlign="Center" />
                            <CellStyle Font-Names="Verdana" Font-Size="8pt">
                            </CellStyle>
                        </dx:GridViewDataTextColumn>
                        <dx:GridViewDataTextColumn Caption="DELIVERY QTY (BOX)" FieldName="delqty"
                            Name="delqty" VisibleIndex="11" Width="100px">
                            <PropertiesTextEdit DisplayFormatString="{0:n0}">
                                
<MaskSettings ErrorText="Please input valid value !" Mask="<0..9999999999999g>"
                                    IncludeLiterals="DecimalSymbol" />
                            
</PropertiesTextEdit>
                            <HeaderStyle Font-Names="Verdana" Font-Size="8pt" Wrap="True" HorizontalAlign="Center" />
                            <CellStyle Font-Names="Verdana" Font-Size="8pt">
                            </CellStyle>
                        </dx:GridViewDataTextColumn>
                        <dx:GridViewDataTextColumn Caption="PO KANBAN" FieldName="pokanban" Name="colpokanban"
                            VisibleIndex="2" Width="70px">
                            <HeaderStyle Font-Names="Tahoma" Font-Size="8pt" HorizontalAlign="Center" VerticalAlign="Middle"
                                Wrap="True" />
                            <CellStyle Font-Names="Tahoma" Font-Size="8pt">
                            </CellStyle>
                        </dx:GridViewDataTextColumn>
                        <dx:GridViewDataTextColumn Caption="PASI INVOICE QTY" FieldName="invqty" Name="invqty"
                            VisibleIndex="10" Width="100px">
                            <PropertiesTextEdit DisplayFormatString="{0:n0}">
                                
                            <MaskSettings ErrorText="Please input valid value !" Mask="<0..9999999999999g>"
                                    IncludeLiterals="DecimalSymbol" />
                            
                            </PropertiesTextEdit>
                            <HeaderStyle Font-Names="Tahoma" Font-Size="8pt" HorizontalAlign="Center" VerticalAlign="Middle"
                                Wrap="True" />
                            <CellStyle Font-Names="Tahoma" Font-Size="8pt">
                            </CellStyle>
                        </dx:GridViewDataTextColumn>
                        <dx:GridViewBandColumn Caption="PASI INVOICE" VisibleIndex="12">
                            <Columns>
                                <dx:GridViewDataTextColumn Caption="CURR" FieldName="pasicurr" VisibleIndex="0" 
                                    Width="70px">
                                    <HeaderStyle Font-Names="Tahoma" Font-Size="8pt" HorizontalAlign="Center" />
                                    <CellStyle Font-Names="Tahoma" Font-Size="8pt">
                                    </CellStyle>
                                </dx:GridViewDataTextColumn>
                                <dx:GridViewDataTextColumn Caption="PRICE" FieldName="pasiprice" 
                                    VisibleIndex="1" Width="70px">
                                    <HeaderStyle Font-Names="Tahoma" Font-Size="8pt" HorizontalAlign="Center" />
                                    <PropertiesTextEdit DisplayFormatString="{0:n2}">
                                
                            <MaskSettings ErrorText="Please input valid value !" Mask="<0..9999999999999g>"
                                    IncludeLiterals="DecimalSymbol" />
                            
                            </PropertiesTextEdit>
                                    <CellStyle Font-Names="Tahoma" Font-Size="8pt">
                                    </CellStyle>
                                </dx:GridViewDataTextColumn>
                                <dx:GridViewDataTextColumn Caption="AMOUNT" FieldName="pasiamount" 
                                    VisibleIndex="2" Width="90px">
                                    <HeaderStyle Font-Names="Tahoma" Font-Size="8pt" />
                                    <PropertiesTextEdit DisplayFormatString="{0:n2}">
                                        
<MaskSettings ErrorText="Please input valid value !" Mask="<0..9999999999999g>"
                                            IncludeLiterals="DecimalSymbol" />
                                    
</PropertiesTextEdit>
                                    <CellStyle Font-Names="Tahoma" Font-Size="8pt">
                                    </CellStyle>
                                </dx:GridViewDataTextColumn>
                            </Columns>
                            <HeaderStyle Font-Names="Tahoma" Font-Size="8pt" HorizontalAlign="Center" />
                        </dx:GridViewBandColumn>
                    </Columns>
                    <SettingsPager Visible="False" PageSize="13" Position="Top">
                        <Summary Text="Page {0} of {1} [{2} record(s)]" AllPagesText="Page {0} of {1} " />
                        <Summary AllPagesText="Page {0} of {1} " Text="Page {0} of {1} [{2} record(s)]">
                        </Summary>
                    </SettingsPager>
                    <SettingsEditing Mode="Batch" NewItemRowPosition="Bottom">
                    </SettingsEditing>
                    <Settings ShowGroupButtons="False" ShowVerticalScrollBar="True" ShowHorizontalScrollBar="True"
                        VerticalScrollableHeight="135" ShowStatusBar="Hidden"></Settings>
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
                                Text="BACK">
                            </dx1:ASPxButton>
                        </td>
                        <td align="right" class="style27">
                            &nbsp;</td>
                        <td align="right">
                            <dx1:ASPxButton ID="btnsubmit" runat="server" Width="90px" Font-Names="Tahoma" Font-Size="8pt"
                                Text="DOWNLOAD" ClientInstanceName="btnsubmit" AutoPostBack="False">
                                <ClientSideEvents Click="function(s, e) {
	                                Grid.PerformCallback('save')
                                }" />
                            </dx1:ASPxButton>
                        </td>
                    </tr>
                </table>
            </td>
        </tr>
    </table>
</asp:Content>
