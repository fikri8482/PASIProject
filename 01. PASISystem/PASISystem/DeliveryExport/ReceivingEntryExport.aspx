<%@ Page Title="" Language="vb" AutoEventWireup="false" MasterPageFile="~/SiteContent.Master"
    CodeBehind="ReceivingEntryExport.aspx.vb" Inherits="PASISystem.ReceivingEntryExport" %>

<%@ Register Assembly="DevExpress.Web.v14.1, Version=14.1.6.0, Culture=neutral, PublicKeyToken=b88d1754d700e49a"
    Namespace="DevExpress.Web.ASPxGridView.Export" TagPrefix="dx" %>
<%@ Register Assembly="DevExpress.Web.v14.1, Version=14.1.6.0, Culture=neutral, PublicKeyToken=b88d1754d700e49a"
    Namespace="DevExpress.Web.ASPxEditors" TagPrefix="dx1" %>
<%@ Register Assembly="DevExpress.Web.v14.1, Version=14.1.6.0, Culture=neutral, PublicKeyToken=b88d1754d700e49a"
    Namespace="DevExpress.Web.ASPxGridView" TagPrefix="dx" %>
    <%@ Register Assembly="DevExpress.Web.v14.1, Version=14.1.6.0, Culture=neutral, PublicKeyToken=b88d1754d700e49a" Namespace="DevExpress.Web.ASPxCallback" TagPrefix="dx" %>
<%@ Register assembly="DevExpress.Web.v14.1, Version=14.1.6.0, Culture=neutral, PublicKeyToken=b88d1754d700e49a" namespace="DevExpress.Web.ASPxHiddenField" tagprefix="dx" %>
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
//            Grid.PerformCallback("Update");
        }

        function OnCancelClick(s, e) {
//            Grid.PerformCallback("Cancel");
        }

        function OnBatchEditStartEditing(s, e) {
            currentColumnName = e.focusedColumn.fieldName;

            if (currentColumnName == "colno" || currentColumnName == "colorderno" || currentColumnName == "collabelno" || currentColumnName == "colpartno"
            || currentColumnName == "colpartname" || currentColumnName == "coluom" || currentColumnName == "colqtybox" || currentColumnName == "coldelqty"
            || currentColumnName == "colremaining" ) {
                e.cancel = true;
            }

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
                                Text="RECEIVED DATE">
                            </dx1:ASPxLabel>
                        </td>
                        <td align="left">
                            <dx1:ASPxTextBox ID="txtrecdate" runat="server" Width="165px" Font-Names="Tahoma"
                                Font-Size="8pt" BackColor="#CCCCCC" ReadOnly="True" Height="16px" 
                                ClientInstanceName="txtrecdate">
                                <ClientSideEvents TextChanged="function(s, e) {
	                                lblerrmessage.SetText('');
                                }" />
                            </dx1:ASPxTextBox>
                        </td>
                        <td align="left">
                            &nbsp;
                            <dx1:ASPxTextBox ID="txtkanbanno" runat="server" Width="165px" Font-Names="Tahoma"
                                Font-Size="8pt" BackColor="#CCCCCC" ReadOnly="True" Height="16px" 
                                ClientInstanceName="txtkanbanno" Visible="False">
                                <ClientSideEvents TextChanged="function(s, e) {
	                                lblerrmessage.SetText('');
                                }" />
                            </dx1:ASPxTextBox>
                        </td>
                        <td align="left">
                            &nbsp;
                            <dx1:ASPxTextBox ID="txtaffiliate" runat="server" Width="165px" Font-Names="Tahoma"
                                Font-Size="8pt" BackColor="#CCCCCC" ReadOnly="True" Height="16px" 
                                ClientInstanceName="txtaffiliate" Visible="False">
                                <ClientSideEvents TextChanged="function(s, e) {
	                                lblerrmessage.SetText('');
                                }" />
                            </dx1:ASPxTextBox>
                        </td>
                        <td align="left">
                            &nbsp;
                            <dx1:ASPxTextBox ID="txtstatus" runat="server" Width="165px" Font-Names="Tahoma"
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
                        </td>
                    </tr>
                    <tr>
                        <td align="left">
                            <dx1:ASPxLabel ID="ASPxLabel2" runat="server" Font-Names="Tahoma" Font-Size="8pt"
                                Text="SUPPLIER">
                            </dx1:ASPxLabel>
                        </td>
                        <td align="left">
                            <dx1:ASPxTextBox ID="txtsupp" runat="server" Width="165px" Font-Names="Tahoma"
                                Font-Size="8pt" BackColor="#CCCCCC" ReadOnly="True" Height="16px" 
                                ClientInstanceName="txtsupp">
                                <ClientSideEvents TextChanged="function(s, e) {
	                                lblerrmessage.SetText('');
                                }" />
                            </dx1:ASPxTextBox>
                        </td>
                        <td align="left">
                            <dx1:ASPxTextBox ID="txtsuppliername" runat="server" Width="210px" Font-Names="Tahoma"
                                Font-Size="8pt" BackColor="#CCCCCC" ReadOnly="True" Height="16px" 
                                ClientInstanceName="txtsuppliername">
                                <ClientSideEvents TextChanged="function(s, e) {
	                                lblerrmessage.SetText('');
                                }" />
                            </dx1:ASPxTextBox>
                        </td>
                        <td align="left">
                            <dx1:ASPxLabel ID="ASPxLabel5" runat="server" Font-Names="Tahoma" Font-Size="8pt"
                                Text="SUPPLIER PLAN DELIVERY DATE">
                            </dx1:ASPxLabel>
                        </td>
                        <td align="left">
                            <dx1:ASPxTextBox ID="txtplandeliverydate" runat="server" Width="165px" Font-Names="Tahoma"
                                Font-Size="8pt" BackColor="#CCCCCC" ReadOnly="True" Height="16px" ClientInstanceName="txtplandeliverydate">
                                <ClientSideEvents TextChanged="function(s, e) {
	                                lblerrmessage.SetText('');
                                }" />
                            </dx1:ASPxTextBox>
                        </td>
                        <td align="left">
                            <table style="width:100%;">
                                <tr>
                                    <td>
                            <dx1:ASPxLabel ID="ASPxLabel7" runat="server" Font-Names="Tahoma" Font-Size="8pt"
                                Text="SUPPLIER DELIVERY DATE">
                            </dx1:ASPxLabel>
                                    </td>
                                    <td>
                            <dx1:ASPxTextBox ID="txtsupplierdeliverydate" runat="server" Width="165px" Font-Names="Tahoma"
                                Font-Size="8pt" BackColor="#CCCCCC" ReadOnly="True" Height="16px" ClientInstanceName="txtsupplierdeliverydate">
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
                    </tr>
                    <tr>
                        <td align="left">
                            <dx1:ASPxLabel ID="ASPxLabel3" runat="server" Font-Names="Tahoma" Font-Size="8pt"
                                Text="SURAT JALAN NO.">
                            </dx1:ASPxLabel>
                        </td>
                        <td align="left">
                            <dx1:ASPxTextBox ID="txtsuratjalanno" runat="server" Width="165px" Font-Names="Tahoma"
                                Font-Size="8pt" BackColor="#CCCCCC" ReadOnly="True" Height="16px" ClientInstanceName="txtsuratjalanno">
                                <ClientSideEvents TextChanged="function(s, e) {
	                                lblerrmessage.SetText('');
                                }" />
                            </dx1:ASPxTextBox>
                        </td>
                        <td align="left" colspan="2">
                            &nbsp;
                            <table style="width:100%;">
                                <tr>
                                    <td>
                                        <dx1:ASPxLabel ID="ASPxLabel12" runat="server" Font-Names="Tahoma" Font-Size="8pt"
                                            Text="PERFORMANCE CLS" Visible="False">
                                        </dx1:ASPxLabel>
                                    </td>
                                    <td>
                            <dx1:ASPxComboBox ID="cbocls" runat="server" ClientInstanceName="cbocls"
                                Font-Names="Tahoma" TextFormatString="{0}" Font-Size="8pt" Width="70px" 
                                            Visible="False">
                                <ClientSideEvents SelectedIndexChanged="function(s, e) {
	                                            txtcls.SetText(cbocls.GetSelectedItem().GetColumnText(1));
                                                }" TextChanged="function(s, e) {
	txtcls.SetText(cbocls.GetSelectedItem().GetColumnText(1));
}" ValueChanged="function(s, e) {
	txtcls.SetText(cbocls.GetSelectedItem().GetColumnText(1));
}" />
                            </dx1:ASPxComboBox>
                                    </td>
                                    <td>
                            <dx1:ASPxTextBox ID="txtcls" runat="server" Width="150px" Font-Names="Tahoma"
                                Font-Size="8pt" BackColor="#CCCCCC" ReadOnly="True" Height="16px" 
                                ClientInstanceName="txtcls" Visible="False">
                            </dx1:ASPxTextBox>
                                    </td>
                                </tr>
                            </table>
                            &nbsp;
                            </td>
                        <td align="left">
                            &nbsp;
                            </td>
                        <td align="left">
                            &nbsp;
                            <dx1:ASPxTextBox ID="txtpono" runat="server" Width="165px" Font-Names="Tahoma"
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
                                        <dx1:ASPxLabel ID="ASPxLabel10" runat="server" Font-Names="Tahoma" Font-Size="8pt"
                                            Text="TOTAL BOX">
                                        </dx1:ASPxLabel>
                        </td>
                        <td align="left">
                                        <dx1:ASPxTextBox ID="txttotalbox" runat="server" Width="100px" Font-Names="Tahoma"
                                            Font-Size="8pt" BackColor="White" Height="16px" MaxLength="20"
                                            ClientInstanceName="txttotalbox" ReadOnly="True">
                                            <ClientSideEvents TextChanged="function(s, e) {
	                                lblerrmessage.SetText('');
                                }" />
                                        </dx1:ASPxTextBox>
                        </td>
                        <td align="left">
                            <table style="width: 100%;">
                                <tr>
                                    <td>
                                        <dx1:ASPxLabel ID="ASPxLabel6" runat="server" Font-Names="Tahoma" Font-Size="8pt"
                                            Text="DRIVER CONTACT" Visible="False">
                                        </dx1:ASPxLabel>
                                    </td>
                                    <td>
                                        <dx1:ASPxTextBox ID="txtdrivercontact" runat="server" Width="100px" Font-Names="Tahoma"
                                            Font-Size="8pt" BackColor="White" Height="16px" MaxLength="20" 
                                            ClientInstanceName="txtdrivercontact" Visible="False">
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
                                            Text="NO. POL" Visible="False">
                                        </dx1:ASPxLabel>
                                    </td>
                                    <td>
                                        <dx1:ASPxTextBox ID="txtnopol" runat="server" Width="100px" Font-Names="Tahoma" Font-Size="8pt"
                                            BackColor="White" Height="16px" MaxLength="20" 
                                            ClientInstanceName="txtnopol" Visible="False">
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
                                            Text="JENIS ARMADA" Visible="False">
                                        </dx1:ASPxLabel>
                                    </td>
                                    <td>
                                        <dx1:ASPxTextBox ID="txtjenisarmada" runat="server" Width="100px" Font-Names="Tahoma"
                                            Font-Size="8pt" BackColor="White" Height="16px" MaxLength="20" 
                                            ClientInstanceName="txtjenisarmada" Visible="False">
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
                                        &nbsp;</td>
                                    <td>
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
                            <dx1:ASPxLabel ID="ASPxLabel13" runat="server" Font-Names="Tahoma" Font-Size="8pt"
                                Text="AFFILIATE">
                            </dx1:ASPxLabel>
                        </td>
                        <td align="left">
                            <dx1:ASPxTextBox ID="txtaffiliatecode" runat="server" Width="165px" Font-Names="Tahoma"
                                Font-Size="8pt" BackColor="#CCCCCC" ReadOnly="True" Height="16px" 
                                ClientInstanceName="txtaffiliatecode">
                                <ClientSideEvents TextChanged="function(s, e) {
	                                lblerrmessage.SetText('');
                                }" />
                            </dx1:ASPxTextBox>
                        </td>
                        <td align="left">
                            <dx1:ASPxTextBox ID="txtsuppliercode1" runat="server" Width="165px" Font-Names="Tahoma"
                                Font-Size="8pt" BackColor="#CCCCCC" ReadOnly="True" Height="16px" 
                                ClientInstanceName="txtsuppliercode" Visible="False">
                                <ClientSideEvents TextChanged="function(s, e) {
	                                lblerrmessage.SetText('');
                                }" />
                            </dx1:ASPxTextBox>
                        </td>
                        <td align="left">
                            <dx1:ASPxLabel ID="ASPxLabel4" runat="server" Font-Names="Tahoma" Font-Size="8pt"
                                Text="DRIVER NAME" Visible="False">
                            </dx1:ASPxLabel>
                        </td>
                        <td align="left">
                            <dx1:ASPxTextBox ID="txtdrivername" runat="server" Width="165px" Font-Names="Tahoma"
                                Font-Size="8pt" BackColor="White" Height="16px" 
                                ClientInstanceName="txtdrivername" Visible="False">
                                <ClientSideEvents TextChanged="function(s, e) {
	                                lblerrmessage.SetText('');
                                }" />
                            </dx1:ASPxTextBox>
                        </td>
                        <td align="left">
                            &nbsp;</td>
                        <td align="left">
                            &nbsp;</td>
                    </tr>
                    <tr>
                        <td align="left">
                            <dx1:ASPxLabel ID="ASPxLabel14" runat="server" Font-Names="Tahoma" Font-Size="8pt"
                                Text="DELIVERY LOCATION">
                            </dx1:ASPxLabel>
                        </td>
                        <td align="left">
                            <dx1:ASPxTextBox ID="txtdeliverycode" runat="server" Width="165px" Font-Names="Tahoma"
                                Font-Size="8pt" BackColor="#CCCCCC" ReadOnly="True" Height="16px" 
                                ClientInstanceName="txtdeliverycode">
                                <ClientSideEvents TextChanged="function(s, e) {
	                                lblerrmessage.SetText('');
                                }" />
                            </dx1:ASPxTextBox>
                        </td>
                        <td align="left">
                            <dx1:ASPxTextBox ID="txtdeliveryname" runat="server" Width="210px" Font-Names="Tahoma"
                                Font-Size="8pt" BackColor="#CCCCCC" ReadOnly="True" Height="16px" 
                                ClientInstanceName="txtdeliveryname">
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
                <dx:ASPxGridView ID="Grid" runat="server" AutoGenerateColumns="False" KeyFieldName="idx;colno;colorderno;colpartno;colpono;LabelNo1;LabelNo2"
                    Width="100%" ClientInstanceName="Grid">
                    <ClientSideEvents BatchEditStartEditing="OnBatchEditStartEditing" 
                        CallbackError="function(s, e) {e.handled = true; }" 
                        
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
}" />

                    <Settings ShowHorizontalScrollBar="True" ShowVerticalScrollBar="True" ShowGroupButtons="False"
                        ShowStatusBar="Hidden" VerticalScrollableHeight="135" />

                    <Columns>
                        <dx:GridViewDataCheckColumn Caption=" " FieldName="colpilih" VisibleIndex="0" 
                            Width="30px">
                            <propertiescheckedit valuechecked="1" valuetype="System.Int32" 
                                valueunchecked="0"></propertiescheckedit>
                        </dx:GridViewDataCheckColumn>
                        <dx:GridViewDataTextColumn Caption="NO." FieldName="colno" Name="colno" VisibleIndex="1"
                            Width="30px">
                            <HeaderStyle Font-Names="Tahoma" Font-Size="8pt" Wrap="True" HorizontalAlign="Center" />
                            <CellStyle Font-Names="Tahoma" Font-Size="8pt">
                            </CellStyle>
                        </dx:GridViewDataTextColumn>
                        <dx:GridViewDataTextColumn Caption="ORDER NO." FieldName="colorderno" 
                            Name="colorderno" VisibleIndex="2"
                            Width="100px">
                            <HeaderStyle Font-Names="Verdana" Font-Size="8pt" Wrap="True" HorizontalAlign="Center" />
                            <CellStyle Font-Names="Verdana" Font-Size="8pt">
                            </CellStyle>
                        </dx:GridViewDataTextColumn>
                        <dx:GridViewDataTextColumn Caption="PART NO." FieldName="colpartno" Name="colpartno"
                            VisibleIndex="4" Width="100px">
                            <HeaderStyle Font-Names="Verdana" Font-Size="8pt" Wrap="True" HorizontalAlign="Center" />
                            <CellStyle Font-Names="Verdana" Font-Size="8pt">
                            </CellStyle>
                        </dx:GridViewDataTextColumn>
                        <dx:GridViewDataTextColumn Caption="PART NAME" FieldName="colpartname" Name="colpartname"
                            VisibleIndex="5" Width="150px">
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
                        <dx:GridViewDataTextColumn Caption="QTY/BOX" FieldName="colqtybox" Name="colqtybox"
                            VisibleIndex="9" Width="70px">
                            <PropertiesTextEdit DisplayFormatString="{0:n0}">
                                

<MaskSettings ErrorText="Please input valid value !" Mask="<0..9999999999999g>"
                                    IncludeLiterals="DecimalSymbol" />
                            

</PropertiesTextEdit>
                            <HeaderStyle Font-Names="Verdana" Font-Size="8pt" Wrap="True" HorizontalAlign="Center" />
                            <CellStyle Font-Names="Verdana" Font-Size="8pt" HorizontalAlign="Right">
                            </CellStyle>
                        </dx:GridViewDataTextColumn>
                        <dx:GridViewDataTextColumn Caption="SUPPLIER DELIVERY QTY" FieldName="coldelqty"
                            Name="coldelqty" VisibleIndex="10" Width="100px">
                            <PropertiesTextEdit DisplayFormatString="{0:n0}">
                                

<MaskSettings ErrorText="Please input valid value !" Mask="<0..9999999999999g>"
                                    IncludeLiterals="DecimalSymbol" />
                            

</PropertiesTextEdit>
                            <HeaderStyle Font-Names="Verdana" Font-Size="8pt" Wrap="True" HorizontalAlign="Center" />
                            <CellStyle Font-Names="Verdana" Font-Size="8pt" HorizontalAlign="Right">
                            </CellStyle>
                        </dx:GridViewDataTextColumn>
                        <dx:GridViewDataTextColumn Caption="GOOD RECEIVING QTY" FieldName="colgoodreceiving"
                            Name="colgoodreceiving" VisibleIndex="11" Width="100px">
                            <PropertiesTextEdit DisplayFormatString="{0:n0}">
                                
<MaskSettings ErrorText="Please input valid value !" Mask="<0..9999999999999g>"
                                    IncludeLiterals="DecimalSymbol" />
                                

<ValidationSettings ErrorDisplayMode="None"></ValidationSettings>
                            
</PropertiesTextEdit>
                            <HeaderStyle Font-Names="Verdana" Font-Size="8pt" Wrap="True" HorizontalAlign="Center" />
                            <CellStyle Font-Names="Verdana" Font-Size="8pt" HorizontalAlign="Right">
                            </CellStyle>
                        </dx:GridViewDataTextColumn>
                        <dx:GridViewDataTextColumn Caption="DEFECT RECEIVING QTY (BOX)" FieldName="coldefect"
                            Name="coldefect" VisibleIndex="14" Width="100px">
                            <PropertiesTextEdit DisplayFormatString="{0:n0}">
                                

<MaskSettings ErrorText="Please input valid value !" Mask="<0..9999999999999g>"
                                    IncludeLiterals="DecimalSymbol" />
                            

</PropertiesTextEdit>
                            <HeaderStyle Font-Names="Verdana" Font-Size="8pt" Wrap="True" HorizontalAlign="Center" />
                            <CellStyle Font-Names="Verdana" Font-Size="8pt" HorizontalAlign="Right">
                            </CellStyle>
                        </dx:GridViewDataTextColumn>
                        <dx:GridViewDataTextColumn Caption="GOOD RECEIVING QTY (BOX)" 
                            FieldName="colreceivingbox" Name="colreceivingbox"
                            VisibleIndex="13" Width="100px">
                            <PropertiesTextEdit DisplayFormatString="{0:n0}">
                                

<MaskSettings ErrorText="Please input valid value !" Mask="<0..9999999999999g>"
                                    IncludeLiterals="DecimalSymbol" />
                            

</PropertiesTextEdit>
                            <HeaderStyle Font-Names="Verdana" Font-Size="8pt" Wrap="True" HorizontalAlign="Center" />
                            <CellStyle Font-Names="Verdana" Font-Size="8pt" HorizontalAlign="Right">
                            </CellStyle>
                        </dx:GridViewDataTextColumn>
                        <dx:GridViewDataTextColumn Caption="DEFECT RECEIVING QTY" 
                            FieldName="coldefectreceiving" Name="coldefectreceiving"
                            VisibleIndex="12" Width="100px">
                            <PropertiesTextEdit DisplayFormatString="{0:n0}">
                                
<MaskSettings ErrorText="Please input valid value !" Mask="<0..9999999999999g>"
                                    IncludeLiterals="DecimalSymbol" />
                                

<ValidationSettings ErrorDisplayMode="None"></ValidationSettings>
                            
</PropertiesTextEdit>
                        </dx:GridViewDataTextColumn>
                        <dx:GridViewDataTextColumn FieldName="colHgood" Name="colHgood" 
                            VisibleIndex="15" Width="0px">
                        </dx:GridViewDataTextColumn>
                        <dx:GridViewDataTextColumn FieldName="colHdefect" Name="colHdefect" 
                            VisibleIndex="16" Width="0px">
                        </dx:GridViewDataTextColumn>
                        <dx:GridViewDataTextColumn FieldName="colpono" Name="colpono" VisibleIndex="17" 
                            Width="0px">
                        </dx:GridViewDataTextColumn>
                        <dx:GridViewDataTextColumn Caption="BOX NO FROM" FieldName="LabelNo1" 
                            VisibleIndex="6">
                            <HeaderStyle HorizontalAlign="Center" />
                            <CellStyle HorizontalAlign="Center">
                            </CellStyle>
                        </dx:GridViewDataTextColumn>
                        <dx:GridViewDataTextColumn Caption="BOX NO TO" FieldName="LabelNo2"
                            VisibleIndex="7">
                            <HeaderStyle Font-Names="Tahoma" Font-Size="8pt" HorizontalAlign="Center" VerticalAlign="Middle"
                                Wrap="True" />
                            <CellStyle Font-Names="Tahoma" Font-Size="8pt" HorizontalAlign="Right">
                            </CellStyle>
                            <HeaderStyle HorizontalAlign="Center" />
                            <CellStyle HorizontalAlign="Center">
                            </CellStyle>
                        </dx:GridViewDataTextColumn>
                        <dx:GridViewDataTextColumn FieldName="PO" VisibleIndex="18" Width="0px">
                        </dx:GridViewDataTextColumn>
                        <dx:GridViewDataTextColumn FieldName="PART" VisibleIndex="19" Width="0px">
                        </dx:GridViewDataTextColumn>
                        <dx:GridViewDataTextColumn Caption=" LABEL" FieldName="LABEL" VisibleIndex="20" 
                            Width="0px" Name="LABEL">
                        </dx:GridViewDataTextColumn>
                        <dx:GridViewDataTextColumn Caption="StatusDefect" FieldName="StatusDefect" 
                            VisibleIndex="21" Width="0px">
                        </dx:GridViewDataTextColumn>
                    </Columns>
                    <SettingsPager Visible="False" PageSize="13" Position="Top" 
                        mode="ShowAllRecords">
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
                                Text="SUB MENU">
                            </dx1:ASPxButton>
                        </td>
                        <td align="right" class="style27">
               <dx1:ASPxTextBox ID="lblStatus" runat="server" Width="0px" Font-Names="Verdana"
                                Font-Size="8pt" ForeColor="White" Height="16px" 
                    ClientInstanceName="lblStatus" BackColor="White" ReadOnly="True">
                   <Border BorderColor="White" />
               </dx1:ASPxTextBox>
                        </td>
                        <td align="right">
                            <table style="width:100%;">
                                <tr>
                                <td valign="top" align="right" style="width: 50px;">
                                        <dx1:ASPxButton ID="btnaddcarton" runat="server" AutoPostBack="False" 
                                            ClientInstanceName="btnaddcarton" Text="ADD CARTON" Visible="False">
<ClientSideEvents Click="function(s, e) {

                                        HF.Set('hfTest', 'add');

                                            Grid.UpdateEdit();

											Grid.PerformCallback('addrow');
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

                                        }"></ClientSideEvents>
                                        </dx1:ASPxButton>
            </td>
                                <td valign="top" align="right" style="width: 50px;">
                <dx1:ASPxButton ID="btnPrintGR" runat="server" Text="PRINT GOOD RECEIVING"
                    Font-Names="Tahoma"
                    Width="90px" AutoPostBack="False" Font-Size="8pt" Visible="False">
                </dx1:ASPxButton>
            </td>
                                <td valign="top" align="right" style="width: 50px;">
                <dx1:ASPxButton ID="btnSendGR" runat="server" Text="SEND GOOD RECEIVING TO SUPPLIER"                              
                    Font-Names="Tahoma" Width="80px" AutoPostBack="False" Font-Size="8pt" Visible="False">
                    <ClientSideEvents Click="function(s, e) {
                        ButtonApprove.PerformCallback();
                    }" />
                </dx1:ASPxButton>
            </td>  
                                    <td valign="top" align="right" >
                            <dx1:ASPxButton ID="btndelete" runat="server" Width="90px" Font-Names="Tahoma" Font-Size="8pt"
                                Text="DELETE" ClientInstanceName="btndelete" AutoPostBack="False">
                                <ClientSideEvents Click="function(s, e) {
    HF.Set('hfTest', 'delete');
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
}" />
                            </dx1:ASPxButton>
                                    </td>
                                    <td valign="top" align="right" >
                                        <dx1:ASPxButton ID="btnsubmit" runat="server" AutoPostBack="False" 
                                            ClientInstanceName="btnsubmit" Text="SUBMIT" Visible="False">
                                            <ClientSideEvents Click="function(s, e) {
    HF.Set('hfTest', 'save');
	Grid.UpdateEdit();
	lblerrmessage.SetText(''); 
	Grid.PerformCallback('save')
	Grid.CancelEdit();
}" />
                                        </dx1:ASPxButton>
                                    </td>
                                </tr>
                            </table>
                        </td>
                    </tr>
                </table>
            </td>
        </tr>
    </table>
    <dx:ASPxHiddenField ID="HF" runat="server" ClientInstanceName="HF">
    </dx:ASPxHiddenField>
    <dx:ASPxCallback ID="ButtonApprove" runat="server" ClientInstanceName="ButtonApprove">
        <ClientSideEvents EndCallback="function(s, e) {
	        var pMsg = s.cpMessage;
            if (pMsg != '') {
                if (pMsg.substring(1,5) == '1010') {
                    lblerrmessage.GetMainElement().style.color = 'Blue';
                } else {
                    lblerrmessage.GetMainElement().style.color = 'Red';
                }
                lblerrmessage.SetText(pMsg);
            } else {
                lblerrmessage.SetText('');
            }
            delete s.cpMessage;
        }" />
    </dx:ASPxCallback>
</asp:Content>
