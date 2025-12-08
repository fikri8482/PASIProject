<%@ Page Title="" Language="vb" AutoEventWireup="false" MasterPageFile="~/SiteContent.Master"
    CodeBehind="ReceivingEntry.aspx.vb" Inherits="PASISystem.ReceivingEntry" %>

<%@ Register Assembly="DevExpress.Web.v14.1, Version=14.1.6.0, Culture=neutral, PublicKeyToken=b88d1754d700e49a"
    Namespace="DevExpress.Web.ASPxGridView.Export" TagPrefix="dx" %>
<%@ Register Assembly="DevExpress.Web.v14.1, Version=14.1.6.0, Culture=neutral, PublicKeyToken=b88d1754d700e49a"
    Namespace="DevExpress.Web.ASPxEditors" TagPrefix="dx1" %>
<%@ Register Assembly="DevExpress.Web.v14.1, Version=14.1.6.0, Culture=neutral, PublicKeyToken=b88d1754d700e49a"
    Namespace="DevExpress.Web.ASPxGridView" TagPrefix="dx" %>
    <%@ Register Assembly="DevExpress.Web.v14.1, Version=14.1.6.0, Culture=neutral, PublicKeyToken=b88d1754d700e49a" Namespace="DevExpress.Web.ASPxCallback" TagPrefix="dx" %>
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

            if (currentColumnName == "colno" || currentColumnName == "colpono" || currentColumnName == "colpokanban" || currentColumnName == "colkanbanno"
            || currentColumnName == "colpartno" || currentColumnName == "colpartname" || currentColumnName == "coluom" || currentColumnName == "colqtybox"
            || currentColumnName == "colremaining" || currentColumnName == "colsupplierqty" || currentColumnName == "colboxqty"){
                e.cancel = true;
            }

            currentEditableVisibleIndex = e.visibleIndex;
        }

		function OnBatchEditEndEditing(s, e) {
            window.setTimeout(function () {
                var reqqty = s.batchEditApi.GetCellValue(e.visibleIndex, "colreceivingqty");
                var supqty = s.batchEditApi.GetCellValue(e.visibleIndex, "colsupplierqty");
                var defqty = s.batchEditApi.GetCellValue(e.visibleIndex, "coldefect");
                var qtybox = s.batchEditApi.GetCellValue(e.visibleIndex, "colqtybox");

                if (reqqty > supqty) {
                    alert("Qty PASI GOOD Receiving Can't bigger than Supplier Delivery Qty !");
                    return;
                } else if (defqty > supqty) {
                    alert("Qty Defect Receiving Can't bigger than Supplier Delivery Qty !");
                    return;
                } else if (reqqty % qtybox > 0) {
                    alert("cannot input multiple MOQ!");
                    return;
                } else if (defqty % qtybox > 0) {
                    alert("cannot input multiple MOQ!");
                    return;
                }

            }, 10);
          
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
                <dx1:ASPxDateEdit ID="dt1" runat="server" ClientInstanceName="dt1" Font-Names="Tahoma"
                    Font-Size="8pt" EditFormat="Custom" EditFormatString="dd MMM yyyy" >
                    <ClientSideEvents Init="function(s, e) { s.SetDate(new Date()); }" />
                </dx1:ASPxDateEdit>
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
                            <dx1:ASPxTextBox ID="txtreceivedate" runat="server" Width="165px" Font-Names="Tahoma"
                                Font-Size="8pt" BackColor="#CCCCCC" ReadOnly="True" Height="16px" 
                                ClientInstanceName="txtreceivedate" Visible="False">
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
                                Text="SUPPLIER">
                            </dx1:ASPxLabel>
                        </td>
                        <td align="left">
                            <dx1:ASPxTextBox ID="txtsuppliercode" runat="server" Width="165px" Font-Names="Tahoma"
                                Font-Size="8pt" BackColor="#CCCCCC" ReadOnly="True" Height="16px" ClientInstanceName="txtsuppliercode">
                                <ClientSideEvents TextChanged="function(s, e) {
	                                lblerrmessage.SetText('');
                                }" />
                            </dx1:ASPxTextBox>
                        </td>
                        <td align="left">
                            <dx1:ASPxTextBox ID="txtsuppliername" runat="server" Width="165px" Font-Names="Tahoma"
                                Font-Size="8pt" BackColor="#CCCCCC" ReadOnly="True" Height="16px" ClientInstanceName="txtsuppliername">
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
                                            Text="PERFORMANCE CLS">
                                        </dx1:ASPxLabel>
                                    </td>
                                    <td>
                            <dx1:ASPxComboBox ID="cbocls" runat="server" ClientInstanceName="cbocls"
                                Font-Names="Tahoma" TextFormatString="{0}" Font-Size="8pt" Width="70px">
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
                                ClientInstanceName="txtcls">
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
                            <dx1:ASPxLabel ID="ASPxLabel4" runat="server" Font-Names="Tahoma" Font-Size="8pt"
                                Text="DRIVER NAME">
                            </dx1:ASPxLabel>
                        </td>
                        <td align="left">
                            <dx1:ASPxTextBox ID="txtdrivername" runat="server" Width="165px" Font-Names="Tahoma"
                                Font-Size="8pt" BackColor="White" Height="16px" ClientInstanceName="txtdrivername">
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
                                            Text="DRIVER CONTACT">
                                        </dx1:ASPxLabel>
                                    </td>
                                    <td>
                                        <dx1:ASPxTextBox ID="txtdrivercontact" runat="server" Width="100px" Font-Names="Tahoma"
                                            Font-Size="8pt" BackColor="White" Height="16px" MaxLength="20" ClientInstanceName="txtdrivercontact">
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
                                            BackColor="White" Height="16px" MaxLength="20" ClientInstanceName="txtnopol">
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
                                            Font-Size="8pt" BackColor="White" Height="16px" MaxLength="20" ClientInstanceName="txtjenisarmada">
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
                                        <dx1:ASPxLabel ID="ASPxLabel10" runat="server" Font-Names="Tahoma" Font-Size="8pt"
                                            Text="TOTAL BOX">
                                        </dx1:ASPxLabel>
                                    </td>
                                    <td>
                                        <dx1:ASPxTextBox ID="txttotalbox" runat="server" Width="100px" Font-Names="Tahoma"
                                            Font-Size="8pt" BackColor="#CCCCCC" ReadOnly="True" Height="16px" MaxLength="20"
                                            ClientInstanceName="txttotalbox">
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
                <dx:ASPxGridView ID="Grid" runat="server" AutoGenerateColumns="False" KeyFieldName="colno;colpono;colpartno;colpokanban;colkanbanno"
                    Width="100%" ClientInstanceName="Grid">
                    <ClientSideEvents BatchEditStartEditing="OnBatchEditStartEditing" BatchEditEndEditing="OnBatchEditEndEditing" CallbackError="function(s, e) {
e.handled = true;
}" EndCallback="function(s, e) {
									Grid.UpdateEdit();
                                    Grid.CancelEdit();
                                    dt1.SetText(s.cpDate);
                                    txtsuppliercode.SetText(s.cpScode); 
                                    txtsuppliername.SetText(s.cpSname);
                                    txtsuratjalanno.SetText(s.cpSJ);
                                    txtdrivername.SetText(s.cpDname);
                                    txtdrivercontact.SetText(s.cpDContact);
                                    txtplandeliverydate.SetText(s.cpPlandeldate);
                                    txtsupplierdeliverydate.SetText(s.cpSdeldate);
                                    txtnopol.SetText(s.cpNopol);
                                    txtjenisarmada.SetText(s.cpJenisarmada);
                                    txttotalbox.SetText(s.cpTotalbox);
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
                        <dx:GridViewDataTextColumn Caption="NO." FieldName="colno" Name="colno" VisibleIndex="0"
                            Width="30px">
                            <HeaderStyle Font-Names="Tahoma" Font-Size="8pt" Wrap="True" HorizontalAlign="Center" />
                            <CellStyle Font-Names="Tahoma" Font-Size="8pt">
                            </CellStyle>
                        </dx:GridViewDataTextColumn>
                        <dx:GridViewDataTextColumn Caption="PO NO." FieldName="colpono" Name="colpono" VisibleIndex="1"
                            Width="100px">
                            <HeaderStyle Font-Names="Verdana" Font-Size="8pt" Wrap="True" HorizontalAlign="Center" />
                            <CellStyle Font-Names="Verdana" Font-Size="8pt">
                            </CellStyle>
                        </dx:GridViewDataTextColumn>
                        <dx:GridViewDataTextColumn Caption="KANBAN NO." FieldName="colkanbanno" Name="colkanbanno"
                            VisibleIndex="3" Width="100px">
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
                        <dx:GridViewDataTextColumn Caption="UOM" FieldName="coluom" Name="coluom" VisibleIndex="6"
                            Width="60px">
                            <HeaderStyle Font-Names="Verdana" Font-Size="8pt" Wrap="True" HorizontalAlign="Center" />
                            <CellStyle Font-Names="Verdana" Font-Size="8pt">
                            </CellStyle>
                        </dx:GridViewDataTextColumn>
                        <dx:GridViewDataTextColumn Caption="QTY/BOX" FieldName="colqtybox" Name="colqtybox"
                            VisibleIndex="7" Width="70px">
                            <PropertiesTextEdit DisplayFormatString="{0:n0}">
                                
<MaskSettings ErrorText="Please input valid value !" Mask="<0..9999999999999g>"
                                    IncludeLiterals="DecimalSymbol" />
                            
</PropertiesTextEdit>
                            <HeaderStyle Font-Names="Verdana" Font-Size="8pt" Wrap="True" HorizontalAlign="Center" />
                            <CellStyle Font-Names="Verdana" Font-Size="8pt" HorizontalAlign="Right">
                            </CellStyle>
                        </dx:GridViewDataTextColumn>
                        <dx:GridViewDataTextColumn Caption="SUPPLIER DELIVERY QTY" FieldName="colsupplierqty"
                            Name="colsupplierqty" VisibleIndex="8" Width="100px">
                            <PropertiesTextEdit DisplayFormatString="{0:n0}">
                                
<MaskSettings ErrorText="Please input valid value !" Mask="<0..9999999999999g>"
                                    IncludeLiterals="DecimalSymbol" />
                            
</PropertiesTextEdit>
                            <HeaderStyle Font-Names="Verdana" Font-Size="8pt" Wrap="True" HorizontalAlign="Center" />
                            <CellStyle Font-Names="Verdana" Font-Size="8pt" HorizontalAlign="Right">
                            </CellStyle>
                        </dx:GridViewDataTextColumn>
                        <dx:GridViewDataTextColumn Caption="PASI GOOD RECEIVING QTY" FieldName="colreceivingqty"
                            Name="colreceivingqty" VisibleIndex="9" Width="100px">
                            <PropertiesTextEdit DisplayFormatString="{0:n0}">
                                <MaskSettings ErrorText="Please input valid value !" Mask="<0..9999999999999g>"
                                    IncludeLiterals="DecimalSymbol" />
                                <ValidationSettings ErrorDisplayMode="None"></ValidationSettings>
                            </PropertiesTextEdit>
                            <HeaderStyle Font-Names="Verdana" Font-Size="8pt" Wrap="True" HorizontalAlign="Center" />
                            <CellStyle Font-Names="Verdana" Font-Size="8pt" HorizontalAlign="Right">
                            </CellStyle>
                        </dx:GridViewDataTextColumn>
                        <dx:GridViewDataTextColumn Caption="PASI REMAINING RECEIVING QTY" FieldName="colremaining"
                            Name="colremaining" VisibleIndex="11" Width="100px">
                            <PropertiesTextEdit DisplayFormatString="{0:n0}">
                                
<MaskSettings ErrorText="Please input valid value !" Mask="<0..9999999999999g>"
                                    IncludeLiterals="DecimalSymbol" />
                            
</PropertiesTextEdit>
                            <HeaderStyle Font-Names="Verdana" Font-Size="8pt" Wrap="True" HorizontalAlign="Center" />
                            <CellStyle Font-Names="Verdana" Font-Size="8pt" HorizontalAlign="Right">
                            </CellStyle>
                        </dx:GridViewDataTextColumn>
                        <dx:GridViewDataTextColumn Caption="PASI RECEIVING QTY (BOX)" FieldName="colboxqty" Name="colboxqty"
                            VisibleIndex="12" Width="100px">
                            <PropertiesTextEdit DisplayFormatString="{0:n0}">
                                
<MaskSettings ErrorText="Please input valid value !" Mask="<0..9999999999999g>"
                                    IncludeLiterals="DecimalSymbol" />
                            
</PropertiesTextEdit>
                            <HeaderStyle Font-Names="Verdana" Font-Size="8pt" Wrap="True" HorizontalAlign="Center" />
                            <CellStyle Font-Names="Verdana" Font-Size="8pt" HorizontalAlign="Right">
                            </CellStyle>
                        </dx:GridViewDataTextColumn>
                        <dx:GridViewDataTextColumn Caption="PO KANBAN" FieldName="colpokanban" Name="colpokanban"
                            VisibleIndex="2" Width="70px">
                            <HeaderStyle Font-Names="Tahoma" Font-Size="8pt" HorizontalAlign="Center" VerticalAlign="Middle"
                                Wrap="True" />
                            <CellStyle Font-Names="Tahoma" Font-Size="8pt">
                            </CellStyle>
                        </dx:GridViewDataTextColumn>
                        <dx:GridViewDataTextColumn Caption="DEFECT RECEIVING QTY" FieldName="coldefect" Name="coldefect"
                            VisibleIndex="10" Width="100px">
                            <PropertiesTextEdit DisplayFormatString="{0:n0}">
                                <MaskSettings ErrorText="Please input valid value !" Mask="<0..9999999999999g>"
                                    IncludeLiterals="DecimalSymbol" />
                                <ValidationSettings ErrorDisplayMode="None"></ValidationSettings>
                            </PropertiesTextEdit>
                            <HeaderStyle Font-Names="Tahoma" Font-Size="8pt" HorizontalAlign="Center" VerticalAlign="Middle"
                                Wrap="True" />
                            <CellStyle Font-Names="Tahoma" Font-Size="8pt" HorizontalAlign="Right">
                            </CellStyle>
                        </dx:GridViewDataTextColumn>
                        <dx:GridViewDataTextColumn Caption="colunitcls" FieldName="colunitcls" 
                            VisibleIndex="13" Width="0px">
                        </dx:GridViewDataTextColumn>
                        <dx:GridViewDataTextColumn FieldName="colHgood" Name="colHgood" 
                            VisibleIndex="14" Width="0px">
                        </dx:GridViewDataTextColumn>
                        <dx:GridViewDataTextColumn FieldName="colHdefect" Name="colHdefect" 
                            VisibleIndex="15" Width="0px">
                        </dx:GridViewDataTextColumn>
                        <dx:GridViewDataTextColumn FieldName="colKanbanQty" VisibleIndex="16" 
                            Width="0px">
                        </dx:GridViewDataTextColumn>
						<dx:GridViewDataTextColumn FieldName="colPrice" VisibleIndex="17" 
                            Width="0px">
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
                            &nbsp;</td>
                        <td align="right">
                            <table style="width:100%;">
                                <tr>
                                <td valign="top" align="right" style="width: 50px;">
                <dx1:ASPxButton ID="btnPrintGR" runat="server" Text="PRINT GOOD RECEIVING"
                    Font-Names="Tahoma"
                    Width="90px" AutoPostBack="False" Font-Size="8pt" Enabled="true" 
                    Visible="true">
                </dx1:ASPxButton>
            </td>
                                <td valign="top" align="right" style="width: 50px;">
                <dx1:ASPxButton ID="btnSendGR" runat="server" Text="SEND GOOD RECEIVING TO SUPPLIER"                              
                    Font-Names="Tahoma" Width="80px" AutoPostBack="False" Font-Size="8pt">
                    <ClientSideEvents Click="function(s, e) {
                        ButtonApprove.PerformCallback();
                    }" />
                </dx1:ASPxButton>
            </td>  
                                    <td valign="top" align="right" >
                            <dx1:ASPxButton ID="btndelete" runat="server" Width="90px" Font-Names="Tahoma" Font-Size="8pt"
                                Text="DELETE" ClientInstanceName="btndelete">
                                <ClientSideEvents Click="function(s, e) {
	Grid.PerformCallback('gridload');
}" />
                            </dx1:ASPxButton>
                                    </td>
                                    <td valign="top" align="right" >
                            <dx1:ASPxButton ID="btnsubmit" runat="server" Width="90px" Font-Names="Tahoma" Font-Size="8pt"
                                Text="SAVE" ClientInstanceName="btnsubmit" AutoPostBack="False">
                                <ClientSideEvents Click="function(s, e) {
	Grid.UpdateEdit();
	Grid.PerformCallback('save')
	Grid.PerformCallback('gridload')
	Grid.CancelEdit();
}" />
<ClientSideEvents Click="function(s, e) {
	Grid.UpdateEdit();
	lblerrmessage.SetText(''); 
	Grid.PerformCallback('save')
	Grid.PerformCallback('gridload')
	Grid.CancelEdit();
}"></ClientSideEvents>
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
