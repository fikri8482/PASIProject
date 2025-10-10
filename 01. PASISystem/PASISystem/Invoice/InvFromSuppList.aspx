<%@ Page Title="" Language="vb" AutoEventWireup="false" MasterPageFile="~/SiteContent.Master" CodeBehind="InvFromSuppList.aspx.vb" Inherits="PASISystem.InvFromSuppList" %>
<%@ Register assembly="DevExpress.Web.v14.1, Version=14.1.6.0, Culture=neutral, PublicKeyToken=b88d1754d700e49a" namespace="DevExpress.Web.ASPxEditors" tagprefix="dx1" %>
<%@ Register assembly="DevExpress.Web.v14.1, Version=14.1.6.0, Culture=neutral, PublicKeyToken=b88d1754d700e49a" namespace="DevExpress.Web.ASPxGridView" tagprefix="dx" %>
<%@ Register assembly="DevExpress.Web.v14.1, Version=14.1.6.0, Culture=neutral, PublicKeyToken=b88d1754d700e49a" namespace="DevExpress.Web.ASPxGridView.Export" tagprefix="dx2" %>
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
        
        .style1
        {
            height: 331px;
        }
        
        </style>
</asp:Content>
<asp:Content ID="Content2" ContentPlaceHolderID="Content" runat="server">
    <table style="width:100%;" width="100%">
        <tr>
            <td align="left">
                            <dx1:ASPxLabel ID="ASPxLabel1" runat="server" Font-Names="Tahoma" Font-Size="8pt"
                                Text="SUPPLIER DELIVERY DATE">
                            </dx1:ASPxLabel>
                        </td>
            <td align="left">
                <dx1:ASPxCheckBox ID="checkbox1" runat="server" CheckState="Unchecked" 
                    ClientInstanceName="checkbox1" Text=" ">
                </dx1:ASPxCheckBox>
            </td>
            <td align="left">
                <table style="width:100%;">
                    <tr>
                        <td>
                <dx1:ASPxDateEdit ID="dt1" runat="server" ClientInstanceName="dt1" 
                    Font-Names="Tahoma" Font-Size="8pt" EditFormat="Custom" 
                    EditFormatString="dd MMM yyyy" Width="100px">
                </dx1:ASPxDateEdit>
                        </td>
                        <td>
                            <dx1:ASPxLabel ID="ASPxLabel12" runat="server" Font-Names="Tahoma" Font-Size="8pt"
                                Text="~">
                            </dx1:ASPxLabel>
                        </td>
                        <td style="font-weight: 700">
                <dx1:ASPxDateEdit ID="dt2" runat="server" ClientInstanceName="dt2" 
                    Font-Names="Tahoma" Font-Size="8pt" EditFormat="Custom" 
                    EditFormatString="dd MMM yyyy" Width="100px">
                </dx1:ASPxDateEdit>
                        </td>
                    </tr>
                </table>
            </td>
            <td align="left">
                            <dx1:ASPxLabel ID="ASPxLabel7" runat="server" Font-Names="Tahoma" Font-Size="8pt"
                                Text="SUPPLIER INVOICE NO">
                            </dx1:ASPxLabel>
                        </td>
            <td align="left">
                <dx1:ASPxTextBox ID="txtsupinvno" runat="server" Font-Names="Tahoma" 
                    Font-Size="8pt" Width="170px" ClientInstanceName="txtsupinvno">
                </dx1:ASPxTextBox>
            </td>
            <td align="left">
                &nbsp;</td>
        </tr>
        <tr>
            <td align="left">
                            <dx1:ASPxLabel ID="ASPxLabel2" runat="server" Font-Names="Tahoma" Font-Size="8pt"
                                Text="SUPPLIER ALREADY INVOICE ">
                            </dx1:ASPxLabel>
                        </td>
            <td align="left">
                &nbsp;</td>
            <td align="left">
                <dx1:ASPxRadioButtonList ID="rbinvoice" runat="server" 
                    ClientInstanceName="rbinvoice" Font-Names="Tahoma" Font-Size="8pt" 
                    Height="16px" RepeatDirection="Horizontal">
                    <Items>
                        <dx1:ListEditItem Text="ALL" Value="ALL" Selected="True"/>
                        <dx1:ListEditItem Text="YES" Value="YES" />
                        <dx1:ListEditItem Text="NO" Value="NO" />
                    </Items>
                    <Border BorderStyle="None" />
                </dx1:ASPxRadioButtonList>
            </td>
            <td align="left">
                            <dx1:ASPxLabel ID="ASPxLabel14" runat="server" Font-Names="Tahoma" Font-Size="8pt"
                                Text="SUPPLIER CODE / NAME">
                            </dx1:ASPxLabel>
                        </td>
            <td align="left">
                <dx1:ASPxComboBox ID="cbosupplier" runat="server" ClientInstanceName="cbosupplier"
                                Font-Names="Tahoma" TextFormatString="{0}" Font-Size="8pt">
                                <ClientSideEvents SelectedIndexChanged="function(s, e) {
	                                            txtsupplier.SetText(cbosupplier.GetSelectedItem().GetColumnText(1));
                                                }" />
                            </dx1:ASPxComboBox>
            </td>
            <td align="left">
                <dx1:ASPxTextBox ID="txtsupplier" runat="server" Font-Names="Tahoma" 
                    Font-Size="8pt" Width="170px" BackColor="Silver" 
                    ClientInstanceName="txtsupplier" ReadOnly="True">
                    <ClientSideEvents TextChanged="function(s, e) {
	lblerrmessage.SetText('');
}" />
                </dx1:ASPxTextBox>
            </td>
        </tr>
        <tr>
            <td align="left">
                            <dx1:ASPxLabel ID="ASPxLabel4" runat="server" Font-Names="Tahoma" Font-Size="8pt"
                                Text="SUPPLIER DIFFERENCE QTY">
                            </dx1:ASPxLabel>
                        </td>
            <td align="left">
                &nbsp;</td>
            <td align="left">
                <dx1:ASPxRadioButtonList ID="rbdeliveryqty" runat="server" 
                    ClientInstanceName="rbdeliveryqty" Font-Names="Tahoma" Font-Size="8pt" 
                    Height="16px" RepeatDirection="Horizontal">
                    <Items>
                        <dx1:ListEditItem Text="ALL" Value="ALL" Selected="True"/>
                        <dx1:ListEditItem Text="YES" Value="YES" />
                        <dx1:ListEditItem Text="NO" Value="NO" />
                    </Items>
                    <Border BorderStyle="None" />
                </dx1:ASPxRadioButtonList>
            </td>
            <td align="left">
                            <dx1:ASPxLabel ID="ASPxLabel13" runat="server" Font-Names="Tahoma" Font-Size="8pt"
                                Text="AFFILIATE CODE / NAME">
                            </dx1:ASPxLabel>
                        </td>
            <td align="left">
                <dx1:ASPxComboBox ID="cboaffiliate" runat="server" ClientInstanceName="cboaffiliate"
                                Font-Names="Tahoma" TextFormatString="{0}" Font-Size="8pt">
                                <ClientSideEvents SelectedIndexChanged="function(s, e) {
	                                            txtaffiliate.SetText(cboaffiliate.GetSelectedItem().GetColumnText(1));
                                                }" />
                            </dx1:ASPxComboBox>
            </td>
            <td align="left">
                <dx1:ASPxTextBox ID="txtaffiliate" runat="server" Font-Names="Tahoma" 
                    Font-Size="8pt" Width="170px" BackColor="Silver" 
                    ClientInstanceName="txtaffiliate" ReadOnly="True">
                    <ClientSideEvents TextChanged="function(s, e) {
	lblerrmessage.SetText('');
}" />
                </dx1:ASPxTextBox>
            </td>
        </tr>
        <tr>
            <td align="left">
                            <dx1:ASPxLabel ID="ASPxLabel5" runat="server" Font-Names="Tahoma" Font-Size="8pt"
                                Text="SUPPLIER SURAT JALAN NO">
                            </dx1:ASPxLabel>
                        </td>
            <td align="left">
                &nbsp;</td>
            <td align="left">
                <dx1:ASPxTextBox ID="txtsj" runat="server" ClientInstanceName="txtsj" 
                    Font-Size="8pt" Width="170px">
                </dx1:ASPxTextBox>
            </td>
            <td align="left">
                            <dx1:ASPxLabel ID="ASPxLabel9" runat="server" Font-Names="Tahoma" Font-Size="8pt"
                                Text="PO NO.">
                            </dx1:ASPxLabel>
                        </td>
            <td align="left">
                <dx1:ASPxTextBox ID="txtpono" runat="server" Font-Names="Tahoma" 
                    Font-Size="8pt" Width="170px" ClientInstanceName="txtpono">
                </dx1:ASPxTextBox>
            </td>
            <td align="left">
                <table style="width:100%;">
                    <tr>
                        <td>
                            <dx1:ASPxButton ID="btnsearch" runat="server" Text="SEARCH" Width="85px" Font-Names="Tahoma"
                                Font-Size="8pt" ClientInstanceName="btnsearch" AutoPostBack="False">
                                <ClientSideEvents Click="function(s, e) {                                        
                                                            
	                                        grid.PerformCallback('gridload');
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
                        <td>
                            <dx1:ASPxButton ID="btnclear" runat="server" Text="CLEAR" Width="85px" Font-Names="Tahoma"
                                Font-Size="8pt">
                            </dx1:ASPxButton>
                        </td>
                    </tr>
                </table>
            </td>
        </tr>
    </table>
    <br />
            <table align="left" width="100%">
                <tr align="left">
                    <td width="100%" height="16px" style="border-top-style: solid; border-top-width: thin;
                        border-top-color: #808080; border-bottom-style: solid; border-bottom-width: thin;
                        border-bottom-color: #808080" align="left">
                        <dx1:ASPxLabel ID="lblerrmessage" runat="server" Font-Names="Tahoma" Font-Size="8pt"
                            Text="ERROR MESSAGE" Width="100%" ClientInstanceName="lblerrmessage">
                        </dx1:ASPxLabel>
                    </td>
                </tr>
            </table>
        <br />
    <br />
    <table style="width:100%;" width="100%">
        <tr>
            <td align="left" class="style1">
                <dx:ASPxGridView ID="grid" runat="server" AutoGenerateColumns="False" 
                    Width="100%" KeyFieldName="sortPONo;KanbanNo;PartNo;SupplierCode;H_SupSj" ClientInstanceName="grid">

                    <%--<SettingsEditing Mode="Batch" NewItemRowPosition="Bottom">
                        <BatchEditSettings ShowConfirmOnLosingChanges="False" />
                    </SettingsEditing>--%>

                    <ClientSideEvents EndCallback="function(s, e) {
						var pMsg = s.cpMessage;
                        if (pMsg != '') {
                            if (pMsg.substring(1,5) == '1001' || pMsg.substring(1,5) == '1002' || pMsg.substring(1,5) == '1003' || pMsg.substring(1,5) == '2001') {
                                lblerrmessage.GetMainElement().style.color = 'Blue';
                            } else {
                                lblerrmessage.GetMainElement().style.color = 'Red';
                            }
                            lblerrmessage.SetText(pMsg);
                        } else {
                            lblerrmessage.SetText('');
                        }
                        delete s.cpMessage;
                        }" RowClick="function(s, e) {
	                        lblerrmessage.SetText('');
                        }" />
                    
                    <SettingsBehavior AllowSelectByRowClick="True" AllowSort="False"
                        ColumnResizeMode="Control" EnableRowHotTrack="True" />

                    <Settings ShowHorizontalScrollBar="True" ShowVerticalScrollBar="True"
                        ShowGroupButtons="False" ShowStatusBar="Hidden"
                        VerticalScrollableHeight="190" />

<ClientSideEvents RowClick="function(s, e) {
	                        lblerrmessage.SetText(&#39;&#39;);
                        }" EndCallback="function(s, e) {
						var pMsg = s.cpMessage;
                        if (pMsg != &#39;&#39;) {
                            if (pMsg.substring(1,5) == &#39;1001&#39; || pMsg.substring(1,5) == &#39;1002&#39; || pMsg.substring(1,5) == &#39;1003&#39; || pMsg.substring(1,5) == &#39;2001&#39;) {
                                lblerrmessage.GetMainElement().style.color = &#39;Blue&#39;;
                            } else {
                                lblerrmessage.GetMainElement().style.color = &#39;Red&#39;;
                            }
                            lblerrmessage.SetText(pMsg);
                        } else {
                            lblerrmessage.SetText(&#39;&#39;);
                        }
                        delete s.cpMessage;
                        }"></ClientSideEvents>

                    <Columns>
                        <dx:GridViewDataHyperLinkColumn Caption=" " FieldName="url" VisibleIndex="0" 
                            Width="50px">
                            <PropertiesHyperLinkEdit TextField="coldetail">
                            </PropertiesHyperLinkEdit>
                            <HeaderStyle Font-Names="Tahoma" Font-Size="8pt" Font-Underline="False" 
                                Wrap="True" />
                            <CellStyle Font-Names="Tahoma" Font-Size="8pt" VerticalAlign="Middle">
                            </CellStyle>
                        </dx:GridViewDataHyperLinkColumn>
                        <dx:GridViewDataTextColumn Caption="NO" FieldName="no" 
                            VisibleIndex="1" Width="0px">
                            <HeaderStyle Font-Names="Tahoma" Font-Size="8pt" Font-Underline="False" 
                                Wrap="True" HorizontalAlign="Center" />
                            <CellStyle Font-Names="Tahoma" Font-Size="8pt">
                            </CellStyle>
                        </dx:GridViewDataTextColumn>
                        <dx:GridViewDataTextColumn Caption="PERIOD" FieldName="period" VisibleIndex="2" 
                            Width="60px" Name="period">
                            <HeaderStyle Font-Names="Tahoma" Font-Size="8pt" Font-Underline="False" 
                                Wrap="True" HorizontalAlign="Center"/>
                            <CellStyle Font-Names="Tahoma" Font-Size="8pt">
                            </CellStyle>
                        </dx:GridViewDataTextColumn>
                        <dx:GridViewDataTextColumn Caption="AFFILIATE CODE" 
                            FieldName="affiliatecode" VisibleIndex="3" 
                            Width="70px">
                            <HeaderStyle Font-Names="Tahoma" Font-Size="8pt" Font-Underline="False" 
                                Wrap="True" HorizontalAlign="Center"/>
                            <CellStyle Font-Names="Tahoma" Font-Size="8pt">
                            </CellStyle>
                        </dx:GridViewDataTextColumn>
                        <dx:GridViewDataTextColumn Caption="AFFILIATE NAME" 
                            FieldName="affiliatename" VisibleIndex="4" 
                            Width="120px">
                            <HeaderStyle Font-Names="Tahoma" Font-Size="8pt" Font-Underline="False" 
                                Wrap="True" HorizontalAlign="Center"/>
                            <CellStyle Font-Names="Tahoma" Font-Size="8pt">
                            </CellStyle>
                        </dx:GridViewDataTextColumn>
                        <dx:GridViewDataTextColumn Caption="PO NO." FieldName="pono" 
                            VisibleIndex="5" Width="80px">
                            <HeaderStyle Font-Names="Tahoma" Font-Size="8pt" Font-Underline="False" 
                                Wrap="True" HorizontalAlign="Center"/>
                            <CellStyle Font-Names="Tahoma" Font-Size="8pt">
                            </CellStyle>
                        </dx:GridViewDataTextColumn>
                        <dx:GridViewDataTextColumn Caption="SUPPLIER CODE" FieldName="suppliercode" 
                            VisibleIndex="6" Width="70px">
                            <HeaderStyle Font-Names="Tahoma" Font-Size="8pt" Font-Underline="False" 
                                Wrap="True" HorizontalAlign="Center"/>
                            <CellStyle Font-Names="Tahoma" Font-Size="8pt">
                            </CellStyle>
                        </dx:GridViewDataTextColumn>
                        <dx:GridViewDataTextColumn Caption="SUPPLIER NAME" FieldName="suppliername" 
                            VisibleIndex="7" Width="100px">
                            <HeaderStyle Font-Names="Tahoma" Font-Size="8pt" Font-Underline="False" 
                                Wrap="True" HorizontalAlign="Center"/>
                            <CellStyle Font-Names="Tahoma" Font-Size="8pt" HorizontalAlign="Center">
                            </CellStyle>
                        </dx:GridViewDataTextColumn>
                        <dx:GridViewDataTextColumn Caption="PO KANBAN" FieldName="pokanban" 
                            VisibleIndex="8" Width="50px">
                            <HeaderStyle Font-Names="Tahoma" Font-Size="8pt" Font-Underline="False" 
                                Wrap="True" HorizontalAlign="Center" />
                            <CellStyle Font-Names="Tahoma" Font-Size="8pt">
                            </CellStyle>
                            
                        </dx:GridViewDataTextColumn>
                        <dx:GridViewDataTextColumn Caption="KANBAN NO." FieldName="kanbanno" 
                            VisibleIndex="9" Width="80px">
                            <HeaderStyle Font-Names="Tahoma" Font-Size="8pt" Font-Underline="False" 
                                Wrap="True" HorizontalAlign="Center"/>
                            <CellStyle Font-Names="Tahoma" Font-Size="8pt" HorizontalAlign="Center">
                            </CellStyle>
                        </dx:GridViewDataTextColumn>
                        <dx:GridViewDataTextColumn Caption="SUPPLIER PLAN DELIVERY DATE" 
                            FieldName="suppplandeldate" VisibleIndex="10" 
                            Width="80px">
                            <HeaderStyle Font-Names="Tahoma" Font-Size="8pt" Font-Underline="False" 
                                Wrap="True" HorizontalAlign="Center"/>
                            <CellStyle Font-Names="Tahoma" Font-Size="8pt" HorizontalAlign="Center">
                            </CellStyle>
                        </dx:GridViewDataTextColumn>
                        <dx:GridViewDataTextColumn Caption="SUPPLIER DELIVERY DATE" 
                            FieldName="suppdeldate" VisibleIndex="11" 
                            Width="80px">
                            <HeaderStyle Font-Names="Tahoma" Font-Size="8pt" Font-Underline="False" 
                                Wrap="True" HorizontalAlign="Center"/>
                            <CellStyle Font-Names="Tahoma" Font-Size="8pt">
                            </CellStyle>
                        </dx:GridViewDataTextColumn>
                        <dx:GridViewDataTextColumn Caption="SUPPLIER SURAT JALAN NO." 
                            FieldName="suppsj" VisibleIndex="12" Width="100px">
                            <HeaderStyle Font-Names="Tahoma" Font-Size="8pt" Font-Underline="False" 
                                Wrap="True" HorizontalAlign="Center"/>
                            <CellStyle Font-Names="Tahoma" Font-Size="8pt">
                            </CellStyle>
                        </dx:GridViewDataTextColumn>
                        <dx:GridViewDataTextColumn Caption="PART NO." FieldName="partno" 
                            VisibleIndex="15" Width="90px" Name="partno">
                            <HeaderStyle Font-Names="Tahoma" Font-Size="8pt" Font-Underline="False" 
                                Wrap="True" HorizontalAlign="Center"/>
                            <CellStyle Font-Names="Tahoma" Font-Size="8pt">
                            </CellStyle>
                        </dx:GridViewDataTextColumn>
                        <dx:GridViewDataTextColumn Caption="PART NAME" 
                            FieldName="partname" VisibleIndex="16" Width="140px">
                            <HeaderStyle Font-Names="Tahoma" Font-Size="8pt" Font-Underline="False" 
                                Wrap="True" HorizontalAlign="Center"/>
                            <CellStyle Font-Names="Tahoma" Font-Size="8pt">
                            </CellStyle>
                        </dx:GridViewDataTextColumn>
                        <dx:GridViewDataTextColumn Caption="UOM" FieldName="uom" 
                            VisibleIndex="17" Width="40px">
                            <HeaderStyle Font-Names="Tahoma" Font-Size="8pt" Font-Underline="False" 
                                Wrap="True" HorizontalAlign="Center"/>
                            <CellStyle Font-Names="Tahoma" Font-Size="8pt" HorizontalAlign="Right">
                            </CellStyle>
                        </dx:GridViewDataTextColumn>
                        <dx:GridViewDataTextColumn Caption="SUPPLIER DELIVERY QTY" 
                            FieldName="suppdelqty" VisibleIndex="18" Width="70px">
                            <HeaderStyle Font-Names="Tahoma" Font-Size="8pt" Font-Underline="False" 
                                Wrap="True" HorizontalAlign="Center"/>
                            <PropertiesTextEdit DisplayFormatString="{0:n0}">
                                
                            <MaskSettings ErrorText="Please input valid value !" Mask="<0..9999999999999g>"
                                                                IncludeLiterals="DecimalSymbol" />
                            
                            </PropertiesTextEdit>
                            <CellStyle Font-Names="Tahoma" Font-Size="8pt" HorizontalAlign = "Right">
                            </CellStyle>
                        </dx:GridViewDataTextColumn>
                        <dx:GridViewDataTextColumn Caption="PASI RECEIVING QTY" FieldName="pasirecqty" 
                            VisibleIndex="19" Width="70px">
                            <HeaderStyle Font-Names="Tahoma" Font-Size="8pt" Font-Underline="False" 
                                Wrap="True" HorizontalAlign="Center"/>
                            <PropertiesTextEdit DisplayFormatString="{0:n0}">
                                
                            <MaskSettings ErrorText="Please input valid value !" Mask="<0..9999999999999g>"
                                                                IncludeLiterals="DecimalSymbol" />
                            
                            </PropertiesTextEdit>
                            <CellStyle Font-Names="Tahoma" Font-Size="8pt" HorizontalAlign="Right">
                            </CellStyle>
                        </dx:GridViewDataTextColumn>
                        <dx:GridViewDataTextColumn Caption="SUPPLIER INVOICE QTY" 
                            FieldName="suppinvqty" VisibleIndex="21" 
                            Width="0px">
                            <HeaderStyle Font-Names="Tahoma" Font-Size="8pt" Font-Underline="False" 
                                Wrap="True" HorizontalAlign="Center"/>
                            <PropertiesTextEdit DisplayFormatString="{0:n0}">
                                
                            <MaskSettings ErrorText="Please input valid value !" Mask="<0..9999999999999g>"
                                                                IncludeLiterals="DecimalSymbol" />
                            
                            </PropertiesTextEdit>
                            <CellStyle Font-Names="Tahoma" Font-Size="8pt" HorizontalAlign="Right">
                            </CellStyle>
                        </dx:GridViewDataTextColumn>
                        <dx:GridViewDataTextColumn Caption="SUPPLIER INVOICE NO" 
                            FieldName="suppinvno" VisibleIndex="22" Width="80px">
                            <HeaderStyle Font-Names="Tahoma" Font-Size="8pt" Font-Underline="False" 
                                Wrap="True" HorizontalAlign="Center"/>
                            <CellStyle Font-Names="Tahoma" Font-Size="8pt">
                            </CellStyle>
                        </dx:GridViewDataTextColumn>
                        <dx:GridViewDataTextColumn Caption="SUPPLIER INVOICE DATE" 
                            FieldName="suppinvdate" VisibleIndex="23" Width="80px">
                        </dx:GridViewDataTextColumn>
                        <dx:GridViewDataTextColumn 
                            FieldName="remrecqty" 
                            VisibleIndex="20" Width="70px" Caption="REMAINING RECEIVING QTY">
                            <CellStyle HorizontalAlign="Right">
                            </CellStyle>
                        </dx:GridViewDataTextColumn>
                        
                        <dx:GridViewDataTextColumn Caption="PASI DELIVERY DATE" FieldName="pasideldate" 
                            VisibleIndex="13" Width="0px">
                            <HeaderStyle Font-Names="Tahoma" Font-Size="8pt" HorizontalAlign="Center" />
                            <CellStyle Font-Names="Tahoma" Font-Size="8pt">
                            </CellStyle>
                        </dx:GridViewDataTextColumn>
                        
                        <dx:GridViewDataTextColumn Caption="PASI SURAT JALAN NO" 
                            FieldName="pasisj" VisibleIndex="14" Width="0px">
                            <HeaderStyle Font-Names="Tahoma" Font-Size="8pt" HorizontalAlign="Center" />
                            <CellStyle Font-Names="Tahoma" Font-Size="8pt">
                            </CellStyle>
                        </dx:GridViewDataTextColumn>
                        <dx:GridViewBandColumn Caption="PASI RECEIVING" VisibleIndex="24" 
                            Name="pasirec">
                            <Columns>
                                <dx:GridViewDataTextColumn Caption="CURR" FieldName="pasireccurr" 
                                    VisibleIndex="1">
                                    <HeaderStyle Font-Names="Tahoma" Font-Size="8pt" HorizontalAlign="Center" />
                                    <CellStyle HorizontalAlign="Right">
                                    </CellStyle>
                                </dx:GridViewDataTextColumn>
                                <dx:GridViewDataTextColumn Caption="AMOUNT" FieldName="pasirecamount" 
                                    VisibleIndex="2">
                                    <HeaderStyle Font-Names="Tahoma" Font-Size="8pt" HorizontalAlign="Center" />
                                    <PropertiesTextEdit DisplayFormatString="{0:n2}">
                                
                                    <MaskSettings ErrorText="Please input valid value !" Mask="<0..9999999999999g>"
                                                                        IncludeLiterals="DecimalSymbol" />
                            
                                    </PropertiesTextEdit>

                                    <CellStyle HorizontalAlign="Right">
                                    </CellStyle>

                                </dx:GridViewDataTextColumn>
                            </Columns>
                            <HeaderStyle Font-Names="Tahoma" Font-Size="8pt" HorizontalAlign="Center" />
                        </dx:GridViewBandColumn>
                        <dx:GridViewBandColumn Caption="SUPPLIER INVOICE" VisibleIndex="25" 
                            Name="suppinv" Visible="False">
                            <Columns>
                                <dx:GridViewDataTextColumn Caption="CURR" FieldName="suppinvcurr" 
                                    VisibleIndex="1">
                                    <HeaderStyle Font-Names="Tahoma" Font-Size="8pt" HorizontalAlign="Center" />
                                    <CellStyle HorizontalAlign="Right">
                                    </CellStyle>
                                </dx:GridViewDataTextColumn>
                                <dx:GridViewDataTextColumn Caption="AMOUNT" FieldName="suppinvamount" 
                                    VisibleIndex="2">
                                    <HeaderStyle Font-Names="Tahoma" Font-Size="8pt" HorizontalAlign="Center" />
                                    <PropertiesTextEdit DisplayFormatString="{0:n2}">
                                
                                    <MaskSettings ErrorText="Please input valid value !" Mask="<0..9999999999999g>"
                                                                        IncludeLiterals="DecimalSymbol" />
                            
                                    </PropertiesTextEdit>
                                    <CellStyle HorizontalAlign="Right">
                                    </CellStyle>
                                </dx:GridViewDataTextColumn>
                            </Columns>
                            <HeaderStyle Font-Names="Tahoma" Font-Size="8pt" Font-Underline="False" 
                                HorizontalAlign="Center" Wrap="True" />
                        </dx:GridViewBandColumn>
                        
                        <dx:GridViewDataTextColumn FieldName="sortpono" VisibleIndex="26" Width="0px">
                        </dx:GridViewDataTextColumn>
                        
                        <dx:GridViewDataTextColumn FieldName="H_SupSj" 
                            VisibleIndex="27" Width="0px">
                        </dx:GridViewDataTextColumn>
                        
                    </Columns>
                    
<SettingsBehavior AllowSort="False" AllowSelectByRowClick="True" ColumnResizeMode="Control" 
                        EnableRowHotTrack="True"></SettingsBehavior>

                    <SettingsPager PageSize="16" Position="Top">
                        <Summary Text="Page {0} of {1} [{2} record(s)]"
                                  AllPagesText="Page {0} of {1} " />
                        <Summary AllPagesText="Page {0} of {1} " Text="Page {0} of {1} [{2} record(s)]"></Summary>
                    </SettingsPager>

<Settings ShowGroupButtons="False" ShowVerticalScrollBar="True" ShowHorizontalScrollBar="True" 
                        VerticalScrollableHeight="190" ShowStatusBar="Hidden"></Settings>

                    <Styles>
                        <SelectedRow ForeColor="Black">
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
                <dx1:ASPxButton ID="btnsubmenu" runat="server" Text="SUB MENU" Width="90px" Font-Names="Tahoma"
                    Font-Size="8pt">
                </dx1:ASPxButton>
            </td>
        </tr>
    </table>
</asp:Content>
