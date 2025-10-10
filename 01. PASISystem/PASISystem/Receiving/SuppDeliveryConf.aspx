<%@ Page Title="" Language="vb" AutoEventWireup="false" MasterPageFile="~/SiteContent.Master" CodeBehind="SuppDeliveryConf.aspx.vb" Inherits="PASISystem.SuppDeliveryConf" %>
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
                                Text="SUPPLIER PLAN DELIVERY DATE (UNTIL)">
                            </dx1:ASPxLabel>
                        </td>
            <td align="left">
                <dx1:ASPxCheckBox ID="checkbox1" runat="server" CheckState="Unchecked" 
                    ClientInstanceName="checkbox1" Text=" ">
                </dx1:ASPxCheckBox>
            </td>
            <td align="left">
                <dx1:ASPxDateEdit ID="dt1" runat="server" ClientInstanceName="dt1" 
                    Font-Names="Tahoma" Font-Size="8pt" EditFormat="Custom" 
                    EditFormatString="dd MMM yyyy" Width="100px">
                </dx1:ASPxDateEdit>
            </td>
            <td align="left">
                            <dx1:ASPxLabel ID="ASPxLabel7" runat="server" Font-Names="Tahoma" Font-Size="8pt"
                                Text="SUPPLIER CODE/NAME">
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
                            <dx1:ASPxLabel ID="ASPxLabel2" runat="server" Font-Names="Tahoma" Font-Size="8pt"
                                Text="SUPPLIER ALREADY DELIVER">
                            </dx1:ASPxLabel>
                        </td>
            <td align="left">
                &nbsp;</td>
            <td align="left">
                <dx1:ASPxRadioButtonList ID="rbdeliver" runat="server" 
                    ClientInstanceName="rbdeliver" Font-Names="Tahoma" Font-Size="8pt" 
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
                            <dx1:ASPxLabel ID="ASPxLabel8" runat="server" Font-Names="Tahoma" Font-Size="8pt"
                                Text="PART CODE/NAME">
                            </dx1:ASPxLabel>
                        </td>
            <td align="left">
                <dx1:ASPxComboBox ID="cbopart" runat="server" ClientInstanceName="cbopart"
                                Font-Names="Tahoma" TextFormatString="{0}" Font-Size="8pt">
                                <ClientSideEvents SelectedIndexChanged="function(s, e) {
	                                            txtpart.SetText(cbopart.GetSelectedItem().GetColumnText(1));
                                                }" />
                            </dx1:ASPxComboBox>
            </td>
            <td align="left">
                <dx1:ASPxTextBox ID="txtpart" runat="server" Font-Names="Tahoma" 
                    Font-Size="8pt" Width="170px" BackColor="Silver" 
                    ClientInstanceName="txtpart" ReadOnly="True">
                    <ClientSideEvents TextChanged="function(s, e) {
	lblerrmessage.SetText('');
}" />
                </dx1:ASPxTextBox>
            </td>
        </tr>
        <tr>
            <td align="left">
                            <dx1:ASPxLabel ID="ASPxLabel3" runat="server" Font-Names="Tahoma" Font-Size="8pt"
                                Text="REMAINING RECEIVING QTY">
                            </dx1:ASPxLabel>
                        </td>
            <td align="left">
                &nbsp;</td>
            <td align="left">
                <dx1:ASPxRadioButtonList ID="rbreceiving" runat="server" 
                    ClientInstanceName="rbreceiving" Font-Names="Tahoma" Font-Size="8pt" 
                    RepeatDirection="Horizontal">
                    <Items>
                        <dx1:ListEditItem Text="ALL" Value="ALL" Selected="True"/>
                        <dx1:ListEditItem Text="YES" Value="YES" />
                        <dx1:ListEditItem Text="NO" Value="NO" />
                    </Items>
                    <Border BorderStyle="None" />
                </dx1:ASPxRadioButtonList>
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
                &nbsp;</td>
        </tr>
        <tr>
            <td align="left">
                            <dx1:ASPxLabel ID="ASPxLabel4" runat="server" Font-Names="Tahoma" Font-Size="8pt"
                                Text="SUPPLIER SURAT JALAN NO.">
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
                            <dx1:ASPxLabel ID="ASPxLabel10" runat="server" Font-Names="Tahoma" Font-Size="8pt"
                                Text="PO KANBAN">
                            </dx1:ASPxLabel>
                        </td>
            <td align="left">
                <dx1:ASPxRadioButtonList ID="rbkanban" runat="server" 
                    Font-Names="Tahoma" Font-Size="8pt" RepeatDirection="Horizontal" 
                    ClientInstanceName="rbkanban">
                    <Items>
                        <dx1:ListEditItem Text="ALL" Value="ALL" Selected="True" />
                        <dx1:ListEditItem Text="YES" Value="YES" />
                        <dx1:ListEditItem Text="NO" Value="NO" />
                    </Items>
                    <Border BorderStyle="None" />
                </dx1:ASPxRadioButtonList>
            </td>
            <td align="left">
                &nbsp;</td>
        </tr>
        <tr>
            <td align="left">
                            <dx1:ASPxLabel ID="ASPxLabel5" runat="server" Font-Names="Tahoma" Font-Size="8pt"
                                Text="SUPPLIER DELIVERY DATE">
                            </dx1:ASPxLabel>
                        </td>
            <td align="left">
                <dx1:ASPxCheckBox ID="checkbox2" runat="server" CheckState="Unchecked" 
                    ClientInstanceName="checkbox2" Text=" ">
                </dx1:ASPxCheckBox>
            </td>
            <td align="left">
                <table style="width:100%;">
                    <tr>
                        <td>
                <dx1:ASPxDateEdit ID="dtfrom" runat="server" ClientInstanceName="dtfrom" 
                    Font-Size="8pt" EditFormat="Custom" EditFormatString="dd MMM yyyy" Width="100px">
                </dx1:ASPxDateEdit>
                        </td>
                        <td>
                            <dx1:ASPxLabel ID="ASPxLabel6" runat="server" Font-Names="Tahoma" Font-Size="8pt"
                                Text="~">
                            </dx1:ASPxLabel>
                        </td>
                        <td>
                <dx1:ASPxDateEdit ID="dtto" runat="server" ClientInstanceName="dtto" 
                    Font-Size="8pt" EditFormat="Custom" EditFormatString="dd MMM yyyy" Width="100px">
                </dx1:ASPxDateEdit>
                        </td>
                    </tr>
                </table>
            </td>
            <td align="left">
                            <dx1:ASPxLabel ID="ASPxLabel11" runat="server" Font-Names="Tahoma" Font-Size="8pt"
                                Text="GOOD RECEIVE SENT">
                            </dx1:ASPxLabel>
                        </td>
            <td align="left">
                <dx1:ASPxRadioButtonList ID="rbgoodreceive" runat="server" 
                    Font-Names="Tahoma" Font-Size="8pt" RepeatDirection="Horizontal" 
                    ClientInstanceName="rbgoodreceive">
                    <Items>
                        <dx1:ListEditItem Text="ALL" Value="ALL" Selected="True" />
                        <dx1:ListEditItem Text="YES" Value="YES" />
                        <dx1:ListEditItem Text="NO" Value="NO" />
                    </Items>
                    <Border BorderStyle="None" />
                </dx1:ASPxRadioButtonList>
            </td>
            <td align="left">
                <table style="width:100%;">
                    <tr>
                        <td>
                            <dx1:ASPxButton ID="btnsearch" runat="server" Text="SEARCH" Width="85px" Font-Names="Tahoma"
                                Font-Size="8pt" ClientInstanceName="btnsearch" AutoPostBack="False">
                                <ClientSideEvents Click="function(s, e) {
                                            var pDateFrom = dtfrom.GetText();
	                                        var pDateTo = dtto.GetText();
                                            var pSupplier = cbosupplier.GetText();
                                            var pPlandate = dt1.GetText();
                                            var pSupplierDeliver = rbdeliver.GetValue();
                                            var pKanban = rbkanban.GetValue();
                                            var pPono = txtpono.GetText();
                                            var psj = txtsj.GetText();
                                            var pRemaining = rbreceiving.GetValue();
                                            var pPartcode = cbopart.GetText();
                                            
                                                            
	                                        grid.PerformCallback('gridload' + '|' + pPlandate + '|' + pSupplierDeliver + '|' + pRemaining + '|' + psj + '|' + pDateFrom + '|' + pDateTo + '|' + pSupplier + '|' + pPartcode + '|' + pPono + '|' + pKanban);
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
                    Width="100%" KeyFieldName="colno;colpartno;h_poorder;H_SURATJALAN;H_SJ;h_idxorder;H_SUPPLIER" ClientInstanceName="grid">
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
}" Init="function(s, e) {
	dtfrom.SetText(s.cpdtfrom);
    dtto.SetText(s.cpdtto); 
	dt1.SetText(s.cpdt1);
	rbdeliver.SetValue(s.cpdeliver);
	rbreceiving.SetValue(s.cpreceive);
	rbkanban.SetValue(s.cpkanban);

}" />
                    <Columns>
                        <dx:GridViewDataHyperLinkColumn Caption=" " FieldName="coldetail" 
                            Name="coldetail" VisibleIndex="1" Width="70px">
                            <PropertiesHyperLinkEdit TextField="coldetailname">
                            </PropertiesHyperLinkEdit>
                            <HeaderStyle Font-Names="Tahoma" Font-Size="8pt" Font-Underline="False" 
                                Wrap="True" />
                            <CellStyle Font-Names="Tahoma" Font-Size="8pt" VerticalAlign="Middle">
                            </CellStyle>
                        </dx:GridViewDataHyperLinkColumn>
                        <dx:GridViewDataTextColumn Caption="NO" FieldName="colno" Name="colno" 
                            VisibleIndex="2" Width="30px">
                            <HeaderStyle Font-Names="Tahoma" Font-Size="8pt" Font-Underline="False" 
                                Wrap="True" HorizontalAlign="Center" />
                            <CellStyle Font-Names="Tahoma" Font-Size="8pt">
                            </CellStyle>
                        </dx:GridViewDataTextColumn>
                        <dx:GridViewDataTextColumn Caption="PERIOD" FieldName="colperiod" 
                            Name="colperiod" VisibleIndex="3" Width="60px">
                            <HeaderStyle Font-Names="Tahoma" Font-Size="8pt" Font-Underline="False" 
                                Wrap="True" HorizontalAlign="Center"/>
                            <CellStyle Font-Names="Tahoma" Font-Size="8pt">
                            </CellStyle>
                        </dx:GridViewDataTextColumn>
                        <dx:GridViewDataTextColumn Caption="AFFILIATE CODE" 
                            FieldName="colaffiliatecode" Name="colaffiliatecode" VisibleIndex="4" 
                            Width="70px">
                            <HeaderStyle Font-Names="Tahoma" Font-Size="8pt" Font-Underline="False" 
                                Wrap="True" HorizontalAlign="Center"/>
                            <CellStyle Font-Names="Tahoma" Font-Size="8pt">
                            </CellStyle>
                        </dx:GridViewDataTextColumn>
                        <dx:GridViewDataTextColumn Caption="AFFILIATE NAME" 
                            FieldName="colaffiliatename" Name="colaffiliatename" VisibleIndex="5" 
                            Width="0px">
                            <HeaderStyle Font-Names="Tahoma" Font-Size="8pt" Font-Underline="False" 
                                Wrap="True" HorizontalAlign="Center"/>
                            <CellStyle Font-Names="Tahoma" Font-Size="8pt">
                            </CellStyle>
                        </dx:GridViewDataTextColumn>
                        <dx:GridViewDataTextColumn Caption="PO NO." FieldName="colpono" Name="colpono" 
                            VisibleIndex="6" Width="120px">
                            <HeaderStyle Font-Names="Tahoma" Font-Size="8pt" Font-Underline="False" 
                                Wrap="True" HorizontalAlign="Center"/>
                            <CellStyle Font-Names="Tahoma" Font-Size="8pt">
                            </CellStyle>
                        </dx:GridViewDataTextColumn>
                        <dx:GridViewDataTextColumn Caption="SUPPLIER CODE" FieldName="colsuppliercode" 
                            Name="colsuppliercode" VisibleIndex="7" Width="70px">
                            <HeaderStyle Font-Names="Tahoma" Font-Size="8pt" Font-Underline="False" 
                                Wrap="True" HorizontalAlign="Center"/>
                            <CellStyle Font-Names="Tahoma" Font-Size="8pt">
                            </CellStyle>
                        </dx:GridViewDataTextColumn>
                        <dx:GridViewDataTextColumn Caption="SUPPLIER NAME" FieldName="colsuppliername" 
                            Name="colsuppliername" VisibleIndex="8" Width="0px">
                            <HeaderStyle Font-Names="Tahoma" Font-Size="8pt" Font-Underline="False" 
                                Wrap="True" HorizontalAlign="Center"/>
                            <CellStyle Font-Names="Tahoma" Font-Size="8pt" HorizontalAlign="Center">
                            </CellStyle>
                        </dx:GridViewDataTextColumn>
                        <dx:GridViewDataTextColumn Caption="PO KANBAN" FieldName="colpokanban" Name="colpokanban" 
                            VisibleIndex="11" Width="50px">
                            <HeaderStyle Font-Names="Tahoma" Font-Size="8pt" Font-Underline="False" 
                                Wrap="True" HorizontalAlign="Center" />
                            <CellStyle Font-Names="Tahoma" Font-Size="8pt">
                            </CellStyle>
                            
                        </dx:GridViewDataTextColumn>
                        <dx:GridViewDataTextColumn Caption="KANBAN NO." FieldName="colkanbanno" 
                            Name="colkanbanno" VisibleIndex="12" Width="0px">
                            <HeaderStyle Font-Names="Tahoma" Font-Size="8pt" Font-Underline="False" 
                                Wrap="True" HorizontalAlign="Center"/>
                            <CellStyle Font-Names="Tahoma" Font-Size="8pt" HorizontalAlign="Center">
                            </CellStyle>
                        </dx:GridViewDataTextColumn>
                        <dx:GridViewDataTextColumn Caption="SUPPLIER PLAN DELIVERY DATE" 
                            FieldName="colplandeldate" Name="colplandeldate" VisibleIndex="13" 
                            Width="0px">
                            <HeaderStyle Font-Names="Tahoma" Font-Size="8pt" Font-Underline="False" 
                                Wrap="True" HorizontalAlign="Center"/>
                            <CellStyle Font-Names="Tahoma" Font-Size="8pt" HorizontalAlign="Center">
                            </CellStyle>
                        </dx:GridViewDataTextColumn>
                        <dx:GridViewDataTextColumn Caption="SUPPLIER DELIVERY DATE" 
                            FieldName="coldeldate" Name="coldeldate" VisibleIndex="10" Width="80px">
                            <HeaderStyle Font-Names="Tahoma" Font-Size="8pt" Font-Underline="False" 
                                Wrap="True" HorizontalAlign="Center"/>
                            <CellStyle Font-Names="Tahoma" Font-Size="8pt">
                            </CellStyle>
                        </dx:GridViewDataTextColumn>
                        <dx:GridViewDataTextColumn Caption="SUPPLIER SURAT JALAN NO." FieldName="colsj" 
                            Name="colsj" VisibleIndex="9" Width="100px">
                            <HeaderStyle Font-Names="Tahoma" Font-Size="8pt" Font-Underline="False" 
                                Wrap="True" HorizontalAlign="Center"/>
                            <CellStyle Font-Names="Tahoma" Font-Size="8pt">
                            </CellStyle>
                        </dx:GridViewDataTextColumn>
                        <dx:GridViewDataTextColumn Caption="PART NO." FieldName="colpartno" 
                            Name="colpartno" VisibleIndex="14" Width="0px">
                            <HeaderStyle Font-Names="Tahoma" Font-Size="8pt" Font-Underline="False" 
                                Wrap="True" HorizontalAlign="Center"/>
                            <CellStyle Font-Names="Tahoma" Font-Size="8pt">
                            </CellStyle>
                        </dx:GridViewDataTextColumn>
                        <dx:GridViewDataTextColumn Caption="PART NAME" 
                            FieldName="colpartname" Name="colpartname" VisibleIndex="15" Width="0px">
                            <HeaderStyle Font-Names="Tahoma" Font-Size="8pt" Font-Underline="False" 
                                Wrap="True" HorizontalAlign="Center"/>
                            <CellStyle Font-Names="Tahoma" Font-Size="8pt">
                            </CellStyle>
                        </dx:GridViewDataTextColumn>
                        <dx:GridViewDataTextColumn Caption="UOM" FieldName="coluom" Name="coluom" 
                            VisibleIndex="16" Width="0px">
                            <HeaderStyle Font-Names="Tahoma" Font-Size="8pt" Font-Underline="False" 
                                Wrap="True" HorizontalAlign="Center"/>
                            <CellStyle Font-Names="Tahoma" Font-Size="8pt" HorizontalAlign="Right">
                            </CellStyle>
                        </dx:GridViewDataTextColumn>
                        <dx:GridViewDataTextColumn Caption="SUPPLIER DELIVERY QTY" FieldName="coldeliveryqty" 
                            Name="coldeliveryqty" VisibleIndex="17" Width="0px">
                            <HeaderStyle Font-Names="Tahoma" Font-Size="8pt" Font-Underline="False" 
                                Wrap="True" HorizontalAlign="Center"/>
                            <CellStyle Font-Names="Tahoma" Font-Size="8pt" HorizontalAlign="Right">
                            </CellStyle>
                        </dx:GridViewDataTextColumn>
                        <dx:GridViewDataTextColumn Caption="GOOD RECEIVEING QTY" FieldName="colreceiveqty" 
                            Name="colreceiveqty" VisibleIndex="18" Width="0px">
                            <HeaderStyle Font-Names="Tahoma" Font-Size="8pt" Font-Underline="False" 
                                Wrap="True" HorizontalAlign="Center"/>
                            <CellStyle Font-Names="Tahoma" Font-Size="8pt" HorizontalAlign="Right">
                            </CellStyle>
                        </dx:GridViewDataTextColumn>
                        <dx:GridViewDataTextColumn Caption="REMAINING RECEIVE QTY" 
                            FieldName="colremaining" Name="colremaining" VisibleIndex="20" 
                            Width="0px">
                            <HeaderStyle Font-Names="Tahoma" Font-Size="8pt" Font-Underline="False" 
                                Wrap="True" HorizontalAlign="Center"/>
                            <CellStyle Font-Names="Tahoma" Font-Size="8pt" HorizontalAlign="Right">
                            </CellStyle>
                        </dx:GridViewDataTextColumn>
                        <dx:GridViewDataTextColumn Caption="RECEIVED DATE" FieldName="colreceivedate" 
                            Name="colreceivedate" VisibleIndex="21" Width="80px">
                            <HeaderStyle Font-Names="Tahoma" Font-Size="8pt" Font-Underline="False" 
                                Wrap="True" HorizontalAlign="Center"/>
                            <CellStyle Font-Names="Tahoma" Font-Size="8pt">
                            </CellStyle>
                        </dx:GridViewDataTextColumn>
                        <dx:GridViewDataTextColumn Caption="RECEIVED BY" FieldName="colreceiveby" 
                            Name="colreceiveby" VisibleIndex="22" Width="80px">
                        </dx:GridViewDataTextColumn>
                        <dx:GridViewDataTextColumn 
                            FieldName="H_POORDER" Name="H_POORDER" 
                            VisibleIndex="23" Width="0px">
                            <HeaderStyle Font-Names="Tahoma" Font-Size="8pt" Font-Underline="False" 
                                Wrap="True" HorizontalAlign="Center"/>
                            <CellStyle Font-Names="Tahoma" Font-Size="8pt">
                            </CellStyle>
                        </dx:GridViewDataTextColumn>
                        
                        <dx:GridViewDataTextColumn Caption="DEFECT RECEIVING QTY" FieldName="coldefect" 
                            Name="coldefect" VisibleIndex="19" Width="0px">
                            <CellStyle HorizontalAlign="Right">
                            </CellStyle>
                        </dx:GridViewDataTextColumn>
                        
                    </Columns>
                    
                   <%--<SettingsPager PageSize="9" 
                        NumericButtonCount="100">
                        <Summary Text="Page {0} of {1} [{2} record(s)]" 
                        AllPagesText="Page {0} of {1} " />
                        <Summary AllPagesText="Page {0} of {1} " Text="Page {0} of {1} [{2} record(s)]"></Summary>
                   </SettingsPager> --%>

                    
                    <Settings ShowHorizontalScrollBar="True" ShowVerticalScrollBar="True" ShowGroupButtons="False"
                        ShowStatusBar="Hidden" VerticalScrollableHeight="250" />
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
