<%@ Page Title="" Language="vb" AutoEventWireup="false" MasterPageFile="~/SiteContent.Master" CodeBehind="InvFromPASIList.aspx.vb" Inherits="AffiliateSystem.InvFromPASIList" %>
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
                                Text="PASI DELIVERY DATE">
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
                            <dx1:ASPxLabel ID="ASPxLabel13" runat="server" Font-Names="Tahoma" Font-Size="8pt"
                                Text="AFFILIATE CODE / NAME" Visible="False">
                            </dx1:ASPxLabel>
                        </td>
            <td align="left">
                <dx1:ASPxComboBox ID="cboaffiliate" runat="server" ClientInstanceName="cboaffiliate"
                                Font-Names="Tahoma" TextFormatString="{0}" Font-Size="8pt" 
                    Visible="False">
                                <ClientSideEvents SelectedIndexChanged="function(s, e) {
	                                            txtaffiliate.SetText(cboaffiliate.GetSelectedItem().GetColumnText(1));
                                                }" />
                            </dx1:ASPxComboBox>
            </td>
            <td align="left">
                <dx1:ASPxTextBox ID="txtaffiliate" runat="server" Font-Names="Tahoma" 
                    Font-Size="8pt" Width="170px" BackColor="Silver" 
                    ClientInstanceName="txtaffiliate" ReadOnly="True" Visible="False">
                    <ClientSideEvents TextChanged="function(s, e) {
	lblerrmessage.SetText('');
}" />
                </dx1:ASPxTextBox>
            </td>
        </tr>
        <tr>
            <td align="left">
                            <dx1:ASPxLabel ID="ASPxLabel2" runat="server" Font-Names="Tahoma" Font-Size="8pt"
                                Text="ALREADY INVOICE BY PASI">
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
                            <dx1:ASPxLabel ID="ASPxLabel5" runat="server" Font-Names="Tahoma" Font-Size="8pt"
                                Text="PASI SURAT JALAN NO">
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
                            &nbsp;</td>
            <td align="left">
                &nbsp;</td>
            <td align="left">
                &nbsp;</td>
        </tr>
        <tr>
            <td align="left">
                            <dx1:ASPxLabel ID="ASPxLabel15" runat="server" Font-Names="Tahoma" Font-Size="8pt"
                                Text="PASI INVOICE DATE">
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
                <dx1:ASPxDateEdit ID="dt3" runat="server" ClientInstanceName="dt3" 
                    Font-Names="Tahoma" Font-Size="8pt" EditFormat="Custom" 
                    EditFormatString="dd MMM yyyy" Width="100px">
                </dx1:ASPxDateEdit>
                        </td>
                        <td>
                            <dx1:ASPxLabel ID="ASPxLabel16" runat="server" Font-Names="Tahoma" Font-Size="8pt"
                                Text="~">
                            </dx1:ASPxLabel>
                        </td>
                        <td style="font-weight: 700">
                <dx1:ASPxDateEdit ID="dt4" runat="server" ClientInstanceName="dt4" 
                    Font-Names="Tahoma" Font-Size="8pt" EditFormat="Custom" 
                    EditFormatString="dd MMM yyyy" Width="100px">
                </dx1:ASPxDateEdit>
                        </td>
                    </tr>
                </table>
            </td>
            <td align="left">
                            &nbsp;</td>
            <td align="left">
                &nbsp;</td>
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
                    Width="100%" KeyFieldName="no;url" ClientInstanceName="grid">
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
                            VisibleIndex="1" Width="30px">
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
                        <dx:GridViewDataTextColumn Caption="PASI INVOICE NO" 
                            FieldName="pasiinvno" VisibleIndex="4" Width="130px">
                            <HeaderStyle Font-Names="Tahoma" Font-Size="8pt" HorizontalAlign="Center" />
                            <CellStyle Font-Names="Tahoma" Font-Size="8pt">
                            </CellStyle>
                        </dx:GridViewDataTextColumn>                        
                        <dx:GridViewDataTextColumn Caption="PASI DELIVERY DATE" FieldName="pasideldate" 
                            VisibleIndex="13">
                            <HeaderStyle Font-Names="Tahoma" Font-Size="8pt" HorizontalAlign="Center" />
                            <CellStyle Font-Names="Tahoma" Font-Size="8pt">
                            </CellStyle>
                        </dx:GridViewDataTextColumn>                        
                        <dx:GridViewDataTextColumn Caption="PASI SURAT JALAN NO" 
                            FieldName="pasisj" VisibleIndex="5" Width="130px">
                            <HeaderStyle Font-Names="Tahoma" Font-Size="8pt" HorizontalAlign="Center" />
                            <CellStyle Font-Names="Tahoma" Font-Size="8pt">
                            </CellStyle>
                        </dx:GridViewDataTextColumn>                                               
                        <dx:GridViewDataTextColumn Caption="PASI INVOICE DATE" FieldName="pasiinvdate" 
                            VisibleIndex="6">
                            <HeaderStyle Font-Names="Tahoma" Font-Size="8pt" HorizontalAlign="Center" />
                            <cellstyle font-names="Tahoma" font-size="8pt">
                            </cellstyle>
                        </dx:GridViewDataTextColumn>
                        <dx:GridViewDataTextColumn Caption="AFFILIATE RECEIVE DATE" 
                            FieldName="affrecdate" VisibleIndex="7" 
                            Width="70px">
                            <HeaderStyle Font-Names="Tahoma" Font-Size="8pt" Font-Underline="False" 
                                Wrap="True" HorizontalAlign="Center"/>
                            <CellStyle Font-Names="Tahoma" Font-Size="8pt" HorizontalAlign="Center">
                            </CellStyle>
                        </dx:GridViewDataTextColumn>
                        <dx:GridViewDataTextColumn Caption="AFFILIATE RECEIVE BY" 
                            FieldName="affrecby" VisibleIndex="8" Width="80px">
                            <HeaderStyle Font-Names="Tahoma" Font-Size="8pt" Font-Underline="False" 
                                Wrap="True" HorizontalAlign="Center"/>
                            <CellStyle Font-Names="Tahoma" Font-Size="8pt">
                            </CellStyle>
                        </dx:GridViewDataTextColumn>                                                
                    </Columns>                    
                    <SettingsBehavior AllowSelectByRowClick="True" AllowSort="False"
                        ColumnResizeMode="Control" EnableRowHotTrack="True" />

<SettingsBehavior AllowSort="False" AllowSelectByRowClick="True" ColumnResizeMode="Control" 
                        EnableRowHotTrack="True"></SettingsBehavior>

                    <SettingsPager PageSize="16" Position="Top">
                        <Summary Text="Page {0} of {1} [{2} record(s)]"
                                  AllPagesText="Page {0} of {1} " />
                        <Summary AllPagesText="Page {0} of {1} " Text="Page {0} of {1} [{2} record(s)]"></Summary>
                    </SettingsPager>
                    <%--<SettingsEditing Mode="Batch" NewItemRowPosition="Bottom">
                        <BatchEditSettings ShowConfirmOnLosingChanges="False" />
                    </SettingsEditing>--%>
                    <Settings ShowHorizontalScrollBar="True" ShowVerticalScrollBar="True"
                        ShowGroupButtons="False" ShowStatusBar="Hidden"
                        VerticalScrollableHeight="190" />

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
