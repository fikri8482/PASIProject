<%@ Page Title="" Language="vb" AutoEventWireup="false" MasterPageFile="~/SiteContent.Master" CodeBehind="AffReceivingConfExport.aspx.vb" Inherits="PASISystem.AffReceivingConfExport" %>
<%@ Register assembly="DevExpress.Web.v14.1, Version=14.1.6.0, Culture=neutral, PublicKeyToken=b88d1754d700e49a" namespace="DevExpress.Web.ASPxEditors" tagprefix="dx1" %>
<%@ Register assembly="DevExpress.Web.v14.1, Version=14.1.6.0, Culture=neutral, PublicKeyToken=b88d1754d700e49a" namespace="DevExpress.Web.ASPxGridView" tagprefix="dx" %>
<%@ Register assembly="DevExpress.Web.v14.1, Version=14.1.6.0, Culture=neutral, PublicKeyToken=b88d1754d700e49a" namespace="DevExpress.Web.ASPxGridView.Export" tagprefix="dx2" %>
<%@ Register Assembly="DevExpress.Web.v14.1, Version=14.1.6.0, Culture=neutral, PublicKeyToken=b88d1754d700e49a" namespace="DevExpress.Web.ASPxGlobalEvents" TagPrefix="dx" %>
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
        height = height - (height * 58 / 100)
        grid.SetHeight(height);
    }
</script>
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
                <dx1:ASPxDateEdit ID="dtDeliveryDateFrom" runat="server" ClientInstanceName="dtDeliveryDateFrom" 
                    Font-Size="8pt" EditFormat="Custom" EditFormatString="dd MMM yyyy">
                </dx1:ASPxDateEdit>
            </td>
            <td align="left">
                <dx1:ASPxLabel ID="ASPxLabel11" runat="server" Font-Names="Tahoma" Font-Size="8pt"
                    Text="~">
                </dx1:ASPxLabel>
            </td>
            <td align="left">
                <dx1:ASPxDateEdit ID="dtDeliveryDateTo" runat="server" ClientInstanceName="dtDeliveryDateTo" 
                    Font-Size="8pt" EditFormat="Custom" EditFormatString="dd MMM yyyy">
                </dx1:ASPxDateEdit>
            </td>
            <td align="left">
                            <dx1:ASPxLabel ID="ASPxLabel7" runat="server" Font-Names="Tahoma" Font-Size="8pt"
                                Text="AFFILIATE CODE/NAME">
                            </dx1:ASPxLabel>
                        </td>
            <td align="left">
                &nbsp;</td>
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
                            <dx1:ASPxLabel ID="ASPxLabel2" runat="server" Font-Names="Tahoma" Font-Size="8pt"
                                Text="ALREADY INVOICE BY PASI">
                            </dx1:ASPxLabel>
                        </td>
            <td align="left">
                &nbsp;</td>
            <td align="left">
                <dx1:ASPxRadioButtonList ID="rbInvoiceByPasi" runat="server" 
                    ClientInstanceName="rbInvoiceByPasi" Font-Names="Tahoma" Font-Size="8pt" 
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
                &nbsp;</td>
            <td align="left">
                &nbsp;</td>
            <td align="left">
                            <dx1:ASPxLabel ID="ASPxLabel8" runat="server" Font-Names="Tahoma" Font-Size="8pt"
                                Text="ORDER NO.">
                            </dx1:ASPxLabel>
                        </td>
            <td align="left">
                &nbsp;</td>
            <td align="left">
                <dx1:ASPxTextBox ID="txtorderno" runat="server" Font-Names="Tahoma" 
                    Font-Size="8pt" Width="170px" ClientInstanceName="txtorderno">
                </dx1:ASPxTextBox>
            </td>
            <td align="left">
                
            </td>
        </tr>
        <tr>
            <td align="left">
                            <dx1:ASPxLabel ID="ASPxLabel3" runat="server" Font-Names="Tahoma" Font-Size="8pt"
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
                &nbsp;</td>
            <td align="left">
                &nbsp;</td>
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
                            <dx1:ASPxLabel ID="ASPxLabel4" runat="server" Font-Names="Tahoma" Font-Size="8pt"
                                Text="PASI INVOICE DATE">
                            </dx1:ASPxLabel>
                        </td>
            <td align="left">
                <dx1:ASPxCheckBox ID="checkbox2" runat="server" CheckState="Unchecked" 
                    ClientInstanceName="checkbox2" Text=" ">
                </dx1:ASPxCheckBox>
            </td>
            <td align="left">
                <dx1:ASPxDateEdit ID="dtPasiInvoiceDateFrom" runat="server" ClientInstanceName="dtPasiInvoiceDateFrom" 
                    Font-Size="8pt" EditFormat="Custom" EditFormatString="dd MMM yyyy">
                </dx1:ASPxDateEdit>
            </td>
            <td align="left">
                <dx1:ASPxLabel ID="ASPxLabel12" runat="server" Font-Names="Tahoma" Font-Size="8pt"
                    Text="~">
                </dx1:ASPxLabel>
            </td>
            <td align="left">
                <dx1:ASPxDateEdit ID="dtPasiInvoiceDateTo" runat="server" ClientInstanceName="dtPasiInvoiceDateTo" 
                    Font-Size="8pt" EditFormat="Custom" EditFormatString="dd MMM yyyy">
                </dx1:ASPxDateEdit>    
            </td>
            <td align="left">                            
                &nbsp;</td>
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
            <td align="left" class="style1" colspan = "2">
                <dx:ASPxGridView ID="grid" runat="server" AutoGenerateColumns="False" 
                    Width="100%" KeyFieldName="colorderno;colaffiliatecode;colno" ClientInstanceName="grid">
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
	dtDeliveryDateFrom.SetText(s.cpdtfrom);
    dtDeliveryDateTo.SetText(s.cpdtto);
    
    dtPasiInvoiceDateFrom.SetText(s.cpdtfrom);
    dtPasiInvoiceDateTo.SetText(s.cpdtto); 
}" />
                    <Columns>   
                        <dx:GridViewDataHyperLinkColumn Caption=" " FieldName="coldetail" 
                            Name="coldetail" VisibleIndex="1" Width="65px">
                            <PropertiesHyperLinkEdit TextField="coldetailname">
                            </PropertiesHyperLinkEdit>
                            <HeaderStyle Font-Names="Tahoma" Font-Size="8pt" Font-Underline="False" 
                                Wrap="True" />
                            <CellStyle Font-Names="Tahoma" Font-Size="8pt" VerticalAlign="Middle">
                            </CellStyle>
                        </dx:GridViewDataHyperLinkColumn> 
                        <dx:GridViewDataCheckColumn Caption=" " FieldName="act" 
                            Name="act" VisibleIndex="2" Width="30px">
                                <PropertiesCheckEdit ValueChecked="1" ValueType="System.Int32" 
                                    ValueUnchecked="0">
                                </PropertiesCheckEdit>
                        </dx:GridViewDataCheckColumn>
                        <dx:GridViewDataTextColumn Caption="NO" FieldName="colno" Name="colno" 
                            VisibleIndex="3" Width="30px">
                            <HeaderStyle Font-Names="Tahoma" Font-Size="8pt" Font-Underline="False" 
                                Wrap="True" HorizontalAlign="Center" />
                            <CellStyle Font-Names="Tahoma" Font-Size="8pt">
                            </CellStyle>
                        </dx:GridViewDataTextColumn>
                        <dx:GridViewDataTextColumn Caption="PERIOD" FieldName="colperiod" 
                            Name="colperiod" VisibleIndex="4" Width="60px">
                            <HeaderStyle Font-Names="Tahoma" Font-Size="8pt" Font-Underline="False" 
                                Wrap="True" HorizontalAlign="Center"/>
                            <CellStyle Font-Names="Tahoma" Font-Size="8pt">
                            </CellStyle>
                        </dx:GridViewDataTextColumn>
                        <dx:GridViewDataTextColumn Caption="AFFILIATE CODE" 
                            FieldName="colaffiliatecode" Name="colaffiliatecode" VisibleIndex="5" 
                            Width="70px">
                            <HeaderStyle Font-Names="Tahoma" Font-Size="8pt" Font-Underline="False" 
                                Wrap="True" HorizontalAlign="Center"/>
                            <CellStyle Font-Names="Tahoma" Font-Size="8pt">
                            </CellStyle>
                        </dx:GridViewDataTextColumn>
                        <dx:GridViewDataTextColumn Caption="AFFILIATE NAME" 
                            FieldName="colaffiliatename" Name="colaffiliatename" VisibleIndex="6" 
                            Width="120px">
                            <HeaderStyle Font-Names="Tahoma" Font-Size="8pt" Font-Underline="False" 
                                Wrap="True" HorizontalAlign="Center"/>
                            <CellStyle Font-Names="Tahoma" Font-Size="8pt">
                            </CellStyle>
                        </dx:GridViewDataTextColumn>
                        <dx:GridViewDataTextColumn Caption="ORDER NO." 
                            FieldName="colorderno" VisibleIndex="7" 
                            Width="120px">
                            <HeaderStyle Font-Names="Tahoma" Font-Size="8pt" Font-Underline="False" 
                                Wrap="True" HorizontalAlign="Center"/>
                            <CellStyle Font-Names="Tahoma" Font-Size="8pt">
                            </CellStyle>
                        </dx:GridViewDataTextColumn>
                        <dx:GridViewDataTextColumn Caption="SUPPLIER CODE" 
                            FieldName="colsupplierid" VisibleIndex="8" 
                            Width="120px">
                            <HeaderStyle Font-Names="Tahoma" Font-Size="8pt" Font-Underline="False" 
                                Wrap="True" HorizontalAlign="Center"/>
                            <CellStyle Font-Names="Tahoma" Font-Size="8pt">
                            </CellStyle>
                        </dx:GridViewDataTextColumn>
                        <dx:GridViewDataTextColumn Caption="SUPPLIER NAME" FieldName="colsuppliername" Name="colsuppliername" 
                            VisibleIndex="9" Width="80px">
                            <HeaderStyle Font-Names="Tahoma" Font-Size="8pt" Font-Underline="False" 
                                Wrap="True" HorizontalAlign="Center"/>
                            <CellStyle Font-Names="Tahoma" Font-Size="8pt">
                            </CellStyle>
                        </dx:GridViewDataTextColumn>
                        <dx:GridViewDataTextColumn Caption="SUPPLIER PLAN DELIVERY DATE" 
                            FieldName="colsuppplandeldate" VisibleIndex="10" Width="80px">
                            <HeaderStyle Font-Names="Tahoma" Font-Size="8pt" Font-Underline="False" 
                                Wrap="True" HorizontalAlign="Center"/>
                            <CellStyle Font-Names="Tahoma" Font-Size="8pt">
                            </CellStyle>
                        </dx:GridViewDataTextColumn>
                        <dx:GridViewDataTextColumn Caption="SUPPLIER DELIVERY DATE" FieldName="coldeldate" 
                            Name="coldeldate" VisibleIndex="11" Width="80px">
                            <HeaderStyle Font-Names="Tahoma" Font-Size="8pt" Font-Underline="False" 
                                Wrap="True" HorizontalAlign="Center"/>
                            <CellStyle Font-Names="Tahoma" Font-Size="8pt" HorizontalAlign="Center">
                            </CellStyle>
                        </dx:GridViewDataTextColumn>
                        <dx:GridViewDataTextColumn Caption="SUPPLIER SURAT JALAN NO." 
                            FieldName="colsj" Name="colsj" VisibleIndex="12" 
                            Width="80px">
                            <HeaderStyle Font-Names="Tahoma" Font-Size="8pt" Font-Underline="False" 
                                Wrap="True" HorizontalAlign="Center"/>
                            <CellStyle Font-Names="Tahoma" Font-Size="8pt">
                            </CellStyle>
                        </dx:GridViewDataTextColumn>
                        <dx:GridViewDataTextColumn Caption="PASI INVOICE NO." 
                            FieldName="colpasiinvno" 
                            VisibleIndex="13" Width="80px">
                            <HeaderStyle Font-Names="Tahoma" Font-Size="8pt" Font-Underline="False" 
                                Wrap="True" HorizontalAlign="Center" />
                            <CellStyle Font-Names="Tahoma" Font-Size="8pt" HorizontalAlign="Center">
                            </CellStyle>
                            
                        </dx:GridViewDataTextColumn>
                        <dx:GridViewDataTextColumn Caption="PASI INVOICE DATE" 
                            FieldName="colpasiinvdate" VisibleIndex="14" Width="80px">
                            <HeaderStyle Font-Names="Tahoma" Font-Size="8pt" Font-Underline="False" 
                                Wrap="True" HorizontalAlign="Center"/>
                            <CellStyle Font-Names="Tahoma" Font-Size="8pt">
                            </CellStyle>
                        </dx:GridViewDataTextColumn>
                        <dx:GridViewDataTextColumn FieldName="colshipping" VisibleIndex="15" 
                            Width="0px">
                        </dx:GridViewDataTextColumn>
                    </Columns>
                    <SettingsPager Visible="False" Position="Top">
                        <Summary Text="Page {0} of {1} [{2} record(s)]" AllPagesText="Page {0} of {1} " />
                        <Summary AllPagesText="Page {0} of {1} " Text="Page {0} of {1} [{2} record(s)]">
                        </Summary>
                    </SettingsPager>
                    <SettingsEditing Mode="Batch">
                    </SettingsEditing>
                    <Settings ShowHorizontalScrollBar="True" ShowVerticalScrollBar="True" ShowGroupButtons="False"
                        ShowStatusBar="Hidden" VerticalScrollableHeight="250" />
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
            <td align="right" width="50px">
                <dx1:ASPxButton ID="btndeliver" runat="server" Text="INVOICE" Width="90px" Font-Names="Tahoma"
                    Font-Size="8pt" AutoPostBack="False">
                    <ClientSideEvents Click="function(s, e) {
	if (grid.batchEditApi.HasChanges()) {
		grid.UpdateEdit();
        var millisecondsToWait = 200;
            setTimeout(function() {
                document.location.href = 'InvToAffExport.aspx';
            }, millisecondsToWait);  
    } else {    
    lblerrmessage.SetText('[6011] Please Select Data to Deliver!');
	lblerrmessage.GetMainElement().style.color = 'Red';        
    }
}" />
                </dx1:ASPxButton>
            </td>
        </tr>
    </table>
    <dx:ASPxGlobalEvents ID="ge" runat="server" >
        <ClientSideEvents ControlsInitialized="function(s, e) {
	OnControlsInitializedSplitter();
	OnControlsInitializedGrid();
}" />
    </dx:ASPxGlobalEvents>
</asp:Content>
