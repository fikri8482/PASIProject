<%@ Page Title="" Language="vb" AutoEventWireup="false" MasterPageFile="~/SiteContent.Master" CodeBehind="PrintForecast.aspx.vb" Inherits="PASISystem.PrintForecast" %>

<%@ Register Assembly="DevExpress.Web.v14.1, Version=14.1.6.0, Culture=neutral, PublicKeyToken=b88d1754d700e49a"
    Namespace="DevExpress.Web.ASPxEditors" TagPrefix="dx" %>
<%@ Register Assembly="DevExpress.Web.v14.1, Version=14.1.6.0, Culture=neutral, PublicKeyToken=b88d1754d700e49a"
    Namespace="DevExpress.Web.ASPxPager" TagPrefix="dx" %>
<%@ Register Assembly="DevExpress.Web.v14.1, Version=14.1.6.0, Culture=neutral, PublicKeyToken=b88d1754d700e49a"
    Namespace="DevExpress.Web.ASPxGridView" TagPrefix="dx" %>
<%@ Register Assembly="DevExpress.Web.v14.1, Version=14.1.6.0, Culture=neutral, PublicKeyToken=b88d1754d700e49a"
    Namespace="DevExpress.Web.ASPxRoundPanel" TagPrefix="dx" %>
<%@ Register Assembly="DevExpress.Web.v14.1, Version=14.1.6.0, Culture=neutral, PublicKeyToken=b88d1754d700e49a"
    Namespace="DevExpress.Web.ASPxPanel" TagPrefix="dx" %>

<%@ Register assembly="DevExpress.Web.v14.1, Version=14.1.6.0, Culture=neutral, PublicKeyToken=b88d1754d700e49a" namespace="DevExpress.Web.ASPxGridView.Export" tagprefix="dx" %>

<asp:Content ID="Content1" ContentPlaceHolderID="head" runat="server">
    <style type="text/css">
        .style6
        {
            width: 137px;
        }
        .style11
        {
            width: 84px;
        }
        .style12
        {
            width: 173px;
        }
        
        .dxflEmptyItem
        {
            height: 21px;
        }
        
        .style25
        {
            width: 1001px;
            height: 20px;
        }
        .style26
        {
            width: 708px;
        }
        .style28
        {
            width: 2px;
        }
        .style29
        {
            width: 179px;
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
            if (currentColumnName == "cols" ) {

                e.cancel = false;
            } else {

                e.cancel = true;
            }


            currentEditableVisibleIndex = e.visibleIndex;
        }
    </script>
    <table align="center" width="100%">
        <tr>
            <td align="left" class="style26" width="100%">
                <table style="border: thin solid #808080; width: 100%;" width="100%">
                    <tr>
                        <td class="style11" align="left" height="26px">
                            <dx:ASPxLabel ID="ASPxLabel1" runat="server" Font-Names="Tahoma" Font-Size="8pt"
                                Text="PERIOD">
                            </dx:ASPxLabel>
                        </td>
                        <td class="style6" rowspan="1" height="26px">
                            <dx:ASPxTimeEdit ID="dtPeriodFrom" runat="server" 
                                ClientInstanceName="dtPeriodFrom" DisplayFormatString="yyyy-MM" 
                                EditFormat="Custom" EditFormatString="yyyy-MM" Width="110px">                                              
                            </dx:ASPxTimeEdit>
                        </td>
                        <td align="center" class="style28" height="26px">
                            &nbsp;</td>
                        <td class="style12" align="left" height="26px">                            
                        </td>
                        <td align="left" class="style29" height="26px">
                        </td>
                        <td align="left" class="style12" height="26px">
                            &nbsp;
                        </td>
                    </tr>
                    <tr>
                        <td class="style11" align="left" height="26px">
                            <dx:ASPxLabel ID="ASPxLabel14" runat="server" Font-Names="Tahoma" Font-Size="8pt"
                                Text="AFFILIATE CODE/NAME" Width="137px">
                            </dx:ASPxLabel>
                        </td>
                        <td class="style6" align="left" height="26px">
                            <dx:ASPxComboBox ID="cboaffiliate" runat="server" ClientInstanceName="cboaffiliate"
                                Font-Names="Tahoma" TextFormatString="{0}" Font-Size="8pt">
                                <ClientSideEvents SelectedIndexChanged="function(s, e) {
	                                            txtaffiliate.SetText(cboaffiliate.GetSelectedItem().GetColumnText(1));
                                                }" />
                            </dx:ASPxComboBox>
                        </td>
                        <td align="center" class="style28" height="26px">
                            &nbsp;</td>
                        <td align="left" class="style12" height="26px">
                            <dx:ASPxTextBox ID="txtaffiliate" runat="server" Width="230px" Font-Names="Tahoma"
                                Font-Size="8pt" BackColor="#CCCCCC" ReadOnly="True" Height="16px" 
                                ClientInstanceName="txtaffiliate">
                                <ClientSideEvents TextChanged="function(s, e) {
	                                lblerrmessage.SetText('');
                                }" />
                            </dx:ASPxTextBox>
                        </td>
                        <td align="right" class="style29" height="26px">
                            <dx:ASPxButton ID="btnsearch" runat="server" Text="SEARCH" Width="85px" Font-Names="Tahoma"
                                Font-Size="8pt" ClientInstanceName="btnsearch" AutoPostBack="False">
                            </dx:ASPxButton>
                        </td>
                        <td align="left" class="style12" height="26px">
                            <dx:ASPxButton ID="btnclear" runat="server" Text="CLEAR" Width="85px" Font-Names="Tahoma"
                                Font-Size="8pt">
                            </dx:ASPxButton>
                        </td>
                    </tr>
                    </table>
            </td>
        </tr>
    </table>
    <table width="100%">
        <td align="left" bgcolor="White" class="style25" width="100%">
            <table align="left" width="100%">
                <tr align="left">
                    <td width="100%" height="16px" style="border-top-style: solid; border-top-width: thin;
                        border-top-color: #808080; border-bottom-style: solid; border-bottom-width: thin;
                        border-bottom-color: #808080" align="left">
                        <dx:ASPxLabel ID="lblerrmessage" runat="server" Font-Names="Tahoma" Font-Size="8pt"
                            Text="ERROR MESSAGE" Width="100%" ClientInstanceName="lblerrmessage">
                        </dx:ASPxLabel>
                    </td>
                </tr>
            </table>
        </td>
    </table>
    <table style="width: 100%; height: 195px;" align="left">
        <tr>
            <td align="left" width="100%" height="0px">
                <dx:ASPxGridView ID="grid" runat="server" AutoGenerateColumns="False" Width="100%"
                    KeyFieldName="PartNo;URUT;no;Period" ClientInstanceName="grid">                    
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
	                        lblerrmessage.SetText('');}" 
                        BatchEditStartEditing="OnBatchEditStartEditing" 
                        CallbackError="function(s, e) {e.handled = true;}" />
                    
                    <SettingsPager Mode="ShowAllRecords" Visible="False">
                    </SettingsPager>
                    
                    <SettingsEditing Mode="Batch" NewItemRowPosition="Bottom">                        
                        <BatchEditSettings ShowConfirmOnLosingChanges="False"></BatchEditSettings>
                    </SettingsEditing>

                    <Settings ShowHorizontalScrollBar="True" ShowVerticalScrollBar="True" ShowGroupButtons="False"
                        ShowStatusBar="Hidden" VerticalScrollableHeight="250" />

                    <StylesEditors ButtonEditCellSpacing="0">
                        <ProgressBar Height="21px">
                        </ProgressBar>
                    </StylesEditors>
                </dx:ASPxGridView>
            </td>
        </tr>
    </table>
    <table style="width: 100%;" width="100%">
        <tr>
            <td align="left">
                <dx:ASPxButton ID="btnsubmenu" runat="server" Text="SUB MENU" Width="90px" Font-Names="Tahoma"
                    Font-Size="8pt">
                </dx:ASPxButton>
            </td>
            <td align="right">
                <dx:ASPxButton ID="btnprintcard" runat="server" Text="DOWNLOAD TO EXCEL" Width="100px"
                    Font-Names="Tahoma" Font-Size="8pt" AutoPostBack="False">
                </dx:ASPxButton>                
                <dx:ASPxGridViewExporter ID="GridExporter" runat="server" GridViewID="grid">
                </dx:ASPxGridViewExporter>
            </td>
        </tr>
    </table>
    <br />
</asp:Content>
