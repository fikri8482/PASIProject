<%@ Page Title="" Language="vb" AutoEventWireup="false" MasterPageFile="~/SiteContent.Master"
    CodeBehind="PrintPreBookingList.aspx.vb" Inherits="PASISystem.PrintPreBookingList" %>

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

<%@ Register assembly="DevExpress.Web.v14.1, Version=14.1.6.0, Culture=neutral, PublicKeyToken=b88d1754d700e49a" namespace="DevExpress.Web.ASPxGlobalEvents" tagprefix="dx" %>
<%@ Register assembly="DevExpress.Web.v14.1, Version=14.1.6.0, Culture=neutral, PublicKeyToken=b88d1754d700e49a" namespace="DevExpress.Web.ASPxCallback" tagprefix="dx" %>

<asp:Content ID="Content1" ContentPlaceHolderID="head" runat="server">
    <style type="text/css">
        .style6
        {
        width: 127px;
    }
        .style11
        {
        width: 60px;
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
        .style27
    {
        width: 169px;
    }
        </style>
</asp:Content>
<asp:Content ID="Content2" ContentPlaceHolderID="Content" runat="server">
    <script language="javascript" type="text/javascript">
        
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
                        <td class="style27" align="left" height="26px">                            
                            <dx:ASPxCallback ID="Approve" runat="server" ClientInstanceName="Approve">
                                <ClientSideEvents EndCallback="function(s, e) {           

            var pMsg = s.cpMessage;
            if (pMsg != '') {
                if (pMsg.substring(1,5) == '1006' || pMsg.substring(1,5) == '1009' ) {
                    lblerrmessage.GetMainElement().style.color = 'Blue';
                } else {
                    lblerrmessage.GetMainElement().style.color = 'Red';
                }
                lblerrmessage.SetText(pMsg);
            } else {
                lblerrmessage.SetText('');
            }

            btnprintcard.SetText(s.cpButton);
            delete s.cpMessage;
        }" />
                            </dx:ASPxCallback>
                        </td>
                        <td align="left" class="style12" height="26px">
                            &nbsp;</td>
                    </tr>
                    <tr>
                        <td class="style11" align="left" height="26px">
                            <dx:ASPxLabel ID="ASPxLabel14" runat="server" Font-Names="Tahoma" Font-Size="8pt"
                                Text="AFFILIATE CODE/NAME" Width="137px">
                            </dx:ASPxLabel>
                        </td>
                        <td class="style6" align="left" height="26px">
                            <table style="width:100%;">
                                <tr>
                                    <td>
                            <dx:ASPxComboBox ID="cboaffiliate" runat="server" ClientInstanceName="cboaffiliate"
                                Font-Names="Tahoma" TextFormatString="{0}" Font-Size="8pt" Width="120px">
                                <ClientSideEvents SelectedIndexChanged="function(s, e) {
	                                            txtaffiliate.SetText(cboaffiliate.GetSelectedItem().GetColumnText(1));
                                                }" />
                            </dx:ASPxComboBox>
                                    </td>
                                    <td>
                            <dx:ASPxTextBox ID="txtaffiliate" runat="server" Width="230px" Font-Names="Tahoma"
                                Font-Size="8pt" BackColor="#CCCCCC" ReadOnly="True" Height="16px" 
                                ClientInstanceName="txtaffiliate">
                                <ClientSideEvents TextChanged="function(s, e) {
	                                lblerrmessage.SetText('');
                                }" />
                            </dx:ASPxTextBox>
                                    </td>
                                    <td>
                            <dx:ASPxButton ID="btnclear" runat="server" Text="CLEAR" Width="85px" Font-Names="Tahoma"
                                Font-Size="8pt">
                            </dx:ASPxButton>
                                    </td>
                                </tr>
                            </table>
                        </td>
                        <td align="left" class="style27" height="26px">
                            &nbsp;</td>
                        <td align="left" class="style12" height="26px">
                            &nbsp;</td>
                    </tr>
                    <tr>
                        <td class="style11" align="left" height="26px">
                            &nbsp;</td>
                        <td class="style6" align="left" height="26px">
                            &nbsp;</td>
                        <td align="left" class="style27" height="26px" colspan="1" rowspan="1">
                            &nbsp;</td>
                        <td align="right" class="style12" height="26px" width="85px">
                            &nbsp;</td>
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
    
    <table style="width: 100%;" width="100%">
        <tr>
            <td align="left">
                <dx:ASPxButton ID="btnsubmenu" runat="server" Text="SUB MENU" Width="90px" Font-Names="Tahoma"
                    Font-Size="8pt">
                </dx:ASPxButton>
            </td>
            <td align="right">
                <dx:ASPxButton ID="btnprintcard" runat="server" Text="PRINT PRE-BOOKING" Width="100px"
                    Font-Names="Tahoma" Font-Size="8pt" AutoPostBack="False">
                    <ClientSideEvents Click="function(s, e) {
	Approve.PerformCallback();
}" />
                </dx:ASPxButton>                
            </td>
        </tr>
    </table>
    <br />
</asp:Content>
