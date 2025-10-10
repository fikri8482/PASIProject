<%@ Page Title="" Language="vb" AutoEventWireup="false" MasterPageFile="~/SiteContent.Master" CodeBehind="GoodReceivingReportCR.aspx.vb" Inherits="AffiliateSystem.ShippingViewReportExportCR" %>
<%@ Register assembly="DevExpress.Web.v14.1, Version=14.1.6.0, Culture=neutral, PublicKeyToken=b88d1754d700e49a" namespace="DevExpress.Web.ASPxEditors" tagprefix="dx1" %>
<%@ Register assembly="CrystalDecisions.Web, Version=13.0.2000.0, Culture=neutral, PublicKeyToken=692fbea5521e1304" namespace="CrystalDecisions.Web" tagprefix="CR" %>
<asp:Content ID="Content1" ContentPlaceHolderID="head" runat="server">
    <style type="text/css">
    .style1
    {
        height: 17px;
    }
</style>
</asp:Content>
<asp:Content ID="Content2" ContentPlaceHolderID="Content" runat="server">
    <div style="width: 100%; margin-left: auto; margin-right: auto;">

        <CR:CrystalReportViewer ID="CrystalReportViewer1" runat="server" 
                AutoDataBind="true" EnableDatabaseLogonPrompt="False" 
                EnableParameterPrompt="False" ToolPanelView="None" BestFitPage="False" 
                HasCrystalLogo="False" HasDrilldownTabs="False" />

    </div>
    <table style="width:100%;">
        
<%--        <tr>
            <td align="left" class="style1">
                
            </td>

        </tr>--%>
        <tr>
            <td align="left">
                <dx1:ASPxButton ID="btnsubmenu" runat="server" Width="90px" 
                    Font-Names="Tahoma" Font-Size="8pt"
                                Text="BACK">
                </dx1:ASPxButton>
            </td>
        </tr>
    </table>
</asp:Content>
