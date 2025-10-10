<%@ Page Title="" Language="vb" AutoEventWireup="false" MasterPageFile="~/SiteContent.Master" CodeBehind="ShippingViewReportExport.aspx.vb" Inherits="PASISystem.ShippingViewReportExport" %>
<%@ Register assembly="DevExpress.XtraReports.v14.1.Web, Version=14.1.6.0, Culture=neutral, PublicKeyToken=b88d1754d700e49a" namespace="DevExpress.XtraReports.Web" tagprefix="dx" %>
<%@ Register assembly="DevExpress.Web.v14.1, Version=14.1.6.0, Culture=neutral, PublicKeyToken=b88d1754d700e49a" namespace="DevExpress.Web.ASPxEditors" tagprefix="dx1" %>
<asp:Content ID="Content1" ContentPlaceHolderID="head" runat="server">
</asp:Content>
<asp:Content ID="Content2" ContentPlaceHolderID="Content" runat="server">
    <table style="width:100%;">
        <tr>
            <td align="left">
                <dx:ASPxDocumentViewer ID="ASPxDocumentViewer1" runat="server" Width="100%">
                </dx:ASPxDocumentViewer>
            </td>
        </tr>
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
