<%@ Page Title="" Language="vb" AutoEventWireup="false" MasterPageFile="~/SiteContent.Master" CodeBehind="viewDeliveryToFor.aspx.vb" Inherits="PASISystem.viewDeliveryToFor" %>
<%@ Register Assembly="DevExpress.XtraReports.v14.1.Web, Version=14.1.6.0, Culture=neutral, PublicKeyToken=b88d1754d700e49a"
    Namespace="DevExpress.XtraReports.Web" TagPrefix="dx" %>
<%@ Register Assembly="DevExpress.Web.v14.1, Version=14.1.6.0, Culture=neutral, PublicKeyToken=b88d1754d700e49a" Namespace="DevExpress.Web.ASPxEditors" TagPrefix="dx1" %>
<asp:Content ID="Content1" ContentPlaceHolderID="head" runat="server">
</asp:Content>
<asp:Content ID="Content2" ContentPlaceHolderID="Content" runat="server">
<table style="width:100%;">
    <tr>
            <td align="left">
                <dx:ASPxDocumentViewer ID="ASPxDocumentViewer1" runat="server" Width="100%">
                    <SettingsReportViewer TableLayout="False" />
                </dx:ASPxDocumentViewer>
            </td>
        </tr>
    <tr>
        <td align = "left">              
            <dx1:ASPxButton ID="btnBack" runat="server" Width="90px" Font-Names="Tahoma" Font-Size="8pt"
                Text="BACK" ClientInstanceName="btnBack" AutoPostBack="False">     
            </dx1:ASPxButton>
        </td>
    </tr>
</table>
</asp:Content>
