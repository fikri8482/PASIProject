<%@ Page Title="" Language="vb" AutoEventWireup="false" MasterPageFile="~/SiteContent.Master" CodeBehind="GoodReceivingReport.aspx.vb" Inherits="PASISystem.GoodReceivingReport" %>

<%@ Register Assembly="DevExpress.XtraReports.v14.1.Web, Version=14.1.6.0, Culture=neutral, PublicKeyToken=b88d1754d700e49a"
    Namespace="DevExpress.XtraReports.Web" TagPrefix="dx" %>
<%@ Register assembly="DevExpress.Web.v14.1, Version=14.1.6.0, Culture=neutral, PublicKeyToken=b88d1754d700e49a" namespace="DevExpress.Web.ASPxEditors" tagprefix="dx" %>
<asp:Content ID="Content1" ContentPlaceHolderID="head" runat="server">
</asp:Content>
<asp:Content ID="Content2" ContentPlaceHolderID="Content" runat="server">
    <table style="width:100%;">
        <tr>
            <td>
                <dx:ASPxDocumentViewer ID="Viewer" runat="server">
                </dx:ASPxDocumentViewer>
            </td>
        </tr>
    </table>

    <div style="height:8px;"></div>
    
    <table id="button" style=" width:100%;">
        <tr>
            <td valign="top" align="left">
                <dx:ASPxButton ID="btnSubMenu" runat="server" Text="BACK"
                    Font-Names="Tahoma" Width="85px" Font-Size="8pt">
                </dx:ASPxButton>
            </td>
            <td valign="top" align="right" style="width: 50px;">
            </td>
            <td valign="top" align="right" style="width: 50px;">      
                                     
            </td>
            <td valign="top" align="right" style="width: 50px;">      
                
            </td>
            <td valign="top" align="right" style="width: 50px;">
                
            </td>
            <td valign="top" align="right" style="width: 50px;">
               
            </td>
           
            <td valign="top" align="right" style="width: 50px;">
                
            </td>            
            <td align="right" style="width:80px;">                                   
                
            </td>
            <td valign="top" align="right" style="width: 50px;">
                
            </td>
        </tr>
    </table>
</asp:Content>
