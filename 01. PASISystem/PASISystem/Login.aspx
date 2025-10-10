<%@ Page Language="vb" AutoEventWireup="false" CodeBehind="Login.aspx.vb" Inherits="PASISystem.Login" %>

<%@ Register Assembly="DevExpress.Web.v14.1, Version=14.1.6.0, Culture=neutral, PublicKeyToken=b88d1754d700e49a" Namespace="DevExpress.Web.ASPxEditors" TagPrefix="dx" %>

<!DOCTYPE html PUBLIC "-//W3C//DTD XHTML 1.0 Transitional//EN" "http://www.w3.org/TR/xhtml1/DTD/xhtml1-transitional.dtd">

<html xmlns="http://www.w3.org/1999/xhtml">
<head runat="server">
    <title>PURCHASING SYSTEM - AFFILIATE</title>
    <link href="~/Styles/images/yazakiIcon.ico" rel="SHORTCUT ICON" type="image/icon" />
    <style type="text/css">
        #LoginTitle
        {
            height: 26px;
        }
        .style1
        {
            width: 192px;
            height: 22px;
        }
    </style>

    <script language="javascript" type="text/javascript" src="Scripts/jsInputValidation.js"></script>
</head>
<body bgcolor="WHITE">
    <form id="frmLogin" runat="server">
    <div>
    <br /><br /><br /><br /><br />
        <center>            
            <div id="LoginFrame" style="width:450px; height:200px; background-image:url(Images/LoginBackgroundNew.jpg);">
                <table id="LoginTitle" width="440px">
                    <tr>
                        <td align="left">
                            <dx:ASPxLabel ID="lblTitle" runat="server" Text="LOGIN - PT. AUTOCOMP SYSTEMS INDONESIA" 
                                Font-Bold="true" ForeColor="White"></dx:ASPxLabel>
                        </td>
                    </tr>
                </table>
                <table id="loginBody" width="440px">
                    <tr>
                        <td class="style1">                            
                        </td>
                        <td>                            
                        </td>
                    </tr>
                    <tr>
                        <td class="style1">                            
                        </td>
                        <td>                            
                        </td>
                    </tr>
                    <tr>
                        <td class="style1">                            
                        </td>
                        <td rowspan="4" valign="top" align="left">        
                            <table>
                                <tr>
                                    <td align="left">
                                        <dx:ASPxLabel ID="lblUserID" runat="server" Text="USER ID"></dx:ASPxLabel>
                                    </td>
                                    <td>
                                        <dx:ASPxTextBox ID="txtUserID" runat="server" Text="" Width="170px" onkeypress="return validChar(event)" >
                                            <ClientSideEvents KeyDown="function(s, e) {	lblErrMsg.SetText('');}" />
                                        </dx:ASPxTextBox>
                                    </td>
                                </tr>
                                <tr>
                                    <td align="left">
                                        <dx:ASPxLabel ID="lblPassword" runat="server" Text="PASSWORD" ></dx:ASPxLabel>
                                    </td>
                                    <td>
                                        <dx:ASPxTextBox ID="txtPassword" runat="server" Text="" Width="170px" Password="true" onkeypress="return validChar(event)">
                                            <ClientSideEvents KeyDown="function(s, e) {	lblErrMsg.SetText('');}" />
                                        </dx:ASPxTextBox>
                                    </td>
                                </tr>
                                <tr>
                                    <td colspan="2" align="right">
                                        <dx:ASPxButton ID="btnLogin" runat="server" Text="LOGIN" Width="80px"></dx:ASPxButton>
                                    </td>
                                </tr>
                            </table>
                        </td>
                    </tr>
                    <tr>
                        <td class="style1">                            
                        </td>
                    </tr>
                    <tr>
                        <td class="style1">                            
                        </td>
                    </tr>
                    <tr>
                        <td class="style1">                            
                        </td>
                    </tr>
                    <tr>
                        <td class="style1">                            
                        </td>
                        <td>                            
                        </td>
                    </tr>
                </table>
            </div>            
        </center>
    </div>

    <div>
        <center>
            <br />
            <dx:ASPxLabel ID="lblErrMsg" runat="server" Text="" Font-Italic="true" ClientInstanceName="lblErrMsg"
                ForeColor="Black"></dx:ASPxLabel>
        </center>
    </div>
    </form>
</body>
</html>
