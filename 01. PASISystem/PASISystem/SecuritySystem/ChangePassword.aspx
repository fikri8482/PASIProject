<%@ Page Title="" Language="vb" AutoEventWireup="false" MasterPageFile="~/SiteContent.Master" CodeBehind="ChangePassword.aspx.vb" Inherits="PASISystem.ChangePassword" %>
<%@ Register assembly="DevExpress.Web.v14.1, Version=14.1.6.0, Culture=neutral, PublicKeyToken=b88d1754d700e49a" namespace="DevExpress.Web.ASPxEditors" tagprefix="dx" %>
<%@ Register Assembly="DevExpress.Web.v14.1, Version=14.1.6.0, Culture=neutral, PublicKeyToken=b88d1754d700e49a" namespace="DevExpress.Web.ASPxGlobalEvents" TagPrefix="dx" %>
<%@ Register assembly="DevExpress.Web.v14.1, Version=14.1.6.0, Culture=neutral, PublicKeyToken=b88d1754d700e49a" namespace="DevExpress.Web.ASPxCallback" tagprefix="dx" %>
<asp:Content ID="Content1" ContentPlaceHolderID="head" runat="server">
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
        height = height - (height * 36 / 100)
        grid.SetHeight(height);
    }
    </script>
</asp:Content>
<asp:Content ID="Content2" ContentPlaceHolderID="Content" runat="server">
    <table style="border-width: thin; border-style: inset hidden ridge hidden; border-color:#9598A1; width:100%; height: 15px;">
        <tr>
            <td>
                <!-- MESSAGE AREA #C0C0C0 -->
                <table id="tblMsg" style="width:100%; height: 15px;">
                    <tr>
                        <td align="left" valign="middle" style="height:15px;">
                            <dx:ASPxLabel ID="lblErrMsg" runat="server" Text="" Font-Names="Tahoma" 
                                ClientInstanceName="lblErrMsg" Font-Italic="True" Font-Bold="True" 
                                Font-Size="8pt">
                            </dx:ASPxLabel>                            
                        </td>
                    </tr>
                </table>
            </td>
        </tr>
    </table>
    <div style="height:5px;"></div>
    <table style="border-width: 1pt thin thin thin; border-style: ridge; border-color:#9598A1; width:100%;">
        <tr>
            <td width="100px">
                &nbsp;</td>
            <td width="250px">
                &nbsp;</td>
            <td width="100px">
                &nbsp;</td>
        </tr>
        <tr>
            <td width="100px">
                &nbsp;</td>
            <td width="250px">
                &nbsp;</td>
            <td width="100px">
                &nbsp;</td>
        </tr>
        <tr>
            <td width="100px">
                &nbsp;</td>
            <td width="250px">
                &nbsp;</td>
            <td width="100px">
                &nbsp;</td>
        </tr>
        <tr>
            <td width="100px">
                &nbsp;</td>
            <td width="250px">
                &nbsp;</td>
            <td width="100px">
                &nbsp;</td>
        </tr>
        <tr>
            <td width="100px">
                &nbsp;</td>
            <td width="250px">
                &nbsp;</td>
            <td width="100px">
                &nbsp;</td>
        </tr>
        <tr>
            <td width="100px">
                &nbsp;</td>
            <td width="250px">
                &nbsp;</td>
            <td width="100px">
                &nbsp;</td>
        </tr>
        <tr>
            <td width="100px">
                &nbsp;</td>
            <td>
                <table style="width:100%;">
                    <tr>
                        <td align="left" width="150px">
                            &nbsp;</td>
                        <td align="left" width="200px">
                            <dx:ASPxLabel ID="ASPxLabel1" runat="server" Font-Names="Tahoma" 
                                Font-Size="8pt" Text="CURRENT PASSWORD">
                            </dx:ASPxLabel>
                        </td>
                        <td align="left">
                            <dx:ASPxTextBox ID="txtCurrentPassword" runat="server" Font-Names="Tahoma" 
                                Width="200px" ClientInstanceName="txtCurrentPassword" Password="True">
                            </dx:ASPxTextBox>
                        </td>
                    </tr>
                    <tr>
                        <td align="left" width="150px">
                            &nbsp;</td>
                        <td align="left" width="200px">
                            <dx:ASPxLabel ID="ASPxLabel2" runat="server" Font-Names="Tahoma" 
                                Font-Size="8pt" Text="NEW PASSWORD">
                            </dx:ASPxLabel>
                        </td>
                        <td align="left">
                            <dx:ASPxTextBox ID="txtNewPassword" runat="server" Font-Names="Tahoma" 
                                Width="200px" ClientInstanceName="txtNewPassword" Password="True">
                            </dx:ASPxTextBox>
                        </td>
                    </tr>
                    <tr>
                        <td align="left" width="150px">
                            &nbsp;</td>
                        <td align="left" width="200px">
                            <dx:ASPxLabel ID="ASPxLabel3" runat="server" Font-Names="Tahoma" 
                                Font-Size="8pt" Text="CONFIRM NEW PASSWORD">
                            </dx:ASPxLabel>
                        </td>
                        <td align="left">
                            <dx:ASPxTextBox ID="txtConfirmPassword" runat="server" Font-Names="Tahoma" 
                                Width="200px" ClientInstanceName="txtConfirmPassword" Password="True">
                            </dx:ASPxTextBox>
                        </td>
                    </tr>
                </table>
            </td>
            <td width="100px">
                <dx:ASPxCallback ID="cbProgress" runat="server" ClientInstanceName="cbProgress">
                <ClientSideEvents CallbackComplete="function(s, e) {
                lblErrMsg.GetMainElement().style.color = 'Red';
                lblErrMsg.SetText(s.cpMessage);
                    if (lblErrMsg.GetText() == '[1002] Data Updated Successfully!') {
                        txtCurrentPassword.SetText('');
                        txtNewPassword.SetText('');
                        txtConfirmPassword.SetText('');	
                    }        
                }" EndCallback="function(s, e) {
	var pMsg = s.cpMessage;        
    if (pMsg.substring(1,5) == '6011' || pMsg.substring(1,5) == '1002' || pMsg.substring(1,5) == '7001') {
        lblErrMsg.GetMainElement().style.color = 'Blue';
    } else {
        lblErrMsg.GetMainElement().style.color = 'Red';
    }

    lblErrMsg.SetText(pMsg);
}" />
                                            
                </dx:ASPxCallback>
            </td>
        </tr>
        <tr>
            <td height="270px">
                &nbsp;</td>
            <td width="200px">
                &nbsp;</td>
            <td width="100px">
                &nbsp;</td>
        </tr>
    </table>

    <%--<table style="width: 100%;">
        <tr>
            <td>
                <!-- MESSAGE AREA #C0C0C0 -->
                <table id="tblMsg" style="border: thin ridge #9598A1; width:100%;">
                    <tr>
                        <td align="center" valign="middle" style="height:25px;">
                            <dx:ASPxLabel ID="lblErrMsg" runat="server" Text="" Font-Names="Tahoma" 
                                ClientInstanceName="lblErrMsg" Font-Italic="True" Font-Bold="true" Font-Size="8pt">
                            </dx:ASPxLabel>                            
                        </td>
                    </tr>
                </table>
            </td>
        </tr>
    </table>--%>
    <div style="height:8px;"></div>
    <table style="width:100%;">
        <tr>
            <td align="left">
                <dx:ASPxButton ID="btnSubMenu" runat="server" Text="SUB MENU" Font-Names="Tahoma" 
                    Width="90px" Font-Size="8pt" AutoPostBack="False" UseSubmitBehavior="false" >
                </dx:ASPxButton>
            </td>
            <td align="right" style="width:85px;">
                <dx:ASPxButton ID="btnClear" runat="server" Text="Clear" AutoPostBack="False" 
                    Width="80px" Font-Names="Tahoma" Font-Size="8pt" UseSubmitBehavior="false">
                    <ClientSideEvents Click="function(s, e) {
	txtCurrentPassword.SetText('');
	txtNewPassword.SetText('');
	txtConfirmPassword.SetText('');
    lblErrMsg.SetText('');    
}" 
/>
                </dx:ASPxButton>
            </td>
            <td align="right" style="width:80px;">
                <dx:ASPxButton ID="btnSubmit" runat="server" Text="SAVE"
                    Width="80px" Font-Names="Tahoma" Font-Size="8pt" AutoPostBack="False"
                    UseSubmitBehavior="False">
                    <ClientSideEvents Click="function(s, e) {
		if (txtCurrentPassword.GetText() == ''){
        lblErrMsg.GetMainElement().style.color = 'Red';
		lblErrMsg.SetText('[6012] Please input Current Password first!');
		e.processOnServer = false;
        return;
	    }
		if (txtNewPassword.GetText() == '') {
        lblErrMsg.GetMainElement().style.color = 'Red';
        lblErrMsg.SetText('[6012] Please input New Password first!');
                e.ProcessOnServer = false;
                return;
        }
		if (txtConfirmPassword.GetText() == '') {
        lblErrMsg.GetMainElement().style.color = 'Red';
        lblErrMsg.SetText('[6012] Please input Confirm Password first!');
                e.ProcessOnServer = false;
                return;
        }
		cbProgress.PerformCallback();	
        
        
}" />
                </dx:ASPxButton>
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
