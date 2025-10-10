<%@ Page Title="" Language="vb" AutoEventWireup="false" MasterPageFile="~/SiteContent.Master" CodeBehind="EmailSetting.aspx.vb" Inherits="PASISystem.EmailSetting" %>

<%@ Register Assembly="DevExpress.Web.v14.1, Version=14.1.6.0, Culture=neutral, PublicKeyToken=b88d1754d700e49a" Namespace="DevExpress.Web.ASPxCallback" TagPrefix="dx" %>
<%@ Register Assembly="DevExpress.Web.v14.1, Version=14.1.6.0, Culture=neutral, PublicKeyToken=b88d1754d700e49a" Namespace="DevExpress.Web.ASPxGridView.Export" TagPrefix="dx" %>
<%@ Register Assembly="DevExpress.Web.v14.1, Version=14.1.6.0, Culture=neutral, PublicKeyToken=b88d1754d700e49a" Namespace="DevExpress.Web.ASPxGridView" TagPrefix="dx" %>
<%@ Register assembly="DevExpress.Web.v14.1, Version=14.1.6.0, Culture=neutral, PublicKeyToken=b88d1754d700e49a" namespace="DevExpress.Web.ASPxEditors" tagprefix="dx" %>
<%@ Register Assembly="DevExpress.Web.v14.1, Version=14.1.6.0, Culture=neutral, PublicKeyToken=b88d1754d700e49a" Namespace="DevExpress.Web.ASPxPanel" TagPrefix="dx" %>
<%@ Register Assembly="DevExpress.Web.v14.1, Version=14.1.6.0, Culture=neutral, PublicKeyToken=b88d1754d700e49a" Namespace="DevExpress.Web.ASPxPopupControl" TagPrefix="dx" %>
<%@ Register Assembly="DevExpress.Web.v14.1, Version=14.1.6.0, Culture=neutral, PublicKeyToken=b88d1754d700e49a" namespace="DevExpress.Web.ASPxGlobalEvents" TagPrefix="dx" %>
<%@ Register assembly="DevExpress.Web.v14.1, Version=14.1.6.0, Culture=neutral, PublicKeyToken=b88d1754d700e49a" namespace="DevExpress.Web.ASPxHiddenField" tagprefix="dx1" %>
<%@ Register assembly="DevExpress.Web.v14.1, Version=14.1.6.0, Culture=neutral, PublicKeyToken=b88d1754d700e49a" namespace="DevExpress.Web.ASPxFileManager" tagprefix="dx" %>
<%@ Register assembly="DevExpress.Web.v14.1, Version=14.1.6.0, Culture=neutral, PublicKeyToken=b88d1754d700e49a" namespace="DevExpress.Web.ASPxUploadControl" tagprefix="dx" %>
<asp:Content ID="Content1" ContentPlaceHolderID="head" runat="server">
<style type="text/css">
.dxeHLC, .dxeHC, .dxeHFC
{
display: none;
}
    .style1
    {
        width: 199px;
    }
    .style2
    {
        width: 319px;
    }
    .style4
    {
        width: 52px;
    }
    .style5
    {
        width: 51px;
    }
    .style7
    {
        width: 153px;
        text-align: left;
    }
    .style9
    {
        width: 130px;
    }
    .style10
    {
        width: 42px;
        text-align: left;
    }
    .style11
    {
        width: 199px;
        height: 32px;
    }
    .style12
    {
        width: 319px;
        height: 32px;
    }
    .style13
    {
        height: 32px;
    }
</style> 

<script language="javascript" type="text/javascript">


    function singlequote(e) {
        var unicode = e.charCode ? e.charCode : e.keyCode
        if (unicode == 39) {
            return false //disable key press
        }
    }

    function numbersonly(e) {
        var unicode = e.charCode ? e.charCode : e.keyCode
        if (unicode != 8) { //if the key isn't the backspace key (which we should allow)
            if (unicode < 45 || unicode > 57) //if not a number
                return false //disable key press
        }
    }

    function clear() {
        txtEmailAddress.SetText('');
        txtUserName.SetText('');
        txtPassword.SetText('');
        txtPort.SetText('');
        txtPop3.SetText('');
        txtAttachmentSave.SetText('');
        txtAttachmentBackup.SetText('');
        txtSchedule.SetText('');
        txtUserNameSMTP.SetText('');
        txtPasswordSMTP.SetText('');
        txtSMTP.SetText('');
        txtPortSMTP.SetText('');
        txtTemplate.SetText('');
        txtResult.SetText('');
        txtSendExcel.SetText('');
        txtPOInterval.SetText('');
        txtPO.SetText('');
        txtPORevision.SetText('');
        txtPORevisionInterval.SetText('');
        txtKanban.SetText('');
        txtKanbanInterval.SetText('');
    }

    function validasubmit() {
        lblInfo.GetMainElement().style.color = 'Red';
        
        if (txtEmailAddress.GetText() == "") {
            lblInfo.SetText("[6011] Please Input Email Address first!");
            txtEmailAddress.Focus();
            e.ProcessOnServer = false;
            return false;
        }
        
        if (txtUserName.GetText() == "") {
            lblInfo.SetText("[6011] Please Input Username first!");
            txtUserName.Focus();
            e.ProcessOnServer = false;
            return false;
        }
        
        if (txtPassword.GetText() == "") {
            lblInfo.SetText("[6011] Please Input Password first!");
            txtPassword.Focus();
            e.ProcessOnServer = false;
            return false;
        }
        
        if (txtPop3.GetText() == "") {
            lblInfo.SetText("[6011] Please Input POP3 first!");
            txtPop3.Focus();
            e.ProcessOnServer = false;
            return false;
        }
        
        if (txtPort.GetText() == "") {
            lblInfo.SetText("[6011] Please Input Port first!");
            txtPort.Focus();
            e.ProcessOnServer = false;
            return false;
        }
        
        if (txtAttachmentSave.GetText() == "") {
            lblInfo.SetText("[6011] Please Attachment Folder Path first!");
            txtAttachmentSave.Focus();
            e.ProcessOnServer = false;
            return false;
        }
       
        if (txtSchedule.GetText() == "") {
            lblInfo.SetText("[6011] Please Schedule Interval first!");
            txtSchedule.Focus();
            e.ProcessOnServer = false;
            return false;
        }
        
        if (txtAttachmentBackup.GetText() == "") {
            lblInfo.SetText("[6011] Please Attachment Backup Folder Path first!");
            txtAttachmentBackup.Focus();
            e.ProcessOnServer = false;
            return false;
        }
        
        if (txtSMTP.GetText() == "") {
            lblInfo.SetText("[6011] Please Input SMTP first!");
            txtSMTP.Focus();
            e.ProcessOnServer = false;
            return false;
        }
        
        if (txtPortSMTP.GetText() == "") {
            lblInfo.SetText("[6011] Please Input Port SMTP first!");
            txtPortSMTP.Focus();
            e.ProcessOnServer = false;
            return false;
        }

        if (txtTemplate.GetText() == "") {
            lblInfo.SetText("[6011] Please Input Template Attachment Folder first!");
            txtTemplate.Focus();
            e.ProcessOnServer = false;
            return false;
        }

        if (txtResult.GetText() == "") {
            lblInfo.SetText("[6011] Please Input Result Attachment Folder first!");
            txtResult.Focus();
            e.ProcessOnServer = false;
            return false;
        }

        if (txtPOInterval.GetText() == "") {
            lblInfo.SetText("[6011] Please Input Po Interval first!");
            txtPOInterval.Focus();
            e.ProcessOnServer = false;
            return false;
        }

        if (txtPO.GetText() == "") {
            lblInfo.SetText("[6011] Please Input PO Approval Date first!");
            txtPO.Focus();
            e.ProcessOnServer = false;
            return false;
        }       
    }

    function up_Insert() {
  
        var pIsUpdate = '';
        var pEmailAddress = txtEmailAddress.GetText();
        var pUserName = txtUserName.GetText();
        var pPassword = txtPassword.GetText();
        var pPOP3 = txtPop3.GetText();
        var pPort = txtPort.GetText();
        var pAttachmentSave = txtAttachmentSave.GetText();
        var pAttachmentBackup = txtAttachmentBackup.GetText();
        var pInterval = txtSchedule.GetText();
        var pusernameSMTP = txtUserNameSMTP.GetText();
        var pPasswordSMTP = txtPasswordSMTP.GetText();
        var pSMTP = txtSMTP.GetText();
        var pPortSMTP = txtPortSMTP.GetText();
        var pTemplate = txtTemplate.GetText();
        var pResult = txtResult.GetText();
        var pSendExcel = txtSendExcel.GetText();
        var pPOApprovalDate = txtPO.GetText();
        var pIntervalPOApproval = txtPOInterval.GetText();
        var pPORevisionApprovalDate = txtPORevision.GetText();
        var pIntervalPORevisionApproval = txtPORevisionInterval.GetText();
        var pKanbanApprovalHour = txtKanban.GetText();
        var pIntervalKanbanApproval = txtKanbanInterval.GetText();


        AffiliateSubmit.PerformCallback('save|' + pIsUpdate + '|' + pEmailAddress + '|' + pUserName + '|' + pPassword + '|' + pPOP3 + '|' + pPort + '|' + pAttachmentSave + '|' + pAttachmentBackup + '|' + pInterval + '|' + pusernameSMTP + '|' + pPasswordSMTP + '|' + pSMTP + '|' + pPortSMTP + '|' + pTemplate + '|' + pResult + '|' + pSendExcel + '|' + pPOApprovalDate + '|' + pIntervalPOApproval + '|' + pPORevisionApprovalDate + '|' + pIntervalPORevisionApproval + '|' + pKanbanApprovalHour + '|' + pIntervalKanbanApproval);
        
    }

    function up_delete() {

        var msg = confirm('Are you sure want to delete this data ?');
        if (msg == false) {
            e.processOnServer = false;
            return;
        }

        var pEmailAddress = txtEmailAddress.GetText();
        var pUserName = txtUserName.GetText();
        var pPassword = txtPassword.GetText();
        var pPOP3 = txtPop3.GetText();
        var pPort = txtPort.GetText();
        var pAttachmentSave = txtAttachmentSave.GetText();
        var pAttachmentBackup = txtAttachmentBackup.GetText();
        var pInterval = txtSchedule.GetText();
        var pusernameSMTP = txtUserNameSMTP.GetText();
        var pPasswordSMTP = txtPasswordSMTP.GetText();
        var pSMTP = txtSMTP.GetText();
        var pPortSMTP = txtPortSMTP.GetText();
        var pTemplate = txtTemplate.GetText();
        var pResult = txtResult.GetText();
        var pSendExcel = txtSendExcel.GetText();
        var pPOApprovalDate = txtPO.GetText();
        var pIntervalPOApproval = txtPOInterval.GetText();
        var pPORevisionApprovalDate = txtPORevision.GetText();
        var pIntervalPORevisionApproval = txtPORevisionInterval.GetText();
        var pKanbanApprovalHour = txtKanban.GetText();
        var pIntervalKanbanApproval = txtKanbanInterval.GetText();
        var ptype = cbotype.GetText();

        AffiliateSubmit.PerformCallback('delete|' + pEmailAddress + '|' + pUserName + '|' + pPassword + '|' + pPOP3 + '|' + pPort + '|' + pAttachmentSave + '|' + pAttachmentBackup + '|' + pInterval + '|' + pusernameSMTP + '|' + pPasswordSMTP + '|' + pSMTP + '|' + pPortSMTP + '|' + pTemplate + '|' + pResult + '|' + pSendExcel + '|' + pPOApprovalDate + '|' + pIntervalPOApproval + '|' + pPORevisionApprovalDate + '|' + pIntervalPORevisionApproval + '|' + pKanbanApprovalHour + '|' + pIntervalKanbanApproval + '|' + ptype);
      
    }

    function ShowFilename() {
        var file = filemanager.GetSelectedFile();
        if (file) {
            var folder = filemanager.GetCurrentFolderPath();
            var relativePath = "~\\" + folder + "\\" + file.name;
            txbFilename.SetText(relativePath);
            popup.Hide();
        }
    }


    </script>
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
            height = height - (height * 52 / 100)
            
        }
    </script>

</asp:Content>

<asp:Content ID="Content2" ContentPlaceHolderID="Content" runat="server">
    <table style="border-width: thin; border-style: inset hidden ridge hidden; border-color:#9598A1; width:100%; height: 15px;">
        <tr>
            <td colspan="8" height="15">
                <%--error message--%>
                <table id="info" style="width:100%; height: 15px;">
                    <tr>
                        <td align="left" valign="middle" style="height:15px;">
                            <dx:ASPxLabel ID="lblInfo" runat="server" Font-Names="Tahoma"
                                ClientInstanceName="lblInfo" Font-Bold="True" Font-Italic="True" 
                                Font-Size="8pt" >
                            </dx:ASPxLabel>
                        </td>
                    </tr>         
                </table>
            </td>            
        </tr>
    </table>

    <div style="height:5px;"></div> 

    <table style="border-width: 1pt thin thin thin; border-style: ridge; border-color:#9598A1; width:100%; height: 450px;">
        <tr>
            <td>
                <table style="width:100%;">
                    <tr>
                        <td align="left" class="style1">
                            <dx:ASPxLabel ID="ASPxLabel34" runat="server" Font-Names="Tahoma" 
                                Font-Size="8pt" Text="PO TYPE">
                            </dx:ASPxLabel>
                        </td>
                        <td align="left" class="style2">
                            <dx:ASPxComboBox ID="cbotype" runat="server" ClientInstanceName="cbotype" 
                                Height="25px">
                                <ClientSideEvents SelectedIndexChanged="function(s, e) {                                    

if (cbotype.GetValue() == 'DOMESTIC')
		{   
            txtPO.SetVisible(true);
            txtPORevision.SetVisible(true);
            txtKanban.SetVisible(true);
            txtPOInterval.SetVisible(true);
            txtPORevisionInterval.SetVisible(true);
            txtKanbanInterval.SetVisible(true);
            ASPxLabel22.SetVisible(true);
            ASPxLabel25.SetVisible(true);
            ASPxLabel28.SetVisible(true);
            ASPxLabel23.SetVisible(true);
            ASPxLabel26.SetVisible(true);
            ASPxLabel29.SetVisible(true);
            ASPxLabel24.SetVisible(true);
            ASPxLabel27.SetVisible(true);
            ASPxLabel30.SetVisible(true);
	} else {
            ASPxLabel25.SetVisible(false);
            ASPxLabel28.SetVisible(false);
            ASPxLabel26.SetVisible(false);
            ASPxLabel29.SetVisible(false);
            ASPxLabel27.SetVisible(false);
            ASPxLabel30.SetVisible(false);
            txtPORevision.SetVisible(false);
            txtKanban.SetVisible(false);
            txtPORevisionInterval.SetVisible(false);
            txtKanbanInterval.SetVisible(false);
	}                                     

cbBind.PerformCallback(cbotype.GetText);

}" />
                            </dx:ASPxComboBox>
                        </td>
                        <td align="left">
                            &nbsp;</td>
                    </tr>
                    <tr>
                        <td align="left" class="style1">
                            <dx:ASPxLabel ID="ASPxLabel10" runat="server" Font-Names="Tahoma" 
                                Font-Size="8pt" Text="EMAIL ADDRESS">
                            </dx:ASPxLabel>
                        </td>
                        <td align="left" class="style2">
                            <dx:ASPxTextBox ID="txtEmailAddress" runat="server" Width="400px" 
                                ClientInstanceName="txtEmailAddress" Font-Names="Tahoma" 
                                MaxLength="100" Height="25px" >
                            </dx:ASPxTextBox>
                        </td>
                        <td align="left">
                            &nbsp;</td>
                    </tr>
                    <tr>
                        <td align="left" class="style1">
                            <dx:ASPxLabel ID="ASPxLabel11" runat="server" Font-Names="Tahoma" 
                                Font-Size="8pt" Text="USER NAME">
                            </dx:ASPxLabel>
                        </td>
                        <td align="left" class="style2">
                            <dx:ASPxTextBox ID="txtUserName" runat="server" Width="400px" 
                                ClientInstanceName="txtUserName" Font-Names="Tahoma" 
                                MaxLength="100" Height="25px">
                            </dx:ASPxTextBox>
                        </td>
                        <td align="left">
                            &nbsp;</td>
                    </tr>
                    <tr>
                        <td align="left" class="style1">
                            <dx:ASPxLabel ID="ASPxLabel12" runat="server" Font-Names="Tahoma" 
                                Font-Size="8pt" Text="PASSWORD">
                            </dx:ASPxLabel>
                        </td>
                        <td align="left" class="style2">
                            <dx:ASPxTextBox ID="txtPassword" runat="server" Width="400px" 
                                ClientInstanceName="txtPassword" Font-Names="Tahoma" 
                                MaxLength="20" Height="25px" 
                                Password="True">
                                <ClientSideEvents Init="function(s, e) {
	s.SetValue(s.cp_myPassword);
}" />
                            </dx:ASPxTextBox>
                        </td>
                        <td align="left">
                            &nbsp;</td>
                    </tr>
                    <tr>
                        <td align="left" class="style1">
                            <dx:ASPxLabel ID="ASPxLabel14" runat="server" Font-Names="Tahoma" 
                                Font-Size="8pt" Text="PORT">
                            </dx:ASPxLabel>
                        </td>
                        <td align="left" class="style2">
                            <dx:ASPxTextBox ID="txtPort" runat="server" Width="100px" 
                                ClientInstanceName="txtPort" Font-Names="Tahoma" 
                                MaxLength="4" onkeypress="return numbersonly(event)" Height="25px">
                                <ClientSideEvents LostFocus="function(s, e) { lblInfo.SetText(&quot;&quot;); }" />
                            </dx:ASPxTextBox>
                        </td>
                        <td align="left">
                            &nbsp;</td>
                    </tr>
                    <tr>
                        <td align="left" class="style1">
                            <dx:ASPxLabel ID="ASPxLabel13" runat="server" Font-Names="Tahoma" 
                                Font-Size="8pt" Text="POP 3">
                            </dx:ASPxLabel>
                        </td>
                        <td align="left" class="style2">
                            <dx:ASPxTextBox ID="txtPop3" runat="server" Width="400px" 
                                ClientInstanceName="txtPop3" Font-Names="Tahoma" 
                                MaxLength="100" Height="25px">
                            </dx:ASPxTextBox>
                        </td>
                        <td align="left">
                            &nbsp;</td>
                    </tr>
                    <tr>
                        <td align="left" class="style1">
                            <dx:ASPxLabel ID="ASPxLabel15" runat="server" Font-Names="Tahoma" 
                                Font-Size="8pt" Text="GET ATTACHMENT SAVE FOLDER">
                            </dx:ASPxLabel>
                        </td>
                        <td align="left" class="style2">
                            <dx:ASPxTextBox ID="txtAttachmentSave" runat="server" Width="400px" 
                                ClientInstanceName="txtAttachmentSave" Font-Names="Tahoma" 
                                MaxLength="100" Height="25px" Text="D:\">
                            </dx:ASPxTextBox>
                        </td>
                        <td align="left">
                            &nbsp;</td>
                    </tr>
                    <tr>
                        <td align="left" class="style1">
                            <dx:ASPxLabel ID="ASPxLabel17" runat="server" Font-Names="Tahoma" 
                                Font-Size="8pt" Text="GET ATTACHMENT BACKUP">
                            </dx:ASPxLabel>
                        </td>
                        <td align="left" class="style2">
                            <dx:ASPxTextBox ID="txtAttachmentBackup" runat="server" Width="400px" 
                                ClientInstanceName="txtAttachmentBackup" Font-Names="Tahoma" 
                                MaxLength="100" Height="25px" Text="D:\">
                            
                            </dx:ASPxTextBox>
                        </td>
                        <td align="left">
                            &nbsp;</td>
                    </tr>
                    <tr>
                        <td align="left" class="style11">
                            <dx:ASPxLabel ID="ASPxLabel16" runat="server" Font-Names="Tahoma" 
                                Font-Size="8pt" Text="GET EMAIL SCHEDULE (SECONDS)" >
                            </dx:ASPxLabel>
                        </td>
                        <td align="left" class="style12">
                            <dx:ASPxTextBox ID="txtSchedule" runat="server" Width="100px" 
                                ClientInstanceName="txtSchedule" onkeypress="return numbersonly(event)" Font-Names="Tahoma" 
                                MaxLength="20"  Height="25px">
                            </dx:ASPxTextBox>
                        </td>
                        <td align="left" class="style13">
                            </td>
                    </tr>
                    <tr>
                        <td align="left" class="style1">
                            <dx:ASPxLabel ID="ASPxLabel20" runat="server" Font-Names="Tahoma" 
                                Font-Size="8pt" Text="USER NAME SMTP">
                            </dx:ASPxLabel>
                        </td>
                        <td align="left" class="style2">
                            <dx:ASPxTextBox ID="txtUserNameSMTP" runat="server" Width="400px" 
                                ClientInstanceName="txtUserNameSMTP" Font-Names="Tahoma" 
                                MaxLength="100" Height="25px">
                            </dx:ASPxTextBox>
                        </td>
                        <td align="left">
                            &nbsp;</td>
                    </tr>
                    <tr>
                        <td align="left" class="style1">
                            <dx:ASPxLabel ID="ASPxLabel21" runat="server" Font-Names="Tahoma" 
                                Font-Size="8pt" Text="PASSWORD SMTP">
                            </dx:ASPxLabel>
                        </td>
                        <td align="left" class="style2">
                            <dx:ASPxTextBox ID="txtPasswordSMTP" runat="server" Width="400px" 
                                ClientInstanceName="txtPasswordSMTP" Font-Names="Tahoma" 
                                MaxLength="20" Height="25px" Password="True">
                                <ClientSideEvents Init="function(s, e) {
	s.SetValue(s.cp_myPasswords);
}" />
                            </dx:ASPxTextBox>
                        </td>
                        <td align="left">
                            &nbsp;</td>
                    </tr>
                    <tr>
                        <td align="left" class="style1">
                            <dx:ASPxLabel ID="ASPxLabel19" runat="server" Font-Names="Tahoma" 
                                Font-Size="8pt" Text="SMTP">
                            </dx:ASPxLabel>
                        </td>
                        <td align="left" class="style2">
                            <dx:ASPxTextBox ID="txtSMTP" runat="server" Width="400px" 
                                ClientInstanceName="txtSMTP" Font-Names="Tahoma" 
                                MaxLength="25" Height="25px">
                            </dx:ASPxTextBox>
                        </td>
                        <td align="left">
                            &nbsp;</td>
                    </tr>
                    <tr>
                        <td align="left" class="style1">
                            <dx:ASPxLabel ID="ASPxLabel18" runat="server" Font-Names="Tahoma" 
                                Font-Size="8pt" Text="PORT SMTP">
                            </dx:ASPxLabel>
                        </td>
                        <td align="left" class="style2">
                            <dx:ASPxTextBox ID="txtPortSMTP" runat="server" Width="100px" 
                                ClientInstanceName="txtPortSMTP" Font-Names="Tahoma" 
                                MaxLength="4" onkeypress="return numbersonly(event)" Height="25px">
                                <ClientSideEvents LostFocus="function(s, e) { lblInfo.SetText(&quot;&quot;); }" />
                            </dx:ASPxTextBox>
                        </td>
                        <td align="left">
                            &nbsp;</td>
                    </tr>
                    <tr>
                        <td align="left" class="style1">
                            <dx:ASPxLabel ID="ASPxLabel31" runat="server" Font-Names="Tahoma" 
                                Font-Size="8pt" Text="TEMPLATE ATTACHMENT FOLDER">
                            </dx:ASPxLabel>
                        </td>
                        <td align="left" class="style2">
                            <dx:ASPxTextBox ID="txtTemplate" runat="server" Width="400px" 
                                ClientInstanceName="txtTemplate" Font-Names="Tahoma" 
                                MaxLength="100" Height="25px" Text="D:\">
                            </dx:ASPxTextBox>
                        </td>
                        <td align="left">
                            &nbsp;</td>
                    </tr>
                    <tr>
                        <td align="left" class="style1">
                            <dx:ASPxLabel ID="ASPxLabel32" runat="server" Font-Names="Tahoma" 
                                Font-Size="8pt" Text="RESULT ATTACHMENT FOLDER">
                            </dx:ASPxLabel>
                        </td>
                        <td align="left" class="style2">
                            <dx:ASPxTextBox ID="txtResult" runat="server" Width="400px" 
                                ClientInstanceName="txtResult" Font-Names="Tahoma" 
                                MaxLength="100" Height="25px" Text="D:\">
                            </dx:ASPxTextBox>
                        </td>
                        <td align="left">
                            &nbsp;</td>
                    </tr>
                    <tr>
                        <td align="left" class="style1">
                            <dx:ASPxLabel ID="ASPxLabel33" runat="server" Font-Names="Tahoma" 
                                Font-Size="8pt" Text="INTERVAL SEND EXCEL (SECONDS)">
                            </dx:ASPxLabel>
                        </td>
                        <td align="left" class="style2">
                            <dx:ASPxTextBox ID="txtSendExcel" runat="server" Width="100px" 
                                ClientInstanceName="txtSendExcel" onkeypress="return numbersonly(event)" Font-Names="Tahoma" 
                                MaxLength="20"  Height="25px">
                            </dx:ASPxTextBox>
                        </td>
                        <td align="left">
                            &nbsp;</td>
                    </tr>
                    </table>

                <table style="width: 100%;">
                    <tr>
                        <td class="style7">
                            <dx:ASPxLabel ID="ASPxLabel22" runat="server" Font-Names="Tahoma" 
                                Font-Size="8pt" Text="AUTO APPROVE PO" style="text-align: left" 
                                ClientInstanceName="ASPxLabel22">
                            </dx:ASPxLabel>
                        </td>
                        <td class="style10">
                            &nbsp;</td>
                        <td class="style4">
                            <dx:ASPxTextBox ID="txtPO" runat="server" Width="50px" 
                                ClientInstanceName="txtPO" Font-Names="Tahoma" 
                                MaxLength="3" Height="25px" onkeypress="return numbersonly(event)">                                
                            </dx:ASPxTextBox>
                        </td>
                        <td class="style4">
                            &nbsp;
                            <dx:ASPxLabel ID="ASPxLabel23" runat="server" Font-Names="Tahoma" 
                                Font-Size="8pt" Text="DAYS" ClientInstanceName="ASPxLabel23">
                            </dx:ASPxLabel>
                        </td>
                        <td class="style4">
                            &nbsp;</td>
                        <td class="style5">
                            <dx:ASPxTextBox ID="txtPOInterval" runat="server" Width="50px" 
                                ClientInstanceName="txtPOInterval" Font-Names="Tahoma" 
                                MaxLength="3" onkeypress="return numbersonly(event)" Height="25px">
                                
                            </dx:ASPxTextBox>
                        </td>
                        <td class="style9">
                            <dx:ASPxLabel ID="ASPxLabel24" runat="server" Text="INTERVAL (SECONDS)" 
                                ClientInstanceName="ASPxLabel24">
                            </dx:ASPxLabel>
                        </td>
                        <td>
                            &nbsp;</td>
                    </tr>
                    <tr>
                        <td class="style7">
                            <dx:ASPxLabel ID="ASPxLabel25" runat="server" Font-Names="Tahoma" 
                                Font-Size="8pt" Text="AUTO APPROVE PO REVISION" style="text-align: left" 
                                ClientInstanceName="ASPxLabel25">
                            </dx:ASPxLabel>
                        </td>
                        <td class="style10">
                            &nbsp;</td>
                        <td class="style4">
                            <dx:ASPxTextBox ID="txtPORevision" runat="server" Width="50px" 
                                ClientInstanceName="txtPORevision" Font-Names="Tahoma" 
                                MaxLength="3" onkeypress="return numbersonly(event)" Height="25px">
                                
                            </dx:ASPxTextBox>
                        </td>
                        <td class="style4">
                            &nbsp;
                            <dx:ASPxLabel ID="ASPxLabel26" runat="server" Font-Names="Tahoma" 
                                Font-Size="8pt" Text="DAYS" ClientInstanceName="ASPxLabel26">
                            </dx:ASPxLabel>
                        </td>
                        <td class="style4">
                            &nbsp;</td>
                        <td class="style5">
                            <dx:ASPxTextBox ID="txtPORevisionInterval" runat="server" Width="50px" 
                                ClientInstanceName="txtPORevisionInterval" Font-Names="Tahoma" 
                                MaxLength="3" onkeypress="return numbersonly(event)" Height="25px">
                                
                            </dx:ASPxTextBox>
                        </td>
                        <td class="style9">
                            <dx:ASPxLabel ID="ASPxLabel27" runat="server" Text="INTERVAL (SECONDS)" 
                                ClientInstanceName="ASPxLabel27">
                            </dx:ASPxLabel>
                        </td>
                        <td>
                            &nbsp;</td>
                    </tr>
                    <tr>
                        <td class="style7">
                            <dx:ASPxLabel ID="ASPxLabel28" runat="server" Font-Names="Tahoma" 
                                Font-Size="8pt" Text="AUTO APPROVE KANBAN" style="text-align: left" 
                                ClientInstanceName="ASPxLabel28">
                            </dx:ASPxLabel>
                        &nbsp;</td>
                        <td class="style10">
                            &nbsp;</td>
                        <td class="style4">
                            <dx:ASPxTextBox ID="txtKanban" runat="server" Width="50px" 
                                ClientInstanceName="txtKanban" Font-Names="Tahoma" 
                                MaxLength="2" Height="25px" onkeypress="return numbersonly(event)">
                               
                            </dx:ASPxTextBox>
                        </td>
                        <td class="style4">
                            &nbsp;
                            <dx:ASPxLabel ID="ASPxLabel29" runat="server" Font-Names="Tahoma" 
                                Font-Size="8pt" Text="HOURS" ClientInstanceName="ASPxLabel29">
                            </dx:ASPxLabel>
                        </td>
                        <td class="style4">
                            &nbsp;</td>
                        <td class="style5">
                            <dx:ASPxTextBox ID="txtKanbanInterval" runat="server" Width="50px" 
                                ClientInstanceName="txtKanbanInterval" Font-Names="Tahoma" 
                                MaxLength="3" onkeypress="return numbersonly(event)" Height="25px">
                                
                            </dx:ASPxTextBox>
                        </td>
                        <td class="style9">
                            <dx:ASPxLabel ID="ASPxLabel30" runat="server" Text="INTERVAL (SECONDS)" 
                                ClientInstanceName="ASPxLabel30">
                            </dx:ASPxLabel>
                        </td>
                        <td>
                            &nbsp;</td>
                    </tr>
                </table>

            </td>
        </tr>

    </table> 

    <div style="height:8px;"></div>      

    <%--Button--%> 
    <table id="button" style=" width:100%;">
        <tr>                        
            <td valign="top" align="left">
                <dx:ASPxButton ID="btnSubMenu" runat="server" Text="SUB MENU"                     
                    Font-Names="Tahoma"
                    Width="90px" Font-Size="8pt" ClientInstanceName="btnSubmenu">
                </dx:ASPxButton>   
                </td>                     
            
            <td valign="top" align="right" style="width: 50px;">                                  
            </td>
            <td valign="top" align="right" style="width: 50px;">                                   
                &nbsp;</td>
            <td align="right" style="width:80px;">                                   
                <dx:ASPxButton ID="btnDelete" runat="server" Text="DELETE"                    
                    Font-Names="Tahoma"
                    Width="90px" AutoPostBack="False" Font-Size="8pt" 
                    ClientInstanceName="btnDelete" Visible="False">
                    <ClientSideEvents Click="function(s, e) {
                  
                        validasubmit();
                     
                        up_delete();
                     
                        clear();
                      
                    }" />
                </dx:ASPxButton>
            </td>
            <td valign="top" align="right" style="width: 50px;">                                   
                <dx:ASPxButton ID="btnSubmit" runat="server" Text="SAVE"                    
                    Font-Names="Tahoma"
                    Width="90px" AutoPostBack="False" Font-Size="8pt" 
                    ClientInstanceName="btnSubmit">
                    <ClientSideEvents Click="function(s, e) {
                        validasubmit();
                        up_Insert();
                      
                    }" />
                </dx:ASPxButton>
            </td>
        </tr>
    </table>
    <div style="height:8px;"></div>  
    
    <dx:ASPxGlobalEvents ID="ge" runat="server" >
        <ClientSideEvents ControlsInitialized="function(s, e) {
	    OnControlsInitializedSplitter();
	    OnControlsInitializedGrid();
    }" />
    </dx:ASPxGlobalEvents>

    <dx:ASPxCallback ID="AffiliateSubmit" runat="server" ClientInstanceName = "AffiliateSubmit">
        <ClientSideEvents EndCallback="function(s, e) {
            var pMsg = s.cpMessage;        
            if (pMsg != '') {
                if (s.cpType == 'error'){
                    lblInfo.GetMainElement().style.color = 'Red';
                }
                else if (s.cpType == 'info'){
                    lblInfo.GetMainElement().style.color = 'Blue';
                }
                else {
                    lblInfo.GetMainElement().style.color = 'Red';
                }
        
                lblInfo.SetText(pMsg);
           
               if (s.cpFunction == 'delete'){
                    if (s.cpType != 'error'){
                        clear();
                    }
                }else if(s.cpFunction == 'insert'){

                }
                
            } else {
                lblInfo.SetText('');
            }  
        }" />
    </dx:ASPxCallback>

    <dx:ASPxCallback ID="cbBind" runat="server" 
                            ClientInstanceName="cbBind">
                            <ClientSideEvents EndCallback="function(s, e) {
	                            txtEmailAddress.SetText(s.cptxtEmailAddress);
                                txtUserName.SetText(s.cptxtUserName);
                                txtPassword.SetText(s.cptxtPassword);
                                txtPort.SetText(s.cptxtPort);
                                txtPop3.SetText(s.cptxtPop3);
                                txtAttachmentSave.SetText(s.cptxtAttachmentSave);
                                txtAttachmentBackup.SetText(s.cptxtAttachmentBackup);
                                txtSchedule.SetText(s.cptxtSchedule);
                                txtUserNameSMTP.SetText(s.cptxtUserNameSMTP);
                                txtPasswordSMTP.SetText(s.cptxtPasswordSMTP);
                                txtSMTP.SetText(s.cptxtSMTP);
                                txtPortSMTP.SetText(s.cptxtPortSMTP);
                                txtTemplate.SetText(s.cptxtTemplate);
                                txtResult.SetText(s.cptxtResult);
                                txtSendExcel.SetText(s.cptxtSendExcel);
                                txtPO.SetText(s.cptxtPO);
                                txtPOInterval.SetText(s.cptxtPOInterval);
                                txtPORevision.SetText(s.cptxtPORevision);
                                txtPORevisionInterval.SetText(s.cptxtPORevisionInterval);
                                txtKanban.SetText(s.cptxtKanban);
                                txtKanbanInterval.SetText(s.cptxtKanbanInterval);

}" />
                        </dx:ASPxCallback>

    </asp:Content>
