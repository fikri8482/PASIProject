<%@ Page Title="" Language="vb" AutoEventWireup="false" MasterPageFile="~/SiteContent.Master" CodeBehind="EmailMasterSetting.aspx.vb" Inherits="PASISystem.EmailMasterSetting" %>

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
        width: 192px;
    }
    .style2
    {
        width: 319px;
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
        
        if (txtUserNameSMTP.GetText() == "") {
            lblInfo.SetText("[6011] Please Input Username SMTP first!");
            txtUserNameSMTP.Focus();
            e.ProcessOnServer = false;
            return false;
        }
        
        if (txtPasswordSMTP.GetText() == "") {
            lblInfo.SetText("[6011] Please Input Password SMTP first!");
            txtPasswordSMTP.Focus();
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
        
     
        AffiliateSubmit.PerformCallback('save|' + pIsUpdate + '|' + pEmailAddress + '|' + pUserName + '|' + pPassword + '|' + pPOP3 + '|' + pPort + '|' + pAttachmentSave + '|' + pAttachmentBackup + '|' + pInterval);
        
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
        
        AffiliateSubmit.PerformCallback('delete|' + pEmailAddress + '|' + pUserName + '|' + pPassword + '|' + pPOP3 + '|' + pPort + '|' + pAttachmentSave + '|' + pAttachmentBackup + '|' + pInterval);
      
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
            grid.SetHeight(height);
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
                            <dx:ASPxLabel ID="lblInfo" runat="server" Font-Names="Verdana"
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

    <table style="border-width: 1pt thin thin thin; border-style: ridge; border-color:#9598A1; width:100%; height: 250px;">
        <tr>
            <td>
                <table style="width:100%;">
                    <tr>
                        <td align="left" class="style1">
                            <dx:ASPxLabel ID="ASPxLabel10" runat="server" Font-Names="Verdana" 
                                Font-Size="8pt" Text="EMAIL ADDRESS">
                            </dx:ASPxLabel>
                        </td>
                        <td align="left" class="style2">
                            <dx:ASPxTextBox ID="txtEmailAddress" runat="server" Width="400px" 
                                ClientInstanceName="txtEmailAddress" Font-Names="Verdana" 
                                MaxLength="100" Height="25px">
                                <ClientSideEvents ValueChanged="function(s, e) {
	                             alert ('1');
	cbSetData.PerformCallback(); 
    alert ('2');
	                                            lblInfo.SetText('');
}" LostFocus="function(s, e) {
	cbSetData.PerformCallback(); 
}" TextChanged="function(s, e) {
	cbSetData.PerformCallback(); 
}" />
                            </dx:ASPxTextBox>
                        </td>
                        <td align="left">
                            &nbsp;</td>
                    </tr>
                    <tr>
                        <td align="left" class="style1">
                            <dx:ASPxLabel ID="ASPxLabel11" runat="server" Font-Names="Verdana" 
                                Font-Size="8pt" Text="USER NAME">
                            </dx:ASPxLabel>
                        </td>
                        <td align="left" class="style2">
                            <dx:ASPxTextBox ID="txtUserName" runat="server" Width="400px" 
                                ClientInstanceName="txtUserName" Font-Names="Verdana" 
                                MaxLength="100" onkeypress="return singlequote(event)" Height="25px">
                            </dx:ASPxTextBox>
                        </td>
                        <td align="left">
                            &nbsp;</td>
                    </tr>
                    <tr>
                        <td align="left" class="style1">
                            <dx:ASPxLabel ID="ASPxLabel12" runat="server" Font-Names="Verdana" 
                                Font-Size="8pt" Text="PASSWORD">
                            </dx:ASPxLabel>
                        </td>
                        <td align="left" class="style2">
                            <dx:ASPxTextBox ID="txtPassword" runat="server" Width="400px" 
                                ClientInstanceName="txtPassword" Font-Names="Verdana" 
                                MaxLength="20" onkeypress="return singlequote(event)" Height="25px">
                            </dx:ASPxTextBox>
                        </td>
                        <td align="left">
                            &nbsp;</td>
                    </tr>
                    <tr>
                        <td align="left" class="style1">
                            <dx:ASPxLabel ID="ASPxLabel14" runat="server" Font-Names="Verdana" 
                                Font-Size="8pt" Text="PORT">
                            </dx:ASPxLabel>
                        </td>
                        <td align="left" class="style2">
                            <dx:ASPxTextBox ID="txtPort" runat="server" Width="100px" 
                                ClientInstanceName="txtPort" Font-Names="Verdana" 
                                MaxLength="4" onkeypress="return numbersonly(event)" Height="25px">
                                <ClientSideEvents LostFocus="function(s, e) { lblInfo.SetText(&quot;&quot;); }" />
                            </dx:ASPxTextBox>
                        </td>
                        <td align="left">
                            &nbsp;</td>
                    </tr>
                    <tr>
                        <td align="left" class="style1">
                            <dx:ASPxLabel ID="ASPxLabel13" runat="server" Font-Names="Verdana" 
                                Font-Size="8pt" Text="POP 3">
                            </dx:ASPxLabel>
                        </td>
                        <td align="left" class="style2">
                            <dx:ASPxTextBox ID="txtPop3" runat="server" Width="400px" 
                                ClientInstanceName="txtPop3" Font-Names="Verdana" 
                                MaxLength="100" Height="25px">
                            </dx:ASPxTextBox>
                        </td>
                        <td align="left">
                            &nbsp;</td>
                    </tr>
                    <tr>
                        <td align="left" class="style1">
                            <dx:ASPxLabel ID="ASPxLabel15" runat="server" Font-Names="Verdana" 
                                Font-Size="8pt" Text="ATTACHMENT SAVE FOLDER">
                            </dx:ASPxLabel>
                        </td>
                        <td align="left" class="style2">
                            <dx:ASPxTextBox ID="txtAttachmentSave" runat="server" Width="400px" 
                                ClientInstanceName="txtAttachmentSave" Font-Names="Verdana" 
                                MaxLength="100" Height="25px" Text="D:\">
                            </dx:ASPxTextBox>
                        </td>
                        <td align="left">
                            <dx:ASPxButton ID="btnBrowse" runat="server" Text="..." 
                            Font-Names="Verdana">
                                <clientsideevents click="function(s, e) {
	&quot;btnBrowse_Click&quot;
}" />
                            </dx:ASPxButton>
                        </td>
                    </tr>
                    <tr>
                        <td align="left" class="style1">
                            <dx:ASPxLabel ID="ASPxLabel17" runat="server" Font-Names="Verdana" 
                                Font-Size="8pt" Text="ATTACHMENT BACKUP">
                            </dx:ASPxLabel>
                        </td>
                        <td align="left" class="style2">
                            <dx:ASPxTextBox ID="txtAttachmentBackup" runat="server" Width="400px" 
                                ClientInstanceName="txtAttachmentBackup" Font-Names="Verdana" 
                                MaxLength="100" Height="25px" Text="D:\">
                            
                            </dx:ASPxTextBox>
                        </td>
                        <td align="left">
                            <dx:ASPxButton ID="ASPxButton2" runat="server" Text="...">
                            </dx:ASPxButton>
                        </td>
                    </tr>
                    <tr>
                        <td align="left" class="style1">
                            <dx:ASPxLabel ID="ASPxLabel16" runat="server" Font-Names="Verdana" 
                                Font-Size="8pt" Text="SCHEDULE EVERY (SECONDS)">
                            </dx:ASPxLabel>
                        </td>
                        <td align="left" class="style2">
                            <dx:ASPxTextBox ID="txtSchedule" runat="server" Width="100px" 
                                ClientInstanceName="txtSchedule" onkeypress="return numbersonly(event)" Font-Names="Verdana" 
                                MaxLength="20"  Height="25px">
                            </dx:ASPxTextBox>
                        </td>
                        <td align="left">
                            &nbsp;</td>
                    </tr>
                    <tr>
                        <td align="left" class="style1">
                            &nbsp;</td>
                        <td align="left" class="style2">
                            &nbsp;</td>
                        <td align="left">
                            &nbsp;</td>
                    </tr>
                    <tr>
                        <td align="left" class="style1">
                            &nbsp;</td>
                        <td align="left" class="style2">
                            &nbsp;</td>
                        <td align="left">
                            &nbsp;</td>
                    </tr>
                    <tr>
                        <td align="left" class="style1">
                            &nbsp;</td>
                        <td align="left" class="style2">
                            &nbsp;</td>
                        <td align="left">
                            &nbsp;</td>
                    </tr>
                    </table>
            </td>
        </tr>

    </table> 

    

   

    <table style="width: 100%;">
        <tr>
            <td>
                &nbsp;
            </td>
            <td>
                &nbsp;
            </td>
            <td>
                &nbsp;
            </td>
        </tr>
    </table>

    <div style="height:8px;"></div>      

    <%--Button--%> 
    <table id="button" style=" width:100%;">
        <tr>                        
            <td valign="top" align="left">
                <dx:ASPxButton ID="btnSubMenu" runat="server" Text="SUB MENU"                     
                    Font-Names="Verdana"
                    Width="90px" Font-Size="8pt">
                </dx:ASPxButton>   
                </td>                     
            
            <td valign="top" align="right" style="width: 50px;">                                  
            </td>
            <td valign="top" align="right" style="width: 50px;">                                   
                &nbsp;</td>
            <td align="right" style="width:80px;">                                   
                <dx:ASPxButton ID="btnDelete" runat="server" Text="DELETE"                    
                    Font-Names="Verdana"
                    Width="90px" AutoPostBack="False" Font-Size="8pt" 
                    ClientInstanceName="btnDelete">
                    <ClientSideEvents Click="function(s, e) {
                  
                        validasubmit();
                     
                        up_delete();
                     
                        clear();
                      
                    }" />
                </dx:ASPxButton>
            </td>
            <td valign="top" align="right" style="width: 50px;">                                   
                <dx:ASPxButton ID="btnSubmit" runat="server" Text="SUBMIT"                    
                    Font-Names="Verdana"
                    Width="90px" AutoPostBack="False" Font-Size="8pt">
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
    <dx:ASPxCallback ID="cbBrowse" runat="server" 
                            ClientInstanceName="cbBrowse">
                            <ClientSideEvents CallbackComplete="function(s, e) {
	ASPxProgressBar1.SetValue(e.result);
}" />
                        </dx:ASPxCallback>

    <dx:ASPxCallback ID="cbSetData" runat="server" ClientInstanceName="cbSetData">
        <ClientSideEvents CallbackComplete="function(s, e) {
                if (s.cpUsername) {
				txtUserName.SetText(s.cpUsername);
				}
                
	            if (s.cpPassword) {
				txtPassword.SetText(s.cpPassword);
				}

                if (s.cpPort) {
				txtPort.SetText(s.cpPort);
				}

                if (s.cpPOP3) {
				txtPop3.SetText(s.cpPOP3);
				}

                if (s.cpAttachmentFolder) {
				txtAttachmentSave.SetText(s.cpAttachmentFolder);
				}

                if (s.cpAttachmentBackupFolder) {
				txtAttachmentBackup.SetText(s.cpAttachmentBackupFolder);
				}

                if (s.cpInterval) {
				txtSchedule.SetText(s.cpInterval);
				}

                
}" EndCallback="function(s, e) {

                if (s.cpUsername) {
				txtUserName.SetText(s.cpUsername);
				}

				if (s.cpPassword) {
				txtPassword.SetText(s.cpPassword);
				}

                if (s.cpPort) {
				txtPort.SetText(s.cpPort);
				}

                if (s.cpPOP3) {
				txtPop3.SetText(s.cpPOP3);
				}

                if (s.cpAttachmentFolder) {
				txtAttachmentSave.SetText(s.cpAttachmentFolder);
				}

                if (s.cpAttachmentBackupFolder) {
				txtAttachmentBackup.SetText(s.cpAttachmentBackupFolder);
				}

                if (s.cpInterval) {
				txtSchedule.SetText(s.cpInterval);
				}

}" />
    </dx:ASPxCallback>
    </asp:Content>
