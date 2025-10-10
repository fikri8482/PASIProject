<%@ Page Title="" Language="vb" AutoEventWireup="false" MasterPageFile="~/SiteContent.Master" CodeBehind="ForwarderEmailMaster.aspx.vb" Inherits="PASISystem.ForwarderEmailMaster" %>
<%@ Register Assembly="DevExpress.Web.v14.1, Version=14.1.6.0, Culture=neutral, PublicKeyToken=b88d1754d700e49a" Namespace="DevExpress.Web.ASPxCallback" TagPrefix="dx" %>
<%@ Register Assembly="DevExpress.Web.v14.1, Version=14.1.6.0, Culture=neutral, PublicKeyToken=b88d1754d700e49a" Namespace="DevExpress.Web.ASPxGridView.Export" TagPrefix="dx" %>
<%@ Register Assembly="DevExpress.Web.v14.1, Version=14.1.6.0, Culture=neutral, PublicKeyToken=b88d1754d700e49a" Namespace="DevExpress.Web.ASPxGridView" TagPrefix="dx" %>
<%@ Register assembly="DevExpress.Web.v14.1, Version=14.1.6.0, Culture=neutral, PublicKeyToken=b88d1754d700e49a" namespace="DevExpress.Web.ASPxEditors" tagprefix="dx" %>
<%@ Register Assembly="DevExpress.Web.v14.1, Version=14.1.6.0, Culture=neutral, PublicKeyToken=b88d1754d700e49a" Namespace="DevExpress.Web.ASPxPanel" TagPrefix="dx" %>
<%@ Register Assembly="DevExpress.Web.v14.1, Version=14.1.6.0, Culture=neutral, PublicKeyToken=b88d1754d700e49a" Namespace="DevExpress.Web.ASPxPopupControl" TagPrefix="dx" %>
<%@ Register Assembly="DevExpress.Web.v14.1, Version=14.1.6.0, Culture=neutral, PublicKeyToken=b88d1754d700e49a" namespace="DevExpress.Web.ASPxGlobalEvents" TagPrefix="dx" %>
<%@ Register assembly="DevExpress.Web.v14.1, Version=14.1.6.0, Culture=neutral, PublicKeyToken=b88d1754d700e49a" namespace="DevExpress.Web.ASPxHiddenField" tagprefix="dx1" %>
<%@ Register assembly="DevExpress.Web.v14.1, Version=14.1.6.0, Culture=neutral, PublicKeyToken=b88d1754d700e49a" namespace="DevExpress.Web.ASPxGridView.Export" tagprefix="dx2" %>
<asp:Content ID="Content1" ContentPlaceHolderID="head" runat="server">
<style type="text/css">
.dxeHLC, .dxeHC, .dxeHFC
{
display: none;
}
    .style1
    {
        width: 188px;
    }
    .style3
    {
        width: 56px;
    }
    .style4
    {
        width: 148px;
    }
    .style5
    {
        width: 200px;
    }
    .style6
    {
        width: 70px;
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

    function validasubmit() {
        lblInfo.GetMainElement().style.color = 'Red';
        if (cboForwarderID.GetText() == "") {
            lblInfo.SetText("[6011] Please Input Forwarder Code first!");
            cboForwarderID.Focus();
            e.ProcessOnServer = false;
            return false;
        }

    }

    function up_Insert() {
        var pIsUpdate = '';
        var pForwarderID = cboForwarderID.GetText();
        var pForwarderExportTO = txtForwarderPOExportTO.GetText();
        var pForwarderExportCC = txtForwarderExportCC.GetText();
        var pForwarderRevisionTO = txtForwarderRevisionTO.GetText();
        var pForwarderRevisionCC = txtForwarderRevisionCC.GetText();
        var pSupplierDeliveryTO = txtSupplierDeliveryTO.GetText();
        var pSupplierDeliveryCC = txtSupplierDeliveryCC.GetText();
        var pForwarderReceivingTO = txtForwarderReceivingTO.GetText();
        var pForwarderReceivingCC = txtForwarderReceivingCC.GetText();

        AffiliateSubmit.PerformCallback('save|' + pIsUpdate + '|' + pForwarderID + '|' + pForwarderExportTO + '|' + pForwarderExportCC + '|' + pForwarderRevisionTO + '|' + pForwarderRevisionCC + '|' + pSupplierDeliveryTO + '|' + pSupplierDeliveryCC + '|' + pForwarderReceivingTO + '|' + pForwarderReceivingCC);

    }

    function readonly() {
        txtForwarderName.GetInputElement().setAttribute('style', 'background:#CCCCCC;foreground:#CCCCCC;');
        txtForwarderName.GetInputElement().readOnly = true;
        lblInfo.SetText('');
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
    <table style="border-width: 1pt thin thin thin; border-style: ridge;  border-color:#9598A1; width:100%; height: 27px;">
        <tr>
            <td>
    <table style="width:100%;">
            <tr>
                <td align="left" class="style5">
                            <dx:ASPxLabel ID="ASPxLabel20" runat="server" Text="FORWARDER CODE/NAME" 
                                Font-Names="Tahoma" font-size="8pt" Width="150px" >
                            </dx:ASPxLabel>
                        </td>
                <td align="left" class="style6">
                            <dx:ASPxComboBox ID="cboForwarderID" runat="server" Height="16px"  
                                            ClientInstanceName="cboForwarderID" Width="130px"
                                            Font-Size="8pt" 
                                            Font-Names="Tahoma" TextFormatString="{0}" TabIndex="1">
                                            <ClientSideEvents SelectedIndexChanged="function(s, e) {
	txtForwarderName.SetText(cboForwarderID.GetSelectedItem().GetColumnText(1));
	cbSetData.PerformCallback(cboForwarderID.GetSelectedItem().GetColumnText(0)); 
	lblInfo.SetText('');
				}" />
                                            <LoadingPanelStyle ImageSpacing="5px">
                                            </LoadingPanelStyle>
                                     
                            </dx:ASPxComboBox>
                        </td>
                <td align="left" width="201px">
                            <dx:ASPxTextBox ID="txtForwarderName" runat="server" Width="320px" 
                                Font-Names="Tahoma" font-size="8pt" ClientInstanceName="txtForwarderName" 
                                MaxLength="2000" BackColor="#CCCCCC" style="margin-left: 0px" 
                                TabIndex="2" >
                                <ClientSideEvents KeyPress="function(s, e) {
	readonly();
}" />
                            </dx:ASPxTextBox>
                        </td>
                <td align="left" width="201px">
                            &nbsp;</td>
                <td align="left" width="201px">
                            &nbsp;</td>
                <td align="left" width="180px">
                            &nbsp;</td>
            </tr>
        </table>
        </td>            
        </tr>
    </table>
    
    <div style="height:5px;"></div> 

    <table style="border-width: thin; border-style: inset hidden ridge hidden; border-color:#9598A1; width:100%; height: 15px;">
        <tr>
            <td colspan="8" height="15">
                <%--error message--%>
                <table id="info" style="width:100%;">
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

    <table style="border-width: 1pt thin thin thin; border-style: ridge;  border-color:#9598A1; width:100%; height: 450px;">
        <tr>
            <td>
                <table style="width:100%; height: 420px;">
                    <tr>
                        <td align="left" class="style4" valign="top">
                            <dx:ASPxLabel ID="ASPxLabel10" runat="server" Font-Names="Tahoma" 
                                Font-Size="8pt" Text="FORWARDER PO EXPORT" Width="200px">
                            </dx:ASPxLabel>
                        </td>
                        <td align="left" class="style3" valign="top">
                            <dx:ASPxLabel ID="ASPxLabel21" runat="server" Font-Names="Tahoma" 
                                Font-Size="8pt" Text="EMAIL TO">
                            </dx:ASPxLabel>
                        </td>
                        <td align="left" class="style1" valign="top">
                            <dx:ASPxMemo ID="txtForwarderPOExportTO" runat="server" Height="42px" Width="320px"
                            ClientInstanceName="txtForwarderPOExportTO" Font-Names="Tahoma" TabIndex="3">
                            </dx:ASPxMemo>
                        </td>
                        <td align="right" valign="top">
                            <dx:ASPxLabel ID="ASPxLabel30" runat="server" Font-Names="Tahoma" 
                                Font-Size="8pt" Text="EMAIL CC">
                            </dx:ASPxLabel>
                        </td>
                        <td align="left" valign="top">     
                        <dx:ASPxMemo ID="txtForwarderExportCC" runat="server" Height="42px" Width="320px"
                        ClientInstanceName="txtForwarderExportCC" Font-Names="Tahoma" MaxLength="2000" 
                                TabIndex="4">
                        </dx:ASPxMemo>
                        </td>
                    </tr>
                    <tr>
                        <td align="left" class="style4" valign="top">
                            <dx:ASPxLabel ID="ASPxLabel11" runat="server" Font-Names="Tahoma" 
                                Font-Size="8pt" Text="FORWARDER PO REVISION" Width="200px">
                            </dx:ASPxLabel>
                        </td>
                        <td align="left" class="style3" valign="top">
                            <dx:ASPxLabel ID="ASPxLabel22" runat="server" Font-Names="Tahoma" 
                                Font-Size="8pt" Text="EMAIL TO">
                            </dx:ASPxLabel>
                        </td>
                        <td align="left" class="style1" valign="top">
                            <dx:ASPxMemo ID="txtForwarderRevisionTO" runat="server" Height="42px" Width="320px"
                            ClientInstanceName="txtForwarderRevisionTO" Font-Names="Tahoma" 
                                MaxLength="2000" TabIndex="5">
                            </dx:ASPxMemo>
                        </td>
                        <td align="right" valign="top">
                            <dx:ASPxLabel ID="ASPxLabel31" runat="server" Font-Names="Tahoma" 
                                Font-Size="8pt" Text="EMAIL CC">
                            </dx:ASPxLabel>
                        </td>
                        <td align="left" valign="top">
                        <dx:ASPxMemo ID="txtForwarderRevisionCC" runat="server" Height="42px" Width="320px"
                            ClientInstanceName="txtForwarderRevisionCC" Font-Names="Tahoma" 
                                MaxLength="2000" TabIndex="6">
                            </dx:ASPxMemo>
                        </td>
                    </tr>
                    <tr>
                        <td align="left" class="style4" valign="top">
                            <dx:ASPxLabel ID="ASPxLabel12" runat="server" Font-Names="Tahoma" 
                                Font-Size="8pt" Text="SUPPLIER DELIVERY CONFIRMATION" Width="200px">
                            </dx:ASPxLabel>
                        </td>
                        <td align="left" class="style3" valign="top">
                            <dx:ASPxLabel ID="ASPxLabel23" runat="server" Font-Names="Tahoma" 
                                Font-Size="8pt" Text="EMAIL TO">
                            </dx:ASPxLabel>
                        </td>
                        <td align="left" class="style1" valign="top">
                        <dx:ASPxMemo ID="txtSupplierDeliveryTO" runat="server" Height="42px" Width="320px"
                            ClientInstanceName="txtSupplierDeliveryTO" Font-Names="Tahoma" MaxLength="2000" 
                                TabIndex="7">
                            </dx:ASPxMemo>
                        </td>
                        <td align="right" valign="top">
                            <dx:ASPxLabel ID="ASPxLabel32" runat="server" Font-Names="Tahoma" 
                                Font-Size="8pt" Text="EMAIL CC">
                            </dx:ASPxLabel>
                        </td>
                        <td align="left" valign="top">
                        <dx:ASPxMemo ID="txtSupplierDeliveryCC" runat="server" Height="42px" Width="320px"
                            ClientInstanceName="txtSupplierDeliveryCC" Font-Names="Tahoma" MaxLength="2000" 
                                TabIndex="8">
                            </dx:ASPxMemo>
                        </td>
                    </tr>
                    <tr>
                        <td align="left" class="style4" valign="top">
                            <dx:ASPxLabel ID="ASPxLabel13" runat="server" Font-Names="Tahoma" 
                                Font-Size="8pt" Text="FORWARDER RECEIVING" Width="200px">
                            </dx:ASPxLabel>
                        </td>
                        <td align="left" class="style3" valign="top">
                            <dx:ASPxLabel ID="ASPxLabel24" runat="server" Font-Names="Tahoma" 
                                Font-Size="8pt" Text="EMAIL TO">
                            </dx:ASPxLabel>
                        </td>
                        <td align="left" class="style1" valign="top">
                        <dx:ASPxMemo ID="txtForwarderReceivingTO" runat="server" Height="42px" Width="320px"
                            ClientInstanceName="txtForwarderReceivingTO" Font-Names="Tahoma" 
                                MaxLength="2000" TabIndex="9">
                            </dx:ASPxMemo>
                        </td>
                        <td align="right" valign="top">
                            <dx:ASPxLabel ID="ASPxLabel33" runat="server" Font-Names="Tahoma" 
                                Font-Size="8pt" Text="EMAIL CC">
                            </dx:ASPxLabel>
                        </td>
                        <td align="left" valign="top">
                        <dx:ASPxMemo ID="txtForwarderReceivingCC" runat="server" Height="42px" Width="320px"
                            ClientInstanceName="txtForwarderReceivingCC" Font-Names="Tahoma" 
                                MaxLength="2000" TabIndex="10">
                            </dx:ASPxMemo>
                        </td>
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
                    Width="90px" Font-Size="8pt" ClientInstanceName="btnSubMenu" TabIndex="23">
                </dx:ASPxButton>   
                </td>                     
            
            <td valign="top" align="right" style="width: 50px;">                                  
            </td>
            <td valign="top" align="right" style="width: 50px;">                                   
                &nbsp;</td>
            <td align="right" style="width:80px;">                                   
                <dx:ASPxButton ID="btnClear" runat="server" Text="CLEAR"                    
                    Font-Names="Tahoma"
                    Width="90px" AutoPostBack="False" Font-Size="8pt" 
                    ClientInstanceName="btnClear" TabIndex="22">
                    <ClientSideEvents Click="function(s, e) {
                    cboForwarderID.SetText('');
                    txtForwarderName.SetText('');
                    txtForwarderPOExportTO.SetText('');
                    txtForwarderExportCC.SetText('');
                    txtForwarderRevisionTO.SetText('');
                    txtForwarderRevisionCC.SetText('');
                    txtSupplierDeliveryTO.SetText('');
                    txtSupplierDeliveryCC.SetText('');
                    txtForwarderReceivingTO.SetText('');
                    txtForwarderReceivingCC.SetText('');
                    lblInfo.SetText('');
}" />
                </dx:ASPxButton>
            </td>
            <td valign="top" align="right" style="width: 50px;">                                   
                <dx:ASPxButton ID="btnSubmit" runat="server" Text="SAVE"                    
                    Font-Names="Tahoma"
                    Width="90px" AutoPostBack="False" Font-Size="8pt" 
                    ClientInstanceName="btnSubmit" TabIndex="21">
                    <ClientSideEvents Click="function(s, e) {
                        validasubmit();
                        up_Insert();
                       
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
               if(s.cpFunction == 'insert'){
                }
            } else {
                lblInfo.SetText('');
            }  
        }" />
    </dx:ASPxCallback>
    <dx:ASPxCallback ID="cbSetData" runat="server" ClientInstanceName = "cbSetData">
        <ClientSideEvents CallbackComplete="function(s, e) {
                txtForwarderPOExportTO.SetText(s.cpPOExportTO);
                txtForwarderExportCC.SetText(s.cpPOExportCC);
                txtForwarderRevisionTO.SetText(s.cpPORevisionTO);
                txtForwarderRevisionCC.SetText(s.cpPORevisionCC);
                txtSupplierDeliveryTO.SetText(s.cpSupplierDeliveryTO);
                txtSupplierDeliveryCC.SetText(s.cpSupplierDeliveryCC);
                txtForwarderReceivingTO.SetText(s.cpForwarderReceivingTO);
                txtForwarderReceivingCC.SetText(s.cpForwarderReceivingCC);
          }" />
    </dx:ASPxCallback>
</asp:Content>
