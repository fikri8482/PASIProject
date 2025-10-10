<%@ Page Title="" Language="vb" AutoEventWireup="false" MasterPageFile="~/SiteContent.Master" CodeBehind="NotificationMaster.aspx.vb" Inherits="PASISystem.NotificationMaster" %>

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
    .style4
    {
        width: 148px;
    }
    .style5
    {
        width: 238px;
    }
    .style6
    {
        width: 70px;
    }
    .style7
    {
        width: 114px;
    }
</style> 

<script language="javascript" type="text/javascript">
//    function OnGridFocusedRowChanged() {
//        grid.GetRowValues(grid.GetFocusedRowIndex(), "AffiliateID;AffiliateName;Address;City;PostalCode;Phone1;Phone2;Fax;NPWP;PODeliveryBy", OnGetRowValues);
//    }
//    function OnGetRowValues(values) {
//        if (values[0] != "" && values[0] != null && values[0] != "null") {
//            txtAffiliateID.SetText(values[0]);
//            txtAffiliateName.SetText(values[1]);
//            txtAddress1.SetText(values[2]);
//            txtAddress2.SetText(values[3]);
//            txtCity.SetText(values[4]);
//            txtProvince.SetText(values[5]);
//            txtPostalCode.SetText(values[6]);
//            txtPhone.SetText(values[7]);
//            txtFax.SetText(values[8]);
//            lblInfo.SetText('');
//            txtAffiliateID.GetInputElement().setAttribute('style', 'background:#CCCCCC;');
//            txtAffiliateID.GetInputElement().readOnly = true;
//        }
//    }

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
        cboNotificationCode.SetText('');
        txtNotificationCode.SetText('');
        txtLine1.SetText('');
        cboLine1.SetText('');
        txtLine2.SetText('');
        cboLine2.SetText('');
        txtLine3.SetText('');
        cboLine3.SetText('');
        txtLine4.SetText('');
        cboLine4.SetText('');
        txtLine5.SetText('');
        cboLine5.SetText('');
        txtLine6.SetText('');
        cboLine6.SetText('');
        txtLine7.SetText('');
        cboLine7.SetText('');
        txtLine8.SetText('');
        cboLine8.SetText('');

        txtNotificationCode.GetInputElement().setAttribute('style', 'background:#FFFFFF;foreground:#FFFFFF;');
        txtNotificationCode.GetInputElement().readOnly = true;

        cboNotificationCode.GetInputElement().readOnly = true;
    }

//    function clear2() {
//        txtSupplierCode.GetInputElement().setAttribute('style', 'background:#CCCCCC;foreground:#CCCCCC;');
//        txtSupplierCode.GetInputElement().readOnly = false;
//    }

    function validasubmit() {
        lblInfo.GetMainElement().style.color = 'Red';
        if (cboNotificationCode.GetText() == "") {
            lblInfo.SetText("[6011] Please Input Notification Code first!");
            cboNotificationCode.Focus();
            e.ProcessOnServer = false;
            return false;
          
        }
             
    }

    function up_Insert() {
        var pIsUpdate = '';
        var pNotificationCode = cboNotificationCode.GetText();
        var pLine1 = txtLine1.GetText();
        var pLineCls1 = cboLine1.GetValue();
        var pLine2 = txtLine2.GetText();
        var pLineCls2 = cboLine2.GetValue();
        var pLine3 = txtLine3.GetText();
        var pLineCls3 = cboLine3.GetValue();
        var pLine4 = txtLine4.GetText();
        var pLineCls4 = cboLine4.GetValue();
        var pLine5 = txtLine5.GetText();
        var pLineCls5 = cboLine5.GetValue();
        var pLine6 = txtLine6.GetText();
        var pLineCls6 = cboLine6.GetValue();
        var pLine7 = txtLine7.GetText();
        var pLineCls7 = cboLine7.GetValue();
        var pLine8 = txtLine8.GetText();
        var pLineCls8 = cboLine8.GetValue();
        var pPODel = '';

        AffiliateSubmit.PerformCallback('save|' + pIsUpdate + '|' + pNotificationCode + '|' + pLine1 + '|' + pLineCls1 + '|' + pLine2 + '|' + pLineCls2 + '|' + pLine3 + '|' + pLineCls3 + '|' + pLine4 + '|' + pLineCls4 + '|' + pLine5 + '|' + pLineCls5 + '|' + pLine6 + '|' + pLineCls6 + '|' + pLine7 + '|' + pLineCls7 + '|' + pLine8 + '|' + pLineCls8);

    }

    function readonly() {
        txtNotificationCode.GetInputElement().setAttribute('style', 'background:#CCCCCC;foreground:#CCCCCC;');
        txtNotificationCode.GetInputElement().readOnly = true;
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
                            <dx:ASPxLabel ID="ASPxLabel20" runat="server" Text="NOTIFICATION CODE" 
                                Font-Names="Tahoma" font-size="8pt" >
                            </dx:ASPxLabel>
                        </td>
                <td align="left" class="style6">
                            <dx:ASPxComboBox ID="cboNotificationCode" runat="server" Height="16px"  
                                            ClientInstanceName="cboNotificationCode" Width="70px"
                                            Font-Size="8pt" 
                                            Font-Names="Tahoma" TextFormatString="{0}">
                                            <ClientSideEvents SelectedIndexChanged="function(s, e) {
txtNotificationCode.SetText(cboNotificationCode.GetSelectedItem().GetColumnText(1));

cbSetData.PerformCallback(cboNotificationCode.GetSelectedItem().GetColumnText(0)); 

lblInfo.SetText('');

                                            }" />
                                            <LoadingPanelStyle ImageSpacing="5px">
                                            </LoadingPanelStyle>
                                     
                            </dx:ASPxComboBox>
                        </td>
                <td align="left" width="201px">
                            <dx:ASPxTextBox ID="txtNotificationCode" runat="server" Width="330px" 
                                Font-Names="Tahoma" font-size="8pt" ClientInstanceName="txtNotificationCode" 
                                MaxLength="25" BackColor="#CCCCCC" style="margin-left: 0px" >
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

    <table style="border-width: 1pt thin thin thin; border-style: ridge; border-color:#9598A1; width:100%; height: 100px;">
        <tr>
            <td>
                <table style="width:100%; height: 450px;">
                    <tr>
                        <td align="left" class="style4" valign="top">
                            <dx:ASPxLabel ID="ASPxLabel10" runat="server" Font-Names="Tahoma" 
                                Font-Size="8pt" Text="LINE 1">
                            </dx:ASPxLabel>
                        </td>
                        <td align="left" valign="top">
                            <dx:ASPxMemo ID="txtLine1" runat="server" ClientInstanceName="txtLine1" 
                                Height="42px" Width="300px"
                            Font-Names="Tahoma" MaxLength="200">
                            </dx:ASPxMemo>
                        </td>
                        <td align="right" valign="top" class="style7">
                            <dx:ASPxLabel ID="ASPxLabel30" runat="server" Font-Names="Tahoma" 
                                Font-Size="8pt" Text="INCLUDE">
                            </dx:ASPxLabel>
                        </td>
                        <td align="left" valign="top">     
                            <dx:ASPxComboBox ID="cboLine1" runat="server" 
                                Width="200px" ClientInstanceName="cboLine1" TextFormatString="{1)">
                                <ClientSideEvents SelectedIndexChanged="function(s, e) {
                               
    if (cboLine1.GetText() == 'TEXT') {
        txtLine1.SetEnabled(true);

    } else { 
        txtLine1.SetEnabled(false);
        txtLine1.SetText('');
    }
    lblInfo.SetText('');
}"  />
                            </dx:ASPxComboBox>
                        </td>
                        <td align="left" valign="top">     
                            &nbsp;</td>
                    </tr>
                    <tr>
                        <td align="left" valign="top">
                            <dx:ASPxLabel ID="ASPxLabel11" runat="server" Font-Names="Tahoma" 
                                Font-Size="8pt" Text="LINE 2">
                            </dx:ASPxLabel>
                        </td>
                        <td align="left" valign="top">
                            <dx:ASPxMemo ID="txtLine2" runat="server" Height="42px" Width="300px"
                            ClientInstanceName="txtLine2" Font-Names="Tahoma" MaxLength="200">
                            </dx:ASPxMemo>
                        </td>
                        <td align="right" valign="top" class="style7">
                            <dx:ASPxLabel ID="ASPxLabel31" runat="server" Font-Names="Tahoma" 
                                Font-Size="8pt" Text="INCLUDE">
                            </dx:ASPxLabel>
                        </td>
                        <td align="left" valign="top">
                            <dx:ASPxComboBox ID="cboLine2" runat="server" 
                                Width="200px" ClientInstanceName="cboLine2" TextFormatString="{1)">
                                <ClientSideEvents SelectedIndexChanged="function(s, e) {
                               
    if (cboLine2.GetText() == 'TEXT') {
        txtLine2.SetEnabled(true);

    } else {
        txtLine2.SetEnabled(false);
        txtLine2.SetText('');

    }
    lblInfo.SetText('');
}" />
                            </dx:ASPxComboBox>
                        </td>
                        <td align="left" valign="top">
                            &nbsp;</td>
                    </tr>
                    <tr>
                        <td align="left" valign="top">
                            <dx:ASPxLabel ID="ASPxLabel12" runat="server" Font-Names="Tahoma" 
                                Font-Size="8pt" Text="LINE 3">
                            </dx:ASPxLabel>
                        </td>
                        <td align="left" valign="top">
                        <dx:ASPxMemo ID="txtLine3" runat="server" Height="42px" Width="300px"
                            ClientInstanceName="txtLine3" Font-Names="Tahoma" MaxLength="200">
                            </dx:ASPxMemo>
                        </td>
                        <td align="right" valign="top" class="style7">
                            <dx:ASPxLabel ID="ASPxLabel32" runat="server" Font-Names="Tahoma" 
                                Font-Size="8pt" Text="INCLUDE">
                            </dx:ASPxLabel>
                        </td>
                        <td align="left" valign="top">
                            <dx:ASPxComboBox ID="cboLine3" runat="server" 
                                Width="200px" ClientInstanceName="cboLine3" TextFormatString="{1)">
                                <ClientSideEvents SelectedIndexChanged="function(s, e) {
                               
    if (cboLine3.GetText() == 'TEXT') {
        txtLine3.SetEnabled(true);

    } else {
        txtLine3.SetEnabled(false);
        txtLine3.SetText('');

    }
    lblInfo.SetText('');
}" />
                            </dx:ASPxComboBox>
                        </td>
                        <td align="left" valign="top">
                            &nbsp;</td>
                    </tr>
                    <tr>
                        <td align="left" valign="top">
                            <dx:ASPxLabel ID="ASPxLabel13" runat="server" Font-Names="Tahoma" 
                                Font-Size="8pt" Text="LINE 4">
                            </dx:ASPxLabel>
                        </td>
                        <td align="left" valign="top">
                        <dx:ASPxMemo ID="txtLine4" runat="server" Height="42px" Width="300px"
                            ClientInstanceName="txtLine4" Font-Names="Tahoma" MaxLength="200">
                            </dx:ASPxMemo>
                        </td>
                        <td align="right" valign="top" class="style7">
                            <dx:ASPxLabel ID="ASPxLabel33" runat="server" Font-Names="Tahoma" 
                                Font-Size="8pt" Text="INCLUDE">
                            </dx:ASPxLabel>
                        </td>
                        <td align="left" valign="top">
                            <dx:ASPxComboBox ID="cboLine4" runat="server" 
                                Width="200px" ClientInstanceName="cboLine4" TextFormatString="{1)">
                                <ClientSideEvents SelectedIndexChanged="function(s, e) {
                               
    if (cboLine4.GetText() == 'TEXT') {
        txtLine4.SetEnabled(true);

    } else {
        txtLine4.SetEnabled(false);
        txtLine4.SetText('');

    }
    lblInfo.SetText('');
}" />
                            </dx:ASPxComboBox>
                        </td>
                        <td align="left" valign="top">
                            &nbsp;</td>
                    </tr>
                    <tr>
                        <td align="left" valign="top">
                            <dx:ASPxLabel ID="ASPxLabel14" runat="server" Font-Names="Tahoma" 
                                Font-Size="8pt" Text="LINE 5">
                            </dx:ASPxLabel>
                        </td>
                        <td align="left" class="style1" valign="top">
                        <dx:ASPxMemo ID="txtLine5" runat="server" Height="42px" Width="300px"
                            ClientInstanceName="txtLine5" Font-Names="Tahoma" MaxLength="200">
                            </dx:ASPxMemo>
                        </td>
                        <td align="right" valign="top" class="style7">
                            <dx:ASPxLabel ID="ASPxLabel34" runat="server" Font-Names="Tahoma" 
                                Font-Size="8pt" Text="INCLUDE">
                            </dx:ASPxLabel>
                        </td>
                        <td align="left" valign="top">
                            <dx:ASPxComboBox ID="cboLine5" runat="server" 
                                Width="200px" ClientInstanceName="cboLine5" TextFormatString="{1)">
                                <ClientSideEvents SelectedIndexChanged="function(s, e) {
                               
    if (cboLine5.GetText() == 'TEXT') {
        txtLine5.SetEnabled(true);

    } else {
        txtLine5.SetEnabled(false);
        txtLine5.SetText('');

    }
    lblInfo.SetText('');
}" />
                            </dx:ASPxComboBox>
                        </td>
                        <td align="left" valign="top">
                            &nbsp;</td>
                    </tr>
                    <tr>
                        <td align="left" class="style4" valign="top">
                            <dx:ASPxLabel ID="ASPxLabel15" runat="server" Font-Names="Tahoma" 
                                Font-Size="8pt" Text="LINE 6">
                            </dx:ASPxLabel>
                        </td>
                        <td align="left" class="style1" valign="top">
                        <dx:ASPxMemo ID="txtLine6" runat="server" Height="42px" Width="300px"
                            ClientInstanceName="txtLine6" Font-Names="Tahoma" MaxLength="200">
                            </dx:ASPxMemo>
                        </td>
                        <td align="right" valign="top" class="style7">
                            <dx:ASPxLabel ID="ASPxLabel35" runat="server" Font-Names="Tahoma" 
                                Font-Size="8pt" Text="INCLUDE">
                            </dx:ASPxLabel>
                        </td>
                        <td align="left" valign="top">
                            <dx:ASPxComboBox ID="cboLine6" runat="server" 
                                Width="200px" ClientInstanceName="cboLine6" TextFormatString="{1)">
                                <ClientSideEvents SelectedIndexChanged="function(s, e) {
                               
    if (cboLine6.GetText() == 'TEXT') {
        txtLine6.SetEnabled(true);

    } else {
        txtLine6.SetEnabled(false);
        txtLine6.SetText('');

    }
    lblInfo.SetText('');
}" />
                            </dx:ASPxComboBox>
                        </td>
                        <td align="left" valign="top">
                            &nbsp;</td>
                    </tr>
                    <tr>
                        <td align="left" class="style4" valign="top">
                            <dx:ASPxLabel ID="ASPxLabel16" runat="server" Font-Names="Tahoma" 
                                Font-Size="8pt" Text="LINE 7">
                            </dx:ASPxLabel>
                        </td>
                        <td align="left" class="style1" valign="top">
                        <dx:ASPxMemo ID="txtLine7" runat="server" Height="42px" Width="300px"
                            ClientInstanceName="txtLine7" Font-Names="Tahoma" MaxLength="200" 
                                BackColor="White" Theme="Default">
                            </dx:ASPxMemo>
                        </td>
                        <td align="right" valign="top" class="style7">
                            <dx:ASPxLabel ID="ASPxLabel36" runat="server" Font-Names="Tahoma" 
                                Font-Size="8pt" Text="INCLUDE">
                            </dx:ASPxLabel>
                        </td>
                        <td align="left" valign="top">
                            <dx:ASPxComboBox ID="cboLine7" runat="server" 
                                Width="200px" ClientInstanceName="cboLine7" TextFormatString="{1)">
                               <ClientSideEvents SelectedIndexChanged="function(s, e) {
                               
    if (cboLine7.GetText() == 'TEXT') {
        txtLine7.SetEnabled(true);

    } else {
        txtLine7.SetEnabled(false);
        txtLine7.SetText('');

    }
    lblInfo.SetText('');
}" />
                            </dx:ASPxComboBox>
                        </td>
                        <td align="left" valign="top">
                            &nbsp;</td>
                    </tr>
                    <tr>
                        <td align="left" class="style4" valign="top">
                            <dx:ASPxLabel ID="ASPxLabel17" runat="server" Font-Names="Tahoma" 
                                Font-Size="8pt" Text="LINE 8">
                            </dx:ASPxLabel>
                        </td>
                        <td align="left" class="style1" valign="top">
                        <dx:ASPxMemo ID="txtLine8" runat="server" Height="42px" Width="300px"
                            ClientInstanceName="txtLine8" Font-Names="Tahoma" MaxLength="200">
                            </dx:ASPxMemo>
                        </td>
                        <td align="right" valign="top" class="style7">
                            <dx:ASPxLabel ID="ASPxLabel37" runat="server" Font-Names="Tahoma" 
                                Font-Size="8pt" Text="INCLUDE">
                            </dx:ASPxLabel>
                        </td>
                        <td align="left" valign="top">
                            <dx:ASPxComboBox ID="cboLine8" runat="server" 
                                Width="200px" ClientInstanceName="cboLine8" TextFormatString="{1)">
                                <ClientSideEvents SelectedIndexChanged="function(s, e) {
                               
   if (cboLine8.GetText() == 'TEXT') {
        txtLine8.SetEnabled(true);

    } else {
        txtLine8.SetEnabled(false);
        txtLine8.SetText('');
  

    }
    lblInfo.SetText('');
}" />
                            </dx:ASPxComboBox>
                        </td>
                        <td align="left" valign="top">
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
                    Width="90px" Font-Size="8pt" ClientInstanceName="btnSubMenu">
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
                    ClientInstanceName="btnClear">
                    <ClientSideEvents Click="function(s, e) {
                    cboNotificationCode.SetText('');
                    txtNotificationCode.SetText('');
                    cboLine1.SetText('');
                    txtLine1.SetText('');
                    cboLine2.SetText('');
                    txtLine2.SetText('');
                    cboLine3.SetText('');
                    txtLine3.SetText('');
                    cboLine4.SetText('');
                    txtLine4.SetText('');
                    cboLine5.SetText('');
                    txtLine5.SetText('');
                    cboLine6.SetText('');
                    txtLine6.SetText('');
                    cboLine7.SetText('');
                    txtLine7.SetText('');
                    cboLine8.SetText('');
                    txtLine8.SetText('');
                    lblInfo.SetText('');
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
                cboNotificationCode.GetInputElement().readOnly = true;
                }
            } else {
                lblInfo.SetText('');
            }  
        }" />
    </dx:ASPxCallback>
    <dx:ASPxCallback ID="cbSetData" runat="server" ClientInstanceName = "cbSetData">
        <ClientSideEvents CallbackComplete="function(s, e) {
        txtLine1.SetText(s.cpLine1);
        cboLine1.SetText(s.cpLine1Cls);
        txtLine2.SetText(s.cpLine2);
        cboLine2.SetText(s.cpLine2Cls);
        txtLine3.SetText(s.cpLine3);
        cboLine3.SetText(s.cpLine3Cls);
        txtLine4.SetText(s.cpLine4);
        cboLine4.SetText(s.cpLine4Cls);
        txtLine5.SetText(s.cpLine5);
        cboLine5.SetText(s.cpLine5Cls);
        txtLine6.SetText(s.cpLine6);
        cboLine6.SetText(s.cpLine6Cls);
        txtLine7.SetText(s.cpLine7);
        cboLine7.SetText(s.cpLine7Cls);
        txtLine8.SetText(s.cpLine8);
        cboLine8.SetText(s.cpLine8Cls);
 		

if (s.cpLine1Cls == 'TEXT') {
        txtLine1.SetEnabled(true);

    } else { 
        txtLine1.SetEnabled(false);
        txtLine1.SetText('');
    }
    lblInfo.SetText('');

    if (s.cpLine2Cls == 'TEXT') {
        txtLine2.SetEnabled(true);

    } else { 
        txtLine2.SetEnabled(false);
        txtLine2.SetText('');
    }
    lblInfo.SetText('');

    if (s.cpLine3Cls == 'TEXT') {
        txtLine3.SetEnabled(true);

    } else { 
        txtLine3.SetEnabled(false);
        txtLine3.SetText('');
    }
    lblInfo.SetText('');

    if (s.cpLine4Cls == 'TEXT') {
        txtLine4.SetEnabled(true);

    } else { 
        txtLine4.SetEnabled(false);
        txtLine4.SetText('');
    }
    lblInfo.SetText('');

    if (s.cpLine5Cls == 'TEXT') {
        txtLine5.SetEnabled(true);

    } else { 
        txtLine5.SetEnabled(false);
        txtLine5.SetText('');
    }
    lblInfo.SetText('');

    if (s.cpLine6Cls == 'TEXT') {
        txtLine6.SetEnabled(true);

    } else { 
        txtLine6.SetEnabled(false);
        txtLine6.SetText('');
    }
    lblInfo.SetText('');

    if (s.cpLine7Cls == 'TEXT') {
        txtLine7.SetEnabled(true);

    } else { 
        txtLine7.SetEnabled(false);
        txtLine7.SetText('');
    }
    lblInfo.SetText('');

    if (s.cpLine8Cls == 'TEXT') {
        txtLine8.SetEnabled(true);

    } else { 
        txtLine8.SetEnabled(false);
        txtLine8.SetText('');
    }
    lblInfo.SetText('');

          }" />

    </dx:ASPxCallback>
</asp:Content>
