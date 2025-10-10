<%@ Page Title="" Language="vb" AutoEventWireup="false" MasterPageFile="~/SiteContent.Master" CodeBehind="AffiliateMasterDetail.aspx.vb" Inherits="PASISystem.AffiliateMasterDetail" %>

<%@ Register Assembly="DevExpress.Web.v14.1, Version=14.1.6.0, Culture=neutral, PublicKeyToken=b88d1754d700e49a" Namespace="DevExpress.Web.ASPxCallback" TagPrefix="dx" %>
<%@ Register Assembly="DevExpress.Web.v14.1, Version=14.1.6.0, Culture=neutral, PublicKeyToken=b88d1754d700e49a" Namespace="DevExpress.Web.ASPxGridView.Export" TagPrefix="dx" %>
<%@ Register Assembly="DevExpress.Web.v14.1, Version=14.1.6.0, Culture=neutral, PublicKeyToken=b88d1754d700e49a" Namespace="DevExpress.Web.ASPxGridView" TagPrefix="dx" %>
<%@ Register assembly="DevExpress.Web.v14.1, Version=14.1.6.0, Culture=neutral, PublicKeyToken=b88d1754d700e49a" namespace="DevExpress.Web.ASPxEditors" tagprefix="dx" %>
<%@ Register Assembly="DevExpress.Web.v14.1, Version=14.1.6.0, Culture=neutral, PublicKeyToken=b88d1754d700e49a" Namespace="DevExpress.Web.ASPxPanel" TagPrefix="dx" %>
<%@ Register Assembly="DevExpress.Web.v14.1, Version=14.1.6.0, Culture=neutral, PublicKeyToken=b88d1754d700e49a" Namespace="DevExpress.Web.ASPxPopupControl" TagPrefix="dx" %>
<%@ Register Assembly="DevExpress.Web.v14.1, Version=14.1.6.0, Culture=neutral, PublicKeyToken=b88d1754d700e49a" namespace="DevExpress.Web.ASPxGlobalEvents" TagPrefix="dx" %>
<%@ Register assembly="DevExpress.Web.v14.1, Version=14.1.6.0, Culture=neutral, PublicKeyToken=b88d1754d700e49a" namespace="DevExpress.Web.ASPxHiddenField" tagprefix="dx1" %>
<asp:Content ID="Content1" ContentPlaceHolderID="head" runat="server">
<style type="text/css">
.dxeHLC, .dxeHC, .dxeHFC
{
display: none;
}
</style> 

<script language="javascript" type="text/javascript">

    function singlequote(e) {
        var unicode = e.charCode ? e.charCode : e.keyCode
        if (unicode == 39) {
            return false //disable key press
        }
    }

    function UpperCase(e) {
        e.target.value = e.target.value.toUpperCase();
    }

    function numbersonly(e) {
        var unicode = e.charCode ? e.charCode : e.keyCode
        if (unicode != 8) { //if the key isn't the backspace key (which we should allow)
            if (unicode < 45 || unicode > 57) //if not a number
                return false //disable key press
        }
    }

    function clear() {
        txtAffiliateID.SetText('');
        txtAffiliateName.SetText('');
        txtConsigneeCode.SetText('');
        txtConsigneeAddress.SetText('');
        txtBuyerCode.SetText('');
        txtBuyerAddress.SetText('');
        txtAddress.SetText('');
        txtCity.SetText('');
        txtPostalCode.SetText('');
        txtPhone1.SetText('');
        txtPhone2.SetText('');
        txtFax.SetText('');
        txtNPWP.SetText('');
        txtPath.SetText('');

        txtAtt.SetText('');
        txtPaymentTerm.SetText('');
        txtPOCode.SetText('');
        txtAffCode.SetText('');

        rdrPASI.SetChecked(true);
        rdrYes.SetChecked(true);
        rdrAff1.SetChecked(true);

        txtAffiliateID.GetInputElement().setAttribute('style', 'background:#FFFFFF;foreground:#FFFFFF;');
        txtAffiliateID.GetInputElement().readOnly = false;
    }

    function clear2() {
        txtAffiliateID.GetInputElement().setAttribute('style', 'background:#CCCCCC;foreground:#CCCCCC;');
        txtAffiliateID.GetInputElement().readOnly = true;
    }

    function validasubmit() {
        lblInfo.GetMainElement().style.color = 'Red';
        if (txtAffiliateID.GetText() == "") {
            lblInfo.SetText("[6011] Please Input Affiliate ID first!");
            txtAffiliateID.Focus();
            e.ProcessOnServer = false;
            return false;
        }        

        if (txtAffiliateName.GetText() == "") {
            lblInfo.SetText("[6011] Please Input Affiliate Name first!");
            txtAffiliateName.Focus();
            e.ProcessOnServer = false;
            return false;
        }
        
        if (txtAddress.GetText() == "" && txtAddress.GetText() == "") {
            lblInfo.SetText("[6011] Please Input Address first!");
            txtAddress.Focus();
            e.ProcessOnServer = false;
            return false;
        }
        
        if (txtCity.GetText() == "") {
            lblInfo.SetText("[6011] Please Input City first!");
            txtCity.Focus();
            e.ProcessOnServer = false;
            return false;
        }
        
        if (txtPostalCode.GetText() == "") {
            lblInfo.SetText("[6011] Please Input Postal Code first!");
            txtPostalCode.Focus();
            e.ProcessOnServer = false;
            return false;
        }
        
        if (txtPhone1.GetText() == "") {
            lblInfo.SetText("[6011] Please Input Phone first!");
            txtPhone1.Focus();
            e.ProcessOnServer = false;
            return false;
        }

        if (txtPOCode.GetText() == "") {
            lblInfo.SetText("[6011] Please Input PO Code (Invoice Code) first!");
            txtPOCode.Focus();
            e.ProcessOnServer = false;
            return false;
        }
    }

    function up_Insert() {
        var pIsUpdate = '';
        var pAffiliateID = txtAffiliateID.GetText();
        var pAffiliateName = txtAffiliateName.GetText();
        var pAddress = txtAddress.GetText();
        var pCity = txtCity.GetValue();
        var pPostalCode = txtPostalCode.GetValue();
        var pPhone1 = txtPhone1.GetValue();
        var pPhone2 = txtPhone2.GetValue();
        var pFax = txtFax.GetValue();
        var pNPWP = txtNPWP.GetValue();
        var pPort = txtPort.GetValue();
        var pPortAir = txtPortAir.GetValue();
        var pAffCode = txtAffCode.GetValue();
        var pPODel = '';
        var pOverseasCls = '';
        var pAffiliateCls = '';

        if (rdrPASI.GetValue() == true) {
            pPODel = '1';
        } else {
            pPODel = '0';
        }

        if (rdrYes.GetValue() == true) {
            pOverseasCls = '1';
        } else {
            pOverseasCls = '0';
        }

        if (rdrAff1.GetValue() == true) {
            pAffiliateCls = 'A';
        } else {
            pAffiliateCls = 'B';
        }

        AffiliateSubmit.PerformCallback('save|' + pIsUpdate + '|' + pAffiliateID + '|' + pAffiliateName + '|' + pAddress + '|' + pCity + '|' + pPostalCode + '|' + pPhone1 + '|' + pPhone2 + '|' + pFax + '|' + pNPWP + '|' + pPODel + '|' + pOverseasCls + '|' + pPort + '|' + pPortAir + '|' + pAffiliateCls + '|' + pAffCode);
    }

    function up_delete() {
        if (txtAffiliateID.GetText() == "") {
            lblInfo.GetMainElement().style.color = 'Red';
            lblInfo.SetText("[6011] Please select the data first!");
            e.ProcessOnServer = false;
            return false;
        }
        
        var msg = confirm('Are you sure want to delete this data ?');
        if (msg == false) {
            e.processOnServer = false;
            return;
        }

        var pGroupCode = txtAffiliateID.GetText();
        AffiliateSubmit.PerformCallback('delete|' + pGroupCode);
    }

    function readonly() {
        txtAffiliateID.GetInputElement().setAttribute('style', 'background:#CCCCCC;foreground:#CCCCCC;');
        txtAffiliateID.GetInputElement().readOnly = true;
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
    <table style="border-width: thin; border-style: inset hidden ridge hidden; border-color:#9598A1; width:100%; height: 15px;">
        <tr>
            <td colspan="8" height="15">
                <%--error message--%>
                <table id="info" style="width:100%; height: 15px;">
                    <tr>
                        <td align="left" valign="middle" style="height:15px;">
                            <dx:ASPxLabel ID="lblInfo" runat="server" Text="[lblinfo]" Font-Names="Tahoma"
                                ClientInstanceName="lblInfo" Font-Bold="True" Font-Italic="True" Font-Size="8pt" >
                            </dx:ASPxLabel>
                        </td>
                    </tr>         
                </table>
            </td>            
        </tr>
    </table>

    <div style="height:5px;"></div> 

    <table style="border-width: 1pt thin thin thin; border-style: ridge; border-color:#9598A1; width:100%; height: 470px;">
        <tr>
            <td>
                <table style="width:100%;">
                    <tr>
                        <td align="left" width="50px">&nbsp;</td>
                        <td align="left" width="150px">
                            <dx:ASPxLabel ID="ASPxLabel10" runat="server" Font-Names="Tahoma" 
                                Font-Size="8pt" Text="AFFILIATE ID">
                            </dx:ASPxLabel>
                        </td>
                        <td align="left">
                            <dx:ASPxTextBox ID="txtAffiliateID" runat="server" Width="300px" 
                                ClientInstanceName="txtAffiliateID" Font-Names="Tahoma" 
                                MaxLength="10" onkeypress="return singlequote(event)" Height="25px">
                                <ClientSideEvents LostFocus="function(s, e) { lblInfo.SetText(&quot;&quot;); }" 
                                    KeyDown=" function(s, e) {
                                        if(ASPxClientUtils.GetKeyCode(e.htmlEvent) ===  ASPxKey.Enter){
                                            lblInfo.SetText('');
                                            AffiliateSubmit.PerformCallback('load');                                            
                                        }
                                    }" 
                                />
                            </dx:ASPxTextBox>
                        </td>
                        <td align="left" width="50px">&nbsp;</td>
                        <td align="left" width="150px">
                            <dx:ASPxLabel ID="ASPxLabel6" runat="server" Font-Names="Tahoma" 
                                Font-Size="8pt" Text="CONSIGNEE CODE">
                            </dx:ASPxLabel>
                        </td>
                        <td align="left">
                            <dx:ASPxTextBox ID="txtConsigneeCode" runat="server" Width="300px" 
                                ClientInstanceName="txtConsigneeCode" Font-Names="Tahoma" 
                                MaxLength="10" onkeypress="return singlequote(event)" Height="25px">
                                <ClientSideEvents LostFocus="function(s, e) { lblInfo.SetText(&quot;&quot;); }" />
                            </dx:ASPxTextBox>
                        </td>
                    </tr>
                    <tr>
                        <td align="left" width="50px">&nbsp;</td>
                        <td align="left" width="150px">
                            <dx:ASPxLabel ID="ASPxLabel11" runat="server" Font-Names="Tahoma" 
                                Font-Size="8pt" Text="AFFILIATE NAME">
                            </dx:ASPxLabel>
                        </td>
                        <td align="left">
                            <dx:ASPxTextBox ID="txtAffiliateName" runat="server" Width="300px" 
                                ClientInstanceName="txtAffiliateName" Font-Names="Tahoma" 
                                MaxLength="50" onkeypress="return singlequote(event)" Height="25px">
                                <ClientSideEvents LostFocus="function(s, e) { lblInfo.SetText(&quot;&quot;); }" />
                            </dx:ASPxTextBox>
                        </td>
                        <td align="left" width="50px">&nbsp;</td>
                        <td align="left" width="150px">
                            <dx:ASPxLabel ID="ASPxLabel9" runat="server" Font-Names="Tahoma" 
                                Font-Size="8pt" Text="CONSIGNEE NAME">
                            </dx:ASPxLabel>
                        </td>
                        <td align="left">
                            <dx:ASPxTextBox ID="txtConsigneeName" runat="server" Width="300px" 
                                ClientInstanceName="txtConsigneeName" Font-Names="Tahoma" 
                                MaxLength="50" onkeypress="return singlequote(event)" Height="25px">
                                <ClientSideEvents LostFocus="function(s, e) { lblInfo.SetText(&quot;&quot;); }" />
                            </dx:ASPxTextBox>
                        </td>
                    </tr>
                    <tr>
                        <td align="left" width="50px">&nbsp;</td>
                        <td align="left" width="150px">
                            <dx:ASPxLabel ID="ASPxLabel12" runat="server" Font-Names="Tahoma" 
                                Font-Size="8pt" Text="ADDRESS">
                            </dx:ASPxLabel>
                        </td>
                        <td align="left">
                            <dx:ASPxMemo  ID="txtAddress" runat="server" Width="300px" Height="100px"
                                ClientInstanceName="txtAddress" Font-Names="Tahoma" 
                                MaxLength="500" onkeypress="return singlequote(event)">
                                <ClientSideEvents LostFocus="function(s, e) { lblInfo.SetText(&quot;&quot;); }" />
                            </dx:ASPxMemo>
                        </td>
                        <td align="left" width="50px">&nbsp;</td>
                        <td align="left" width="150px">
                            <dx:ASPxLabel ID="ASPxLabel5" runat="server" Font-Names="Tahoma" 
                                Font-Size="8pt" Text="CONSIGNEE ADDRESS">
                            </dx:ASPxLabel>
                        </td>
                        <td align="left">
                            <dx:ASPxMemo  ID="txtConsigneeAddress" runat="server" Width="300px" Height="100px"
                                ClientInstanceName="txtConsigneeAddress" Font-Names="Tahoma" 
                                MaxLength="500" onkeypress="return singlequote(event)">
                                <ClientSideEvents LostFocus="function(s, e) { lblInfo.SetText(&quot;&quot;); }" />
                            </dx:ASPxMemo>
                        </td>
                    </tr>
                    <tr>
                        <td align="left" width="50px">&nbsp;</td>
                        <td align="left" width="150px">
                            <dx:ASPxLabel ID="ASPxLabel13" runat="server" Font-Names="Tahoma" 
                                Font-Size="8pt" Text="CITY">
                            </dx:ASPxLabel>
                        </td>
                        <td align="left">
                            <dx:ASPxTextBox ID="txtCity" runat="server" Width="300px" 
                                ClientInstanceName="txtCity" Font-Names="Tahoma" 
                                MaxLength="20" onkeypress="return singlequote(event)" Height="25px">
                                <ClientSideEvents LostFocus="function(s, e) { lblInfo.SetText(&quot;&quot;); }" />
                            </dx:ASPxTextBox>
                        </td>
                        <td align="left" width="50px">&nbsp;</td>
                        <td align="left" width="150px">
                            <dx:ASPxLabel ID="ASPxLabel7" runat="server" Font-Names="Tahoma" 
                                Font-Size="8pt" Text="BUYER CODE">
                            </dx:ASPxLabel>
                        </td>
                        <td align="left">
                            <dx:ASPxTextBox ID="txtBuyerCode" runat="server" Width="300px" 
                                ClientInstanceName="txtBuyerCode" Font-Names="Tahoma" 
                                MaxLength="10" onkeypress="return singlequote(event)" Height="25px">
                                <ClientSideEvents LostFocus="function(s, e) { lblInfo.SetText(&quot;&quot;); }" />
                            </dx:ASPxTextBox>
                        </td>
                    </tr>
                    <tr>
                        <td align="left" width="50px">&nbsp;</td>
                        <td align="left" width="150px">
                            <dx:ASPxLabel ID="ASPxLabel14" runat="server" Font-Names="Tahoma" 
                                Font-Size="8pt" Text="POSTAL CODE">
                            </dx:ASPxLabel>
                        </td>
                        <td align="left">
                            <dx:ASPxTextBox ID="txtPostalCode" runat="server" Width="300px" 
                                ClientInstanceName="txtPostalCode" Font-Names="Tahoma" 
                                MaxLength="10" onkeypress="return numbersonly(event)" Height="25px">
                                <ClientSideEvents LostFocus="function(s, e) { lblInfo.SetText(&quot;&quot;); }" />
                            </dx:ASPxTextBox>
                        </td>
                        <td align="left" width="50px">&nbsp;</td>
                        <td align="left" width="150px" >
                            <dx:ASPxLabel ID="ASPxLabel20" runat="server" Font-Names="Tahoma" 
                                Font-Size="8pt" Text="BUYER NAME">
                            </dx:ASPxLabel>
                        </td>
                        <td align="left">
                            <dx:ASPxTextBox ID="txtBuyerName" runat="server" Width="300px" 
                                ClientInstanceName="txtBuyerName" Font-Names="Tahoma" 
                                MaxLength="50" onkeypress="return singlequote(event)" Height="25px">
                                <ClientSideEvents LostFocus="function(s, e) { lblInfo.SetText(&quot;&quot;); }" />
                            </dx:ASPxTextBox>
                        </td>
                    </tr>
                    <tr>
                        <td align="left" width="50px">&nbsp;</td>
                        <td align="left" width="150px">
                            <dx:ASPxLabel ID="ASPxLabel15" runat="server" Font-Names="Tahoma" 
                                Font-Size="8pt" Text="PHONE 1">
                            </dx:ASPxLabel>
                        </td>
                        <td align="left">
                            <dx:ASPxTextBox ID="txtPhone1" runat="server" Width="300px" 
                                ClientInstanceName="txtPhone1" Font-Names="Tahoma" 
                                MaxLength="15" onkeypress="return numbersonly(event)" Height="25px">
                                <ClientSideEvents LostFocus="function(s, e) { lblInfo.SetText(&quot;&quot;); }" />
                            </dx:ASPxTextBox>
                        </td>
                        <td align="left" width="50px">&nbsp;</td>
                        <td align="left" width="150px" rowspan="4">
                            <dx:ASPxLabel ID="ASPxLabel8" runat="server" Font-Names="Tahoma" 
                                Font-Size="8pt" Text="BUYER ADDRESS">
                            </dx:ASPxLabel>
                        </td>
                        <td align="left" rowspan="4">
                            <dx:ASPxMemo  ID="txtBuyerAddress" runat="server" Width="300px" Height="100px"
                                ClientInstanceName="txtBuyerAddress" Font-Names="Tahoma" 
                                MaxLength="500" onkeypress="return singlequote(event)">
                                <ClientSideEvents LostFocus="function(s, e) { lblInfo.SetText(&quot;&quot;); }" />
                            </dx:ASPxMemo>
                        </td>
                    </tr>
                    <tr>
                        <td align="left" width="50px">&nbsp;</td>
                        <td align="left" width="150px">
                            <dx:ASPxLabel ID="ASPxLabel16" runat="server" Font-Names="Tahoma" 
                                Font-Size="8pt" Text="PHONE 2">
                            </dx:ASPxLabel>
                        </td>
                        <td align="left">
                            <dx:ASPxTextBox ID="txtPhone2" runat="server" Width="300px" 
                                ClientInstanceName="txtPhone2" Font-Names="Tahoma" 
                                MaxLength="15" onkeypress="return numbersonly(event)" Height="25px">
                                <ClientSideEvents LostFocus="function(s, e) { lblInfo.SetText(&quot;&quot;); }" />
                            </dx:ASPxTextBox>
                        </td>                        
                    </tr>
                    <tr>
                        <td align="left" width="50px">&nbsp;</td>
                        <td align="left" width="150px">
                            <dx:ASPxLabel ID="ASPxLabel17" runat="server" Font-Names="Tahoma" 
                                Font-Size="8pt" Text="FAX">
                            </dx:ASPxLabel>
                        </td>
                        <td align="left">
                            <dx:ASPxTextBox ID="txtFax" runat="server" Width="300px" 
                                ClientInstanceName="txtFax" Font-Names="Tahoma" 
                                MaxLength="15" onkeypress="return numbersonly(event)" Height="25px">
                                <ClientSideEvents LostFocus="function(s, e) { lblInfo.SetText(&quot;&quot;); }" />
                            </dx:ASPxTextBox>
                        </td>
                    </tr>
                    <tr>
                        <td align="left" width="50px">&nbsp;</td>
                        <td align="left" width="150px">
                            <dx:ASPxLabel ID="ASPxLabel18" runat="server" Font-Names="Tahoma" 
                                Font-Size="8pt" Text="NPWP">
                            </dx:ASPxLabel>
                        </td>
                        <td align="left">
                            <dx:ASPxTextBox ID="txtNPWP" runat="server" Width="300px" 
                                ClientInstanceName="txtNPWP" Font-Names="Tahoma" 
                                MaxLength="25" onkeypress="return numbersonly(event)" Height="25px">
                                <ClientSideEvents LostFocus="function(s, e) { lblInfo.SetText(&quot;&quot;); }" />
                            </dx:ASPxTextBox>
                        </td>
                    </tr>
                    <tr>
                        <td align="left" width="50px">&nbsp;</td>
                        <td align="left" width="150px">
                            <dx:ASPxLabel ID="ASPxLabel2" runat="server" Font-Names="Tahoma" 
                                Font-Size="8pt" Text="KANTOR PABEAN">
                            </dx:ASPxLabel>
                        </td>
                        <td align="left">
                            <dx:ASPxTextBox ID="txtKantorPabean" runat="server" Width="300px" 
                                ClientInstanceName="txtKantorPabean" Font-Names="Tahoma" 
                                MaxLength="60" onkeypress="return singlequote(event)" Height="25px">
                                <ClientSideEvents LostFocus="function(s, e) { lblInfo.SetText(&quot;&quot;); }" />
                            </dx:ASPxTextBox>
                        </td>
                        <td align="left" width="50px">&nbsp;</td>
                        <td align="left" width="150px" >
                            <dx:ASPxLabel ID="ASPxLabel21" runat="server" Font-Names="Tahoma" 
                                Font-Size="8pt" Text="DESTINATION PORT (BOAT)">
                            </dx:ASPxLabel>
                        </td>
                        <td align="left">
                            <dx:ASPxTextBox ID="txtPort" runat="server" Width="300px" 
                                ClientInstanceName="txtPort" Font-Names="Tahoma" 
                                MaxLength="25" onkeypress="return singlequote(event)" Height="25px">
                                <ClientSideEvents LostFocus="function(s, e) { lblInfo.SetText(&quot;&quot;); }" />
                            </dx:ASPxTextBox>
                        </td>
                    </tr>
                    <tr>
                        <td align="left" width="50px">&nbsp;</td>
                        <td align="left" width="150px">
                            <dx:ASPxLabel ID="ASPxLabel3" runat="server" Font-Names="Tahoma" 
                                Font-Size="8pt" Text="IZIN TPB">
                            </dx:ASPxLabel>
                        </td>
                        <td align="left">
                            <dx:ASPxTextBox ID="txtIzinTPB" runat="server" Width="300px" 
                                ClientInstanceName="txtIzinTPB" Font-Names="Tahoma" 
                                MaxLength="60" onkeypress="return singlequote(event)" Height="25px">
                                <ClientSideEvents LostFocus="function(s, e) { lblInfo.SetText(&quot;&quot;); }" />
                            </dx:ASPxTextBox>
                        </td>
                        <td align="left" width="50px">&nbsp;</td>
                        <td align="left" width="150px" >
                            <dx:ASPxLabel ID="ASPxLabel23" runat="server" Font-Names="Tahoma" 
                                Font-Size="8pt" Text="ATTENTION">
                            </dx:ASPxLabel>
                        </td>
                        <td align="left">
                            <dx:ASPxTextBox ID="txtAtt" runat="server" Width="300px" 
                                ClientInstanceName="txtAtt" Font-Names="Tahoma" 
                                MaxLength="100" onkeypress="return singlequote(event)" Height="25px">
                                <ClientSideEvents LostFocus="function(s, e) { lblInfo.SetText(&quot;&quot;); }" />
                            </dx:ASPxTextBox>
                        </td>
                    </tr>
                    <tr>
                        <td align="left" width="50px">&nbsp;</td>
                        <td align="left" width="150px">
                            <dx:ASPxLabel ID="ASPxLabel4" runat="server" Font-Names="Tahoma" 
                                Font-Size="8pt" Text="BC PERSON">
                            </dx:ASPxLabel>
                        </td>
                        <td align="left">
                            <dx:ASPxTextBox ID="txtBCPerson" runat="server" Width="300px" 
                                ClientInstanceName="txtBCPerson" Font-Names="Tahoma" 
                                MaxLength="60" onkeypress="return singlequote(event)" Height="25px">
                                <ClientSideEvents LostFocus="function(s, e) { lblInfo.SetText(&quot;&quot;); }" />
                            </dx:ASPxTextBox>
                        </td>
                        <td align="left" width="50px">&nbsp;</td>
                        <td align="left" width="150px" >
                            <dx:ASPxLabel ID="ASPxLabel24" runat="server" Font-Names="Tahoma" 
                                Font-Size="8pt" Text="PAYMENT TERM">
                            </dx:ASPxLabel>
                        </td>
                        <td align="left">
                            <dx:ASPxTextBox ID="txtPaymentTerm" runat="server" Width="300px" 
                                ClientInstanceName="txtPaymentTerm" Font-Names="Tahoma" 
                                MaxLength="50" onkeypress="return singlequote(event)" Height="25px">
                                <ClientSideEvents LostFocus="function(s, e) { lblInfo.SetText(&quot;&quot;); }" />
                            </dx:ASPxTextBox>
                        </td>
                    </tr>
                    <tr>
                        <td align="left" width="50px">&nbsp;</td>
                        <td align="left" width="150px">
                            <dx:ASPxLabel ID="ASPxLabel19" runat="server" Font-Names="Tahoma" 
                                Font-Size="8pt" Text="PO DELIVERY BY">
                            </dx:ASPxLabel>
                        </td>
                        <td align="left">
                            <table>
                                <tr>
                                    <td>
                                        <dx:ASPxRadioButton ID="rdrPASI" ClientInstanceName="rdrPASI" runat="server" Text="PASI" GroupName="pasi">
                                            <ClientSideEvents LostFocus="function(s, e) { lblInfo.SetText(&quot;&quot;); }" />
                                        </dx:ASPxRadioButton>
                                    </td>
                                    <td>
                                        <dx:ASPxRadioButton ID="rdrSupplier" ClientInstanceName="rdrSupplier" runat="server" Text="SUPPLIER" GroupName="pasi">
                                            <ClientSideEvents LostFocus="function(s, e) { lblInfo.SetText(&quot;&quot;); }" />
                                        </dx:ASPxRadioButton>
                                    </td>
                                </tr>
                            </table>                                                        
                        </td>
                        <td align="left" width="50px">&nbsp;</td>
                        <td align="left" width="150px" >
                            <dx:ASPxLabel ID="ASPxLabel25" runat="server" Font-Names="Tahoma" 
                                Font-Size="8pt" Text="PO CODE (INVOICE)">
                            </dx:ASPxLabel>
                        </td>
                        <td align="left">
                            <dx:ASPxTextBox ID="txtPOCode" runat="server" Width="300px" 
                                ClientInstanceName="txtPOCode" Font-Names="Tahoma" 
                                MaxLength="2" onkeypress="return singlequote(event)" Height="25px">
                                <ClientSideEvents LostFocus="function(s, e) { lblInfo.SetText(&quot;&quot;); }" />
                            </dx:ASPxTextBox>
                        </td>
                    </tr>
                    <tr>
                        <td align="left" width="50px">&nbsp;</td>
                        <td align="left" width="150px">
                            <dx:ASPxLabel ID="ASPxLabel22" runat="server" Font-Names="Tahoma" 
                                Font-Size="8pt" Text="OVERSEAS AFFILIATE">
                            </dx:ASPxLabel>
                        </td>
                        <td align="left">
                            <table>
                                <tr>
                                    <td>
                                        <dx:ASPxRadioButton ID="rdrYes" ClientInstanceName="rdrYes" runat="server" Text="YES" GroupName="pasicls">
                                            <ClientSideEvents LostFocus="function(s, e) { lblInfo.SetText(&quot;&quot;); }" />
                                        </dx:ASPxRadioButton>
                                    </td>
                                    <td>
                                        <dx:ASPxRadioButton ID="rdrNo" ClientInstanceName="rdrNo" runat="server" Text="NO" GroupName="pasicls">
                                            <ClientSideEvents LostFocus="function(s, e) { lblInfo.SetText(&quot;&quot;); }" />
                                        </dx:ASPxRadioButton>
                                    </td>
                                </tr>
                            </table>                                                        
                        </td>
                        <td align="left" width="50px">&nbsp;</td>
                        <td align="left" width="150px">
                            <dx:ASPxLabel ID="ASPxLabel1" runat="server" Font-Names="Tahoma" 
                                Font-Size="8pt" Text="PATH OES FOLDER">
                            </dx:ASPxLabel>
                        </td>
                        <td align="left">
                            <dx:ASPxTextBox ID="txtPath" runat="server" Width="300px" 
                                ClientInstanceName="txtPath" Font-Names="Tahoma" 
                                MaxLength="300" onkeypress="return singlequote(event)" Height="25px">
                                <ClientSideEvents LostFocus="function(s, e) { lblInfo.SetText(&quot;&quot;); }" />
                            </dx:ASPxTextBox>
                        </td>
                    </tr>
                    <tr>
                        <td align="left" width="50px">&nbsp;</td>
                        <td align="left" width="150px">
                            <dx:ASPxLabel ID="ASPxLabel26" runat="server" Font-Names="Tahoma" 
                                Font-Size="8pt" Text="AFFILIATE CLASSIFICATION">
                            </dx:ASPxLabel>
                        </td>
                        <td align="left">
                            <table>
                                <tr>
                                    <td>
                                        <dx:ASPxRadioButton ID="rdrAff1" ClientInstanceName="rdrAff1" runat="server" Text="AFFILIATE" GroupName="pasiaff">
                                            <ClientSideEvents LostFocus="function(s, e) { lblInfo.SetText(&quot;&quot;); }" />
                                        </dx:ASPxRadioButton>
                                    </td>
                                    <td>
                                        <dx:ASPxRadioButton ID="rdrAff2" ClientInstanceName="rdrAff2" runat="server" Text="NON AFFILIATE" GroupName="pasiaff">
                                            <ClientSideEvents LostFocus="function(s, e) { lblInfo.SetText(&quot;&quot;); }" />
                                        </dx:ASPxRadioButton>
                                    </td>                                    
                                </tr>
                            </table>                                                        
                        </td> 
                        <td align="left" width="50px">&nbsp;</td>
                        <td align="left" width="150px" >
                            <dx:ASPxLabel ID="ASPxLabel27" runat="server" Font-Names="Tahoma" 
                                Font-Size="8pt" Text="DESTINATION PORT (AIR)">
                            </dx:ASPxLabel>
                        </td>
                        <td align="left">
                            <dx:ASPxTextBox ID="txtPortAir" runat="server" Width="300px" 
                                ClientInstanceName="txtPortAir" Font-Names="Tahoma" 
                                MaxLength="25" onkeypress="return singlequote(event)" Height="25px">
                                <ClientSideEvents LostFocus="function(s, e) { lblInfo.SetText(&quot;&quot;); }" />
                            </dx:ASPxTextBox>
                        </td>
                    </tr>
                    <tr>
                        <td align="left" width="50px">&nbsp;</td>
                        <td align="left" width="150px">
                            
                        </td>
                        <td align="left">
                            
                        </td>
                        <td align="left" width="50px">&nbsp;</td>
                        <td align="left" width="150px" >
                            <dx:ASPxLabel ID="ASPxLabel29" runat="server" Font-Names="Tahoma" 
                                Font-Size="8pt" Text="Affiliate Code">
                            </dx:ASPxLabel>
                        </td>
                        <td align="left">
                            <dx:ASPxTextBox ID="txtAffCode" runat="server" Width="300px" 
                                ClientInstanceName="txtAffCode" Font-Names="Tahoma" 
                                MaxLength="4" onkeypress="return singlequote(event)" Height="25px" oninput="return UpperCase(event)">
                                <ClientSideEvents LostFocus="function(s, e) { lblInfo.SetText(&quot;&quot;); }" />
                            </dx:ASPxTextBox>
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
                    Width="90px" Font-Size="8pt" UseSubmitBehavior="False">
                </dx:ASPxButton>   
                </td>                     
            
            <td valign="top" align="right" style="width: 50px;">                                  
            </td>
            <td valign="top" align="right" style="width: 50px;">                                   
                <dx:ASPxButton ID="btnClear" runat="server" Text="CLEAR"                    
                    Font-Names="Tahoma"
                    Width="90px" AutoPostBack="False" Font-Size="8pt" 
                    UseSubmitBehavior="False">                   
                </dx:ASPxButton>
            </td>
            <td align="right" style="width:80px;">                                   
                <dx:ASPxButton ID="btnDelete" runat="server" Text="DELETE"                              
                    Font-Names="Tahoma" Width="80px" AutoPostBack="False" Font-Size="8pt" 
                    UseSubmitBehavior="False">
                    <ClientSideEvents Click="function(s, e) {
                        up_delete();
                    }" />
                </dx:ASPxButton>
            </td>
            <td valign="top" align="right" style="width: 50px;">                                   
                <dx:ASPxButton ID="btnSubmit" runat="server" Text="SAVE"                    
                    Font-Names="Tahoma"
                    Width="90px" AutoPostBack="False" Font-Size="8pt" 
                    UseSubmitBehavior="False">
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
            lblInfo.SetText('');   
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
                    clear2();
                }
            } else {
                lblInfo.SetText('');
            }  

            delete s.cpMessage;

            if (s.cpKeyPress == 'ON')
            {
                txtAffiliateID.SetText(s.cpAffiliateID);
                txtConsigneeCode.SetText(s.cpConsigneeCode);
                txtConsigneeAddress.SetText(s.cpConsigneeAddress);
                txtBuyerCode.SetText(s.cpBuyerCode);
                txtBuyerAddress.SetText(s.cpBuyerAddress);
                txtAffiliateName.SetText(s.cpAffiliateName);
                txtAddress.SetText(s.cpAddress);
                txtCity.SetText(s.cpCity);
                txtPostalCode.SetText(s.cpPostalCode);
                txtPhone1.SetText(s.cpPhone1);
                txtPhone2.SetText(s.cpPhone2);
                txtFax.SetText(s.cpFax);
                txtNPWP.SetText(s.cpNPWP);
                txtKantorPabean.SetText(s.cpKantorPabean);
                txtIzinTPB.SetText(s.cpIzinTPB);
                txtBCPerson.SetText(s.cpBCPerson);
                if (s.cpPODeliveryBy == '1') {
                    rdrPASI.SetChecked(true);
                }else {
                    rdrSupplier.SetChecked(true);
                }
                if (s.cpOverseasCls == '1') {
                    rdrYes.SetChecked(true);
                }else {
                    rdrNo.SetChecked(true);
                }

                if (s.cpAffiliateCls == '1') {
                    rdrAff1.SetChecked(true);
                }else {
                    rdrAff2.SetChecked(true);
                }

                txtPaymentTerm.SetText(s.cpPaymentTerm);
                txtPOCode.SetText(s.cpPOCode);
                txtAtt.SetText(s.cpAtt);

                txtPath.SetText(s.cpFolderOES);
                txtPort.SetText(s.cpPort);
                txtPortAir.SetText(s.cpPortAir);
                txtAffiliateID.GetInputElement().setAttribute('style', 'background:#CCCCCC;foreground:#CCCCCC;');
                txtAffiliateID.GetInputElement().readOnly = true;

                delete s.cpKeyPress
            }
        }" />
    </dx:ASPxCallback>
</asp:Content>
