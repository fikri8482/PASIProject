<%@ Page Title="" Language="vb" AutoEventWireup="false" MasterPageFile="~/SiteContent.Master" CodeBehind="POFinalApproval.aspx.vb" Inherits="AffiliateSystem.POFinalApproval" %>
<%@ Register Assembly="DevExpress.Web.v14.1, Version=14.1.6.0, Culture=neutral, PublicKeyToken=b88d1754d700e49a" namespace="DevExpress.Web.ASPxGridView.Export" TagPrefix="dx" %>
<%@ Register Assembly="DevExpress.Web.v14.1, Version=14.1.6.0, Culture=neutral, PublicKeyToken=b88d1754d700e49a" namespace="DevExpress.Web.ASPxGridView" TagPrefix="dx" %>
<%@ Register assembly="DevExpress.Web.v14.1, Version=14.1.6.0, Culture=neutral, PublicKeyToken=b88d1754d700e49a" namespace="DevExpress.Web.ASPxEditors" tagprefix="dx" %>
<%@ Register Assembly="DevExpress.Web.v14.1, Version=14.1.6.0, Culture=neutral, PublicKeyToken=b88d1754d700e49a" namespace="DevExpress.Web.ASPxGlobalEvents" TagPrefix="dx" %>
<%@ Register Assembly="DevExpress.Web.v14.1, Version=14.1.6.0, Culture=neutral, PublicKeyToken=b88d1754d700e49a" namespace="DevExpress.Web.ASPxPanel" TagPrefix="dx" %>
<%@ Register Assembly="DevExpress.Web.v14.1, Version=14.1.6.0, Culture=neutral, PublicKeyToken=b88d1754d700e49a" namespace="DevExpress.Web.ASPxRoundPanel" TagPrefix="dx" %>
<%@ Register Assembly="DevExpress.Web.v14.1, Version=14.1.6.0, Culture=neutral, PublicKeyToken=b88d1754d700e49a" Namespace="DevExpress.Web.ASPxCallback" TagPrefix="dx" %>

<asp:Content ID="Content1" ContentPlaceHolderID="head" runat="server">
    <style type="text/css">
        .dxeHLC, .dxeHC, .dxeHFC
        {
            display: none;
        }
        .style1
        {
            width: 5px;
            height: 20px;
        }
        .style2
        {
            width: 100px;
            height: 20px;
        }
        .style3
        {
            height: 20px;
        }
    </style>
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
        height = height - (height * 35 / 100)
        grid.SetHeight(height);
    }

    function clear() {
        txtUser1.SetText('');
        txtUser2.SetText('');
        txtUser3.SetText('');
        txtUser4.SetText('');
        txtUser5.SetText('');
        txtUser6.SetText('');
        txtUser7.SetText('');
        txtUser8.SetText('');

        txtDate1.SetText('');
        txtDate2.SetText('');
        txtDate3.SetText('');
        txtDate4.SetText('');
        txtDate5.SetText('');
        txtDate6.SetText('');
        txtDate7.SetText('');
        txtDate8.SetText('');

        txtShip.SetText('');
        txtRemarks.SetText('');
        txtDelivery.SetText('');
        txtCommercial.SetText('');
        txtPOKanban.SetText('');
    }
</script>
</asp:Content>

<asp:Content ID="Content2" ContentPlaceHolderID="Content" runat="server">
    <table style="width:100%;">
        <tr>
            <td width="100%">
                <table style="border-width: 1pt thin thin thin; border-style: ridge; border-color:#9598A1; width:100%; height: 70px;">
                    <tr>
                        <td colspan="8" height="20">
                            <table id="Table1">
                                <tr>
                                    <td style="width:5px;"></td>
                                    <td align="left" valign="middle" height="20px" width="100px">
                                        <dx:ASPxLabel ID="ASPxLabel1" runat="server" Text="PERIOD"
                                            Font-Names="Tahoma" Font-Size="8pt">
                                        </dx:ASPxLabel>
                                    </td>
                                    <td align="left" valign="middle" height="20px" width="180px">
                                        <dx:ASPxTimeEdit ID="dtPeriodFrom" runat="server" ClientInstanceName="dtPeriodFrom" 
                                            DisplayFormatString="MMM yyyy" EditFormat="Custom" 
                                            EditFormatString="MMM yyyy" Font-Names="Tahoma" Font-Size="8pt" Width="180px">
                                            <ClientSideEvents ValueChanged="function(s, e) {	                                            
                                                grid.PerformCallback('kosong');
                                                clear();
                                                cboPartNo.PerformCallback(dtPeriodFrom.GetValue().toString());
	                                            lblInfo.SetText('');                                                
                                            }" />
                                        </dx:ASPxTimeEdit>                                                                                   
                                    </td>
                                    <td style="width:5px;"></td>
                                    <td align="left" valign="middle" style="height:20px; width:100px;">
                                        <dx:ASPxLabel ID="ASPxLabel20" runat="server" Text="SHIP BY"
                                            Font-Names="Tahoma" Font-Size="8pt">
                                        </dx:ASPxLabel>
                                    </td>
                                    <td align="left" valign="middle" height="20px" width="180px">
                                        <dx:ASPxTextBox ID="txtShip" runat="server" Width="180px" 
                                            ClientInstanceName="txtShip" Font-Names="Tahoma" Font-Size="8pt"
                                            MaxLength="25" onkeypress="return singlequote(event)" Height="20px" ReadOnly="True">
                                            <readonlystyle backcolor="#CCCCCC">
                                            </readonlystyle>
                                        </dx:ASPxTextBox>
                                    </td>
                                    <td style="width:5px;"></td>
                                    <td align="left" valign="middle" height="20px" width="135px">
                                        <dx:ASPxLabel ID="ASPxLabel4" runat="server" Text="SUPPLIER REMARKS"
                                            Font-Names="Tahoma" Font-Size="8pt">
                                        </dx:ASPxLabel>
                                    </td>
                                    <td align="left" valign="middle" width="250px" colspan="2" rowspan ="3">
                                        <dx:ASPxMemo ID="txtRemarks" ClientInstanceName="txtRemarks" 
                                            Font-Names="Tahoma" Font-Size="8" MaxLength="200" runat="server" Height="80px" 
                                            Width="250px" ReadOnly="True">
                                            <readonlystyle backcolor="#CCCCCC">
                                            </readonlystyle>
                                        </dx:ASPxMemo>    
                                    </td>                                    

                                    <td align="right" width="180px"></td>
                                </tr>
                                <tr>
                                    <td style="width:5px;"></td>
                                    <td align="left" valign="middle" height="20px" width="100px">
                                        <dx:ASPxLabel ID="ASPxLabel6" runat="server" Text="PO NO."
                                            Font-Names="Tahoma" Font-Size="8pt">
                                        </dx:ASPxLabel>
                                    </td>
                                    <td align="left" valign="middle" height="20px" width="180px">
                                        <dx:ASPxComboBox ID="cboPartNo" runat="server" 
                                            ClientInstanceName="cboPartNo" Width="100%"
                                            Font-Size="8pt" 
                                            Font-Names="Verdana" TextFormatString="{0}">
                                            <ClientSideEvents SelectedIndexChanged="function(s, e) {
                                                grid.PerformCallback('kosong');
                                                clear();
                                                ButtonPartNo.PerformCallback(cboPartNo.GetText());
	                                            lblInfo.SetText('');                                                
                                            }" />
                                            <LoadingPanelStyle ImageSpacing="5px">
                                            </LoadingPanelStyle>
                                        </dx:ASPxComboBox>
                                    </td>
                                    <td style="width:5px;"></td>
                                    <td align="left" valign="middle" style="height:20px; width:100px;">
                                        <dx:ASPxLabel ID="ASPxLabel7" runat="server" Text="PO KANBAN"
                                            Font-Names="Tahoma" Font-Size="8pt">
                                        </dx:ASPxLabel>
                                    </td>
                                    <td align="left" valign="middle" height="20px" width="180px">
                                        <dx:ASPxTextBox ID="txtPOKanban" runat="server" Width="180px" 
                                            ClientInstanceName="txtPOKanban" Font-Names="Tahoma" Font-Size="8pt"
                                            MaxLength="25" onkeypress="return singlequote(event)" Height="20px" ReadOnly="True">
                                            <readonlystyle backcolor="#CCCCCC">
                                            </readonlystyle>
                                        </dx:ASPxTextBox>
                                    </td>
                                    <td style="width:5px;"></td>
                                    <td align="left" valign="middle" height="20px" width="135px">
                                    </td>
                                    
                                    <td align="right" width="180px"></td>
                                </tr> 
                                <tr>
                                    <td style="width:5px;"></td>
                                    <td align="left" valign="middle" height="20px" width="100px">
                                        <dx:ASPxLabel ID="ASPxLabel9" runat="server" Text="DELIVERY"
                                            Font-Names="Tahoma" Font-Size="8pt">
                                        </dx:ASPxLabel>
                                    </td>
                                    <td align="left" valign="middle" height="20px" width="180px">
                                        <dx:ASPxTextBox ID="txtDelivery" runat="server" Width="180px" 
                                            ClientInstanceName="txtDelivery" Font-Names="Tahoma" Font-Size="8pt"
                                            MaxLength="20" onkeypress="return singlequote(event)" Height="20px" ReadOnly="True">
                                            <readonlystyle backcolor="#CCCCCC">
                                            </readonlystyle>
                                        </dx:ASPxTextBox>
                                    </td>
                                    <td style="width:5px;"></td>
                                    <td align="left" valign="middle" style="height:20px; width:100px;">
                                        <dx:ASPxLabel ID="ASPxLabel10" runat="server" Text="DIFFERENCE"
                                            Font-Names="Tahoma" Font-Size="8pt">
                                        </dx:ASPxLabel>
                                    </td>
                                    <td align="left" valign="middle" height="20px" width="180px">
                                        <table>
                                            <tr>
                                                <td>
                                                    <dx:ASPxRadioButton ID="rdrDiff1" ClientInstanceName="rdrDiff1" runat="server" Text="ALL" GroupName="Diff" Font-Names="Tahoma" Font-Size="8pt">
                                                        <ClientSideEvents CheckedChanged="function(s, e) {
                                                        grid.PerformCallback('kosong');
                                                        lblInfo.SetText('');
                                                        clear();
                                                        }" />
                                                    </dx:ASPxRadioButton>
                                                </td>
                                                <td>
                                                    <dx:ASPxRadioButton ID="rdrDiff2" ClientInstanceName="rdrDiff2" runat="server" Text="YES" GroupName="Diff" Font-Names="Tahoma" Font-Size="8pt">
                                                        <ClientSideEvents CheckedChanged="function(s, e) {
                                                        grid.PerformCallback('kosong');
                                                        lblInfo.SetText('');
                                                        clear();
                                                        }" />
                                                    </dx:ASPxRadioButton>
                                                </td>
                                                <td>
                                                    <dx:ASPxRadioButton ID="rdrDiff3" ClientInstanceName="rdrDiff3" runat="server" Text="NO" GroupName="Diff" Font-Names="Tahoma" Font-Size="8pt">
                                                        <ClientSideEvents CheckedChanged="function(s, e) {
                                                        grid.PerformCallback('kosong');
                                                        lblInfo.SetText('');
                                                        clear();
                                                        }" />
                                                    </dx:ASPxRadioButton>
                                                </td>
                                            </tr>
                                        </table> 
                                    </td>

                                    <td style="width:5px;"></td>

                                    <td align="left" valign="middle" height="20px" width="135px">
                                    </td>                                

                                    <td align="right" width="180px"></td>
                                </tr> 
                                <tr>
                                    <td style="width:5px;"></td>
                                    <td align="left" valign="middle" height="20px" width="100px">
                                        <dx:ASPxLabel ID="ASPxLabel5" runat="server" Text="COMMERCIAL"
                                            Font-Names="Tahoma" Font-Size="8pt">
                                        </dx:ASPxLabel>
                                    </td>
                                    <td align="left" valign="middle" height="20px" width="180px">
                                        <dx:ASPxTextBox ID="txtCommercial" runat="server" Width="180px" 
                                            ClientInstanceName="txtCommercial" Font-Names="Tahoma" Font-Size="8pt"
                                            MaxLength="20" onkeypress="return singlequote(event)" Height="20px" ReadOnly="True">
                                            <readonlystyle backcolor="#CCCCCC">
                                            </readonlystyle>
                                        </dx:ASPxTextBox>
                                    </td>
                                    <td style="width:5px;"></td>
                                    <td align="left" valign="middle" style="height:20px; width:100px;">                                        
                                    </td>
                                    <td align="left" valign="middle" height="20px" width="180px"></td>
                                    <td style="width:5px;"></td>
                                    <td align="left" valign="middle" height="20px" width="135px">
                                        
                                    </td>

                                    <td align="left" valign="middle" height="20px" width="80px"></td>

                                    <td align="right" valign="middle" style="height:25px; width:200px;">
                                        
                                    </td>
                                    <td align="right" width="180px">
                                        <table style="width:100%;" align="right">
                                            <tr>
                                                <td align="right" valign="middle" style="height:25px; width:90px;">
                                                    <dx:ASPxButton ID="btnRefresh" runat="server" Text="SEARCH"
                                                        Font-Names="Tahoma" Width="85px" AutoPostBack="False" Font-Size="8pt" >
                                                        <ClientSideEvents Click="function(s, e) {
                                                            grid.PerformCallback('load' + '|' + cboPartNo.GetText());
                                                            lblInfo.SetText('');
                                                        }" />
                                                    </dx:ASPxButton>
                                                </td>
                                                <td align="right" valign="middle" style="height:25px; width:90px;">
                                                    <dx:ASPxButton ID="btnClear" runat="server" Text="CLEAR"
                                                        Font-Names="Tahoma" Width="85px" AutoPostBack="False" Font-Size="8pt">                                 
                                                    </dx:ASPxButton>
                                                </td>
                                            </tr>
                                        </table>
                                    </td>
                                </tr>                               
                            </table>
                        </td>
                    </tr>
                </table>
            </td>            
        </tr>
    </table>

    <div style="height:1px;"></div>

    <table style="width:100%; height: 15px;">
        <tr>
            <td colspan="8" height="15">
            <%--error message--%>
                <table id="tblMsg" style="border-width: thin; border-style: inset hidden ridge hidden; border-color:#9598A1; width:100%; height: 15px;">
                    <tr>
                        <td align="left" valign="middle" height="15px">
                            <dx:ASPxLabel ID="lblInfo" runat="server" Text="[lblinfo]" Font-Names="Tahoma" 
                                ClientInstanceName="lblInfo" Font-Bold="True" Font-Italic="True" 
                                Font-Size="8pt">
                            </dx:ASPxLabel>
                        </td>
                    </tr>         
                </table>
            </td>            
        </tr>
    </table>

    <div style="height:1px;"></div>

    <table style="width:100%;">
        <tr>
            <td valign="top" width="40%" align="left">
                <dx:ASPxRoundPanel ID="ASPxRoundPanel1" runat="server" HeaderText="PO STATUS" ShowCollapseButton="true"
                    View="GroupBox" Width="100%" Font-Names="Tahoma">
                    <PanelCollection>
                        <dx:PanelContent ID="PanelContent1" runat="server">
                            <table id="Table2">
                                <tr>
                                    <td align="left" valign="middle">
                                        <dx:ASPxLabel ID="ASPxLabel18" runat="server" Font-Names="Tahoma" Font-Size="8pt"
                                            ForeColor="#003366" Text="(1) AFFILIATE ENTRY">
                                        </dx:ASPxLabel>
                                    </td>
                                    <td align="left" valign="middle">
                                        <dx:ASPxTextBox ID="txtDate1" runat="server" ClientInstanceName="txtDate1" 
                                            Font-Names="Tahoma" Font-Size="8pt" Height="20px" MaxLength="10" 
                                            ReadOnly="True" Width="130px">
                                            <readonlystyle backcolor="#CCCCCC">
                                            </readonlystyle>
                                        </dx:ASPxTextBox>
                                    </td>
                                    <td align="left" valign="middle">
                                        <dx:ASPxTextBox ID="txtUser1" runat="server" ClientInstanceName="txtUser1" 
                                            Font-Names="Tahoma" Font-Size="8pt" Height="20px" MaxLength="10" 
                                            ReadOnly="True" Width="130px">
                                            <readonlystyle backcolor="#CCCCCC">
                                            </readonlystyle>
                                        </dx:ASPxTextBox>
                                    </td>
                                    <td align="left" valign="middle">
                                        &nbsp;</td>
                                    <td align="left" valign="middle">
                                        &nbsp;</td>
                                    <td align="left" valign="middle">
                                        <dx:ASPxLabel ID="ASPxLabel14" runat="server" Font-Names="Tahoma" 
                                            Font-Size="8pt" ForeColor="#003366" Text="(5) SUPPLIER PENDING (PARTIAL)">
                                        </dx:ASPxLabel>
                                    </td>
                                    <td align="left" valign="middle">
                                        <dx:ASPxTextBox ID="txtDate5" runat="server" ClientInstanceName="txtDate5" 
                                            Font-Names="Tahoma" Font-Size="8pt" Height="20px" MaxLength="10" 
                                            ReadOnly="True" Width="130px">
                                            <readonlystyle backcolor="#CCCCCC">
                                            </readonlystyle>
                                        </dx:ASPxTextBox>
                                    </td>
                                    <td align="left" valign="middle">
                                        <dx:ASPxTextBox ID="txtUser5" runat="server" ClientInstanceName="txtUser5" 
                                            Font-Names="Tahoma" Font-Size="8pt" Height="20px" MaxLength="10" 
                                            ReadOnly="True" Width="130px">
                                            <readonlystyle backcolor="#CCCCCC">
                                            </readonlystyle>
                                        </dx:ASPxTextBox>
                                    </td>
                                </tr>
                                <tr>
                                    <td align="left" valign="middle">
                                        <dx:ASPxLabel ID="ASPxLabel11" runat="server" Font-Names="Tahoma" Font-Size="8pt"
                                            ForeColor="#003366" Text="(2) AFFILIATE APPROVAL">
                                        </dx:ASPxLabel>
                                    </td>
                                    <td align="left" valign="middle">
                                        <dx:ASPxTextBox ID="txtDate2" runat="server" ClientInstanceName="txtDate2"
                                            Font-Names="Tahoma" Font-Size="8pt" Height="20px" MaxLength="10"
                                            ReadOnly="True" Width="130px">
                                            <readonlystyle backcolor="#CCCCCC">
                                            </readonlystyle>
                                        </dx:ASPxTextBox>
                                    </td>
                                    <td align="left" valign="middle">
                                        <dx:ASPxTextBox ID="txtUser2" runat="server" ClientInstanceName="txtUser2"
                                            Font-Names="Tahoma" Font-Size="8pt" Height="20px" MaxLength="10"
                                            ReadOnly="True" Width="130px">
                                            <readonlystyle backcolor="#CCCCCC">
                                            </readonlystyle>
                                        </dx:ASPxTextBox>
                                    </td>
                                    <td align="left" valign="middle">
                                        &nbsp;</td>
                                    <td align="left" valign="middle">
                                        &nbsp;</td>
                                    <td align="left" valign="middle">
                                        <dx:ASPxLabel ID="ASPxLabel15" runat="server" Font-Names="Tahoma" 
                                            Font-Size="8pt" ForeColor="#003366" Text="(6) SUPPLIER UNAPPROVE">
                                        </dx:ASPxLabel>
                                    </td>
                                    <td align="left" valign="middle">
                                        <dx:ASPxTextBox ID="txtDate6" runat="server" ClientInstanceName="txtDate6"
                                            Font-Names="Tahoma" Font-Size="8pt" Height="20px" MaxLength="10"
                                            ReadOnly="True" Width="130px">
                                            <readonlystyle backcolor="#CCCCCC">
                                            </readonlystyle>
                                        </dx:ASPxTextBox>
                                    </td>
                                    <td align="left" valign="middle">
                                        <dx:ASPxTextBox ID="txtUser6" runat="server" ClientInstanceName="txtUser6"
                                            Font-Names="Tahoma" Font-Size="8pt" Height="20px" MaxLength="10"
                                            ReadOnly="True" Width="130px">
                                            <readonlystyle backcolor="#CCCCCC">
                                            </readonlystyle>
                                        </dx:ASPxTextBox>
                                    </td>
                                </tr>
                                <tr>
                                    <td align="left" valign="middle">
                                        <dx:ASPxLabel ID="ASPxLabel12" runat="server" Font-Names="Tahoma" Font-Size="8pt"
                                            ForeColor="#003366" Text="(3) PASI SEND AFFILIATE PO TO SUPPLIER">
                                        </dx:ASPxLabel>
                                    </td>
                                    <td align="left" valign="middle">
                                        <dx:ASPxTextBox ID="txtDate3" runat="server" ClientInstanceName="txtDate3"
                                            Font-Names="Tahoma" Font-Size="8pt" Height="20px" MaxLength="10"
                                            ReadOnly="True" Width="130px">
                                            <readonlystyle backcolor="#CCCCCC">
                                            </readonlystyle>
                                        </dx:ASPxTextBox>
                                    </td>
                                    <td align="left" valign="middle">
                                        <dx:ASPxTextBox ID="txtUser3" runat="server" ClientInstanceName="txtUser3"
                                            Font-Names="Tahoma" Font-Size="8pt" Height="20px" MaxLength="10"
                                            ReadOnly="True" Width="130px">
                                            <readonlystyle backcolor="#CCCCCC">
                                            </readonlystyle>
                                        </dx:ASPxTextBox>
                                    </td>
                                    <td align="left" valign="middle">
                                        &nbsp;</td>
                                    <td align="left" valign="middle">
                                        &nbsp;</td>
                                    <td align="left" valign="middle">
                                        <dx:ASPxLabel ID="ASPxLabel16" runat="server" Font-Names="Tahoma" 
                                            Font-Size="8pt" ForeColor="#003366" Text="(7) PASI APPROVAL">
                                        </dx:ASPxLabel>
                                    </td>
                                    <td align="left" valign="middle">
                                        <dx:ASPxTextBox ID="txtDate7" runat="server" ClientInstanceName="txtDate7"
                                            Font-Names="Tahoma" Font-Size="8pt" Height="20px" MaxLength="10"
                                            ReadOnly="True" Width="130px">
                                            <readonlystyle backcolor="#CCCCCC">
                                            </readonlystyle>
                                        </dx:ASPxTextBox>
                                    </td>
                                    <td align="left" valign="middle">
                                        <dx:ASPxTextBox ID="txtUser7" runat="server" ClientInstanceName="txtUser7"
                                            Font-Names="Tahoma" Font-Size="8pt" Height="20px" MaxLength="10"
                                            ReadOnly="True" Width="130px">
                                            <readonlystyle backcolor="#CCCCCC">
                                            </readonlystyle>
                                        </dx:ASPxTextBox>
                                    </td>
                                </tr>
                                <tr>
                                    <td align="left" valign="middle">
                                        <dx:ASPxLabel ID="ASPxLabel13" runat="server" Font-Names="Tahoma" Font-Size="8pt"
                                            ForeColor="#003366" Text="(4) SUPPLIER APPROVAL">
                                        </dx:ASPxLabel>
                                    </td>
                                    <td align="left" valign="middle">
                                        <dx:ASPxTextBox ID="txtDate4" runat="server" ClientInstanceName="txtDate4"
                                            Font-Names="Tahoma" Font-Size="8pt" Height="20px" MaxLength="10"
                                            ReadOnly="True" Width="130px">
                                            <readonlystyle backcolor="#CCCCCC">
                                            </readonlystyle>
                                        </dx:ASPxTextBox>
                                    </td>
                                    <td align="left" valign="middle">
                                        <dx:ASPxTextBox ID="txtUser4" runat="server" ClientInstanceName="txtUser4"
                                            Font-Names="Tahoma" Font-Size="8pt" Height="20px" MaxLength="10"
                                            ReadOnly="True" Width="130px">
                                            <readonlystyle backcolor="#CCCCCC">
                                            </readonlystyle>
                                        </dx:ASPxTextBox>
                                    </td>
                                    <td align="left" valign="middle">
                                        &nbsp;</td>
                                    <td align="left" valign="middle">
                                        &nbsp;</td>
                                    <td align="left" valign="middle">
                                        <dx:ASPxLabel ID="ASPxLabel17" runat="server" Font-Names="Tahoma"
                                            Font-Size="8pt" ForeColor="#003366" Text="(8) AFFILIATE FINAL APPROVAL">
                                        </dx:ASPxLabel>
                                    </td>
                                    <td align="left" valign="middle">
                                        <dx:ASPxTextBox ID="txtDate8" runat="server" ClientInstanceName="txtDate8"
                                            Font-Names="Tahoma" Font-Size="8pt" Height="20px" MaxLength="10"
                                            ReadOnly="True" Width="130px">
                                            <readonlystyle backcolor="#CCCCCC">
                                            </readonlystyle>
                                        </dx:ASPxTextBox>
                                    </td>
                                    <td align="left" valign="middle">
                                        <dx:ASPxTextBox ID="txtUser8" runat="server" ClientInstanceName="txtUser8"
                                            Font-Names="Tahoma" Font-Size="8pt" Height="20px" MaxLength="10"
                                            ReadOnly="True" Width="130px">
                                            <readonlystyle backcolor="#CCCCCC">
                                            </readonlystyle>
                                        </dx:ASPxTextBox>
                                    </td>
                                </tr>
                            </table>
                        </dx:PanelContent>
                    </PanelCollection>
                </dx:ASPxRoundPanel>
            </td>
        </tr>        
    </table>
    <table style="width:100%;">
        <tr>
            <td valign="top" align="left">
                &nbsp;
            </td>
            <td valign="top" align="left">
                &nbsp;
            </td>
            <td valign="top" align="left">
                &nbsp;
            </td>
            <td valign="top" align="left">
                &nbsp;
            </td>
            <td valign="top" align="left">
                &nbsp;
            </td>
            <td valign="top" align="right" width="300px">
                <table style="width: 100%;">
                    <tr>
                        <td align="right" valign="middle" width="30px">
                            <asp:TextBox ID="lSuuplier" runat="server" BackColor="Yellow" BorderStyle="None"
                                ReadOnly="True" Width="30px"></asp:TextBox>
                        </td>
                        <td align="right" valign="middle" width="110PX">
                            <dx:ASPxLabel runat="server" Text=": EDITED BY PASI" Font-Names="Tahoma" Font-Size="8pt"
                                ID="ASPxLabel21" Width="110px">
                            </dx:ASPxLabel>
                        </td>
                        <td width="3px">&nbsp;</td>
                        <td align="right" valign="middle" width="30px">
                            <asp:TextBox ID="lSuuplier0" runat="server" BackColor="GreenYellow" BorderStyle="None"
                                ReadOnly="True" Width="30px"></asp:TextBox>
                        </td>
                        <td align="right" valign="middle" width="140px">
                            <dx:ASPxLabel runat="server" Text=": EDITED BY SUPPLIER" Font-Names="Tahoma" Font-Size="8pt"
                                ID="ASPxLabel2" Width="140px">
                            </dx:ASPxLabel>
                        </td>
                    </tr>
                </table>
            </td>
        </tr>
    </table>
    <div style="height:1px;"></div>

    <table style="width:100%;">
        <tr>
            <td colspan="8" align="left" valign="top" height="220">
                <%--Column : Grid--%>
                <dx:ASPxGridView ID="grid" runat="server" Width="100%"
                    Font-Names="Tahoma" KeyFieldName="PartNo1;AffiliateName"
                    AutoGenerateColumns="False"
                    ClientInstanceName="grid" Font-Size="8pt">
                    <ClientSideEvents Init="OnInit" FocusedRowChanged="function(s, e) {
	                    OnGridFocusedRowChanged();
                    }" CallbackError="function(s, e) {e.handled = true;}" EndCallback="function(s, e) {
                        var pMsg = s.cpMessage;
                        if (pMsg != '') {
                            if (pMsg.substring(1,5) == '1007') {
                                lblInfo.GetMainElement().style.color = 'Blue';
                            } else {
                                lblInfo.GetMainElement().style.color = 'Red';
                            }
                            lblInfo.SetText(pMsg);
                        } else {
                            lblInfo.SetText('');
                        }
                        delete s.cpMessage;

                        if (s.cpSearch != '') {
                            txtUser1.SetText(s.cpUser1);
                            txtUser2.SetText(s.cpUser2);
                            txtUser3.SetText(s.cpUser3);
                            txtUser4.SetText(s.cpUser4);
                            txtUser5.SetText(s.cpUser5);
                            txtUser6.SetText(s.cpUser6);
                            txtUser7.SetText(s.cpUser7);
                            txtUser8.SetText(s.cpUser8);

                            txtDate1.SetText(s.cpDate1);
                            txtDate2.SetText(s.cpDate2);
                            txtDate3.SetText(s.cpDate3);
                            txtDate4.SetText(s.cpDate4);
                            txtDate5.SetText(s.cpDate5);
                            txtDate6.SetText(s.cpDate6);
                            txtDate7.SetText(s.cpDate7);
                            txtDate8.SetText(s.cpDate8);
                        }
                        
                        if(s.cpButton == 'YES'){
                            btnApprove.SetEnabled(false);
                        }else{
                            btnApprove.SetEnabled(true);
                        }

                        delete s.cpButton
                        delete s.cpSearch;
                    }" RowClick="function(s, e) {
	                    lblInfo.SetText('');
                    }" />                    
                    <Columns>
                        <dx:GridViewDataTextColumn VisibleIndex="1" Caption="NO." FieldName="NoUrut" Width="30px" HeaderStyle-HorizontalAlign="Center">
<HeaderStyle HorizontalAlign="Center"></HeaderStyle>

                            <CellStyle Font-Names="Tahoma" Font-Size="8pt">
                            </CellStyle>
                        </dx:GridViewDataTextColumn>
                        <dx:GridViewDataTextColumn VisibleIndex="2" Caption="PART NO." FieldName="PartNo" Width="90px" HeaderStyle-HorizontalAlign="Center">
<HeaderStyle HorizontalAlign="Center"></HeaderStyle>

                            <CellStyle Font-Names="Tahoma" Font-Size="8pt">
                            </CellStyle>
                        </dx:GridViewDataTextColumn>
                        <dx:GridViewDataTextColumn VisibleIndex="3" Caption="PART NAME" FieldName="PartName" Width="180px" HeaderStyle-HorizontalAlign="Center">
<HeaderStyle HorizontalAlign="Center"></HeaderStyle>

                            <CellStyle Font-Names="Tahoma" Font-Size="8pt">
                            </CellStyle>
                        </dx:GridViewDataTextColumn>
                        <dx:GridViewDataTextColumn VisibleIndex="4" Caption="KANBAN CLS" FieldName="KanbanCls" Width="0px" HeaderStyle-HorizontalAlign="Center" CellStyle-HorizontalAlign="Center">
<HeaderStyle HorizontalAlign="Center"></HeaderStyle>

                            <CellStyle Font-Names="Tahoma" Font-Size="8pt">
                            </CellStyle>
                        </dx:GridViewDataTextColumn>
                        <dx:GridViewDataTextColumn VisibleIndex="5" Caption="UOM" FieldName="UnitDesc" Width="40px" HeaderStyle-HorizontalAlign="Center" CellStyle-HorizontalAlign="Center">
<HeaderStyle HorizontalAlign="Center"></HeaderStyle>

                            <CellStyle Font-Names="Tahoma" Font-Size="8pt">
                            </CellStyle>
                        </dx:GridViewDataTextColumn>
                        <dx:GridViewDataTextColumn VisibleIndex="6" Caption="MOQ" FieldName="MOQ" Width="70px" HeaderStyle-HorizontalAlign="Center">
                            <PropertiesTextEdit DisplayFormatString="{0:n0}">
                            </PropertiesTextEdit>

<HeaderStyle HorizontalAlign="Center"></HeaderStyle>

                            <CellStyle Font-Names="Tahoma" Font-Size="8pt">
                            </CellStyle>
                        </dx:GridViewDataTextColumn>
                        
                        <dx:GridViewDataTextColumn VisibleIndex="7" Caption="QTY/BOX" FieldName="QtyBox" Width="70px" HeaderStyle-HorizontalAlign="Center">
                            <PropertiesTextEdit DisplayFormatString="{0:n0}">
                            </PropertiesTextEdit>

<HeaderStyle HorizontalAlign="Center"></HeaderStyle>

                            <CellStyle Font-Names="Tahoma" Font-Size="8pt">
                            </CellStyle>
                        </dx:GridViewDataTextColumn>  
                        <dx:GridViewDataTextColumn VisibleIndex="8" Caption="MAKER" FieldName="Maker" Width="100px" HeaderStyle-HorizontalAlign="Center">
<HeaderStyle HorizontalAlign="Center"></HeaderStyle>

                            <CellStyle Font-Names="Tahoma" Font-Size="8pt">
                            </CellStyle>
                        </dx:GridViewDataTextColumn>

                        <dx:GridViewDataTextColumn VisibleIndex="9" Caption=" " FieldName="AffiliateName" Width="80px" HeaderStyle-HorizontalAlign="Center">
<HeaderStyle HorizontalAlign="Center"></HeaderStyle>

                            <CellStyle Font-Names="Tahoma" Font-Size="8pt">
                            </CellStyle>
                        </dx:GridViewDataTextColumn>

                        <dx:GridViewDataTextColumn VisibleIndex="10" Caption="TOTAL FIRM QTY" FieldName="POQty" Width="100px" HeaderStyle-HorizontalAlign="Center">
                            <PropertiesTextEdit DisplayFormatString="{0:n0}">
                            </PropertiesTextEdit>

<HeaderStyle HorizontalAlign="Center"></HeaderStyle>

                            <CellStyle Font-Names="Tahoma" Font-Size="8pt">
                            </CellStyle>
                        </dx:GridViewDataTextColumn>

                        <dx:GridViewDataTextColumn VisibleIndex="11" Caption="TOTAL FIRM QTY" FieldName="POQtyOld" Width="0px" HeaderStyle-HorizontalAlign="Center">
                            <PropertiesTextEdit DisplayFormatString="{0:n0}">
                            </PropertiesTextEdit>

<HeaderStyle HorizontalAlign="Center"></HeaderStyle>

                            <CellStyle Font-Names="Tahoma" Font-Size="8pt">
                            </CellStyle>
                        </dx:GridViewDataTextColumn>

                        <dx:GridViewDataTextColumn VisibleIndex="12" Caption="CURR" FieldName="CurrDesc" Width="45px" HeaderStyle-HorizontalAlign="Center" CellStyle-HorizontalAlign="Center" Visible="False">                            
<HeaderStyle HorizontalAlign="Center"></HeaderStyle>

                            <CellStyle Font-Names="Tahoma" Font-Size="8pt">
                            </CellStyle>
                        </dx:GridViewDataTextColumn>  
                        <dx:GridViewDataTextColumn VisibleIndex="13" Caption="PRICE" FieldName="Price" Width="80px" HeaderStyle-HorizontalAlign="Center" Visible="False">
                            <PropertiesTextEdit DisplayFormatString="{0:n2}">
                            </PropertiesTextEdit>

<HeaderStyle HorizontalAlign="Center"></HeaderStyle>

                            <CellStyle Font-Names="Tahoma" Font-Size="8pt">
                            </CellStyle>
                        </dx:GridViewDataTextColumn>  
                        <dx:GridViewDataTextColumn VisibleIndex="14" Caption="AMOUNT" FieldName="Amount" Width="110px" HeaderStyle-HorizontalAlign="Center" Visible="False">
                            <PropertiesTextEdit DisplayFormatString="{0:n2}">
                            </PropertiesTextEdit>

<HeaderStyle HorizontalAlign="Center"></HeaderStyle>

                            <CellStyle Font-Names="Tahoma" Font-Size="8pt">
                            </CellStyle>
                        </dx:GridViewDataTextColumn>
                        <dx:GridViewDataTextColumn VisibleIndex="15" Caption="FORECAST N+1" FieldName="ForecastN1" Width="80px" HeaderStyle-HorizontalAlign="Center">
                            <PropertiesTextEdit DisplayFormatString="{0:n0}">
                            </PropertiesTextEdit>

<HeaderStyle HorizontalAlign="Center"></HeaderStyle>

                            <CellStyle Font-Names="Tahoma" Font-Size="8pt">
                            </CellStyle>
                        </dx:GridViewDataTextColumn>
                        <dx:GridViewDataTextColumn VisibleIndex="16" Caption="FORECAST N+2" FieldName="ForecastN2" Width="80px" HeaderStyle-HorizontalAlign="Center">
                            <PropertiesTextEdit DisplayFormatString="{0:n0}">
                            </PropertiesTextEdit>

<HeaderStyle HorizontalAlign="Center"></HeaderStyle>

                            <CellStyle Font-Names="Tahoma" Font-Size="8pt">
                            </CellStyle>
                        </dx:GridViewDataTextColumn>
                        <dx:GridViewDataTextColumn VisibleIndex="17" Caption="FORECAST N+3" FieldName="ForecastN3" Width="80px" HeaderStyle-HorizontalAlign="Center">
                            <PropertiesTextEdit DisplayFormatString="{0:n0}">
                            </PropertiesTextEdit>

<HeaderStyle HorizontalAlign="Center"></HeaderStyle>

                            <CellStyle Font-Names="Tahoma" Font-Size="8pt">
                            </CellStyle>
                        </dx:GridViewDataTextColumn>

                        <dx:GridViewBandColumn Caption="E.T.A SCHEDULE (BASED ON FIRM ORDER)" VisibleIndex="17" HeaderStyle-HorizontalAlign="Center">
                            <Columns>
                                <dx:GridViewDataTextColumn VisibleIndex="18" Caption="1" Width="60px" FieldName="DeliveryD1" HeaderStyle-HorizontalAlign="Center">
                                    <PropertiesTextEdit DisplayFormatString="{0:n0}">
                                    </PropertiesTextEdit>

<HeaderStyle HorizontalAlign="Center"></HeaderStyle>

                                    <CellStyle Font-Names="Tahoma" Font-Size="8pt">
                                    </CellStyle>
                                </dx:GridViewDataTextColumn>
                                <dx:GridViewDataTextColumn VisibleIndex="19" Caption="1O" Width="0px" FieldName="DeliveryD1Old" HeaderStyle-HorizontalAlign="Center">
                                    <PropertiesTextEdit DisplayFormatString="{0:n0}">
                                    </PropertiesTextEdit>

<HeaderStyle HorizontalAlign="Center"></HeaderStyle>

                                    <CellStyle Font-Names="Tahoma" Font-Size="8pt">
                                    </CellStyle>
                                </dx:GridViewDataTextColumn>
                                <dx:GridViewDataTextColumn VisibleIndex="20" Caption="2" Width="60px" FieldName="DeliveryD2" HeaderStyle-HorizontalAlign="Center">
                                    <PropertiesTextEdit DisplayFormatString="{0:n0}">
                                    </PropertiesTextEdit>

<HeaderStyle HorizontalAlign="Center"></HeaderStyle>

                                    <CellStyle Font-Names="Tahoma" Font-Size="8pt">
                                    </CellStyle>
                                </dx:GridViewDataTextColumn>
                                <dx:GridViewDataTextColumn VisibleIndex="21" Caption="2O" Width="0px" FieldName="DeliveryD2Old" HeaderStyle-HorizontalAlign="Center">
                                    <PropertiesTextEdit DisplayFormatString="{0:n0}">
                                    </PropertiesTextEdit>

<HeaderStyle HorizontalAlign="Center"></HeaderStyle>

                                    <CellStyle Font-Names="Tahoma" Font-Size="8pt">
                                    </CellStyle>
                                </dx:GridViewDataTextColumn>
                                <dx:GridViewDataTextColumn VisibleIndex="22" Caption="3" Width="60px" FieldName="DeliveryD3" HeaderStyle-HorizontalAlign="Center">
                                    <PropertiesTextEdit DisplayFormatString="{0:n0}">
                                    </PropertiesTextEdit>

<HeaderStyle HorizontalAlign="Center"></HeaderStyle>

                                    <CellStyle Font-Names="Tahoma" Font-Size="8pt">
                                    </CellStyle>
                                </dx:GridViewDataTextColumn>
                                <dx:GridViewDataTextColumn VisibleIndex="23" Caption="3O" Width="0px" FieldName="DeliveryD3Old" HeaderStyle-HorizontalAlign="Center">
                                    <PropertiesTextEdit DisplayFormatString="{0:n0}">
                                    </PropertiesTextEdit>

<HeaderStyle HorizontalAlign="Center"></HeaderStyle>

                                    <CellStyle Font-Names="Tahoma" Font-Size="8pt">
                                    </CellStyle>
                                </dx:GridViewDataTextColumn>
                                <dx:GridViewDataTextColumn VisibleIndex="24" Caption="4" Width="60px" FieldName="DeliveryD4" HeaderStyle-HorizontalAlign="Center">
                                    <PropertiesTextEdit DisplayFormatString="{0:n0}">
                                    </PropertiesTextEdit>

<HeaderStyle HorizontalAlign="Center"></HeaderStyle>

                                    <CellStyle Font-Names="Tahoma" Font-Size="8pt">
                                    </CellStyle>
                                </dx:GridViewDataTextColumn>
                                <dx:GridViewDataTextColumn VisibleIndex="25" Caption="4O" Width="0px" FieldName="DeliveryD4Old" HeaderStyle-HorizontalAlign="Center">
                                    <PropertiesTextEdit DisplayFormatString="{0:n0}">
                                    </PropertiesTextEdit>

<HeaderStyle HorizontalAlign="Center"></HeaderStyle>

                                    <CellStyle Font-Names="Tahoma" Font-Size="8pt">
                                    </CellStyle>
                                </dx:GridViewDataTextColumn>
                                <dx:GridViewDataTextColumn VisibleIndex="26" Caption="5" Width="60px" FieldName="DeliveryD5" HeaderStyle-HorizontalAlign="Center">
                                    <PropertiesTextEdit DisplayFormatString="{0:n0}">
                                    </PropertiesTextEdit>

<HeaderStyle HorizontalAlign="Center"></HeaderStyle>

                                    <CellStyle Font-Names="Tahoma" Font-Size="8pt">
                                    </CellStyle>
                                </dx:GridViewDataTextColumn>
                                <dx:GridViewDataTextColumn VisibleIndex="27" Caption="5O" Width="0px" FieldName="DeliveryD5Old" HeaderStyle-HorizontalAlign="Center">
                                    <PropertiesTextEdit DisplayFormatString="{0:n0}">
                                    </PropertiesTextEdit>

<HeaderStyle HorizontalAlign="Center"></HeaderStyle>

                                    <CellStyle Font-Names="Tahoma" Font-Size="8pt">
                                    </CellStyle>
                                </dx:GridViewDataTextColumn>
                                <dx:GridViewDataTextColumn VisibleIndex="28" Caption="6" Width="60px" FieldName="DeliveryD6" HeaderStyle-HorizontalAlign="Center">
                                    <PropertiesTextEdit DisplayFormatString="{0:n0}">
                                    </PropertiesTextEdit>

<HeaderStyle HorizontalAlign="Center"></HeaderStyle>

                                    <CellStyle Font-Names="Tahoma" Font-Size="8pt">
                                    </CellStyle>
                                </dx:GridViewDataTextColumn>
                                <dx:GridViewDataTextColumn VisibleIndex="29" Caption="6O" Width="0px" FieldName="DeliveryD6Old" HeaderStyle-HorizontalAlign="Center">
                                    <PropertiesTextEdit DisplayFormatString="{0:n0}">
                                    </PropertiesTextEdit>

<HeaderStyle HorizontalAlign="Center"></HeaderStyle>

                                    <CellStyle Font-Names="Tahoma" Font-Size="8pt">
                                    </CellStyle>
                                </dx:GridViewDataTextColumn>
                                <dx:GridViewDataTextColumn VisibleIndex="30" Caption="7" Width="60px" FieldName="DeliveryD7" HeaderStyle-HorizontalAlign="Center">
                                    <PropertiesTextEdit DisplayFormatString="{0:n0}">
                                    </PropertiesTextEdit>

<HeaderStyle HorizontalAlign="Center"></HeaderStyle>

                                    <CellStyle Font-Names="Tahoma" Font-Size="8pt">
                                    </CellStyle>
                                </dx:GridViewDataTextColumn>
                                <dx:GridViewDataTextColumn VisibleIndex="31" Caption="7O" Width="0px" FieldName="DeliveryD7Old" HeaderStyle-HorizontalAlign="Center">
                                    <PropertiesTextEdit DisplayFormatString="{0:n0}">
                                    </PropertiesTextEdit>

<HeaderStyle HorizontalAlign="Center"></HeaderStyle>

                                    <CellStyle Font-Names="Tahoma" Font-Size="8pt">
                                    </CellStyle>
                                </dx:GridViewDataTextColumn>
                                <dx:GridViewDataTextColumn VisibleIndex="32" Caption="8" Width="60px" FieldName="DeliveryD8" HeaderStyle-HorizontalAlign="Center">
                                    <PropertiesTextEdit DisplayFormatString="{0:n0}">
                                    </PropertiesTextEdit>

<HeaderStyle HorizontalAlign="Center"></HeaderStyle>

                                    <CellStyle Font-Names="Tahoma" Font-Size="8pt">
                                    </CellStyle>
                                </dx:GridViewDataTextColumn>
                                <dx:GridViewDataTextColumn VisibleIndex="33" Caption="8O" Width="0px" FieldName="DeliveryD8Old" HeaderStyle-HorizontalAlign="Center">
                                   <PropertiesTextEdit DisplayFormatString="{0:n0}">
                                    </PropertiesTextEdit>

<HeaderStyle HorizontalAlign="Center"></HeaderStyle>

                                    <CellStyle Font-Names="Tahoma" Font-Size="8pt">
                                    </CellStyle>
                                </dx:GridViewDataTextColumn>
                                <dx:GridViewDataTextColumn VisibleIndex="34" Caption="9" Width="60px" FieldName="DeliveryD9" HeaderStyle-HorizontalAlign="Center">
                                    <PropertiesTextEdit DisplayFormatString="{0:n0}">
                                    </PropertiesTextEdit>

<HeaderStyle HorizontalAlign="Center"></HeaderStyle>

                                    <CellStyle Font-Names="Tahoma" Font-Size="8pt">
                                    </CellStyle>
                                </dx:GridViewDataTextColumn>
                                <dx:GridViewDataTextColumn VisibleIndex="35" Caption="9O" Width="0px" FieldName="DeliveryD9Old" HeaderStyle-HorizontalAlign="Center">
                                    <PropertiesTextEdit DisplayFormatString="{0:n0}">
                                    </PropertiesTextEdit>

<HeaderStyle HorizontalAlign="Center"></HeaderStyle>

                                    <CellStyle Font-Names="Tahoma" Font-Size="8pt">
                                    </CellStyle>
                                </dx:GridViewDataTextColumn>
                                <dx:GridViewDataTextColumn VisibleIndex="36" Caption="10" Width="60px" FieldName="DeliveryD10" HeaderStyle-HorizontalAlign="Center">
                                    <PropertiesTextEdit DisplayFormatString="{0:n0}">
                                    </PropertiesTextEdit>

<HeaderStyle HorizontalAlign="Center"></HeaderStyle>

                                    <CellStyle Font-Names="Tahoma" Font-Size="8pt">
                                    </CellStyle>
                                </dx:GridViewDataTextColumn>
                                <dx:GridViewDataTextColumn VisibleIndex="37" Caption="10O" Width="0px" FieldName="DeliveryD10Old" HeaderStyle-HorizontalAlign="Center">
                                    <PropertiesTextEdit DisplayFormatString="{0:n0}">
                                    </PropertiesTextEdit>

<HeaderStyle HorizontalAlign="Center"></HeaderStyle>

                                    <CellStyle Font-Names="Tahoma" Font-Size="8pt">
                                    </CellStyle>
                                </dx:GridViewDataTextColumn>
                                <dx:GridViewDataTextColumn VisibleIndex="38" Caption="11" Width="60px" FieldName="DeliveryD11" HeaderStyle-HorizontalAlign="Center">
                                    <PropertiesTextEdit DisplayFormatString="{0:n0}">
                                    </PropertiesTextEdit>

<HeaderStyle HorizontalAlign="Center"></HeaderStyle>

                                    <CellStyle Font-Names="Tahoma" Font-Size="8pt">
                                    </CellStyle>
                                </dx:GridViewDataTextColumn>
                                <dx:GridViewDataTextColumn VisibleIndex="39" Caption="11O" Width="0px" FieldName="DeliveryD11Old" HeaderStyle-HorizontalAlign="Center">
                                    <PropertiesTextEdit DisplayFormatString="{0:n0}">
                                    </PropertiesTextEdit>

<HeaderStyle HorizontalAlign="Center"></HeaderStyle>

                                    <CellStyle Font-Names="Tahoma" Font-Size="8pt">
                                    </CellStyle>
                                </dx:GridViewDataTextColumn>
                                <dx:GridViewDataTextColumn VisibleIndex="40" Caption="12" Width="60px" FieldName="DeliveryD12" HeaderStyle-HorizontalAlign="Center">
                                    <PropertiesTextEdit DisplayFormatString="{0:n0}">
                                    </PropertiesTextEdit>

<HeaderStyle HorizontalAlign="Center"></HeaderStyle>

                                    <CellStyle Font-Names="Tahoma" Font-Size="8pt">
                                    </CellStyle>
                                </dx:GridViewDataTextColumn>
                                <dx:GridViewDataTextColumn VisibleIndex="41" Caption="12O" Width="0px" FieldName="DeliveryD12Old" HeaderStyle-HorizontalAlign="Center">
                                    <PropertiesTextEdit DisplayFormatString="{0:n0}">
                                    </PropertiesTextEdit>

<HeaderStyle HorizontalAlign="Center"></HeaderStyle>

                                    <CellStyle Font-Names="Tahoma" Font-Size="8pt">
                                    </CellStyle>
                                </dx:GridViewDataTextColumn>
                                <dx:GridViewDataTextColumn VisibleIndex="42" Caption="13" Width="60px" FieldName="DeliveryD13" HeaderStyle-HorizontalAlign="Center">
                                    <PropertiesTextEdit DisplayFormatString="{0:n0}">
                                    </PropertiesTextEdit>

<HeaderStyle HorizontalAlign="Center"></HeaderStyle>

                                    <CellStyle Font-Names="Tahoma" Font-Size="8pt">
                                    </CellStyle>
                                </dx:GridViewDataTextColumn>
                                <dx:GridViewDataTextColumn VisibleIndex="43" Caption="13O" Width="0px" FieldName="DeliveryD13Old" HeaderStyle-HorizontalAlign="Center">
                                    <PropertiesTextEdit DisplayFormatString="{0:n0}">
                                    </PropertiesTextEdit>

<HeaderStyle HorizontalAlign="Center"></HeaderStyle>

                                    <CellStyle Font-Names="Tahoma" Font-Size="8pt">
                                    </CellStyle>
                                </dx:GridViewDataTextColumn>
                                <dx:GridViewDataTextColumn VisibleIndex="44" Caption="14" Width="60px" FieldName="DeliveryD14" HeaderStyle-HorizontalAlign="Center">
                                    <PropertiesTextEdit DisplayFormatString="{0:n0}">
                                    </PropertiesTextEdit>

<HeaderStyle HorizontalAlign="Center"></HeaderStyle>

                                    <CellStyle Font-Names="Tahoma" Font-Size="8pt">
                                    </CellStyle>
                                </dx:GridViewDataTextColumn>
                                <dx:GridViewDataTextColumn VisibleIndex="45" Caption="14O" Width="0px" FieldName="DeliveryD14Old" HeaderStyle-HorizontalAlign="Center">
                                    <PropertiesTextEdit DisplayFormatString="{0:n0}">
                                    </PropertiesTextEdit>

<HeaderStyle HorizontalAlign="Center"></HeaderStyle>

                                    <CellStyle Font-Names="Tahoma" Font-Size="8pt">
                                    </CellStyle>
                                </dx:GridViewDataTextColumn>
                                <dx:GridViewDataTextColumn VisibleIndex="46" Caption="15" Width="60px" FieldName="DeliveryD15" HeaderStyle-HorizontalAlign="Center">
                                    <PropertiesTextEdit DisplayFormatString="{0:n0}">
                                    </PropertiesTextEdit>

<HeaderStyle HorizontalAlign="Center"></HeaderStyle>

                                    <CellStyle Font-Names="Tahoma" Font-Size="8pt">
                                    </CellStyle>
                                </dx:GridViewDataTextColumn>
                                <dx:GridViewDataTextColumn VisibleIndex="47" Caption="15O" Width="0px" FieldName="DeliveryD15Old" HeaderStyle-HorizontalAlign="Center">
                                    <PropertiesTextEdit DisplayFormatString="{0:n0}">
                                    </PropertiesTextEdit>

<HeaderStyle HorizontalAlign="Center"></HeaderStyle>

                                    <CellStyle Font-Names="Tahoma" Font-Size="8pt">
                                    </CellStyle>
                                </dx:GridViewDataTextColumn>
                                <dx:GridViewDataTextColumn VisibleIndex="48" Caption="16" Width="60px" FieldName="DeliveryD16" HeaderStyle-HorizontalAlign="Center">
                                    <PropertiesTextEdit DisplayFormatString="{0:n0}">
                                    </PropertiesTextEdit>

<HeaderStyle HorizontalAlign="Center"></HeaderStyle>

                                    <CellStyle Font-Names="Tahoma" Font-Size="8pt">
                                    </CellStyle>
                                </dx:GridViewDataTextColumn>
                                <dx:GridViewDataTextColumn VisibleIndex="49" Caption="16O" Width="0px" FieldName="DeliveryD16Old" HeaderStyle-HorizontalAlign="Center">
                                    <PropertiesTextEdit DisplayFormatString="{0:n0}">
                                    </PropertiesTextEdit>

<HeaderStyle HorizontalAlign="Center"></HeaderStyle>

                                    <CellStyle Font-Names="Tahoma" Font-Size="8pt">
                                    </CellStyle>
                                </dx:GridViewDataTextColumn>
                                <dx:GridViewDataTextColumn VisibleIndex="50" Caption="17" Width="60px" FieldName="DeliveryD17" HeaderStyle-HorizontalAlign="Center">
                                    <PropertiesTextEdit DisplayFormatString="{0:n0}">
                                    </PropertiesTextEdit>

<HeaderStyle HorizontalAlign="Center"></HeaderStyle>

                                    <CellStyle Font-Names="Tahoma" Font-Size="8pt">
                                    </CellStyle>
                                </dx:GridViewDataTextColumn>
                                <dx:GridViewDataTextColumn VisibleIndex="51" Caption="17O" Width="0px" FieldName="DeliveryD17Old" HeaderStyle-HorizontalAlign="Center">
                                    <PropertiesTextEdit DisplayFormatString="{0:n0}">
                                    </PropertiesTextEdit>

<HeaderStyle HorizontalAlign="Center"></HeaderStyle>

                                    <CellStyle Font-Names="Tahoma" Font-Size="8pt">
                                    </CellStyle>
                                </dx:GridViewDataTextColumn>
                                <dx:GridViewDataTextColumn VisibleIndex="52" Caption="18" Width="60px" FieldName="DeliveryD18" HeaderStyle-HorizontalAlign="Center">
                                    <PropertiesTextEdit DisplayFormatString="{0:n0}">
                                    </PropertiesTextEdit>

<HeaderStyle HorizontalAlign="Center"></HeaderStyle>

                                    <CellStyle Font-Names="Tahoma" Font-Size="8pt">
                                    </CellStyle>
                                </dx:GridViewDataTextColumn>
                                <dx:GridViewDataTextColumn VisibleIndex="53" Caption="18O" Width="0px" FieldName="DeliveryD18Old" HeaderStyle-HorizontalAlign="Center">
                                    <PropertiesTextEdit DisplayFormatString="{0:n0}">
                                    </PropertiesTextEdit>

<HeaderStyle HorizontalAlign="Center"></HeaderStyle>

                                    <CellStyle Font-Names="Tahoma" Font-Size="8pt">
                                    </CellStyle>
                                </dx:GridViewDataTextColumn>
                                <dx:GridViewDataTextColumn VisibleIndex="54" Caption="19" Width="60px" FieldName="DeliveryD19" HeaderStyle-HorizontalAlign="Center">
                                    <PropertiesTextEdit DisplayFormatString="{0:n0}">
                                    </PropertiesTextEdit>

<HeaderStyle HorizontalAlign="Center"></HeaderStyle>

                                    <CellStyle Font-Names="Tahoma" Font-Size="8pt">
                                    </CellStyle>
                                </dx:GridViewDataTextColumn>
                                <dx:GridViewDataTextColumn VisibleIndex="55" Caption="19O" Width="0px" FieldName="DeliveryD19Old" HeaderStyle-HorizontalAlign="Center">
                                    <PropertiesTextEdit DisplayFormatString="{0:n0}">
                                    </PropertiesTextEdit>

<HeaderStyle HorizontalAlign="Center"></HeaderStyle>

                                    <CellStyle Font-Names="Tahoma" Font-Size="8pt">
                                    </CellStyle>
                                </dx:GridViewDataTextColumn>
                                <dx:GridViewDataTextColumn VisibleIndex="56" Caption="20" Width="60px" FieldName="DeliveryD20" HeaderStyle-HorizontalAlign="Center">
                                    <PropertiesTextEdit DisplayFormatString="{0:n0}">
                                    </PropertiesTextEdit>

<HeaderStyle HorizontalAlign="Center"></HeaderStyle>

                                    <CellStyle Font-Names="Tahoma" Font-Size="8pt">
                                    </CellStyle>
                                </dx:GridViewDataTextColumn>
                                <dx:GridViewDataTextColumn VisibleIndex="57" Caption="20O" Width="0px" FieldName="DeliveryD20Old" HeaderStyle-HorizontalAlign="Center">
                                    <PropertiesTextEdit DisplayFormatString="{0:n0}">
                                    </PropertiesTextEdit>

<HeaderStyle HorizontalAlign="Center"></HeaderStyle>

                                    <CellStyle Font-Names="Tahoma" Font-Size="8pt">
                                    </CellStyle>
                                </dx:GridViewDataTextColumn>
                                <dx:GridViewDataTextColumn VisibleIndex="58" Caption="21" Width="60px" FieldName="DeliveryD21" HeaderStyle-HorizontalAlign="Center">
                                    <PropertiesTextEdit DisplayFormatString="{0:n0}">
                                    </PropertiesTextEdit>

<HeaderStyle HorizontalAlign="Center"></HeaderStyle>

                                    <CellStyle Font-Names="Tahoma" Font-Size="8pt">
                                    </CellStyle>
                                </dx:GridViewDataTextColumn>
                                <dx:GridViewDataTextColumn VisibleIndex="59" Caption="21O" Width="0px" FieldName="DeliveryD21Old" HeaderStyle-HorizontalAlign="Center">
                                    <PropertiesTextEdit DisplayFormatString="{0:n0}">
                                    </PropertiesTextEdit>

<HeaderStyle HorizontalAlign="Center"></HeaderStyle>

                                    <CellStyle Font-Names="Tahoma" Font-Size="8pt">
                                    </CellStyle>
                                </dx:GridViewDataTextColumn>
                                <dx:GridViewDataTextColumn VisibleIndex="60" Caption="22" Width="60px" FieldName="DeliveryD22" HeaderStyle-HorizontalAlign="Center">
                                    <PropertiesTextEdit DisplayFormatString="{0:n0}">
                                    </PropertiesTextEdit>

<HeaderStyle HorizontalAlign="Center"></HeaderStyle>

                                    <CellStyle Font-Names="Tahoma" Font-Size="8pt">
                                    </CellStyle>
                                </dx:GridViewDataTextColumn>
                                <dx:GridViewDataTextColumn VisibleIndex="61" Caption="22O" Width="0px" FieldName="DeliveryD22Old" HeaderStyle-HorizontalAlign="Center">
                                    <PropertiesTextEdit DisplayFormatString="{0:n0}">
                                    </PropertiesTextEdit>

<HeaderStyle HorizontalAlign="Center"></HeaderStyle>

                                    <CellStyle Font-Names="Tahoma" Font-Size="8pt">
                                    </CellStyle>
                                </dx:GridViewDataTextColumn>
                                <dx:GridViewDataTextColumn VisibleIndex="62" Caption="23" Width="60px" FieldName="DeliveryD23" HeaderStyle-HorizontalAlign="Center">
                                    <PropertiesTextEdit DisplayFormatString="{0:n0}">
                                    </PropertiesTextEdit>

<HeaderStyle HorizontalAlign="Center"></HeaderStyle>

                                    <CellStyle Font-Names="Tahoma" Font-Size="8pt">
                                    </CellStyle>
                                </dx:GridViewDataTextColumn>
                                <dx:GridViewDataTextColumn VisibleIndex="63" Caption="23O" Width="0px" FieldName="DeliveryD23Old" HeaderStyle-HorizontalAlign="Center">
                                    <PropertiesTextEdit DisplayFormatString="{0:n0}">
                                    </PropertiesTextEdit>

<HeaderStyle HorizontalAlign="Center"></HeaderStyle>

                                    <CellStyle Font-Names="Tahoma" Font-Size="8pt">
                                    </CellStyle>
                                </dx:GridViewDataTextColumn>
                                <dx:GridViewDataTextColumn VisibleIndex="64" Caption="24" Width="60px" FieldName="DeliveryD24" HeaderStyle-HorizontalAlign="Center">
                                    <PropertiesTextEdit DisplayFormatString="{0:n0}">
                                    </PropertiesTextEdit>

<HeaderStyle HorizontalAlign="Center"></HeaderStyle>

                                    <CellStyle Font-Names="Tahoma" Font-Size="8pt">
                                    </CellStyle>
                                </dx:GridViewDataTextColumn>
                                <dx:GridViewDataTextColumn VisibleIndex="65" Caption="24O" Width="0px" FieldName="DeliveryD24Old" HeaderStyle-HorizontalAlign="Center">
                                    <PropertiesTextEdit DisplayFormatString="{0:n0}">
                                    </PropertiesTextEdit>

<HeaderStyle HorizontalAlign="Center"></HeaderStyle>

                                    <CellStyle Font-Names="Tahoma" Font-Size="8pt">
                                    </CellStyle>
                                </dx:GridViewDataTextColumn>
                                <dx:GridViewDataTextColumn VisibleIndex="66" Caption="25" Width="60px" FieldName="DeliveryD25" HeaderStyle-HorizontalAlign="Center">
                                    <PropertiesTextEdit DisplayFormatString="{0:n0}">
                                    </PropertiesTextEdit>

<HeaderStyle HorizontalAlign="Center"></HeaderStyle>

                                    <CellStyle Font-Names="Tahoma" Font-Size="8pt">
                                    </CellStyle>
                                </dx:GridViewDataTextColumn>
                                <dx:GridViewDataTextColumn VisibleIndex="67" Caption="25O" Width="0px" FieldName="DeliveryD25Old" HeaderStyle-HorizontalAlign="Center">
                                    <PropertiesTextEdit DisplayFormatString="{0:n0}">
                                    </PropertiesTextEdit>

<HeaderStyle HorizontalAlign="Center"></HeaderStyle>

                                    <CellStyle Font-Names="Tahoma" Font-Size="8pt">
                                    </CellStyle>
                                </dx:GridViewDataTextColumn>
                                <dx:GridViewDataTextColumn VisibleIndex="68" Caption="26" Width="60px" FieldName="DeliveryD26" HeaderStyle-HorizontalAlign="Center">
                                    <PropertiesTextEdit DisplayFormatString="{0:n0}">
                                    </PropertiesTextEdit>

<HeaderStyle HorizontalAlign="Center"></HeaderStyle>

                                    <CellStyle Font-Names="Tahoma" Font-Size="8pt">
                                    </CellStyle>
                                </dx:GridViewDataTextColumn>
                                <dx:GridViewDataTextColumn VisibleIndex="69" Caption="26O" Width="0px" FieldName="DeliveryD26Old" HeaderStyle-HorizontalAlign="Center">
                                    <PropertiesTextEdit DisplayFormatString="{0:n0}">
                                    </PropertiesTextEdit>

<HeaderStyle HorizontalAlign="Center"></HeaderStyle>

                                    <CellStyle Font-Names="Tahoma" Font-Size="8pt">
                                    </CellStyle>
                                </dx:GridViewDataTextColumn>
                                <dx:GridViewDataTextColumn VisibleIndex="70" Caption="27" Width="60px" FieldName="DeliveryD27" HeaderStyle-HorizontalAlign="Center">
                                    <PropertiesTextEdit DisplayFormatString="{0:n0}">
                                    </PropertiesTextEdit>

<HeaderStyle HorizontalAlign="Center"></HeaderStyle>

                                    <CellStyle Font-Names="Tahoma" Font-Size="8pt">
                                    </CellStyle>
                                </dx:GridViewDataTextColumn>
                                <dx:GridViewDataTextColumn VisibleIndex="71" Caption="27O" Width="0px" FieldName="DeliveryD27Old" HeaderStyle-HorizontalAlign="Center">
                                    <PropertiesTextEdit DisplayFormatString="{0:n0}">
                                    </PropertiesTextEdit>

<HeaderStyle HorizontalAlign="Center"></HeaderStyle>

                                    <CellStyle Font-Names="Tahoma" Font-Size="8pt">
                                    </CellStyle>
                                </dx:GridViewDataTextColumn>
                                <dx:GridViewDataTextColumn VisibleIndex="72" Caption="28" Width="60px" FieldName="DeliveryD28" HeaderStyle-HorizontalAlign="Center">
                                    <PropertiesTextEdit DisplayFormatString="{0:n0}">
                                    </PropertiesTextEdit>

<HeaderStyle HorizontalAlign="Center"></HeaderStyle>

                                    <CellStyle Font-Names="Tahoma" Font-Size="8pt">
                                    </CellStyle>
                                </dx:GridViewDataTextColumn>
                                <dx:GridViewDataTextColumn VisibleIndex="73" Caption="28O" Width="0px" FieldName="DeliveryD28Old" HeaderStyle-HorizontalAlign="Center">
                                    <PropertiesTextEdit DisplayFormatString="{0:n0}">
                                    </PropertiesTextEdit>

<HeaderStyle HorizontalAlign="Center"></HeaderStyle>

                                    <CellStyle Font-Names="Tahoma" Font-Size="8pt">
                                    </CellStyle>
                                </dx:GridViewDataTextColumn>
                                <dx:GridViewDataTextColumn VisibleIndex="74" Caption="29" Width="60px" FieldName="DeliveryD29" HeaderStyle-HorizontalAlign="Center">
                                    <PropertiesTextEdit DisplayFormatString="{0:n0}">
                                    </PropertiesTextEdit>

<HeaderStyle HorizontalAlign="Center"></HeaderStyle>

                                    <CellStyle Font-Names="Tahoma" Font-Size="8pt">
                                    </CellStyle>
                                </dx:GridViewDataTextColumn>
                                <dx:GridViewDataTextColumn VisibleIndex="75" Caption="29O" Width="0px" FieldName="DeliveryD29Old" HeaderStyle-HorizontalAlign="Center">
                                    <PropertiesTextEdit DisplayFormatString="{0:n0}">
                                    </PropertiesTextEdit>

<HeaderStyle HorizontalAlign="Center"></HeaderStyle>

                                    <CellStyle Font-Names="Tahoma" Font-Size="8pt">
                                    </CellStyle>
                                </dx:GridViewDataTextColumn>
                                <dx:GridViewDataTextColumn VisibleIndex="76" Caption="30" Width="60px" FieldName="DeliveryD30" HeaderStyle-HorizontalAlign="Center">
                                    <PropertiesTextEdit DisplayFormatString="{0:n0}">
                                    </PropertiesTextEdit>

<HeaderStyle HorizontalAlign="Center"></HeaderStyle>

                                    <CellStyle Font-Names="Tahoma" Font-Size="8pt">
                                    </CellStyle>
                                </dx:GridViewDataTextColumn>
                                <dx:GridViewDataTextColumn VisibleIndex="77" Caption="30O" Width="0px" FieldName="DeliveryD30Old" HeaderStyle-HorizontalAlign="Center">
                                    <PropertiesTextEdit DisplayFormatString="{0:n0}">
                                    </PropertiesTextEdit>

<HeaderStyle HorizontalAlign="Center"></HeaderStyle>

                                    <CellStyle Font-Names="Tahoma" Font-Size="8pt">
                                    </CellStyle>
                                </dx:GridViewDataTextColumn>
                                <dx:GridViewDataTextColumn VisibleIndex="78" Caption="31" Width="60px" FieldName="DeliveryD31" HeaderStyle-HorizontalAlign="Center">
                                    <PropertiesTextEdit DisplayFormatString="{0:n0}">
                                    </PropertiesTextEdit>

<HeaderStyle HorizontalAlign="Center"></HeaderStyle>

                                    <CellStyle Font-Names="Tahoma" Font-Size="8pt">
                                    </CellStyle>
                                </dx:GridViewDataTextColumn>
                                <dx:GridViewDataTextColumn VisibleIndex="79" Caption="31O" Width="0px" FieldName="DeliveryD31Old" HeaderStyle-HorizontalAlign="Center">
                                    <PropertiesTextEdit DisplayFormatString="{0:n0}">
                                    </PropertiesTextEdit>

<HeaderStyle HorizontalAlign="Center"></HeaderStyle>

                                    <CellStyle Font-Names="Tahoma" Font-Size="8pt">
                                    </CellStyle>
                                </dx:GridViewDataTextColumn>
                                <dx:GridViewDataTextColumn VisibleIndex="80" Caption="PartNo1" Width="0px" FieldName="PartNo1" HeaderStyle-HorizontalAlign="Center">                                                                       
<HeaderStyle HorizontalAlign="Center"></HeaderStyle>

                                    <CellStyle Font-Names="Tahoma" Font-Size="8pt">
                                    </CellStyle>
                                </dx:GridViewDataTextColumn>
                            </Columns>
                            <HeaderStyle HorizontalAlign="Center" />
                        </dx:GridViewBandColumn>
                    </Columns>
                    <SettingsBehavior AllowSelectByRowClick="True" AllowSort="False"
                        ColumnResizeMode="Control" EnableRowHotTrack="True" />
                    <SettingsPager PageSize="10" Position="Top">
                        <Summary Text="Page {0} of {1} [{2} record(s)]"
                                  AllPagesText="Page {0} of {1} " />
                        <Summary AllPagesText="Page {0} of {1} " Text="Page {0} of {1} [{2} record(s)]"></Summary>
                    </SettingsPager>
                    <SettingsEditing Mode="Batch" NewItemRowPosition="Bottom">
                        <BatchEditSettings ShowConfirmOnLosingChanges="False" />
                    </SettingsEditing>
                    <Settings ShowHorizontalScrollBar="True" ShowVerticalScrollBar="True"
                        ShowGroupButtons="False" ShowStatusBar="Hidden"
                        VerticalScrollableHeight="220" />
                    <Styles>
                        <SelectedRow ForeColor="Black">
                        </SelectedRow>
                    </Styles>
                    <StylesEditors ButtonEditCellSpacing="0">
                        <ProgressBar Height="21px">
                        </ProgressBar>
                    </StylesEditors>
                </dx:ASPxGridView>
            </td>
        </tr>

        <tr>
            <td valign="top" align="left">
                <dx:ASPxButton ID="btnSubMenu" runat="server" Text="SUB MENU"
                    Font-Names="Tahoma" Width="85px" Font-Size="8pt">
                </dx:ASPxButton>
            </td>
            <td valign="top" align="right" style="width: 50px;">
            </td>       
            <td valign="top" align="right" style="width: 50px;">
                <dx:ASPxButton ID="btnApprove" runat="server" Text="FINAL APPROVE"
                    Font-Names="Tahoma"
                    Width="90px" AutoPostBack="False" Font-Size="8pt" ClientInstanceName="btnApprove">
                 <ClientSideEvents Click="function(s, e) {                    
                    lblInfo.SetText('');
                    btnApprove.SetEnabled(false);
                    grid.PerformCallback('save' + '|' + cboPartNo.GetText() + '|' + txtPOKanban.GetText());
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

    <dx:ASPxCallback ID="ButtonApprove" runat="server" ClientInstanceName="ButtonApprove">
        <ClientSideEvents EndCallback="function(s, e) {
	        txtUser1.SetText(s.cpUser1);
            txtUser2.SetText(s.cpUser2);
            txtUser3.SetText(s.cpUser3);
            txtUser4.SetText(s.cpUser4);
            txtUser5.SetText(s.cpUser5);
            txtUser6.SetText(s.cpUser6);
            txtUser7.SetText(s.cpUser7);
            txtUser8.SetText(s.cpUser8);

            txtDate1.SetText(s.cpDate1);
            txtDate2.SetText(s.cpDate2);
            txtDate3.SetText(s.cpDate3);
            txtDate4.SetText(s.cpDate4);
            txtDate5.SetText(s.cpDate5);
            txtDate6.SetText(s.cpDate6);
            txtDate7.SetText(s.cpDate7);
            txtDate8.SetText(s.cpDate8);

            var pMsg = s.cpMessage;
            if (pMsg != '') {
                if (pMsg.substring(1,5) == '1007') {
                    lblInfo.GetMainElement().style.color = 'Blue';
                } else {
                    lblInfo.GetMainElement().style.color = 'Red';
                }
                lblInfo.SetText(pMsg);
            } else {
                lblInfo.SetText('');
            }
            delete s.cpMessage;
        }" />
    </dx:ASPxCallback>
    <dx:ASPxCallback ID="ButtonPartNo" runat="server" ClientInstanceName="ButtonPartNo">
        <ClientSideEvents EndCallback="function(s, e) {
	         if (s.cpDelivery != '') 
            {
                txtDelivery.SetText(s.cpDelivery);
                txtCommercial.SetText(s.cpCommercial);
                txtShip.SetText(s.cpShip);
                txtRemarks.SetText(s.cpRemarks);
                txtPOKanban.SetText(s.cpPOKanban);
            } else {
                lblInfo.SetText('');
            }
            delete s.cpDelivery;
        }" />
    </dx:ASPxCallback>
</asp:Content>

