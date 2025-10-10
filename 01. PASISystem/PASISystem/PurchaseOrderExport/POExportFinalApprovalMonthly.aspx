<%@ Page Title="" Language="vb" AutoEventWireup="false" MasterPageFile="~/SiteContent.Master"
    CodeBehind="POExportFinalApprovalMonthly.aspx.vb" Inherits="PASISystem.POExportFinalApprovalMonthly" %>

<%@ Register Assembly="DevExpress.Web.v14.1, Version=14.1.6.0, Culture=neutral, PublicKeyToken=b88d1754d700e49a"
    Namespace="DevExpress.Web.ASPxGridView.Export" TagPrefix="dx" %>
<%@ Register Assembly="DevExpress.Web.v14.1, Version=14.1.6.0, Culture=neutral, PublicKeyToken=b88d1754d700e49a"
    Namespace="DevExpress.Web.ASPxGridView" TagPrefix="dx" %>
<%@ Register Assembly="DevExpress.Web.v14.1, Version=14.1.6.0, Culture=neutral, PublicKeyToken=b88d1754d700e49a"
    Namespace="DevExpress.Web.ASPxEditors" TagPrefix="dx" %>
<%@ Register Assembly="DevExpress.Web.v14.1, Version=14.1.6.0, Culture=neutral, PublicKeyToken=b88d1754d700e49a"
    Namespace="DevExpress.Web.ASPxGlobalEvents" TagPrefix="dx" %>
<%@ Register Assembly="DevExpress.Web.v14.1, Version=14.1.6.0, Culture=neutral, PublicKeyToken=b88d1754d700e49a"
    Namespace="DevExpress.Web.ASPxRoundPanel" TagPrefix="dx" %>
<%@ Register Assembly="DevExpress.Web.v14.1, Version=14.1.6.0, Culture=neutral, PublicKeyToken=b88d1754d700e49a"
    Namespace="DevExpress.Web.ASPxPanel" TagPrefix="dx" %>
<%@ Register Assembly="DevExpress.Web.v14.1, Version=14.1.6.0, Culture=neutral, PublicKeyToken=b88d1754d700e49a"
    Namespace="DevExpress.Web.ASPxCallback" TagPrefix="dx" %>
<asp:Content ID="Content1" ContentPlaceHolderID="head" runat="server">
    <style type="text/css">
        .dxeHLC, .dxeHC, .dxeHFC
        {
            display: none;
        }
        .style1
        {
            width: 170px;   
        }
        .style2
        {
            width: 120px;   
        }
        .style3
        {
            width: 250px;   
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
            height = height - (height * 65 / 100)
            grid.SetHeight(height);
        }

        function SelectedIndexChangedAff() {
            txtAffiliateName.SetText(cboAffiliateCode.GetSelectedItem().GetColumnText(1));
            lblInfo.SetText('');
        }
        
    </script>
</asp:Content>
<asp:Content ID="Content2" ContentPlaceHolderID="Content" runat="server">
    <table style="width: 100%;">
        <tr>
            <td colspan="16" width="100%">
                <table style="border-width: 1pt thin thin thin; border-style: ridge; border-color:#9598A1; width:100%; height: 70px;">
                    <tr>
                        <td colspan="8" height="20">
                            <table id="Table1" width="100%">
                                <tr>
                                    <td align="left" valign="middle" class="style1">
                                        <dx:ASPxLabel ID="ASPxLabel1" runat="server" Text="PERIOD"
                                            Font-Names="Tahoma" Font-Size="8pt" Width="100%">
                                        </dx:ASPxLabel>
                                    </td>
                                    <td align="left" valign="middle" colspan="2" class="style2">
                                        <dx:ASPxTimeEdit ID="dtPeriodFrom" runat="server"
                                            ClientInstanceName="dtPeriodFrom" DisplayFormatString="yyyy-MM" 
                                            EditFormat="Custom" EditFormatString="yyyy-MM" Width="75px" 
                                            Height="21px">
                                            <ClientSideEvents DateChanged="function(s, e) {
	                                            gridPerformCallback('kosong');
                                            }" />
                                        </dx:ASPxTimeEdit>
                                    </td>
                                    <td width="15">&nbsp;</td>
                                    <td align="left" colspan="3" class="style3">
                                        <table style="width:100%;">
                                            <tr>
                                                <td width="50%">
                                                    <dx:ASPxLabel ID="ASPxLabel4" runat="server" Text="COMMERCIAL"
                                                        Font-Names="Tahoma" Font-Size="8pt" Width="100%">
                                                    </dx:ASPxLabel>
                                                </td>
                                                <td align="left" width="25%">
                                                    <dx:ASPxRadioButton ID="rdrCom1" ClientInstanceName="rdrCom1" runat="server" Text="YES" GroupName="Commercial" Font-Names="Tahoma" Font-Size="8pt">
                                                        <ClientSideEvents LostFocus="function(s, e) {
                                                            lblInfo.SetText('');
                                                        }" />
                                                    </dx:ASPxRadioButton>
                                                </td>
                                                <td align="left" width="25%">
                                                    <dx:ASPxRadioButton ID="rdrCom2" ClientInstanceName="rdrCom2" runat="server" Text="NO" GroupName="Commercial" Font-Names="Tahoma" Font-Size="8pt">
                                                        <ClientSideEvents LostFocus="function(s, e) {
                                                            lblInfo.SetText('');                                                         
                                                        }" />
                                                    </dx:ASPxRadioButton>
                                                </td>
                                            </tr>
                                        </table>
                                    </td>
                                    <td width="15">&nbsp;</td>
                                    <td align="left" colspan="3" class="style3">
                                        <table style="width:100%;">
                                            <tr>
                                                <td align="left" valign="middle" colspan="2" width="50%">
                                                    <dx:ASPxLabel ID="ASPxLabel32" runat="server" Text="DELIVERY LOCATION"
                                                        Font-Names="Tahoma" Font-Size="8pt" width="100%">
                                                    </dx:ASPxLabel>
                                                </td>
                                                <td align="left" valign="middle" colspan="2" width="50%">
                                                    <dx:ASPxComboBox ID="cboDelLoc" width="100%" runat="server" 
                                                        Font-Size="8pt" Font-Names="Tahoma" TextFormatString="{0}" 
                                                        ClientInstanceName="cboDelLoc" TabIndex="3">
                                                        <ClientSideEvents SelectedIndexChanged="function(s, e) {
	                                                        txtDelLoc.SetText(cboDelLoc.GetSelectedItem().GetColumnText(1));
	                                                        lblInfo.SetText('');
                                                        }" />
                                                        <LoadingPanelStyle ImageSpacing="5px">
                                                        </LoadingPanelStyle>
                                                    </dx:ASPxComboBox>
                                                </td>
                                            </tr>
                                        </table>
                                    </td>
                                    <td align="left" valign="middle" colspan="3" class="style3">
                                        <dx:ASPxTextBox ID="txtDelLoc" runat="server" Width="100%" 
                                            ClientInstanceName="txtDelLoc" Font-Names="Tahoma" Font-Size="8pt"
                                            ReadOnly="True" MaxLength="100" Height="20px">
                                            <readonlystyle backcolor="#CCCCCC">
                                            </readonlystyle>
                                        </dx:ASPxTextBox>
                                    </td>
                                    <td>&nbsp;</td>
                                </tr>

                                <tr>
                                    <td align="left" valign="middle" class="style1">
                                        <dx:ASPxLabel ID="ASPxLabel5" runat="server" Text="PO MONTHLY /EMERGENCY"
                                            Font-Names="Tahoma" Font-Size="8pt" width="100%">
                                        </dx:ASPxLabel>
                                    </td>
                                    <td align="left" colspan="3" class="style2">
                                        <table style="width:100%;">
                                            <tr>
                                                <td align="left" valign="middle" height="20px" width="50%">
                                                    <dx:ASPxRadioButton ID="rdMonthly" runat="server" 
                                                        ClientInstanceName="rdMonthly" Text="M">
                                                    </dx:ASPxRadioButton>
                                                </td>
                                                <td align="left" valign="middle" height="20px" width="50%">
                                                    <dx:ASPxRadioButton ID="rdEmergency" runat="server" 
                                                        ClientInstanceName="rdEmergency" Text="E">
                                                    </dx:ASPxRadioButton>
                                                </td>
                                            </tr>
                                        </table>
                                    </td>
                                    <td align="left" colspan="3" class="style3">
                                        <table style="width:100%;">
                                            <tr>
                                                <td width="50%">
                                                    <dx:ASPxLabel ID="ASPxLabel2" runat="server" Text="SHIP BY"
                                                        Font-Names="Tahoma" Font-Size="8pt" Width="100%">
                                                    </dx:ASPxLabel>
                                                </td>
                                                <td width="25%">
                                                    <dx:ASPxRadioButton ID="rdrShipBy2" ClientInstanceName="rdrShipBy2" runat="server" 
                                                        Text="BOAT" GroupName="POEm" Font-Names="Tahoma" Font-Size="8pt">
                                                        <ClientSideEvents LostFocus="function(s, e) {                                                        
                                                            lblInfo.SetText('');
                                                        }" />
                                                    </dx:ASPxRadioButton>
                                                </td>   
                                                <td width="25%">
                                                    <dx:ASPxRadioButton ID="rdrShipBy3" ClientInstanceName="rdrShipBy3" runat="server" 
                                                        Text="AIR" GroupName="POEm" Font-Names="Tahoma" Font-Size="8pt">
                                                        <ClientSideEvents LostFocus="function(s, e) {
                                                            lblInfo.SetText('');                                                         
                                                        }" />
                                                    </dx:ASPxRadioButton>
                                                </td>
                                            </tr>
                                        </table>
                                    </td>
                                    <td>&nbsp;</td>
                                    <td>&nbsp;</td>
                                </tr>
                                        
                                <tr>
                                    <td align="left" valign="middle" class="style1">
                                        <dx:ASPxLabel ID="ASPxLabel20" runat="server" Text="AFFILIATE CODE/NAME"
                                            Font-Names="Tahoma" Font-Size="8pt">
                                        </dx:ASPxLabel>
                                    </td>
                                    <td align="left" valign="middle" colspan="2" class="style2">
                                        <dx:ASPxComboBox ID="cboAffiliate" width="100%" runat="server" 
                                            Font-Size="8pt" Font-Names="Tahoma" TextFormatString="{0}" 
                                            ClientInstanceName="cboAffiliate" TabIndex="3">
                                            <ClientSideEvents SelectedIndexChanged="function(s, e) {
	                                            txtAffiliate.SetText(cboAffiliate.GetSelectedItem().GetColumnText(1));
                                                txtconsignee.SetText(cboAffiliate.GetSelectedItem().GetColumnText(2));
	                                            lblInfo.SetText('');
                                            }" />
                                            <LoadingPanelStyle ImageSpacing="5px">
                                            </LoadingPanelStyle>
                                        </dx:ASPxComboBox>
                                    </td>
                                    <td>&nbsp;</td>
                                    <td align="left" colspan="3" class="style3">
                                        <dx:ASPxTextBox ID="txtAffiliate" runat="server" Width="100%" 
                                            ClientInstanceName="txtAffiliate" Font-Names="Tahoma" Font-Size="8pt"
                                            ReadOnly="True" MaxLength="100" Height="20px">
                                            <readonlystyle backcolor="#CCCCCC">
                                            </readonlystyle>
                                        </dx:ASPxTextBox>
                                    </td>
                                    <td>&nbsp;</td>
                                    <td align="left" valign="middle" class="style3">
                                        <dx:ASPxTextBox ID="txtconsignee" runat="server" Width="100%" 
                                            ClientInstanceName="txtconsignee" Font-Names="Tahoma" Font-Size="8pt"
                                            ReadOnly="True" MaxLength="100" Height="20px" ForeColor="White">
                                            <Border BorderStyle="None" />
                                        </dx:ASPxTextBox>
                                    </td>
                                </tr>
                                
                                <tr>
                                    <td align="left" valign="middle" class="style1">
                                        <dx:ASPxLabel ID="ASPxLabel34" runat="server" Text="ORDER NO"
                                            Font-Names="Tahoma" Font-Size="8pt">
                                        </dx:ASPxLabel>
                                    </td>
                                    <td align="left" valign="middle" colspan="2" class="style2">
                                        <dx:ASPxTextBox ID="txtpono" runat="server" Width="100%" 
                                            ClientInstanceName="txtpono">
                                        </dx:ASPxTextBox>
                                    </td>
                                    <td>&nbsp;</td>
                                    <td align="left" colspan="3" class="style3">
                                        <table style="width:100%;">
                                            <tr>
                                                <td align="left" valign="middle" width="50%">
                                                    <dx:ASPxLabel ID="ASPxLabel35" runat="server" Text="ETD VENDOR"
                                                        Font-Names="Tahoma" Font-Size="8pt" Width="100%">
                                                    </dx:ASPxLabel>
                                                </td>
                                                <td align="right" valign="middle" width="50%">
                                                    <dx:ASPxDateEdit ID="dtETDVendor" runat="server" Width="100%" 
                                                        ClientInstanceName="dtETDVendor" DisplayFormatString="yyyy-MM-dd" 
                                                        EditFormat="Custom" EditFormatString="yyyy-MM-dd">
                                                    </dx:ASPxDateEdit>
                                                </td>
                                            </tr>
                                        </table>
                                    </td>
                                    <td>&nbsp;</td>
                                    <td align="left" colspan="3" class="style3">
                                        <table style="width:100%;">
                                            <tr>
                                                <td align="left" width="50%">
                                                    <dx:ASPxLabel ID="ASPxLabel36" runat="server" Text="ETD PORT"
                                                        Font-Names="Tahoma" Font-Size="8pt" Width="100%">
                                                    </dx:ASPxLabel>
                                                </td>
                                                <td align="right" valign="middle" width="50%">
                                                    <dx:ASPxDateEdit ID="dtETDPort" runat="server" Width="100%" 
                                                        ClientInstanceName="dtETDPort" DisplayFormatString="yyyy-MM-dd" 
                                                        EditFormat="Custom" EditFormatString="yyyy-MM-dd">
                                                    </dx:ASPxDateEdit>
                                                </td>
                                            </tr>
                                        </table>
                                    </td>
                                    <td align="left" colspan="3" class="style3">
                                        <table style="width:100%;">
                                            <tr>
                                                <td align="left" valign="middle" width="50%">
                                                    <dx:ASPxLabel ID="ASPxLabel37" runat="server" Text="ETA PORT"
                                                        Font-Names="Tahoma" Font-Size="8pt" Width="100%">
                                                    </dx:ASPxLabel>
                                                </td>
                                                <td align="right" valign="middle" width="50%">
                                                    <dx:ASPxDateEdit ID="dtETAPort" runat="server" Width="100%" 
                                                        ClientInstanceName="dtETAPort" DisplayFormatString="yyyy-MM-dd" 
                                                        EditFormat="Custom" EditFormatString="yyyy-MM-dd">
                                                    </dx:ASPxDateEdit>
                                                </td>
                                            </tr>
                                        </table>
                                    </td>
                                    <td align="left" colspan="3" class="style3">
                                        <table style="width:100%;">
                                            <tr>
                                                <td align="left" valign="middle" width="50%">
                                                    <dx:ASPxLabel ID="ASPxLabel38" runat="server" Text="ETA FACTORY"
                                                        Font-Names="Tahoma" Font-Size="8pt" Width="100%">
                                                    </dx:ASPxLabel>
                                                </td>
                                                <td align="right" valign="middle" width="50%">
                                                    <dx:ASPxDateEdit ID="dtETAFactory" runat="server" Width="100%" 
                                                        ClientInstanceName="dtETAFactory" DisplayFormatString="yyyy-MM-dd" 
                                                        EditFormat="Custom" EditFormatString="yyyy-MM-dd">
                                                    </dx:ASPxDateEdit>
                                                </td>
                                            </tr>
                                        </table>
                                    </td>
                                    <td>&nbsp;</td>
                                </tr>                                                                                 
                                <tr>
                                    <td align="left" valign="middle" class="style1">
                                        <dx:ASPxLabel ID="ASPxLabel39" runat="server" Text="ORIGINAL O/NO"
                                            Font-Names="Tahoma" Font-Size="8pt">
                                        </dx:ASPxLabel>
                                    </td>
                                    <td align="left" valign="middle" colspan="2" class="style2">
                                            <dx:ASPxTextBox ID="txtOrderNo" runat="server" Width="100%" 
                                            ClientInstanceName="txtOrderNo">
                                        </dx:ASPxTextBox>
                                    </td>
                                </tr>                               
                            </table>
                        </td>
                    </tr>
                </table>
            </td>            
        </tr>
    </table>
    <div style="height: 1px;">
    </div>
    <table style="width: 100%; height: 15px;">
        <tr>
            <td colspan="8" height="15">
                <table id="tblMsg" style="border-width: thin; border-style: inset hidden ridge hidden;
                    border-color: #9598A1; width: 100%; height: 15px;">
                    <tr>
                        <td align="left" valign="middle" height="15px">
                            <dx:ASPxLabel ID="lblInfo" runat="server" Text="[lblinfo]" Font-Names="Tahoma" ClientInstanceName="lblInfo"
                                Font-Bold="True" Font-Italic="True" Font-Size="8pt">
                            </dx:ASPxLabel>
                        </td>
                    </tr>
                </table>
            </td>
        </tr>
    </table>
    <div style="height: 1px;">
    </div>
    <table style="width: 100%;">
        <tr>
            <td valign="top" align="left">
                &nbsp;
            </td>
            <td valign="top" align="left" width="30px">
                &nbsp;
            </td>
            <td valign="top" align="left" width="200px">
                &nbsp;
            </td>
            <td valign="top" align="left" width="200px">
                &nbsp;
            </td>
            <td valign="top" align="left" width="200px">
                &nbsp;
            </td>
            <td valign="top" align="right" width="30px">
                <table style="width: 100%;">
                    <tr>                        
                        <td width="3px">
                            &nbsp;
                        </td>
                        <td align="right" valign="middle" width="30px">
                            <asp:TextBox ID="lSuuplier0" runat="server" BackColor="GreenYellow" BorderStyle="None"
                                ReadOnly="True" Width="30px"></asp:TextBox>
                        </td>
                        <td align="right" valign="middle" width="140px">
                            <dx:ASPxLabel ID="ASPxLabel24" runat="server" Font-Names="Tahoma" Font-Size="8pt"
                                Text=": EDIT BY SUPPLIER" Width="140px">
                            </dx:ASPxLabel>
                        </td>
                    </tr>
                </table>
            </td>
        </tr>
    </table>
    <div style="height: 1px;">
    </div>
    <table style="width: 100%;">
        <tr>
            <td colspan="8" align="left" valign="top" height="220">
                <%--Column : Grid--%>
                <dx:ASPxGridView ID="grid" runat="server" Width="100%"
                    Font-Names="Tahoma" KeyFieldName="NoUrutAbal;NoUrut;PartNo1"
                    AutoGenerateColumns="False"
                    ClientInstanceName="grid" Font-Size="8pt">
                    <ClientSideEvents Init="OnInit" 
                        CallbackError="function(s, e) {e.handled = true;}" 
                        EndCallback="function(s, e) {
                        var pMsg = s.cpMessage;
                        if (pMsg != '') {
                            if (pMsg.substring(1,5) == '1001' || pMsg.substring(1,5) == '1002' || pMsg.substring(1,5) == '1003' || pMsg.substring(1,5) == '2001') {
                                lblInfo.GetMainElement().style.color = 'Blue';  
                            } else {
                                lblInfo.GetMainElement().style.color = 'Red';
                            }
                            lblInfo.SetText(pMsg);
                        } else {
                            lblInfo.SetText('');
                        }
                        delete s.cpMessage;
                    }" RowClick="function(s, e) {
	                    lblInfo.SetText('');
                    }" />                    
                    <Columns>
                        <dx:GridViewDataTextColumn VisibleIndex="1" Caption="NO." FieldName="NoUrut" 
                            Name="NoUrut"
                            Width="30px" HeaderStyle-HorizontalAlign="Center">                            
                            <HeaderStyle HorizontalAlign="Center" Wrap="True"></HeaderStyle>
                            <CellStyle Font-Names="Tahoma" Font-Size="8pt" HorizontalAlign="Left">
                            </CellStyle>
                        </dx:GridViewDataTextColumn>
                        <dx:GridViewDataTextColumn VisibleIndex="1" Caption="NO." FieldName="NoUrutAbal" 
                            Name="NoUrut"
                            Width="0px" HeaderStyle-HorizontalAlign="Center">                            
                            <HeaderStyle HorizontalAlign="Center" Wrap="True"></HeaderStyle>
                            <CellStyle Font-Names="Tahoma" Font-Size="8pt" HorizontalAlign="Left">
                            </CellStyle>
                        </dx:GridViewDataTextColumn>
                        <dx:GridViewDataTextColumn VisibleIndex="2" Caption="PART NO." FieldName="PartNo" Width="90px" 
                        Name="PartNo" HeaderStyle-HorizontalAlign="Center">                            
                            <HeaderStyle HorizontalAlign="Center" Wrap="True"></HeaderStyle>
                            <CellStyle Font-Names="Tahoma" Font-Size="8pt">
                            </CellStyle>
                            </dx:GridViewDataTextColumn>
                        <dx:GridViewDataTextColumn VisibleIndex="2" Caption="PART NO." FieldName="PartNo1" Width="0px" 
                        Name="PartNo" HeaderStyle-HorizontalAlign="Center">                            
                            <HeaderStyle HorizontalAlign="Center" Wrap="True"></HeaderStyle>
                            <CellStyle Font-Names="Tahoma" Font-Size="8pt">
                            </CellStyle>
                            </dx:GridViewDataTextColumn>
                        <dx:GridViewDataTextColumn VisibleIndex="3" Caption="PART NAME" FieldName="PartName" Width="180px" 
                        Name="PartName" HeaderStyle-HorizontalAlign="Center">                            
                            <HeaderStyle HorizontalAlign="Center"></HeaderStyle>
                            <CellStyle Font-Names="Tahoma" Font-Size="8pt">
                            </CellStyle>
                        </dx:GridViewDataTextColumn>
                        <dx:GridViewDataTextColumn VisibleIndex="4" Caption="UOM" Width="35px" 
                            HeaderStyle-HorizontalAlign="Center" CellStyle-HorizontalAlign="Center" 
                            FieldName="UnitDesc" Name="Description">                            
                            <HeaderStyle HorizontalAlign="Center"></HeaderStyle>
                            <CellStyle Font-Names="Tahoma" Font-Size="8pt">
                            </CellStyle>
                        </dx:GridViewDataTextColumn>
                        <dx:GridViewDataTextColumn VisibleIndex="5" Caption="MOQ" FieldName="MOQ" Name="MOQ"
                            Width="60px" HeaderStyle-HorizontalAlign="Center">                            
                            <PropertiesTextEdit DisplayFormatString="{0:n0}">
                            </PropertiesTextEdit>
                            <HeaderStyle HorizontalAlign="Center"></HeaderStyle>
                            <CellStyle Font-Names="Tahoma" Font-Size="8pt">
                            </CellStyle>
                        </dx:GridViewDataTextColumn>
                        <dx:GridViewDataTextColumn VisibleIndex="6" Caption="QTY BOX" FieldName="QtyBox" Name="QtyBox"
                            Width="75px" HeaderStyle-HorizontalAlign="Center">                            
                            <PropertiesTextEdit DisplayFormatString="{0:n0}">
                            </PropertiesTextEdit>
                            <HeaderStyle HorizontalAlign="Center"></HeaderStyle>
                            <CellStyle Font-Names="Tahoma" Font-Size="8pt">
                            </CellStyle>
                        </dx:GridViewDataTextColumn>
                        <dx:GridViewDataTextColumn VisibleIndex="6" Caption=" " FieldName="AffiliateName" Width="180px" 
                        Name="AffiliateName" HeaderStyle-HorizontalAlign="Center">                            
                            <HeaderStyle HorizontalAlign="Center"></HeaderStyle>
                            <CellStyle Font-Names="Tahoma" Font-Size="8pt">
                            </CellStyle>
                        </dx:GridViewDataTextColumn>
                        <dx:GridViewDataTextColumn VisibleIndex="7" Caption="TOTAL FIRM QTY *" 
                            FieldName="POQty" Name="POQty" Width="60px" HeaderStyle-HorizontalAlign="Center">                            
                            <PropertiesTextEdit DisplayFormatString="{0:n0}" MaxLength="18">
                            </PropertiesTextEdit>
                            <HeaderStyle Font-Names="Tahoma" Font-Size="8pt" Font-Underline="False" 
                                Wrap="True" HorizontalAlign="Center"/>
                            <CellStyle Font-Names="Tahoma" Font-Size="8pt">
                            </CellStyle>
                        </dx:GridViewDataTextColumn>
                        <dx:GridViewDataTextColumn VisibleIndex="8" Caption="PREVIOUS FORECAST N" 
                            Width="75px" HeaderStyle-HorizontalAlign="Center" 
                            FieldName="PreviousForecast" Name="PreviousForecast">                            
                            <PropertiesTextEdit DisplayFormatString="{0:n0}">
                            </PropertiesTextEdit>
                            <HeaderStyle Font-Names="Tahoma" Font-Size="8pt" Font-Underline="False" 
                                Wrap="True" HorizontalAlign="Center"/>
                            <CellStyle Font-Names="Tahoma" Font-Size="8pt">
                            </CellStyle>
                        </dx:GridViewDataTextColumn>
                        <dx:GridViewDataTextColumn VisibleIndex="9" Caption="VARIANCE" Width="70px"
                            HeaderStyle-HorizontalAlign="Center" CellStyle-HorizontalAlign="Center" 
                            FieldName="Variance" Name="Variance">
                            <PropertiesTextEdit DisplayFormatString="{0:n0}">
                            </PropertiesTextEdit>                            
                            <HeaderStyle HorizontalAlign="Center"></HeaderStyle>
                            <CellStyle Font-Names="Tahoma" Font-Size="8pt" HorizontalAlign="Right">
                            </CellStyle>
                        </dx:GridViewDataTextColumn>
                        <dx:GridViewDataTextColumn VisibleIndex="10" Caption="VARIANCE %" Width="70px"
                            FieldName="VarPecentage" Name="VarPecentage">
                            <HeaderStyle HorizontalAlign="Center"></HeaderStyle>
                            <PropertiesTextEdit DisplayFormatString="{0:n2}">
                            </PropertiesTextEdit>
                            <CellStyle Font-Names="Tahoma" Font-Size="8pt" HorizontalAlign="Right">
                            </CellStyle>
                        </dx:GridViewDataTextColumn>
                         <dx:GridViewDataTextColumn VisibleIndex="11" Caption="FORECAST N+1" 
                            FieldName="Forecast1" Name="Forecast1" Width="60px" HeaderStyle-HorizontalAlign="Center">                            
                            <PropertiesTextEdit DisplayFormatString="{0:n0}" MaxLength="18">
                            </PropertiesTextEdit>
                            <HeaderStyle Font-Names="Tahoma" Font-Size="8pt" Font-Underline="False" 
                                Wrap="True" HorizontalAlign="Center"/>
                            <CellStyle Font-Names="Tahoma" Font-Size="8pt">
                            </CellStyle>
                        </dx:GridViewDataTextColumn>
                        <dx:GridViewDataTextColumn VisibleIndex="12" Caption="FORECAST N+2" 
                            Width="75px" HeaderStyle-HorizontalAlign="Center" FieldName="Forecast2" Name="Forecast2">
                            <HeaderStyle Font-Names="Tahoma" Font-Size="8pt" Font-Underline="False" 
                                Wrap="True" HorizontalAlign="Center"/>
                            <CellStyle Font-Names="Tahoma" Font-Size="8pt">
                            </CellStyle>
                            <PropertiesTextEdit DisplayFormatString="{0:n0}" MaxLength="18">
                            <MaskSettings ErrorText="Please input valid value !"
                            Mask="<0..999999g>" IncludeLiterals="DecimalSymbol" />
                            </PropertiesTextEdit>
                        </dx:GridViewDataTextColumn>
                        <dx:GridViewDataTextColumn VisibleIndex="13" Caption="FORECAST N+3" 
                            Width="75px" HeaderStyle-HorizontalAlign="Center" FieldName="Forecast3" Name="Forecast3">
                            <HeaderStyle Font-Names="Tahoma" Font-Size="8pt" Font-Underline="False" 
                                Wrap="True" HorizontalAlign="Center"/>
                            <CellStyle Font-Names="Tahoma" Font-Size="8pt">
                            </CellStyle>
                            <PropertiesTextEdit DisplayFormatString="{0:n0}" MaxLength="18">
                            <MaskSettings ErrorText="Please input valid value !"
                            Mask="<0..999999g>" IncludeLiterals="DecimalSymbol" />
                            </PropertiesTextEdit>
                        </dx:GridViewDataTextColumn>
                        <dx:GridViewDataTextColumn Caption="SUPPLIER CODE" FieldName="SupplierID" 
                            Name="SupplierID"
                            VisibleIndex="14" Width="100px">
                        </dx:GridViewDataTextColumn>
                        <%--<dx:GridViewDataTextColumn Caption="AFFILIATE CODE" FieldName="AffiliateID" 
                        Name="AffiliateID"
                            VisibleIndex="13" Width="0px">
                        </dx:GridViewDataTextColumn>--%>
                        <%--<dx:GridViewDataTextColumn Caption="SUPPLIER CODE" FieldName="SupplierID" 
                        Name="SupplierID"
                            VisibleIndex="14" Width="0px">
                        </dx:GridViewDataTextColumn>--%>
                        <%--<dx:GridViewDataTextColumn FieldName="PONo" Width="0px" Caption="PO NO." 
                        Name="PONo"
                            VisibleIndex="17">
                            <HeaderStyle HorizontalAlign="Center"></HeaderStyle>
                            <CellStyle Font-Names="Tahoma" Font-Size="8pt"></CellStyle>
                        </dx:GridViewDataTextColumn>--%>
                       <%-- <dx:GridViewDataTextColumn Caption="AdaData" FieldName="AdaData" 
                            Name="AdaData" VisibleIndex="20" Width="0px">
                            <HeaderStyle Font-Names="Tahoma" Font-Size="8pt" Font-Underline="False" 
                                Wrap="True" HorizontalAlign="Center"/>
                            <CellStyle Font-Names="Tahoma" Font-Size="8pt">
                            </CellStyle>
                        </dx:GridViewDataTextColumn>--%>
                    </Columns>
                    <SettingsPager Visible="False" Position="Top">
                        <Summary Text="Page {0} of {1} [{2} record(s)]" AllPagesText="Page {0} of {1} " />
                        <Summary AllPagesText="Page {0} of {1} " Text="Page {0} of {1} [{2} record(s)]">
                        </Summary>
                    </SettingsPager>
                    <SettingsEditing Mode="Batch">
                    </SettingsEditing>
                    <Settings ShowHorizontalScrollBar="True" ShowVerticalScrollBar="True" ShowGroupButtons="False"
                        ShowStatusBar="Hidden" VerticalScrollableHeight="250" />

                    <Settings ShowGroupButtons="False" ShowVerticalScrollBar="True" ShowHorizontalScrollBar="True" 
                        VerticalScrollableHeight="250" ShowStatusBar="Hidden"></Settings>
                    <Styles>
                                <Header BackColor="#FFD2A6" Font-Names="Verdana" Font-Size="8pt"></Header>
                                <Row BackColor="#FFFFE1" Font-Names="Verdana" Font-Size="8pt" Wrap="False"></Row>
                                <RowHotTrack BackColor="#E8EFFD" Font-Names="Verdana" Font-Size="8pt" Wrap="False"></RowHotTrack>
                                <SelectedRow Wrap="False">
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
                <dx:ASPxButton ID="btnSubMenu" runat="server" Text="SUB MENU" Font-Names="Tahoma"
                    Width="90px" Font-Size="8pt">
                </dx:ASPxButton>
            </td>
            <td valign="top" align="right" style="width: 50px;">
            </td>
            <td valign="top" align="right" style="width: 50px;">
            </td>            
            <td valign="top" align="right" style="width: 50px;">
                <dx:ASPxButton ID="btnFinalApp" runat="server" Text="FINAL APPROVE" Font-Names="Tahoma" Width="90px"
                    ClientInstanceName="btnSubmit" AutoPostBack="False" Font-Size="8pt">
                    <ClientSideEvents Click="function(s, e) {
                      grid.PerformCallback('save');
                    }" />
                </dx:ASPxButton>
            </td>
        </tr>
    </table>
    <dx:ASPxGlobalEvents ID="ge" runat="server">
        <ClientSideEvents ControlsInitialized="function(s, e) {
	        OnControlsInitializedSplitter();
	        OnControlsInitializedGrid();
        }" />
    </dx:ASPxGlobalEvents>    
</asp:Content>
