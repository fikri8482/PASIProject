<%@ Page Title="" Language="vb" AutoEventWireup="false" MasterPageFile="~/SiteContent.Master"
    CodeBehind="POExportFinalApprovalEmergency.aspx.vb" Inherits="PASISystem.POExportFinalApprovalEmergency" %>

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
            height: 16px;
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
            height = height - (height * 63 / 100)
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
            <td valign="top" width="60%">
                <table style="border-width: 1pt thin thin thin; border-style: ridge; border-color: #9598A1;
                    width: 100%;">
                    <tr>
                        <td height="100">
                            <table id="Table1">
                                <tr style="height: 15px" >
                                    <td style="width: 5px; height:15px;">
                                        &nbsp;
                                    </td>
                                    <td align="left" valign="middle" width="160px" height="15px">
                                        <dx:ASPxLabel ID="ASPxLabel3" runat="server" Text="PERIOD" Font-Names="Tahoma" Font-Size="8pt"
                                            Width="160px">
                                        </dx:ASPxLabel>
                                    </td>
                                    <td align="left" valign="middle" width="120px" height="15px">
                                        <dx:ASPxTimeEdit ID="dtPeriodFrom" runat="server" ClientInstanceName="dtPeriodFrom"
                                            DisplayFormatString="yyyy-MM" EditFormat="Custom" EditFormatString="yyyy-MM"
                                            Width="110px" Font-Names="Tahoma" Font-Size="8pt" ReadOnly="True"  
                                            height="15px">
                                            <ReadOnlyStyle BackColor="#CCCCCC">
                                            </ReadOnlyStyle>
                                        </dx:ASPxTimeEdit>
                                    </td>
                                    <td align="left" valign="middle" width="80px"  height="15px">
                                        <dx:ASPxLabel ID="ASPxLabel7" runat="server" Text="COMMERCIAL" Font-Names="Tahoma"
                                            Font-Size="8pt" Width="80px" height="15px">
                                        </dx:ASPxLabel>
                                    </td>
                                    <td align="left" valign="middle" width="150px" height="15px">
                                        <dx:ASPxRadioButtonList ID="rblCommercial" runat="server" RepeatDirection="Horizontal"
                                            Width="130px" ClientInstanceName="rblCommercial" SelectedIndex="0" TabIndex="9"
                                            Font-Names="Tahoma" Font-Size="8pt" ReadOnly="True" height="15px">
                                            <RadioButtonStyle HorizontalAlign="Left">
                                            </RadioButtonStyle>
                                            <Items>                                                
                                                <dx:ListEditItem Text="YES" Value="1" />
                                                <dx:ListEditItem Text="NO" Value="0" />
                                            </Items>
                                            <Border BorderStyle="None"></Border>
                                        </dx:ASPxRadioButtonList>
                                    </td>
                                    <td align="left" valign="middle" width="120px" height="15px">
                                        <dx:ASPxLabel ID="ASPxLabel28" runat="server" Text="DELIVERY LOCATION" Font-Names="Tahoma"
                                            Font-Size="8pt" Width="120px" height="15px">
                                        </dx:ASPxLabel>
                                    </td>
                                    <td align="left" valign="middle" width="120px" height="15px">
                                        <dx:ASPxComboBox ID="cboDeliveryLoc" runat="server" TextFormatString="{0}" DropDownStyle="DropDown"
                                            Width="120px" ClientInstanceName="cboDeliveryLoc" IncrementalFilteringMode="StartsWith"
                                            Font-Names="Tahoma" Font-Size="8pt" MaxLength="10" TabIndex="4" 
                                            ReadOnly="True" height="15px">
                                            <ClientSideEvents SelectedIndexChanged="function(s, e) {SelectedIndexChangedAff();}"
                                                LostFocus="function(s, e) { lblInfo.SetText(&quot;&quot;); }" />
                                            <ReadOnlyStyle BackColor="#CCCCCC">
                                            </ReadOnlyStyle>
                                        </dx:ASPxComboBox>
                                    </td>
                                    <td align="left" valign="middle" width="170px" height="15px">
                                        <dx:ASPxTextBox ID="txtDeliveryLoc" runat="server" ClientInstanceName="txtDeliveryLoc"
                                            Width="170px" Height="15px" ReadOnly="True" TabIndex="5" Font-Names="Tahoma"
                                            Font-Size="8pt">
                                            <ReadOnlyStyle BackColor="#CCCCCC">
                                            </ReadOnlyStyle>
                                        </dx:ASPxTextBox>
                                    </td>
                                </tr>
                                <tr style="height: 15px" >
                                    <td style="width: 5px; height: 15px;">
                                        &nbsp;
                                    </td>
                                    <td align="left" valign="middle" width="160px" height="15px">
                                        <dx:ASPxLabel ID="ASPxLabel2" runat="server" Text="PO MONTHLY / EMERGENCY" Font-Names="Tahoma"
                                            Font-Size="8pt" Width="160px" height="15px">
                                        </dx:ASPxLabel>
                                    </td>
                                    <td align="left" valign="middle" width="120px" height="15px">
                                        <dx:ASPxTextBox ID="txtPOEmergency" runat="server" ClientInstanceName="txtPOEmergency"
                                            Width="110px" Height="15px" ReadOnly="True" TabIndex="5" Font-Names="Tahoma"
                                            Font-Size="8pt">
                                            <ReadOnlyStyle BackColor="#CCCCCC">
                                            </ReadOnlyStyle>
                                        </dx:ASPxTextBox>
                                    </td>
                                    <td align="left" valign="middle" width="80px" height="15px">
                                        <dx:ASPxLabel ID="ASPxLabel27" runat="server" Text="SHIP BY" Font-Names="Tahoma"
                                            Font-Size="8pt" Width="80px" height="15px">
                                        </dx:ASPxLabel>
                                    </td>
                                    <td align="left" valign="middle" width="150px" height="15px">
                                        <dx:ASPxTextBox ID="txtShipBy" runat="server" ClientInstanceName="txtShipBy" Width="130px"
                                            Height="15px" ReadOnly="True" TabIndex="5" Font-Names="Tahoma" 
                                            Font-Size="8pt">
                                            <ReadOnlyStyle BackColor="#CCCCCC">
                                            </ReadOnlyStyle>
                                        </dx:ASPxTextBox>
                                    </td>
                                    <td align="left" valign="middle" width="120px" height="15px">
                                        <dx:ASPxLabel ID="ASPxLabel4" runat="server" Font-Names="Tahoma" 
                                            Font-Size="8pt" Text="SUPPLIER REMARKS" height="15px" Width="120px">
                                        </dx:ASPxLabel>
                                    </td>
                                    <td align="left" valign="top" width="280px" colspan="2" rowspan="2">
                                        <dx:ASPxMemo ID="txtRemarks" runat="server" ClientInstanceName="txtRemarks" 
                                            Font-Names="Tahoma" Font-Size="8" Height="50px" MaxLength="200" ReadOnly="True" 
                                            Width="280px">
                                            <readonlystyle backcolor="#CCCCCC">
                                            </readonlystyle>
                                        </dx:ASPxMemo>
                                    </td>
                                </tr>
                                <tr style="height: 15px" >
                                    <td style="width: 5px; height: 15px;">
                                        &nbsp;
                                    </td>
                                    <td align="left" valign="middle" width="160px" height="15px">
                                        <dx:ASPxLabel ID="ASPxLabel1" runat="server" Text="AFFILIATE CODE/NAME" Font-Names="Tahoma"
                                            Font-Size="8pt" Width="160px" height="15px">
                                        </dx:ASPxLabel>
                                    </td>
                                    <td align="left" valign="middle" width="110px" height="15px">
                                        <dx:ASPxComboBox ID="cboAffiliateCode" runat="server" TextFormatString="{0}" DropDownStyle="DropDown"
                                            Width="110px" ClientInstanceName="cboAffiliateCode" IncrementalFilteringMode="StartsWith"
                                            Font-Names="Tahoma" Font-Size="8pt" MaxLength="10" TabIndex="4" 
                                            ReadOnly="True" height="15px">
                                            <ClientSideEvents SelectedIndexChanged="function(s, e) {SelectedIndexChangedAff();}"
                                                LostFocus="function(s, e) { lblInfo.SetText(&quot;&quot;); }" />
                                            <ReadOnlyStyle BackColor="#CCCCCC">
                                            </ReadOnlyStyle>
                                        </dx:ASPxComboBox>
                                    </td>
                                    <td align="left" valign="middle" colspan="2" width="230px" height="15px">
                                        <dx:ASPxTextBox ID="txtAffiliateName" runat="server" ClientInstanceName="txtAffiliateName"
                                            Width="210px" Height="15px" ReadOnly="True" TabIndex="5" Font-Names="Tahoma"
                                            Font-Size="8pt">
                                            <ReadOnlyStyle BackColor="#CCCCCC">
                                            </ReadOnlyStyle>
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
            <td colspan="8" align="center" valign="top">
                <table style="width: 100%;">
                    <tr>
                        <td width="150px">
                            &nbsp;
                        </td>
                        <td width="150px">
                            &nbsp;
                        </td>
                        <td rowspan="5">
                            <table border="0" cellpadding="0" cellspacing="0" width="250px">                                
                                <tr>
                                    <td width="150px" height="16px">
                                        <dx:ASPxTextBox runat="server" Width="150px" MaxLength="10" Height="16px" Font-Names="Tahoma"
                                            Font-Size="8pt" ID="ASPxTextBox5" BackColor="#FFD2A6" Text="ORDER NO" ReadOnly="True"
                                            Font-Bold="True">
                                        </dx:ASPxTextBox>
                                    </td>
                                    <td width="100px" height="16px">
                                        <dx:ASPxTextBox runat="server" Width="100px" MaxLength="10" Height="16px" ClientInstanceName="txtOrderNoWeek1"
                                            Font-Names="Tahoma" Font-Size="8pt" ID="txtOrderNoWeek1"
                                            ReadOnly="True">
                                            <ReadOnlyStyle BackColor="#CCCCCC">
                                            </ReadOnlyStyle>
                                        </dx:ASPxTextBox>
                                    </td>
                                </tr>
                                <tr>
                                    <td width="150px" height="16px">
                                        <dx:ASPxTextBox runat="server" Width="150px" MaxLength="10" Height="16px" Font-Names="Tahoma"
                                            Font-Size="8pt" ID="ASPxTextBox6" BackColor="#FFD2A6" Text="ETD VENDOR (ORDER)" ReadOnly="True"
                                            Font-Bold="True">
                                        </dx:ASPxTextBox>
                                    </td>
                                    <td width="100px" height="16px">
                                        <dx:ASPxDateEdit ID="dtWeekETDVendorOld1" runat="server" ClientInstanceName="dtWeekETDVendorOld1"
                                            DisplayFormatString="yyyy-MM-dd" EditFormat="Custom" EditFormatString="yyyy-MM-dd"
                                            Font-Names="Tahoma" Font-Size="8pt" Width="100px" ReadOnly="True">
                                            <ReadOnlyStyle BackColor="#CCCCCC">
                                            </ReadOnlyStyle>                                            
                                        </dx:ASPxDateEdit>
                                    </td>
                                </tr>                                
                                <tr>
                                    <td width="150px" height="16px">
                                        <dx:ASPxTextBox runat="server" Width="150px" MaxLength="10" Height="16px" Font-Names="Tahoma"
                                            Font-Size="8pt" ID="ASPxTextBox7" BackColor="#FFD2A6" Text="ETD PORT" ReadOnly="True"
                                            Font-Bold="True">
                                        </dx:ASPxTextBox>
                                    </td>
                                    <td width="100px" height="16px">
                                        <dx:ASPxDateEdit ID="dtWeekETDPortOld1" runat="server" ClientInstanceName="dtWeekETDPortOld1"
                                            DisplayFormatString="yyyy-MM-dd" EditFormat="Custom" EditFormatString="yyyy-MM-dd"
                                            Font-Names="Tahoma" Font-Size="8pt" Width="100px" ReadOnly="True">
                                            <ReadOnlyStyle BackColor="#CCCCCC">
                                            </ReadOnlyStyle>
                                        </dx:ASPxDateEdit>
                                    </td>
                                </tr>                                
                                <tr>
                                    <td width="150px" height="16px">
                                        <dx:ASPxTextBox runat="server" Width="150px" MaxLength="10" Height="16px" Font-Names="Tahoma"
                                            Font-Size="8pt" ID="ASPxTextBox8" BackColor="#FFD2A6" Text="ETA PORT" ReadOnly="True"
                                            Font-Bold="True">
                                        </dx:ASPxTextBox>
                                    </td>
                                    <td width="100px" height="16px">
                                        <dx:ASPxDateEdit ID="dtWeekETAPortOld1" runat="server" ClientInstanceName="dtWeekETAPortOld1"
                                            DisplayFormatString="yyyy-MM-dd" EditFormat="Custom" EditFormatString="yyyy-MM-dd"
                                            Font-Names="Tahoma" Font-Size="8pt" Width="100px" ReadOnly="True">
                                            <ReadOnlyStyle BackColor="#CCCCCC">
                                            </ReadOnlyStyle>
                                        </dx:ASPxDateEdit>
                                    </td>
                                </tr>                                
                                <tr>
                                    <td width="150px" height="16px">
                                        <dx:ASPxTextBox runat="server" Width="150px" MaxLength="10" Height="16px" Font-Names="Tahoma"
                                            Font-Size="8pt" ID="ASPxTextBox9" BackColor="#FFD2A6" Text="ETA FACTORY" ReadOnly="True"
                                            Font-Bold="True">
                                        </dx:ASPxTextBox>
                                    </td>
                                    <td width="100px" height="16px">
                                        <dx:ASPxDateEdit ID="dtETAFactWeekOld1" runat="server" ClientInstanceName="dtETAFactWeekOld1"
                                            DisplayFormatString="yyyy-MM-dd" EditFormat="Custom" EditFormatString="yyyy-MM-dd"
                                            Font-Names="Tahoma" Font-Size="8pt" Width="100px" ReadOnly="True">
                                            <ReadOnlyStyle BackColor="#CCCCCC">
                                            </ReadOnlyStyle>
                                        </dx:ASPxDateEdit>
                                    </td>                                    
                                </tr>                                
                             </table>
                        </td>
                        <td width="150px">
                            &nbsp;
                        </td>
                        <td width="150px">
                            &nbsp;
                        </td>
                    </tr>                    
                </table>
            </td>
        </tr>

        <tr>
            <td colspan="8" align="left" valign="top" height="220">
                <%--Column : Grid--%>
                <dx:ASPxGridView ID="grid" runat="server" AutoGenerateColumns="False" ClientInstanceName="grid"
                    Font-Names="Tahoma" Font-Size="8pt" KeyFieldName="NoUrut" Width="100%">
                    <ClientSideEvents CallbackError="function(s, e) {e.handled = true;}" EndCallback="function(s, e) {                       
                        var pMsg = s.cpMessage;
                        if (pMsg != '') {
                            if (pMsg.substring(1,5) == '1001' || pMsg.substring(1,5) == '1002' || pMsg.substring(1,5) == '1003' || pMsg.substring(1,5) == '2001' || pMsg.substring(1,5) == '1008' || pMsg.substring(1,5) == '1009') {
                                lblInfo.GetMainElement().style.color = 'Blue';
                            } else {
                                lblInfo.GetMainElement().style.color = 'Red';
                            }
                            lblInfo.SetText(pMsg);
                        } else {
                            lblInfo.SetText('');
                        }
                        delete s.cpMessage;
                        delete s.cpSearch;
                    }" FocusedRowChanged="function(s, e) {
	                    OnGridFocusedRowChanged();
                    }" Init="OnInit" RowClick="function(s, e) {
	                    lblInfo.SetText('');
                    }" />
                    <Columns>
                        <dx:GridViewDataTextColumn Caption="NO." FieldName="NoUrut" HeaderStyle-HorizontalAlign="Center"
                            VisibleIndex="1" Width="30px">
                            <HeaderStyle HorizontalAlign="Center"></HeaderStyle>
                            <CellStyle Font-Names="Tahoma" Font-Size="8pt">
                            </CellStyle>
                        </dx:GridViewDataTextColumn>
                        <dx:GridViewDataTextColumn Caption="PART NO." FieldName="PartNo" HeaderStyle-HorizontalAlign="Center"
                            VisibleIndex="2" Width="90px">
                            <HeaderStyle HorizontalAlign="Center"></HeaderStyle>
                            <CellStyle Font-Names="Tahoma" Font-Size="8pt">
                            </CellStyle>
                        </dx:GridViewDataTextColumn>
                        <dx:GridViewDataTextColumn Caption="PART NAME" FieldName="PartName" HeaderStyle-HorizontalAlign="Center"
                            VisibleIndex="3" Width="180px">
                            <HeaderStyle HorizontalAlign="Center"></HeaderStyle>
                            <CellStyle Font-Names="Tahoma" Font-Size="8pt">
                            </CellStyle>
                        </dx:GridViewDataTextColumn>
                        <dx:GridViewDataTextColumn Caption="UOM" CellStyle-HorizontalAlign="Center" FieldName="UnitDesc"
                            HeaderStyle-HorizontalAlign="Center" VisibleIndex="4" Width="40px">
                            <HeaderStyle HorizontalAlign="Center"></HeaderStyle>
                            <CellStyle Font-Names="Tahoma" Font-Size="8pt">
                            </CellStyle>
                        </dx:GridViewDataTextColumn>
                        <dx:GridViewDataTextColumn Caption="MOQ" FieldName="MOQ" HeaderStyle-HorizontalAlign="Center"
                            VisibleIndex="5" Width="50px">
                            <PropertiesTextEdit DisplayFormatString="{0:n0}">
                            </PropertiesTextEdit>
                            <HeaderStyle HorizontalAlign="Center"></HeaderStyle>
                            <CellStyle Font-Names="Tahoma" Font-Size="8pt">
                            </CellStyle>
                        </dx:GridViewDataTextColumn>
                        <dx:GridViewDataTextColumn Caption="QTY/BOX" FieldName="QtyBox" HeaderStyle-HorizontalAlign="Center"
                            VisibleIndex="6" Width="70px">
                            <PropertiesTextEdit DisplayFormatString="{0:n0}">
                            </PropertiesTextEdit>
                            <HeaderStyle HorizontalAlign="Center" Wrap="True"></HeaderStyle>
                            <CellStyle Font-Names="Tahoma" Font-Size="8pt">
                            </CellStyle>
                        </dx:GridViewDataTextColumn>                        
                        <dx:GridViewDataTextColumn FieldName="POQtyOld" Width="80px"
                            Caption="TOTAL FIRM QTY" VisibleIndex="7">
                            <PropertiesTextEdit DisplayFormatString="{0:n0}">
                            </PropertiesTextEdit>
                            <HeaderStyle HorizontalAlign="Center"></HeaderStyle>
                            <CellStyle Font-Names="Tahoma" Font-Size="8pt">
                            </CellStyle>
                        </dx:GridViewDataTextColumn>
                        <dx:GridViewDataTextColumn VisibleIndex="8" Caption="ETD VENDOR" FieldName="ETDVendor1" Width="80px"
                            HeaderStyle-HorizontalAlign="Center">
                            <PropertiesTextEdit DisplayFormatString="yyyy-MM-dd">
                            </PropertiesTextEdit>
                            <HeaderStyle HorizontalAlign="Center"></HeaderStyle>
                            <CellStyle Font-Names="Tahoma" Font-Size="8pt">
                            </CellStyle>
                        </dx:GridViewDataTextColumn>
                        <dx:GridViewDataTextColumn VisibleIndex="9" Caption="ETD VENDOR" FieldName="ETDVendor1Old" Width="0px"
                            HeaderStyle-HorizontalAlign="Center">
                            <PropertiesTextEdit DisplayFormatString="yyyy-MM-dd">
                            </PropertiesTextEdit>
                            <HeaderStyle HorizontalAlign="Center"></HeaderStyle>
                            <CellStyle Font-Names="Tahoma" Font-Size="8pt">
                            </CellStyle>
                        </dx:GridViewDataTextColumn>
                        <dx:GridViewDataTextColumn FieldName="POQty" Width="110px"
                            Caption="SUPPLIER CONFIRMATION QTY" VisibleIndex="10">
                            <PropertiesTextEdit DisplayFormatString="{0:n0}">
                            </PropertiesTextEdit>
                            <HeaderStyle HorizontalAlign="Center"></HeaderStyle>
                            <CellStyle Font-Names="Tahoma" Font-Size="8pt">
                            </CellStyle>
                        </dx:GridViewDataTextColumn>                        
                    </Columns>
                    <SettingsBehavior AllowSelectByRowClick="True" AllowSort="False" ColumnResizeMode="Control"
                        EnableRowHotTrack="True" />
                    <SettingsPager Mode="ShowAllRecords">
                        <Summary AllPagesText="Page {0} of {1} " Text="Page {0} of {1} [{2} record(s)]" />                        
                    </SettingsPager>
                    <SettingsEditing Mode="Batch" NewItemRowPosition="Bottom">
                        <BatchEditSettings ShowConfirmOnLosingChanges="False" />
                    </SettingsEditing>
                    <Settings ShowGroupButtons="False" ShowHorizontalScrollBar="True" ShowStatusBar="Hidden"
                        ShowVerticalScrollBar="True" VerticalScrollableHeight="190" />
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
