<%@ Page Title="" Language="vb" AutoEventWireup="false" MasterPageFile="~/SiteContent.Master" CodeBehind="RemainingPOExport.aspx.vb" Inherits="PASISystem.RemainingPOExport1" %>
<%@ Register Assembly="DevExpress.Web.v14.1, Version=14.1.6.0, Culture=neutral, PublicKeyToken=b88d1754d700e49a"
    Namespace="DevExpress.Web.ASPxEditors" TagPrefix="dx" %>
<%@ Register Assembly="DevExpress.Web.v14.1, Version=14.1.6.0, Culture=neutral, PublicKeyToken=b88d1754d700e49a"
    Namespace="DevExpress.Web.ASPxGridView" TagPrefix="dx" %>
<%@ Register Assembly="DevExpress.Web.v14.1, Version=14.1.6.0, Culture=neutral, PublicKeyToken=b88d1754d700e49a"
    Namespace="DevExpress.Web.ASPxGridView.Export" TagPrefix="dx" %>
<%@ Register Assembly="DevExpress.Web.v14.1, Version=14.1.6.0, Culture=neutral, PublicKeyToken=b88d1754d700e49a"
    Namespace="DevExpress.Web.ASPxGlobalEvents" TagPrefix="dx" %>
<%@ Register Assembly="DevExpress.Web.v14.1, Version=14.1.6.0, Culture=neutral, PublicKeyToken=b88d1754d700e49a"
    Namespace="DevExpress.Web.ASPxHiddenField" TagPrefix="dx" %>
<%@ Register assembly="DevExpress.Web.v14.1, Version=14.1.6.0, Culture=neutral, PublicKeyToken=b88d1754d700e49a" 
    Namespace="DevExpress.Web.ASPxRoundPanel" tagprefix="dx" %>
<%@ Register assembly="DevExpress.Web.v14.1, Version=14.1.6.0, Culture=neutral, PublicKeyToken=b88d1754d700e49a"
    Namespace="DevExpress.Web.ASPxPanel" tagprefix="dx" %>
<asp:Content ID="Content1" ContentPlaceHolderID="head" runat="server">
    <style type="text/css">
        .dxeBase
        {
            font: 12px Tahoma, Geneva, sans-serif;
        }
        
        .dxeBase
        {
            font: 12px Tahoma, Geneva, sans-serif;
        }
        
        .style3
        {
            height: 200px;
        }
        .style2
        {
            height: 25px;
        }
    </style>
</asp:Content>
<asp:Content ID="Content2" ContentPlaceHolderID="Content" runat="server">

    <script type="text/javascript">

        function OnAllCheckedChanged(s, e) {
            if (s.GetValue() == -1) s.SetValue(1);
            for (var i = 0; i < grid.GetVisibleRowsOnPage(); i++) {
                grid.batchEditApi.SetCellValue(i, "Act", s.GetValue());
            }
        }

        function OnUpdateClick(s, e) {
            Grid.PerformCallback("Update");
        }

        function OnCancelClick(s, e) {
            Grid.PerformCallback("Cancel");
        }

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
            height = height - (height * 54 / 100)
            grid.SetHeight(height);
        }     
    </script>
    <table style="width:100%; height: 150px;">
    <tr>
    <td align="left" style="width: 100%;">
        <dx:ASPxRoundPanel ID="ASPxRoundPanel1" runat="server" HeaderText="" 
            ShowCollapseButton="true" View="GroupBox" Width="100%" Height="150" 
            ShowHeader="False" BackColor="White">
            <ContentPaddings PaddingLeft="5px" PaddingRight="5px" />
            <PanelCollection>
                <dx:PanelContent ID="PanelContent1" runat="server">
                <table id="Table2">
                    <tr>
                        <td align="left" height="20px" valign="middle" width="150px">
                            <dx:ASPxLabel ID="ASPxLabel13" runat="server" 
                                Text="PERIOD" Width="150px">
                            </dx:ASPxLabel>
                        </td>
                        <td align="left" height="20px" valign="middle" width="110px">
                            <dx:ASPxTimeEdit ID="dtPeriodFrom1" runat="server" ClientInstanceName="dtPeriodFrom1" 
                                DisplayFormatString="MMM yyyy" EditFormat="Custom" 
                                EditFormatString="MMM yyyy" Font-Names="Verdana" Font-Size="8pt" Width="110px">
                                <ClientSideEvents ValueChanged="function(s, e) {
	                                cboRequestNo.PerformCallback(cboPersoninCharge.GetText().toString() + '|' + cboFactory.GetText().toString() + '|' + dtPeriod.GetValue().toString() + '|' + cboDepartment.GetText().toString() + '|' + dtRequestFrom.GetValue().toString() + '|' + dtRequestTo.GetValue().toString());
                                }" />
                            </dx:ASPxTimeEdit>
                        </td>
                        <td>~</td>
                        <td align="left" style="width: 160px">
                            <dx:ASPxTimeEdit ID="dtPeriodTo1" runat="server" ClientInstanceName="dtPeriodTo1" 
                                DisplayFormatString="MMM yyyy" EditFormat="Custom" 
                                EditFormatString="MMM yyyy" Font-Names="Verdana" Font-Size="8pt" Width="110px">
                                <ClientSideEvents ValueChanged="function(s, e) {
	                            cboRequestNo.PerformCallback(cboPersoninCharge.GetText().toString() + '|' + cboFactory.GetText().toString() + '|' + dtPeriod.GetValue().toString() + '|' + cboDepartment.GetText().toString() + '|' + dtRequestFrom.GetValue().toString() + '|' + dtRequestTo.GetValue().toString());
                            }" />
                            </dx:ASPxTimeEdit>
                        </td>
                        <td align="left" height="20px" valign="middle" width="250"> </td>
                        <td align="left" height="20px" valign="middle"></td>
                        <td align="left" height="20px" valign="middle" width="170">
                            <dx:ASPxLabel ID="ASPxLabel22" runat="server" Text="PO MONTHLY / EMERGENCY" 
                                Width="170px">
                            </dx:ASPxLabel>
                        </td>
                        <td align="left" height="20px" valign="middle">
                            <dx:ASPxRadioButton ID="rdAll" runat="server" Text="ALL" 
                                ClientInstanceName="rdAll" Checked="True" GroupName="EmergencyCls">
                                <ClientSideEvents CheckedChanged="function(s, e) {
	lblErrMsg.SetText('');
}" />
                            </dx:ASPxRadioButton>
                        </td>
                        <td align="left" height="20px" valign="middle">
                            <dx:ASPxRadioButton ID="rbMonthly" runat="server" Text="MONTHLY" 
                                ClientInstanceName="rbMonthly" GroupName="EmergencyCls">
                                <ClientSideEvents CheckedChanged="function(s, e) {
	lblErrMsg.SetText('');
}" />
                            </dx:ASPxRadioButton>
                        </td>
                        <td align="left" height="20px" valign="middle">
                            <dx:ASPxRadioButton ID="rbEmergency" runat="server" Text="EMERGENCY" 
                                ClientInstanceName="rbEmergency" GroupName="EmergencyCls">
                                <ClientSideEvents CheckedChanged="function(s, e) {
	lblErrMsg.SetText('');
}" />
                            </dx:ASPxRadioButton>
                        </td>
                    </tr>
                    <tr>
                        <td align="left" height="20px" valign="middle" width="150px">
                            <dx:ASPxLabel ID="ASPxLabel14" runat="server" Text="AFFILIATE CODE/NAME">
                            </dx:ASPxLabel>
                        </td>
                        <td align="left" height="20px" valign="middle" width="110">
                            <dx:ASPxComboBox ID="cboAffiliateCode" runat="server" Width="110px" 
                                ClientInstanceName="cboAffiliateCode" TextFormatString="{0}">
                                <ClientSideEvents SelectedIndexChanged="function(s, e) {
	txtAffiliateName.SetText(cboAffiliateCode.GetSelectedItem().GetColumnText(1));
	lblErrMsg.SetText('');
}" />
                            </dx:ASPxComboBox>
                        </td>
                        <td align="left" height="20px" valign="middle" width="250" colspan="3">
                            <dx:ASPxTextBox ID="txtAffiliateName" runat="server" BackColor="#CCCCCC" 
                                ClientInstanceName="txtAffiliateName" ReadOnly="True" Width="250px">
                            </dx:ASPxTextBox>
                        </td>
                        <td align="left" height="20px" valign="middle"></td>
                        <td align="left" height="20px" valign="middle" width="170">
                            <dx:ASPxLabel ID="ASPxLabel23" runat="server" Text="ORDER NO" Width="170px">
                            </dx:ASPxLabel>
                        </td>
                        <td align="left" height="20px" valign="middle" colspan="3">
                            <dx:ASPxTextBox ID="txtOrderNo" runat="server" ClientInstanceName="txtOrderNo" 
                                Width="130px">
                            </dx:ASPxTextBox>
                        </td>
                    </tr>
                    <tr>
                        <td>
                            <dx:ASPxLabel ID="ASPxLabel19" runat="server" Text="SUPPLIER CODE/NAME" 
                                Width="150px">
                            </dx:ASPxLabel>
                        </td>
                        <td align="left" height="20px" valign="middle" width="110px">
                            <dx:ASPxComboBox ID="cboSupplierCode" runat="server" 
                                ClientInstanceName="cboSupplierCode" TextFormatString="{0}" Width="110px">
                                <ClientSideEvents SelectedIndexChanged="function(s, e) {
	txtSupplierName.SetText(cboSupplierCode.GetSelectedItem().GetColumnText(1));
	lblErrMsg.SetText('');
}" />
                            </dx:ASPxComboBox>
                        </td>
                        <td align="left" height="20px" valign="middle" width="250" colspan="3">
                            <dx:ASPxTextBox ID="txtSupplierName" runat="server" BackColor="#CCCCCC" 
                                ClientInstanceName="txtSupplierName" ReadOnly="True" Width="250px">
                            </dx:ASPxTextBox>
                        </td>
                        <td></td>
                        <td>
                            &nbsp;</td>
                        <td>
                            &nbsp;</td>
                        <td>
                            &nbsp;</td>
                        <td>
                            &nbsp;</td>
                    </tr>
                    <tr>
                        <td>
                            <dx:ASPxLabel ID="ASPxLabel15" runat="server" Text="FORWARDER" width="150px">
                            </dx:ASPxLabel>
                        </td>
                        <td align="left" height="20px" valign="middle" width="110px">
                            <dx:ASPxComboBox ID="cboForwarder" runat="server" 
                                ClientInstanceName="cboForwarder" TextFormatString="{0}" Width="110px">
                                <ClientSideEvents SelectedIndexChanged="function(s, e) {
	txtForwarder.SetText(cboForwarder.GetSelectedItem().GetColumnText(1));
	lblErrMsg.SetText('');
}" />
                            </dx:ASPxComboBox>
                        </td>
                        <td align="left" height="20px" valign="middle" width="250" colspan="3">
                            <dx:ASPxTextBox ID="txtForwarder" runat="server" BackColor="#CCCCCC" 
                                ClientInstanceName="txtForwarder" ReadOnly="True" Width="250px">
                            </dx:ASPxTextBox>
                        </td>
                        <td></td>
                        <td>
                            &nbsp;</td>
                        <td>
                            &nbsp;</td>
                        <td>
                            &nbsp;</td>
                        <td>
                            &nbsp;</td>
                    </tr>
                    <tr>
                        <td>
                            <dx:ASPxLabel ID="ASPxLabel20" runat="server" Text="PART CODE/NAME" width="150px">
                            </dx:ASPxLabel>
                        </td>
                        <td align="left" height="20px" valign="middle" width="110px">
                            <dx:ASPxComboBox ID="cboPartNo" runat="server" ClientInstanceName="cboPartNo" 
                                TextFormatString="{0}" Width="110px">
                                <ClientSideEvents SelectedIndexChanged="function(s, e) {
	txtPartName.SetText(cboPartNo.GetSelectedItem().GetColumnText(1));
	lblErrMsg.SetText('');
}" />
                            </dx:ASPxComboBox>
                        </td>
                        <td align="left" height="20px" valign="middle" width="250" colspan="3">
                            <dx:ASPxTextBox ID="txtPartName" runat="server" BackColor="#CCCCCC" 
                                ClientInstanceName="txtPartName" ReadOnly="True" Width="250px">
                            </dx:ASPxTextBox>
                        </td>
                        <td></td>
                        <td></td>
                        <td></td>
                        <td style="margin-left: 40px">
                            <dx:ASPxButton ID="btnSearch" runat="server" AutoPostBack="False" 
                                ClientInstanceName="btnSearch" Text="SEARCH" Width="85px">
                                <ClientSideEvents Click="function(s, e) {
	grid.PerformCallback('loaddata');
	lblErrMsg.SetText('');
}" />
                            </dx:ASPxButton>
                        </td>
                        <td style="margin-left: 40px" align="right">
                            <dx:ASPxButton ID="btnClearAll0" runat="server" AutoPostBack="False" 
                                ClientInstanceName="btnClearAll" Text="CLEAR" Width="85px">
                                <ClientSideEvents Click="function(s, e) {
dtPeriodFrom1.SetDate(new Date());
dtPeriodTo1.SetDate(new Date());
cboAffiliateCode.SetText('==ALL==');
txtAffiliateName.SetText('==ALL==');
cboSupplierCode.SetText('==ALL==');
txtSupplierName.SetText('==ALL==');
cboForwarder.SetText('==ALL==');
txtForwarder.SetText('==ALL==');
cboPartNo.SetText('==ALL==');
txtPartName.SetText('==ALL==');
rdAll.SetChecked(true);
rbMonthly.SetChecked(false);
rbEmergency.SetChecked(false);
txtOrderNo.SetText('');
lblErrMsg.SetText('');
}" />
                            </dx:ASPxButton>
                        </td>
                    </tr>
                </table>
                
                </dx:PanelContent>
            </PanelCollection>
        </dx:ASPxRoundPanel>
    </td>
    </tr>
    </table>
    
    <div style="height:8px;">
    </div>
    <table align="left" width="100%">
        <tr align="left">
            <td width="100%" height="16px" style="border-top-style: solid; border-top-width: thin;
                border-top-color: #808080; border-bottom-style: solid; border-bottom-width: thin;
                border-bottom-color: #808080" align="left">
                <dx:ASPxLabel ID="lblErrMsg" runat="server" Font-Names="Tahoma" Font-Size="8pt"
                    Text="ERROR MESSAGE" Width="100%" ClientInstanceName="lblErrMsg" 
                    Height="16px">
                </dx:ASPxLabel>
            </td>
        </tr>
    </table>
    <br />
    <br />
    <table style="width: 100%;">
        <tr>
            <td align="left" class="style3" colspan="5">
                <dx:ASPxGridView ID="grid" runat="server" AutoGenerateColumns="False" 
                    Width="100%" KeyFieldName="NoUrut;AffiliateID;SupplierID;ForwarderID;OrderNo;PartNo" ClientInstanceName="grid">
                    <ClientSideEvents EndCallback="function(s, e) {
						var pMsg = s.cpMessage;
                        if (pMsg != '') {
                            if (pMsg.substring(1,5) == '1001' || pMsg.substring(1,5) == '1002' || pMsg.substring(1,5) == '1003' || pMsg.substring(1,5) == '2001') {
                                lblErrMsg.GetMainElement().style.color = 'Blue';
                            } else {
                                lblErrMsg.GetMainElement().style.color = 'Red';
                            }
                            lblErrMsg.SetText(pMsg);
                        } else {
                            lblErrMsg.SetText('');
                        }

var pSending = s.cpSending;
if (pSending == 'ALREADY SEND') {
alert('yes');
	txtSending.SetText('AREADY SEND');
    alert('oke');
}
                        delete s.cpMessage;
}" RowClick="function(s, e) {
	lblErrMsg.SetText('');
}" CallbackError="function(s, e) {
	e.handled = true;
}" Init="OnInit" />

<ClientSideEvents RowClick="function(s, e) {
	lblErrMsg.SetText(&#39;&#39;);
}" EndCallback="function(s, e) {
						var pMsg = s.cpMessage;
                        if (pMsg != '') {
                            if (pMsg.substring(1,5) == '1001' || pMsg.substring(1,5) == '1002' || pMsg.substring(1,5) == '1003' || pMsg.substring(1,5) == '2001') {
                                lblErrMsg.GetMainElement().style.color = 'Blue';
                            } else {
                                lblErrMsg.GetMainElement().style.color = 'Red';
                            }
                            lblErrMsg.SetText(pMsg);
                        } else {
                            lblErrMsg.SetText('');
                        }

var pSending = s.cpSending;
if (pSending == 'ALREADY SEND') {
	txtSend.SetText('ALREADY SEND');
    } else {
    txtSend.SetText('');
}
                        delete s.cpMessage;
}" CallbackError="function(s, e) {
	e.handled = true;
}" Init="OnInit"></ClientSideEvents>

                    <Columns>
                        <dx:GridViewDataTextColumn Caption="NO." 
                            FieldName="NoUrut" Name="NoUrut" VisibleIndex="0" 
                            Width="40px">
                            <HeaderStyle Font-Names="Tahoma" Font-Size="8pt" Font-Underline="False" 
                                Wrap="True" HorizontalAlign="Center"/>
                            <CellStyle Font-Names="Tahoma" Font-Size="8pt" HorizontalAlign="Center">
                            </CellStyle>
                        </dx:GridViewDataTextColumn>
                        <dx:GridViewDataDateColumn Caption="PERIOD" FieldName="Period" 
                            Name="Period" VisibleIndex="1">
                            <PropertiesDateEdit DisplayFormatString="MMM yyyy" Spacing="0" 
                                EditFormat="Custom" EditFormatString="MMM yyyy">
                            </PropertiesDateEdit>
                            <HeaderStyle Font-Names="Tahoma" Font-Size="8pt" Font-Underline="False" 
                                Wrap="True" HorizontalAlign="Center"/>
                            <CellStyle Font-Names="Verdana" Font-Size="8pt">
                            </CellStyle>
                        </dx:GridViewDataDateColumn>
                        <dx:GridViewDataTextColumn Caption="AFFILIATE" 
                            FieldName="AffiliateID" Name="AffiliateID" VisibleIndex="2" 
                            Width="120px">
                            <HeaderStyle Font-Names="Tahoma" Font-Size="8pt" Font-Underline="False" 
                                Wrap="True" HorizontalAlign="Center"/>
                            <CellStyle Font-Names="Tahoma" Font-Size="8pt">
                            </CellStyle>
                        </dx:GridViewDataTextColumn>
                        <dx:GridViewDataTextColumn Caption="FORWARDER" 
                            FieldName="ForwarderID" Name="ForwarderID" VisibleIndex="3" 
                            Width="120px">
                            <HeaderStyle Font-Names="Tahoma" Font-Size="8pt" Font-Underline="False" 
                                Wrap="True" HorizontalAlign="Center"/>
                            <CellStyle Font-Names="Tahoma" Font-Size="8pt">
                            </CellStyle>
                        </dx:GridViewDataTextColumn>
                        <dx:GridViewDataTextColumn Caption="ORDER NO." 
                            FieldName="OrderNo" Name="OrderNo" VisibleIndex="4" 
                            Width="120px">
                            <HeaderStyle Font-Names="Tahoma" Font-Size="8pt" Font-Underline="False" 
                                Wrap="True" HorizontalAlign="Center"/>
                            <CellStyle Font-Names="Tahoma" Font-Size="8pt">
                            </CellStyle>
                        </dx:GridViewDataTextColumn>
                        <dx:GridViewDataTextColumn VisibleIndex="5" Caption="PO EMERGENCY (EMERGENCY/   MONTHLY)" 
                            FieldName="EmergencyCls" Name="EmergencyCls" Width="120px"
                            HeaderStyle-HorizontalAlign="Center">
                            <HeaderStyle HorizontalAlign="Center" Wrap="True"></HeaderStyle>
                            <CellStyle Font-Names="Tahoma" Font-Size="8pt" HorizontalAlign="Left">
                            </CellStyle>
                        </dx:GridViewDataTextColumn>
                        <dx:GridViewDataTextColumn Caption="SUPPLIER" FieldName="SupplierID" 
                            Name="SupplierID" VisibleIndex="6" Width="70px">
                            <HeaderStyle Font-Names="Tahoma" Font-Size="8pt" Font-Underline="False" 
                                Wrap="True" HorizontalAlign="Center"/>
                            <CellStyle Font-Names="Tahoma" Font-Size="8pt" HorizontalAlign = "Left">
                            </CellStyle>
                        </dx:GridViewDataTextColumn>
                        <dx:GridViewDataTextColumn Caption="ETD VENDOR" FieldName="ETDVendor" 
                            Name="ETDVendor" VisibleIndex="7" Width="70px">
                            <HeaderStyle Font-Names="Tahoma" Font-Size="8pt" Font-Underline="False" 
                                Wrap="True" HorizontalAlign="Center"/>
                            <CellStyle Font-Names="Tahoma" Font-Size="8pt" HorizontalAlign="Center">
                            </CellStyle>
                        </dx:GridViewDataTextColumn>
                        <dx:GridViewDataTextColumn Caption="ETD PORT" FieldName="ETDPort" 
                            Name="ETDPort" VisibleIndex="8" Width="70px">
                            <HeaderStyle Font-Names="Tahoma" Font-Size="8pt" Font-Underline="False" 
                                Wrap="True" HorizontalAlign="Center"/>
                            <CellStyle Font-Names="Tahoma" Font-Size="8pt" HorizontalAlign="Center">
                            </CellStyle>
                        </dx:GridViewDataTextColumn>
                        <dx:GridViewDataTextColumn Caption="PART NO." FieldName="PartNo" 
                            Name="PartNo" VisibleIndex="9" Width="120px">
                            <HeaderStyle Font-Names="Tahoma" Font-Size="8pt" Font-Underline="False" 
                                Wrap="True" HorizontalAlign="Center"/>
                            <CellStyle Font-Names="Tahoma" Font-Size="8pt">
                            </CellStyle>
                        </dx:GridViewDataTextColumn>
                        <dx:GridViewDataTextColumn Caption="ORDER QTY" FieldName="OrderQty" 
                            Name="OrderQty" VisibleIndex="10" Width="80px">
                            <HeaderStyle Font-Names="Tahoma" Font-Size="8pt" Font-Underline="False" 
                                Wrap="True" HorizontalAlign="Center"/>
                            <CellStyle Font-Names="Tahoma" Font-Size="8pt" HorizontalAlign="Right">
                            </CellStyle>
                        </dx:GridViewDataTextColumn>
                        <dx:GridViewDataTextColumn Caption="SUPPLIER DELIVERY QTY" FieldName="SuppDelQty" 
                            Name="SuppDelQty" VisibleIndex="11" Width="70px">
                            <HeaderStyle Font-Names="Tahoma" Font-Size="8pt" Font-Underline="False" 
                                Wrap="True" HorizontalAlign="Center"/>
                            <CellStyle Font-Names="Tahoma" Font-Size="8pt" HorizontalAlign="Right">
                            </CellStyle>
                        </dx:GridViewDataTextColumn>
                        <dx:GridViewDataTextColumn Caption="FWD GOOD RECEIVING QTY" FieldName="GoodRecQty" 
                            Name="GoodRecQty" VisibleIndex="12" Width="70px">
                            <HeaderStyle Font-Names="Tahoma" Font-Size="8pt" Font-Underline="False" 
                                Wrap="True" HorizontalAlign="Center"/>
                            <CellStyle Font-Names="Tahoma" Font-Size="8pt" HorizontalAlign="Right">
                            </CellStyle>
                        </dx:GridViewDataTextColumn>
                        <dx:GridViewDataTextColumn Caption="FWD DEFECT RECEIVING QTY" FieldName="DefRecQty" 
                            Name="DefRecQty" VisibleIndex="13" Width="70px">
                            <HeaderStyle Font-Names="Tahoma" Font-Size="8pt" Font-Underline="False" 
                                Wrap="True" HorizontalAlign="Center"/>
                            <CellStyle Font-Names="Tahoma" Font-Size="8pt" HorizontalAlign="Right">
                            </CellStyle>
                        </dx:GridViewDataTextColumn>
                        <dx:GridViewDataTextColumn Caption="FWD REMAINING RECEIVING QTY" FieldName="RemainRecQty" 
                            Name="RemainRecQty" VisibleIndex="14" Width="70px">
                            <HeaderStyle Font-Names="Tahoma" Font-Size="8pt" Font-Underline="False" 
                                Wrap="True" HorizontalAlign="Center"/>
                            <CellStyle Font-Names="Tahoma" Font-Size="8pt" HorizontalAlign="Right">
                            </CellStyle>
                        </dx:GridViewDataTextColumn>
                    </Columns>
                    <SettingsPager Visible="False" Position="Top">
                        <Summary Text="Page {0} of {1} [{2} record(s)]" AllPagesText="Page {0} of {1} " />
                        <Summary AllPagesText="Page {0} of {1} " Text="Page {0} of {1} [{2} record(s)]">
                        </Summary>
                    </SettingsPager>
                    <Settings ShowHorizontalScrollBar="True" ShowVerticalScrollBar="True" ShowGroupButtons="False"
                        ShowStatusBar="Hidden" VerticalScrollableHeight="250" />

                    <Settings ShowGroupButtons="False" ShowVerticalScrollBar="True" ShowHorizontalScrollBar="True" 
                        VerticalScrollableHeight="250" ShowStatusBar="Hidden"></Settings>
                    <Styles>
                                <Header BackColor="#FFD2A6" Font-Names="Verdana" Font-Size="8pt" Wrap="True"></Header>
                                <Row BackColor="#FFFFE1" Font-Names="Verdana" Font-Size="8pt" Wrap="True"></Row>
                                <RowHotTrack BackColor="#E8EFFD" Font-Names="Verdana" Font-Size="8pt" Wrap="False"></RowHotTrack>
                                <SelectedRow Wrap="False">
                                </SelectedRow>
                            </Styles>
                    <StylesEditors ButtonEditCellSpacing="0">
                        <ProgressBar Height="21px">
                        </ProgressBar>
                    </StylesEditors>
                </dx:ASPxGridView>
                <dx:ASPxGridViewExporter ID="gridExport" runat="server" GridViewID="gridExport" >
                </dx:ASPxGridViewExporter>
                <br />
            </td>
        </tr>
        <tr>
            <td align="left">
                <dx:ASPxButton ID="btnsubmenu" runat="server" Text="SUB MENU" Width="90px" Font-Names="Tahoma"
                    Font-Size="8pt">
                </dx:ASPxButton>
            </td>
            <td align="right" width="50px">
                <dx:ASPxButton ID="btnSendShipping" runat="server" 
                    Text="SEND SHIPPING INSTRUCTION" Width="90px" 
                    ClientInstanceName="btnSendShipping" AutoPostBack="False" Visible="False">
                    <ClientSideEvents Click="function(s, e) {
	grid.PerformCallback('sendshipping');
}" />
                </dx:ASPxButton>
            </td>
            <td align="right" width="50px">
                <dx:ASPxButton ID="btnExcel" runat="server" Text="EXCEL" Width="90px" 
                    ClientInstanceName="btnExcel" AutoPostBack="False" Visible="False">
                    <ClientSideEvents Click="function(s, e) {
	grid.PerformCallback('excelshipping');
}" />
                </dx:ASPxButton>
            </td>
            <td align="right" width="50px">
                <dx:ASPxButton ID="btnDelete" runat="server" Text="DELETE" Width="90px" 
                    ClientInstanceName="btnDelete" AutoPostBack="False" Visible="False">
                    <ClientSideEvents Click="function(s, e) {
	grid.UpdateEdit();
	grid.PerformCallback('loadaftersubmit');
}" />
                </dx:ASPxButton>
            </td>
            <td align="right" width="50px">
                <dx:ASPxButton ID="btnSave" runat="server" Text="SAVE TO EXCEL" Width="90px" 
                    ClientInstanceName="btnSave" AutoPostBack="False">
                    <ClientSideEvents Click="function(s, e) {
	grid.PerformCallback('excelremaining');
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
    <dx:ASPxHiddenField ID="HF" runat="server" ClientInstanceName="HF">
    </dx:ASPxHiddenField> 


</asp:Content>
