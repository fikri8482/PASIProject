<%@ Page Title="" Language="vb" AutoEventWireup="false" MasterPageFile="~/SiteContent.Master"
    CodeBehind="ForwarderMapping.aspx.vb" Inherits="PASISystem.ForwarderMapping" %>

<%@ Register Assembly="DevExpress.Web.v14.1, Version=14.1.6.0, Culture=neutral, PublicKeyToken=b88d1754d700e49a"
    Namespace="DevExpress.Web.ASPxPopupControl" TagPrefix="dx" %>
<%@ Register Assembly="DevExpress.Web.v14.1, Version=14.1.6.0, Culture=neutral, PublicKeyToken=b88d1754d700e49a"
    Namespace="DevExpress.Web.ASPxGridView" TagPrefix="dx" %>
<%@ Register Assembly="DevExpress.Web.v14.1, Version=14.1.6.0, Culture=neutral, PublicKeyToken=b88d1754d700e49a"
    Namespace="DevExpress.Web.ASPxEditors" TagPrefix="dx" %>
<%@ Register Assembly="DevExpress.Web.v14.1, Version=14.1.6.0, Culture=neutral, PublicKeyToken=b88d1754d700e49a"
    Namespace="DevExpress.Web.ASPxTabControl" TagPrefix="dx" %>
<%@ Register Assembly="DevExpress.Web.v14.1, Version=14.1.6.0, Culture=neutral, PublicKeyToken=b88d1754d700e49a"
    Namespace="DevExpress.Web.ASPxClasses" TagPrefix="dx" %>
<%@ Register Assembly="DevExpress.Web.v14.1, Version=14.1.6.0, Culture=neutral, PublicKeyToken=b88d1754d700e49a"
    Namespace="DevExpress.Web.ASPxPanel" TagPrefix="dx" %>
<%@ Register Assembly="DevExpress.Web.v14.1, Version=14.1.6.0, Culture=neutral, PublicKeyToken=b88d1754d700e49a"
    Namespace="DevExpress.Web.ASPxGlobalEvents" TagPrefix="dx" %>
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
            width: 46px;
        }
    </style>
    <script language="javascript" type="text/javascript">
        function OnGridFocusedRowChangedUser() {
            gridUser.GetRowValues(gridUser.GetFocusedRowIndex(), 'AffiliateID', OnGetRowValuesUser);
        }
        function OnGetRowValuesUser(values) {
            cboAffiliateID.SetText(values[0].trim());
            txtUserId.SetText(values[1].trim());
            txtUserIDTemp.SetText(values[1].trim());
            txtCCTemp.SetText(values[1].trim());
            txtFullName.SetText(values[2].trim());
            txtPasswordUS.SetText(values[3].trim());
            txtConfPassword.SetText(values[3].trim());
            cboUserGroup.SetText('');
            cbAccount.SetValue(values[5].trim());
            rblAdminStatus.SetValue(values[6].trim());
            txtDesc.SetValue(values[7].trim());
            lblErrMsg.SetText("");
            ASPxCallback1.PerformCallback();
            gridMenu.PerformCallback('load|' + cboAffiliateID.GetText());
        }

        function OnBatchEditStartEditing(s, e) {
            currentColumnName = e.focusedColumn.fieldName;

            if (currentColumnName == "fwdid") {
                e.cancel = true;
            }

            currentEditableVisibleIndex = e.visibleIndex;
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
            height = height - (height * 70 / 100)
            gridMenu.SetHeight(height);

            var heightUser = Math.max(0, myHeight);
            heightUser = heightUser - (heightUser * 65 / 100)
            gridUser.SetHeight(heightUser);
        }
    </script>
</asp:Content>
<asp:Content ID="Content2" ContentPlaceHolderID="Content" runat="server">
    <table style="border-width: thin; border-style: inset hidden ridge hidden; border-color: #9598A1;
        width: 100%; height: 15px;">
        <tr>
            <td>
                <!-- MESSAGE AREA #C0C0C0 -->
                <table id="tblMsg" style="width: 100%; height: 15px;">
                    <tr>
                        <td align="left" valign="middle" style="height: 15px;">
                            <dx:ASPxLabel ID="lblErrMsg" runat="server" Text="" Font-Names="Tahoma" ClientInstanceName="lblErrMsg"
                                Font-Italic="True" Font-Bold="true" Font-Size="8pt">
                            </dx:ASPxLabel>
                        </td>
                    </tr>
                </table>
            </td>
        </tr>
    </table>
    <div style="height: 5px;">
    </div>
    <table style="width: 100%;">
        <tr>
            <td valign="top" width="50%">
                <table style="width: 100%;" frame="box">
                    <tr>
                        <td align="left" class="style1">
                            <table style="width:100%;">
                                <tr>
                                    <td>
                            <dx:ASPxLabel ID="ASPxLabel1" runat="server" Font-Names="Tahoma" Font-Size="8pt"
                                Text="AFFILIATE ID" Width="70px">
                            </dx:ASPxLabel>
                                    </td>
                                    <td>
                            <dx:ASPxComboBox ID="cboAffiliateID" runat="server" ClientInstanceName="cboAffiliateID"
                                Width="100px" Font-Names="Tahoma" DataSourceID="AffiliateUser" TextField="AffiliateID"
                                ValueField="AffiliateID" TextFormatString="{0}" Font-Size="8pt">
                                <ClientSideEvents SelectedIndexChanged="function(s, e) {
	                                gridMenu.PerformCallback('load|' + cboAffiliateID.GetText());
                                }" />
                                <Columns>
                                    <dx:ListBoxColumn Caption="AFFILIATE ID" FieldName="AffiliateID" Width="70px" />
                                    <dx:ListBoxColumn Caption="AFFILIATE NAME" FieldName="AffiliateName" Width="180px" />
                                </Columns>
                                <LoadingPanelStyle ImageSpacing="5px">
                                </LoadingPanelStyle>
                            </dx:ASPxComboBox>
                                    </td>
                                </tr>
                            </table>
                        </td>
                        <td align="left" width="200px">
                            <asp:SqlDataSource ID="AffiliateUser" runat="server" ConnectionString="<%$ ConnectionStrings:KonString %>"
                                SelectCommand="SELECT RTRIM(AffiliateID)AffiliateID, AffiliateName FROM dbo.MS_Affiliate ORDER BY AffiliateID">
                            </asp:SqlDataSource>
                        </td>
                    </tr>
                </table>
            </td>
        </tr>
        <tr>
            <td valign="top" width="50%">
                &nbsp;
            </td>
        </tr>
        <tr>
            <td align="left" valign="top">
                <dx:ASPxGridView runat="server" ClientInstanceName="gridMenu" KeyFieldName="ForwarderID"
                    AutoGenerateColumns="False" Width="100%" Font-Names="Tahoma" Font-Size="8pt"
                    ID="gridMenu">
                    <ClientSideEvents BatchEditStartEditing="OnBatchEditStartEditing" EndCallback="function(s, e) {
                                        gridMenu.CancelEdit();
                                        var pMsg = s.cpMessage;
                                        if (pMsg != '') {
                                                    if (pMsg.substring(1,5) == '6011') {
                                                        lblErrMsg.GetMainElement().style.color = 'Red';        
                                                    } else {
                                                        lblErrMsg.GetMainElement().style.color = 'Blue';
                                                    }

                                                    lblErrMsg.SetText(s.cpMessage);
                                                } else {
                                                    lblErrMsg.SetText('');
                                                }
	                                            delete s.cpMessage;
                                        
}" CallbackError="function(s, e) {
e.Cancel=True;
}"></ClientSideEvents>
                    <Columns>
                        <dx:GridViewDataTextColumn FieldName="ForwarderID" ReadOnly="True" Width="150px" Caption="FORWARDER ID"
                            VisibleIndex="0" Name="fwdid">
                            <CellStyle Font-Names="Tahoma" Font-Size="8pt">
                            </CellStyle>
                        </dx:GridViewDataTextColumn>
                        <dx:GridViewDataCheckColumn FieldName="AirCls" Width="100px" Caption="AIR" 
                            VisibleIndex="1" Name="air">
                            <PropertiesCheckEdit ValueType="System.Int32" ValueChecked="1" ValueUnchecked="0">
                            </PropertiesCheckEdit>
                            <Settings HeaderFilterMode="CheckedList"></Settings>
                            <HeaderStyle HorizontalAlign="Center"></HeaderStyle>
                        </dx:GridViewDataCheckColumn>
                        <dx:GridViewDataCheckColumn FieldName="BoatCls" Width="100px" Caption="BOAT"
                            VisibleIndex="2">
                            <PropertiesCheckEdit ValueType="System.Int32" ValueChecked="1" ValueUnchecked="0">
                            </PropertiesCheckEdit>
                            <HeaderStyle HorizontalAlign="Center"></HeaderStyle>
                        </dx:GridViewDataCheckColumn>
                    </Columns>
                    <SettingsBehavior AllowSort="False" AllowFocusedRow="True" ColumnResizeMode="Control"
                        EnableRowHotTrack="True"></SettingsBehavior>
                    <SettingsPager Mode="ShowAllRecords">
                        <Summary AllPagesText="Page {0} of {1} " Text="Page {0} of {1} [{2} record(s)]">
                        </Summary>
                    </SettingsPager>
                    <SettingsEditing Mode="Batch" NewItemRowPosition="Bottom">
                        <BatchEditSettings ShowConfirmOnLosingChanges="False"></BatchEditSettings>
                    </SettingsEditing>
                    <Settings ShowVerticalScrollBar="True" VerticalScrollableHeight="500" ShowStatusBar="Hidden"
                        HorizontalScrollBarMode="Visible" VerticalScrollBarMode="Visible"></Settings>
                    <Styles>
                        <Header BackColor="#FFD2A6" Font-Names="Tahoma" Font-Size="8pt">
                        </Header>
                        <Row Wrap="False" BackColor="#FFFFE1" Font-Names="Tahoma" Font-Size="8pt">
                        </Row>
                        <RowHotTrack Wrap="False" BackColor="#E8EFFD" Font-Names="Tahoma" Font-Size="8pt">
                        </RowHotTrack>
                        <SelectedRow Wrap="False">
                        </SelectedRow>
                        <FocusedRow BackColor="#DCE7FC" ForeColor="Black">
                        </FocusedRow>
                    </Styles>
                    <StylesEditors ButtonEditCellSpacing="0">
                        <ProgressBar Height="21px">
                        </ProgressBar>
                    </StylesEditors>
                </dx:ASPxGridView>
            </td>
        </tr>
    </table>
    <div style="height: 2px;">
    </div>
    <table style="width: 100%;">
        <tr>
            <td align="left">
                <dx:ASPxButton ID="btnSubMenu" runat="server" ClientInstanceName="btnSubMenu" Font-Names="Tahoma"
                    Text="SUB MENU" Width="90px" Font-Size="8pt">
                </dx:ASPxButton>
            </td>
            <td align="left">
                <dx:ASPxTextBox ID="txtCCTemp" runat="server" ClientInstanceName="txtCCTemp" Width="170px"
                    BackColor="White" ForeColor="White">
                    <Border BorderColor="White" />
                </dx:ASPxTextBox>
            </td>
            <td align="left">
                <dx:ASPxTextBox ID="txtUserIDTemp" runat="server" ClientInstanceName="txtUserIDTemp"
                    Width="170px" BackColor="White" ForeColor="White">
                    <Border BorderColor="White" />
                </dx:ASPxTextBox>
            </td>
            <td align="right" width="85px">
            </td>
            <td align="right" width="85px">
                &nbsp;
            </td>
            <td align="center" width="85px">
                &nbsp;
            </td>
            <td align="right" width="85px">
                <dx:ASPxButton ID="btnSubmit" runat="server" Text="SAVE" Width="80px" Font-Names="Tahoma"
                    AutoPostBack="false" Font-Size="8pt" ClientInstanceName="btnSubmit">
                    <ClientSideEvents Click="function(s, e) {
                        if (cboAffiliateID.GetText() == '') {
                            lblErrMsg.GetMainElement().style.color = 'Red';
		                    lblErrMsg.SetText('[6010] Please choose Affiliate ID first!');
                            e.processOnServer = false;
                            return;
                        }

                        gridMenu.UpdateEdit();
                        gridMenu.PerformCallback('load' + '|' + cboAffiliateID.GetText());
	                    }" />
                </dx:ASPxButton>
            </td>
        </tr>
    </table>
    <div style="height: 8px;">
    </div>
    <table style="width: 100%;">
        <tr>
            <td>
                &nbsp;
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
