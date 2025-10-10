<%@ Page Title="" Language="vb" AutoEventWireup="false" MasterPageFile="~/SiteContent.Master" CodeBehind="UserSetup.aspx.vb" Inherits="AffiliateSystem.UserSetup" %>
<%@ Register Assembly="DevExpress.Web.v14.1, Version=14.1.6.0, Culture=neutral, PublicKeyToken=b88d1754d700e49a" Namespace="DevExpress.Web.ASPxPopupControl" TagPrefix="dx" %>
<%@ Register assembly="DevExpress.Web.v14.1, Version=14.1.6.0, Culture=neutral, PublicKeyToken=b88d1754d700e49a" namespace="DevExpress.Web.ASPxGridView" tagprefix="dx" %>
<%@ Register assembly="DevExpress.Web.v14.1, Version=14.1.6.0, Culture=neutral, PublicKeyToken=b88d1754d700e49a" namespace="DevExpress.Web.ASPxEditors" tagprefix="dx" %>
<%@ Register assembly="DevExpress.Web.v14.1, Version=14.1.6.0, Culture=neutral, PublicKeyToken=b88d1754d700e49a" namespace="DevExpress.Web.ASPxTabControl" tagprefix="dx" %>
<%@ Register assembly="DevExpress.Web.v14.1, Version=14.1.6.0, Culture=neutral, PublicKeyToken=b88d1754d700e49a" namespace="DevExpress.Web.ASPxClasses" tagprefix="dx" %>
<%@ Register assembly="DevExpress.Web.v14.1, Version=14.1.6.0, Culture=neutral, PublicKeyToken=b88d1754d700e49a" namespace="DevExpress.Web.ASPxPanel" tagprefix="dx" %>
<%@ Register Assembly="DevExpress.Web.v14.1, Version=14.1.6.0, Culture=neutral, PublicKeyToken=b88d1754d700e49a" namespace="DevExpress.Web.ASPxGlobalEvents" TagPrefix="dx" %>
<%@ Register assembly="DevExpress.Web.v14.1, Version=14.1.6.0, Culture=neutral, PublicKeyToken=b88d1754d700e49a" namespace="DevExpress.Web.ASPxCallback" tagprefix="dx" %>

<asp:Content ID="Content1" ContentPlaceHolderID="head" runat="server">
    <style type="text/css">
    .dxeHLC, .dxeHC, .dxeHFC
    {
    	display : none;
    }
    </style>
    
    <script language="javascript" type="text/javascript">
    function OnGridFocusedRowChangedUser() {
        gridUser.GetRowValues(gridUser.GetFocusedRowIndex(), 'AffiliateID;UserID;FullName;Password;InvalidLogin;Locked;StatusAdmin;Description', OnGetRowValuesUser);
    }
    function OnGetRowValuesUser(values) {
        cboAffiliateID.SetText(values[0].trim());
        txtUserId.SetText(values[1].trim());
        txtUserIDTemp.SetText(values[1].trim());
        txtFullName.SetText(values[2].trim());
        txtPasswordUS.SetText(values[3].trim());
        txtConfPassword.SetText(values[3].trim());
        cboUserGroup.SetText('');
        cbAccount.SetValue(values[5].trim());
        rblAdminStatus.SetValue(values[6].trim());
        txtDesc.SetValue(values[7].trim());
        lblErrMsg.SetText("");
        ASPxCallback1.PerformCallback();
        gridMenu.PerformCallback('load');
    }
    
    function OnUpdateClick(s, e) {
        gridMenu.PerformCallback("Update");
    }

    function OnCancelClick(s, e) {
        gridMenu.PerformCallback("Cancel");
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
    <table style="border-width: thin; border-style: inset hidden ridge hidden; border-color:#9598A1; width:100%; height: 15px;">
        <tr>
            <td>
                <!-- MESSAGE AREA #C0C0C0 -->
                <table id="tblMsg" style="width:100%; height: 15px;">
                    <tr>
                        <td align="left" valign="middle" style="height:15px;">
                            <dx:ASPxLabel ID="lblErrMsg" runat="server" Text="" Font-Names="Verdana" 
                                ClientInstanceName="lblErrMsg" Font-Italic="True" Font-Bold="true" Font-Size="8pt">
                            </dx:ASPxLabel>                            
                        </td>
                    </tr>
                </table>
            </td>
        </tr>
    </table>

    <div style="height:5px;"></div>

    <table style="width:100%;">
        <tr>            
            <td valign="top" width="35%">
                <dx:ASPxGridView ID="gridUser" runat="server" Width="100%" 
                    Font-Names="Verdana" AutoGenerateColumns="False" 
                    ClientIDMode="Predictable" KeyFieldName="AffiliateID;UserID" 
                    ClientInstanceName="gridUser" Font-Size="8pt">
                    <ClientSideEvents 
                        Init="OnInit"
                        FocusedRowChanged="function(s, e) {OnGridFocusedRowChangedUser();}" 
                        EndCallback="function(s, e) { 
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
                        }" />
                    <Columns>                        
                        <dx:GridViewDataTextColumn Caption="Affiliate ID" FieldName="AffiliateID" Name="AffiliateID" 
                            Visible="false" Width="0px" VisibleIndex="0">
                        </dx:GridViewDataTextColumn>
                        <dx:GridViewDataTextColumn Caption="User ID" 
                            FieldName="UserID" VisibleIndex="1" Name="UserID" ReadOnly="True" 
                            Width="100px">
                            <CellStyle Font-Names="Verdana" Font-Size="8pt">
                            </CellStyle>
                        </dx:GridViewDataTextColumn>
                        <dx:GridViewDataTextColumn Caption="User Name" FieldName="FullName" Name="FullName" 
                            VisibleIndex="2" Width="100px">
                            <CellStyle Font-Names="Verdana" Font-Size="8pt" HorizontalAlign="Left">
                            </CellStyle>
                        </dx:GridViewDataTextColumn>
                        <dx:GridViewDataTextColumn Caption="Password" FieldName="Password" 
                            Name="Password" Visible="False" VisibleIndex="3">
                        </dx:GridViewDataTextColumn>
                        <dx:GridViewDataTextColumn Caption="InvalidLogin" FieldName="InvalidLogin" 
                            Name="InvalidLogin" Visible="False" VisibleIndex="4">
                        </dx:GridViewDataTextColumn>
                        <dx:GridViewDataTextColumn Caption="Locked" FieldName="Locked" Name="Locked" 
                            Visible="False" VisibleIndex="5">
                        </dx:GridViewDataTextColumn>
                        <dx:GridViewDataTextColumn Caption="StatusAdmin" FieldName="StatusAdmin" 
                            Name="StatusAdmin" Visible="False" VisibleIndex="6">
                        </dx:GridViewDataTextColumn>
                        <dx:GridViewDataTextColumn Caption="Description" FieldName="Description" 
                            Name="Description" Visible="False" VisibleIndex="7">
                        </dx:GridViewDataTextColumn>
                    </Columns>
                    <SettingsBehavior AllowFocusedRow="True" AllowSort="False" 
                        ColumnResizeMode="Control" EnableRowHotTrack="True" />
                    <SettingsPager PageSize="11" NumericButtonCount="10">
                        <Summary AllPagesText="Page {0} of {1} " Text="Page {0} of {1} [{2} record(s)]"></Summary>
                    </SettingsPager>
                    <Settings ShowVerticalScrollBar="True" />
                    <Styles>
                        <Header BackColor="#FFD2A6" Font-Names="Verdana" Font-Size="8pt"></Header>
                        <Row BackColor="#FFFFFF" Font-Names="Verdana" Font-Size="8pt" Wrap="False"></Row>
                        <RowHotTrack BackColor="#E8EFFD" Font-Names="Verdana" Font-Size="8pt" Wrap="False"></RowHotTrack>
                        <SelectedRow Wrap="False">
                        </SelectedRow>
                        <FocusedRow ForeColor="Black" BackColor="#DCE7FC">
                        </FocusedRow>
                    </Styles>
                    <StylesEditors ButtonEditCellSpacing="0">
                        <ProgressBar Height="21px">
                        </ProgressBar>
                    </StylesEditors>
                </dx:ASPxGridView>
            </td>
            <td valign="top" width="50%">
                <table style="width:100%;" frame="box">
                    <tr>
                        <td align="left" width="10px">
                            &nbsp;</td>
                        <td align="left" width="100px">
                            <dx:ASPxLabel ID="ASPxLabel1" runat="server" Font-Names="Verdana" 
                                Font-Size="8pt" Text="Affiliate ID">
                            </dx:ASPxLabel>
                        </td>
                        <td align="left" width="200px">
                            <dx:ASPxComboBox ID="cboAffiliateID" runat="server" 
                                ClientInstanceName="cboAffiliateID" Width="100px" Font-Names="Verdana" 
                                DataSourceID="AffiliateUser" TextField="AffiliateID" 
                                ValueField="AffiliateID" TextFormatString="{0}" Font-Size="8pt">
                                <Columns>
                                    <dx:ListBoxColumn Caption="Affiliate ID" FieldName="AffiliateID" Width="70px" />
                                    <dx:ListBoxColumn Caption="Affiliate Name" FieldName="AffiliateName" Width="180px" />
                                </Columns>
                                <LoadingPanelStyle ImageSpacing="5px">
                                </LoadingPanelStyle>
                            </dx:ASPxComboBox>  
                            <asp:SqlDataSource ID="AffiliateUser" runat="server" 
                                ConnectionString="<%$ ConnectionStrings:KonString %>" 
                                SelectCommand="SELECT AffiliateID, AffiliateName FROM dbo.MS_Affiliate ORDER BY AffiliateID">
                            </asp:SqlDataSource>
                        </td>
                    </tr>
                    <tr>
                        <td align="left" width="10px">
                            &nbsp;</td>
                        <td align="left" width="100px">
                            <dx:ASPxLabel ID="ASPxLabel2" runat="server" Font-Names="Verdana" 
                                Font-Size="8pt" Text="User ID">
                            </dx:ASPxLabel>
                        </td>
                        <td align="left" width="100px">
                            <dx:ASPxTextBox ID="txtUserId" runat="server" Font-Names="Verdana" 
                                Font-Size="8pt" Width="170px" ClientInstanceName="txtUserId" 
                                MaxLength="30">
                            </dx:ASPxTextBox>
                        </td>
                    </tr>
                    <tr>
                        <td align="left" width="10px">
                            &nbsp;</td>
                        <td align="left" width="100px">
                            <dx:ASPxLabel ID="ASPxLabel3" runat="server" Font-Names="Verdana" 
                                Font-Size="8pt" Text="Full Name">
                            </dx:ASPxLabel>
                        </td>
                        <td align="left" width="100px">
                            <dx:ASPxTextBox ID="txtFullName" runat="server" Font-Names="Verdana" 
                                Font-Size="8pt" Width="170px" ClientInstanceName="txtFullName" 
                                MaxLength="25">
                            </dx:ASPxTextBox>
                        </td>
                    </tr>
                    <tr>
                        <td align="left" width="10px">
                            &nbsp;</td>
                        <td align="left" width="100px">
                            <dx:ASPxLabel ID="ASPxLabel4" runat="server" Font-Names="Verdana" 
                                Font-Size="8pt" Text="Password">
                            </dx:ASPxLabel>
                        </td>
                        <td align="left" width="100px">
                            <dx:ASPxTextBox ID="txtPasswordUS" runat="server" Font-Names="Verdana" 
                                Font-Size="8pt" Width="170px" ClientInstanceName="txtPasswordUS" 
                                Password="True">
                                <ClientSideEvents Init="function(s, e) {
	                                s.SetValue(s.cp_myPassword);
                                }" />
                            </dx:ASPxTextBox>
                        </td>
                    </tr>
                    <tr>
                        <td align="left" width="10px">
                            &nbsp;</td>
                        <td align="left" width="100px">
                            <dx:ASPxLabel ID="ASPxLabel5" runat="server" Font-Names="Verdana" 
                                Font-Size="8pt" Text="Confirm Password">
                            </dx:ASPxLabel>
                        </td>
                        <td align="left" width="100px">
                            <dx:ASPxTextBox ID="txtConfPassword" runat="server" Font-Names="Verdana" 
                                Font-Size="8pt" Width="170px" ClientInstanceName="txtConfPassword" 
                                Password="True">
                                <ClientSideEvents Init="function(s, e) {
	                                s.SetValue(s.cp_myPassword);
                                }" />
                            </dx:ASPxTextBox>
                        </td>
                    </tr>
                    <tr>
                        <td align="left" width="10px">
                            &nbsp;</td>
                        <td align="left" width="100px">
                            <dx:ASPxLabel ID="ASPxLabel6" runat="server" Font-Names="Verdana" 
                                Font-Size="8pt" Text="User Group">
                            </dx:ASPxLabel>
                        </td>
                        <td align="left" width="100px">
                            <dx:ASPxComboBox ID="cboUserGroup" runat="server" 
                                ClientInstanceName="cboUserGroup" Width="100px" Font-Names="Verdana" 
                                DataSourceID="UserGroup" TextField="UserID" 
                                ValueField="UserID" TextFormatString="{0}" Font-Size="8pt">
                                <ClientSideEvents SelectedIndexChanged="function(s, e) {
	                                gridMenu.PerformCallback('loadPrevilege');
                                }" />
                                <Columns>
                                    <dx:ListBoxColumn Caption="User ID" FieldName="UserID" Width="50px" />
                                    <dx:ListBoxColumn Caption="Full Name" FieldName="FullName" Width="50px" />
                                </Columns>
                                <LoadingPanelStyle ImageSpacing="5px">
                                </LoadingPanelStyle>
                            </dx:ASPxComboBox>  
                            <asp:SqlDataSource ID="UserGroup" runat="server" 
                                ConnectionString="<%$ ConnectionStrings:KonString %>" 
                                SelectCommand="SELECT UserID, FullName FROM dbo.SC_UserSetup WHERE UserCls = '0' ORDER BY UserID">
                            </asp:SqlDataSource>
                        </td>
                    </tr>
                    <tr>
                        <td align="left" width="10px">
                            &nbsp;</td>
                        <td align="left" width="100px">
                            <dx:ASPxLabel ID="ASPxLabel7" runat="server" Font-Names="Verdana" 
                                Font-Size="8pt" Text="Administrator Status">
                            </dx:ASPxLabel>
                        </td>
                        <td align="left" width="100px">
                            <dx:ASPxRadioButtonList ID="rblAdminStatus" runat="server" 
                                Font-Names="Verdana" Font-Size="8pt" RepeatDirection="Horizontal" 
                                Width="106px" ClientInstanceName="rblAdminStatus">
                                <Items>
                                    <dx:ListEditItem Text="Yes" Value="1" />
                                    <dx:ListEditItem Text="No" Value="0" />
                                </Items>
                            </dx:ASPxRadioButtonList>
                        </td>
                    </tr>
                    <tr>
                        <td align="left" width="10px">
                            &nbsp;</td>
                        <td align="left" width="100px">
                            <dx:ASPxLabel ID="ASPxLabel8" runat="server" Font-Names="Verdana" 
                                Font-Size="8pt" Text="Account Lock">
                            </dx:ASPxLabel>
                        </td>
                        <td align="left" width="100px">
                            <dx:ASPxCheckBox ID="cbAccount" runat="server" Font-Names="Verdana" 
                                Font-Size="8pt" Text="Unlock" ClientInstanceName="cbAccount" 
                                ValueChecked="1" ValueType="System.String" ValueUnchecked="0">
                            </dx:ASPxCheckBox>
                        </td>
                    </tr>
                    <tr>
                        <td align="left" width="10px">
                            &nbsp;</td>
                        <td align="left" width="100px">
                            <dx:ASPxLabel ID="ASPxLabel9" runat="server" Font-Names="Verdana" 
                                Font-Size="8pt" Text="Description">
                            </dx:ASPxLabel>
                        </td>
                        <td align="left" width="100px">
                            <dx:ASPxTextBox ID="txtDesc" runat="server" Font-Names="Verdana" 
                                Font-Size="8pt" Width="170px" ClientInstanceName="txtDesc" MaxLength="25">
                            </dx:ASPxTextBox>
                        </td>
                    </tr>
                    </table>
            </td>
        </tr>
        <tr>
            <td align="left" valign="top" colspan="2">
                <dx:aspxpagecontrol id="ASPxPageControl1" runat="server" 
                    font-names="Verdana" font-size="8pt" width="100%" 
                    ActiveTabIndex="0">
                    <TabPages>
                        <dx:tabpage Name="MenuPrivillege" Text="Menu Privillege">
                            <ContentCollection>
                                <dx:ContentControl runat="server" SupportsDisabledAttribute="True">
                                    <dx:ASPxGridView ID="gridMenu" runat="server" AutoGenerateColumns="False" 
                                        ClientInstanceName="gridMenu" Font-Names="Verdana" Font-Size="8pt" 
                                        KeyFieldName="MenuID" Width="100%">
                                        <ClientSideEvents Init="OnInit" 
                                        EndCallback="function(s, e) { 
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
                                        }" />
                                        <Columns>
                                            <dx:GridViewDataTextColumn Caption="Menu Group" FieldName="GroupID" 
                                                Name="GroupID" ReadOnly="True" ShowInCustomizationForm="True" VisibleIndex="0" 
                                                Width="150px">
                                                <CellStyle Font-Names="Verdana" Font-Size="8pt">
                                                </CellStyle>
                                            </dx:GridViewDataTextColumn>
                                            <dx:GridViewDataTextColumn Caption="Menu ID" FieldName="MenuID" Name="MenuID" 
                                                ShowInCustomizationForm="True" VisibleIndex="1" Width="100px">
                                                <CellStyle Font-Names="Verdana" Font-Size="8pt" HorizontalAlign="Left">
                                                </CellStyle>
                                            </dx:GridViewDataTextColumn>
                                            <dx:GridViewDataTextColumn Caption="Menu Name" FieldName="MenuDesc" 
                                                Name="MenuDesc" ShowInCustomizationForm="True" VisibleIndex="2" 
                                                Width="200px">
                                            </dx:GridViewDataTextColumn>
                                            <dx:GridViewDataCheckColumn Caption="Allow Access" FieldName="AllowAccess" 
                                                Name="AllowAccess" VisibleIndex="3" Width="100px" >
                                                <PropertiesCheckEdit ValueChecked="1" ValueType="System.Int32" 
                                                    ValueUnchecked="0">
                                                </PropertiesCheckEdit>
                                                <HeaderStyle HorizontalAlign="Center"></HeaderStyle>
                                            </dx:GridViewDataCheckColumn>
                                            <dx:GridViewDataCheckColumn Caption="Allow Update" FieldName="AllowUpdate" 
                                                Name="AllowUpdate" VisibleIndex="4" Width="100px" >
                                                <PropertiesCheckEdit ValueChecked="1" ValueType="System.Int32" 
                                                    ValueUnchecked="0">
                                                </PropertiesCheckEdit>
                                                <HeaderStyle HorizontalAlign="Center"></HeaderStyle>
                                            </dx:GridViewDataCheckColumn>
                                            <dx:GridViewDataCheckColumn Caption="Allow Delete" FieldName="AllowDelete" 
                                                Name="AllowDelete" VisibleIndex="5" Width="100px" >
                                                <PropertiesCheckEdit ValueChecked="1" ValueType="System.Int32" 
                                                    ValueUnchecked="0">
                                                </PropertiesCheckEdit>
                                                <HeaderStyle HorizontalAlign="Center"></HeaderStyle>
                                            </dx:GridViewDataCheckColumn>
                                            <dx:GridViewDataCheckColumn Caption="Allow Confirm" FieldName="AllowConfirm" 
                                                Name="AllowConfirm" VisibleIndex="6" Width="100px" >
                                                <PropertiesCheckEdit ValueChecked="1" ValueType="System.Int32" 
                                                    ValueUnchecked="0">
                                                </PropertiesCheckEdit>
                                                <HeaderStyle HorizontalAlign="Center"></HeaderStyle>
                                            </dx:GridViewDataCheckColumn>
                                        </Columns>
                                        <SettingsBehavior AllowFocusedRow="True" AllowSort="False" 
                                            ColumnResizeMode="Control" EnableRowHotTrack="True" />
                                        <SettingsPager NumericButtonCount="10" Mode="ShowAllRecords">
                                            <Summary AllPagesText="Page {0} of {1} " Text="Page {0} of {1} [{2} record(s)]"></Summary>
                                        </SettingsPager>
                                        <SettingsEditing Mode="Batch" NewItemRowPosition="Bottom">
                                            <BatchEditSettings ShowConfirmOnLosingChanges="False" />
                                        </SettingsEditing>
                                        <Settings ShowVerticalScrollBar="True" ShowStatusBar="Hidden" 
                                            VerticalScrollableHeight="160" HorizontalScrollBarMode="Visible" 
                                            VerticalScrollBarMode="Visible" />
                                        <Styles>
                                            <Header BackColor="#FFD2A6" Font-Names="Verdana" Font-Size="8pt">
                                            </Header>
                                            <Row BackColor="#FFFFE1" Font-Names="Verdana" Font-Size="8pt" Wrap="False">
                                            </Row>
                                            <RowHotTrack BackColor="#E8EFFD" Font-Names="Verdana" Font-Size="8pt" 
                                                Wrap="False">
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
                                </dx:ContentControl>
                            </ContentCollection>
                        </dx:tabpage>                        
                    </TabPages>
                    <ClientSideEvents ActiveTabChanged="function(s, e) {
	                    gridMenu.PerformCallback('loadPrevilege');	                    
                    }" ActiveTabChanging="function(s, e) {
	                    gridMenu.PerformCallback('loadPrevilege');
                    }" />
                </dx:aspxpagecontrol>
            </td>
        </tr>
        </table>

    <div style="height:2px;"></div>
    
    <table style="width:100%;">
        <tr>
            <td align="left">
                <dx:ASPxButton ID="btnSubMenu" runat="server" ClientInstanceName="btnSubMenu" 
                Font-Names="Verdana" Text="Sub Menu" Width="90px" Font-Size="8pt">
                </dx:ASPxButton>
            </td>
            <td align="left">
                <dx:ASPxTextBox ID="txtCCTemp" runat="server" ClientInstanceName="txtCCTemp" 
                Width="170px" BackColor="White" ForeColor="White">
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
                <dx:ASPxButton ID="btnClear" runat="server" Text="Clear" AutoPostBack="False" 
                Width="80px" Font-Names="Verdana" Font-Size="8pt">
                <ClientSideEvents Click="function(s, e) {
	            cboAffiliateID.SetText('');
                txtUserId.SetText('');
                txtFullName.SetText('');
                txtPasswordUS.SetText('');
                txtConfPassword.SetText('');
                cboUserGroup.SetText('');
                rblAdminStatus.SetValue(0);
                cbAccount.SetValue(0);
                txtDesc.SetText('');
                txtCCTemp.SetText('');
                txtUserIDTemp.SetText('');

                gridUser.SetFocusedRowIndex(-1);
   	            gridUser.PerformCallback('load');
                gridMenu.PerformCallback('load');
	                            
   	            lblErrMsg.SetText('');
                
                cboAffiliateID.GetInputElement().setAttribute('style', 'background:#FFFFFF;');
                cboAffiliateID.GetInputElement().readOnly = false;

                txtUserId.GetInputElement().setAttribute('style', 'background:#FFFFFF;');
                txtUserId.GetInputElement().readOnly = false;
                }" />
                </dx:ASPxButton>
            </td>
            <td align="center" width="85px">
                <dx:ASPxButton ID="btnDelete" runat="server" Text="Delete" Width="80px" 
                Font-Names="Verdana" Font-Size="8pt" 
                AutoPostBack="False">
                <ClientSideEvents Click="function(s, e) {
                    if (gridUser.GetFocusedRowIndex() == -1) {
                        lblErrMsg.GetMainElement().style.color = 'Red';
		                lblErrMsg.SetText('[6010] Please select the data first!');
                        e.processOnServer = false;
                        return;
                    } 
    
                    var msg = confirm('Are you sure want to delete this data ?');                
                    if (msg == false) {
                        e.processOnServer = false;
                        return;
                    } 
                    var pCCCode = cboAffiliateID.GetText();
                    var pUserID = txtUserId.GetText();

                    gridUser.PerformCallback('delete|'+pCCCode+'|'+pUserID);      
                    gridMenu.PerformCallback('load');
                    cboAffiliateID.SetText('');
                    txtUserId.SetText('');
                    txtFullName.SetText('');
                    txtPasswordUS.SetText('');
                    txtConfPassword.SetText('');
                    cboUserGroup.SetText('');
                    rblAdminStatus.SetValue(0);
                    cbAccount.SetValue(0);
                    txtDesc.SetText('');                    


                    cboAffiliateID.GetInputElement().setAttribute('style', 'background:#FFFFFF;');
                    cboAffiliateID.GetInputElement().readOnly = false;

                    txtUserId.GetInputElement().setAttribute('style', 'background:#FFFFFF;');
                    txtUserId.GetInputElement().readOnly = false;
                }" />
                </dx:ASPxButton>
            </td>
            <td align="right" width="85px">
                <dx:ASPxButton ID="btnSubmit" runat="server" Text="Submit" Width="80px" 
                    Font-Names="Verdana" AutoPostBack="false"
                    Font-Size="8pt" ClientInstanceName="btnSubmit">
                    <ClientSideEvents Click="function(s, e) {
                        if (cboAffiliateID.GetText() == '') {
                            lblErrMsg.GetMainElement().style.color = 'Red';
		                    lblErrMsg.SetText('[6010] Please choose Affiliate ID first!');
                            e.processOnServer = false;
                            return;
                        }
                        if (txtUserId.GetText() == '') {
                            lblErrMsg.GetMainElement().style.color = 'Red';
		                    lblErrMsg.SetText('[6010] Please insert User Id first!');
                            e.processOnServer = false;
                            return;
                        }
                        if (txtFullName.GetText() == '') {
                            lblErrMsg.GetMainElement().style.color = 'Red';
		                    lblErrMsg.SetText('[6010] Please insert Full name first!');
                            e.processOnServer = false;
                            return;
                        }
                        if (txtPasswordUS.GetText() == '') {
                            lblErrMsg.GetMainElement().style.color = 'Red';
		                    lblErrMsg.SetText('[6010] Please insert Password first!');
                            e.processOnServer = false;
                            return;
                        }
                        if (txtConfPassword.GetText() == '') {
                            lblErrMsg.GetMainElement().style.color = 'Red';
		                    lblErrMsg.SetText('[6010] Please insert Confirmation Password first!');
                            e.processOnServer = false;
                            return;
                        }
                        if (txtPasswordUS.GetText() != txtConfPassword.GetText()) {
                            lblErrMsg.GetMainElement().style.color = 'Red';
		                    lblErrMsg.SetText('[6010] Please insert Confirmation Password the same as Password code first!');
                            e.processOnServer = false;
                            return;
                        }

                        var pCCCode = cboAffiliateID.GetText();
                        var pUserID = txtUserId.GetText();
                        var pFullName = txtFullName.GetText();
                        var pPWd = txtPasswordUS.GetText();
                        var pLocked = cbAccount.GetValue();
                        var pStatusAdmin= rblAdminStatus.GetValue();
                        var pDesc = txtDesc.GetText();
                        var pByUserID = cboUserGroup.GetText();
                        
                        gridUser.PerformCallback('save|'+ pCCCode + '|'+ pUserID + '|' + pFullName + '|' + pPWd + '|' + pLocked + '|' + pStatusAdmin + '|' + pDesc + '|' + pByUserID);
	                    gridMenu.UpdateEdit();
                        gridMenu.PerformCallback('load');                        
	                    }" />
                </dx:ASPxButton>
            </td>
        </tr>
    </table>

    <div style="height:8px;"></div> 

    <table style="width:100%;">
        <tr>
            <td>
                <dx:ASPxPopupControl ID="search" runat="server" CloseAction="CloseButton" Modal="True"
                    PopupHorizontalAlign="WindowCenter" PopupVerticalAlign="WindowCenter" ClientInstanceName="search"
                    HeaderText="Please input User ID!" AllowDragging="True" 
                    HeaderStyle-BackColor="#255FDC" HeaderStyle-ForeColor="White" headerstyle-Font-Names="Verdana"
                    PopupAnimationType="None" EnableViewState="False" Width="280px" 
                    BackColor="#96C8FF" Font-Names="Verdana">
                    <ClientSideEvents PopUp="function(s, e) { txtsearch.Focus(); }" />
                    <HeaderStyle BackColor="#255FDC" Font-Names="Verdana" ForeColor="White"></HeaderStyle>
                    <ContentCollection>
                        <dx:PopupControlContentControl ID="PopupControlContentControl1" runat="server">
                            <dx:ASPxPanel ID="Panel1" runat="server" DefaultButton="btOK">
                                <PanelCollection>
                                    <dx:PanelContent ID="PanelContent1" runat="server" Width="280px">
                                        <table width="280px">
                                            <tr>                                                            
                                                <td valign="top" width="100px">
                                                    <dx:ASPxLabel ID="lblUsername1" runat="server" Text="User ID:" 
                                                        Font-Names="Verdana" >
                                                    </dx:ASPxLabel>
                                                </td>
                                                <td width="230px">
                                                    <dx:ASPxTextBox ID="txtsearch" runat="server" Width="230px" 
                                                        ClientInstanceName="txtsearch" MaxLength="15" Font-Names="Verdana">                                                                    
                                                    </dx:ASPxTextBox>
                                                </td>
                                            </tr>
                                            <tr>
                                                <td colspan="2" >                                                                
                                                </td>
                                            </tr>                                                        
                                            <tr>
                                                <td colspan="2">
                                                    <table width="280px">
                                                    <tr>
                                                    <td align="right">
                                                        <dx:ASPxButton ID="btOK" runat="server" Text="OK" Width="80px" 
                                                            AutoPostBack="False" CssFilePath="~/App_Themes/DevEx/{0}/styles.css" 
                                                            CssPostfix="DevEx" SpriteCssFilePath="~/App_Themes/DevEx/{0}/sprite.css" 
                                                            Font-Names="Verdana" >
                                                            <ClientSideEvents Click="function(s, e) { 
                                                            lblErrMsg.SetText('');
                                                            ASPxCallback1.PerformCallback('search');
                                                            ASPxCallback2.PerformCallback('search');
                                                            search.Hide();
                                                            txtsearch.SetText('');}" />                                                        
                                                        </dx:ASPxButton>
                                                    </td>
                                                    <td align="right" style="width:85px;">
                                                        &nbsp;</td>
                                                    </tr>
                                                    </table>
                                                </td>
                                            </tr>
                                        </table>
                                    </dx:PanelContent>
                                </PanelCollection>
                            </dx:ASPxPanel>                                        
                        </dx:PopupControlContentControl>
                    </ContentCollection>                                
                    <Border BorderColor="#255FDC"></Border>
                </dx:ASPxPopupControl>
                        <dx:ASPxCallback ID="ASPxCallback1" runat="server" 
                            ClientInstanceName="ASPxCallback1">
                            <ClientSideEvents CallbackComplete="function(s, e) {
	                            txtPasswordUS.SetText(e.result);
                                txtConfPassword.SetText(e.result);
                            }" />
                        </dx:ASPxCallback>
                        <dx:ASPxCallback ID="ASPxCallback2" runat="server" 
                            ClientInstanceName="ASPxCallback2">
                            <ClientSideEvents EndCallback="function(s, e) {
	                            cboAffiliateID.SetText(s.cpCCCode);
                                txtUserId.SetText(s.cpUserId);
                                txtFullName.SetText(s.cpFullName);
                                rblAdminStatus.SetValue(s.cpStatusAdmin);
                                cbAccount.SetValue(s.cpLocked);
                                txtDesc.SetText(s.cpDescription);
                                gridMenu.PerformCallback('loadPrevilege');                                                           
                            }" />
                        </dx:ASPxCallback>
            </td>
        </tr>
    </table>

    <dx:ASPxGlobalEvents ID="ge" runat="server" >
        <ClientSideEvents ControlsInitialized="function(s, e) {
        OnControlsInitializedSplitter();
        OnControlsInitializedGrid();
        }" />
    </dx:ASPxGlobalEvents> 
</asp:Content>
