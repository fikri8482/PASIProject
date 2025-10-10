<%@ Page Title="" Language="vb" AutoEventWireup="false" MasterPageFile="~/SiteContent.Master" CodeBehind="UserSetup.aspx.vb" Inherits="PASISystem.UserSetup" %>
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
        gridMenu.PerformCallback('load');
    }
    
    function OnUpdateClick(s, e) {
        gridMenu.PerformCallback("Update");
    }

    function OnCancelClick(s, e) {
        gridMenu.PerformCallback("Cancel");
    }

    function grid_SelectionChanged(s, e) {
        s.GetSelectedFieldValues("IDNo", GetSelectedFieldValuesCallback);

    }

    function GetSelectedFieldValuesCallback(values) {
        selList.BeginUpdate();
        try {
            selList.ClearItems();
            for (var i = 0; i < values.length; i++) {
                selList.AddItem(values[i]);
            }
        } finally {
            selList.EndUpdate();
        }
        document.getElementById("selCount").innerHTML = grid.GetSelectedRowCount();
    }

    function OnBatchEditStartEditing(s, e) {
        currentColumnName = e.focusedColumn.fieldName;

        if (currentColumnName == "GroupID" || currentColumnName == "MenuID" || currentColumnName == "MenuDesc") {
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
    <table style="border-width: thin; border-style: inset hidden ridge hidden; border-color:#9598A1; width:100%; height: 15px;">
        <tr>
            <td>
                <!-- MESSAGE AREA #C0C0C0 -->
                <table id="tblMsg" style="width:100%; height: 15px;">
                    <tr>
                        <td align="left" valign="middle" style="height:15px;">
                            <dx:ASPxLabel ID="lblErrMsg" runat="server" Text="" Font-Names="Tahoma" 
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
            <td valign="top" width="40%">
                <dx:ASPxGridView ID="gridUser" runat="server" Width="100%" 
                    Font-Names="Tahoma" AutoGenerateColumns="False" 
                    ClientIDMode="Predictable" KeyFieldName="AffiliateID;UserID" 
                    ClientInstanceName="gridUser" Font-Size="8pt">
                    <ClientSideEvents 
                        Init="OnInit"
                        FocusedRowChanged="function(s, e) {OnGridFocusedRowChangedUser();}" 
                        EndCallback="function(s, e) { 
                        var pMsg = s.cpMessage;
                        if (pMsg != '') {
                            if (pMsg.substring(1,5) == '6011' || pMsg.substring(1,5) == '9999') {
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
                        <dx:GridViewDataTextColumn Caption="AFFILIATE ID" FieldName="AffiliateID" Name="AffiliateID" 
                            Width="80px" VisibleIndex="0">
                            <CellStyle Font-Names="Tahoma" Font-Size="8pt">
                            </CellStyle>
                        </dx:GridViewDataTextColumn>
                        <dx:GridViewDataTextColumn Caption="USER ID" 
                            FieldName="UserID" VisibleIndex="1" Name="UserID" ReadOnly="True" 
                            Width="100px">
                            <CellStyle Font-Names="Tahoma" Font-Size="8pt">
                            </CellStyle>
                        </dx:GridViewDataTextColumn>
                        <dx:GridViewDataTextColumn Caption="USER NAME" FieldName="FullName" Name="FullName" 
                            VisibleIndex="2" Width="100px">
                            <CellStyle Font-Names="Tahoma" Font-Size="8pt" HorizontalAlign="Left">
                            </CellStyle>
                        </dx:GridViewDataTextColumn>
                        <dx:GridViewDataTextColumn Caption="PASSWORD" FieldName="Password" 
                            Name="Password" Visible="False" VisibleIndex="3">
                        </dx:GridViewDataTextColumn>
                        <dx:GridViewDataTextColumn Caption="INVALID LOGIN" FieldName="InvalidLogin" 
                            Name="InvalidLogin" Visible="False" VisibleIndex="4">
                        </dx:GridViewDataTextColumn>
                        <dx:GridViewDataTextColumn Caption="LOCKED" FieldName="Locked" Name="Locked" 
                            Visible="False" VisibleIndex="5">
                        </dx:GridViewDataTextColumn>
                        <dx:GridViewDataTextColumn Caption="STATUS ADMIN" FieldName="StatusAdmin" 
                            Name="StatusAdmin" Visible="False" VisibleIndex="6">
                        </dx:GridViewDataTextColumn>
                        <dx:GridViewDataTextColumn Caption="DESCRIPTION" FieldName="Description" 
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
                        <Header BackColor="#FFD2A6" Font-Names="Tahoma" Font-Size="8pt"></Header>
                        <Row BackColor="#FFFFFF" Font-Names="Tahoma" Font-Size="8pt" Wrap="False"></Row>
                        <RowHotTrack BackColor="#E8EFFD" Font-Names="Tahoma" Font-Size="8pt" Wrap="False"></RowHotTrack>
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
                            <dx:ASPxLabel ID="ASPxLabel1" runat="server" Font-Names="Tahoma" 
                                Font-Size="8pt" Text="AFFILIATE ID">
                            </dx:ASPxLabel>
                        </td>
                        <td align="left" width="200px">
                            <dx:ASPxComboBox ID="cboAffiliateID" runat="server" 
                                ClientInstanceName="cboAffiliateID" Width="100px" Font-Names="Tahoma" 
                                DataSourceID="AffiliateUser" TextField="AffiliateID" 
                                ValueField="AffiliateID" TextFormatString="{0}" Font-Size="8pt">
                                <ClientSideEvents SelectedIndexChanged="function(s, e) {
	                                gridMenu.PerformCallback('load');
                                    cboUserGroup.PerformCallback();
                                }" />
                                <Columns>
                                    <dx:ListBoxColumn Caption="AFFILIATE ID" FieldName="AffiliateID" Width="70px" />
                                    <dx:ListBoxColumn Caption="AFFILIATE NAME" FieldName="AffiliateName" Width="180px" />
                                </Columns>
                                <LoadingPanelStyle ImageSpacing="5px">
                                </LoadingPanelStyle>
                            </dx:ASPxComboBox>  
                            <asp:SqlDataSource ID="AffiliateUser" runat="server" 
                                ConnectionString="<%$ ConnectionStrings:KonString %>" 
                                SelectCommand="SELECT RTRIM(AffiliateID)AffiliateID, AffiliateName FROM dbo.MS_Affiliate ORDER BY AffiliateID">
                            </asp:SqlDataSource>
                        </td>
                    </tr>
                    <tr>
                        <td align="left" width="10px">
                            &nbsp;</td>
                        <td align="left" width="100px">
                            <dx:ASPxLabel ID="ASPxLabel2" runat="server" Font-Names="Tahoma" 
                                Font-Size="8pt" Text="USER ID">
                            </dx:ASPxLabel>
                        </td>
                        <td align="left" width="100px">
                            <dx:ASPxTextBox ID="txtUserId" runat="server" Font-Names="Tahoma" 
                                Font-Size="8pt" Width="170px" ClientInstanceName="txtUserId" 
                                MaxLength="30">
                                <ClientSideEvents LostFocus="function(s, e) {
	                                txtUserIDTemp.SetText(txtUserId.GetText());
                                }" />
                            </dx:ASPxTextBox>
                        </td>
                    </tr>
                    <tr>
                        <td align="left" width="10px">
                            &nbsp;</td>
                        <td align="left" width="100px">
                            <dx:ASPxLabel ID="ASPxLabel3" runat="server" Font-Names="Tahoma" 
                                Font-Size="8pt" Text="FULL NAME">
                            </dx:ASPxLabel>
                        </td>
                        <td align="left" width="100px">
                            <dx:ASPxTextBox ID="txtFullName" runat="server" Font-Names="Tahoma" 
                                Font-Size="8pt" Width="170px" ClientInstanceName="txtFullName" 
                                MaxLength="25">
                            </dx:ASPxTextBox>
                        </td>
                    </tr>
                    <tr>
                        <td align="left" width="10px">
                            &nbsp;</td>
                        <td align="left" width="100px">
                            <dx:ASPxLabel ID="ASPxLabel4" runat="server" Font-Names="Tahoma" 
                                Font-Size="8pt" Text="PASSWORD">
                            </dx:ASPxLabel>
                        </td>
                        <td align="left" width="100px">
                            <dx:ASPxTextBox ID="txtPasswordUS" runat="server" Font-Names="Tahoma" 
                                Font-Size="8pt" Width="170px" ClientInstanceName="txtPasswordUS" AutoCompleteType ="None"
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
                            <dx:ASPxLabel ID="ASPxLabel5" runat="server" Font-Names="Tahoma" 
                                Font-Size="8pt" Text="CONFIRM PASSWORD">
                            </dx:ASPxLabel>
                        </td>
                        <td align="left" width="100px">
                            <dx:ASPxTextBox ID="txtConfPassword" runat="server" Font-Names="Tahoma" 
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
                            <dx:ASPxLabel ID="ASPxLabel6" runat="server" Font-Names="Tahoma" 
                                Font-Size="8pt" Text="USER GROUP">
                            </dx:ASPxLabel>
                        </td>
                        <td align="left" width="100px">
                            <dx:ASPxComboBox ID="cboUserGroup" runat="server" 
                                ClientInstanceName="cboUserGroup" Width="210px" Font-Names="Tahoma" 
                                TextFormatString="{0}" Font-Size="8pt">
                                <ClientSideEvents SelectedIndexChanged="function(s, e) {
	                                gridMenu.PerformCallback('loadPrevilege');
                                }" />
                                <LoadingPanelStyle ImageSpacing="5px">
                                </LoadingPanelStyle>
                            </dx:ASPxComboBox>                           
                        </td>
                    </tr>
                    <tr>
                        <td align="left" width="10px">
                            &nbsp;</td>
                        <td align="left" width="100px">
                            <dx:ASPxLabel ID="ASPxLabel7" runat="server" Font-Names="Tahoma" 
                                Font-Size="8pt" Text="ADMINISTRATOR STATUS">
                            </dx:ASPxLabel>
                        </td>
                        <td align="left" width="100px">
                            <dx:ASPxRadioButtonList ID="rblAdminStatus" runat="server" 
                                Font-Names="Tahoma" Font-Size="8pt" RepeatDirection="Horizontal" 
                                Width="106px" ClientInstanceName="rblAdminStatus">
                                <Items>
                                    <dx:ListEditItem Text="YES" Value="1" />
                                    <dx:ListEditItem Text="NO" Value="0" />
                                </Items>
                            </dx:ASPxRadioButtonList>
                        </td>
                    </tr>
                    <tr>
                        <td align="left" width="10px">
                            &nbsp;</td>
                        <td align="left" width="100px">
                            <dx:ASPxLabel ID="ASPxLabel8" runat="server" Font-Names="Tahoma" 
                                Font-Size="8pt" Text="ACCOUNT LOCK">
                            </dx:ASPxLabel>
                        </td>
                        <td align="left" width="100px">
                            <dx:ASPxCheckBox ID="cbAccount" runat="server" Font-Names="Tahoma" 
                                Font-Size="8pt" Text="UNLOCK" ClientInstanceName="cbAccount" 
                                ValueChecked="1" ValueType="System.String" ValueUnchecked="0">
                            </dx:ASPxCheckBox>
                        </td>
                    </tr>
                    <tr>
                        <td align="left" width="10px">
                            &nbsp;</td>
                        <td align="left" width="100px">
                            <dx:ASPxLabel ID="ASPxLabel9" runat="server" Font-Names="Tahoma" 
                                Font-Size="8pt" Text="DESCRIPTION">
                            </dx:ASPxLabel>
                        </td>
                        <td align="left" width="100px">
                            <dx:ASPxTextBox ID="txtDesc" runat="server" Font-Names="Tahoma" 
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
                    Font-Names="Tahoma" font-size="8pt" width="100%" 
                    ActiveTabIndex="0">
                    <TabPages>
                        <dx:tabpage Name="MenuPrivillege" Text="MENU PRIVILLEGE">
                            <ContentCollection>
                                <dx:ContentControl runat="server" SupportsDisabledAttribute="True">
                                    <dx:ASPxGridView ID="gridMenu" runat="server" AutoGenerateColumns="False" 
                                        ClientInstanceName="gridMenu" Font-Names="Tahoma" Font-Size="8pt" 
                                        KeyFieldName="MenuID" Width="100%">
                                        <ClientSideEvents Init="OnInit" BatchEditStartEditing="OnBatchEditStartEditing" 
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
                                            <dx:GridViewDataTextColumn Caption="MENU GROUP" FieldName="GroupID" 
                                                Name="GroupID" ReadOnly="True" ShowInCustomizationForm="True" VisibleIndex="0" 
                                                Width="150px">
                                                <CellStyle Font-Names="Tahoma" Font-Size="8pt">
                                                </CellStyle>
                                            </dx:GridViewDataTextColumn>
                                            <dx:GridViewDataTextColumn Caption="MENU ID" FieldName="MenuID" Name="MenuID" 
                                                ShowInCustomizationForm="True" VisibleIndex="1" Width="100px">
                                                <CellStyle Font-Names="Tahoma" Font-Size="8pt" HorizontalAlign="Left">
                                                </CellStyle>
                                            </dx:GridViewDataTextColumn>
                                            <dx:GridViewDataTextColumn Caption="MENU NAME" FieldName="MenuDesc" 
                                                Name="MenuDesc" ShowInCustomizationForm="True" VisibleIndex="2" 
                                                Width="200px">
                                            </dx:GridViewDataTextColumn>
                                            <dx:GridViewDataCheckColumn Caption="ALLOW ACCESS" FieldName="AllowAccess" 
                                                Name="AllowAccess" VisibleIndex="3" Width="100px" >
                                                <PropertiesCheckEdit ValueChecked="1" ValueType="System.Int32" 
                                                    ValueUnchecked="0">
                                                </PropertiesCheckEdit>
                                                <Settings HeaderFilterMode="CheckedList" />
                                                <HeaderStyle HorizontalAlign="Center"></HeaderStyle>
                                            </dx:GridViewDataCheckColumn>
                                            <dx:GridViewDataCheckColumn Caption="ALLOW UPDATE" FieldName="AllowUpdate" 
                                                Name="AllowUpdate" VisibleIndex="4" Width="100px" >
                                                <PropertiesCheckEdit ValueChecked="1" ValueType="System.Int32" 
                                                    ValueUnchecked="0">
                                                </PropertiesCheckEdit>
                                                <HeaderStyle HorizontalAlign="Center"></HeaderStyle>
                                            </dx:GridViewDataCheckColumn>
                                            <dx:GridViewDataCheckColumn Caption="ALLOW DELETE" FieldName="AllowDelete" 
                                                Name="AllowDelete" VisibleIndex="5" Width="100px" >
                                                <PropertiesCheckEdit ValueChecked="1" ValueType="System.Int32" 
                                                    ValueUnchecked="0">
                                                </PropertiesCheckEdit>
                                                <HeaderStyle HorizontalAlign="Center"></HeaderStyle>
                                            </dx:GridViewDataCheckColumn>
                                            <dx:GridViewDataCheckColumn Caption="ALLOW DOWNLOAD" FieldName="AllowConfirm" 
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
                                            <Header BackColor="#FFD2A6" Font-Names="Tahoma" Font-Size="8pt">
                                            </Header>
                                            <Row BackColor="#FFFFE1" Font-Names="Tahoma" Font-Size="8pt" Wrap="False">
                                            </Row>
                                            <RowHotTrack BackColor="#E8EFFD" Font-Names="Tahoma" Font-Size="8pt" 
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
                    Font-Names="Tahoma" Text="SUB MENU" Width="90px" Font-Size="8pt">
                </dx:ASPxButton>
<%--                <dx:ASPxListBox ID="selList" ClientInstanceName="selList" runat="server" Height="0px"
			        Width="100%" EnableTheming="True" Theme="Default" ClientVisible="False" >
                    <Columns>
                        <dx:ListBoxColumn Caption="ID" FieldName="IDNo" Width="100px" />
                    </Columns>
                </dx:ASPxListBox>--%>
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
                <dx:ASPxButton ID="btnClear" runat="server" Text="CLEAR" AutoPostBack="False" 
                Width="80px" Font-Names="Tahoma" Font-Size="8pt">
                <ClientSideEvents Click="function(s, e) {
	            cboAffiliateID.SetText('PASI');
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
                <dx:ASPxButton ID="btnDelete" runat="server" Text="DELETE" Width="80px" 
                Font-Names="Tahoma" Font-Size="8pt" 
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
                    txtCCTemp.SetText('');
                    txtUserIDTemp.SetText('');

                    cboAffiliateID.GetInputElement().setAttribute('style', 'background:#FFFFFF;');
                    cboAffiliateID.GetInputElement().readOnly = false;

                    txtUserId.GetInputElement().setAttribute('style', 'background:#FFFFFF;');
                    txtUserId.GetInputElement().readOnly = false;
                }" />
                </dx:ASPxButton>
            </td>
            <td align="right" width="85px">
                <dx:ASPxButton ID="btnSubmit" runat="server" Text="SAVE" Width="80px" 
                    Font-Names="Tahoma" AutoPostBack="false"
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
                        var pUserCLs; 
                        
                        if (cboAffiliateID.GetText() == 'PASI' || cboAffiliateID.GetText() == 'PASI-AW'){
                            pUserCLs = '1';
                        } else {
                            pUserCLs = '0';
                        }
                        
                        gridUser.PerformCallback('save|'+ pCCCode + '|'+ pUserID + '|' + pFullName + '|' + pPWd + '|' + pLocked + '|' + pStatusAdmin + '|' + pDesc + '|' + pByUserID + '|' + pUserCLs);
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
                <dx:ASPxCallback ID="ASPxCallback1" runat="server" 
                    ClientInstanceName="ASPxCallback1">
                    <ClientSideEvents CallbackComplete="function(s, e) {
	                    txtPasswordUS.SetText(e.result);
                        txtConfPassword.SetText(e.result);
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
