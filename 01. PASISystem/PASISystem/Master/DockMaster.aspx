<%@ Page Title="" Language="vb" AutoEventWireup="false" MasterPageFile="~/SiteContent.Master"
    CodeBehind="DockMaster.aspx.vb" Inherits="PASISystem.DockMaster" %>

<%@ Register Assembly="DevExpress.Web.v14.1, Version=14.1.6.0, Culture=neutral, PublicKeyToken=b88d1754d700e49a"
    Namespace="DevExpress.Web.ASPxGridView.Export" TagPrefix="dx" %>
<%@ Register Assembly="DevExpress.Web.v14.1, Version=14.1.6.0, Culture=neutral, PublicKeyToken=b88d1754d700e49a"
    Namespace="DevExpress.Web.ASPxGridView" TagPrefix="dx" %>
<%@ Register Assembly="DevExpress.Web.v14.1, Version=14.1.6.0, Culture=neutral, PublicKeyToken=b88d1754d700e49a"
    Namespace="DevExpress.Web.ASPxEditors" TagPrefix="dx" %>
<%@ Register Assembly="DevExpress.Web.v14.1, Version=14.1.6.0, Culture=neutral, PublicKeyToken=b88d1754d700e49a"
    Namespace="DevExpress.Web.ASPxGlobalEvents" TagPrefix="dx" %>
<%@ Register Assembly="DevExpress.Web.v14.1, Version=14.1.6.0, Culture=neutral, PublicKeyToken=b88d1754d700e49a"
    Namespace="DevExpress.Web.ASPxCallback" TagPrefix="dx" %>
    <%@ Register Assembly="DevExpress.Web.v14.1, Version=14.1.6.0, Culture=neutral, PublicKeyToken=b88d1754d700e49a" Namespace="DevExpress.Web.ASPxHiddenField" TagPrefix="dx" %>
<asp:Content ID="Content1" ContentPlaceHolderID="head" runat="server">
    <style type="text/css">
        .dxeHLC, .dxeHC, .dxeHFC
        {
            display: none;
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
            height = height - (height * 58 / 100)
            grid.SetHeight(height);
        }

        function OnBatchEditStartEditing(s, e) {
            currentColumnName = e.focusedColumn.fieldName;

            if (currentColumnName == "NoUrut" || currentColumnName == "AffiliateID" || currentColumnName == "AffiliateName" || currentColumnName == "DockID") {
                e.cancel = true;
            }

            currentEditableVisibleIndex = e.visibleIndex;
        }

        function OnGridFocusedRowChanged() {
            grid.GetRowValues(grid.GetFocusedRowIndex(), "AffiliateID;AffiliateName;DockID", OnGetRowValues);
        }
        function OnGetRowValues(values) {
            if (values[0] != "" && values[0] != null && values[0] != "null") {

                cboAffiliate2.SetText(values[0]);
                txtAffiliate2.SetText(values[1]);                
                txtDock.SetText(values[2]);

                HF.Set('hfTest', values[2]);

                lblInfo.SetText('');

                cboAffiliate2.GetInputElement().setAttribute('style', 'background:#CCCCCC;foreground:#CCCCCC;');
                cboAffiliate2.GetInputElement().readOnly = true;
                cboAffiliate2.SetEnabled(false);
            }
        }

        function up_delete() {
            if (cboAffiliate2.GetText() == "") {
                lblInfo.GetMainElement().style.color = 'Red';
                lblInfo.SetText("[6011] Please select the data first!");
                e.ProcessOnServer = false;
                return false;
            }

//            if (grid.GetFocusedRowIndex() == -1) {
//                lblInfo.GetMainElement().style.color = 'Red';
//                lblInfo.SetText("[6011] Please select the data first!");
//                e.ProcessOnServer = false;
//                return false;
//            }

            var msg = confirm('Are you sure want to delete this data ?');
            if (msg == false) {
                e.processOnServer = false;
                return;
            }

            
            var pAffiliateID = cboAffiliate2.GetText();
            var pSupplierID = txtDock.GetText();

            grid.PerformCallback('delete|' + pAffiliateID + '|' + pSupplierID);
        }

        function validasubmit() {
            lblInfo.GetMainElement().style.color = 'Red';

            if (cboAffiliate2.GetText() == "") {
                lblInfo.SetText("[6011] Please Select Affiliate first!");
                cboAffiliate2.Focus();
                e.ProcessOnServer = false;
                return false;
            }

            if (txtDock.GetText() == "") {
                lblInfo.SetText("[6011] Please Input Dock ID!");
                txtDock.Focus();
                e.ProcessOnServer = false;
                return false;
            }
        }

        function up_Insert() {            
            var pAffiliateID = cboAffiliate2.GetSelectedItem().GetColumnText(0);            
            var pQuota = txtDock.GetValue();

            grid.PerformCallback('save|' + pAffiliateID + '|' + pQuota);
        }
    </script>
</asp:Content>
<asp:Content ID="Content2" ContentPlaceHolderID="Content" runat="server">
    <table style="width: 100%;">
        <tr>
            <td>
                <table style="border-width: 1pt thin thin thin; border-style: ridge; border-color: #9598A1;
                    width: 100%;">
                    <tr>
                        <td colspan="8" height="30">
                            <table id="Table1">                                
                                <tr>
                                    <td style="width: 5px;">
                                    </td>
                                    <td align="left" valign="middle" height="25px" width="80px">
                                        <dx:ASPxLabel ID="ASPxLabel2" runat="server" Text="AFFILIATE CODE" Font-Names="Verdana"
                                            Font-Size="8pt">
                                        </dx:ASPxLabel>
                                    </td>
                                    <td align="left" valign="middle" height="25px" width="150px">
                                        <dx:ASPxComboBox ID="cboAffiliate" runat="server" ClientInstanceName="cboAffiliate"
                                            Width="100%" Font-Size="8pt" Font-Names="Verdana" TextFormatString="{0}">
                                            <ClientSideEvents SelectedIndexChanged="function(s, e) {
	                                            txtAffiliate.SetText(cboAffiliate.GetSelectedItem().GetColumnText(1));
                                                grid.PerformCallback('kosong');	                                
	                                            lblInfo.SetText('');	
                                            }" />
                                            <LoadingPanelStyle ImageSpacing="5px">
                                            </LoadingPanelStyle>
                                        </dx:ASPxComboBox>
                                    </td>
                                    <td align="left" valign="middle" height="25px" width="200px">
                                        <dx:ASPxTextBox ID="txtAffiliate" runat="server" Width="100%" Height="20px" ClientInstanceName="txtAffiliate"
                                            Font-Names="Verdana" Font-Size="8pt" MaxLength="100" BackColor="#CCCCCC" ReadOnly="True">
                                        </dx:ASPxTextBox>
                                    </td>
                                    <td align="right" valign="middle" style="height: 25px; width: 100px;">
                                        <table style="width: 100%;">
                                            <tr>
                                                <td align="right" valign="middle" style="height: 25px; width: 90px;">
                                                    <dx:ASPxButton ID="btnRefresh" runat="server" Text="SEARCH" Font-Names="Verdana"
                                                        Width="85px" AutoPostBack="False" Font-Size="8pt">
                                                        <ClientSideEvents Click="function(s, e) {
                                                            grid.PerformCallback('load');
                                                            lblInfo.SetText('');
                                                        }" />
                                                    </dx:ASPxButton>
                                                </td>
                                                <td align="right" valign="middle" style="height: 25px; width: 90px;">
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
    <div style="height: 1px;">
    </div>
    <table style="width: 100%; height: 15px;">
        <tr>
            <td colspan="8" height="15">
                <%--error message--%>
                <table id="tblMsg" style="border-width: thin; border-style: inset hidden ridge hidden;
                    border-color: #9598A1; width: 100%; height: 15px;">
                    <tr>
                        <td align="left" valign="middle" height="15px">
                            <dx:ASPxLabel ID="lblInfo" runat="server" Text="[lblinfo]" Font-Names="Verdana" ClientInstanceName="lblInfo"
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
            <td colspan="8" align="left" valign="top" height="220">
                <%--Column : Grid--%>
                <dx:ASPxGridView ID="grid" runat="server" Width="100%" Font-Names="Verdana" KeyFieldName="NoUrut"
                    AutoGenerateColumns="False" ClientInstanceName="grid" Font-Size="8pt">
                    <ClientSideEvents Init="OnInit" FocusedRowChanged="function(s, e) {
	                    OnGridFocusedRowChanged();
                    }" EndCallback="function(s, e) {
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

                            if (s.cpFunction == 'delete'){
                                if (s.cpType != 'error'){
                                    clear();
                                }
                            }else if(s.cpFunction == 'insert'){
                                clear();
                            }
                        } else {
                            lblInfo.SetText('');
                        }  
                    }" RowClick="function(s, e) {
	                    lblInfo.SetText('');
                    }" BatchEditStartEditing="OnBatchEditStartEditing" />
                    <Columns>
                        <dx:GridViewDataTextColumn VisibleIndex="0" Caption="NO" FieldName="NoUrut" Width="50px"
                            HeaderStyle-HorizontalAlign="Center">
                            <HeaderStyle HorizontalAlign="Center"></HeaderStyle>
                        </dx:GridViewDataTextColumn>                        
                        <dx:GridViewDataTextColumn VisibleIndex="1" Caption="AFFILIATE CODE" FieldName="AffiliateID"
                            Width="80px" HeaderStyle-HorizontalAlign="Center">
                            <HeaderStyle HorizontalAlign="Center"></HeaderStyle>
                        </dx:GridViewDataTextColumn>
                        <dx:GridViewDataTextColumn VisibleIndex="2" Caption="AFFILIATE NAME" FieldName="AffiliateName"
                            Width="210px" HeaderStyle-HorizontalAlign="Center">
                            <HeaderStyle HorizontalAlign="Center"></HeaderStyle>
                        </dx:GridViewDataTextColumn>                        
                        <dx:GridViewDataTextColumn VisibleIndex="3" Caption="DOCK ID" FieldName="DockID"
                            Width="80px" HeaderStyle-HorizontalAlign="Center">
                            <HeaderStyle HorizontalAlign="Center"></HeaderStyle>
                        </dx:GridViewDataTextColumn>
                    </Columns>
                    <SettingsBehavior AllowSelectByRowClick="True" AllowSort="False" ColumnResizeMode="Control"
                        EnableRowHotTrack="True" />
                    <SettingsPager Visible="False" PageSize="13" Position="Top">
                        <Summary Text="Page {0} of {1} [{2} record(s)]" AllPagesText="Page {0} of {1} " />
                        <Summary AllPagesText="Page {0} of {1} " Text="Page {0} of {1} [{2} record(s)]">
                        </Summary>
                    </SettingsPager>
                    <SettingsEditing Mode="Batch" NewItemRowPosition="Bottom">
                        <BatchEditSettings ShowConfirmOnLosingChanges="False" />
                    </SettingsEditing>
                    <Settings ShowHorizontalScrollBar="True" ShowVerticalScrollBar="True" ShowGroupButtons="False"
                        ShowStatusBar="Hidden" VerticalScrollableHeight="190" />
                    <SettingsCommandButton EditButton-ButtonType="Link">
                        <EditButton Text="Detail">
                        </EditButton>
                    </SettingsCommandButton>
                    <Styles>
                        <SelectedRow ForeColor="Black">
                        </SelectedRow>
                        <FocusedRow ForeColor="Black">
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
    <div style="height: 8px;">
    </div>
    <table style="width: 100%;">
        <tr>
            <td colspan="8" height="70">
                <!-- INPUT AREA -->
                <table id="tbl1" style="border-width: 1pt thin thin thin; border-style: ridge; border-color: #9598A1;
                    width: 100%; height: 25px; background-color: #FFD2A6">
                    <tr>
                        <td valign="top" style="width: 130px;">
                            <dx:ASPxLabel ID="ASPxLabel56" runat="server" Text="AFFILIATE CODE" Font-Names="Verdana"
                                Font-Size="8pt" Width="100%">
                            </dx:ASPxLabel>
                        </td>
                        <td valign="top" style="width: 200px;">
                            <dx:ASPxLabel ID="ASPxLabel57" runat="server" Text="AFFILIATE NAME" Font-Names="Verdana"
                                Font-Size="8pt" Width="100%">
                            </dx:ASPxLabel>
                        </td>                        
                        <td valign="top" style="width: 150px;">
                            <dx:ASPxLabel ID="ASPxLabel5" runat="server" Text="DOCK ID" Font-Names="Verdana"
                                Font-Size="8pt" Width="100%">
                            </dx:ASPxLabel>
                        </td>
                        <td>
                        </td>
                    </tr>
                </table>
                <table id="tbl2" style="border-width: 1pt thin thin thin; border-style: ridge; border-color: #9598A1;
                    width: 100%; height: 35px;">
                    <tr>                        
                        <td style="width: 130px;">
                            <dx:ASPxComboBox ID="cboAffiliate2" runat="server" ClientInstanceName="cboAffiliate2"
                                Width="130px" Font-Size="8pt" Font-Names="Verdana" TextFormatString="{0}">
                                <ClientSideEvents SelectedIndexChanged="function(s, e) {
	                                txtAffiliate2.SetText(cboAffiliate2.GetSelectedItem().GetColumnText(1));	                                
	                                lblInfo.SetText('');	
                                }" />
                                <LoadingPanelStyle ImageSpacing="5px">
                                </LoadingPanelStyle>
                            </dx:ASPxComboBox>
                        </td>
                        <td style="width: 200px;">
                            <dx:ASPxTextBox ID="txtAffiliate2" runat="server" Width="200px" Height="20px" ClientInstanceName="txtAffiliate2"
                                Font-Names="Verdana" Font-Size="8pt" MaxLength="50" BackColor="#CCCCCC" ReadOnly="True">
                            </dx:ASPxTextBox>
                        </td>                        
                        <td style="width: 150px;">
                            <dx:ASPxTextBox ID="txtDock" runat="server" Width="100%" Height="20px" ClientInstanceName="txtDock"
                                Font-Names="Verdana" Font-Size="8pt" MaxLength="5">
                                <ClientSideEvents TextChanged="function(s, e) {
	                                lblInfo.SetText('');
                                }" />
                            </dx:ASPxTextBox>
                        </td>
                        <td>
                        </td>
                    </tr>
                </table>
            </td>
        </tr>
    </table>
    <div style="height: 8px;">
    </div>
    <table id="button" style="width: 100%;">
        <tr>
            <td valign="top" align="left">
                <dx:ASPxButton ID="btnSubMenu" runat="server" Text="SUB MENU" Font-Names="Verdana"
                    Width="85px" Font-Size="8pt">
                </dx:ASPxButton>
            </td>
            <td valign="top" align="left" style="width: 50px;">
                <dx:ASPxButton ID="btnClear" runat="server" Text="CLEAR" Font-Names="Verdana" Width="90px"
                    AutoPostBack="False" Font-Size="8pt">
                </dx:ASPxButton>
            </td>
            <td valign="top" align="left" style="width: 80px;">
                <dx:ASPxButton ID="btnDelete" runat="server" Text="DELETE" Font-Names="Verdana" Width="80px"
                    AutoPostBack="False" Font-Size="8pt">
                    <ClientSideEvents Click="function(s, e) {
                        up_delete();
                        grid.PerformCallback('load');

                        
                        cboAffiliate2.SetText('');                        
                        txtAffiliate2.SetText('');                                           
                        txtDock.SetText('');

                        HF.Set('hfTest', '');
                                                                    
                        cboAffiliate2.GetInputElement().setAttribute('style', 'background:#FFFFFF;');
                        cboAffiliate2.GetInputElement().readOnly = false;
                        cboAffiliate2.SetEnabled(true);
                    }" />
                </dx:ASPxButton>
            </td>
            <td valign="top" align="right" style="width: 50px;">
                <dx:ASPxButton ID="btnSubmit" runat="server" Text="SAVE" Font-Names="Verdana" Width="90px"
                    AutoPostBack="False" Font-Size="8pt">
                    <ClientSideEvents Click="function(s, e) {
                        validasubmit();
                        up_Insert();
                        
                        cboAffiliate2.SetText('');                        
                        txtAffiliate2.SetText('');                                           
                        txtDock.SetText('');

                        HF.Set('hfTest', '');                                         

                        cboAffiliate2.GetInputElement().setAttribute('style', 'background:#FFFFFF;');
                        cboAffiliate2.GetInputElement().readOnly = false;
                        cboAffiliate2.SetEnabled(true);
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
    <dx:ASPxCallback ID="AffiliateSubmit" runat="server" ClientInstanceName="AffiliateSubmit">
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

                if (s.cpFunction == 'delete'){
                    if (s.cpType != 'error'){
                        clear();
                    }
                }else if(s.cpFunction == 'insert'){
                    clear();
                }
            } else {
                lblInfo.SetText('');
            }  
        }" />
    </dx:ASPxCallback>
    <dx:ASPxHiddenField ID="HF" runat="server" ClientInstanceName="HF">
    </dx:ASPxHiddenField>
</asp:Content>
