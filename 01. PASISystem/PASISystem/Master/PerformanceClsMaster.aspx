<%@ Page Title="" Language="vb" AutoEventWireup="false" MasterPageFile="~/SiteContent.Master"
    CodeBehind="PerformanceClsMaster.aspx.vb" Inherits="PASISystem.PerformanceClsMaster" %>

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
<asp:Content ID="Content1" ContentPlaceHolderID="head" runat="server">
    <style type="text/css">
        .dxeHLC, .dxeHC, .dxeHFC
        {
            display: none;
        }
        #Table1
        {
            width: 986px;
            margin-left: 0px;
        }
        .style45
        {
            width: 178px;
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
        function numbersonly(e) {
            var unicode = e.charCode ? e.charCode : e.keyCode
            if (unicode != 8) { //if the key isn't the backspace key (which we should allow)
                if (unicode < 45 || unicode > 57) //if not a number
                    return false //disable key press
            }
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
            height = height - (height * 45 / 100)
            grid.SetHeight(height);
        }

        function OnBatchEditStartEditing(s, e) {
            currentColumnName = e.focusedColumn.fieldName;

            if (currentColumnName == "NoUrut" || currentColumnName == "PerformanceCls" || currentColumnName == "Description") {

                e.cancel = true;
            }

            currentEditableVisibleIndex = e.visibleIndex;
        }

        function OnGridFocusedRowChanged() {
            grid.GetRowValues(grid.GetFocusedRowIndex(), "PerformanceCls;Description;", OnGetRowValues);
        }

        function OnGetRowValues(values) {
            if (values[0] != "" && values[0] != null && values[0] != "null") {

                txtPerformanceCls.SetText(values[0]);
                txtPerformanceCls2.SetText(values[1]);

                txtPerformanceCls.GetInputElement().setAttribute('style', 'background:#CCCCCC;foreground:#CCCCCC;');
                txtPerformanceCls.GetInputElement().readOnly = true;
                txtPerformanceCls.SetEnabled(false);
                lblInfo.SetText('');
            }
        }

        function up_delete() {
            if (txtPerformanceCls.GetText() == "") {
                lblInfo.GetMainElement().style.color = 'Red';
                lblInfo.SetText("[6011] Please select the data first!");
                e.ProcessOnServer = false;
                return false;
            }
            if (grid.GetFocusedRowIndex() == -1) {
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
            var pPerformanceClsCode = txtPerformanceCls.GetText();
            grid.PerformCallback('delete|' + pPerformanceClsCode);

        }

        function readonly() {
            txtPerformanceCls2.GetInputElement().setAttribute('style', 'background:#CCCCCC;foreground:#CCCCCC;');
            txtPerformanceCls2.GetInputElement().readOnly = true;

        }

        function validasubmit() {
            lblInfo.GetMainElement().style.color = 'Red';
            if (txtPerformanceCls.GetText() == "") {
                lblInfo.SetText("[6011] Please Select Supplier Group Code first!");
                txtPerformanceCls.Focus();
                e.ProcessOnServer = false;
                return false;
            }


            lblInfo.GetMainElement().style.color = 'Red';
            if (txtPerformanceCls2.GetText() == "") {
                lblInfo.SetText("[6011] Please Input the Supplier Group Name first!");
                txtPerformanceCls2.Focus();
                e.ProcessOnServer = false;
                return false;
            }
        }

        function up_Insert() {
            var pIsUpdate = '';
            var pPerformanceClsCode = txtPerformanceCls.GetText();
            var pPerformanceClsName = txtPerformanceCls2.GetText();

            grid.PerformCallback('save|' + pIsUpdate + '|' + pPerformanceClsCode + '|' + pPerformanceClsName);

        }

        function clear() {
            txtPerformanceCls.SetText('');
            txtPerformanceCls2.SetText('');
            txtPerformanceCls.GetInputElement().setAttribute('style', 'background:#FFFFFF;foreground:#FFFFFF;');
            txtPerformanceCls.GetInputElement().readOnly = false;
            txtPerformanceCls.SetEnabled(true);
            lblInfo.SetText('');

        }

        function clear2() {
            txtPerformanceCls.SetText('');
            txtPerformanceCls2.SetText('');
            txtPerformanceCls.GetInputElement().setAttribute('style', 'background:#FFFFFF;foreground:#FFFFFF;');
            txtPerformanceCls.GetInputElement().readOnly = false;
            txtPerformanceCls.SetEnabled(true);

        }
    </script>
</asp:Content>
<asp:Content ID="Content2" ContentPlaceHolderID="Content" runat="server">
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
                            <dx:ASPxLabel ID="lblInfo" runat="server" Font-Names="Tahoma" ClientInstanceName="lblInfo"
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
            <td colspan="2" align="left" valign="top" height="100px">
                <%--Column : Grid--%>
                <dx:ASPxGridView ID="grid" runat="server" Width="100%" Font-Names="Tahoma" KeyFieldName="PerformanceCls"
                    AutoGenerateColumns="False" ClientInstanceName="grid" Font-Size="8pt">
                    <ClientSideEvents Init="OnInit" FocusedRowChanged="function(s, e) {OnGridFocusedRowChanged();
}" 
                        EndCallback="function(s, e) {
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
 
                                    }
                                }else if(s.cpFunction == 'insert'){

                                }
                            } else {
                                lblInfo.SetText('');}  
                        }" 
                        RowClick="function(s, e) {lblInfo.SetText('');}" 
                        BatchEditStartEditing="OnBatchEditStartEditing" />
                    <Columns>
                        <dx:GridViewDataTextColumn VisibleIndex="0" Caption="NO" FieldName="NoUrut" Width="30px"
                            HeaderStyle-HorizontalAlign="Center">
                            <HeaderStyle HorizontalAlign="Center" Font-Names="Tahoma" Font-Size="8pt"></HeaderStyle>
                            <CellStyle Font-Names="Tahoma" Font-Size="8pt">
                            </CellStyle>
                        </dx:GridViewDataTextColumn>
                        <dx:GridViewDataTextColumn VisibleIndex="1" Caption="PERFORMANCE CLS CODE" FieldName="PerformanceCls"
                            Width="170px" HeaderStyle-HorizontalAlign="Center">
                            <HeaderStyle HorizontalAlign="Center" Font-Names="Tahoma" Font-Size="8pt"></HeaderStyle>
                            <CellStyle Font-Names="Tahoma" Font-Size="8pt">
                            </CellStyle>
                        </dx:GridViewDataTextColumn>
                        <dx:GridViewDataTextColumn VisibleIndex="2" Caption="PERFORMANCE CLS NAME" FieldName="Description"
                            Width="310px" HeaderStyle-HorizontalAlign="Center">
                            <HeaderStyle HorizontalAlign="Center" Font-Names="Tahoma" Font-Size="8pt"></HeaderStyle>
                            <CellStyle Font-Names="Tahoma" Font-Size="8pt">
                            </CellStyle>
                        </dx:GridViewDataTextColumn>
                    </Columns>
                    <SettingsBehavior AllowSelectByRowClick="True" AllowSort="False" ColumnResizeMode="Control" />
                    <SettingsPager Visible="False" PageSize="14" NumericButtonCount="10" AlwaysShowPager="True"
                        Mode="ShowAllRecords">
                        <Summary Text="Page {0} of {1} [{2} record(s)]" AllPagesText="Page {0} of {1} " />
                        <Summary AllPagesText="Page {0} of {1} " Text="Page {0} of {1} [{2} record(s)]">
                        </Summary>
                    </SettingsPager>
                    <Settings ShowHorizontalScrollBar="True" ShowVerticalScrollBar="True" ShowGroupButtons="False"
                        ShowStatusBar="Hidden" VerticalScrollableHeight="190" />
                    <SettingsCommandButton EditButton-ButtonType="Link">
                        <EditButton Text="Detail">
                        </EditButton>
                    </SettingsCommandButton>
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
    </table>
    <div style="height: 8px;">
    </div>
    <table style="width: 100%;">
        <tr>
            <td height="50">
                <!-- INPUT AREA -->
                <table id="Table1" style="border-width: 1pt thin thin thin; border-style: ridge;
                    border-color: #9598A1; width: 100%; height: 25px;">
                    <td bgcolor="#FFD2A6" align="center">
                        <dx:ASPxLabel ID="ASPxLabel56" runat="server" Text="PERFORMANCE CLS CODE" Font-Names="Tahoma"
                            Font-Size="8pt" Width="166px">
                        </dx:ASPxLabel>
                    </td>
                    <td bgcolor="#FFD2A6" align="center" width="380px">
                        <dx:ASPxLabel ID="ASPxLabel5" runat="server" Text="PERFORMANCE CLS NAME" Font-Names="Tahoma"
                            Font-Size="8pt" Width="380px">
                        </dx:ASPxLabel>
                    </td>
                    <td align="center" width="380px" colspan="2" style="width: 760px">
                        &nbsp;</td>
                    <tr>
                        <td width="170px">
                            <dx:ASPxTextBox ID="txtPerformanceCls" runat="server" Width="170px" Height="20px"
                                ClientInstanceName="txtPerformanceCls" Font-Names="Tahoma" Font-Size="8pt" MaxLength="3"
                                BackColor="White" TabIndex="1">
                                <ClientSideEvents GotFocus="function(s, e) {lblInfo.SetText('');}" />
                            </dx:ASPxTextBox>
                        </td>
                        <td width="380px">
                            <dx:ASPxTextBox ID="txtPerformanceCls2" runat="server" Width="380px" Height="20px"
                                ClientInstanceName="txtPerformanceCls2" Font-Names="Tahoma" Font-Size="8pt" MaxLength="25"
                                BackColor="White" TabIndex="2">
                                <ClientSideEvents GotFocus="function(s, e) {lblInfo.SetText('');}" />
                            </dx:ASPxTextBox>
                        </td>
                        <td width="380px">
                            &nbsp;</td>
                        <td width="380px">
                            &nbsp;</td>
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
                <dx:ASPxButton ID="btnSubMenu" runat="server" Text="SUB MENU" Font-Names="Tahoma"
                    Width="85px" Font-Size="8pt" TabIndex="6" ClientInstanceName="btnSubMenu">
                </dx:ASPxButton>
            </td>
            <td valign="top" align="right" style="width: 50px;">
                <dx:ASPxButton ID="btnClear" runat="server" Text="CLEAR" Font-Names="Tahoma" Width="90px"
                    AutoPostBack="False" Font-Size="8pt" TabIndex="5" ClientInstanceName="btnClear">
                    <ClientSideEvents Click="function(s, e) {clear();}" />
                </dx:ASPxButton>
            </td>
            <td align="right" style="width: 80px;">
                <dx:ASPxButton ID="btnDelete" runat="server" Text="DELETE" Font-Names="Tahoma" Width="80px"
                    AutoPostBack="False" Font-Size="8pt" TabIndex="4" ClientInstanceName="btnDelete">
                    <ClientSideEvents Click="function(s, e) {
                    up_delete();

                    grid.PerformCallback('load');
                    
                    grid.SetFocusedRowIndex(-1);
                    
                        txtPerformanceCls.SetText('');
                        txtPerformanceCls2.SetText('');
                        
                        txtPerformanceCls.GetInputElement().setAttribute('style', 'background:#FFFFFF;foreground:#FFFFFF;');
                        txtPerformanceCls.GetInputElement().readOnly = false;
                        txtPerformanceCls.SetEnabled(true);

                        grid.SetFocusedRowIndex(-1);
                    }" />
                </dx:ASPxButton>
            </td>
            <td valign="top" align="right" style="width: 50px;">
                <dx:ASPxButton ID="btnSubmit" runat="server" Text="SAVE" Font-Names="Tahoma" Width="90px"
                    AutoPostBack="False" Font-Size="8pt" TabIndex="3" ClientInstanceName="btnSubmit">
                    <ClientSideEvents Click="function(s, e) {
                        grid.SetFocusedRowIndex(-1);
                        validasubmit();
                        up_Insert();
                       
                        grid.PerformCallback('load');
                        
                        grid.SetFocusedRowIndex(-1);

                        txtPerformanceCls.SetText('');
                        txtPerformanceCls2.SetText('');
                        
                        txtPerformanceCls.GetInputElement().setAttribute('style',  'background:#FFFFFF;foreground:#FFFFFF;');
                        txtPerformanceCls.GetInputElement().readOnly = false;
                        txtPerformanceCls.SetEnabled(true);
                        
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
