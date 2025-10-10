<%@ Page Title="" Language="vb" AutoEventWireup="false" MasterPageFile="~/SiteContent.Master"
    CodeBehind="ForecastRegularMonthly.aspx.vb" Inherits="AffiliateSystem.ForecastRegularMonthly" %>

<%@ Register Assembly="DevExpress.Web.v14.1, Version=14.1.6.0, Culture=neutral, PublicKeyToken=b88d1754d700e49a"
    Namespace="DevExpress.Web.ASPxGridView.Export" TagPrefix="dx" %>
<%@ Register Assembly="DevExpress.Web.v14.1, Version=14.1.6.0, Culture=neutral, PublicKeyToken=b88d1754d700e49a"
    Namespace="DevExpress.Web.ASPxGridView" TagPrefix="dx" %>
<%@ Register Assembly="DevExpress.Web.v14.1, Version=14.1.6.0, Culture=neutral, PublicKeyToken=b88d1754d700e49a"
    Namespace="DevExpress.Web.ASPxEditors" TagPrefix="dx" %>
<%@ Register Assembly="DevExpress.Web.v14.1, Version=14.1.6.0, Culture=neutral, PublicKeyToken=b88d1754d700e49a"
    Namespace="DevExpress.Web.ASPxGlobalEvents" TagPrefix="dx" %>
<%@ Register Assembly="DevExpress.Web.v14.1, Version=14.1.6.0, Culture=neutral, PublicKeyToken=b88d1754d700e49a"
    Namespace="DevExpress.Web.ASPxPanel" TagPrefix="dx" %>
<%@ Register Assembly="DevExpress.Web.v14.1, Version=14.1.6.0, Culture=neutral, PublicKeyToken=b88d1754d700e49a"
    Namespace="DevExpress.Web.ASPxRoundPanel" TagPrefix="dx" %>
<%@ Register Assembly="DevExpress.Web.v14.1, Version=14.1.6.0, Culture=neutral, PublicKeyToken=b88d1754d700e49a"
    Namespace="DevExpress.Web.ASPxCallback" TagPrefix="dx" %>
<%@ Register Assembly="DevExpress.Web.v14.1, Version=14.1.6.0, Culture=neutral, PublicKeyToken=b88d1754d700e49a"
    Namespace="DevExpress.Web.ASPxEditors" TagPrefix="dx1" %>
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
            height = height - (height * 45 / 100)
            grid.SetHeight(height);
        }

        function OnBatchEditStartEditing(s, e) {
            currentColumnName = e.focusedColumn.fieldName;

            if (currentColumnName == "NoUrut" || currentColumnName == "PartNo" || currentColumnName == "PartName" || currentColumnName == "UnitDesc"
            || currentColumnName == "MinOrderQty" || currentColumnName == "Maker" || currentColumnName == "KanbanCls" || currentColumnName == "PONo" || currentColumnName == "QtyBox"
            || currentColumnName == "CurrDesc" || currentColumnName == "Price" || currentColumnName == "Amount"
            || currentColumnName == "ForecastN1" || currentColumnName == "ForecastN2" || currentColumnName == "ForecastN3") {
                e.cancel = true;
            }

            if (currentColumnName == "url") {
                var pDeliveryByPASICls = s.batchEditApi.GetCellValue(e.visibleIndex, "DeliveryByPASICls");
                var pSupplierSJNo = s.batchEditApi.GetCellValue(e.visibleIndex, "SupplierSJNo");
                var pPASISJNo = s.batchEditApi.GetCellValue(e.visibleIndex, "PASISJNo");

                if (pDeliveryByPASICls == "1") {
                    if (pPASISJNo == "") {
                        e.cancel = true;
                    }
                } else {
                    if (pSupplierSJNo == "") {
                        e.cancel = true;
                    }
                }
            }

            currentEditableVisibleIndex = e.visibleIndex;
        }

        function OnBatchEditEndEditing(s, e) {
            window.setTimeout(function () {
                var pPrice = s.batchEditApi.GetCellValue(e.visibleIndex, "Price");
                var pQty = s.batchEditApi.GetCellValue(e.visibleIndex, "POQty");

                s.batchEditApi.SetCellValue(e.visibleIndex, "Amount", pPrice * pQty);
            }, 10);
        }

        function validasubmit() {
            lblInfo.GetMainElement().style.color = 'Red';
            if (txtPONo.GetText() == "") {
                lblInfo.SetText("[6011] Please Input PO No. first!");
                txtPONo.Focus();
                e.ProcessOnServer = false;
                return false;
            }

            if (txtShip.GetText() == "") {
                lblInfo.SetText("[6011] Please Input Ship By first!");
                txtShip.Focus();
                e.ProcessOnServer = false;
                return false;
            }
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
        }

        function up_delete() {
            if (txtPONo.GetText() == "") {
                lblInfo.GetMainElement().style.color = 'Red';
                lblInfo.SetText("[6011] Please input PO No first!");
                e.ProcessOnServer = false;
                return false;
            }

            if (txtDate2.GetText() != "") {
                lblInfo.GetMainElement().style.color = 'Red';
                lblInfo.SetText("[6011] Can't delete, because this PO already Approve!");
                e.ProcessOnServer = false;
                return false;
            }

            var msg = confirm('Are you sure want to delete this data ?');
            if (msg == false) {
                e.processOnServer = false;
                return;
            }

            var pGroupCode = txtPONo.GetText();
            ButtonDelete.PerformCallback('delete|' + pGroupCode);
        }
    </script>
</asp:Content>
<asp:Content ID="Content2" ContentPlaceHolderID="Content" runat="server">
    <table style="width: 100%;">
        <tr>
            <td width="100%">
                <table style="border-width: 1pt thin thin thin; border-style: ridge; border-color: #9598A1;
                    width: 100%; height: 70px;">
                    <tr>
                        <td colspan="8" height="20">
                            <table id="Table1">
                                <!-- ROW 1 -->
                                <tr>
                                    <td style="width: 5px;">
                                    </td>
                                    <td align="left" valign="middle" height="20px" width="200px">
                                        <dx:ASPxLabel ID="ASPxLabel4" runat="server" Text="PERIOD FORECAST" Font-Names="Tahoma"
                                            Font-Size="8pt">
                                        </dx:ASPxLabel>
                                    </td>
                                    <td align="left" valign="middle" height="20px" width="280px">
                                        <table>
                                            <tr>
                                                <td>
                                                    <dx:ASPxTimeEdit ID="dtPOPeriod" runat="server" ClientInstanceName="dtPOPeriod"
                                                        DisplayFormatString="yyyy" EditFormat="Custom" EditFormatString="yyyy"
                                                        Width="100px" HorizontalAlign="Left">
                                                        <ClientSideEvents DateChanged="function(s, e) {
	cbopart.PerformCallback(dtPOPeriod.GetText());
}" />
                                                    </dx:ASPxTimeEdit>
                                                </td>
                                                <td>
                                                   
                                                </td>
                                            </tr>
                                        </table>
                                    </td>
                                    <td style="width: 5px;">
                                    </td>
                                    <td>
                                    </td>
                                    <td>
                                    </td>
                                    <td style="width: 5px;">
                                    </td>
                                    <td align="left" valign="middle" height="20px" width="130px">
                                    </td>
                                    <td align="right" valign="middle" height="20px" width="150px">
                                    </td>
                                </tr>                                                                  
                                <!-- ROW 2 -->
                                <tr>
                                    <td style="width: 5px;">
                                    </td>
                                    <td align="left" valign="middle">
                                        <dx:ASPxLabel ID="ASPxLabel3" runat="server" Text="PART NO" Font-Names="Tahoma"
                                            Font-Size="8pt">
                                        </dx:ASPxLabel>
                                    </td>
                                    <td align="left" valign="middle">
                                        <table>
                                            <tr>
                                                <td>
                                                    <dx1:ASPxComboBox ID="cbopart" runat="server" ClientInstanceName="cbopart" Font-Names="Tahoma"
                                                        TextFormatString="{0}" Font-Size="8pt">
                                                        <ClientSideEvents SelectedIndexChanged="function(s, e) {
	                                                                                txtpart.SetText(cbopart.GetSelectedItem().GetColumnText(1));
                                                                                    }" />
                                                    </dx1:ASPxComboBox>
                                                </td>  
                                                <td>
                                                    <dx1:ASPxTextBox ID="txtpart" runat="server" Font-Names="Tahoma" Font-Size="8pt"
                                                        Width="170px" BackColor="Silver" ClientInstanceName="txtpart" ReadOnly="True">
                                                        <ClientSideEvents TextChanged="function(s, e) {
	                                    lblerrmessage.SetText('');
                                    }" />
                                                    </dx1:ASPxTextBox>
                                                </td>                                              
                                            </tr>
                                        </table>
                                    </td>
                                    <td style="width: 5px;">                                        
                                    </td>
                                    <td>
                                        
                                    </td>
                                    <td>
                                        
                                    </td>
                                    <td style="width: 5px;">
                                    </td>
                                    <td align="left" valign="middle" height="20px" width="130px">
                                        
                                    </td>
                                    <td align="left" valign="middle">
                                        <table>
                                            <tr>
                                                <td>
                                                    <dx:ASPxButton ID="btnSearch" runat="server" Text="SEARCH" Font-Names="Tahoma" Width="85px"
                                                        AutoPostBack="False" Font-Size="8pt">
                                                        <ClientSideEvents Click="function(s, e) {                                         
                                                            grid.PerformCallback('load');
                                                            lblInfo.SetText('');
                                                        }" />
                                                    </dx:ASPxButton>
                                                </td>
                                                <td>
                                                    <dx:ASPxButton ID="btnClear" runat="server" Text="CLEAR" Font-Names="Tahoma" Width="85px"
                                                        AutoPostBack="False" Font-Size="8pt">
                                                        <ClientSideEvents Click="function(s, e) {
                                                            
                                                            cboAffiliateCode.SetText('==ALL==');

                                                            cboSupplierCode.SetText('==ALL==');

                                                            txtPartNo.SetText('');

                                                            lblInfo.SetText('');
                                                            grid.PerformCallback('clear');
                                                        }" />
                                                    </dx:ASPxButton>
                                                </td>
                                            </tr>
                                        </table>                                    
                                    </td>
                                </tr>                               
                                <!-- ROW 3 -->
                                <tr>
                                    <td style="width: 5px;">
                                    </td>
                                    <td align="left" valign="middle">
                                       
                                    </td>
                                    <td align="left" valign="middle">
                                        
                                    </td>
                                    <td style="width: 5px;">
                                    </td>
                                    <td>
                                        
                                    </td>
                                    <td>
                                        
                                    </td>
                                    <td style="width: 5px;">
                                    </td>
                                    <td align="left" valign="middle">
                                        
                                    </td>
                                    <td align="left" valign="middle">
                                       
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
            <td align="right">
                &nbsp
            </td>
            <td align="right">
            </td>
            <td align="right">
                <dx:ASPxImage ID="ASPxImage1" runat="server" ShowLoadingImage="true" ImageUrl="~/Images/fuchsia.jpg"
                    Height="15px" Width="15px">
                </dx:ASPxImage>
                <dx:ASPxLabel ID="ASPxLabel5" runat="server" Text=" : DIFFERENCE" Font-Names="Tahoma"
                    ClientInstanceName="difference" Font-Bold="True" Font-Size="8pt">
                </dx:ASPxLabel>
            </td>
        </tr>
        <tr>
            <td colspan="3" align="left" valign="top" height="220">
                <%--Column : Grid--%>
                <dx:ASPxGridView ID="grid" runat="server" Width="100%" Font-Names="Tahoma" KeyFieldName="colNo"
                    AutoGenerateColumns="False" ClientInstanceName="grid" Font-Size="8pt">
                    <ClientSideEvents Init="OnInit" FocusedRowChanged="function(s, e) {
	                    OnGridFocusedRowChanged();
                    }" CallbackError="function(s, e) {e.handled = true;}" BatchEditEndEditing="OnBatchEditEndEditing"
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
                        delete s.cpPONo;
                    }" RowClick="function(s, e) {
	                    lblInfo.SetText('');
                    }" BatchEditStartEditing="OnBatchEditStartEditing" />
                    <Columns>
                        <%--<dx:GridViewDataTextColumn VisibleIndex="0" Caption="NO" FieldName="ColNo" Width="0px"
                            HeaderStyle-HorizontalAlign="Center">
                            <HeaderStyle HorizontalAlign="Center" Wrap="True"></HeaderStyle>
                            <CellStyle Font-Names="Tahoma" Font-Size="8pt" HorizontalAlign="Center">
                            </CellStyle>
                        </dx:GridViewDataTextColumn>--%>
                        <dx:GridViewDataTextColumn VisibleIndex="1" Caption="PART NO." FieldName="PartNo"
                            Width="100px" HeaderStyle-HorizontalAlign="Center">
                            <HeaderStyle HorizontalAlign="Center"></HeaderStyle>
                            <CellStyle Font-Names="Tahoma" Font-Size="8pt">
                            </CellStyle>
                        </dx:GridViewDataTextColumn>
                        <dx:GridViewDataTextColumn VisibleIndex="2" Caption="AFFILIATE CODE" FieldName="AffiliateID"
                            Width="120px" HeaderStyle-HorizontalAlign="Center">
                            <HeaderStyle HorizontalAlign="Center" Wrap="True"></HeaderStyle>
                            <CellStyle Font-Names="Tahoma" Font-Size="8pt">
                            </CellStyle>
                        </dx:GridViewDataTextColumn>                    
                        <dx:GridViewDataTextColumn VisibleIndex="3" Caption="SUPPLIER CODE" FieldName="SupplierID"
                            Width="90px" HeaderStyle-HorizontalAlign="Center">
                            <HeaderStyle HorizontalAlign="Center" Wrap="True"></HeaderStyle>
                            <CellStyle Font-Names="Tahoma" Font-Size="8pt">
                            </CellStyle>
                        </dx:GridViewDataTextColumn>                        
                        <dx:GridViewDataTextColumn Caption="DATA" FieldName="Data" VisibleIndex="4" Width="110px"
                            HeaderStyle-HorizontalAlign="Center">
                            <HeaderStyle HorizontalAlign="Center"></HeaderStyle>
                            <CellStyle Font-Names="Tahoma" Font-Size="8pt" HorizontalAlign="Right">
                            </CellStyle>
                        </dx:GridViewDataTextColumn>                        
                        <dx:GridViewDataTextColumn Caption="7" FieldName="F7" VisibleIndex="5" Width="60px" HeaderStyle-HorizontalAlign="Center">
                            <PropertiesTextEdit DisplayFormatString="{0:n0}">
                            </PropertiesTextEdit>

<HeaderStyle HorizontalAlign="Center"></HeaderStyle>

                            <CellStyle HorizontalAlign="Right" VerticalAlign="Middle" Wrap="True">
                            </CellStyle>
                        </dx:GridViewDataTextColumn>
                        <dx:GridViewDataTextColumn Caption="8" FieldName="F8" VisibleIndex="6" Width="60px" HeaderStyle-HorizontalAlign="Center">
                            <PropertiesTextEdit DisplayFormatString="{0:n0}">
                            </PropertiesTextEdit>

<HeaderStyle HorizontalAlign="Center"></HeaderStyle>

                            <CellStyle HorizontalAlign="Right" VerticalAlign="Middle" Wrap="True">
                            </CellStyle>
                        </dx:GridViewDataTextColumn>
                        <dx:GridViewDataTextColumn Caption="9" FieldName="F9" VisibleIndex="7" Width="60px" HeaderStyle-HorizontalAlign="Center">
                            <PropertiesTextEdit DisplayFormatString="{0:n0}">
                            </PropertiesTextEdit>

<HeaderStyle HorizontalAlign="Center"></HeaderStyle>

                            <CellStyle HorizontalAlign="Right" VerticalAlign="Middle" Wrap="True">
                            </CellStyle>
                        </dx:GridViewDataTextColumn>
                        <dx:GridViewDataTextColumn Caption="10" FieldName="F10" VisibleIndex="8" Width="60px" HeaderStyle-HorizontalAlign="Center">
                            <PropertiesTextEdit DisplayFormatString="{0:n0}">
                            </PropertiesTextEdit>

<HeaderStyle HorizontalAlign="Center"></HeaderStyle>

                            <CellStyle HorizontalAlign="Right" VerticalAlign="Middle" Wrap="True">
                            </CellStyle>
                        </dx:GridViewDataTextColumn>
                        <dx:GridViewDataTextColumn Caption="11" FieldName="F11" VisibleIndex="9" Width="60px" HeaderStyle-HorizontalAlign="Center">
                            <PropertiesTextEdit DisplayFormatString="{0:n0}">
                            </PropertiesTextEdit>

<HeaderStyle HorizontalAlign="Center"></HeaderStyle>

                            <CellStyle HorizontalAlign="Right" VerticalAlign="Middle" Wrap="True">
                            </CellStyle>
                        </dx:GridViewDataTextColumn>
                        <dx:GridViewDataTextColumn Caption="12" FieldName="F12" VisibleIndex="10" Width="60px" HeaderStyle-HorizontalAlign="Center">
                            <PropertiesTextEdit DisplayFormatString="{0:n0}">
                            </PropertiesTextEdit>

<HeaderStyle HorizontalAlign="Center"></HeaderStyle>

                            <CellStyle HorizontalAlign="Right" VerticalAlign="Middle" Wrap="True">
                            </CellStyle>
                        </dx:GridViewDataTextColumn>    
                        <dx:GridViewDataTextColumn Caption="1" FieldName="F1" VisibleIndex="11" Width="60px" HeaderStyle-HorizontalAlign="Center">
                            <PropertiesTextEdit DisplayFormatString="{0:n0}">
                            </PropertiesTextEdit>

<HeaderStyle HorizontalAlign="Center"></HeaderStyle>

                            <CellStyle HorizontalAlign="Right" VerticalAlign="Middle" Wrap="True">
                            </CellStyle>
                        </dx:GridViewDataTextColumn>
                        <dx:GridViewDataTextColumn Caption="2" FieldName="F2" VisibleIndex="12" Width="60px" HeaderStyle-HorizontalAlign="Center">
                            <PropertiesTextEdit DisplayFormatString="{0:n0}">
                            </PropertiesTextEdit>

<HeaderStyle HorizontalAlign="Center"></HeaderStyle>

                            <CellStyle HorizontalAlign="Right" VerticalAlign="Middle" Wrap="True">
                            </CellStyle>
                        </dx:GridViewDataTextColumn>
                        <dx:GridViewDataTextColumn Caption="3" FieldName="F3" VisibleIndex="13" Width="60px" HeaderStyle-HorizontalAlign="Center">
                            <PropertiesTextEdit DisplayFormatString="{0:n0}">
                            </PropertiesTextEdit>

<HeaderStyle HorizontalAlign="Center"></HeaderStyle>

                            <CellStyle HorizontalAlign="Right" VerticalAlign="Middle" Wrap="True">
                            </CellStyle>
                        </dx:GridViewDataTextColumn>
                        <dx:GridViewDataTextColumn Caption="4" FieldName="F4" VisibleIndex="14" Width="60px" HeaderStyle-HorizontalAlign="Center">
                            <PropertiesTextEdit DisplayFormatString="{0:n0}">
                            </PropertiesTextEdit>

<HeaderStyle HorizontalAlign="Center"></HeaderStyle>

                            <CellStyle HorizontalAlign="Right" VerticalAlign="Middle" Wrap="True">
                            </CellStyle>
                        </dx:GridViewDataTextColumn>
                        <dx:GridViewDataTextColumn Caption="5" FieldName="F5" VisibleIndex="15" Width="60px" HeaderStyle-HorizontalAlign="Center">
                            <PropertiesTextEdit DisplayFormatString="{0:n0}">
                            </PropertiesTextEdit>

<HeaderStyle HorizontalAlign="Center"></HeaderStyle>

                            <CellStyle HorizontalAlign="Right" VerticalAlign="Middle" Wrap="True">
                            </CellStyle>
                        </dx:GridViewDataTextColumn>
                        <dx:GridViewDataTextColumn Caption="6" FieldName="F6" VisibleIndex="16" Width="60px" HeaderStyle-HorizontalAlign="Center">
                            <PropertiesTextEdit DisplayFormatString="{0:n0}">
                            </PropertiesTextEdit>

<HeaderStyle HorizontalAlign="Center"></HeaderStyle>

                            <CellStyle HorizontalAlign="Right" VerticalAlign="Middle" Wrap="True">
                            </CellStyle>
                        </dx:GridViewDataTextColumn>                    
                    </Columns>
                    <SettingsPager PageSize="100" Position="Top">
                        <Summary Text="Page {0} of {1} [{2} record(s)]" AllPagesText="Page {0} of {1} " />
                        <Summary AllPagesText="Page {0} of {1} " Text="Page {0} of {1} [{2} record(s)]">
                        </Summary>
                    </SettingsPager>
                    <Settings ShowHorizontalScrollBar="True" ShowVerticalScrollBar="True" ShowGroupButtons="False"
                        ShowStatusBar="Hidden" VerticalScrollableHeight="190" />
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
    <div style="height: 1px;">
    </div>
    <table id="button" style="width: 100%;">
        <tr>
            <td valign="top" align="left">
                <dx:ASPxButton ID="btnSubMenu" runat="server" Text="SUB MENU" Font-Names="Tahoma"
                    Width="85px" Font-Size="8pt">
                </dx:ASPxButton>
            </td>
            <td valign="top" align="right" style="width: 50px;">
            </td>
            <td valign="top" align="right" style="width: 50px;">
            </td>
            <td valign="top" align="right" style="width: 50px;">
            </td>
            <td valign="top" align="right" style="width: 50px;">
            </td>
            <td valign="top" align="right" style="width: 50px;">
            </td>
            <td valign="top" align="right" style="width: 50px;">
            </td>
            <td align="right" style="width: 80px;">
                &nbsp;</td>
            <td valign="top" align="right" style="width: 50px;">
                <dx:ASPxButton ID="btnexcel" runat="server" Text="EXCEL" Font-Names="Tahoma" Width="85px"
                    AutoPostBack="False" Font-Size="8pt" ClientInstanceName="btnexcel">
                    <ClientSideEvents Click="function(s, e) {                                         
                                                            grid.PerformCallback('excel');
                                                            lblInfo.SetText('');
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
