<%@ Page Title="" Language="vb" AutoEventWireup="false" MasterPageFile="~/SiteContent.Master"
    CodeBehind="SummaryInvoiceAffiliateExp.aspx.vb" Inherits="PASISystem.SummaryInvoiceAffiliateExp" %>

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
            height = height - (height * 50 / 100)
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
                                        <dx:ASPxLabel ID="ASPxLabel4" runat="server" Text="AFFILIATE CODE/NAME" Font-Names="Tahoma"
                                            Font-Size="8pt">
                                        </dx:ASPxLabel>
                                    </td>
                                    <td align="left" valign="middle" height="20px" width="280px">
                                        <table>
                                            <tr>
                                                <td>
                                                    <dx:ASPxComboBox ID="cboAffiliateCode" runat="server" ClientInstanceName="cboAffiliateCode"
                                                        Font-Names="Tahoma" TextFormatString="{0}" Font-Size="8pt" Width="120px" 
                                                        DropDownStyle="DropDown">
                                                        <ClientSideEvents SelectedIndexChanged="function(s, e) {
	                                                        txtAffiliateName.SetText(cboAffiliateCode.GetSelectedItem().GetColumnText(1));
                                                            grid.PerformCallback('clear');
                                                            lblInfo.SetText('');
                                                            }" />
                                                    </dx:ASPxComboBox>
                                                </td>
                                                <td>
                                                    <dx:ASPxTextBox ID="txtAffiliateName" runat="server" Width="200px" Font-Names="Tahoma"
                                                        Font-Size="8pt" BackColor="#CCCCCC" ReadOnly="True" MaxLength="20" ClientInstanceName="txtAffiliateName">
                                                    </dx:ASPxTextBox>
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
                                    <td align="left" valign="middle" height="20px" width="200px">
                                        <dx:ASPxLabel ID="ASPxLabel2" runat="server" Text="SUPPLIER CODE" Font-Names="Tahoma"
                                            Font-Size="8pt">
                                        </dx:ASPxLabel>
                                    </td>
                                    <td align="left" valign="middle" height="20px" width="280px">
                                        <table>
                                            <tr>
                                                <td>
                                                    <dx:ASPxComboBox ID="cboSupplierCode" runat="server" ClientInstanceName="cboSupplierCode"
                                                        Font-Names="Tahoma" TextFormatString="{0}" Font-Size="8pt" Width="120px" 
                                                        DropDownStyle="DropDown">
                                                        <ClientSideEvents SelectedIndexChanged="function(s, e) {
	                                                        txtSupplierName.SetText(cboSupplierCode.GetSelectedItem().GetColumnText(1));
                                                            grid.PerformCallback('clear');
                                                            lblInfo.SetText('');
                                                            }" />
                                                    </dx:ASPxComboBox>
                                                </td>
                                                <td>
                                                    <dx:ASPxTextBox ID="txtSupplierName" runat="server" Width="200px" Font-Names="Tahoma"
                                                        Font-Size="8pt" BackColor="#CCCCCC" ReadOnly="True" MaxLength="20" ClientInstanceName="txtSupplierName">
                                                    </dx:ASPxTextBox>
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
                                <!-- ROW 3 -->
                                <tr>
                                    <td style="width: 5px;">
                                    </td>
                                    <td align="left" valign="middle">
                                        <dx:ASPxLabel ID="ASPxLabel1" runat="server" Text="AFFILIATE PO PERIOD" Font-Names="Tahoma"
                                            Font-Size="8pt">
                                        </dx:ASPxLabel>
                                    </td>
                                    <td align="left" valign="middle">
                                        <table>
                                            <tr>
                                                <td>
                                                    <dx:ASPxCheckBox ID="chkPOPeriod" runat="server" ClientInstanceName="chkPOPeriod"
                                                        Text=" " Checked="True" CheckState="Checked">
                                                        <ClientSideEvents CheckedChanged="function (s, e) {
                                                                if (s.GetChecked()==true) {
                                                                    dtPOPeriodFrom.SetEnabled(true);
                                                                    dtPOPeriodTo.SetEnabled(true);
                                                                } else {
                                                                    dtPOPeriodFrom.SetEnabled(false);
                                                                    dtPOPeriodTo.SetEnabled(false);
                                                                }
                                                                grid.PerformCallback('clear');
                                                          }" />
                                                    </dx:ASPxCheckBox>
                                                </td>
                                                <td>
                                                    <%--<dx:ASPxDateEdit ID="dtPOPeriodFrom" runat="server" Font-Names="Tahoma" Font-Size="8pt" Width="100px" HorizontalAlign="Center"
                                                        EditFormat="Custom" EditFormatString="MMM yyyy" ClientInstanceName="dtPOPeriodFrom">
                                                    </dx:ASPxDateEdit>--%>
                                                    <dx:ASPxTimeEdit ID="dtPOPeriodFrom" runat="server" ClientInstanceName="dtPOPeriodFrom"
                                                        DisplayFormatString="MMM yyyy" EditFormat="Custom" EditFormatString="MMM yyyy"
                                                        Width="100px" HorizontalAlign="Center">
                                                    </dx:ASPxTimeEdit>
                                                </td>
                                                <td>
                                                    ~
                                                </td>
                                                <td>
                                                    <%--<dx:ASPxDateEdit ID="dtPOPeriodTo" runat="server" Font-Names="Tahoma" Font-Size="8pt" Width="100px" HorizontalAlign="Center"
                                                        EditFormat="Custom" EditFormatString="MMM yyyy" ClientInstanceName="dtPOPeriodTo">
                                                    </dx:ASPxDateEdit>--%>
                                                    <dx:ASPxTimeEdit ID="dtPOPeriodTo" runat="server" ClientInstanceName="dtPOPeriodTo"
                                                        DisplayFormatString="MMM yyyy" EditFormat="Custom" EditFormatString="MMM yyyy"
                                                        Width="100px" HorizontalAlign="Center">
                                                    </dx:ASPxTimeEdit>
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
                                    <td align="left" valign="middle">
                                        
                                    </td>
                                    <td align="left" valign="middle">
                                       
                                    </td>
                                </tr>
                                <!-- ROW 4 -->
                                <tr>
                                    <td style="width: 5px;">
                                    </td>
                                    <td align="left" valign="middle">
                                        <dx:ASPxLabel ID="ASPxLabel8" runat="server" Text="SUPPLIER DELIVERY DATE" Font-Names="Tahoma"
                                            Font-Size="8pt">
                                        </dx:ASPxLabel>
                                    </td>
                                    <td align="left" valign="middle">
                                        <table>
                                            <tr>
                                                <td>
                                                    <dx:ASPxCheckBox ID="chkSupplierDelDate" runat="server" ClientInstanceName="chkSupplierDelDate"
                                                        Text=" " Checked="true">
                                                        <ClientSideEvents CheckedChanged="function (s, e) {
                                                                if (s.GetChecked()==true) {
                                                                    dtSupplierDelDateFrom.SetEnabled(true);
                                                                    dtSupplierDelDateTo.SetEnabled(true);
                                                                } else {
                                                                    dtSupplierDelDateFrom.SetEnabled(false);
                                                                    dtSupplierDelDateTo.SetEnabled(false);
                                                                }
                                                                grid.PerformCallback('clear');
                                                          }" />
                                                    </dx:ASPxCheckBox>
                                                </td>
                                                <td>
                                                    <dx:ASPxDateEdit ID="dtSupplierDelDateFrom" runat="server" Font-Names="Tahoma" Font-Size="8pt"
                                                        Width="100px" EditFormat="Custom" EditFormatString="dd MMM yyyy" ClientInstanceName="dtSupplierDelDateFrom">
                                                    </dx:ASPxDateEdit>
                                                </td>
                                                <td>
                                                    ~
                                                </td>
                                                <td>
                                                    <dx:ASPxDateEdit ID="dtSupplierDelDateTo" runat="server" Font-Names="Tahoma" Font-Size="8pt"
                                                        Width="100px" EditFormat="Custom" EditFormatString="dd MMM yyyy" ClientInstanceName="dtSupplierDelDateTo">
                                                    </dx:ASPxDateEdit>
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
                                        
                                    </td>
                                </tr>                                
                                <!-- ROW 6 -->
                                <tr>
                                    <td style="width: 5px;">
                                    </td>
                                    <td align="left" valign="middle">
                                        <dx:ASPxLabel ID="ASPxLabel14" runat="server" Text="PO No" Font-Names="Tahoma"
                                            Font-Size="8pt">
                                        </dx:ASPxLabel>
                                    </td>
                                    <td align="left" valign="middle">
                                        <table>
                                            <tr>
                                                <td>
                                                    <dx:ASPxTextBox ID="txtPONo" runat="server" Width="200px" Font-Names="Tahoma"
                                                        Font-Size="8pt" MaxLength="20" ClientInstanceName="txtPONo">
                                                    </dx:ASPxTextBox>
                                                </td>
                                                <td>
                                                   
                                                </td>
                                                <td>
                                                    
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
                                                            chkPOPeriod.SetChecked(true);
                                                            dtPOPeriodFrom.SetEnabled(true);
                                                            dtPOPeriodTo.SetEnabled(true);

                                                            chkSupplierDelDate.SetChecked(false);
                                                            dtSupplierDelDateFrom.SetEnabled(false);
                                                            dtSupplierDelDateTo.SetEnabled(false);

                                                            cboAffiliateCode.SetText('==ALL==');
                                                            txtAffiliateName.SetText('==ALL==');

                                                            cboSupplierCode.SetText('==ALL==');
                                                            txtSupplierName.SetText('==ALL==');
                                                            txtPONo.SetText('');

                                                            lblInfo.SetText('');
                                                            grid.PerformCallback('clear');
                                                        }" />
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
                        <dx:GridViewDataTextColumn VisibleIndex="1" Caption="PERIOD" FieldName="Period" Width="70px"
                            HeaderStyle-HorizontalAlign="Center">
                            <HeaderStyle HorizontalAlign="Center" Wrap="True"></HeaderStyle>
                            <PropertiesTextEdit DisplayFormatString="MMM yyyy">
                            </PropertiesTextEdit>
                            <CellStyle Font-Names="Tahoma" Font-Size="8pt" HorizontalAlign="Center">
                            </CellStyle>
                        </dx:GridViewDataTextColumn>
                        <dx:GridViewDataTextColumn VisibleIndex="2" Caption="PO NO." FieldName="PONo" Width="100px"
                            HeaderStyle-HorizontalAlign="Center">
                            <HeaderStyle HorizontalAlign="Center" Wrap="True"></HeaderStyle>
                            <CellStyle Font-Names="Tahoma" Font-Size="8pt">
                            </CellStyle>
                        </dx:GridViewDataTextColumn>
                        <dx:GridViewDataTextColumn VisibleIndex="3" Caption="AFFILIATE CODE" FieldName="AffiliateID"
                            Width="120px" HeaderStyle-HorizontalAlign="Center">
                            <HeaderStyle HorizontalAlign="Center" Wrap="True"></HeaderStyle>
                            <CellStyle Font-Names="Tahoma" Font-Size="8pt">
                            </CellStyle>
                        </dx:GridViewDataTextColumn>                    
                        <dx:GridViewDataTextColumn VisibleIndex="5" Caption="SUPPLIER CODE" FieldName="SupplierID"
                            Width="70px" HeaderStyle-HorizontalAlign="Center">
                            <HeaderStyle HorizontalAlign="Center" Wrap="True"></HeaderStyle>
                            <CellStyle Font-Names="Tahoma" Font-Size="8pt">
                            </CellStyle>
                        </dx:GridViewDataTextColumn>                        
                        <%--<dx:GridViewDataTextColumn VisibleIndex="6" Caption="PO KANBAN" FieldName="POKanban"
                            Width="0px" HeaderStyle-HorizontalAlign="Center">
                            <HeaderStyle HorizontalAlign="Center" Wrap="True"></HeaderStyle>
                            <CellStyle Font-Names="Tahoma" Font-Size="8pt" HorizontalAlign="Center">
                            </CellStyle>
                        </dx:GridViewDataTextColumn>--%>
                        <%--<dx:GridViewDataTextColumn VisibleIndex="7" Caption="PO ISSUE DATE AFFILIATE"
                            FieldName="EntryDate" Width="110px" 
                            HeaderStyle-HorizontalAlign="Center">
                            <HeaderStyle HorizontalAlign="Center" Wrap="True"></HeaderStyle>
                            <PropertiesTextEdit DisplayFormatString="dd MMM yyyy">
                            </PropertiesTextEdit>
                            <CellStyle Font-Names="Tahoma" Font-Size="8pt" HorizontalAlign="Center">
                            </CellStyle>
                        </dx:GridViewDataTextColumn>--%>
                        <dx:GridViewDataTextColumn VisibleIndex="8" Caption="PO ISSUE DATE AFFILIATE"
                            FieldName="POIssueDateAffiliate" Width="110px" 
                            HeaderStyle-HorizontalAlign="Center">
                            <HeaderStyle HorizontalAlign="Center" Wrap="True"></HeaderStyle>
                            <PropertiesTextEdit DisplayFormatString="dd MMM yyyy">
                            </PropertiesTextEdit>
                            <CellStyle Font-Names="Tahoma" Font-Size="8pt" HorizontalAlign="Center">
                            </CellStyle>
                        </dx:GridViewDataTextColumn>
                        <dx:GridViewDataTextColumn VisibleIndex="9" Caption="PART NO." FieldName="PartNo"
                            Width="100px" HeaderStyle-HorizontalAlign="Center">
                            <HeaderStyle HorizontalAlign="Center"></HeaderStyle>
                            <CellStyle Font-Names="Tahoma" Font-Size="8pt">
                            </CellStyle>
                        </dx:GridViewDataTextColumn>
                        <dx:GridViewDataTextColumn VisibleIndex="10" Caption="PART NAME" FieldName="PartName"
                            Width="150px" HeaderStyle-HorizontalAlign="Center">
                            <HeaderStyle HorizontalAlign="Center" Wrap="True"></HeaderStyle>
                            <CellStyle Font-Names="Tahoma" Font-Size="8pt">
                            </CellStyle>
                        </dx:GridViewDataTextColumn>
                        <dx:GridViewDataTextColumn Caption="PO QTY" FieldName="QtyPO" VisibleIndex="11" Width="90px"
                            HeaderStyle-HorizontalAlign="Center">
                            <PropertiesTextEdit DisplayFormatString="{0:n0}">
                            </PropertiesTextEdit>
                            <HeaderStyle HorizontalAlign="Center"></HeaderStyle>
                            <CellStyle Font-Names="Tahoma" Font-Size="8pt" HorizontalAlign="Right">
                            </CellStyle>
                        </dx:GridViewDataTextColumn>
                        <%--<dx:GridViewDataTextColumn VisibleIndex="12" Caption="PCS/BOX" FieldName="QtyBox"
                            Width="80px" HeaderStyle-HorizontalAlign="Center">
                            <PropertiesTextEdit DisplayFormatString="{0:n0}">
                            </PropertiesTextEdit>
                            <HeaderStyle HorizontalAlign="Center" Wrap="True"></HeaderStyle>
                            <CellStyle Font-Names="Tahoma" Font-Size="8pt">
                            </CellStyle>
                        </dx:GridViewDataTextColumn>
                        <dx:GridViewDataTextColumn Caption="BOX/PALLET" FieldName="BoxPallet"
                            VisibleIndex="13" Width="90px" HeaderStyle-HorizontalAlign="Center">
                            <PropertiesTextEdit DisplayFormatString="{0:n0}">
                            </PropertiesTextEdit>
                            <HeaderStyle HorizontalAlign="Center"></HeaderStyle>
                            <CellStyle Font-Names="Tahoma" Font-Size="8pt" HorizontalAlign="Right">
                            </CellStyle>
                        </dx:GridViewDataTextColumn>
                        <dx:GridViewDataTextColumn Caption="VOLUME (PALLET)" FieldName="VolumePallet"
                            VisibleIndex="14" Width="90px" HeaderStyle-HorizontalAlign="Center">
                            <PropertiesTextEdit DisplayFormatString="{0:n0}">
                            </PropertiesTextEdit>
                            <HeaderStyle HorizontalAlign="Center"></HeaderStyle>
                            <CellStyle Font-Names="Tahoma" Font-Size="8pt" HorizontalAlign="Right">
                            </CellStyle>
                        </dx:GridViewDataTextColumn>--%>  
                        <dx:GridViewDataTextColumn VisibleIndex="15" Caption="SCHEDULE ETD SUPPLIER"
                            FieldName="ETDSupp" Width="110px" 
                            HeaderStyle-HorizontalAlign="Center">
                            <HeaderStyle HorizontalAlign="Center" Wrap="True"></HeaderStyle>
                            <PropertiesTextEdit DisplayFormatString="dd MMM yyyy">
                            </PropertiesTextEdit>
                            <CellStyle Font-Names="Tahoma" Font-Size="8pt" HorizontalAlign="Center">
                            </CellStyle>
                        </dx:GridViewDataTextColumn>
                        <dx:GridViewDataTextColumn VisibleIndex="16" Caption="SCHEDULE ETA AFFILIATE"
                            FieldName="ETAAff" Width="110px" 
                            HeaderStyle-HorizontalAlign="Center">
                            <HeaderStyle HorizontalAlign="Center" Wrap="True"></HeaderStyle>
                            <PropertiesTextEdit DisplayFormatString="dd MMM yyyy">
                            </PropertiesTextEdit>
                            <CellStyle Font-Names="Tahoma" Font-Size="8pt" HorizontalAlign="Center">
                            </CellStyle>
                        </dx:GridViewDataTextColumn>                    
                        <dx:GridViewDataTextColumn VisibleIndex="17" Caption="ACTUAL ETD SUPPLIER" FieldName="ActualETDSupplier"
                            Width="110px" HeaderStyle-HorizontalAlign="Center">
                            <HeaderStyle HorizontalAlign="Center" Wrap="True"></HeaderStyle>
                            <PropertiesTextEdit DisplayFormatString="dd MMM yyyy">
                            </PropertiesTextEdit>
                            <CellStyle Font-Names="Tahoma" Font-Size="8pt" HorizontalAlign="Center">
                            </CellStyle>
                        </dx:GridViewDataTextColumn>
                        <dx:GridViewDataTextColumn VisibleIndex="18" Caption="SURAT JALAN NO." FieldName="SupplierDNNo"
                            Width="100px" HeaderStyle-HorizontalAlign="Center">
                            <HeaderStyle HorizontalAlign="Center" Wrap="True"></HeaderStyle>
                            <CellStyle Font-Names="Tahoma" Font-Size="8pt" HorizontalAlign="Left">
                            </CellStyle>
                        </dx:GridViewDataTextColumn>
                        <%--<dx:GridViewDataTextColumn VisibleIndex="19" Caption="PASI INVOICE NO." FieldName="PASIInvoiceNo"
                            Width="100px" HeaderStyle-HorizontalAlign="Center">
                            <HeaderStyle HorizontalAlign="Center" Wrap="True"></HeaderStyle>
                            <CellStyle Font-Names="Tahoma" Font-Size="8pt" HorizontalAlign="Left">
                            </CellStyle>
                        </dx:GridViewDataTextColumn>--%>
                        <%--<dx:GridViewDataTextColumn VisibleIndex="20" Caption="PASI DELIVERY QTY" FieldName="PASIDeliveryQty"
                            Width="80px" HeaderStyle-HorizontalAlign="Center">
                            <PropertiesTextEdit DisplayFormatString="{0:n0}">
                            </PropertiesTextEdit>
                            <HeaderStyle HorizontalAlign="Center" Wrap="True"></HeaderStyle>
                            <CellStyle Font-Names="Tahoma" Font-Size="8pt" HorizontalAlign="Right">
                            </CellStyle>
                        </dx:GridViewDataTextColumn>--%>                        
                        <dx:GridViewDataTextColumn VisibleIndex="23" Caption="RECEIVE DATE" FieldName="ReceiveDate"
                            Width="80px" HeaderStyle-HorizontalAlign="Center">
                            <PropertiesTextEdit DisplayFormatString="dd MMM yyyy">
                            </PropertiesTextEdit>
                            <HeaderStyle HorizontalAlign="Center" Wrap="True"></HeaderStyle>
                            <CellStyle Font-Names="Tahoma" Font-Size="8pt" HorizontalAlign="Right">
                            </CellStyle>
                        </dx:GridViewDataTextColumn>
                        <dx:GridViewDataTextColumn VisibleIndex="28" Caption="RECEIVE QTY" FieldName="ReceivingQty"
                            Width="80px" HeaderStyle-HorizontalAlign="Center">
                            <HeaderStyle HorizontalAlign="Center" Wrap="True"></HeaderStyle>
                            <PropertiesTextEdit DisplayFormatString="{0:n0}">
                            </PropertiesTextEdit>
                            <CellStyle Font-Names="Tahoma" Font-Size="8pt" HorizontalAlign="Center">
                            </CellStyle>
                        </dx:GridViewDataTextColumn>
                        <dx:GridViewDataTextColumn VisibleIndex="29" Caption="INVOICE NO" FieldName="InvoiceNo"
                            Width="80px" HeaderStyle-HorizontalAlign="Center">
                            <HeaderStyle HorizontalAlign="Center" Wrap="True"></HeaderStyle>                            
                            <CellStyle Font-Names="Tahoma" Font-Size="8pt" HorizontalAlign="Center">
                            </CellStyle>
                        </dx:GridViewDataTextColumn>
                        <dx:GridViewDataTextColumn VisibleIndex="30" Caption="INVOICE DATE" FieldName="InvoiceDate"
                            Width="80px" HeaderStyle-HorizontalAlign="Center">
                            <PropertiesTextEdit DisplayFormatString="dd MMM yyyy">
                            </PropertiesTextEdit>
                            <HeaderStyle HorizontalAlign="Center" Wrap="True"></HeaderStyle>                            
                            <CellStyle Font-Names="Tahoma" Font-Size="8pt" HorizontalAlign="Center">
                            </CellStyle>
                        </dx:GridViewDataTextColumn>
                        <dx:GridViewDataTextColumn VisibleIndex="31" Caption="PRICE/PCS" FieldName="Price"
                            Width="80px" HeaderStyle-HorizontalAlign="Center">
                            <HeaderStyle HorizontalAlign="Center" Wrap="True"></HeaderStyle>
                            <CellStyle Font-Names="Tahoma" Font-Size="8pt" HorizontalAlign="Center">
                            </CellStyle>
                        </dx:GridViewDataTextColumn>
                        <dx:GridViewDataTextColumn VisibleIndex="32" Caption="TOTAL AMOUNT" FieldName="Total"
                            Width="80px" HeaderStyle-HorizontalAlign="Center">
                            <HeaderStyle HorizontalAlign="Center" Wrap="True"></HeaderStyle>
                            <PropertiesTextEdit DisplayFormatString="{0:n0}">
                            </PropertiesTextEdit>
                            <CellStyle Font-Names="Tahoma" Font-Size="8pt" HorizontalAlign="Center">
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
