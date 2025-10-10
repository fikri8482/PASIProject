<%@ Page Title="" Language="vb" AutoEventWireup="false" MasterPageFile="~/SiteContent.Master"
    CodeBehind="SummaryOutstanding.aspx.vb" Inherits="PASISystem.SummaryOutstanding" %>

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
                                                        Text=" " Checked="true">
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
                                        <dx:ASPxLabel ID="ASPxLabel2" runat="server" Text="PO NO." Font-Names="Tahoma" Font-Size="8pt">
                                        </dx:ASPxLabel>
                                    </td>
                                    <td>
                                        <dx:ASPxTextBox ID="txtPONo" runat="server" Width="120px" Font-Names="Tahoma" Font-Size="8pt"
                                            ClientInstanceName="txtPONo">
                                        </dx:ASPxTextBox>
                                    </td>
                                    <td style="width: 5px;">
                                    </td>
                                    <td align="left" valign="middle">
                                        <%--<dx:ASPxLabel ID="ASPxLabel4" runat="server" Text="PO PROGRESS"
                                            Font-Names="Tahoma" Font-Size="8pt">
                                        </dx:ASPxLabel>--%>
                                        <dx:ASPxLabel ID="ASPxLabel7" runat="server" Text="PASI RECEIVE" Font-Names="Tahoma"
                                            Font-Size="8pt">
                                        </dx:ASPxLabel>
                                    </td>
                                    <td align="left" valign="middle">
                                        <%--<table>
                                            <tr>
                                                <td>
                                                    <dx:ASPxRadioButton ID="rdrPOPAll" ClientInstanceName="rdrPOPAll" runat="server" Text="ALL" GroupName="POProgress" Font-Names="Tahoma" Font-Size="8pt">
                                                        <ClientSideEvents CheckedChanged="function(s, e) {
                                                            grid.PerformCallback('clear');
                                                            lblInfo.SetText('');
                                                        }" />
                                                    </dx:ASPxRadioButton>
                                                </td>
                                                <td>
                                                    <dx:ASPxRadioButton ID="rdrPOPComplete" ClientInstanceName="rdrPOPComplete" runat="server" Text="COMPLETE" GroupName="POProgress" Font-Names="Tahoma" Font-Size="8pt">
                                                        <ClientSideEvents CheckedChanged="function(s, e) {
                                                            grid.PerformCallback('clear');
                                                            lblInfo.SetText('');
                                                        }" />
                                                    </dx:ASPxRadioButton>
                                                </td>
                                                <td>
                                                    <dx:ASPxRadioButton ID="rdrPOPRemaining" ClientInstanceName="rdrPOPRemaining" runat="server" Text="REMAINING" GroupName="POProgress" Font-Names="Tahoma" Font-Size="8pt">
                                                        <ClientSideEvents CheckedChanged="function(s, e) {
                                                            grid.PerformCallback('clear');
                                                            lblInfo.SetText('');
                                                        }" />
                                                    </dx:ASPxRadioButton>
                                                </td>
                                                <td>
                                                    <dx:ASPxRadioButton ID="rdrPOPDiff" ClientInstanceName="rdrPOPDiff" runat="server" Text="DIFF." GroupName="POProgress" Font-Names="Tahoma" Font-Size="8pt">
                                                        <ClientSideEvents CheckedChanged="function(s, e) {
                                                            grid.PerformCallback('clear');
                                                            lblInfo.SetText('');
                                                        }" />
                                                    </dx:ASPxRadioButton>
                                                </td>
                                            </tr>
                                        </table>--%>
                                        <table>
                                            <tr>
                                                <td>
                                                    <dx:ASPxRadioButton ID="rdrPRAll" ClientInstanceName="rdrPRAll" runat="server" Text="ALL"
                                                        GroupName="PASIReceive" Font-Names="Tahoma" Font-Size="8pt">
                                                        <ClientSideEvents CheckedChanged="function(s, e) {
                                                            grid.PerformCallback('clear');
                                                            lblInfo.SetText('');
                                                        }" />
                                                    </dx:ASPxRadioButton>
                                                </td>
                                                <td>
                                                    <dx:ASPxRadioButton ID="rdrPRComplete" ClientInstanceName="rdrPRComplete" runat="server"
                                                        Text="COMPLETE" GroupName="PASIReceive" Font-Names="Tahoma" Font-Size="8pt">
                                                        <ClientSideEvents CheckedChanged="function(s, e) {
                                                            grid.PerformCallback('clear');
                                                            lblInfo.SetText('');
                                                        }" />
                                                    </dx:ASPxRadioButton>
                                                </td>
                                                <td>
                                                    <dx:ASPxRadioButton ID="rdrPRRemaining" ClientInstanceName="rdrPRRemaining" runat="server"
                                                        Text="REMAINING" GroupName="PASIReceive" Font-Names="Tahoma" Font-Size="8pt">
                                                        <ClientSideEvents CheckedChanged="function(s, e) {
                                                            grid.PerformCallback('clear');
                                                            lblInfo.SetText('');
                                                        }" />
                                                    </dx:ASPxRadioButton>
                                                </td>
                                                <td>
                                                    <dx:ASPxRadioButton ID="rdrPRDiff" ClientInstanceName="rdrPRDiff" runat="server"
                                                        Text="DIFF." GroupName="PASIReceive" Font-Names="Tahoma" Font-Size="8pt">
                                                        <ClientSideEvents CheckedChanged="function(s, e) {
                                                            grid.PerformCallback('clear');
                                                            lblInfo.SetText('');
                                                        }" />
                                                    </dx:ASPxRadioButton>
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
                                        <dx:ASPxLabel ID="ASPxLabel3" runat="server" Text="SUPPLIER PLAN DELIVERY DATE" Font-Names="Tahoma"
                                            Font-Size="8pt">
                                        </dx:ASPxLabel>
                                    </td>
                                    <td align="left" valign="middle">
                                        <table>
                                            <tr>
                                                <td>
                                                    <dx:ASPxCheckBox ID="chkSupplierPlanDelDate" runat="server" ClientInstanceName="chkSupplierPlanDelDate"
                                                        Text=" " Checked="true">
                                                        <ClientSideEvents CheckedChanged="function (s, e) {
                                                                if (s.GetChecked()==true) {
                                                                    dtSupplierPlanDelDateFrom.SetEnabled(true);
                                                                    dtSupplierPlanDelDateTo.SetEnabled(true);
                                                                } else {
                                                                    dtSupplierPlanDelDateFrom.SetEnabled(false);
                                                                    dtSupplierPlanDelDateTo.SetEnabled(false);
                                                                }
                                                                grid.PerformCallback('clear');
                                                          }" />
                                                    </dx:ASPxCheckBox>
                                                </td>
                                                <td>
                                                    <dx:ASPxDateEdit ID="dtSupplierPlanDelDateFrom" runat="server" Font-Names="Tahoma"
                                                        Font-Size="8pt" Width="100px" EditFormat="Custom" EditFormatString="dd MMM yyyy"
                                                        ClientInstanceName="dtSupplierPlanDelDateFrom">
                                                    </dx:ASPxDateEdit>
                                                </td>
                                                <td>
                                                    ~
                                                </td>
                                                <td>
                                                    <dx:ASPxDateEdit ID="dtSupplierPlanDelDateTo" runat="server" Font-Names="Tahoma"
                                                        Font-Size="8pt" Width="100px" EditFormat="Custom" EditFormatString="dd MMM yyyy"
                                                        ClientInstanceName="dtSupplierPlanDelDateTo">
                                                    </dx:ASPxDateEdit>
                                                </td>
                                            </tr>
                                        </table>
                                    </td>
                                    <td style="width: 5px;">
                                    </td>
                                    <td>
                                        <dx:ASPxLabel ID="ASPxLabel6" runat="server" Text="SUPPLIER SJ. NO." Font-Names="Tahoma"
                                            Font-Size="8pt">
                                        </dx:ASPxLabel>
                                    </td>
                                    <td>
                                        <dx:ASPxTextBox ID="txtSupplierSJNo" runat="server" Width="120px" Font-Names="Tahoma"
                                            Font-Size="8pt" ClientInstanceName="txtSupplierSJNo">
                                        </dx:ASPxTextBox>
                                    </td>
                                    <td style="width: 5px;">
                                    </td>
                                    <td align="left" valign="middle">
                                        <dx:ASPxLabel ID="ASPxLabel10" runat="server" Text="PASI DELIVERY" Font-Names="Tahoma"
                                            Font-Size="8pt">
                                        </dx:ASPxLabel>
                                    </td>
                                    <td align="left" valign="middle">
                                        <table>
                                            <tr>
                                                <td>
                                                    <dx:ASPxRadioButton ID="rdrPDAll" ClientInstanceName="rdrPDAll" runat="server" Text="ALL"
                                                        GroupName="PASIDelivery" Font-Names="Tahoma" Font-Size="8pt">
                                                        <ClientSideEvents CheckedChanged="function(s, e) {
                                                            grid.PerformCallback('clear');
                                                            lblInfo.SetText('');
                                                        }" />
                                                    </dx:ASPxRadioButton>
                                                </td>
                                                <td>
                                                    <dx:ASPxRadioButton ID="rdrPDComplete" ClientInstanceName="rdrPDComplete" runat="server"
                                                        Text="COMPLETE" GroupName="PASIDelivery" Font-Names="Tahoma" Font-Size="8pt">
                                                        <ClientSideEvents CheckedChanged="function(s, e) {
                                                            grid.PerformCallback('clear');
                                                            lblInfo.SetText('');
                                                        }" />
                                                    </dx:ASPxRadioButton>
                                                </td>
                                                <td>
                                                    <dx:ASPxRadioButton ID="rdrPDRemaining" ClientInstanceName="rdrPDRemaining" runat="server"
                                                        Text="REMAINING" GroupName="PASIDelivery" Font-Names="Tahoma" Font-Size="8pt">
                                                        <ClientSideEvents CheckedChanged="function(s, e) {
                                                            grid.PerformCallback('clear');
                                                            lblInfo.SetText('');
                                                        }" />
                                                    </dx:ASPxRadioButton>
                                                </td>
                                                <td>
                                                    <%--<dx:ASPxRadioButton ID="rdrPDDiff" ClientInstanceName="rdrPDDiff" runat="server" Text="DIFF." GroupName="PASIDelivery" Font-Names="Tahoma" Font-Size="8pt">
                                                        <ClientSideEvents CheckedChanged="function(s, e) {
                                                            grid.PerformCallback('clear');
                                                            lblInfo.SetText('');
                                                        }" />
                                                    </dx:ASPxRadioButton>--%>
                                                </td>
                                            </tr>
                                        </table>
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
                                        <dx:ASPxLabel ID="ASPxLabel9" runat="server" Text="PASI SJ. NO." Font-Names="Tahoma"
                                            Font-Size="8pt">
                                        </dx:ASPxLabel>
                                    </td>
                                    <td>
                                        <dx:ASPxTextBox ID="txtPASISJNo" runat="server" Width="120px" Font-Names="Tahoma"
                                            Font-Size="8pt" ClientInstanceName="txtPASISJNo">
                                        </dx:ASPxTextBox>
                                    </td>
                                    <td style="width: 5px;">
                                    </td>
                                    <td align="left" valign="middle" height="20px" width="130px">
                                        <dx:ASPxLabel ID="ASPxLabel13" runat="server" Text="AFFILIATE RECEIVE" Font-Names="Tahoma"
                                            Font-Size="8pt">
                                        </dx:ASPxLabel>
                                    </td>
                                    <td align="left" valign="middle">
                                        <table>
                                            <tr>
                                                <td>
                                                    <dx:ASPxRadioButton ID="rdrARAll" ClientInstanceName="rdrARAll" runat="server" Text="ALL"
                                                        GroupName="AffiliateReceive" Font-Names="Tahoma" Font-Size="8pt">
                                                        <ClientSideEvents CheckedChanged="function(s, e) {
                                                            grid.PerformCallback('clear');
                                                            lblInfo.SetText('');
                                                        }" />
                                                    </dx:ASPxRadioButton>
                                                </td>
                                                <td>
                                                    <dx:ASPxRadioButton ID="rdrARComplete" ClientInstanceName="rdrARComplete" runat="server"
                                                        Text="COMPLETE" GroupName="AffiliateReceive" Font-Names="Tahoma" Font-Size="8pt">
                                                        <ClientSideEvents CheckedChanged="function(s, e) {
                                                            grid.PerformCallback('clear');
                                                            lblInfo.SetText('');
                                                        }" />
                                                    </dx:ASPxRadioButton>
                                                </td>
                                                <td>
                                                    <dx:ASPxRadioButton ID="rdrARRemaining" ClientInstanceName="rdrARRemaining" runat="server"
                                                        Text="REMAINING" GroupName="AffiliateReceive" Font-Names="Tahoma" Font-Size="8pt">
                                                        <ClientSideEvents CheckedChanged="function(s, e) {
                                                            grid.PerformCallback('clear');
                                                            lblInfo.SetText('');
                                                        }" />
                                                    </dx:ASPxRadioButton>
                                                </td>
                                                <td>
                                                    <dx:ASPxRadioButton ID="rdrARDiff" ClientInstanceName="rdrARDiff" runat="server"
                                                        Text="DIFF." GroupName="AffiliateReceive" Font-Names="Tahoma" Font-Size="8pt">
                                                        <ClientSideEvents CheckedChanged="function(s, e) {
                                                            grid.PerformCallback('clear');
                                                            lblInfo.SetText('');
                                                        }" />
                                                    </dx:ASPxRadioButton>
                                                </td>
                                            </tr>
                                        </table>
                                    </td>
                                </tr>
                                <!-- ROW 5 -->
                                <tr>
                                    <td style="width: 5px;">
                                    </td>
                                    <td align="left" valign="middle">
                                        <dx:ASPxLabel ID="ASPxLabel11" runat="server" Text="PASI RECEIVE DATE" Font-Names="Tahoma"
                                            Font-Size="8pt">
                                        </dx:ASPxLabel>
                                    </td>
                                    <td align="left" valign="middle">
                                        <table>
                                            <tr>
                                                <td>
                                                    <dx:ASPxCheckBox ID="chkPASIRecDate" runat="server" ClientInstanceName="chkPASIRecDate"
                                                        Text=" " Checked="true">
                                                        <ClientSideEvents CheckedChanged="function (s, e) {
                                                                if (s.GetChecked()==true) {
                                                                    dtPASIRecDateFrom.SetEnabled(true);
                                                                    dtPASIRecDateTo.SetEnabled(true);
                                                                } else {
                                                                    dtPASIRecDateFrom.SetEnabled(false);
                                                                    dtPASIRecDateTo.SetEnabled(false);
                                                                }
                                                                grid.PerformCallback('clear');
                                                          }" />
                                                    </dx:ASPxCheckBox>
                                                </td>
                                                <td>
                                                    <dx:ASPxDateEdit ID="dtPASIRecDateFrom" runat="server" Font-Names="Tahoma" Font-Size="8pt"
                                                        Width="100px" EditFormat="Custom" EditFormatString="dd MMM yyyy" ClientInstanceName="dtPASIRecDateFrom">
                                                    </dx:ASPxDateEdit>
                                                </td>
                                                <td>
                                                    ~
                                                </td>
                                                <td>
                                                    <dx:ASPxDateEdit ID="dtPASIRecDateTo" runat="server" Font-Names="Tahoma" Font-Size="8pt"
                                                        Width="100px" EditFormat="Custom" EditFormatString="dd MMM yyyy" ClientInstanceName="dtPASIRecDateTo">
                                                    </dx:ASPxDateEdit>
                                                </td>
                                            </tr>
                                        </table>
                                    </td>
                                    <td style="width: 5px;">
                                    </td>
                                    <td>
                                        <dx:ASPxLabel ID="ASPxLabel12" runat="server" Text="SUPPLIER INV. NO." Font-Names="Tahoma"
                                            Font-Size="8pt">
                                        </dx:ASPxLabel>
                                    </td>
                                    <td>
                                        <dx:ASPxTextBox ID="txtSupplierInvNo" runat="server" Width="120px" Font-Names="Tahoma"
                                            Font-Size="8pt" ClientInstanceName="txtSupplierInvNo">
                                        </dx:ASPxTextBox>
                                    </td>
                                    <td style="width: 5px;">
                                    </td>
                                    <td align="left" valign="middle" height="20px" width="130px">
                                        <dx:ASPxLabel ID="ASPxLabel21" runat="server" Text="SUPPLIER INVOICE" Font-Names="Tahoma"
                                            Font-Size="8pt">
                                        </dx:ASPxLabel>
                                    </td>
                                    <td align="left" valign="middle">
                                        <table>
                                            <tr>
                                                <td>
                                                    <dx:ASPxRadioButton ID="rdrSIAll" ClientInstanceName="rdrSIAll" runat="server" Text="ALL"
                                                        GroupName="SupplierInvoice" Font-Names="Tahoma" Font-Size="8pt">
                                                        <ClientSideEvents CheckedChanged="function(s, e) {
                                                            grid.PerformCallback('clear');
                                                            lblInfo.SetText('');
                                                        }" />
                                                    </dx:ASPxRadioButton>
                                                </td>
                                                <td>
                                                    <dx:ASPxRadioButton ID="rdrSIComplete" ClientInstanceName="rdrSIComplete" runat="server"
                                                        Text="COMPLETE" GroupName="SupplierInvoice" Font-Names="Tahoma" Font-Size="8pt">
                                                        <ClientSideEvents CheckedChanged="function(s, e) {
                                                            grid.PerformCallback('clear');
                                                            lblInfo.SetText('');
                                                        }" />
                                                    </dx:ASPxRadioButton>
                                                </td>
                                                <td>
                                                    <dx:ASPxRadioButton ID="rdrSIRemaining" ClientInstanceName="rdrSIRemaining" runat="server"
                                                        Text="REMAINING" GroupName="SupplierInvoice" Font-Names="Tahoma" Font-Size="8pt">
                                                        <ClientSideEvents CheckedChanged="function(s, e) {
                                                            grid.PerformCallback('clear');
                                                            lblInfo.SetText('');
                                                        }" />
                                                    </dx:ASPxRadioButton>
                                                </td>
                                                <td>
                                                    <%--<dx:ASPxRadioButton ID="rdrPIDiff" ClientInstanceName="rdrPIDiff" runat="server" Text="DIFF." GroupName="PASIInvoice" Font-Names="Tahoma" Font-Size="8pt">
                                                        <ClientSideEvents CheckedChanged="function(s, e) {
                                                            grid.PerformCallback('clear');
                                                            lblInfo.SetText('');
                                                        }" />
                                                    </dx:ASPxRadioButton>--%>
                                                </td>
                                            </tr>
                                        </table>
                                    </td>
                                </tr>
                                <!-- ROW 6 -->
                                <tr>
                                    <td style="width: 5px;">
                                    </td>
                                    <td align="left" valign="middle">
                                        <dx:ASPxLabel ID="ASPxLabel14" runat="server" Text="PASI DELIVERY DATE" Font-Names="Tahoma"
                                            Font-Size="8pt">
                                        </dx:ASPxLabel>
                                    </td>
                                    <td align="left" valign="middle">
                                        <table>
                                            <tr>
                                                <td>
                                                    <dx:ASPxCheckBox ID="chkPASIDelDate" runat="server" ClientInstanceName="chkPASIDelDate"
                                                        Text=" " Checked="true">
                                                        <ClientSideEvents CheckedChanged="function (s, e) {
                                                                if (s.GetChecked()==true) {
                                                                    dtPASIDelDateFrom.SetEnabled(true);
                                                                    dtPASIDelDateTo.SetEnabled(true);
                                                                } else {
                                                                    dtPASIDelDateFrom.SetEnabled(false);
                                                                    dtPASIDelDateTo.SetEnabled(false);
                                                                }
                                                                grid.PerformCallback('clear');
                                                          }" />
                                                    </dx:ASPxCheckBox>
                                                </td>
                                                <td>
                                                    <dx:ASPxDateEdit ID="dtPASIDelDateFrom" runat="server" Font-Names="Tahoma" Font-Size="8pt"
                                                        Width="100px" EditFormat="Custom" EditFormatString="dd MMM yyyy" ClientInstanceName="dtPASIDelDateFrom">
                                                    </dx:ASPxDateEdit>
                                                </td>
                                                <td>
                                                    ~
                                                </td>
                                                <td>
                                                    <dx:ASPxDateEdit ID="dtPASIDelDateTo" runat="server" Font-Names="Tahoma" Font-Size="8pt"
                                                        Width="100px" EditFormat="Custom" EditFormatString="dd MMM yyyy" ClientInstanceName="dtPASIDelDateTo">
                                                    </dx:ASPxDateEdit>
                                                </td>
                                            </tr>
                                        </table>
                                    </td>
                                    <td style="width: 5px;">
                                    </td>
                                    <td>
                                        <dx:ASPxLabel ID="ASPxLabel20" runat="server" Text="PASI INV. NO." Font-Names="Tahoma"
                                            Font-Size="8pt">
                                        </dx:ASPxLabel>
                                    </td>
                                    <td>
                                        <dx:ASPxTextBox ID="txtPASIInvNo" runat="server" Width="120px" Font-Names="Tahoma"
                                            Font-Size="8pt" ClientInstanceName="txtPASIInvNo">
                                        </dx:ASPxTextBox>
                                    </td>
                                    <td style="width: 5px;">
                                    </td>
                                    <td align="left" valign="middle" height="20px" width="130px">
                                        <dx:ASPxLabel ID="ASPxLabel16" runat="server" Text="PASI INVOICE" Font-Names="Tahoma"
                                            Font-Size="8pt">
                                        </dx:ASPxLabel>
                                    </td>
                                    <td align="left" valign="middle">
                                        <table>
                                            <tr>
                                                <td>
                                                    <dx:ASPxRadioButton ID="rdrPIAll" ClientInstanceName="rdrPIAll" runat="server" Text="ALL"
                                                        GroupName="PASIInvoice" Font-Names="Tahoma" Font-Size="8pt">
                                                        <ClientSideEvents CheckedChanged="function(s, e) {
                                                            grid.PerformCallback('clear');
                                                            lblInfo.SetText('');
                                                        }" />
                                                    </dx:ASPxRadioButton>
                                                </td>
                                                <td>
                                                    <dx:ASPxRadioButton ID="rdrPIComplete" ClientInstanceName="rdrPIComplete" runat="server"
                                                        Text="COMPLETE" GroupName="PASIInvoice" Font-Names="Tahoma" Font-Size="8pt">
                                                        <ClientSideEvents CheckedChanged="function(s, e) {
                                                            grid.PerformCallback('clear');
                                                            lblInfo.SetText('');
                                                        }" />
                                                    </dx:ASPxRadioButton>
                                                </td>
                                                <td>
                                                    <dx:ASPxRadioButton ID="rdrPIRemaining" ClientInstanceName="rdrPIRemaining" runat="server"
                                                        Text="REMAINING" GroupName="PASIInvoice" Font-Names="Tahoma" Font-Size="8pt">
                                                        <ClientSideEvents CheckedChanged="function(s, e) {
                                                            grid.PerformCallback('clear');
                                                            lblInfo.SetText('');
                                                        }" />
                                                    </dx:ASPxRadioButton>
                                                </td>
                                                <td>
                                                    <%--<dx:ASPxRadioButton ID="rdrPIDiff" ClientInstanceName="rdrPIDiff" runat="server" Text="DIFF." GroupName="PASIInvoice" Font-Names="Tahoma" Font-Size="8pt">
                                                        <ClientSideEvents CheckedChanged="function(s, e) {
                                                            grid.PerformCallback('clear');
                                                            lblInfo.SetText('');
                                                        }" />
                                                    </dx:ASPxRadioButton>--%>
                                                </td>
                                            </tr>
                                        </table>
                                    </td>
                                </tr>
                                <!-- ROW 7 -->
                                <tr>
                                    <td style="width: 5px;">
                                    </td>
                                    <td align="left" valign="middle">
                                        <dx:ASPxLabel ID="ASPxLabel15" runat="server" Text="AFFILIATE RECEIVE DATE" Font-Names="Tahoma"
                                            Font-Size="8pt">
                                        </dx:ASPxLabel>
                                    </td>
                                    <td align="left" valign="middle">
                                        <table>
                                            <tr>
                                                <td>
                                                    <dx:ASPxCheckBox ID="chkAffiliateRecDate" runat="server" ClientInstanceName="chkAffiliateRecDate"
                                                        Text=" " Checked="true">
                                                        <ClientSideEvents CheckedChanged="function (s, e) {
                                                                if (s.GetChecked()==true) {
                                                                    dtAffiliateRecDateFrom.SetEnabled(true);
                                                                    dtAffiliateRecDateTo.SetEnabled(true);
                                                                } else {
                                                                    dtAffiliateRecDateFrom.SetEnabled(false);
                                                                    dtAffiliateRecDateTo.SetEnabled(false);
                                                                }
                                                                grid.PerformCallback('clear');
                                                          }" />
                                                    </dx:ASPxCheckBox>
                                                </td>
                                                <td>
                                                    <dx:ASPxDateEdit ID="dtAffiliateRecDateFrom" runat="server" Font-Names="Tahoma" Font-Size="8pt"
                                                        Width="100px" EditFormat="Custom" EditFormatString="dd MMM yyyy" ClientInstanceName="dtAffiliateRecDateFrom">
                                                    </dx:ASPxDateEdit>
                                                </td>
                                                <td>
                                                    ~
                                                </td>
                                                <td>
                                                    <dx:ASPxDateEdit ID="dtAffiliateRecDateTo" runat="server" Font-Names="Tahoma" Font-Size="8pt"
                                                        Width="100px" EditFormat="Custom" EditFormatString="dd MMM yyyy" ClientInstanceName="dtAffiliateRecDateTo">
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

                                <!-- ROW 10 -->
                                <tr>
                                    <td style="width: 5px;">
                                    </td>
                                    <td align="left" valign="middle">
                                        <dx:ASPxLabel ID="ASPxLabel18" runat="server" Text="PART CODE/NAME" Font-Names="Tahoma"
                                            Font-Size="8pt">
                                        </dx:ASPxLabel>
                                    </td>
                                    <td align="left" valign="middle">
                                        <table>
                                            <tr>
                                                <td>
                                                    <dx:ASPxComboBox ID="cboPart" runat="server" ClientInstanceName="cboPart" Font-Names="Tahoma"
                                                        TextFormatString="{0}" Font-Size="8pt" Width="120px">
                                                        <ClientSideEvents SelectedIndexChanged="function(s, e) {
	                                                        txtPartName.SetText(cboPart.GetSelectedItem().GetColumnText(1));
                                                            grid.PerformCallback('clear');
                                                            }" />
                                                    </dx:ASPxComboBox>
                                                </td>
                                                <td>
                                                    <dx:ASPxTextBox ID="txtPartName" runat="server" Width="200px" Font-Names="Tahoma"
                                                        Font-Size="8pt" BackColor="#CCCCCC" ReadOnly="True" MaxLength="20" ClientInstanceName="txtPartName">
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
                                    <td align="right" valign="middle">
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
                                                            chkSupplierPeriod.SetChecked(true);
                                                            dtSupplierPeriod.SetEnabled(true);
                                                            rdrSADAll.SetChecked(true);
                                                            rdrPADAll.SetChecked(true);
                                                            rdrRRQAll.SetChecked(true);
                                                            txtSupplierSJNo.SetText('');
                                                            chkRecDate.SetChecked(true);
                                                            dtRecDateFrom.SetEnabled(true);
                                                            dtRecDateTo.SetEnabled(true);
                                                            cboSupplier.SetText('==ALL==');
                                                            txtSupplierName.SetText('==ALL==');
                                                            cboPart.SetText('==ALL==');
                                                            txtPartName.SetText('==ALL==');
                                                            txtPONo.SetText('');
                                                            rdrSDAll.SetChecked(true);
                                                            rdrPOKAll.SetChecked(true);
                                                            rdrMCPAll.SetChecked(true);
                                                            rdrGRSAll.SetChecked(true);

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
                        <dx:GridViewDataTextColumn VisibleIndex="0" Caption="NO" FieldName="ColNo" Width="0px"
                            HeaderStyle-HorizontalAlign="Center">
                            <HeaderStyle HorizontalAlign="Center" Wrap="True"></HeaderStyle>
                            <CellStyle Font-Names="Tahoma" Font-Size="8pt" HorizontalAlign="Center">
                            </CellStyle>
                        </dx:GridViewDataTextColumn>
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
                        <dx:GridViewDataTextColumn VisibleIndex="6" Caption="PO KANBAN" FieldName="POKanban"
                            Width="60px" HeaderStyle-HorizontalAlign="Center">
                            <HeaderStyle HorizontalAlign="Center" Wrap="True"></HeaderStyle>
                            <CellStyle Font-Names="Tahoma" Font-Size="8pt" HorizontalAlign="Center">
                            </CellStyle>
                        </dx:GridViewDataTextColumn>
                        <dx:GridViewDataTextColumn VisibleIndex="7" Caption="PO ISSUE DATE AFFILIATE"
                            FieldName="EntryDate" Width="110px" 
                            HeaderStyle-HorizontalAlign="Center">
                            <HeaderStyle HorizontalAlign="Center" Wrap="True"></HeaderStyle>
                            <PropertiesTextEdit DisplayFormatString="dd MMM yyyy">
                            </PropertiesTextEdit>
                            <CellStyle Font-Names="Tahoma" Font-Size="8pt" HorizontalAlign="Center">
                            </CellStyle>
                        </dx:GridViewDataTextColumn>
                        <dx:GridViewDataTextColumn VisibleIndex="8" Caption="PO SEND TO SUPPLIER DATE"
                            FieldName="PASISendAffiliateDate" Width="110px" 
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
                        <dx:GridViewDataTextColumn VisibleIndex="12" Caption="KANBAN NO." FieldName="KanbanNo"
                            Width="80px" HeaderStyle-HorizontalAlign="Center">
                            <HeaderStyle HorizontalAlign="Center" Wrap="True"></HeaderStyle>
                            <CellStyle Font-Names="Tahoma" Font-Size="8pt">
                            </CellStyle>
                        </dx:GridViewDataTextColumn>
                        <dx:GridViewDataTextColumn Caption="SCHEDULE QTY" FieldName="KanbanQty"
                            VisibleIndex="13" Width="90px" HeaderStyle-HorizontalAlign="Center">
                            <PropertiesTextEdit DisplayFormatString="{0:n0}">
                            </PropertiesTextEdit>
                            <HeaderStyle HorizontalAlign="Center"></HeaderStyle>
                            <CellStyle Font-Names="Tahoma" Font-Size="8pt" HorizontalAlign="Right">
                            </CellStyle>
                        </dx:GridViewDataTextColumn>  
                        <dx:GridViewDataTextColumn VisibleIndex="14" Caption="SCHEDULE ETD SUPPLIER"
                            FieldName="ETDSupp" Width="110px" 
                            HeaderStyle-HorizontalAlign="Center">
                            <HeaderStyle HorizontalAlign="Center" Wrap="True"></HeaderStyle>
                            <PropertiesTextEdit DisplayFormatString="dd MMM yyyy">
                            </PropertiesTextEdit>
                            <CellStyle Font-Names="Tahoma" Font-Size="8pt" HorizontalAlign="Center">
                            </CellStyle>
                        </dx:GridViewDataTextColumn>
                        <dx:GridViewDataTextColumn VisibleIndex="15" Caption="SCHEDULE ETA AFFILIATE"
                            FieldName="ETAAff" Width="110px" 
                            HeaderStyle-HorizontalAlign="Center">
                            <HeaderStyle HorizontalAlign="Center" Wrap="True"></HeaderStyle>
                            <PropertiesTextEdit DisplayFormatString="dd MMM yyyy">
                            </PropertiesTextEdit>
                            <CellStyle Font-Names="Tahoma" Font-Size="8pt" HorizontalAlign="Center">
                            </CellStyle>
                        </dx:GridViewDataTextColumn>                    
                        <dx:GridViewDataTextColumn VisibleIndex="16" Caption="SUPPLIER DELIVERY DATE" FieldName="SupplierDeliveryDate"
                            Width="110px" HeaderStyle-HorizontalAlign="Center">
                            <HeaderStyle HorizontalAlign="Center" Wrap="True"></HeaderStyle>
                            <PropertiesTextEdit DisplayFormatString="dd MMM yyyy">
                            </PropertiesTextEdit>
                            <CellStyle Font-Names="Tahoma" Font-Size="8pt" HorizontalAlign="Center">
                            </CellStyle>
                        </dx:GridViewDataTextColumn>
                        <dx:GridViewDataTextColumn VisibleIndex="17" Caption="SUPPLIER SURAT JALAN NO." FieldName="SupplierSuratJalanNo"
                            Width="100px" HeaderStyle-HorizontalAlign="Center">
                            <HeaderStyle HorizontalAlign="Center" Wrap="True"></HeaderStyle>
                            <CellStyle Font-Names="Tahoma" Font-Size="8pt" HorizontalAlign="Left">
                            </CellStyle>
                        </dx:GridViewDataTextColumn>
                        <dx:GridViewDataTextColumn VisibleIndex="18" Caption="SUPPLIER DELIVERY QTY" FieldName="SupplierDeliveryQty"
                            Width="80px" HeaderStyle-HorizontalAlign="Center">
                            <PropertiesTextEdit DisplayFormatString="{0:n0}">
                            </PropertiesTextEdit>
                            <HeaderStyle HorizontalAlign="Center" Wrap="True"></HeaderStyle>
                            <CellStyle Font-Names="Tahoma" Font-Size="8pt" HorizontalAlign="Right">
                            </CellStyle>
                        </dx:GridViewDataTextColumn>
                        <dx:GridViewDataTextColumn VisibleIndex="19" Caption="REMAINING" FieldName="RemainingQtyPOPASI"
                            Width="80px" HeaderStyle-HorizontalAlign="Center">
                            <PropertiesTextEdit DisplayFormatString="{0:n0}">
                            </PropertiesTextEdit>
                            <HeaderStyle HorizontalAlign="Center" Wrap="True"></HeaderStyle>
                            <CellStyle Font-Names="Tahoma" Font-Size="8pt" HorizontalAlign="Right">
                            </CellStyle>
                        </dx:GridViewDataTextColumn>
                        <dx:GridViewDataTextColumn VisibleIndex="20" Caption="PASI RECEIVE DATE" FieldName="PASIReceiveDate"
                            Width="80px" HeaderStyle-HorizontalAlign="Center">
                            <HeaderStyle HorizontalAlign="Center" Wrap="True"></HeaderStyle>
                            <PropertiesTextEdit DisplayFormatString="dd MMM yyyy">
                            </PropertiesTextEdit>
                            <CellStyle Font-Names="Tahoma" Font-Size="8pt" HorizontalAlign="Center">
                            </CellStyle>
                        </dx:GridViewDataTextColumn>
                        <dx:GridViewDataTextColumn VisibleIndex="21" Caption="PASI RECEIVING QTY" FieldName="PASIReceivingQty"
                            Width="80px" HeaderStyle-HorizontalAlign="Center">
                            <PropertiesTextEdit DisplayFormatString="{0:n0}">
                            </PropertiesTextEdit>
                            <HeaderStyle HorizontalAlign="Center" Wrap="True"></HeaderStyle>
                            <CellStyle Font-Names="Tahoma" Font-Size="8pt" HorizontalAlign="Right">
                            </CellStyle>
                        </dx:GridViewDataTextColumn>
                        <dx:GridViewDataTextColumn VisibleIndex="22" Caption="INVOICE NO. FROM SUPPLIER"
                            FieldName="InvoiceNoFromSupplier" Width="120px" 
                            HeaderStyle-HorizontalAlign="Center">
                            <HeaderStyle HorizontalAlign="Center"></HeaderStyle>
                            <CellStyle Font-Names="Tahoma" Font-Size="8pt">
                            </CellStyle>
                        </dx:GridViewDataTextColumn>
                        <dx:GridViewDataTextColumn VisibleIndex="22" Caption="INVOICE DATE FROM SUPPLIER"
                            FieldName="InvoiceDateFromSupplier" Width="100px" 
                            HeaderStyle-HorizontalAlign="Center">
                            <HeaderStyle HorizontalAlign="Center"></HeaderStyle>
                            <PropertiesTextEdit DisplayFormatString="dd MMM yyyy">
                            </PropertiesTextEdit>
                            <CellStyle Font-Names="Tahoma" Font-Size="8pt" HorizontalAlign="Center">
                            </CellStyle>
                        </dx:GridViewDataTextColumn>
                        <dx:GridViewBandColumn Caption="INVOICE FROM SUPPLIER " VisibleIndex="23" 
                            HeaderStyle-HorizontalAlign="Center">
                            <Columns>
                                <dx:GridViewDataTextColumn VisibleIndex="0" Caption="CURR" FieldName="InvoiceFromSupplierCurr"
                                    Width="60px" HeaderStyle-HorizontalAlign="Center">
                                    <HeaderStyle HorizontalAlign="Center"></HeaderStyle>
                                    <CellStyle Font-Names="Tahoma" Font-Size="8pt">
                                    </CellStyle>
                                </dx:GridViewDataTextColumn>
                                <dx:GridViewDataTextColumn VisibleIndex="1" Caption="AMOUNT" FieldName="InvoiceFromSupplierAmount"
                                    Width="110px" HeaderStyle-HorizontalAlign="Center">
                                    <PropertiesTextEdit DisplayFormatString="{0:n2}">
                                    </PropertiesTextEdit>
                                    <HeaderStyle HorizontalAlign="Center"></HeaderStyle>
                                    <CellStyle Font-Names="Tahoma" Font-Size="8pt" HorizontalAlign="Right">
                                    </CellStyle>
                                </dx:GridViewDataTextColumn>
                            </Columns>
                            <HeaderStyle HorizontalAlign="Center"></HeaderStyle>
                        </dx:GridViewBandColumn>
<%--                        <dx:GridViewDataTextColumn VisibleIndex="24" Caption="PASI DELIVERY DATE" FieldName="PASIDeliveryDate"
                            Width="0px" HeaderStyle-HorizontalAlign="Center">
                            <HeaderStyle HorizontalAlign="Center" Wrap="True"></HeaderStyle>
                            <CellStyle Font-Names="Tahoma" Font-Size="8pt" HorizontalAlign="Center">
                            </CellStyle>
                        </dx:GridViewDataTextColumn>--%>
                        <dx:GridViewDataTextColumn VisibleIndex="25" Caption="PASI SURAT JALAN NO." FieldName="PASISuratJalanNo"
                            Width="120px" HeaderStyle-HorizontalAlign="Center">
                            <HeaderStyle HorizontalAlign="Center" Wrap="True"></HeaderStyle>
                            <CellStyle Font-Names="Tahoma" Font-Size="8pt">
                            </CellStyle>
                        </dx:GridViewDataTextColumn>
                        <dx:GridViewDataTextColumn VisibleIndex="26" Caption="PASI DELIVERY QTY" FieldName="PASIDeliveryQty"
                            Width="80px" HeaderStyle-HorizontalAlign="Center">
                            <PropertiesTextEdit DisplayFormatString="{0:n0}">
                            </PropertiesTextEdit>
                            <HeaderStyle HorizontalAlign="Center" Wrap="True"></HeaderStyle>
                            <CellStyle Font-Names="Tahoma" Font-Size="8pt" HorizontalAlign="Right">
                            </CellStyle>
                        </dx:GridViewDataTextColumn>
                        <dx:GridViewDataTextColumn VisibleIndex="27" Caption="AFFILIATE RECEIVE DATE" FieldName="AffiliateReceiveDate"
                            Width="80px" HeaderStyle-HorizontalAlign="Center">
                            <HeaderStyle HorizontalAlign="Center" Wrap="True"></HeaderStyle>
                            <PropertiesTextEdit DisplayFormatString="dd MMM yyyy">
                            </PropertiesTextEdit>
                            <CellStyle Font-Names="Tahoma" Font-Size="8pt" HorizontalAlign="Center">
                            </CellStyle>
                        </dx:GridViewDataTextColumn>
                        <dx:GridViewDataTextColumn VisibleIndex="28" Caption="AFFILIATE RECEIVING QTY" FieldName="AffiliateReceivingQty"
                            Width="80px" HeaderStyle-HorizontalAlign="Center">
                            <PropertiesTextEdit DisplayFormatString="{0:n0}">
                            </PropertiesTextEdit>
                            <HeaderStyle HorizontalAlign="Center" Wrap="True"></HeaderStyle>
                            <CellStyle Font-Names="Tahoma" Font-Size="8pt" HorizontalAlign="Right">
                            </CellStyle>
                        </dx:GridViewDataTextColumn>
                        <dx:GridViewDataTextColumn VisibleIndex="29" Caption="INVOICE NO. TO AFFILIATE" FieldName="InvoiceNoToAffiliate"
                            Width="120px" HeaderStyle-HorizontalAlign="Center">
                            <HeaderStyle HorizontalAlign="Center"></HeaderStyle>
                            <CellStyle Font-Names="Tahoma" Font-Size="8pt">
                            </CellStyle>
                        </dx:GridViewDataTextColumn>
                        <dx:GridViewDataTextColumn VisibleIndex="30" Caption="INVOICE DATE TO AFFILIATE"
                            FieldName="InvoiceDateToAffiliate" Width="100px" 
                            HeaderStyle-HorizontalAlign="Center">
                            <PropertiesTextEdit DisplayFormatString="dd MMM yyyy">
                            </PropertiesTextEdit>
                            <HeaderStyle HorizontalAlign="Center"></HeaderStyle>
                            <CellStyle Font-Names="Tahoma" Font-Size="8pt" HorizontalAlign="Center">
                            </CellStyle>
                        </dx:GridViewDataTextColumn>
                        <dx:GridViewBandColumn Caption="INVOICE TO AFFILIATE" VisibleIndex="31" 
                            HeaderStyle-HorizontalAlign="Center">
                            <Columns>
                                <dx:GridViewDataTextColumn VisibleIndex="0" Caption="CURR" FieldName="InvoiceToAffiliateCurr"
                                    Width="60px" HeaderStyle-HorizontalAlign="Center">
                                    <HeaderStyle HorizontalAlign="Center"></HeaderStyle>
                                    <CellStyle Font-Names="Tahoma" Font-Size="8pt">
                                    </CellStyle>
                                </dx:GridViewDataTextColumn>
                                <dx:GridViewDataTextColumn VisibleIndex="1" Caption="AMOUNT" FieldName="InvoiceToAffiliateAmount"
                                    Width="110px" HeaderStyle-HorizontalAlign="Center">
                                    <PropertiesTextEdit DisplayFormatString="{0:n2}">
                                    </PropertiesTextEdit>
                                    <HeaderStyle HorizontalAlign="Center"></HeaderStyle>
                                    <CellStyle Font-Names="Tahoma" Font-Size="8pt" HorizontalAlign="Right">
                                    </CellStyle>
                                </dx:GridViewDataTextColumn>
                            </Columns>
                            <HeaderStyle HorizontalAlign="Center"></HeaderStyle>
                        </dx:GridViewBandColumn>
<%--                        <dx:GridViewDataTextColumn VisibleIndex="32" Caption="UOM" FieldName="UOM" Width="0px"
                            HeaderStyle-HorizontalAlign="Center">
                            <HeaderStyle HorizontalAlign="Center" Wrap="True"></HeaderStyle>
                            <CellStyle Font-Names="Tahoma" Font-Size="8pt" HorizontalAlign="Center">
                            </CellStyle>
                        </dx:GridViewDataTextColumn>                                                                        
                        <dx:GridViewDataTextColumn VisibleIndex="33" Caption="INVOICE FROM SUPPLIER QTY"
                            FieldName="InvoiceFromSupplierQty" Width="0px" 
                            HeaderStyle-HorizontalAlign="Center">
                            <PropertiesTextEdit DisplayFormatString="{0:n0}">
                            </PropertiesTextEdit>
                            <HeaderStyle HorizontalAlign="Center" Wrap="True"></HeaderStyle>
                            <CellStyle Font-Names="Tahoma" Font-Size="8pt" HorizontalAlign="Right">
                            </CellStyle>
                        </dx:GridViewDataTextColumn>
                        <dx:GridViewDataTextColumn VisibleIndex="34" Caption="INVOICE TO AFFILIATE QTY" FieldName="InvoiceToAffiliateQty"
                            Width="0px" HeaderStyle-HorizontalAlign="Center">
                            <PropertiesTextEdit DisplayFormatString="{0:n0}">
                            </PropertiesTextEdit>
                            <HeaderStyle HorizontalAlign="Center"></HeaderStyle>
                            <CellStyle Font-Names="Tahoma" Font-Size="8pt" HorizontalAlign="Right">
                            </CellStyle>
                        </dx:GridViewDataTextColumn>     --%>                                                                   
<%--                        <dx:GridViewDataTextColumn VisibleIndex="35" Caption="SortPONo" FieldName="SortPONo"
                            Width="0px" HeaderStyle-HorizontalAlign="Center">
                            <HeaderStyle HorizontalAlign="Center" Wrap="True"></HeaderStyle>
                            <CellStyle Font-Names="Tahoma" Font-Size="8pt" HorizontalAlign="Center">
                            </CellStyle>
                        </dx:GridViewDataTextColumn>
                        <dx:GridViewDataTextColumn VisibleIndex="36" Caption="SortKanbanNo" FieldName="SortKanbanNo"
                            Width="0px" HeaderStyle-HorizontalAlign="Center">
                            <HeaderStyle HorizontalAlign="Center" Wrap="True"></HeaderStyle>
                            <CellStyle Font-Names="Tahoma" Font-Size="8pt">
                            </CellStyle>
                        </dx:GridViewDataTextColumn>
                        <dx:GridViewDataTextColumn VisibleIndex="37" Caption="SortHeader" FieldName="SortHeader"
                            Width="0px" HeaderStyle-HorizontalAlign="Center">
                            <HeaderStyle HorizontalAlign="Center"></HeaderStyle>
                        </dx:GridViewDataTextColumn>    --%>                                          
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
