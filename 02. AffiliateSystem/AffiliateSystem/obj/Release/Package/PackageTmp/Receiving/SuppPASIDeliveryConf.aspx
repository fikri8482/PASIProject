<%@ Page Title="" Language="vb" AutoEventWireup="false" MasterPageFile="~/SiteContent.Master" CodeBehind="SuppPASIDeliveryConf.aspx.vb" Inherits="AffiliateSystem.SuppPASIDeliveryConf" %>
<%@ Register Assembly="DevExpress.Web.v14.1, Version=14.1.6.0, Culture=neutral, PublicKeyToken=b88d1754d700e49a" namespace="DevExpress.Web.ASPxGridView.Export" TagPrefix="dx" %>
<%@ Register Assembly="DevExpress.Web.v14.1, Version=14.1.6.0, Culture=neutral, PublicKeyToken=b88d1754d700e49a" namespace="DevExpress.Web.ASPxGridView" TagPrefix="dx" %>
<%@ Register assembly="DevExpress.Web.v14.1, Version=14.1.6.0, Culture=neutral, PublicKeyToken=b88d1754d700e49a" namespace="DevExpress.Web.ASPxEditors" tagprefix="dx" %>
<%@ Register Assembly="DevExpress.Web.v14.1, Version=14.1.6.0, Culture=neutral, PublicKeyToken=b88d1754d700e49a" namespace="DevExpress.Web.ASPxGlobalEvents" TagPrefix="dx" %>
<%@ Register Assembly="DevExpress.Web.v14.1, Version=14.1.6.0, Culture=neutral, PublicKeyToken=b88d1754d700e49a" namespace="DevExpress.Web.ASPxPanel" TagPrefix="dx" %>
<%@ Register Assembly="DevExpress.Web.v14.1, Version=14.1.6.0, Culture=neutral, PublicKeyToken=b88d1754d700e49a" namespace="DevExpress.Web.ASPxRoundPanel" TagPrefix="dx" %>
<%@ Register Assembly="DevExpress.Web.v14.1, Version=14.1.6.0, Culture=neutral, PublicKeyToken=b88d1754d700e49a" Namespace="DevExpress.Web.ASPxCallback" TagPrefix="dx" %>

<%@ Register assembly="DevExpress.Web.v14.1, Version=14.1.6.0, Culture=neutral, PublicKeyToken=b88d1754d700e49a" namespace="DevExpress.Web.ASPxEditors" tagprefix="dx1" %>

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
    <table style="width:100%;">
        <tr>
            <td width="100%">
                <table style="border-width: 1pt thin thin thin; border-style: ridge; border-color:#9598A1; width:100%; height: 70px;">
                    <tr>
                        <td colspan="8" height="20">
                            <table id="Table1">
                                <tr>
                                    <td style="width:5px;"></td>
                                    <td align="left" valign="middle" height="20px" width="200px">
                                        <dx:ASPxLabel ID="ASPxLabel1" runat="server" Text="SUPPLIER PLAN DELIVERY DATE (UNTIL)"
                                            Font-Names="Tahoma" Font-Size="8pt">
                                        </dx:ASPxLabel>
                                    </td>
                                    <td align="left" valign="middle" height="20px" width="300px">
                                        <table>
                                            <tr>
                                                <td>
                                                    <dx:ASPxCheckBox ID="chkSupplierPeriod" runat="server" ClientInstanceName="chkSupplierPeriod" Text="" Checked="true">
                                                        <ClientSideEvents
                                                            CheckedChanged="function (s, e) {
                                                                if (s.GetChecked()==true) {
                                                                    dtSupplierPeriod.SetEnabled(true);
                                                                } else {
                                                                    dtSupplierPeriod.SetEnabled(false);
                                                                }
                                                                grid.PerformCallback('clear');
                                                          }"
                                                        />
                                                    </dx:ASPxCheckBox>
                                                </td>
                                                <td>
                                                    <dx:ASPxDateEdit ID="dtSupplierPeriod" runat="server" Font-Names="Tahoma" Font-Size="8pt" Width="100px"
                                                        EditFormat="Custom" EditFormatString="dd MMM yyyy" ClientInstanceName="dtSupplierPeriod">
                                                    </dx:ASPxDateEdit>
                                                </td>
                                            </tr>
                                        </table>                                         
                                    </td>                                    
                                    <td style="width:5px;"></td>
                                    <td align="left" valign="middle" height="20px" width="130px">
                                        <dx:ASPxLabel ID="ASPxLabel4" runat="server" Text="SUPPLIER CODE/NAME"
                                            Font-Names="Tahoma" Font-Size="8pt">
                                        </dx:ASPxLabel>
                                    </td>
                                    <td align="left" valign="middle" height="20px" width="150px">
                                        <table>
                                            <tr>
                                                <td>
                                                    <dx:ASPxComboBox ID="cboSupplier" runat="server" ClientInstanceName="cboSupplier"
                                                        Font-Names="Tahoma" TextFormatString="{0}" Font-Size="8pt" Width="100px">
                                                        <ClientSideEvents SelectedIndexChanged="function(s, e) {
                                                            grid.PerformCallback('clear');
	                                                        txtSupplierName.SetText(cboSupplier.GetSelectedItem().GetColumnText(1));                                                            
                                                            }" />
                                                    </dx:ASPxComboBox>
                                                </td>
                                                <td>
                                                    <dx:ASPxTextBox ID="txtSupplierName" runat="server" Width="200px" Font-Names="Tahoma"
                                                        Font-Size="8pt" BackColor="#CCCCCC" ReadOnly="True" MaxLength="20"
                                                        ClientInstanceName="txtSupplierName">
                                                        <ClientSideEvents TextChanged="function(s, e) {
	                                                        lblInfo.SetText('');
                                                        }" />
                                                    </dx:ASPxTextBox>
                                                </td>
                                            </tr>
                                        </table>                                         
                                    </td>
                                </tr>
                                <tr>
                                    <td style="width:5px;"></td>
                                    <td align="left" valign="middle" height="20px" width="200px">
                                        <dx:ASPxLabel ID="ASPxLabel6" runat="server" Text="SUPPLIER ALREADY DELIVER"
                                            Font-Names="Tahoma" Font-Size="8pt">
                                        </dx:ASPxLabel>
                                    </td>
                                    <td align="left" valign="middle" >
                                        <table>
                                            <tr>
                                                <td>
                                                    <dx:ASPxRadioButton ID="rdrSADAll" ClientInstanceName="rdrSADAll" runat="server" Text="ALL" GroupName="SupplierAlreadyDeliver" Font-Names="Tahoma" Font-Size="8pt">
                                                        <ClientSideEvents CheckedChanged="function(s, e) {
                                                            grid.PerformCallback('clear');
                                                            lblInfo.SetText('');
                                                        }" />
                                                    </dx:ASPxRadioButton>
                                                </td>
                                                <td>
                                                    <dx:ASPxRadioButton ID="rdrSADYes" ClientInstanceName="rdrSADYes" runat="server" Text="YES" GroupName="SupplierAlreadyDeliver" Font-Names="Tahoma" Font-Size="8pt">
                                                        <ClientSideEvents CheckedChanged="function(s, e) {
                                                            grid.PerformCallback('clear');
                                                            lblInfo.SetText('');
                                                        }" />
                                                    </dx:ASPxRadioButton>
                                                </td>
                                                <td>
                                                    <dx:ASPxRadioButton ID="rdrSADNo" ClientInstanceName="rdrSADNo" runat="server" Text="NO" GroupName="SupplierAlreadyDeliver" Font-Names="Tahoma" Font-Size="8pt">
                                                        <ClientSideEvents CheckedChanged="function(s, e) {
                                                            grid.PerformCallback('clear');
                                                            lblInfo.SetText('');
                                                        }" />
                                                    </dx:ASPxRadioButton>
                                                </td>
                                            </tr>
                                        </table> 
                                    </td>                                    
                                    <td style="width:5px;"></td>
                                    <td align="left" valign="middle" height="20px" width="130px">
                                        <dx:ASPxLabel ID="ASPxLabel7" runat="server" Text="PART CODE/NAME"
                                            Font-Names="Tahoma" Font-Size="8pt">
                                        </dx:ASPxLabel>
                                    </td>
                                    <td align="left" valign="middle" height="20px" width="150px">
                                         <table>
                                            <tr>
                                                <td>
                                                    <dx:ASPxComboBox ID="cboPart" runat="server" ClientInstanceName="cboPart"
                                                        Font-Names="Tahoma" TextFormatString="{0}" Font-Size="8pt" Width="100px">
                                                        <ClientSideEvents SelectedIndexChanged="function(s, e) {
	                                                        txtPartName.SetText(cboPart.GetSelectedItem().GetColumnText(1));
                                                            grid.PerformCallback('clear');
                                                            }" />
                                                    </dx:ASPxComboBox>
                                                </td>
                                                <td>
                                                    <dx:ASPxTextBox ID="txtPartName" runat="server" Width="200px" Font-Names="Tahoma"
                                                        Font-Size="8pt" BackColor="#CCCCCC" ReadOnly="True" MaxLength="20"
                                                        ClientInstanceName="txtPartName">
                                                    </dx:ASPxTextBox>
                                                </td>
                                            </tr>
                                        </table> 
                                    </td>
                                </tr>
                                <tr>
                                    <td style="width:5px;"></td>
                                    <td align="left" valign="middle" height="20px" width="200px">
                                        <dx:ASPxLabel ID="ASPxLabel8" runat="server" Text="PASI ALREADY DELIVER"
                                            Font-Names="Tahoma" Font-Size="8pt">
                                        </dx:ASPxLabel>
                                    </td>
                                    <td align="left" valign="middle" height="20px" width="300px">
                                        <table>
                                            <tr>
                                                <td>
                                                    <dx:ASPxRadioButton ID="rdrPADAll" ClientInstanceName="rdrPADAll" runat="server" Text="ALL" GroupName="PASIAlreadyDeliver" Font-Names="Tahoma" Font-Size="8pt">
                                                        <ClientSideEvents CheckedChanged="function(s, e) {
                                                            grid.PerformCallback('clear');
                                                            lblInfo.SetText('');
                                                        }" />
                                                    </dx:ASPxRadioButton>
                                                </td>
                                                <td>
                                                    <dx:ASPxRadioButton ID="rdrPADYes" ClientInstanceName="rdrPADYes" runat="server" Text="YES" GroupName="PASIAlreadyDeliver" Font-Names="Tahoma" Font-Size="8pt">
                                                        <ClientSideEvents CheckedChanged="function(s, e) {
                                                            grid.PerformCallback('clear');
                                                            lblInfo.SetText('');
                                                        }" />
                                                    </dx:ASPxRadioButton>
                                                </td>
                                                <td>
                                                    <dx:ASPxRadioButton ID="rdrPADNo" ClientInstanceName="rdrPADNo" runat="server" Text="NO" GroupName="PASIAlreadyDeliver" Font-Names="Tahoma" Font-Size="8pt">
                                                        <ClientSideEvents CheckedChanged="function(s, e) {
                                                            grid.PerformCallback('clear');
                                                            lblInfo.SetText('');
                                                        }" />
                                                    </dx:ASPxRadioButton>
                                                </td>
                                            </tr>
                                        </table> 
                                    </td>                                    
                                    <td style="width:5px;"></td>
                                    <td align="left" valign="middle" height="20px" width="130px">
                                        <dx:ASPxLabel ID="ASPxLabel9" runat="server" Text="PO NO."
                                            Font-Names="Tahoma" Font-Size="8pt">
                                        </dx:ASPxLabel>
                                    </td>
                                    <td align="left" valign="middle" height="20px" width="150px">                                        
                                        <dx:ASPxTextBox ID="txtPONo" runat="server" Width="130px" Font-Names="Tahoma"
                                            Font-Size="8pt" BackColor="White" MaxLength="20"
                                            ClientInstanceName="txtPONo">
                                            <ClientSideEvents TextChanged="function(s, e) {
                                                grid.PerformCallback('clear');
	                                            lblInfo.SetText('');
                                            }" />
                                        </dx:ASPxTextBox>
                                    </td>
                                </tr>
                                <tr>
                                    <td style="width:5px;"></td>
                                    <td align="left" valign="middle" height="20px" width="200px">
                                        <dx:ASPxLabel ID="ASPxLabel10" runat="server" Text="REMAINING RECEIVING QTY"
                                            Font-Names="Tahoma" Font-Size="8pt">
                                        </dx:ASPxLabel>
                                    </td>
                                    <td align="left" valign="middle" height="20px" width="300px">
                                        <table>
                                            <tr>
                                                <td>
                                                    <dx:ASPxRadioButton ID="rdrRRQAll" ClientInstanceName="rdrRRQAll" runat="server" Text="ALL" GroupName="RemainingReceivingQty" Font-Names="Tahoma" Font-Size="8pt">
                                                        <ClientSideEvents CheckedChanged="function(s, e) {
                                                            grid.PerformCallback('clear');
                                                            lblInfo.SetText('');
                                                        }" />
                                                    </dx:ASPxRadioButton>
                                                </td>
                                                <td>
                                                    <dx:ASPxRadioButton ID="rdrRRQYes" ClientInstanceName="rdrRRQYes" runat="server" Text="YES" GroupName="RemainingReceivingQty" Font-Names="Tahoma" Font-Size="8pt">
                                                        <ClientSideEvents CheckedChanged="function(s, e) {
                                                            grid.PerformCallback('clear');
                                                            lblInfo.SetText('');
                                                        }" />
                                                    </dx:ASPxRadioButton>
                                                </td>
                                                <td>
                                                    <dx:ASPxRadioButton ID="rdrRRQNo" ClientInstanceName="rdrRRQNo" runat="server" Text="NO" GroupName="RemainingReceivingQty" Font-Names="Tahoma" Font-Size="8pt">
                                                        <ClientSideEvents CheckedChanged="function(s, e) {
                                                            grid.PerformCallback('clear');
                                                            lblInfo.SetText('');
                                                        }" />
                                                    </dx:ASPxRadioButton>
                                                </td>
                                            </tr>
                                        </table>
                                    </td>                                    
                                    <td style="width:5px;"></td>
                                    <td align="left" valign="middle" height="20px" width="130px">
                                        <dx:ASPxLabel ID="ASPxLabel11" runat="server" Text="SUPPLIER DELIVERY"
                                            Font-Names="Tahoma" Font-Size="8pt">
                                        </dx:ASPxLabel>
                                    </td>
                                    <td align="left" valign="middle" height="20px" width="150px">
                                        <table>
                                            <tr>
                                                <td>
                                                    <dx:ASPxRadioButton ID="rdrSDAll" ClientInstanceName="rdrSDAll" runat="server" Text="ALL" GroupName="SupplierDelivery" Font-Names="Tahoma" Font-Size="8pt">
                                                        <ClientSideEvents CheckedChanged="function(s, e) {
                                                            grid.PerformCallback('clear');
                                                            lblInfo.SetText('');
                                                        }" />
                                                    </dx:ASPxRadioButton>
                                                </td>
                                                <td>
                                                    <dx:ASPxRadioButton ID="rdrSDDirect" ClientInstanceName="rdrSDDirect" runat="server" Text="DIRECT TO AFFILATE" GroupName="SupplierDelivery" Font-Names="Tahoma" Font-Size="8pt">
                                                        <ClientSideEvents CheckedChanged="function(s, e) {
                                                            grid.PerformCallback('clear');
                                                            lblInfo.SetText('');
                                                        }" />
                                                    </dx:ASPxRadioButton>
                                                </td>
                                                <td>
                                                    <dx:ASPxRadioButton ID="rdrSDPasi" ClientInstanceName="rdrSDPasi" runat="server" Text="VIA PASI" GroupName="SupplierDelivery" Font-Names="Tahoma" Font-Size="8pt">
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
                                <tr>
                                    <td style="width:5px;"></td>
                                    <td align="left" valign="middle" height="20px" width="200px">
                                        <dx:ASPxLabel ID="ASPxLabel12" runat="server" Text="SUPPLIER / PASI SURAT JALAN NO."
                                            Font-Names="Tahoma" Font-Size="8pt">
                                        </dx:ASPxLabel>
                                    </td>
                                    <td align="left" valign="middle" height="20px" width="300px">
                                        <dx:ASPxTextBox ID="txtSupplierSJNo" runat="server" Width="130px" Font-Names="Tahoma"
                                            Font-Size="8pt" BackColor="White" MaxLength="20"
                                            ClientInstanceName="txtSupplierSJNo">
                                            <ClientSideEvents TextChanged="function(s, e) {
                                                grid.PerformCallback('clear');
	                                            lblInfo.SetText('');
                                            }" />
                                        </dx:ASPxTextBox>                                        
                                    </td>                                    
                                    <td style="width:5px;"></td>
                                    <td align="left" valign="middle" height="20px" width="130px">
                                        <dx:ASPxLabel ID="ASPxLabel3" runat="server" Text="PO KANBAN"
                                            Font-Names="Tahoma" Font-Size="8pt">
                                        </dx:ASPxLabel>
                                    </td>
                                    <td align="left" valign="middle" height="20px" width="150px">
                                        <table>
                                            <tr>
                                                <td>
                                                    <dx:ASPxRadioButton ID="rdrPOKAll" ClientInstanceName="rdrPOKAll" runat="server" Text="ALL" GroupName="POKanban" Font-Names="Tahoma" Font-Size="8pt">
                                                        <ClientSideEvents CheckedChanged="function(s, e) {
                                                        grid.PerformCallback('clear');
                                                        lblInfo.SetText('');
                                                        }" />
                                                    </dx:ASPxRadioButton>
                                                </td>
                                                <td>
                                                    <dx:ASPxRadioButton ID="rdrPOKYes" ClientInstanceName="rdrPOKYes" runat="server" Text="YES" GroupName="POKanban" Font-Names="Tahoma" Font-Size="8pt">
                                                        <ClientSideEvents CheckedChanged="function(s, e) {
                                                        grid.PerformCallback('clear');
                                                        lblInfo.SetText('');
                                                        }" />
                                                    </dx:ASPxRadioButton>
                                                </td>
                                                <td>
                                                    <dx:ASPxRadioButton ID="rdrPOKNo" ClientInstanceName="rdrPOKNo" runat="server" Text="NO" GroupName="POKanban" Font-Names="Tahoma" Font-Size="8pt">
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
                                <tr>
                                    <td style="width:5px;"></td>
                                    <td align="left" valign="middle" height="20px" width="200px">
                                        <dx:ASPxLabel ID="ASPxLabel14" runat="server" Text="PASI DELIVERY DATE"
                                            Font-Names="Tahoma" Font-Size="8pt">
                                        </dx:ASPxLabel>
                                    </td>
                                    <td align="left" valign="middle" height="20px" width="300px">
                                        <table>
                                            <tr>
                                                <td>
                                                    <dx:ASPxCheckBox ID="chkRecDate" runat="server" ClientInstanceName="chkRecDate" Text="" Checked="true">
                                                        <ClientSideEvents
                                                            CheckedChanged="function (s, e) {                                                                                                                                
                                                                if (s.GetChecked()==true) {
                                                                    dtRecDateFrom.SetEnabled(true);
                                                                    dtRecDateTo.SetEnabled(true);
                                                                } else {
                                                                    dtRecDateFrom.SetEnabled(false);
                                                                    dtRecDateTo.SetEnabled(false);
                                                                }
                                                                grid.PerformCallback('clear');
                                                          }"
                                                        />
                                                    </dx:ASPxCheckBox>
                                                </td>
                                                <td>                                                    
                                                    <dx:ASPxDateEdit ID="dtRecDateFrom" runat="server" Font-Names="Tahoma" Font-Size="8pt" Width="100px"
                                                        EditFormat="Custom" EditFormatString="dd MMM yyyy" ClientInstanceName="dtRecDateFrom">
                                                    </dx:ASPxDateEdit>
                                                </td>
                                                <td> ~
                                                </td>
                                                <td>
                                                    <dx:ASPxDateEdit ID="dtRecDateTo" runat="server" Font-Names="Tahoma" Font-Size="8pt" Width="100px"
                                                        EditFormat="Custom" EditFormatString="dd MMM yyyy" ClientInstanceName="dtRecDateTo">
                                                    </dx:ASPxDateEdit>
                                                </td>
                                            </tr>
                                        </table> 
                                    </td>                                    
                                    <td style="width:5px;"></td>
                                    <td align="left" valign="middle" height="20px" width="130px">
                                        <dx:ASPxLabel ID="ASPxLabel15" runat="server" Text="MANUAL CLOSE PO"
                                            Font-Names="Tahoma" Font-Size="8pt">
                                        </dx:ASPxLabel>
                                    </td>
                                    <td align="left" valign="middle" height="20px" width="150px">
                                        <table>
                                            <tr>
                                                <td>
                                                    <dx:ASPxRadioButton ID="rdrMCPAll" ClientInstanceName="rdrMCPAll" runat="server" Text="ALL" GroupName="ManualClosePO" Font-Names="Tahoma" Font-Size="8pt">
                                                        <ClientSideEvents CheckedChanged="function(s, e) {
                                                        grid.PerformCallback('clear');
                                                        lblInfo.SetText('');
                                                        }" />
                                                    </dx:ASPxRadioButton>
                                                </td>
                                                <td>
                                                    <dx:ASPxRadioButton ID="rdrMCPYes" ClientInstanceName="rdrMCPYes" runat="server" Text="YES" GroupName="ManualClosePO" Font-Names="Tahoma" Font-Size="8pt">
                                                        <ClientSideEvents CheckedChanged="function(s, e) {
                                                        grid.PerformCallback('clear');
                                                        lblInfo.SetText('');
                                                        }" />
                                                    </dx:ASPxRadioButton>
                                                </td>
                                                <td>
                                                    <dx:ASPxRadioButton ID="rdrMCPNo" ClientInstanceName="rdrMCPNo" runat="server" Text="NO" GroupName="ManualClosePO" Font-Names="Tahoma" Font-Size="8pt">
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
                                <tr>
                                    <td style="width:5px;"></td>
                                    <td align="left" valign="middle" style="height:20px;">
                                        &nbsp;</td>
                                    <td align="left" valign="middle" >
                                        &nbsp;</td>                                    
                                    <td style="width:5px;"></td>
                                    <td align="left" valign="middle" >
                                        <dx:ASPxLabel ID="ASPxLabel2" runat="server" Text="GOOD RECEIVING SENT"
                                            Font-Names="Tahoma" Font-Size="8pt">
                                        </dx:ASPxLabel>
                                    </td>
                                    <td align="left" valign="middle" >
                                        <table>
                                            <tr>
                                                <td>
                                                    <dx:ASPxRadioButton ID="rdrGRSAll" ClientInstanceName="rdrGRSAll" runat="server" Text="ALL" GroupName="GoodReceivingSent" Font-Names="Tahoma" Font-Size="8pt">
                                                        <ClientSideEvents CheckedChanged="function(s, e) {
                                                        grid.PerformCallback('clear');
                                                        lblInfo.SetText('');
                                                        }" />
                                                    </dx:ASPxRadioButton>
                                                </td>
                                                <td>
                                                    <dx:ASPxRadioButton ID="rdrGRSYes" ClientInstanceName="rdrGRSYes" runat="server" Text="YES" GroupName="GoodReceivingSent" Font-Names="Tahoma" Font-Size="8pt">
                                                        <ClientSideEvents CheckedChanged="function(s, e) {
                                                        grid.PerformCallback('clear');
                                                        lblInfo.SetText('');
                                                        }" />
                                                    </dx:ASPxRadioButton>
                                                </td>
                                                <td>
                                                    <dx:ASPxRadioButton ID="rdrGRSNo" ClientInstanceName="rdrGRSNo" runat="server" Text="NO" GroupName="GoodReceivingSent" Font-Names="Tahoma" Font-Size="8pt">
                                                        <ClientSideEvents CheckedChanged="function(s, e) {
                                                        grid.PerformCallback('clear');
                                                        lblInfo.SetText('');
                                                        }" />
                                                    </dx:ASPxRadioButton>
                                                </td>
                                                <td style="width:50px;"></td>
                                                <td>
                                                    <dx:ASPxButton ID="btnSearch" runat="server" Text="SEARCH"
                                                        Font-Names="Tahoma" Width="85px" AutoPostBack="False" Font-Size="8pt" >
                                                        <ClientSideEvents Click="function(s, e) {                                         
                                                            grid.PerformCallback('load');
                                                            lblInfo.SetText('');
                                                        }" />
                                                    </dx:ASPxButton>
                                                </td>
                                                <td>
                                                    <dx:ASPxButton ID="btnClear" runat="server" Text="CLEAR"
                                                        Font-Names="Tahoma" Width="85px" AutoPostBack="False" Font-Size="8pt" >
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
                                                <td></td>
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

    <div style="height:1px;"></div>

    <table style="width:100%; height: 15px;">
        <tr>
            <td colspan="8" height="15">
            <%--error message--%>
                <table id="tblMsg" style="border-width: thin; border-style: inset hidden ridge hidden; border-color:#9598A1; width:100%; height: 15px;">
                    <tr>
                        <td align="left" valign="middle" height="15px">
                            <dx:ASPxLabel ID="lblInfo" runat="server" Text="[lblinfo]" Font-Names="Tahoma" 
                                ClientInstanceName="lblInfo" Font-Bold="True" Font-Italic="True" 
                                Font-Size="8pt">
                            </dx:ASPxLabel>
                        </td>
                    </tr>         
                </table>
            </td>            
        </tr>
    </table>
        
    <div style="height:1px;"></div>

    <table style="width:100%;">
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
                <dx:ASPxGridView ID="grid" runat="server" Width="100%"
                    Font-Names="Tahoma" KeyFieldName="PONo;KanbanNo;PartNo;SupplierCode"
                    AutoGenerateColumns="False"
                    ClientInstanceName="grid" Font-Size="8pt">
                    <ClientSideEvents Init="OnInit" FocusedRowChanged="function(s, e) {
	                    OnGridFocusedRowChanged();
                    }" CallbackError="function(s, e) {e.handled = true;}" BatchEditEndEditing="OnBatchEditEndEditing" EndCallback="function(s, e) {
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
                        <dx:GridViewDataHyperLinkColumn Caption=" " FieldName="url" 
                            VisibleIndex="0" Width="70px">
                            <PropertiesHyperLinkEdit TextField="detail"> </PropertiesHyperLinkEdit>
                            <HeaderStyle HorizontalAlign="Center" Wrap="True" />
                            <CellStyle Font-Names="Tahoma" Font-Size="8pt">
                            </CellStyle>
                        </dx:GridViewDataHyperLinkColumn>
                        <dx:GridViewDataTextColumn VisibleIndex="1" Caption="NO" FieldName="ColNo" Width="30px"
                            HeaderStyle-HorizontalAlign="Center">
                            <HeaderStyle HorizontalAlign="Center" Wrap="True"></HeaderStyle>
                            <CellStyle Font-Names="Tahoma" Font-Size="8pt" HorizontalAlign="Center">
                            </CellStyle>                            
                        </dx:GridViewDataTextColumn>
                        <dx:GridViewDataTextColumn VisibleIndex="2" Caption="PERIOD" FieldName="Period" Width="70px"
                            HeaderStyle-HorizontalAlign="Center">
                            <HeaderStyle HorizontalAlign="Center" Wrap="True"></HeaderStyle>
                            <CellStyle Font-Names="Tahoma" Font-Size="8pt" HorizontalAlign="Center">
                            </CellStyle>
                        </dx:GridViewDataTextColumn>
                        <dx:GridViewDataTextColumn VisibleIndex="3" Caption="PO NO." FieldName="PONo" Width="100px"
                            HeaderStyle-HorizontalAlign="Center">
                            <HeaderStyle HorizontalAlign="Center" Wrap="True"></HeaderStyle>
                            <CellStyle Font-Names="Tahoma" Font-Size="8pt">
                            </CellStyle>
                        </dx:GridViewDataTextColumn>

                        <dx:GridViewDataTextColumn VisibleIndex="6" Caption="DELIVERY LOCATION CODE" 
                            FieldName="DeliveryLocationCode" Width="80px"
                            HeaderStyle-HorizontalAlign="Center">
                            <HeaderStyle HorizontalAlign="Center" Wrap="True"></HeaderStyle>
                            <CellStyle Font-Names="Tahoma" Font-Size="8pt">
                            </CellStyle>
                        </dx:GridViewDataTextColumn>
                        <dx:GridViewDataTextColumn VisibleIndex="7" Caption="DELIVERY LOCATION NAME" 
                            FieldName="DeliveryLocationName" Width="0px"
                            HeaderStyle-HorizontalAlign="Center">
                            <HeaderStyle HorizontalAlign="Center" Wrap="True"></HeaderStyle>
                            <CellStyle Font-Names="Tahoma" Font-Size="8pt">
                            </CellStyle>
                        </dx:GridViewDataTextColumn>
                        <dx:GridViewDataTextColumn VisibleIndex="8" Caption="SUPPLIER CODE" 
                            FieldName="SupplierCode" Width="70px"
                            HeaderStyle-HorizontalAlign="Center">
                            <HeaderStyle HorizontalAlign="Center" Wrap="True"></HeaderStyle>
                            <CellStyle Font-Names="Tahoma" Font-Size="8pt">
                            </CellStyle>
                        </dx:GridViewDataTextColumn>
                        <dx:GridViewDataTextColumn VisibleIndex="9" Caption="SUPPLIER NAME" 
                            FieldName="SupplierName" Width="0px"
                            HeaderStyle-HorizontalAlign="Center">
                            <HeaderStyle HorizontalAlign="Center" Wrap="True"></HeaderStyle>
                            <CellStyle Font-Names="Tahoma" Font-Size="8pt">
                            </CellStyle>
                        </dx:GridViewDataTextColumn>
                        <dx:GridViewDataTextColumn VisibleIndex="10" Caption="PO KANBAN" 
                            FieldName="POKanban" Width="60px"
                            HeaderStyle-HorizontalAlign="Center">
                            <HeaderStyle HorizontalAlign="Center" Wrap="True"></HeaderStyle>
                            <CellStyle Font-Names="Tahoma" Font-Size="8pt" HorizontalAlign="Center" >
                            </CellStyle>
                        </dx:GridViewDataTextColumn>
                        <dx:GridViewDataTextColumn VisibleIndex="11" Caption="KANBAN NO." 
                            FieldName="KanbanNo" Width="80px"
                            HeaderStyle-HorizontalAlign="Center">
                            <HeaderStyle HorizontalAlign="Center" Wrap="True"></HeaderStyle>
                            <CellStyle Font-Names="Tahoma" Font-Size="8pt">
                            </CellStyle>
                        </dx:GridViewDataTextColumn>
                        <dx:GridViewDataTextColumn VisibleIndex="12" 
                            Caption="SUPPLIER PLAN DELIVERY DATE" FieldName="SupplierPlanDeliveryDate" Width="110px"
                            HeaderStyle-HorizontalAlign="Center">
                            <HeaderStyle HorizontalAlign="Center" Wrap="True"></HeaderStyle>
                            <CellStyle Font-Names="Tahoma" Font-Size="8pt" HorizontalAlign="Center">
                            </CellStyle>
                        </dx:GridViewDataTextColumn>
                        <dx:GridViewDataTextColumn VisibleIndex="13" Caption="SUPPLIER DELIVERY DATE" 
                            FieldName="SupplierDeliveryDate" Width="0px"
                            HeaderStyle-HorizontalAlign="Center">
                            <HeaderStyle HorizontalAlign="Center" Wrap="True"></HeaderStyle>
                            <CellStyle Font-Names="Tahoma" Font-Size="8pt" HorizontalAlign="Center">
                            </CellStyle>
                        </dx:GridViewDataTextColumn>
                        <dx:GridViewDataTextColumn VisibleIndex="14" Caption="SUPPLIER SURAT JALAN NO." 
                            FieldName="SupplierSJNo" Width="0px"
                            HeaderStyle-HorizontalAlign="Center">
                            <HeaderStyle HorizontalAlign="Center" Wrap="True"></HeaderStyle>
                            <CellStyle Font-Names="Tahoma" Font-Size="8pt">
                            </CellStyle>
                        </dx:GridViewDataTextColumn>
                        <dx:GridViewDataTextColumn VisibleIndex="15" Caption="PASI DELIVERY DATE" 
                            FieldName="PASIDeliveryDate" Width="80px"
                            HeaderStyle-HorizontalAlign="Center">
                            <HeaderStyle HorizontalAlign="Center" Wrap="True"></HeaderStyle>
                            <CellStyle Font-Names="Tahoma" Font-Size="8pt" HorizontalAlign="Center">
                            </CellStyle>
                        </dx:GridViewDataTextColumn>
                        <dx:GridViewDataTextColumn VisibleIndex="16" Caption="PASI SURAT JALAN NO." 
                            FieldName="PASISJNo" Width="100px"
                            HeaderStyle-HorizontalAlign="Center">
                            <HeaderStyle HorizontalAlign="Center" Wrap="True"></HeaderStyle>
                            <CellStyle Font-Names="Tahoma" Font-Size="8pt">
                            </CellStyle>
                        </dx:GridViewDataTextColumn>
                        <dx:GridViewDataTextColumn VisibleIndex="17" Caption="PART NO." 
                            FieldName="PartNo" Width="0px"
                            HeaderStyle-HorizontalAlign="Center">
                            <HeaderStyle HorizontalAlign="Center" Wrap="True"></HeaderStyle>
                            <CellStyle Font-Names="Tahoma" Font-Size="8pt">
                            </CellStyle>
                        </dx:GridViewDataTextColumn>
                        <dx:GridViewDataTextColumn VisibleIndex="18" Caption="PART NAME" 
                            FieldName="PartName" Width="0px"
                            HeaderStyle-HorizontalAlign="Center">
                            <HeaderStyle HorizontalAlign="Center" Wrap="True"></HeaderStyle>
                            <CellStyle Font-Names="Tahoma" Font-Size="8pt">
                            </CellStyle>
                        </dx:GridViewDataTextColumn>
                        <dx:GridViewDataTextColumn VisibleIndex="19" Caption="UOM" FieldName="UOM" Width="0px"
                            HeaderStyle-HorizontalAlign="Center">
                            <HeaderStyle HorizontalAlign="Center" Wrap="True"></HeaderStyle>
                            <CellStyle Font-Names="Tahoma" Font-Size="8pt" HorizontalAlign="Center">
                            </CellStyle>
                        </dx:GridViewDataTextColumn>
                        <dx:GridViewDataTextColumn VisibleIndex="20" Caption="SUPPLIER DELIVERY QTY" 
                            FieldName="SupplierDeliveryQty" Width="0px"
                            HeaderStyle-HorizontalAlign="Center">
                            <HeaderStyle HorizontalAlign="Center" Wrap="True"></HeaderStyle>
                            <CellStyle Font-Names="Tahoma" Font-Size="8pt" HorizontalAlign="Right">
                            </CellStyle>                            
                        </dx:GridViewDataTextColumn>
                        <dx:GridViewDataTextColumn VisibleIndex="21" Caption="PASI GOOD RECEIVING QTY" 
                            FieldName="PASIGoodReceivingQty" Width="0px"
                            HeaderStyle-HorizontalAlign="Center">
                            <HeaderStyle HorizontalAlign="Center" Wrap="True"></HeaderStyle>
                            <CellStyle Font-Names="Tahoma" Font-Size="8pt" HorizontalAlign="Right">
                            </CellStyle>                            
                        </dx:GridViewDataTextColumn>
                        <dx:GridViewDataTextColumn VisibleIndex="22" 
                            Caption="PASI DEFECT RECEIVING QTY" FieldName="PASIDefectQty" Width="0px"
                            HeaderStyle-HorizontalAlign="Center">
                            <HeaderStyle HorizontalAlign="Center" Wrap="True"></HeaderStyle>
                            <CellStyle Font-Names="Tahoma" Font-Size="8pt" HorizontalAlign="Right">
                            </CellStyle>
                            
                        </dx:GridViewDataTextColumn>
                        <dx:GridViewDataTextColumn VisibleIndex="23" Caption="PASI DELIVERY QTY" 
                            FieldName="PASIDeliveryQty" Width="0px"
                            HeaderStyle-HorizontalAlign="Center">
                            <HeaderStyle HorizontalAlign="Center" Wrap="True"></HeaderStyle>
                            <CellStyle Font-Names="Tahoma" Font-Size="8pt" HorizontalAlign="Right">
                            </CellStyle>
                            
                        </dx:GridViewDataTextColumn>
                        <dx:GridViewDataTextColumn VisibleIndex="24" Caption="GOOD RECEIVING QTY" 
                            FieldName="GoodReceivingQty" Width="0px"
                            HeaderStyle-HorizontalAlign="Center">
                            <HeaderStyle HorizontalAlign="Center" Wrap="True"></HeaderStyle>
                            <CellStyle Font-Names="Tahoma" Font-Size="8pt" HorizontalAlign="Right">
                            </CellStyle>
                            
                        </dx:GridViewDataTextColumn>
                        <dx:GridViewDataTextColumn VisibleIndex="25" Caption="DEFECT RECEIVING QTY" 
                            FieldName="DefectReceivingQty" Width="0px"
                            HeaderStyle-HorizontalAlign="Center">
                            <HeaderStyle HorizontalAlign="Center" Wrap="True"></HeaderStyle>
                            <CellStyle Font-Names="Tahoma" Font-Size="8pt" HorizontalAlign="Right">
                            </CellStyle>
                            
                        </dx:GridViewDataTextColumn>
                        <dx:GridViewDataTextColumn VisibleIndex="26" Caption="REMAINING RECEIVING QTY" 
                            FieldName="RemainingReceivingQty" Width="0px"
                            HeaderStyle-HorizontalAlign="Center">
                            <HeaderStyle HorizontalAlign="Center" Wrap="True"></HeaderStyle>
                            <CellStyle Font-Names="Tahoma" Font-Size="8pt" HorizontalAlign="Right">
                            </CellStyle>
                            
                        </dx:GridViewDataTextColumn>
                        <dx:GridViewDataTextColumn VisibleIndex="27" Caption="RECEIVED DATE" 
                            FieldName="ReceivedDate" Width="140px"
                            HeaderStyle-HorizontalAlign="Center">
                            <HeaderStyle HorizontalAlign="Center" Wrap="True"></HeaderStyle>
                            <CellStyle Font-Names="Tahoma" Font-Size="8pt" HorizontalAlign="Center">
                            </CellStyle>
                        </dx:GridViewDataTextColumn>
                        <dx:GridViewDataTextColumn VisibleIndex="28" Caption="RECEIVED BY" 
                            FieldName="ReceivedBy" Width="100px"
                            HeaderStyle-HorizontalAlign="Center">
                            <HeaderStyle HorizontalAlign="Center" Wrap="True"></HeaderStyle>
                            <CellStyle Font-Names="Tahoma" Font-Size="8pt">
                            </CellStyle>
                        </dx:GridViewDataTextColumn>
                        <dx:GridViewDataTextColumn VisibleIndex="29" Caption="GOOD RECEIVING SENT" 
                            FieldName="GoodReceivingSent" Width="80px"
                            HeaderStyle-HorizontalAlign="Center">
                            <HeaderStyle HorizontalAlign="Center" Wrap="True"></HeaderStyle>
                            <CellStyle Font-Names="Tahoma" Font-Size="8pt" HorizontalAlign="Center">
                            </CellStyle>
                        </dx:GridViewDataTextColumn>
                        <dx:GridViewDataTextColumn VisibleIndex="30" Caption="SortPONo" 
                            FieldName="SortPONo" Width="0px"
                            HeaderStyle-HorizontalAlign="Center">
                            <HeaderStyle HorizontalAlign="Center"></HeaderStyle>
                            <CellStyle Font-Names="Tahoma" Font-Size="8pt">
                            </CellStyle>
                        </dx:GridViewDataTextColumn>                        
                        <dx:GridViewDataTextColumn VisibleIndex="31" Caption="SortKanbanNo" 
                            FieldName="SortKanbanNo" Width="0px"
                            HeaderStyle-HorizontalAlign="Center">
                            <HeaderStyle HorizontalAlign="Center"></HeaderStyle>
                            <CellStyle Font-Names="Tahoma" Font-Size="8pt">
                            </CellStyle>
                        </dx:GridViewDataTextColumn>
                        <dx:GridViewDataTextColumn VisibleIndex="32" Caption="SortHeader" 
                            FieldName="SortHeader" Width="0px"
                            HeaderStyle-HorizontalAlign="Center">
                            <HeaderStyle HorizontalAlign="Center"></HeaderStyle>
                            <CellStyle Font-Names="Tahoma" Font-Size="8pt">
                            </CellStyle>
                        </dx:GridViewDataTextColumn>                        
                        <dx:GridViewDataTextColumn VisibleIndex="33" Caption="DeliveryByPASICls" 
                            FieldName="DeliveryByPASICls" Width="0px"
                            HeaderStyle-HorizontalAlign="Center">
                            <HeaderStyle HorizontalAlign="Center"></HeaderStyle>
                            <CellStyle Font-Names="Tahoma" Font-Size="8pt">
                            </CellStyle>
                        </dx:GridViewDataTextColumn>
                    </Columns>
                    <SettingsBehavior AllowSelectByRowClick="True" AllowSort="False"
                        ColumnResizeMode="Control" EnableRowHotTrack="True" />
                    <SettingsPager PageSize="16" Position="Top">
                        <Summary Text="Page {0} of {1} [{2} record(s)]"
                                  AllPagesText="Page {0} of {1} " />
                        <Summary AllPagesText="Page {0} of {1} " Text="Page {0} of {1} [{2} record(s)]"></Summary>
                    </SettingsPager>
                    <%--<SettingsEditing Mode="Batch" NewItemRowPosition="Bottom">
                        <BatchEditSettings ShowConfirmOnLosingChanges="False" />
                    </SettingsEditing>--%>
                    <Settings ShowHorizontalScrollBar="True" ShowVerticalScrollBar="True"
                        ShowGroupButtons="False" ShowStatusBar="Hidden"
                        VerticalScrollableHeight="190" />
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
    
    <div style="height:8px;"></div>
    
    <table id="button" style=" width:100%;">
        <tr>
            <td valign="top" align="left">
                <dx:ASPxButton ID="btnSubMenu" runat="server" Text="SUB MENU"
                    Font-Names="Tahoma" Width="85px" Font-Size="8pt">
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
            <%--<td valign="top" align="right" style="width: 50px;">
                <dx:ASPxButton ID="btnUpload" runat="server" Text="UPLOAD"
                    Font-Names="Tahoma"
                    Width="90px" AutoPostBack="False" Font-Size="8pt" Enabled="False" 
                    Visible="False">
                </dx:ASPxButton>
            </td>
            <td valign="top" align="right" style="width: 50px;">
                <dx:ASPxButton ID="btnDownload" runat="server" Text="DOWNLOAD"
                    Font-Names="Tahoma"
                    Width="90px" AutoPostBack="False" Font-Size="8pt" Enabled="False" 
                    Visible="False">
                </dx:ASPxButton>
            </td>--%>
            <td valign="top" align="right" style="width: 50px;">
                
            </td>            
            <td align="right" style="width:80px;">                                   
                <%--<dx:ASPxButton ID="btnDelete" runat="server" Text="DELETE"                              
                    Font-Names="Tahoma" Width="80px" AutoPostBack="False" Font-Size="8pt">
                    <ClientSideEvents Click="function(s, e) {
                        up_delete();
                    }" />
                </dx:ASPxButton>--%>
            </td>
            <td valign="top" align="right" style="width: 50px;">
                <%--<dx:ASPxButton ID="btnSubmit" runat="server" Text="SUBMIT"
                    Font-Names="Tahoma"
                    Width="90px" AutoPostBack="False" Font-Size="8pt">
                    <ClientSideEvents Click="function(s, e) {
                        validasubmit();
                        grid.UpdateEdit();
                        grid.PerformCallback('load');
                    }" />
                </dx:ASPxButton>--%>
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

