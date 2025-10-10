<%@ Page Title="" Language="vb" AutoEventWireup="false" MasterPageFile="~/SiteContent.Master" CodeBehind="RemainingPOExport.aspx.vb" Inherits="PASISystem.RemainingPOExport" %>
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
        .style1
        {
            width: 49px;
        }
        .style2
        {
            width: 54px;
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


    function clear() {
        
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
                                <!-- ROW 1 -->
                                <!-- ROW 2 -->
                                <tr>
                                    <td style="width:5px;"></td>
                                    <td align="left" valign="middle" >
                                        <dx:ASPxLabel ID="ASPxLabel22" runat="server" Text="PERIOD"
                                            Font-Names="Tahoma" Font-Size="8pt">
                                        </dx:ASPxLabel>
                                    </td>
                                    <td align="left" valign="middle" >
                                        <table>
                                            <tr>
                                                <td>
                                                    &nbsp;</td>
                                                <td>
                            <dx1:ASPxTimeEdit ID="dtPeriodFrom" runat="server" 
                                ClientInstanceName="dtPeriodFrom" DisplayFormatString="yyyy-MM" 
                                EditFormat="Custom" EditFormatString="yyyy-MM" Width="110px">                                              
                            </dx1:ASPxTimeEdit>
                                                </td>
                                                <td>~</td>
                                                <td>
                            <dx1:ASPxTimeEdit ID="dtPeriodTo" runat="server" 
                                ClientInstanceName="dtPeriodTo" DisplayFormatString="yyyy-MM" 
                                EditFormat="Custom" EditFormatString="yyyy-MM" Width="110px">                                              
                            </dx1:ASPxTimeEdit>
                                                </td>
                                            </tr>
                                        </table>                                         
                                    </td>

                                    <td style="width:5px;"></td>

                                    <td align="left" valign="middle" >
                                        <%--<dx:ASPxLabel ID="ASPxLabel4" runat="server" Text="PO PROGRESS"
                                            Font-Names="Tahoma" Font-Size="8pt">
                                        </dx:ASPxLabel>--%>
                                        <dx:ASPxLabel ID="ASPxLabel7" runat="server" Text="SUPPLIER CODE/NAME"
                                            Font-Names="Tahoma" Font-Size="8pt">
                                        </dx:ASPxLabel>
                                    </td>
                                    <td align="left" valign="middle" >
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
                                                    <dx:ASPxComboBox ID="cbosupplier" runat="server" ClientInstanceName="cbosupplier"
                                                        Font-Names="Tahoma" TextFormatString="{0}" Font-Size="8pt" Width="120px">
                                                        <ClientSideEvents SelectedIndexChanged="function(s, e) {
	                                                        txtsupplier.SetText(cbosupplier.GetSelectedItem().GetColumnText(1));
                                                            grid.PerformCallback('clear');
                                                            lblInfo.SetText('');
                                                            }" />
                                                    </dx:ASPxComboBox>
                                                </td>
                                                <td>
                                                    <dx:ASPxTextBox ID="txtsupplier" runat="server" Width="200px" Font-Names="Tahoma"
                                                        Font-Size="8pt" BackColor="#CCCCCC" ReadOnly="True" MaxLength="20"
                                                        ClientInstanceName="txtsupplier">
                                                    </dx:ASPxTextBox>
                                                </td>                                                
                                            </tr>
                                        </table>                                         
                                    </td>
                                </tr>
                                <!-- ROW 3 -->
                                <tr>
                                    <td style="width:5px;"></td>
                                    <td align="left" valign="middle" >
                                        <dx:ASPxLabel ID="ASPxLabel1" runat="server" Text="REMAINING FORWARDER RECEIVING QTY"
                                            Font-Names="Tahoma" Font-Size="8pt">
                                        </dx:ASPxLabel>
                                    </td>
                                    <td align="left" valign="middle" >
                                        <table>
                                            <tr>
                                                <td>
                                                    <dx:ASPxRadioButton ID="rdrRAll" ClientInstanceName="rdrRAll" runat="server" 
                                                        Text="ALL" GroupName="REMAINING" Font-Names="Tahoma" Font-Size="8pt">
                                                        <ClientSideEvents CheckedChanged="function(s, e) {
                                                            grid.PerformCallback('clear');
                                                            lblInfo.SetText('');
                                                        }" />
                                                    </dx:ASPxRadioButton>
                                                </td>
                                                <td>
                                                    <dx:ASPxRadioButton ID="rdrRYes" ClientInstanceName="rdrRYes" 
                                                        runat="server" Text="YES" GroupName="REMAINING" Font-Names="Tahoma" 
                                                        Font-Size="8pt">
                                                        <ClientSideEvents CheckedChanged="function(s, e) {
                                                            grid.PerformCallback('clear');
                                                            lblInfo.SetText('');
                                                        }" />
                                                    </dx:ASPxRadioButton>
                                                </td>
                                                <td>
                                                    <dx:ASPxRadioButton ID="rdrRNo" ClientInstanceName="rdrRNo" 
                                                        runat="server" Text="NO" GroupName="REMAINING" Font-Names="Tahoma" 
                                                        Font-Size="8pt">
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

                                    <td align="left" valign="middle" >
                                        <dx:ASPxLabel ID="ASPxLabel10" runat="server" Text="AFFILIATE CODE/NAME"
                                            Font-Names="Tahoma" Font-Size="8pt">
                                        </dx:ASPxLabel>
                                    </td>
                                    <td align="left" valign="middle" >
                                        <table>
                                            <tr>
                                                <td>
                                                    <dx:ASPxComboBox ID="cboAffiliate" runat="server" ClientInstanceName="cboAffiliate"
                                                        Font-Names="Tahoma" TextFormatString="{0}" Font-Size="8pt" Width="120px">
                                                        <ClientSideEvents SelectedIndexChanged="function(s, e) {
	                                                        txtAffiliate.SetText(cboAffiliate.GetSelectedItem().GetColumnText(1));
                                                            grid.PerformCallback('clear');
                                                            lblInfo.SetText('');
                                                            }" />
                                                    </dx:ASPxComboBox>
                                                </td>
                                                <td>
                                                    <dx:ASPxTextBox ID="txtAffiliate" runat="server" Width="200px" Font-Names="Tahoma"
                                                        Font-Size="8pt" BackColor="#CCCCCC" ReadOnly="True" MaxLength="20"
                                                        ClientInstanceName="txtAffiliate">
                                                    </dx:ASPxTextBox>
                                                </td>                                                
                                            </tr>
                                        </table>                                         
                                    </td>
                                </tr>
                                <!-- ROW 4 -->
                                <tr>
                                    <td style="width:5px;"></td>
                                    <td align="left" valign="middle" >
                                        <dx:ASPxLabel ID="ASPxLabel18" runat="server" Text="PART CODE/NAME"
                                            Font-Names="Tahoma" Font-Size="8pt">
                                        </dx:ASPxLabel>
                                    </td>
                                    <td align="left" valign="middle" >
                                        <table>
                                            <tr>
                                                <td>
                                                    <dx:ASPxComboBox ID="cboPart" runat="server" ClientInstanceName="cboPart"
                                                        Font-Names="Tahoma" TextFormatString="{0}" Font-Size="8pt" Width="120px">
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

                                    <td style="width:5px;"></td>

                                    <td align="left" valign="middle" height="20px" width="130px">
                                        <dx:ASPxLabel ID="ASPxLabel16" runat="server" Text="PO EMERGENCY"
                                            Font-Names="Tahoma" Font-Size="8pt">
                                        </dx:ASPxLabel>
                                    </td>
                                    <td align="left" valign="middle" >
                                        <table style="width: 49%;">
                                            <tr>
                                                <td class="style1">
                                                    <dx:ASPxRadioButton ID="rdrEAll" ClientInstanceName="rdrEAll" runat="server" 
                                                        Text="ALL" GroupName="EMERGENCY" Font-Names="Tahoma" Font-Size="8pt">
                                                        <ClientSideEvents CheckedChanged="function(s, e) {
                                                            grid.PerformCallback('clear');
                                                            lblInfo.SetText('');
                                                        }" />
                                                    </dx:ASPxRadioButton>
                                                </td>
                                                <td class="style2">
                                                    <dx:ASPxRadioButton ID="rdrEyes" ClientInstanceName="rdrEYes" 
                                                        runat="server" Text="YES" GroupName="EMERGENCY" Font-Names="Tahoma" 
                                                        Font-Size="8pt">
                                                        <ClientSideEvents CheckedChanged="function(s, e) {
                                                            grid.PerformCallback('clear');
                                                            lblInfo.SetText('');
                                                        }" />
                                                    </dx:ASPxRadioButton>
                                                </td>
                                                <td>
                                                    <dx:ASPxRadioButton ID="rdrENo" ClientInstanceName="rdrENo" 
                                                        runat="server" Text="NO" GroupName="EMERGENCY" Font-Names="Tahoma" 
                                                        Font-Size="8pt">
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
                                    <td style="width:5px;"></td>
                                    <td align="left" valign="middle" >
                                        <dx:ASPxLabel ID="ASPxLabel23" runat="server" Text="BOX NO."
                                            Font-Names="Tahoma" Font-Size="8pt">
                                        </dx:ASPxLabel>
                                    </td>
                                    <td align="left" valign="middle" >
                                        <dx:ASPxTextBox ID="txtboxno" runat="server" Width="120px" Font-Names="Tahoma"
                                            Font-Size="8pt" ClientInstanceName="txtboxno">
                                        </dx:ASPxTextBox>
                                    </td>

                                    <td style="width:5px;"></td>

                                    <td align="left" valign="middle" height="20px" width="130px">
                                        <dx:ASPxLabel ID="ASPxLabel21" runat="server" Text="PO NO."
                                            Font-Names="Tahoma" Font-Size="8pt">
                                        </dx:ASPxLabel>
                                    </td>
                                    <td align="left" valign="middle" >
                                        <dx:ASPxTextBox ID="txtpono" runat="server" Width="120px" Font-Names="Tahoma"
                                            Font-Size="8pt" ClientInstanceName="txtpono">
                                        </dx:ASPxTextBox>
                                    </td>
                                </tr>
                                <!-- ROW 6 -->
                                <!-- ROW 7 -->
                                <!-- ROW 8 -->
                                <!-- ROW 9 -->
                                <!-- ROW 10 -->
                                <tr>
                                    <td style="width:5px;"></td>
                                    <td align="left" valign="middle" >
                                        &nbsp;</td>
                                    <td align="left" valign="middle" >
                                        &nbsp;</td>

                                    <td style="width:5px;"></td>

                                    <td align="left" valign="middle" height="20px" width="130px">
                                        
                                    </td>
                                    <td align="right" valign="middle" >
                                        <table>
                                            <tr>
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
            <td align="right">&nbsp</td>
            <td align="right">
            </td>
            <td align="right">
                &nbsp;</td>
        </tr>
        <tr>
            <td colspan="3" align="left" valign="top" height="220">
                <%--Column : Grid--%>
                <dx:ASPxGridView ID="grid" runat="server" Width="100%"
                    Font-Names="Tahoma" KeyFieldName="AffiliateCode;PONo;KanbanNo;PartNo;SupplierCode;SupplierSJNo;PASISJNo"
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
<ClientSideEvents FocusedRowChanged="function(s, e) {
	                    OnGridFocusedRowChanged();
                    }" RowClick="function(s, e) {
	                    lblInfo.SetText(&#39;&#39;);
                    }" BatchEditStartEditing="OnBatchEditStartEditing" 
                        BatchEditEndEditing="OnBatchEditEndEditing" EndCallback="function(s, e) {
                        var pMsg = s.cpMessage;
                        if (pMsg != &#39;&#39;) {
                            if (pMsg.substring(1,5) == &#39;1001&#39; || pMsg.substring(1,5) == &#39;1002&#39; || pMsg.substring(1,5) == &#39;1003&#39; || pMsg.substring(1,5) == &#39;2001&#39;) {
                                lblInfo.GetMainElement().style.color = &#39;Blue&#39;;                                    
                            } else {
                                lblInfo.GetMainElement().style.color = &#39;Red&#39;;
                            }
                            lblInfo.SetText(pMsg);
                        } else {
                            lblInfo.SetText(&#39;&#39;);
                        }
                        delete s.cpMessage;
                        delete s.cpPONo;
                    }" CallbackError="function(s, e) {e.handled = true;}" Init="OnInit"></ClientSideEvents>
                    
                    <Columns>
                                                                  
                        <dx:GridViewDataTextColumn VisibleIndex="0" Caption="NO" FieldName="No" Width="30px"
                            HeaderStyle-HorizontalAlign="Center">
                            <HeaderStyle HorizontalAlign="Center" Wrap="True"></HeaderStyle>
                            <CellStyle Font-Names="Tahoma" Font-Size="8pt" HorizontalAlign="Center">
                            </CellStyle>                            
                        </dx:GridViewDataTextColumn>
                        <dx:GridViewDataTextColumn VisibleIndex="1" Caption="PERIOD" FieldName="Period" Width="70px"
                            HeaderStyle-HorizontalAlign="Center">
                            <HeaderStyle HorizontalAlign="Center" Wrap="True"></HeaderStyle>
                            <CellStyle Font-Names="Tahoma" Font-Size="8pt" HorizontalAlign="Center">
                            </CellStyle>
                        </dx:GridViewDataTextColumn>
                        <dx:GridViewDataTextColumn VisibleIndex="2" Caption="AFFILIATE CODE" 
                            FieldName="AffiliateID" Width="120px"
                            HeaderStyle-HorizontalAlign="Center">
                            <HeaderStyle HorizontalAlign="Center" Wrap="True"></HeaderStyle>
                            <CellStyle Font-Names="Tahoma" Font-Size="8pt">
                            </CellStyle>
                        </dx:GridViewDataTextColumn>
                        <dx:GridViewDataTextColumn VisibleIndex="3" Caption="AFFILIATE NAME" FieldName="AffiliateName" Width="170px"
                            HeaderStyle-HorizontalAlign="Center">
                            <HeaderStyle HorizontalAlign="Center" Wrap="True"></HeaderStyle>
                            <CellStyle Font-Names="Tahoma" Font-Size="8pt">
                            </CellStyle>
                        </dx:GridViewDataTextColumn>
                        <dx:GridViewDataTextColumn VisibleIndex="6" Caption="PO NO." 
                            FieldName="OrderNo" Width="70px"
                            HeaderStyle-HorizontalAlign="Center">
                            <HeaderStyle HorizontalAlign="Center" Wrap="True"></HeaderStyle>
                            <CellStyle Font-Names="Tahoma" Font-Size="8pt">
                            </CellStyle>
                        </dx:GridViewDataTextColumn>
                        <dx:GridViewDataTextColumn VisibleIndex="7" Caption="PO EMERGENCY" 
                            FieldName="EmergencyCls" Width="100px"
                            HeaderStyle-HorizontalAlign="Center">
<HeaderStyle HorizontalAlign="Center"></HeaderStyle>
                        </dx:GridViewDataTextColumn>
                        <dx:GridViewDataTextColumn VisibleIndex="8" Caption="SUPPLIER CODE" 
                            FieldName="SupplierID" Width="100px"
                            HeaderStyle-HorizontalAlign="Center">
                            <HeaderStyle HorizontalAlign="Center" Wrap="True"></HeaderStyle>
                            <CellStyle Font-Names="Tahoma" Font-Size="8pt" HorizontalAlign="Center" >
                            </CellStyle>
                        </dx:GridViewDataTextColumn>
                        <dx:GridViewDataTextColumn VisibleIndex="9" Caption="SUPPLIER NAME" 
                            FieldName="SupplierName" Width="150px"
                            HeaderStyle-HorizontalAlign="Center">
                            <HeaderStyle HorizontalAlign="Center" Wrap="True"></HeaderStyle>
                            <CellStyle Font-Names="Tahoma" Font-Size="8pt">
                            </CellStyle>
                        </dx:GridViewDataTextColumn>
                        <dx:GridViewDataTextColumn VisibleIndex="10" Caption="ETD VENDOR" 
                            FieldName="ETDVendor" Width="90px"
                            HeaderStyle-HorizontalAlign="Center">
                            <HeaderStyle HorizontalAlign="Center" Wrap="True"></HeaderStyle>
                            <CellStyle Font-Names="Tahoma" Font-Size="8pt" HorizontalAlign="Left">
                            </CellStyle>
                        </dx:GridViewDataTextColumn>

                        <dx:GridViewDataTextColumn VisibleIndex="11" Caption="ETD PORT" 
                            FieldName="ETDPort" Width="90px"
                            HeaderStyle-HorizontalAlign="Center">
                            <HeaderStyle HorizontalAlign="Center" Wrap="True"></HeaderStyle>
                            <CellStyle Font-Names="Tahoma" Font-Size="8pt" HorizontalAlign="Center">
                            </CellStyle>
                        </dx:GridViewDataTextColumn>
                        <dx:GridViewDataTextColumn VisibleIndex="12" Caption="ETA PORT" 
                            FieldName="ETAPort" Width="90px"
                            HeaderStyle-HorizontalAlign="Center">
                            <HeaderStyle HorizontalAlign="Center" Wrap="True"></HeaderStyle>
                            <CellStyle Font-Names="Tahoma" Font-Size="8pt">
                            </CellStyle>
                        </dx:GridViewDataTextColumn>
                        <dx:GridViewDataTextColumn VisibleIndex="13" Caption="ETA FACTORY" 
                            FieldName="ETAFactory" Width="90px"
                            HeaderStyle-HorizontalAlign="Center">
                            <HeaderStyle HorizontalAlign="Center" Wrap="True"></HeaderStyle>
                            <CellStyle Font-Names="Tahoma" Font-Size="8pt" HorizontalAlign="Center">
                            </CellStyle>
                        </dx:GridViewDataTextColumn>
                        <dx:GridViewDataTextColumn VisibleIndex="17" Caption="PART NO." 
                            FieldName="PartNo" Width="100px"
                            HeaderStyle-HorizontalAlign="Center">
                            <HeaderStyle HorizontalAlign="Center"></HeaderStyle>
                            <CellStyle Font-Names="Tahoma" Font-Size="8pt">
                            </CellStyle>
                        </dx:GridViewDataTextColumn>
                        <dx:GridViewDataTextColumn VisibleIndex="18" Caption="PART NAME" 
                            FieldName="PartName" Width="150px"
                            HeaderStyle-HorizontalAlign="Center">
                            <HeaderStyle HorizontalAlign="Center" Wrap="True"></HeaderStyle>
                            <CellStyle Font-Names="Tahoma" Font-Size="8pt">
                            </CellStyle>
                        </dx:GridViewDataTextColumn>
                        <dx:GridViewDataTextColumn VisibleIndex="19" Caption="UOM" FieldName="UOM" Width="40px"
                            HeaderStyle-HorizontalAlign="Center">
                            <HeaderStyle HorizontalAlign="Center" Wrap="True"></HeaderStyle>
                            <CellStyle Font-Names="Tahoma" Font-Size="8pt" HorizontalAlign="Center">
                            </CellStyle>
                        </dx:GridViewDataTextColumn>
                        <dx:GridViewDataTextColumn VisibleIndex="22" Caption="ORDER QTY" 
                            FieldName="POQty" Width="80px"
                            HeaderStyle-HorizontalAlign="Center">
                            <HeaderStyle HorizontalAlign="Center" Wrap="True"></HeaderStyle>
                            <CellStyle Font-Names="Tahoma" Font-Size="8pt" HorizontalAlign="Right">
                            </CellStyle>
                            
                        </dx:GridViewDataTextColumn>
                        <dx:GridViewDataTextColumn VisibleIndex="23" Caption="SUPPLIER DELIVERY QTY" 
                            FieldName="DOQty" Width="80px"
                            HeaderStyle-HorizontalAlign="Center">
                            <HeaderStyle HorizontalAlign="Center" Wrap="True"></HeaderStyle>
                            <CellStyle Font-Names="Tahoma" Font-Size="8pt" HorizontalAlign="Right">
                            </CellStyle>                            
                        </dx:GridViewDataTextColumn>
                        <dx:GridViewDataTextColumn VisibleIndex="25" 
                            Caption="GOOD RECEIVING QTY" FieldName="GoodRecQty" Width="80px"
                            HeaderStyle-HorizontalAlign="Center">
                            <HeaderStyle HorizontalAlign="Center" Wrap="True"></HeaderStyle>
                            <CellStyle Font-Names="Tahoma" Font-Size="8pt" HorizontalAlign="Right">
                            </CellStyle>
                            
                        </dx:GridViewDataTextColumn>
                        <dx:GridViewDataTextColumn VisibleIndex="26" Caption="DEFECT RECEIVING QTY" 
                            FieldName="DefectRecQty" Width="90px"
                            HeaderStyle-HorizontalAlign="Center">
                            <HeaderStyle HorizontalAlign="Center" Wrap="True"></HeaderStyle>
                            <CellStyle Font-Names="Tahoma" Font-Size="8pt" HorizontalAlign="Right">
                            </CellStyle>
                            
                        </dx:GridViewDataTextColumn>
                        <dx:GridViewDataTextColumn VisibleIndex="27" Caption="REMAINING RECEIVING QTY" 
                            FieldName="Remaining" Width="90px"
                            HeaderStyle-HorizontalAlign="Center">
                            <HeaderStyle HorizontalAlign="Center"></HeaderStyle>
                            <CellStyle Font-Names="Tahoma" Font-Size="8pt" HorizontalAlign="Right">
                            </CellStyle>
                            
                        </dx:GridViewDataTextColumn>
                        <dx:GridViewDataTextColumn VisibleIndex="20" Caption="QTY/BOX" 
                            FieldName="QtyBox"
                            HeaderStyle-HorizontalAlign="Center" Width="80px">
                            <HeaderStyle HorizontalAlign="Center" Wrap="True"></HeaderStyle>
                            <CellStyle Font-Names="Tahoma" Font-Size="8pt" HorizontalAlign="Center">
                            </CellStyle>
                        </dx:GridViewDataTextColumn>
                        <dx:GridViewDataTextColumn VisibleIndex="21" Caption="LABEL NO" 
                            FieldName="BoxNo"
                            HeaderStyle-HorizontalAlign="Center" Width="150px">
                    <HeaderStyle HorizontalAlign="Center"></HeaderStyle>
                        </dx:GridViewDataTextColumn>
                    </Columns>
                    
                    <SettingsPager PageSize="16" Position="Top">
                        <Summary Text="Page {0} of {1} [{2} record(s)]"
                                  AllPagesText="Page {0} of {1} " />
                        <Summary AllPagesText="Page {0} of {1} " Text="Page {0} of {1} [{2} record(s)]"></Summary>
                    </SettingsPager>

                    <Settings ShowHorizontalScrollBar="True" ShowVerticalScrollBar="True" ShowGroupButtons="False"
                        ShowStatusBar="Hidden" VerticalScrollableHeight="250" />                   

<Settings ShowGroupButtons="False" ShowVerticalScrollBar="True" ShowHorizontalScrollBar="True" 
                        VerticalScrollableHeight="250" ShowStatusBar="Hidden"></Settings>

                    <Styles>
                        <SelectedRow ForeColor="Black" Wrap="False">
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

    <div style="height:1px;"></div>
    
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
            <%--<dx:GridViewDataTextColumn VisibleIndex="37" Caption="DeliveryByPASICls" 
                            FieldName="DeliveryByPASICls" Width="0px"
                            HeaderStyle-HorizontalAlign="Center">
                            <HeaderStyle HorizontalAlign="Center"></HeaderStyle>
                            <CellStyle Font-Names="Tahoma" Font-Size="8pt">
                            </CellStyle>
                        </dx:GridViewDataTextColumn>  --%>
            <td valign="top" align="right" style="width: 50px;">
                
            </td>            
            <td align="right" style="width:80px;">                                   
                <%--<SettingsBehavior AllowSelectByRowClick="True" AllowSort="False"
                        ColumnResizeMode="Control" EnableRowHotTrack="True" />--%>
            </td>
            <td valign="top" align="right" style="width: 50px;">
                <dx:ASPxButton ID="btnExcel" runat="server" Text="EXCEL"
                    Font-Names="Tahoma"
                    Width="90px" AutoPostBack="False" Font-Size="8pt">
                    <ClientSideEvents Click="function(s, e) {
                       lblInfo.SetText('');
                       grid.PerformCallback('excel');
                    }" />
                </dx:ASPxButton>
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
