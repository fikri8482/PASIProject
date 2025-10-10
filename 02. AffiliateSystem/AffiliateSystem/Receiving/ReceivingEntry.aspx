<%@ Page Title="" Language="vb" AutoEventWireup="false" MasterPageFile="~/SiteContent.Master" CodeBehind="ReceivingEntry.aspx.vb" Inherits="AffiliateSystem.ReceivingEntry" %>
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
        height = height - (height * 53 / 100)
        grid.SetHeight(height);
    }

    function OnBatchEditStartEditing(s, e) {
        currentColumnName = e.focusedColumn.fieldName;

        if (currentColumnName == "ColNo" || currentColumnName == "PONo" || currentColumnName == "POKanban" || currentColumnName == "KanbanNo"
            || currentColumnName == "PartNo" || currentColumnName == "PartName" || currentColumnName == "UOM" || currentColumnName == "QtyBox" || currentColumnName == "SupplierDeliveryQty"
            || currentColumnName == "PASIGoodReceivingQty" || currentColumnName == "PASIDefectQty" || currentColumnName == "PASIDeliveryQty" || currentColumnName == "RemainingReceivingQty"
            || currentColumnName == "ReceivingQtyBox") {
            e.cancel = true;
        }

        currentEditableVisibleIndex = e.visibleIndex;
    }

    function OnBatchEditEndEditing(s, e) {
        window.setTimeout(function () {
            //if (currentColumnName == "GoodReceivingQty" || currentColumnName == "DefectReceivingQty") {
            var pDeliveryByPASICls = s.batchEditApi.GetCellValue(e.visibleIndex, "DeliveryByPASICls");
            var pDelivery;
            var pMsg;
            if (pDeliveryByPASICls == "1") {
                pDelivery = s.batchEditApi.GetCellValue(e.visibleIndex, "PASIDeliveryQty");
                pMsg = "[7001] Value can't greater than PASI Delivery Qty !!";
            } else {
                pDelivery = s.batchEditApi.GetCellValue(e.visibleIndex, "SupplierDeliveryQty");
                pMsg = "[7001] Value can't greater than Supplier Delivery Qty !!";
            }

            var pGRQty = s.batchEditApi.GetCellValue(e.visibleIndex, "GoodReceivingQty");
            var pDefQty = s.batchEditApi.GetCellValue(e.visibleIndex, "DefectReceivingQty");
            s.batchEditApi.SetCellValue(e.visibleIndex, "RemainingReceivingQty", pDelivery - (parseInt(pGRQty) + parseInt(pDefQty)));

            var pRemainingRecQty = s.batchEditApi.GetCellValue(e.visibleIndex, "RemainingReceivingQty");
            if (pRemainingRecQty < 0) {
                lblInfo.SetText(pMsg);
                lblInfo.GetMainElement().style.color = 'Red';
            }
            //}
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
        var msg = confirm('Are you sure want to delete this data ?');
        if (msg == false) {
            e.processOnServer = false;
            return;
        }

        grid.PerformCallback('delete');
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
                                        <dx:ASPxLabel ID="ASPxLabel1" runat="server" Text="RECEIVED DATE"
                                            Font-Names="Tahoma" Font-Size="8pt">
                                        </dx:ASPxLabel>
                                    </td>
                                    <td align="left" valign="middle" height="20px" width="400px">
                                        <dx:ASPxTextBox ID="txtRecDate" runat="server" Width="140px" Font-Names="Tahoma"
                                            Font-Size="8pt" BackColor="#CCCCCC" ReadOnly="True" MaxLength="20"
                                            ClientInstanceName="txtRecDate">
                                        </dx:ASPxTextBox>                                   
                                    </td>                                    
                                    <td style="width:5px;"></td>
                                    <td align="left" valign="middle" height="20px" width="130px">
                                        <dx:ASPxLabel ID="ASPxLabel15" runat="server" Text="PERFORMANCE CLS"
                                            Font-Names="Tahoma" Font-Size="8pt">
                                        </dx:ASPxLabel></td>
                                    <td align="left" valign="middle" height="20px" width="150px">
                                        <table>
                                            <tr>
                                                <td>
                                                    <dx:ASPxComboBox ID="cboPerformanceCls" runat="server" ClientInstanceName="cboPerformanceCls"
                                                        Font-Names="Tahoma" TextFormatString="{0}" Font-Size="8pt" Width="80px">
                                                        <ClientSideEvents SelectedIndexChanged="function(s, e) {
	                                                        txtPerformanceDesc.SetText(cboPerformanceCls.GetSelectedItem().GetColumnText(1));                                                            
                                                            }" />
                                                    </dx:ASPxComboBox></td>
                                                <td>
                                                    <dx:ASPxTextBox ID="txtPerformanceDesc" runat="server" Width="150px" Font-Names="Tahoma"
                                                        Font-Size="8pt" BackColor="#CCCCCC" ReadOnly="True" MaxLength="30"
                                                        ClientInstanceName="txtPerformanceDesc">
                                                    </dx:ASPxTextBox></td>
                                            </tr>
                                        </table>                                         
                                    </td>
                                </tr>
                                <tr>
                                    <td style="width:5px;"></td>
                                    <td align="left" valign="middle" >
                                        <dx:ASPxLabel ID="ASPxLabel8" runat="server" Text="INVOICE PASI NO."
                                            Font-Names="Tahoma" Font-Size="8pt">
                                        </dx:ASPxLabel>
                                    </td>
                                    <td align="left" valign="middle" >
                                        <table>
                                            <tr>
                                                <td>
                                                    <dx:ASPxTextBox ID="txtSupplierSJNo" runat="server" Width="100px" Font-Names="Tahoma"
                                                        Font-Size="8pt" BackColor="#CCCCCC" ReadOnly="True" MaxLength="20"
                                                        ClientInstanceName="txtSupplierSJNo">
                                                    </dx:ASPxTextBox>
                                                </td>
                                                <td style="width:170px;">
                                                    <dx:ASPxLabel ID="ASPxLabel4" runat="server" Text="SUPPLIER PLAN DELIVERY DATE"
                                                        Font-Names="Tahoma" Font-Size="8pt">
                                                    </dx:ASPxLabel>
                                                </td>
                                                <td>
                                                    <dx:ASPxTextBox ID="txtSupplierPlanDeliveryDate" runat="server" Width="80px" Font-Names="Tahoma"
                                                        Font-Size="8pt" BackColor="#CCCCCC" ReadOnly="True" MaxLength="20"
                                                        ClientInstanceName="txtSupplierPlanDeliveryDate">
                                                    </dx:ASPxTextBox>
                                                </td>
                                            </tr>
                                        </table> 
                                    </td>                                    
                                    <td style="width:5px;"></td>
                                    <td align="left" valign="middle" height="20px" width="130px">
                                        <dx:ASPxLabel ID="ASPxLabel7" runat="server" Text="DELIVERY LOCATION"
                                            Font-Names="Tahoma" Font-Size="8pt">
                                        </dx:ASPxLabel>
                                    </td>
                                    <td align="left" valign="middle" height="20px" width="150px">
                                         <table>
                                            <tr>
                                                <td>
                                                    <dx:ASPxTextBox ID="txtDeliveryLocationCode" runat="server" Width="80px" Font-Names="Tahoma"
                                                        Font-Size="8pt" BackColor="#CCCCCC" ReadOnly="True" MaxLength="20"
                                                        ClientInstanceName="txtDeliveryLocationCode">
                                                    </dx:ASPxTextBox>                                                     
                                                </td>
                                                <td>
                                                    <dx:ASPxTextBox ID="txtDeliveryLocationName" runat="server" Width="150px" Font-Names="Tahoma"
                                                        Font-Size="8pt" BackColor="#CCCCCC" ReadOnly="True" MaxLength="20"
                                                        ClientInstanceName="txtDeliveryLocationName">
                                                    </dx:ASPxTextBox>
                                                </td>
                                            </tr>
                                        </table> 
                                    </td>
                                </tr>
                                <tr>
                                    <td style="width:5px;"></td>
                                    <td align="left" valign="middle" >
                                        <dx:ASPxLabel ID="ASPxLabel10" runat="server" Text="PASI SURAT JALAN NO."
                                            Font-Names="Tahoma" Font-Size="8pt">
                                        </dx:ASPxLabel>
                                    </td>
                                    <td align="left" valign="middle" >
                                        <table>
                                            <tr>
                                                <td>
                                                    <dx:ASPxTextBox ID="txtPASISJNo" runat="server" Width="100px" Font-Names="Tahoma"
                                                        Font-Size="8pt" BackColor="#CCCCCC" ReadOnly="True" MaxLength="20"
                                                        ClientInstanceName="txtPASISJNo">
                                                    </dx:ASPxTextBox>
                                                </td>
                                                <td style="width:170px;">
                                                    <dx:ASPxLabel ID="ASPxLabel5" runat="server" Text="PASI DELIVERY DATE"
                                                        Font-Names="Tahoma" Font-Size="8pt">
                                                    </dx:ASPxLabel>
                                                </td>
                                                <td>
                                                    <dx:ASPxTextBox ID="txtPASIDeliveryDate" runat="server" Width="80px" Font-Names="Tahoma"
                                            Font-Size="8pt" BackColor="#CCCCCC" ReadOnly="True" MaxLength="20"
                                            ClientInstanceName="txtPASIDeliveryDate">
                                        </dx:ASPxTextBox>   
                                                </td>
                                            </tr>
                                        </table> 
                                    </td>
                                    <td style="width:5px;"></td>
                                    <td align="left" valign="middle" >
                                        <dx:ASPxLabel ID="ASPxLabel9" runat="server" Text="SUPPLIER DELIVERY DATE"
                                            Font-Names="Tahoma" Font-Size="8pt">
                                        </dx:ASPxLabel>
                                    </td>
                                    <td align="left" valign="middle" >
                                        <dx:ASPxTextBox ID="txtSupplierDeliveryDate" runat="server" Width="80px" Font-Names="Tahoma"
                                            Font-Size="8pt" BackColor="#CCCCCC" ReadOnly="True" MaxLength="20"
                                            ClientInstanceName="txtSupplierDeliveryDate">
                                        </dx:ASPxTextBox> 
                                    </td>
                                </tr>
                                <tr>
                                    <td style="width:5px;"></td>
                                    <td align="left" valign="middle" >
                                        <dx:ASPxLabel ID="ASPxLabel12" runat="server" Text="DRIVER NAME"
                                            Font-Names="Tahoma" Font-Size="8pt">
                                        </dx:ASPxLabel>
                                    </td>
                                    <td align="left" valign="middle" colspan="4">
                                        <table>
                                            <tr>
                                                <td>
                                                    <dx:ASPxTextBox ID="txtDriverName" runat="server" Width="100px" Font-Names="Tahoma"
                                                        Font-Size="8pt" BackColor="White" MaxLength="20"
                                                        ClientInstanceName="txtDriverName">
                                                    </dx:ASPxTextBox>
                                                </td>
                                                <td>
                                                    <dx:ASPxLabel ID="ASPxLabel2" runat="server" Text="DRIVER CONTACT"
                                                        Font-Names="Tahoma" Font-Size="8pt">
                                                    </dx:ASPxLabel>
                                                </td>
                                                <td>
                                                    <dx:ASPxTextBox ID="txtDriverContact" runat="server" Width="100px" Font-Names="Tahoma"
                                                        Font-Size="8pt" BackColor="White" MaxLength="20"
                                                        ClientInstanceName="txtDriverContact">
                                                    </dx:ASPxTextBox>
                                                </td>
                                                <td>
                                                    <dx:ASPxLabel ID="ASPxLabel3" runat="server" Text="NO. POL"
                                                        Font-Names="Tahoma" Font-Size="8pt">
                                                    </dx:ASPxLabel>
                                                </td>
                                                <td>
                                                    <dx:ASPxTextBox ID="txtNoPol" runat="server" Width="70px" Font-Names="Tahoma"
                                                        Font-Size="8pt" BackColor="White" MaxLength="20"
                                                        ClientInstanceName="txtNoPol">
                                                    </dx:ASPxTextBox>
                                                </td>
                                                <td>
                                                    <dx:ASPxLabel ID="ASPxLabel13" runat="server" Text="JENIS ARMADA"
                                                        Font-Names="Tahoma" Font-Size="8pt">
                                                    </dx:ASPxLabel>
                                                </td>
                                                <td>
                                                    <dx:ASPxTextBox ID="txtJenisArmada" runat="server" Width="70px" Font-Names="Tahoma"
                                                        Font-Size="8pt" BackColor="White" MaxLength="20"
                                                        ClientInstanceName="txtJenisArmada">
                                                    </dx:ASPxTextBox>
                                                </td>
                                                <td style="width:50;">
                                                </td>
                                                <td>
                                                    <dx:ASPxLabel ID="ASPxLabel14" runat="server" Text="TOTAL BOX"
                                                        Font-Names="Tahoma" Font-Size="8pt">
                                                    </dx:ASPxLabel>
                                                </td>
                                                <td>
                                                    <dx:ASPxTextBox ID="txtTotalBox" runat="server" Width="50px" Font-Names="Tahoma"
                                                        Font-Size="8pt" BackColor="#CCCCCC" ReadOnly="True" MaxLength="20" HorizontalAlign="Right"
                                                        ClientInstanceName="txtTotalBox">
                                                    </dx:ASPxTextBox>
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
            <td align="right" style="width:91%;">
                <dx:ASPxImage ID="imgDifferent" runat="server" ShowLoadingImage="true" 
                    ImageUrl="~/Images/fuchsia.jpg" Height="15px" Width="15px">
                </dx:ASPxImage>
            </td>
            <td align="right">
                <dx:ASPxLabel ID="ASPxLabel11" runat="server" Text=": DIFFERENCE" Font-Names="Tahoma" 
                    ClientInstanceName="ASPxLabel11" Font-Bold="True" Font-Size="8pt">
                </dx:ASPxLabel>
            </td>
        </tr>
        <tr>
            <td colspan="8" align="left" valign="top" height="220">
                <%--Column : Grid--%>
                <dx:ASPxGridView ID="grid" runat="server" Width="100%"
                    Font-Names="Tahoma" KeyFieldName="PONo;KanbanNo;PartNo;sjpasi"
                    AutoGenerateColumns="False"
                    ClientInstanceName="grid" Font-Size="8pt">
                    <ClientSideEvents Init="OnInit" FocusedRowChanged="function(s, e) {
	                    OnGridFocusedRowChanged();
                    }" CallbackError="function(s, e) {e.handled = true;}" BatchEditEndEditing="OnBatchEditEndEditing" EndCallback="function(s, e) {
                        grid.CancelEdit();
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

                        txtTotalBox.SetText(s.cpTotalBox);

                        delete s.cpTotalBox;
                        delete s.cpMessage;
                        delete s.cpPONo;
                    }" RowClick="function(s, e) {
	                    lblInfo.SetText('');
                    }" BatchEditStartEditing="OnBatchEditStartEditing" />                    
                    <Columns>
                        <dx:GridViewDataTextColumn VisibleIndex="0" Caption="NO" FieldName="ColNo" Width="30px"
                            HeaderStyle-HorizontalAlign="Center">
                            <HeaderStyle HorizontalAlign="Center" Wrap="True"></HeaderStyle>
                            <CellStyle Font-Names="Tahoma" Font-Size="8pt" HorizontalAlign="Center">
                            </CellStyle>                            
                        </dx:GridViewDataTextColumn>
                        <dx:GridViewDataTextColumn VisibleIndex="2" Caption="PO NO." FieldName="PONo" Width="100px"
                            HeaderStyle-HorizontalAlign="Center">
                            <HeaderStyle HorizontalAlign="Center" Wrap="True"></HeaderStyle>
                            <CellStyle Font-Names="Tahoma" Font-Size="8pt">
                            </CellStyle>
                        </dx:GridViewDataTextColumn>
                        <dx:GridViewDataTextColumn VisibleIndex="3" Caption="PO KANBAN" 
                            FieldName="POKanban" Width="60px"
                            HeaderStyle-HorizontalAlign="Center">
                            <HeaderStyle HorizontalAlign="Center" Wrap="True"></HeaderStyle>
                            <CellStyle Font-Names="Tahoma" Font-Size="8pt" HorizontalAlign="Center">
                            </CellStyle>
                        </dx:GridViewDataTextColumn>
                        <dx:GridViewDataTextColumn VisibleIndex="4" Caption="KANBAN NO." 
                            FieldName="KanbanNo" Width="100px"
                            HeaderStyle-HorizontalAlign="Center">
                            <HeaderStyle HorizontalAlign="Center" Wrap="True"></HeaderStyle>
                            <CellStyle Font-Names="Tahoma" Font-Size="8pt">
                            </CellStyle>
                        </dx:GridViewDataTextColumn>
                        <dx:GridViewDataTextColumn VisibleIndex="5" Caption="PART NO." 
                            FieldName="PartNo" Width="100px"
                            HeaderStyle-HorizontalAlign="Center">
                            <HeaderStyle HorizontalAlign="Center" Wrap="True"></HeaderStyle>
                            <CellStyle Font-Names="Tahoma" Font-Size="8pt">
                            </CellStyle>
                        </dx:GridViewDataTextColumn>
                        <dx:GridViewDataTextColumn VisibleIndex="6" Caption="PART NAME" 
                            FieldName="PartName" Width="150px"
                            HeaderStyle-HorizontalAlign="Center">
                            <HeaderStyle HorizontalAlign="Center" Wrap="True"></HeaderStyle>
                            <CellStyle Font-Names="Tahoma" Font-Size="8pt">
                            </CellStyle>
                        </dx:GridViewDataTextColumn>
                        <dx:GridViewDataTextColumn VisibleIndex="7" Caption="UOM" FieldName="UOM" Width="40px"
                            HeaderStyle-HorizontalAlign="Center">
                            <HeaderStyle HorizontalAlign="Center" Wrap="True"></HeaderStyle>
                            <CellStyle Font-Names="Tahoma" Font-Size="8pt" HorizontalAlign="Center">
                            </CellStyle>
                        </dx:GridViewDataTextColumn>
                        <dx:GridViewDataTextColumn VisibleIndex="8" Caption="QTY/ BOX" 
                            FieldName="QtyBox" Width="50px"
                            HeaderStyle-HorizontalAlign="Center">
                            <HeaderStyle HorizontalAlign="Center" Wrap="True"></HeaderStyle>
                            <CellStyle Font-Names="Tahoma" Font-Size="8pt" HorizontalAlign="Right">
                            </CellStyle>
                        </dx:GridViewDataTextColumn>
                        <dx:GridViewDataTextColumn VisibleIndex="9" Caption="SUPPLIER DELIVERY QTY" 
                            FieldName="SupplierDeliveryQty" Width="0px"
                            HeaderStyle-HorizontalAlign="Center">
                            <HeaderStyle HorizontalAlign="Center" Wrap="True"></HeaderStyle>
                            <CellStyle Font-Names="Tahoma" Font-Size="8pt" HorizontalAlign="Right">
                            </CellStyle>
                        </dx:GridViewDataTextColumn>
                        <dx:GridViewDataTextColumn VisibleIndex="10" Caption="PASI GOOD RECEIVING QTY" 
                            FieldName="PASIGoodReceivingQty" Width="0px"
                            HeaderStyle-HorizontalAlign="Center">
                            <HeaderStyle HorizontalAlign="Center" Wrap="True"></HeaderStyle>
                            <CellStyle Font-Names="Tahoma" Font-Size="8pt" HorizontalAlign="Right">
                            </CellStyle>
                        </dx:GridViewDataTextColumn>
                        <dx:GridViewDataTextColumn VisibleIndex="11" 
                            Caption="PASI DEFECT RECEIVING QTY" FieldName="PASIDefectQty" Width="0px"
                            HeaderStyle-HorizontalAlign="Center">
                            <HeaderStyle HorizontalAlign="Center" Wrap="True"></HeaderStyle>
                            <CellStyle Font-Names="Tahoma" Font-Size="8pt" HorizontalAlign="Right">
                            </CellStyle>
                        </dx:GridViewDataTextColumn>
                        <dx:GridViewDataTextColumn VisibleIndex="12" Caption="PASI DELIVERY QTY" 
                            FieldName="PASIDeliveryQty" Width="80px"
                            HeaderStyle-HorizontalAlign="Center">
                            <HeaderStyle HorizontalAlign="Center" Wrap="True"></HeaderStyle>
                            <CellStyle Font-Names="Tahoma" Font-Size="8pt" HorizontalAlign="Right">
                            </CellStyle>
                        </dx:GridViewDataTextColumn>
                        <dx:GridViewDataTextColumn VisibleIndex="13" 
                            Caption="AFFILIATE GOOD RECEIVING QTY*" FieldName="GoodReceivingQty" Width="90px"
                            HeaderStyle-HorizontalAlign="Center">
                            <PropertiesTextEdit>
                                <Style HorizontalAlign="Right"></Style>
                            </PropertiesTextEdit>
                            <HeaderStyle HorizontalAlign="Center" Wrap="True"></HeaderStyle>
                            <CellStyle Font-Names="Tahoma" Font-Size="8pt" HorizontalAlign="Right">
                            </CellStyle>
                        </dx:GridViewDataTextColumn>
                        <dx:GridViewDataTextColumn VisibleIndex="14" 
                            Caption="AFFILIATE DEFECT RECEIVING QTY*" FieldName="DefectReceivingQty" Width="90px"
                            HeaderStyle-HorizontalAlign="Center">
                            <PropertiesTextEdit>
                                <Style HorizontalAlign="Right"></Style>
                            </PropertiesTextEdit>
                            <HeaderStyle HorizontalAlign="Center" Wrap="True"></HeaderStyle>
                            <CellStyle Font-Names="Tahoma" Font-Size="8pt" HorizontalAlign="Right">
                            </CellStyle>
                        </dx:GridViewDataTextColumn>
                        <dx:GridViewDataTextColumn VisibleIndex="15" 
                            Caption="AFFILIATE REMAINING RECEIVING QTY" FieldName="RemainingReceivingQty" Width="90px"
                            HeaderStyle-HorizontalAlign="Center">
                            <HeaderStyle HorizontalAlign="Center" Wrap="True"></HeaderStyle>
                            <CellStyle Font-Names="Tahoma" Font-Size="8pt" HorizontalAlign="Right">
                            </CellStyle>
                        </dx:GridViewDataTextColumn>
                        <dx:GridViewDataTextColumn VisibleIndex="16" 
                            Caption="AFFILIATE RECEIVING QTY (BOX)" FieldName="ReceivingQtyBox" Width="80px"
                            HeaderStyle-HorizontalAlign="Center">
                            <HeaderStyle HorizontalAlign="Center" Wrap="True"></HeaderStyle>
                            <CellStyle Font-Names="Tahoma" Font-Size="8pt" HorizontalAlign="Right">
                            </CellStyle>
                        </dx:GridViewDataTextColumn>
                        <dx:GridViewDataTextColumn VisibleIndex="17" Caption="UnitCls" 
                            FieldName="UnitCls" Width="0px"
                            HeaderStyle-HorizontalAlign="Center">
                            <HeaderStyle HorizontalAlign="Center" Wrap="True"></HeaderStyle>
                            <CellStyle Font-Names="Tahoma" Font-Size="8pt" HorizontalAlign="Center">
                            </CellStyle>
                        </dx:GridViewDataTextColumn>
                        <dx:GridViewDataTextColumn VisibleIndex="18" Caption="DeliveryByPASICls" 
                            FieldName="DeliveryByPASICls" Width="0px"
                            HeaderStyle-HorizontalAlign="Center">
                            <HeaderStyle HorizontalAlign="Center" Wrap="True"></HeaderStyle>
                            <CellStyle Font-Names="Tahoma" Font-Size="8pt" HorizontalAlign="Center">
                            </CellStyle>
                        </dx:GridViewDataTextColumn>
                        <dx:GridViewDataTextColumn VisibleIndex="19" Caption="Is Saved" 
                            FieldName="IsSaved" Width="0px"
                            HeaderStyle-HorizontalAlign="Center">
                            <HeaderStyle HorizontalAlign="Center" Wrap="True"></HeaderStyle>
                            <CellStyle Font-Names="Tahoma" Font-Size="8pt" HorizontalAlign="Center">
                            </CellStyle>
                        </dx:GridViewDataTextColumn>
                        <dx:GridViewDataTextColumn Caption="SUPPLIER" FieldName="Supplier" 
                            VisibleIndex="1" Width="100px">
                            <HeaderStyle HorizontalAlign="Center" VerticalAlign="Middle" />
                        </dx:GridViewDataTextColumn>
                    </Columns>
                    <SettingsBehavior AllowSelectByRowClick="True" AllowSort="False"
                        ColumnResizeMode="Control" EnableRowHotTrack="True" />
                    <SettingsPager PageSize="10" Position="Top">
                        <Summary Text="Page {0} of {1} [{2} record(s)]"
                                  AllPagesText="Page {0} of {1} " />
                        <Summary AllPagesText="Page {0} of {1} " Text="Page {0} of {1} [{2} record(s)]"></Summary>
                    </SettingsPager>
                    <SettingsEditing Mode="Batch" NewItemRowPosition="Bottom">
                        <BatchEditSettings ShowConfirmOnLosingChanges="False" />
                    </SettingsEditing>
                    <Settings ShowHorizontalScrollBar="True" ShowVerticalScrollBar="True"
                        ShowGroupButtons="False" ShowStatusBar="Hidden"
                        VerticalScrollableHeight="220" />
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
                <dx:ASPxButton ID="btnSubMenu" runat="server" Text="BACK"
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
            </td>--%>
            <td valign="top" align="right" style="width: 50px;">
                <dx:ASPxButton ID="btnBC40" runat="server" Text="BC 4.0"
                    Font-Names="Tahoma"
                    Width="90px" AutoPostBack="False" Font-Size="8pt" Enabled="true" 
                    Visible="true">
                    <ClientSideEvents Click="function(s, e) {
	                    grid.PerformCallback('bc40')
                    }" />
                </dx:ASPxButton>
            </td>
            <td valign="top" align="right" style="width: 50px;">
                <dx:ASPxButton ID="btnPrintGR" runat="server" Text="PRINT GOOD RECEIVING"
                    Font-Names="Tahoma"
                    Width="90px" AutoPostBack="False" Font-Size="8pt" Enabled="true" 
                    Visible="true">
                </dx:ASPxButton>
            </td>
            <td valign="top" align="right" style="width: 50px;">
                <dx:ASPxButton ID="btnSendGR" runat="server" Text="SEND GOOD RECEIVING TO SUPPLIER"                              
                    Font-Names="Tahoma" Width="80px" AutoPostBack="False" Font-Size="8pt">
                    <ClientSideEvents Click="function(s, e) {
                        grid.PerformCallback('sendtosupplier');
                    }" />
                </dx:ASPxButton>
            </td>            
            <td align="right" style="width:80px;">                                   
                <dx:ASPxButton ID="btnDelete" runat="server" Text="DELETE"                              
                    Font-Names="Tahoma" Width="80px" AutoPostBack="False" Font-Size="8pt">
                    <ClientSideEvents Click="function(s, e) {
                        up_delete();
                    }" />
                </dx:ASPxButton>
            </td>
            <td valign="top" align="right" style="width: 50px;">
                <dx:ASPxButton ID="btnSubmit" runat="server" Text="SAVE"
                    Font-Names="Tahoma"
                    Width="90px" AutoPostBack="False" Font-Size="8pt">
                    <ClientSideEvents Click="function(s, e) {
                        grid.UpdateEdit();
                        grid.PerformCallback('save');
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

    <dx:ASPxCallback ID="ButtonDelete" runat="server" ClientInstanceName = "ButtonDelete">
        <ClientSideEvents EndCallback="function(s, e) {
            var pMsg = s.cpMessage;
            if (pMsg != '') {
                if (pMsg.substring(1,5) == '1003') {
                    lblInfo.GetMainElement().style.color = 'Blue';
                    clear();
                    grid.PerformCallback('load');
                } else {
                    lblInfo.GetMainElement().style.color = 'Red';
                }
                lblInfo.SetText(pMsg);
            } else {
                lblInfo.SetText('');
            }
            delete s.cpMessage;
        }" />
    </dx:ASPxCallback>

    <dx:ASPxCallback ID="ButtonApprove" runat="server" ClientInstanceName="ButtonApprove">
        <ClientSideEvents EndCallback="function(s, e) {
	        

            var pMsg = s.cpMessage;
            if (pMsg != '') {
                if (pMsg.substring(1,5) == '1007') {
                    lblInfo.GetMainElement().style.color = 'Blue';
                } else {
                    lblInfo.GetMainElement().style.color = 'Red';
                }
                lblInfo.SetText(pMsg);
            } else {
                lblInfo.SetText('');
            }
            delete s.cpMessage;
        }" />
    </dx:ASPxCallback>
</asp:Content>

