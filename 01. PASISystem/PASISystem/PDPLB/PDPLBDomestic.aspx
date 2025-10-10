<%@ Page Title="" Language="vb" AutoEventWireup="false" MasterPageFile="~/SiteContent.Master" CodeBehind="PDPLBDomestic.aspx.vb" Inherits="PASISystem.PDPLBDomestic" %>

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
        function clear() {
            var PeriodTo = new Date();

            dtDeliveryDate.SetDate(PeriodTo);
            txtAffiliateName.SetText('');
            txtInvoice.SetText('');

            cboAffiliateCode.SetText('');
            cboSuratJalan.SetText('');
            cboKategory.SetText('');
            cboRemark.SetText('');

            lblInfo.SetText('');
        }

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
            e.cancel = true;
            currentEditableVisibleIndex = e.visibleIndex;
        }

        function validasi_search() {
            //lblInfo.GetMainElement().style.color = 'Red';
            //if (txtPONo.GetText() == "") {
            //    lblInfo.SetText("[6011] Please Input PO No. first!");
            //    txtPONo.Focus();
            //    e.ProcessOnServer = false;
            //    return false;
            //}
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
                                        <dx:ASPxLabel ID="ASPxLabel4" runat="server" Text="CUSTOMER" Font-Names="Tahoma" Font-Size="8pt">
                                        </dx:ASPxLabel>
                                    </td>
                                    <td align="left" valign="middle" height="20px" width="280px">
                                        <table>
                                            <tr>
                                                <td>
                                                    <dx:ASPxComboBox ID="cboAffiliateCode" runat="server" ClientInstanceName="cboAffiliateCode"
                                                        Font-Names="Tahoma" TextFormatString="{0}" Font-Size="8pt" Width="120px" 
                                                        DropDownStyle="DropDownList">
                                                        <ClientSideEvents SelectedIndexChanged="function(s, e) {
	                                                        txtAffiliateName.SetText(cboAffiliateCode.GetSelectedItem().GetColumnText(1));

                                                            var pAffiD = cboAffiliateCode.GetText();
                                                            var pDateFrom = dtDeliveryDate.GetText();

                                                            cboSuratJalan.PerformCallback(pAffiD + '|' + pDateFrom);

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
                                        <dx:ASPxLabel ID="ASPxLabel2" runat="server" Text="DELIVERY DATE" Font-Names="Tahoma"
                                            Font-Size="8pt">
                                        </dx:ASPxLabel>
                                    </td>
                                    <td align="left" valign="middle" height="20px" width="280px">
                                        <table>
                                            <tr>
                                                <td>
                                                    <dx:ASPxDateEdit ID="dtDeliveryDate" runat="server" ClientInstanceName="dtDeliveryDate" Font-Size="8pt"
                                                        EditFormat="Custom" EditFormatString="dd MMM yyyy" Width="120px">
                                                        <ClientSideEvents DateChanged="function(s, e) {
                                                            var pAffiD = cboAffiliateCode.GetText();
                                                            var pDateFrom = dtDeliveryDate.GetText();

                                                            cboSuratJalan.PerformCallback(pAffiD + '|' + pDateFrom);
                                                            }" />
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
                                    <td align="right" valign="middle" height="20px" width="150px">
                                    </td>
                                </tr>
                                <!-- ROW 3 -->
                                <tr>
                                    <td style="width: 5px;">
                                    </td>
                                    <td align="left" valign="middle">
                                        <dx:ASPxLabel ID="ASPxLabel1" runat="server" Text="SURAT JALAN" Font-Names="Tahoma"
                                            Font-Size="8pt">
                                        </dx:ASPxLabel>
                                    </td>
                                    <td align="left" valign="middle">
                                        <table>
                                            <tr>
                                                <td>
                                                    <dx:ASPxComboBox ID="cboSuratJalan" runat="server" ClientInstanceName="cboSuratJalan"
                                                        Font-Names="Tahoma" TextFormatString="{0}" Font-Size="8pt" Width="200px" 
                                                        DropDownStyle="DropDownList">
                                                        <ClientSideEvents SelectedIndexChanged="function(s, e) {
                                                            grid.PerformCallback('clear');
                                                            lblInfo.SetText('');
                                                            }" />
                                                    </dx:ASPxComboBox>
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
                                        <dx:ASPxLabel ID="ASPxLabel8" runat="server" Text="INVOICE NO" Font-Names="Tahoma"
                                            Font-Size="8pt">
                                        </dx:ASPxLabel>
                                    </td>
                                    <td align="left" valign="middle">
                                        <table>
                                            <tr>
                                                <td>
                                                    <dx:ASPxTextBox ID="txtInvoice" runat="server" Width="200px" Font-Names="Tahoma"
                                                        Font-Size="8pt" MaxLength="50" ClientInstanceName="txtInvoice">
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
                                    <td align="left" valign="middle">
                                        
                                    </td>
                                </tr>                                
                                <!-- ROW 6 -->
                                <tr>
                                    <td style="width: 5px;">
                                    </td>
                                    <td align="left" valign="middle">
                                        <dx:ASPxLabel ID="ASPxLabel14" runat="server" Text="CATEGORY" Font-Names="Tahoma"
                                            Font-Size="8pt">
                                        </dx:ASPxLabel>
                                    </td>
                                    <td align="left" valign="middle">
                                        <table>
                                            <tr>
                                                <td>
                                                    <dx:ASPxComboBox ID="cboKategory" runat="server" ClientInstanceName="cboKategory"
                                                        Font-Names="Tahoma" TextFormatString="{0}" Font-Size="8pt" Width="120px" 
                                                        DropDownStyle="DropDownList">
                                                        <ClientSideEvents SelectedIndexChanged="function(s, e) {
                                                            grid.PerformCallback('clear');
                                                            lblInfo.SetText('');
                                                            }" />
                                                    </dx:ASPxComboBox>
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
                                </tr>

                                <tr>
                                    <td style="width: 5px;">
                                    </td>
                                    <td align="left" valign="middle">
                                        <dx:ASPxLabel ID="ASPxLabel3" runat="server" Text="REMARK ORDER" Font-Names="Tahoma"
                                            Font-Size="8pt">
                                        </dx:ASPxLabel>
                                    </td>
                                    <td align="left" valign="middle">
                                        <table>
                                            <tr>
                                                <td>
                                                    <dx:ASPxComboBox ID="cboRemark" runat="server" ClientInstanceName="cboRemark"
                                                        Font-Names="Tahoma" TextFormatString="{0}" Font-Size="8pt" Width="120px" 
                                                        DropDownStyle="DropDownList">
                                                        <ClientSideEvents SelectedIndexChanged="function(s, e) {
                                                            grid.PerformCallback('clear');
                                                            lblInfo.SetText('');
                                                            }" />
                                                    </dx:ASPxComboBox>
                                                </td>
                                            </tr>
                                        </table>
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
                                                            clear();

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
            <td colspan="3" align="left" valign="top" height="220">
                <%--Column : Grid--%>
                <dx:ASPxGridView ID="grid" runat="server" Width="100%" Font-Names="Tahoma" KeyFieldName="colNo"
                    AutoGenerateColumns="False" ClientInstanceName="grid" Font-Size="8pt">
                    <ClientSideEvents Init="OnInit" CallbackError="function(s, e) {alert('errror');}"
                        EndCallback="function(s, e) {
                        debugger
                        var pMsg = s.cpMessage;
                        if (pMsg != '' && pMsg != undefined) {
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
                        <dx:GridViewDataTextColumn VisibleIndex="1" Caption="Order No" FieldName="PoNo" Width="150px"
                            HeaderStyle-HorizontalAlign="Center">
                            <HeaderStyle HorizontalAlign="Center" Wrap="True"></HeaderStyle>
                            <CellStyle Font-Names="Tahoma" Font-Size="8pt" HorizontalAlign="Left">
                            </CellStyle>
                        </dx:GridViewDataTextColumn>
                        <dx:GridViewDataTextColumn VisibleIndex="2" Caption="Date" FieldName="Datex" Width="100px"
                            HeaderStyle-HorizontalAlign="Center">
                            <HeaderStyle HorizontalAlign="Center" Wrap="True"></HeaderStyle>
                            <CellStyle Font-Names="Tahoma" Font-Size="8pt" HorizontalAlign="Center">
                            </CellStyle>
                        </dx:GridViewDataTextColumn>
                        <dx:GridViewDataTextColumn VisibleIndex="3" Caption="DN No" FieldName="DNNo" Width="120px" 
                            HeaderStyle-HorizontalAlign="Center">
                            <HeaderStyle HorizontalAlign="Center" Wrap="True"></HeaderStyle>
                            <CellStyle Font-Names="Tahoma" Font-Size="8pt" HorizontalAlign="Left">
                            </CellStyle>
                        </dx:GridViewDataTextColumn>
                        <dx:GridViewDataTextColumn VisibleIndex="4" Caption="Part No" FieldName="PartNo" Width="120px" 
                            HeaderStyle-HorizontalAlign="Center">
                            <HeaderStyle HorizontalAlign="Center" Wrap="True"></HeaderStyle>
                            <CellStyle Font-Names="Tahoma" Font-Size="8pt" HorizontalAlign="Left">
                            </CellStyle>
                        </dx:GridViewDataTextColumn>
                        <dx:GridViewDataTextColumn VisibleIndex="5" Caption="Invoice No" FieldName="InvoiceNo" Width="120px" 
                            HeaderStyle-HorizontalAlign="Center">
                            <HeaderStyle HorizontalAlign="Center" Wrap="True"></HeaderStyle>
                            <CellStyle Font-Names="Tahoma" Font-Size="8pt" HorizontalAlign="Left">
                            </CellStyle>
                        </dx:GridViewDataTextColumn>
                        <dx:GridViewDataTextColumn VisibleIndex="6" Caption="Qty" FieldName="Qty" Width="70px" 
                            HeaderStyle-HorizontalAlign="Center">
                            <HeaderStyle HorizontalAlign="Center" Wrap="True"></HeaderStyle>
                            <CellStyle Font-Names="Tahoma" Font-Size="8pt" HorizontalAlign="Center">
                            </CellStyle>
                        </dx:GridViewDataTextColumn>
                        <dx:GridViewDataTextColumn VisibleIndex="7" Caption="UOM" FieldName="UOM" Width="50px" 
                            HeaderStyle-HorizontalAlign="Center">
                            <HeaderStyle HorizontalAlign="Center" Wrap="True"></HeaderStyle>
                            <CellStyle Font-Names="Tahoma" Font-Size="8pt" HorizontalAlign="Left">
                            </CellStyle>
                        </dx:GridViewDataTextColumn>
                        <dx:GridViewDataTextColumn VisibleIndex="8" Caption="Qty/Box" FieldName="QtyBox" Width="70px" 
                            HeaderStyle-HorizontalAlign="Center">
                            <HeaderStyle HorizontalAlign="Center" Wrap="True"></HeaderStyle>
                            <CellStyle Font-Names="Tahoma" Font-Size="8pt" HorizontalAlign="Center">
                            </CellStyle>
                        </dx:GridViewDataTextColumn>
                        <dx:GridViewDataTextColumn VisibleIndex="9" Caption="Box/Pallet" FieldName="BoxPerPallet" Width="70px" 
                            HeaderStyle-HorizontalAlign="Center">
                            <HeaderStyle HorizontalAlign="Center" Wrap="True"></HeaderStyle>
                            <CellStyle Font-Names="Tahoma" Font-Size="8pt" HorizontalAlign="Center">
                            </CellStyle>
                        </dx:GridViewDataTextColumn>
                        <dx:GridViewDataTextColumn VisibleIndex="10" Caption="Kode Factory Order" FieldName="KodeFactoryOrder" Width="120px" 
                            HeaderStyle-HorizontalAlign="Center">
                            <HeaderStyle HorizontalAlign="Center" Wrap="True"></HeaderStyle>
                            <CellStyle Font-Names="Tahoma" Font-Size="8pt" HorizontalAlign="Left">
                            </CellStyle>
                        </dx:GridViewDataTextColumn>
                        <dx:GridViewDataTextColumn VisibleIndex="11" Caption="Jenis Packing" FieldName="Jenispacking" Width="120px" 
                            HeaderStyle-HorizontalAlign="Center">
                            <HeaderStyle HorizontalAlign="Center" Wrap="True"></HeaderStyle>
                            <CellStyle Font-Names="Tahoma" Font-Size="8pt" HorizontalAlign="Left">
                            </CellStyle>
                        </dx:GridViewDataTextColumn>
                        <dx:GridViewDataTextColumn VisibleIndex="12" Caption="Jenis Order" FieldName="JenisOrder" Width="120px" 
                            HeaderStyle-HorizontalAlign="Center">
                            <HeaderStyle HorizontalAlign="Center" Wrap="True"></HeaderStyle>
                            <CellStyle Font-Names="Tahoma" Font-Size="8pt" HorizontalAlign="Left">
                            </CellStyle>
                        </dx:GridViewDataTextColumn>
                        <dx:GridViewDataTextColumn VisibleIndex="13" Caption="Remark Order" FieldName="RemarkOrder" Width="120px" 
                            HeaderStyle-HorizontalAlign="Center">
                            <HeaderStyle HorizontalAlign="Center" Wrap="True"></HeaderStyle>
                            <CellStyle Font-Names="Tahoma" Font-Size="8pt" HorizontalAlign="Left">
                            </CellStyle>
                        </dx:GridViewDataTextColumn>
                        <dx:GridViewDataTextColumn VisibleIndex="14" Caption="Kode Tujuan Factory Order" FieldName="KodeTujuanFactoryOrder" Width="120px" 
                            HeaderStyle-HorizontalAlign="Center">
                            <HeaderStyle HorizontalAlign="Center" Wrap="True"></HeaderStyle>
                            <CellStyle Font-Names="Tahoma" Font-Size="8pt" HorizontalAlign="Left">
                            </CellStyle>
                        </dx:GridViewDataTextColumn>
                        <dx:GridViewDataTextColumn VisibleIndex="15" Caption="Kategori Barang" FieldName="kategoriBarang" Width="120px" 
                            HeaderStyle-HorizontalAlign="Center">
                            <HeaderStyle HorizontalAlign="Center" Wrap="True"></HeaderStyle>
                            <CellStyle Font-Names="Tahoma" Font-Size="8pt" HorizontalAlign="Left">
                            </CellStyle>
                        </dx:GridViewDataTextColumn>
                        <dx:GridViewDataTextColumn VisibleIndex="16" Caption="Keterangan" FieldName="Keterangan" Width="120px" 
                            HeaderStyle-HorizontalAlign="Center">
                            <HeaderStyle HorizontalAlign="Center" Wrap="True"></HeaderStyle>
                            <CellStyle Font-Names="Tahoma" Font-Size="8pt" HorizontalAlign="Left">
                            </CellStyle>
                        </dx:GridViewDataTextColumn>


                        <%--<dx:GridViewDataTextColumn VisibleIndex="8" Caption="PO SEND TO SUPPLIER DATE"
                            FieldName="PASISendAffiliateDate" Width="110px" 
                            HeaderStyle-HorizontalAlign="Center">
                            <HeaderStyle HorizontalAlign="Center" Wrap="True"></HeaderStyle>
                            <PropertiesTextEdit DisplayFormatString="dd MMM yyyy">
                            </PropertiesTextEdit>
                            <CellStyle Font-Names="Tahoma" Font-Size="8pt" HorizontalAlign="Center">
                            </CellStyle>
                        </dx:GridViewDataTextColumn>
                        <dx:GridViewDataTextColumn Caption="PO QTY" FieldName="QtyPO" VisibleIndex="11" Width="90px"
                            HeaderStyle-HorizontalAlign="Center">
                            <PropertiesTextEdit DisplayFormatString="{0:n0}">
                            </PropertiesTextEdit>
                            <HeaderStyle HorizontalAlign="Center"></HeaderStyle>
                            <CellStyle Font-Names="Tahoma" Font-Size="8pt" HorizontalAlign="Right">
                            </CellStyle>
                        </dx:GridViewDataTextColumn>--%>

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
