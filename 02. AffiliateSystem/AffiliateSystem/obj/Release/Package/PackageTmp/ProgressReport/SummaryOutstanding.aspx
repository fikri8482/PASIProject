<%@ Page Title="" Language="vb" AutoEventWireup="false" MasterPageFile="~/SiteContent.Master" CodeBehind="SummaryOutstanding.aspx.vb" Inherits="AffiliateSystem.SummaryOutstanding" %>
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
                                <tr>
                                    <td style="width:5px;"></td>
                                    <td align="left" valign="middle" height="20px" width="200px">
                                        <dx:ASPxLabel ID="ASPxLabel1" runat="server" Text="PO PERIOD"
                                            Font-Names="Tahoma" Font-Size="8pt">
                                        </dx:ASPxLabel>
                                    </td>
                                    <td align="left" valign="middle" height="20px" width="280px">
                                        <table>
                                            <tr>
                                                <td>
                                                    <dx:ASPxCheckBox ID="chkPOPeriod" runat="server" ClientInstanceName="chkPOPeriod" Text=" " Checked="true">
                                                        <ClientSideEvents
                                                            CheckedChanged="function (s, e) {
                                                                if (s.GetChecked()==true) {
                                                                    dtPOPeriodFrom.SetEnabled(true);
                                                                    dtPOPeriodTo.SetEnabled(true);
                                                                } else {
                                                                    dtPOPeriodFrom.SetEnabled(false);
                                                                    dtPOPeriodTo.SetEnabled(false);
                                                                }
                                                                grid.PerformCallback('clear');
                                                          }"
                                                        />
                                                    </dx:ASPxCheckBox>
                                                </td>
                                                <td>                                                    
                                                    <dx:ASPxTimeEdit ID="dtPOPeriodFrom" runat="server" ClientInstanceName="dtPOPeriodFrom"
                                                        DisplayFormatString="MMM yyyy" EditFormat="Custom" EditFormatString="MMM yyyy"
                                                        Width="100px" HorizontalAlign="Center">
                                                    </dx:ASPxTimeEdit>
                                                </td>
                                                <td>~</td>
                                                <td>                                                    
                                                    <dx:ASPxTimeEdit ID="dtPOPeriodTo" runat="server" ClientInstanceName="dtPOPeriodTo"
                                                        DisplayFormatString="MMM yyyy" EditFormat="Custom" EditFormatString="MMM yyyy"
                                                        Width="100px" HorizontalAlign="Center">
                                                    </dx:ASPxTimeEdit>
                                                </td>
                                            </tr>
                                        </table>                                         
                                    </td>

                                    <td style="width:5px;"></td>

                                    <td>
                                        <dx:ASPxLabel ID="ASPxLabel2" runat="server" Text="PO NO."
                                            Font-Names="Tahoma" Font-Size="8pt">
                                        </dx:ASPxLabel>
                                    </td>
                                    <td>
                                        <dx:ASPxTextBox ID="txtPONo" runat="server" Width="120px" Font-Names="Tahoma"
                                            Font-Size="8pt" ClientInstanceName="txtPONo">
                                        </dx:ASPxTextBox>
                                    </td>

                                    <td style="width:5px;"></td>

                                    <td align="left" valign="middle" height="20px" width="130px">                                       
                                    </td>
                                    <td align="left" valign="middle" height="20px" width="150px">
                                    </td>
                                </tr>

                                <!-- ROW 3 -->
                                <tr>
                                    <td style="width:5px;"></td>
                                    <td align="left" valign="middle" >
                                        <dx:ASPxLabel ID="ASPxLabel14" runat="server" Text="PASI DELIVERY DATE"
                                            Font-Names="Tahoma" Font-Size="8pt">
                                        </dx:ASPxLabel>
                                    </td>
                                    <td align="left" valign="middle" >
                                        <table>
                                            <tr>
                                                <td>
                                                    <dx:ASPxCheckBox ID="chkPASIDelDate" runat="server" ClientInstanceName="chkPASIDelDate" Text=" " Checked="true">
                                                        <ClientSideEvents
                                                            CheckedChanged="function (s, e) {
                                                                if (s.GetChecked()==true) {
                                                                    dtPASIDelDateFrom.SetEnabled(true);
                                                                    dtPASIDelDateTo.SetEnabled(true);
                                                                } else {
                                                                    dtPASIDelDateFrom.SetEnabled(false);
                                                                    dtPASIDelDateTo.SetEnabled(false);
                                                                }
                                                                grid.PerformCallback('clear');
                                                          }"
                                                        />
                                                    </dx:ASPxCheckBox>
                                                </td>
                                                <td>
                                                    <dx:ASPxDateEdit ID="dtPASIDelDateFrom" runat="server" Font-Names="Tahoma" Font-Size="8pt" Width="100px"
                                                        EditFormat="Custom" EditFormatString="dd MMM yyyy" ClientInstanceName="dtPASIDelDateFrom">
                                                    </dx:ASPxDateEdit>
                                                </td>
                                                <td>~</td>
                                                <td>
                                                    <dx:ASPxDateEdit ID="dtPASIDelDateTo" runat="server" Font-Names="Tahoma" Font-Size="8pt" Width="100px"
                                                        EditFormat="Custom" EditFormatString="dd MMM yyyy" ClientInstanceName="dtPASIDelDateTo">
                                                    </dx:ASPxDateEdit>
                                                </td>
                                            </tr>
                                        </table>                                         
                                    </td>

                                    <td style="width:5px;"></td>

                                    <td>
                                        <dx:ASPxLabel ID="ASPxLabel9" runat="server" Text="PASI SJ. NO."
                                            Font-Names="Tahoma" Font-Size="8pt">
                                        </dx:ASPxLabel>
                                    </td>
                                    <td>
                                        <dx:ASPxTextBox ID="txtPASISJNo" runat="server" Width="120px" Font-Names="Tahoma"
                                            Font-Size="8pt" ClientInstanceName="txtPASISJNo">
                                        </dx:ASPxTextBox>
                                    </td>

                                    <td style="width:5px;"></td>

                                    <td align="left" valign="middle" height="20px" width="130px">                                        
                                    </td>
                                    <td align="left" valign="middle" >                                         
                                    </td>
                                </tr>
                                <!-- ROW 4 -->
                                <tr>
                                    <td style="width:5px;"></td>
                                    <td align="left" valign="middle" >
                                        <dx:ASPxLabel ID="ASPxLabel15" runat="server" Text="AFFILIATE RECEIVE DATE"
                                            Font-Names="Tahoma" Font-Size="8pt">
                                        </dx:ASPxLabel>
                                    </td>
                                    <td align="left" valign="middle" >
                                        <table>
                                            <tr>
                                                <td>
                                                    <dx:ASPxCheckBox ID="chkAffiliateRecDate" runat="server" ClientInstanceName="chkAffiliateRecDate" Text=" " Checked="true">
                                                        <ClientSideEvents
                                                            CheckedChanged="function (s, e) {
                                                                if (s.GetChecked()==true) {
                                                                    dtAffiliateRecDateFrom.SetEnabled(true);
                                                                    dtAffiliateRecDateTo.SetEnabled(true);
                                                                } else {
                                                                    dtAffiliateRecDateFrom.SetEnabled(false);
                                                                    dtAffiliateRecDateTo.SetEnabled(false);
                                                                }
                                                                grid.PerformCallback('clear');
                                                          }"
                                                        />
                                                    </dx:ASPxCheckBox>
                                                </td>
                                                <td>
                                                    <dx:ASPxDateEdit ID="dtAffiliateRecDateFrom" runat="server" Font-Names="Tahoma" Font-Size="8pt" Width="100px"
                                                        EditFormat="Custom" EditFormatString="dd MMM yyyy" ClientInstanceName="dtAffiliateRecDateFrom">
                                                    </dx:ASPxDateEdit>
                                                </td>
                                                <td>~</td>
                                                <td>
                                                    <dx:ASPxDateEdit ID="dtAffiliateRecDateTo" runat="server" Font-Names="Tahoma" Font-Size="8pt" Width="100px"
                                                        EditFormat="Custom" EditFormatString="dd MMM yyyy" ClientInstanceName="dtAffiliateRecDateTo">
                                                    </dx:ASPxDateEdit>
                                                </td>
                                            </tr>
                                        </table>                                         
                                    </td>

                                    <td style="width:5px;"></td>

                                    <td>
                                        <dx:ASPxLabel ID="ASPxLabel12" runat="server" Text="PASI INV. NO."
                                            Font-Names="Tahoma" Font-Size="8pt">
                                        </dx:ASPxLabel>
                                    </td>
                                    <td>
                                        <dx:ASPxTextBox ID="txtPASIInvNo" runat="server" Width="120px" Font-Names="Tahoma"
                                            Font-Size="8pt" ClientInstanceName="txtPASIInvNo">
                                        </dx:ASPxTextBox>
                                    </td>

                                    <td style="width:5px;"></td>

                                    <td align="left" valign="middle" height="20px" width="130px">
                                    </td>
                                    <td align="left" valign="middle" >
                                    </td>
                                </tr>
                                <!-- ROW 5 -->
                                <tr>
                                    <td style="width:5px;"></td>
                                    <td align="left" valign="middle" >
                                        <dx:ASPxLabel ID="ASPxLabel17" runat="server" Text="PASI INVOICE DATE"
                                            Font-Names="Tahoma" Font-Size="8pt">
                                        </dx:ASPxLabel>
                                    </td>
                                    <td align="left" valign="middle" >
                                        <table>
                                            <tr>
                                                <td>
                                                    <dx:ASPxCheckBox ID="chkPASIInvDate" runat="server" ClientInstanceName="chkPASIInvDate" Text=" " Checked="true">
                                                        <ClientSideEvents
                                                            CheckedChanged="function (s, e) {
                                                                if (s.GetChecked()==true) {
                                                                    dtPASIInvDateFrom.SetEnabled(true);
                                                                    dtPASIInvDateTo.SetEnabled(true);
                                                                } else {
                                                                    dtPASIInvDateFrom.SetEnabled(false);
                                                                    dtPASIInvDateTo.SetEnabled(false);
                                                                }
                                                                grid.PerformCallback('clear');
                                                          }"
                                                        />
                                                    </dx:ASPxCheckBox>
                                                </td>
                                                <td>
                                                    <dx:ASPxDateEdit ID="dtPASIInvDateFrom" runat="server" Font-Names="Tahoma" Font-Size="8pt" Width="100px"
                                                        EditFormat="Custom" EditFormatString="dd MMM yyyy" ClientInstanceName="dtPASIInvDateFrom">
                                                    </dx:ASPxDateEdit>
                                                </td>
                                                <td>~</td>
                                                <td>
                                                    <dx:ASPxDateEdit ID="dtPASIInvDateTo" runat="server" Font-Names="Tahoma" Font-Size="8pt" Width="100px"
                                                        EditFormat="Custom" EditFormatString="dd MMM yyyy" ClientInstanceName="dtPASIInvDateTo">
                                                    </dx:ASPxDateEdit>
                                                </td>
                                            </tr>
                                        </table>                                         
                                    </td>

                                    <td style="width:5px;"></td>

                                    <td>
                                        
                                    </td>
                                    <td>
                                        
                                    </td>

                                    <td style="width:5px;"></td>

                                    <td align="left" valign="middle" height="20px" width="130px">
                                        
                                    </td>
                                    <td align="left" valign="middle" >
                                         
                                    </td>
                                </tr>

                                <!-- ROW 8 -->
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

                                    <td>
                                        
                                    </td>
                                    <td>
                                        
                                    </td>

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
            <td align="right">
                <table style="width:100%;">
                    <tr>
                        <td align="right" width="90%">
                        </td>
                        <td align="right" width="8%">
                <dx:ASPxImage ID="ASPxImage1" runat="server" ShowLoadingImage="true" 
                    ImageUrl="~/Images/fuchsia.jpg" Height="15px" Width="15px">
                </dx:ASPxImage>
                <dx:ASPxLabel ID="ASPxLabel5" runat="server" Text=" : DIFFERENCE" Font-Names="Tahoma" 
                    ClientInstanceName="difference" Font-Bold="True" Font-Size="8pt">
                </dx:ASPxLabel>
                        </td>
                    </tr>
                </table>
            </td>
        </tr>
        <tr>
            <td align="left" valign="top" height="220">
                <%--Column : Grid--%>
                <dx:ASPxGridView ID="grid" runat="server" Width="100%"
                    Font-Names="Tahoma" KeyFieldName="PONo;KanbanNo;PartNo;SupplierCode"
                    AutoGenerateColumns="False"
                    ClientInstanceName="grid" Font-Size="8pt">
                    <ClientSideEvents Init="OnInit" FocusedRowChanged="function(s, e) {
	                    OnGridFocusedRowChanged();
                    }" CallbackError="function(s, e) {e.handled = true;}" EndCallback="function(s, e) {
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
                    }" />                    
                    <Columns>
                        <dx:GridViewDataTextColumn VisibleIndex="0" Caption="NO" FieldName="ColNo" Width="30px"
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
                            <CellStyle Font-Names="Tahoma" Font-Size="8pt" HorizontalAlign="Center" >
                            </CellStyle>
                        </dx:GridViewDataTextColumn>
                        <dx:GridViewDataTextColumn VisibleIndex="4" Caption="KANBAN NO." 
                            FieldName="KanbanNo" Width="80px"
                            HeaderStyle-HorizontalAlign="Center">
                            <HeaderStyle HorizontalAlign="Center" Wrap="True"></HeaderStyle>
                            <CellStyle Font-Names="Tahoma" Font-Size="8pt">
                            </CellStyle>
                        </dx:GridViewDataTextColumn>
                        <dx:GridViewDataTextColumn VisibleIndex="9" Caption="PASI DELIVERY DATE" 
                            FieldName="PASIDeliveryDate" Width="100px"
                            HeaderStyle-HorizontalAlign="Center">
                            <HeaderStyle HorizontalAlign="Center" Wrap="True"></HeaderStyle>
                            <CellStyle Font-Names="Tahoma" Font-Size="8pt" HorizontalAlign="Center">
                            </CellStyle>
                        </dx:GridViewDataTextColumn>                        
                        <dx:GridViewDataTextColumn VisibleIndex="10" Caption="PASI SURAT JALAN NO." 
                            FieldName="PASISJNo" Width="100px"
                            HeaderStyle-HorizontalAlign="Center">
                            <HeaderStyle HorizontalAlign="Center" Wrap="True"></HeaderStyle>
                            <CellStyle Font-Names="Tahoma" Font-Size="8pt">
                            </CellStyle>
                        </dx:GridViewDataTextColumn>
                        <dx:GridViewDataTextColumn VisibleIndex="12" Caption="AFFILIATE RECEIVE DATE" 
                            FieldName="AffiliateReceiveDate"
                            HeaderStyle-HorizontalAlign="Center">
                            <HeaderStyle HorizontalAlign="Center" Wrap="True"></HeaderStyle>
                            <CellStyle Font-Names="Tahoma" Font-Size="8pt" HorizontalAlign="Center">
                            </CellStyle>
                        </dx:GridViewDataTextColumn>
                        <dx:GridViewDataTextColumn VisibleIndex="5" Caption="PART NO." 
                            FieldName="PartNo" Width="100px"
                            HeaderStyle-HorizontalAlign="Center">
                            <HeaderStyle HorizontalAlign="Center"></HeaderStyle>
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
                        <dx:GridViewDataTextColumn VisibleIndex="11" 
                            Caption="PASI DELIVERY QTY" FieldName="PASIDeliveryQty" Width="80px"
                            HeaderStyle-HorizontalAlign="Center">
                            <HeaderStyle HorizontalAlign="Center" Wrap="True"></HeaderStyle>
                            <PropertiesTextEdit DisplayFormatString="{0:n0}">
                            </PropertiesTextEdit>
                            <CellStyle Font-Names="Tahoma" Font-Size="8pt" HorizontalAlign="Right">
                            </CellStyle>                            
                        </dx:GridViewDataTextColumn>
                        <dx:GridViewDataTextColumn VisibleIndex="13" Caption="AFFILIATE RECEIVING QTY" 
                            FieldName="AffiliateReceivingQty" Width="80px"
                            HeaderStyle-HorizontalAlign="Center">
                            <HeaderStyle HorizontalAlign="Center" Wrap="True"></HeaderStyle>
                            <PropertiesTextEdit DisplayFormatString="{0:n0}">
                            </PropertiesTextEdit>
                            <CellStyle Font-Names="Tahoma" Font-Size="8pt" HorizontalAlign="Right">
                            </CellStyle>                            
                        </dx:GridViewDataTextColumn>                        
                        <dx:GridViewDataTextColumn VisibleIndex="14" Caption="INVOICE TO AFFILIATE" 
                            FieldName="InvoiceNoToAffiliate" Width="140px"
                            HeaderStyle-HorizontalAlign="Center">
                            <HeaderStyle HorizontalAlign="Center"></HeaderStyle>
                            <CellStyle Font-Names="Tahoma" Font-Size="8pt">
                            </CellStyle>
                        </dx:GridViewDataTextColumn>
                        <dx:GridViewDataTextColumn VisibleIndex="15" Caption="INVOICE DATE TO AFFILIATE" 
                            FieldName="InvoiceDateToAffiliate" Width="100px"
                            HeaderStyle-HorizontalAlign="Center">
                            <HeaderStyle HorizontalAlign="Center"></HeaderStyle>
                            <CellStyle Font-Names="Tahoma" Font-Size="8pt" HorizontalAlign="Center">
                            </CellStyle>
                        </dx:GridViewDataTextColumn>
                        <dx:GridViewBandColumn Caption="PASI INVOICE" VisibleIndex="16" 
                            HeaderStyle-HorizontalAlign="Center">
                            <Columns>
                                <dx:GridViewDataTextColumn Caption="CURR" FieldName="InvoiceToAffiliateCurr" VisibleIndex="0">
                                    <HeaderStyle HorizontalAlign="Center" Wrap="True"></HeaderStyle>
                                    <CellStyle Font-Names="Tahoma" Font-Size="8pt" HorizontalAlign="Center">
                                    </CellStyle>
                                </dx:GridViewDataTextColumn>
                                <dx:GridViewDataTextColumn Caption="AMOUNT" FieldName="InvoiceToAffiliateAmount" VisibleIndex="1">
                                    <HeaderStyle HorizontalAlign="Center" Wrap="True"></HeaderStyle>
                                    <PropertiesTextEdit DisplayFormatString="{0:n2}">
                                    </PropertiesTextEdit>
                                    <CellStyle Font-Names="Tahoma" Font-Size="8pt" HorizontalAlign="Right">
                                    </CellStyle>
                                </dx:GridViewDataTextColumn>
                            </Columns>
                        </dx:GridViewBandColumn>                    
                        <dx:GridViewDataTextColumn Caption="QTY PO" FieldName="QtyPO" VisibleIndex="7" 
                            Width="80px"
                            HeaderStyle-HorizontalAlign="Center">
                            <HeaderStyle HorizontalAlign="Center" Wrap="True"></HeaderStyle>
                            <PropertiesTextEdit DisplayFormatString="{0:n0}">
                            </PropertiesTextEdit>
                            <CellStyle Font-Names="Tahoma" Font-Size="8pt" HorizontalAlign="Right">
                            </CellStyle>
                        </dx:GridViewDataTextColumn>                        
                    </Columns>
                    <SettingsBehavior AllowSelectByRowClick="True" AllowSort="False"
                        ColumnResizeMode="Control" EnableRowHotTrack="True" />
                    <SettingsPager PageSize="100" Position="Top">
                        <Summary Text="Page {0} of {1} [{2} record(s)]" AllPagesText="Page {0} of {1} " />
                        <Summary AllPagesText="Page {0} of {1} " Text="Page {0} of {1} [{2} record(s)]">
                        </Summary>
                    </SettingsPager>
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
                <dx:ASPxButton ID="btnExcel" runat="server" Text="EXCEL"
                    Font-Names="Tahoma"
                    Width="90px" Font-Size="8pt" AutoPostBack="False">
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

