<%@ Page Title="" Language="vb" AutoEventWireup="false" MasterPageFile="~/SiteContent.Master" CodeBehind="PORevHistory.aspx.vb" Inherits="AffiliateSystem.PORevHistory" %>
<%@ Register Assembly="DevExpress.Web.v14.1, Version=14.1.6.0, Culture=neutral, PublicKeyToken=b88d1754d700e49a" namespace="DevExpress.Web.ASPxGridView.Export" TagPrefix="dx" %>
<%@ Register Assembly="DevExpress.Web.v14.1, Version=14.1.6.0, Culture=neutral, PublicKeyToken=b88d1754d700e49a" namespace="DevExpress.Web.ASPxGridView" TagPrefix="dx" %>
<%@ Register assembly="DevExpress.Web.v14.1, Version=14.1.6.0, Culture=neutral, PublicKeyToken=b88d1754d700e49a" namespace="DevExpress.Web.ASPxEditors" tagprefix="dx" %>
<%@ Register Assembly="DevExpress.Web.v14.1, Version=14.1.6.0, Culture=neutral, PublicKeyToken=b88d1754d700e49a" namespace="DevExpress.Web.ASPxGlobalEvents" TagPrefix="dx" %>
<%@ Register Assembly="DevExpress.Web.v14.1, Version=14.1.6.0, Culture=neutral, PublicKeyToken=b88d1754d700e49a" namespace="DevExpress.Web.ASPxPanel" TagPrefix="dx" %>
<%@ Register Assembly="DevExpress.Web.v14.1, Version=14.1.6.0, Culture=neutral, PublicKeyToken=b88d1754d700e49a" namespace="DevExpress.Web.ASPxRoundPanel" TagPrefix="dx" %>
<%@ Register Assembly="DevExpress.Web.v14.1, Version=14.1.6.0, Culture=neutral, PublicKeyToken=b88d1754d700e49a" Namespace="DevExpress.Web.ASPxCallback" TagPrefix="dx" %>

<asp:Content ID="Content1" ContentPlaceHolderID="head" runat="server">
    <style type="text/css">
        .dxeHLC, .dxeHC, .dxeHFC
        {
            display: none;
        }
        .style1
        {
            width: 5px;
            height: 20px;
        }
        .style2
        {
            width: 100px;
            height: 20px;
        }
        .style3
        {
            height: 20px;
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
        height = height - (height * 44 / 100)
        grid.SetHeight(height);
    }

    function clear() {
        txtPartName.SetText('');
    }
</script>
</asp:Content>

<asp:Content ID="Content2" ContentPlaceHolderID="Content" runat="server">
    <table style="width:100%;">
        <tr>
            <td colspan="8">
                <table style="border-width: 1pt thin thin thin; border-style: ridge; border-color:#9598A1; width:100%; height: 70px;">
                    <tr>
                        <td colspan="8" height="20">
                            <table id="Table1">
                                <tr>
                                    <td style="width:5px;"></td>
                                    <td align="left" valign="middle" height="20px" width="120px">
                                        <dx:ASPxLabel ID="ASPxLabel1" runat="server" Text="PERIOD"
                                            Font-Names="Tahoma" Font-Size="8pt">
                                        </dx:ASPxLabel>
                                    </td>
                                    <td align="left" valign="middle" height="20px" width="180px">
                                        <dx:ASPxTimeEdit ID="dtPeriodFrom" runat="server" ClientInstanceName="dtPeriodFrom" 
                                            DisplayFormatString="MMM yyyy" EditFormat="Custom" 
                                            EditFormatString="MMM yyyy" Font-Names="Tahoma" Font-Size="8pt" Width="180px">
                                            <ClientSideEvents ValueChanged="function(s, e) {	                                            
                                                grid.PerformCallback('kosong');
                                                cboPONo.PerformCallback(dtPeriodFrom.GetValue().toString());                                                
	                                            lblInfo.SetText('');                                                
                                            }" />
                                        </dx:ASPxTimeEdit>                                                                                   
                                    </td>
                                    <td align="left" valign="middle" width="250px" class="style3"></td>
                                    <td style="width:5px;"></td>
                                    <td align="left" valign="middle" height="20px" width="180px">                                       
                                    </td>
                                    <td style="width:5px;"></td> 
                                </tr>
                                <tr>
                                    <td style="width:5px;"></td>
                                    <td align="left" valign="middle" height="20px" width="120px">
                                        <dx:ASPxLabel ID="ASPxLabel6" runat="server" Text="PO NO."
                                            Font-Names="Tahoma" Font-Size="8pt">
                                        </dx:ASPxLabel>
                                    </td>
                                    <td align="left" valign="middle" height="20px" width="180px">
                                        <dx:ASPxComboBox ID="cboPONo" runat="server" 
                                            ClientInstanceName="cboPONo" Width="100%"
                                            Font-Size="8pt" 
                                            Font-Names="Verdana" TextFormatString="{0}">
                                            <ClientSideEvents SelectedIndexChanged="function(s, e) {
                                                grid.PerformCallback('kosong');                                                
	                                            lblInfo.SetText('');                                                
                                            }" />
                                            <LoadingPanelStyle ImageSpacing="5px">
                                            </LoadingPanelStyle>
                                        </dx:ASPxComboBox>
                                    </td>
                                    <td align="left" valign="middle" width="250px" class="style3"></td>
                                    <td style="width:5px;"></td>
                                    <td align="left" valign="middle" height="20px" width="180px">                                        
                                    </td>
                                    <td style="width:5px;"></td> 
                                </tr> 
                                <tr>
                                    <td style="width: 5px;"></td>
                                    <td align="left" valign="middle" height="20px" width="120px">
                                        <dx:ASPxLabel ID="ASPxLabel3" runat="server" Text="PARTS NO. / NAME"
                                            Font-Names="Tahoma" Font-Size="8pt">
                                        </dx:ASPxLabel>
                                    </td>
                                    <td align="left" valign="middle" height="20px" width="180px">
                                        <dx:ASPxComboBox ID="cboPartNo" runat="server" 
                                            ClientInstanceName="cboPartNo" Width="100%"
                                            Font-Size="8pt" 
                                            Font-Names="Verdana" TextFormatString="{0}">
                                            <ClientSideEvents SelectedIndexChanged="function(s, e) {
                                                grid.PerformCallback('kosong');
                                                txtPartName.SetText(cboPartNo.GetSelectedItem().GetColumnText(1))
	                                            lblInfo.SetText('');                                                
                                            }" />
                                            <LoadingPanelStyle ImageSpacing="5px">
                                            </LoadingPanelStyle>
                                        </dx:ASPxComboBox>
                                    </td>
                                    <td align="left" valign="middle" width="250px" class="style3">
                                        <dx:ASPxTextBox ID="txtPartName" runat="server" Width="250px" 
                                            ClientInstanceName="txtPartName" Font-Names="Tahoma" Font-Size="8pt"
                                            MaxLength="20" onkeypress="return singlequote(event)" Height="20px" ReadOnly="True">
                                            <readonlystyle backcolor="#CCCCCC">
                                            </readonlystyle>
                                        </dx:ASPxTextBox>
                                    </td>
                                    <td style="width:5px;"></td>  
                                    <td align="left" valign="middle" height="20px" width="180px">
                                        <table style="width:100%;" align="right">
                                            <tr>
                                                <td align="right" valign="middle" style="height:25px; width:90px;">
                                                    <dx:ASPxButton ID="btnSearch" runat="server" Text="SEARCH"
                                                        Font-Names="Tahoma" Width="85px" AutoPostBack="False" Font-Size="8pt" >
                                                        <ClientSideEvents Click="function(s, e) {
                                                            grid.PerformCallback('load');
                                                            lblInfo.SetText('');
                                                        }" />
                                                    </dx:ASPxButton>
                                                </td>
                                                <td align="right" valign="middle" style="height:25px; width:90px;">
                                                    <dx:ASPxButton ID="btnClear" runat="server" Text="CLEAR"
                                                        Font-Names="Tahoma" Width="85px" AutoPostBack="False" Font-Size="8pt">                                 
                                                    </dx:ASPxButton>
                                                </td>
                                            </tr>
                                        </table>
                                    </td>
                                    <td style="width:5px;"></td> 
                                </tr>                             
                            </table>
                        </td>
                    </tr>
                </table>
            </td>            
        </tr>
    
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
    
        <tr>
            <td valign="top" align="left" colspan="7">
                &nbsp;
            </td>
            <td valign="top" align="right" width="170px">
                <table style="width: 100%;">
                    <tr>                                                
                        <td align="right" valign="middle" width="170px">
                            <asp:TextBox ID="lSuuplier0" runat="server" BackColor="Yellow" BorderStyle="None"
                                ReadOnly="True" Width="30px">
                            </asp:TextBox>
                            <dx:ASPxLabel runat="server" Text=": DIFFERENCE" Font-Names="Tahoma" Font-Size="8pt"
                                ID="ASPxLabel2" Width="90px">
                            </dx:ASPxLabel>
                        </td>
                    </tr>
                </table>
            </td>
        </tr>
    
        <tr>
            <td colspan="8" align="left" valign="top" height="220">
                <%--Column : Grid--%>
                <dx:ASPxGridView ID="grid" runat="server" Width="100%"
                    Font-Names="Tahoma" KeyFieldName="PartNo2;AffiliateName;PONo;NoUrut"
                    AutoGenerateColumns="False"
                    ClientInstanceName="grid" Font-Size="8pt">
                    <ClientSideEvents Init="OnInit" FocusedRowChanged="function(s, e) {
	                    OnGridFocusedRowChanged();
                    }" CallbackError="function(s, e) {e.handled = true;}" EndCallback="function(s, e) {
                        if (s.cpSearch != '') {
                            txtUser1.SetText(s.cpUser1);
                            txtUser2.SetText(s.cpUser2);
                            txtUser3.SetText(s.cpUser3);
                            txtUser4.SetText(s.cpUser4);
                            txtUser5.SetText(s.cpUser5);
                            txtUser6.SetText(s.cpUser6);
                            txtUser7.SetText(s.cpUser7);
                            txtUser8.SetText(s.cpUser8);

                            txtDate1.SetText(s.cpDate1);
                            txtDate2.SetText(s.cpDate2);
                            txtDate3.SetText(s.cpDate3);
                            txtDate4.SetText(s.cpDate4);
                            txtDate5.SetText(s.cpDate5);
                            txtDate6.SetText(s.cpDate6);
                            txtDate7.SetText(s.cpDate7);
                            txtDate8.SetText(s.cpDate8);
                        }
                        
                        delete s.cpSearch;
                    }" RowClick="function(s, e) {
	                    lblInfo.SetText('');
                    }" />                    
                    <Columns>
                        <dx:GridViewDataTextColumn VisibleIndex="1" Caption="NO." FieldName="NoUrut" Width="30px" HeaderStyle-HorizontalAlign="Center">
                            <HeaderStyle HorizontalAlign="Center"></HeaderStyle>
                            <CellStyle Font-Names="Tahoma" Font-Size="8pt">
                            </CellStyle>
                        </dx:GridViewDataTextColumn>
                        <dx:GridViewDataTextColumn VisibleIndex="2" Caption="PART NO." FieldName="PartNo" Width="90px" HeaderStyle-HorizontalAlign="Center">
                            <HeaderStyle HorizontalAlign="Center"></HeaderStyle>
                            <CellStyle Font-Names="Tahoma" Font-Size="8pt">
                            </CellStyle>
                        </dx:GridViewDataTextColumn>
                        <dx:GridViewDataTextColumn VisibleIndex="3" Caption="PART NAME" FieldName="PartName" Width="180px" HeaderStyle-HorizontalAlign="Center">
                            <HeaderStyle HorizontalAlign="Center"></HeaderStyle>
                            <CellStyle Font-Names="Tahoma" Font-Size="8pt">
                            </CellStyle>
                        </dx:GridViewDataTextColumn>
<%--                    <dx:GridViewDataTextColumn VisibleIndex="4" Caption="KANBAN CLS" FieldName="KanbanCls" Width="60px" HeaderStyle-HorizontalAlign="Center" CellStyle-HorizontalAlign="Center">
                            <HeaderStyle HorizontalAlign="Center"></HeaderStyle>
                            <CellStyle Font-Names="Tahoma" Font-Size="8pt">
                            </CellStyle>
                        </dx:GridViewDataTextColumn>--%>
                        <dx:GridViewDataTextColumn VisibleIndex="5" Caption="UOM" FieldName="UnitDesc" Width="40px" HeaderStyle-HorizontalAlign="Center" CellStyle-HorizontalAlign="Center">
                            <HeaderStyle HorizontalAlign="Center"></HeaderStyle>
                            <CellStyle Font-Names="Tahoma" Font-Size="8pt">
                            </CellStyle>
                        </dx:GridViewDataTextColumn>
                        <dx:GridViewDataTextColumn VisibleIndex="6" Caption="MOQ" FieldName="MOQ" Width="70px" HeaderStyle-HorizontalAlign="Center">
                            <HeaderStyle HorizontalAlign="Center"></HeaderStyle>
                            <CellStyle Font-Names="Tahoma" Font-Size="8pt">
                            </CellStyle>
                        </dx:GridViewDataTextColumn>
                        
                        <dx:GridViewDataTextColumn VisibleIndex="7" Caption="QTY/BOX" FieldName="QtyBox" Width="70px" HeaderStyle-HorizontalAlign="Center">
                            <PropertiesTextEdit DisplayFormatString="{0:n0}">
                                <MaskSettings ErrorText="Please input valid value !"
                                Mask="<0..999999g>" IncludeLiterals="DecimalSymbol" />
                            </PropertiesTextEdit>
                            <HeaderStyle HorizontalAlign="Center"></HeaderStyle>
                            <CellStyle Font-Names="Tahoma" Font-Size="8pt">
                            </CellStyle>
                        </dx:GridViewDataTextColumn>  
                        <dx:GridViewDataTextColumn VisibleIndex="8" Caption="MAKER" FieldName="Maker" Width="100px" HeaderStyle-HorizontalAlign="Center">
                            <HeaderStyle HorizontalAlign="Center"></HeaderStyle>
                            <CellStyle Font-Names="Tahoma" Font-Size="8pt">
                            </CellStyle>
                        </dx:GridViewDataTextColumn>

                        <dx:GridViewDataTextColumn VisibleIndex="9" Caption=" " FieldName="AffiliateName" Width="90px" HeaderStyle-HorizontalAlign="Center">
                            <HeaderStyle HorizontalAlign="Center"></HeaderStyle>
                            <CellStyle Font-Names="Tahoma" Font-Size="8pt">
                            </CellStyle>
                        </dx:GridViewDataTextColumn>

                        <dx:GridViewDataTextColumn VisibleIndex="10" Caption="PO No." FieldName="PONo" Width="120px" HeaderStyle-HorizontalAlign="Center" CellStyle-HorizontalAlign="Center">
                            <HeaderStyle HorizontalAlign="Center"></HeaderStyle>
                            <CellStyle Font-Names="Tahoma" Font-Size="8pt">
                            </CellStyle>
                        </dx:GridViewDataTextColumn>

                        <dx:GridViewDataTextColumn VisibleIndex="11" Caption="TOTAL FIRM QTY" 
                            FieldName="POQty" Width="100px" HeaderStyle-HorizontalAlign="Center">
                            <PropertiesTextEdit DisplayFormatString="{0:n0}">
                                <MaskSettings ErrorText="Please input valid value !"
                                Mask="<0..999999g>" IncludeLiterals="DecimalSymbol" />
                            </PropertiesTextEdit>
                            <HeaderStyle HorizontalAlign="Center"></HeaderStyle>
                            <CellStyle Font-Names="Tahoma" Font-Size="8pt">
                            </CellStyle>
                        </dx:GridViewDataTextColumn>

                        <dx:GridViewDataTextColumn VisibleIndex="12" Caption="TOTAL FIRM QTY" 
                            FieldName="POQtyOld" Width="0px" HeaderStyle-HorizontalAlign="Center">
                            <PropertiesTextEdit DisplayFormatString="{0:n0}">
                                <MaskSettings ErrorText="Please input valid value !"
                                Mask="<0..999999g>" IncludeLiterals="DecimalSymbol" />
                            </PropertiesTextEdit>
                            <HeaderStyle HorizontalAlign="Center"></HeaderStyle>
                            <CellStyle Font-Names="Tahoma" Font-Size="8pt">
                            </CellStyle>
                        </dx:GridViewDataTextColumn>

                        <%--<dx:GridViewDataTextColumn VisibleIndex="15" Caption="FORECAST N+1" FieldName="ForecastN1" Width="80px" HeaderStyle-HorizontalAlign="Center">
                            <PropertiesTextEdit DisplayFormatString="{0:n0}">
                                <MaskSettings ErrorText="Please input valid value !"
                                Mask="<0..999999g>" IncludeLiterals="DecimalSymbol" />
                            </PropertiesTextEdit>
                            <HeaderStyle HorizontalAlign="Center"></HeaderStyle>
                            <CellStyle Font-Names="Tahoma" Font-Size="8pt">
                            </CellStyle>
                        </dx:GridViewDataTextColumn>
                        <dx:GridViewDataTextColumn VisibleIndex="16" Caption="FORECAST N+2" FieldName="ForecastN2" Width="80px" HeaderStyle-HorizontalAlign="Center">
                            <PropertiesTextEdit DisplayFormatString="{0:n0}">
                                <MaskSettings ErrorText="Please input valid value !"
                                Mask="<0..999999g>" IncludeLiterals="DecimalSymbol" />
                            </PropertiesTextEdit>
                            <HeaderStyle HorizontalAlign="Center"></HeaderStyle>
                            <CellStyle Font-Names="Tahoma" Font-Size="8pt">
                            </CellStyle>
                        </dx:GridViewDataTextColumn>
                        <dx:GridViewDataTextColumn VisibleIndex="17" Caption="FORECAST N+3" FieldName="ForecastN3" Width="80px" HeaderStyle-HorizontalAlign="Center">
                            <PropertiesTextEdit DisplayFormatString="{0:n0}">
                                <MaskSettings ErrorText="Please input valid value !"
                                Mask="<0..999999g>" IncludeLiterals="DecimalSymbol" />
                            </PropertiesTextEdit>
                            <HeaderStyle HorizontalAlign="Center"></HeaderStyle>
                            <CellStyle Font-Names="Tahoma" Font-Size="8pt">
                            </CellStyle>
                        </dx:GridViewDataTextColumn>--%>

                        <dx:GridViewBandColumn Caption="E.T.A SCHEDULE (BASED ON FIRM ORDER)" VisibleIndex="17" HeaderStyle-HorizontalAlign="Center">
                            <Columns>
                                <dx:GridViewDataTextColumn VisibleIndex="18" Caption="1" Width="60px" FieldName="DeliveryD1" HeaderStyle-HorizontalAlign="Center">
                                    <PropertiesTextEdit DisplayFormatString="{0:n0}">
                                        <MaskSettings ErrorText="Please input valid value !"
                                        Mask="<0..999999g>" IncludeLiterals="DecimalSymbol" />
                                    </PropertiesTextEdit>
                                    <HeaderStyle HorizontalAlign="Center"></HeaderStyle>
                                    <CellStyle Font-Names="Tahoma" Font-Size="8pt">
                                    </CellStyle>
                                </dx:GridViewDataTextColumn>
                                <dx:GridViewDataTextColumn VisibleIndex="19" Caption="1" Width="0px" FieldName="DeliveryD1Old" HeaderStyle-HorizontalAlign="Center">
                                    <PropertiesTextEdit DisplayFormatString="{0:n0}">
                                        <MaskSettings ErrorText="Please input valid value !"
                                        Mask="<0..999999g>" IncludeLiterals="DecimalSymbol" />
                                    </PropertiesTextEdit>
                                    <HeaderStyle HorizontalAlign="Center"></HeaderStyle>
                                    <CellStyle Font-Names="Tahoma" Font-Size="8pt">
                                    </CellStyle>
                                </dx:GridViewDataTextColumn>
                                <dx:GridViewDataTextColumn VisibleIndex="20" Caption="2" Width="60px" FieldName="DeliveryD2" HeaderStyle-HorizontalAlign="Center">
                                    <PropertiesTextEdit DisplayFormatString="{0:n0}">
                                        <MaskSettings ErrorText="Please input valid value !"
                                        Mask="<0..999999g>" IncludeLiterals="DecimalSymbol" />
                                    </PropertiesTextEdit>
                                    <HeaderStyle HorizontalAlign="Center"></HeaderStyle>
                                    <CellStyle Font-Names="Tahoma" Font-Size="8pt">
                                    </CellStyle>
                                </dx:GridViewDataTextColumn>
                                <dx:GridViewDataTextColumn VisibleIndex="21" Caption="2" Width="0px" FieldName="DeliveryD2Old" HeaderStyle-HorizontalAlign="Center">
                                    <PropertiesTextEdit DisplayFormatString="{0:n0}">
                                        <MaskSettings ErrorText="Please input valid value !"
                                        Mask="<0..999999g>" IncludeLiterals="DecimalSymbol" />
                                    </PropertiesTextEdit>
                                    <HeaderStyle HorizontalAlign="Center"></HeaderStyle>
                                    <CellStyle Font-Names="Tahoma" Font-Size="8pt">
                                    </CellStyle>
                                </dx:GridViewDataTextColumn>
                                <dx:GridViewDataTextColumn VisibleIndex="22" Caption="3" Width="60px" FieldName="DeliveryD3" HeaderStyle-HorizontalAlign="Center">
                                    <PropertiesTextEdit DisplayFormatString="{0:n0}">
                                        <MaskSettings ErrorText="Please input valid value !"
                                        Mask="<0..999999g>" IncludeLiterals="DecimalSymbol" />
                                    </PropertiesTextEdit>
                                    <HeaderStyle HorizontalAlign="Center"></HeaderStyle>
                                    <CellStyle Font-Names="Tahoma" Font-Size="8pt">
                                    </CellStyle>
                                </dx:GridViewDataTextColumn>
                                <dx:GridViewDataTextColumn VisibleIndex="23" Caption="3" Width="0px" FieldName="DeliveryD3Old" HeaderStyle-HorizontalAlign="Center">
                                    <PropertiesTextEdit DisplayFormatString="{0:n0}">
                                        <MaskSettings ErrorText="Please input valid value !"
                                        Mask="<0..999999g>" IncludeLiterals="DecimalSymbol" />
                                    </PropertiesTextEdit>
                                    <HeaderStyle HorizontalAlign="Center"></HeaderStyle>
                                    <CellStyle Font-Names="Tahoma" Font-Size="8pt">
                                    </CellStyle>
                                </dx:GridViewDataTextColumn>
                                <dx:GridViewDataTextColumn VisibleIndex="24" Caption="4" Width="60px" FieldName="DeliveryD4" HeaderStyle-HorizontalAlign="Center">
                                    <PropertiesTextEdit DisplayFormatString="{0:n0}">
                                        <MaskSettings ErrorText="Please input valid value !"
                                        Mask="<0..999999g>" IncludeLiterals="DecimalSymbol" />
                                    </PropertiesTextEdit>
                                    <HeaderStyle HorizontalAlign="Center"></HeaderStyle>
                                    <CellStyle Font-Names="Tahoma" Font-Size="8pt">
                                    </CellStyle>
                                </dx:GridViewDataTextColumn>
                                <dx:GridViewDataTextColumn VisibleIndex="25" Caption="4" Width="0px" FieldName="DeliveryD4Old" HeaderStyle-HorizontalAlign="Center">
                                    <PropertiesTextEdit DisplayFormatString="{0:n0}">
                                        <MaskSettings ErrorText="Please input valid value !"
                                        Mask="<0..999999g>" IncludeLiterals="DecimalSymbol" />
                                    </PropertiesTextEdit>
                                    <HeaderStyle HorizontalAlign="Center"></HeaderStyle>
                                    <CellStyle Font-Names="Tahoma" Font-Size="8pt">
                                    </CellStyle>
                                </dx:GridViewDataTextColumn>
                                <dx:GridViewDataTextColumn VisibleIndex="26" Caption="5" Width="60px" FieldName="DeliveryD5" HeaderStyle-HorizontalAlign="Center">
                                    <PropertiesTextEdit DisplayFormatString="{0:n0}">
                                        <MaskSettings ErrorText="Please input valid value !"
                                        Mask="<0..999999g>" IncludeLiterals="DecimalSymbol" />
                                    </PropertiesTextEdit>
                                    <HeaderStyle HorizontalAlign="Center"></HeaderStyle>
                                    <CellStyle Font-Names="Tahoma" Font-Size="8pt">
                                    </CellStyle>
                                </dx:GridViewDataTextColumn>
                                <dx:GridViewDataTextColumn VisibleIndex="27" Caption="5" Width="0px" FieldName="DeliveryD5Old" HeaderStyle-HorizontalAlign="Center">
                                    <PropertiesTextEdit DisplayFormatString="{0:n0}">
                                        <MaskSettings ErrorText="Please input valid value !"
                                        Mask="<0..999999g>" IncludeLiterals="DecimalSymbol" />
                                    </PropertiesTextEdit>
                                    <HeaderStyle HorizontalAlign="Center"></HeaderStyle>
                                    <CellStyle Font-Names="Tahoma" Font-Size="8pt">
                                    </CellStyle>
                                </dx:GridViewDataTextColumn>
                                <dx:GridViewDataTextColumn VisibleIndex="28" Caption="6" Width="60px" FieldName="DeliveryD6" HeaderStyle-HorizontalAlign="Center">
                                    <PropertiesTextEdit DisplayFormatString="{0:n0}">
                                        <MaskSettings ErrorText="Please input valid value !"
                                        Mask="<0..999999g>" IncludeLiterals="DecimalSymbol" />
                                    </PropertiesTextEdit>
                                    <HeaderStyle HorizontalAlign="Center"></HeaderStyle>
                                    <CellStyle Font-Names="Tahoma" Font-Size="8pt">
                                    </CellStyle>
                                </dx:GridViewDataTextColumn>
                                <dx:GridViewDataTextColumn VisibleIndex="29" Caption="6" Width="0px" FieldName="DeliveryD6Old" HeaderStyle-HorizontalAlign="Center">
                                    <PropertiesTextEdit DisplayFormatString="{0:n0}">
                                        <MaskSettings ErrorText="Please input valid value !"
                                        Mask="<0..999999g>" IncludeLiterals="DecimalSymbol" />
                                    </PropertiesTextEdit>
                                    <HeaderStyle HorizontalAlign="Center"></HeaderStyle>
                                    <CellStyle Font-Names="Tahoma" Font-Size="8pt">
                                    </CellStyle>
                                </dx:GridViewDataTextColumn>
                                <dx:GridViewDataTextColumn VisibleIndex="30" Caption="7" Width="60px" FieldName="DeliveryD7" HeaderStyle-HorizontalAlign="Center">
                                    <PropertiesTextEdit DisplayFormatString="{0:n0}">
                                        <MaskSettings ErrorText="Please input valid value !"
                                        Mask="<0..999999g>" IncludeLiterals="DecimalSymbol" />
                                    </PropertiesTextEdit>
                                    <HeaderStyle HorizontalAlign="Center"></HeaderStyle>
                                    <CellStyle Font-Names="Tahoma" Font-Size="8pt">
                                    </CellStyle>
                                </dx:GridViewDataTextColumn>
                                <dx:GridViewDataTextColumn VisibleIndex="31" Caption="7" Width="0px" FieldName="DeliveryD7Old" HeaderStyle-HorizontalAlign="Center">
                                    <PropertiesTextEdit DisplayFormatString="{0:n0}">
                                        <MaskSettings ErrorText="Please input valid value !"
                                        Mask="<0..999999g>" IncludeLiterals="DecimalSymbol" />
                                    </PropertiesTextEdit>
                                    <HeaderStyle HorizontalAlign="Center"></HeaderStyle>
                                    <CellStyle Font-Names="Tahoma" Font-Size="8pt">
                                    </CellStyle>
                                </dx:GridViewDataTextColumn>
                                <dx:GridViewDataTextColumn VisibleIndex="32" Caption="8" Width="60px" FieldName="DeliveryD8" HeaderStyle-HorizontalAlign="Center">
                                    <PropertiesTextEdit DisplayFormatString="{0:n0}">
                                        <MaskSettings ErrorText="Please input valid value !"
                                        Mask="<0..999999g>" IncludeLiterals="DecimalSymbol" />
                                    </PropertiesTextEdit>
                                    <HeaderStyle HorizontalAlign="Center"></HeaderStyle>
                                    <CellStyle Font-Names="Tahoma" Font-Size="8pt">
                                    </CellStyle>
                                </dx:GridViewDataTextColumn>
                                <dx:GridViewDataTextColumn VisibleIndex="33" Caption="8" Width="0px" FieldName="DeliveryD8Old" HeaderStyle-HorizontalAlign="Center">
                                    <PropertiesTextEdit DisplayFormatString="{0:n0}">
                                        <MaskSettings ErrorText="Please input valid value !"
                                        Mask="<0..999999g>" IncludeLiterals="DecimalSymbol" />
                                    </PropertiesTextEdit>
                                    <HeaderStyle HorizontalAlign="Center"></HeaderStyle>
                                    <CellStyle Font-Names="Tahoma" Font-Size="8pt">
                                    </CellStyle>
                                </dx:GridViewDataTextColumn>
                                <dx:GridViewDataTextColumn VisibleIndex="34" Caption="9" Width="60px" FieldName="DeliveryD9" HeaderStyle-HorizontalAlign="Center">
                                    <PropertiesTextEdit DisplayFormatString="{0:n0}">
                                        <MaskSettings ErrorText="Please input valid value !"
                                        Mask="<0..999999g>" IncludeLiterals="DecimalSymbol" />
                                    </PropertiesTextEdit>
                                    <HeaderStyle HorizontalAlign="Center"></HeaderStyle>
                                    <CellStyle Font-Names="Tahoma" Font-Size="8pt">
                                    </CellStyle>
                                </dx:GridViewDataTextColumn>
                                <dx:GridViewDataTextColumn VisibleIndex="35" Caption="9" Width="0px" FieldName="DeliveryD9Old" HeaderStyle-HorizontalAlign="Center">
                                    <PropertiesTextEdit DisplayFormatString="{0:n0}">
                                        <MaskSettings ErrorText="Please input valid value !"
                                        Mask="<0..999999g>" IncludeLiterals="DecimalSymbol" />
                                    </PropertiesTextEdit>
                                    <HeaderStyle HorizontalAlign="Center"></HeaderStyle>
                                    <CellStyle Font-Names="Tahoma" Font-Size="8pt">
                                    </CellStyle>
                                </dx:GridViewDataTextColumn>
                                <dx:GridViewDataTextColumn VisibleIndex="36" Caption="10" Width="60px" FieldName="DeliveryD10" HeaderStyle-HorizontalAlign="Center">
                                    <PropertiesTextEdit DisplayFormatString="{0:n0}">
                                        <MaskSettings ErrorText="Please input valid value !"
                                        Mask="<0..999999g>" IncludeLiterals="DecimalSymbol" />
                                    </PropertiesTextEdit>
                                    <HeaderStyle HorizontalAlign="Center"></HeaderStyle>
                                    <CellStyle Font-Names="Tahoma" Font-Size="8pt">
                                    </CellStyle>
                                </dx:GridViewDataTextColumn>
                                <dx:GridViewDataTextColumn VisibleIndex="37" Caption="10" Width="0px" FieldName="DeliveryD10Old" HeaderStyle-HorizontalAlign="Center">
                                    <PropertiesTextEdit DisplayFormatString="{0:n0}">
                                        <MaskSettings ErrorText="Please input valid value !"
                                        Mask="<0..999999g>" IncludeLiterals="DecimalSymbol" />
                                    </PropertiesTextEdit>
                                    <HeaderStyle HorizontalAlign="Center"></HeaderStyle>
                                    <CellStyle Font-Names="Tahoma" Font-Size="8pt">
                                    </CellStyle>
                                </dx:GridViewDataTextColumn>
                                <dx:GridViewDataTextColumn VisibleIndex="38" Caption="11" Width="60px" FieldName="DeliveryD11" HeaderStyle-HorizontalAlign="Center">
                                    <PropertiesTextEdit DisplayFormatString="{0:n0}">
                                        <MaskSettings ErrorText="Please input valid value !"
                                        Mask="<0..999999g>" IncludeLiterals="DecimalSymbol" />
                                    </PropertiesTextEdit>
                                    <HeaderStyle HorizontalAlign="Center"></HeaderStyle>
                                    <CellStyle Font-Names="Tahoma" Font-Size="8pt">
                                    </CellStyle>
                                </dx:GridViewDataTextColumn>
                                <dx:GridViewDataTextColumn VisibleIndex="39" Caption="11" Width="0px" FieldName="DeliveryD11Old" HeaderStyle-HorizontalAlign="Center">
                                    <PropertiesTextEdit DisplayFormatString="{0:n0}">
                                        <MaskSettings ErrorText="Please input valid value !"
                                        Mask="<0..999999g>" IncludeLiterals="DecimalSymbol" />
                                    </PropertiesTextEdit>
                                    <HeaderStyle HorizontalAlign="Center"></HeaderStyle>
                                    <CellStyle Font-Names="Tahoma" Font-Size="8pt">
                                    </CellStyle>
                                </dx:GridViewDataTextColumn>
                                <dx:GridViewDataTextColumn VisibleIndex="40" Caption="12" Width="60px" FieldName="DeliveryD12" HeaderStyle-HorizontalAlign="Center">
                                    <PropertiesTextEdit DisplayFormatString="{0:n0}">
                                        <MaskSettings ErrorText="Please input valid value !"
                                        Mask="<0..999999g>" IncludeLiterals="DecimalSymbol" />
                                    </PropertiesTextEdit>
                                    <HeaderStyle HorizontalAlign="Center"></HeaderStyle>
                                    <CellStyle Font-Names="Tahoma" Font-Size="8pt">
                                    </CellStyle>
                                </dx:GridViewDataTextColumn>
                                <dx:GridViewDataTextColumn VisibleIndex="41" Caption="12" Width="0px" FieldName="DeliveryD12Old" HeaderStyle-HorizontalAlign="Center">
                                    <PropertiesTextEdit DisplayFormatString="{0:n0}">
                                        <MaskSettings ErrorText="Please input valid value !"
                                        Mask="<0..999999g>" IncludeLiterals="DecimalSymbol" />
                                    </PropertiesTextEdit>
                                    <HeaderStyle HorizontalAlign="Center"></HeaderStyle>
                                    <CellStyle Font-Names="Tahoma" Font-Size="8pt">
                                    </CellStyle>
                                </dx:GridViewDataTextColumn>
                                <dx:GridViewDataTextColumn VisibleIndex="42" Caption="13" Width="60px" FieldName="DeliveryD13" HeaderStyle-HorizontalAlign="Center">
                                    <PropertiesTextEdit DisplayFormatString="{0:n0}">
                                        <MaskSettings ErrorText="Please input valid value !"
                                        Mask="<0..999999g>" IncludeLiterals="DecimalSymbol" />
                                    </PropertiesTextEdit>
                                    <HeaderStyle HorizontalAlign="Center"></HeaderStyle>
                                    <CellStyle Font-Names="Tahoma" Font-Size="8pt">
                                    </CellStyle>
                                </dx:GridViewDataTextColumn>
                                <dx:GridViewDataTextColumn VisibleIndex="43" Caption="13" Width="0px" FieldName="DeliveryD13Old" HeaderStyle-HorizontalAlign="Center">
                                    <PropertiesTextEdit DisplayFormatString="{0:n0}">
                                        <MaskSettings ErrorText="Please input valid value !"
                                        Mask="<0..999999g>" IncludeLiterals="DecimalSymbol" />
                                    </PropertiesTextEdit>
                                    <HeaderStyle HorizontalAlign="Center"></HeaderStyle>
                                    <CellStyle Font-Names="Tahoma" Font-Size="8pt">
                                    </CellStyle>
                                </dx:GridViewDataTextColumn>
                                <dx:GridViewDataTextColumn VisibleIndex="44" Caption="14" Width="60px" FieldName="DeliveryD14" HeaderStyle-HorizontalAlign="Center">
                                    <PropertiesTextEdit DisplayFormatString="{0:n0}">
                                        <MaskSettings ErrorText="Please input valid value !"
                                        Mask="<0..999999g>" IncludeLiterals="DecimalSymbol" />
                                    </PropertiesTextEdit>
                                    <HeaderStyle HorizontalAlign="Center"></HeaderStyle>
                                    <CellStyle Font-Names="Tahoma" Font-Size="8pt">
                                    </CellStyle>
                                </dx:GridViewDataTextColumn>
                                <dx:GridViewDataTextColumn VisibleIndex="45" Caption="14" Width="0px" FieldName="DeliveryD14Old" HeaderStyle-HorizontalAlign="Center">
                                    <PropertiesTextEdit DisplayFormatString="{0:n0}">
                                        <MaskSettings ErrorText="Please input valid value !"
                                        Mask="<0..999999g>" IncludeLiterals="DecimalSymbol" />
                                    </PropertiesTextEdit>
                                    <HeaderStyle HorizontalAlign="Center"></HeaderStyle>
                                    <CellStyle Font-Names="Tahoma" Font-Size="8pt">
                                    </CellStyle>
                                </dx:GridViewDataTextColumn>
                                <dx:GridViewDataTextColumn VisibleIndex="46" Caption="15" Width="60px" FieldName="DeliveryD15" HeaderStyle-HorizontalAlign="Center">
                                    <PropertiesTextEdit DisplayFormatString="{0:n0}">
                                        <MaskSettings ErrorText="Please input valid value !"
                                        Mask="<0..999999g>" IncludeLiterals="DecimalSymbol" />
                                    </PropertiesTextEdit>
                                    <HeaderStyle HorizontalAlign="Center"></HeaderStyle>
                                    <CellStyle Font-Names="Tahoma" Font-Size="8pt">
                                    </CellStyle>
                                </dx:GridViewDataTextColumn>
                                <dx:GridViewDataTextColumn VisibleIndex="47" Caption="15" Width="0px" FieldName="DeliveryD15Old" HeaderStyle-HorizontalAlign="Center">
                                    <PropertiesTextEdit DisplayFormatString="{0:n0}">
                                        <MaskSettings ErrorText="Please input valid value !"
                                        Mask="<0..999999g>" IncludeLiterals="DecimalSymbol" />
                                    </PropertiesTextEdit>
                                    <HeaderStyle HorizontalAlign="Center"></HeaderStyle>
                                    <CellStyle Font-Names="Tahoma" Font-Size="8pt">
                                    </CellStyle>
                                </dx:GridViewDataTextColumn>
                                <dx:GridViewDataTextColumn VisibleIndex="48" Caption="16" Width="60px" FieldName="DeliveryD16" HeaderStyle-HorizontalAlign="Center">
                                    <PropertiesTextEdit DisplayFormatString="{0:n0}">
                                        <MaskSettings ErrorText="Please input valid value !"
                                        Mask="<0..999999g>" IncludeLiterals="DecimalSymbol" />
                                    </PropertiesTextEdit>
                                    <HeaderStyle HorizontalAlign="Center"></HeaderStyle>
                                    <CellStyle Font-Names="Tahoma" Font-Size="8pt">
                                    </CellStyle>
                                </dx:GridViewDataTextColumn>
                                <dx:GridViewDataTextColumn VisibleIndex="49" Caption="16" Width="0px" FieldName="DeliveryD16Old" HeaderStyle-HorizontalAlign="Center">
                                    <PropertiesTextEdit DisplayFormatString="{0:n0}">
                                        <MaskSettings ErrorText="Please input valid value !"
                                        Mask="<0..999999g>" IncludeLiterals="DecimalSymbol" />
                                    </PropertiesTextEdit>
                                    <HeaderStyle HorizontalAlign="Center"></HeaderStyle>
                                    <CellStyle Font-Names="Tahoma" Font-Size="8pt">
                                    </CellStyle>
                                </dx:GridViewDataTextColumn>
                                <dx:GridViewDataTextColumn VisibleIndex="50" Caption="17" Width="60px" FieldName="DeliveryD17" HeaderStyle-HorizontalAlign="Center">
                                    <PropertiesTextEdit DisplayFormatString="{0:n0}">
                                        <MaskSettings ErrorText="Please input valid value !"
                                        Mask="<0..999999g>" IncludeLiterals="DecimalSymbol" />
                                    </PropertiesTextEdit>
                                    <HeaderStyle HorizontalAlign="Center"></HeaderStyle>
                                    <CellStyle Font-Names="Tahoma" Font-Size="8pt">
                                    </CellStyle>
                                </dx:GridViewDataTextColumn>
                                <dx:GridViewDataTextColumn VisibleIndex="51" Caption="17" Width="0px" FieldName="DeliveryD17Old" HeaderStyle-HorizontalAlign="Center">
                                    <PropertiesTextEdit DisplayFormatString="{0:n0}">
                                        <MaskSettings ErrorText="Please input valid value !"
                                        Mask="<0..999999g>" IncludeLiterals="DecimalSymbol" />
                                    </PropertiesTextEdit>
                                    <HeaderStyle HorizontalAlign="Center"></HeaderStyle>
                                    <CellStyle Font-Names="Tahoma" Font-Size="8pt">
                                    </CellStyle>
                                </dx:GridViewDataTextColumn>
                                <dx:GridViewDataTextColumn VisibleIndex="52" Caption="18" Width="60px" FieldName="DeliveryD18" HeaderStyle-HorizontalAlign="Center">
                                    <PropertiesTextEdit DisplayFormatString="{0:n0}">
                                        <MaskSettings ErrorText="Please input valid value !"
                                        Mask="<0..999999g>" IncludeLiterals="DecimalSymbol" />
                                    </PropertiesTextEdit>
                                    <HeaderStyle HorizontalAlign="Center"></HeaderStyle>
                                    <CellStyle Font-Names="Tahoma" Font-Size="8pt">
                                    </CellStyle>
                                </dx:GridViewDataTextColumn>
                                <dx:GridViewDataTextColumn VisibleIndex="53" Caption="18" Width="0px" FieldName="DeliveryD18Old" HeaderStyle-HorizontalAlign="Center">
                                    <PropertiesTextEdit DisplayFormatString="{0:n0}">
                                        <MaskSettings ErrorText="Please input valid value !"
                                        Mask="<0..999999g>" IncludeLiterals="DecimalSymbol" />
                                    </PropertiesTextEdit>
                                    <HeaderStyle HorizontalAlign="Center"></HeaderStyle>
                                    <CellStyle Font-Names="Tahoma" Font-Size="8pt">
                                    </CellStyle>
                                </dx:GridViewDataTextColumn>
                                <dx:GridViewDataTextColumn VisibleIndex="54" Caption="19" Width="60px" FieldName="DeliveryD19" HeaderStyle-HorizontalAlign="Center">
                                    <PropertiesTextEdit DisplayFormatString="{0:n0}">
                                        <MaskSettings ErrorText="Please input valid value !"
                                        Mask="<0..999999g>" IncludeLiterals="DecimalSymbol" />
                                    </PropertiesTextEdit>
                                    <HeaderStyle HorizontalAlign="Center"></HeaderStyle>
                                    <CellStyle Font-Names="Tahoma" Font-Size="8pt">
                                    </CellStyle>
                                </dx:GridViewDataTextColumn>
                                <dx:GridViewDataTextColumn VisibleIndex="55" Caption="19" Width="0px" FieldName="DeliveryD19Old" HeaderStyle-HorizontalAlign="Center">
                                    <PropertiesTextEdit DisplayFormatString="{0:n0}">
                                        <MaskSettings ErrorText="Please input valid value !"
                                        Mask="<0..999999g>" IncludeLiterals="DecimalSymbol" />
                                    </PropertiesTextEdit>
                                    <HeaderStyle HorizontalAlign="Center"></HeaderStyle>
                                    <CellStyle Font-Names="Tahoma" Font-Size="8pt">
                                    </CellStyle>
                                </dx:GridViewDataTextColumn>
                                <dx:GridViewDataTextColumn VisibleIndex="56" Caption="20" Width="60px" FieldName="DeliveryD20" HeaderStyle-HorizontalAlign="Center">
                                    <PropertiesTextEdit DisplayFormatString="{0:n0}">
                                        <MaskSettings ErrorText="Please input valid value !"
                                        Mask="<0..999999g>" IncludeLiterals="DecimalSymbol" />
                                    </PropertiesTextEdit>
                                    <HeaderStyle HorizontalAlign="Center"></HeaderStyle>
                                    <CellStyle Font-Names="Tahoma" Font-Size="8pt">
                                    </CellStyle>
                                </dx:GridViewDataTextColumn>
                                <dx:GridViewDataTextColumn VisibleIndex="57" Caption="20" Width="0px" FieldName="DeliveryD20Old" HeaderStyle-HorizontalAlign="Center">
                                    <PropertiesTextEdit DisplayFormatString="{0:n0}">
                                        <MaskSettings ErrorText="Please input valid value !"
                                        Mask="<0..999999g>" IncludeLiterals="DecimalSymbol" />
                                    </PropertiesTextEdit>
                                    <HeaderStyle HorizontalAlign="Center"></HeaderStyle>
                                    <CellStyle Font-Names="Tahoma" Font-Size="8pt">
                                    </CellStyle>
                                </dx:GridViewDataTextColumn>
                                <dx:GridViewDataTextColumn VisibleIndex="58" Caption="21" Width="60px" FieldName="DeliveryD21" HeaderStyle-HorizontalAlign="Center">
                                    <PropertiesTextEdit DisplayFormatString="{0:n0}">
                                        <MaskSettings ErrorText="Please input valid value !"
                                        Mask="<0..999999g>" IncludeLiterals="DecimalSymbol" />
                                    </PropertiesTextEdit>
                                    <HeaderStyle HorizontalAlign="Center"></HeaderStyle>
                                    <CellStyle Font-Names="Tahoma" Font-Size="8pt">
                                    </CellStyle>
                                </dx:GridViewDataTextColumn>
                                <dx:GridViewDataTextColumn VisibleIndex="59" Caption="21" Width="0px" FieldName="DeliveryD21Old" HeaderStyle-HorizontalAlign="Center">
                                    <PropertiesTextEdit DisplayFormatString="{0:n0}">
                                        <MaskSettings ErrorText="Please input valid value !"
                                        Mask="<0..999999g>" IncludeLiterals="DecimalSymbol" />
                                    </PropertiesTextEdit>
                                    <HeaderStyle HorizontalAlign="Center"></HeaderStyle>
                                    <CellStyle Font-Names="Tahoma" Font-Size="8pt">
                                    </CellStyle>
                                </dx:GridViewDataTextColumn>
                                <dx:GridViewDataTextColumn VisibleIndex="60" Caption="22" Width="60px" FieldName="DeliveryD22" HeaderStyle-HorizontalAlign="Center">
                                    <PropertiesTextEdit DisplayFormatString="{0:n0}">
                                        <MaskSettings ErrorText="Please input valid value !"
                                        Mask="<0..999999g>" IncludeLiterals="DecimalSymbol" />
                                    </PropertiesTextEdit>
                                    <HeaderStyle HorizontalAlign="Center"></HeaderStyle>
                                    <CellStyle Font-Names="Tahoma" Font-Size="8pt">
                                    </CellStyle>
                                </dx:GridViewDataTextColumn>
                                <dx:GridViewDataTextColumn VisibleIndex="61" Caption="22" Width="0px" FieldName="DeliveryD22Old" HeaderStyle-HorizontalAlign="Center">
                                    <PropertiesTextEdit DisplayFormatString="{0:n0}">
                                        <MaskSettings ErrorText="Please input valid value !"
                                        Mask="<0..999999g>" IncludeLiterals="DecimalSymbol" />
                                    </PropertiesTextEdit>
                                    <HeaderStyle HorizontalAlign="Center"></HeaderStyle>
                                    <CellStyle Font-Names="Tahoma" Font-Size="8pt">
                                    </CellStyle>
                                </dx:GridViewDataTextColumn>
                                <dx:GridViewDataTextColumn VisibleIndex="62" Caption="23" Width="60px" FieldName="DeliveryD23" HeaderStyle-HorizontalAlign="Center">
                                    <PropertiesTextEdit DisplayFormatString="{0:n0}">
                                        <MaskSettings ErrorText="Please input valid value !"
                                        Mask="<0..999999g>" IncludeLiterals="DecimalSymbol" />
                                    </PropertiesTextEdit>
                                    <HeaderStyle HorizontalAlign="Center"></HeaderStyle>
                                    <CellStyle Font-Names="Tahoma" Font-Size="8pt">
                                    </CellStyle>
                                </dx:GridViewDataTextColumn>
                                <dx:GridViewDataTextColumn VisibleIndex="63" Caption="23" Width="0px" FieldName="DeliveryD23Old" HeaderStyle-HorizontalAlign="Center">
                                    <PropertiesTextEdit DisplayFormatString="{0:n0}">
                                        <MaskSettings ErrorText="Please input valid value !"
                                        Mask="<0..999999g>" IncludeLiterals="DecimalSymbol" />
                                    </PropertiesTextEdit>
                                    <HeaderStyle HorizontalAlign="Center"></HeaderStyle>
                                    <CellStyle Font-Names="Tahoma" Font-Size="8pt">
                                    </CellStyle>
                                </dx:GridViewDataTextColumn>
                                <dx:GridViewDataTextColumn VisibleIndex="64" Caption="24" Width="60px" FieldName="DeliveryD24" HeaderStyle-HorizontalAlign="Center">
                                    <PropertiesTextEdit DisplayFormatString="{0:n0}">
                                        <MaskSettings ErrorText="Please input valid value !"
                                        Mask="<0..999999g>" IncludeLiterals="DecimalSymbol" />
                                    </PropertiesTextEdit>
                                    <HeaderStyle HorizontalAlign="Center"></HeaderStyle>
                                    <CellStyle Font-Names="Tahoma" Font-Size="8pt">
                                    </CellStyle>
                                </dx:GridViewDataTextColumn>
                                <dx:GridViewDataTextColumn VisibleIndex="65" Caption="24" Width="0px" FieldName="DeliveryD24Old" HeaderStyle-HorizontalAlign="Center">
                                    <PropertiesTextEdit DisplayFormatString="{0:n0}">
                                        <MaskSettings ErrorText="Please input valid value !"
                                        Mask="<0..999999g>" IncludeLiterals="DecimalSymbol" />
                                    </PropertiesTextEdit>
                                    <HeaderStyle HorizontalAlign="Center"></HeaderStyle>
                                    <CellStyle Font-Names="Tahoma" Font-Size="8pt">
                                    </CellStyle>
                                </dx:GridViewDataTextColumn>
                                <dx:GridViewDataTextColumn VisibleIndex="66" Caption="25" Width="60px" FieldName="DeliveryD25" HeaderStyle-HorizontalAlign="Center">
                                    <PropertiesTextEdit DisplayFormatString="{0:n0}">
                                        <MaskSettings ErrorText="Please input valid value !"
                                        Mask="<0..999999g>" IncludeLiterals="DecimalSymbol" />
                                    </PropertiesTextEdit>
                                    <HeaderStyle HorizontalAlign="Center"></HeaderStyle>
                                    <CellStyle Font-Names="Tahoma" Font-Size="8pt">
                                    </CellStyle>
                                </dx:GridViewDataTextColumn>
                                <dx:GridViewDataTextColumn VisibleIndex="67" Caption="25" Width="0px" FieldName="DeliveryD25Old" HeaderStyle-HorizontalAlign="Center">
                                    <PropertiesTextEdit DisplayFormatString="{0:n0}">
                                        <MaskSettings ErrorText="Please input valid value !"
                                        Mask="<0..999999g>" IncludeLiterals="DecimalSymbol" />
                                    </PropertiesTextEdit>
                                    <HeaderStyle HorizontalAlign="Center"></HeaderStyle>
                                    <CellStyle Font-Names="Tahoma" Font-Size="8pt">
                                    </CellStyle>
                                </dx:GridViewDataTextColumn>
                                <dx:GridViewDataTextColumn VisibleIndex="68" Caption="26" Width="60px" FieldName="DeliveryD26" HeaderStyle-HorizontalAlign="Center">
                                    <PropertiesTextEdit DisplayFormatString="{0:n0}">
                                        <MaskSettings ErrorText="Please input valid value !"
                                        Mask="<0..999999g>" IncludeLiterals="DecimalSymbol" />
                                    </PropertiesTextEdit>
                                    <HeaderStyle HorizontalAlign="Center"></HeaderStyle>
                                    <CellStyle Font-Names="Tahoma" Font-Size="8pt">
                                    </CellStyle>
                                </dx:GridViewDataTextColumn>
                                <dx:GridViewDataTextColumn VisibleIndex="69" Caption="26" Width="0px" FieldName="DeliveryD26Old" HeaderStyle-HorizontalAlign="Center">
                                    <PropertiesTextEdit DisplayFormatString="{0:n0}">
                                        <MaskSettings ErrorText="Please input valid value !"
                                        Mask="<0..999999g>" IncludeLiterals="DecimalSymbol" />
                                    </PropertiesTextEdit>
                                    <HeaderStyle HorizontalAlign="Center"></HeaderStyle>
                                    <CellStyle Font-Names="Tahoma" Font-Size="8pt">
                                    </CellStyle>
                                </dx:GridViewDataTextColumn>
                                <dx:GridViewDataTextColumn VisibleIndex="70" Caption="27" Width="60px" FieldName="DeliveryD27" HeaderStyle-HorizontalAlign="Center">
                                    <PropertiesTextEdit DisplayFormatString="{0:n0}">
                                        <MaskSettings ErrorText="Please input valid value !"
                                        Mask="<0..999999g>" IncludeLiterals="DecimalSymbol" />
                                    </PropertiesTextEdit>
                                    <HeaderStyle HorizontalAlign="Center"></HeaderStyle>
                                    <CellStyle Font-Names="Tahoma" Font-Size="8pt">
                                    </CellStyle>
                                </dx:GridViewDataTextColumn>
                                <dx:GridViewDataTextColumn VisibleIndex="71" Caption="27" Width="0px" FieldName="DeliveryD27Old" HeaderStyle-HorizontalAlign="Center">
                                    <PropertiesTextEdit DisplayFormatString="{0:n0}">
                                        <MaskSettings ErrorText="Please input valid value !"
                                        Mask="<0..999999g>" IncludeLiterals="DecimalSymbol" />
                                    </PropertiesTextEdit>
                                    <HeaderStyle HorizontalAlign="Center"></HeaderStyle>
                                    <CellStyle Font-Names="Tahoma" Font-Size="8pt">
                                    </CellStyle>
                                </dx:GridViewDataTextColumn>
                                <dx:GridViewDataTextColumn VisibleIndex="72" Caption="28" Width="60px" FieldName="DeliveryD28" HeaderStyle-HorizontalAlign="Center">
                                    <PropertiesTextEdit DisplayFormatString="{0:n0}">
                                        <MaskSettings ErrorText="Please input valid value !"
                                        Mask="<0..999999g>" IncludeLiterals="DecimalSymbol" />
                                    </PropertiesTextEdit>
                                    <HeaderStyle HorizontalAlign="Center"></HeaderStyle>
                                    <CellStyle Font-Names="Tahoma" Font-Size="8pt">
                                    </CellStyle>
                                </dx:GridViewDataTextColumn>
                                <dx:GridViewDataTextColumn VisibleIndex="73" Caption="28" Width="0px" FieldName="DeliveryD28Old" HeaderStyle-HorizontalAlign="Center">
                                    <PropertiesTextEdit DisplayFormatString="{0:n0}">
                                        <MaskSettings ErrorText="Please input valid value !"
                                        Mask="<0..999999g>" IncludeLiterals="DecimalSymbol" />
                                    </PropertiesTextEdit>
                                    <HeaderStyle HorizontalAlign="Center"></HeaderStyle>
                                    <CellStyle Font-Names="Tahoma" Font-Size="8pt">
                                    </CellStyle>
                                </dx:GridViewDataTextColumn>
                                <dx:GridViewDataTextColumn VisibleIndex="74" Caption="29" Width="60px" FieldName="DeliveryD29" HeaderStyle-HorizontalAlign="Center">
                                    <PropertiesTextEdit DisplayFormatString="{0:n0}">
                                        <MaskSettings ErrorText="Please input valid value !"
                                        Mask="<0..999999g>" IncludeLiterals="DecimalSymbol" />
                                    </PropertiesTextEdit>
                                    <HeaderStyle HorizontalAlign="Center"></HeaderStyle>
                                    <CellStyle Font-Names="Tahoma" Font-Size="8pt">
                                    </CellStyle>
                                </dx:GridViewDataTextColumn>
                                <dx:GridViewDataTextColumn VisibleIndex="75" Caption="29" Width="0px" FieldName="DeliveryD29Old" HeaderStyle-HorizontalAlign="Center">
                                    <PropertiesTextEdit DisplayFormatString="{0:n0}">
                                        <MaskSettings ErrorText="Please input valid value !"
                                        Mask="<0..999999g>" IncludeLiterals="DecimalSymbol" />
                                    </PropertiesTextEdit>
                                    <HeaderStyle HorizontalAlign="Center"></HeaderStyle>
                                    <CellStyle Font-Names="Tahoma" Font-Size="8pt">
                                    </CellStyle>
                                </dx:GridViewDataTextColumn>
                                <dx:GridViewDataTextColumn VisibleIndex="76" Caption="30" Width="60px" FieldName="DeliveryD30" HeaderStyle-HorizontalAlign="Center">
                                    <PropertiesTextEdit DisplayFormatString="{0:n0}">
                                        <MaskSettings ErrorText="Please input valid value !"
                                        Mask="<0..999999g>" IncludeLiterals="DecimalSymbol" />
                                    </PropertiesTextEdit>
                                    <HeaderStyle HorizontalAlign="Center"></HeaderStyle>
                                    <CellStyle Font-Names="Tahoma" Font-Size="8pt">
                                    </CellStyle>
                                </dx:GridViewDataTextColumn>
                                <dx:GridViewDataTextColumn VisibleIndex="77" Caption="30" Width="0px" FieldName="DeliveryD30Old" HeaderStyle-HorizontalAlign="Center">
                                    <PropertiesTextEdit DisplayFormatString="{0:n0}">
                                        <MaskSettings ErrorText="Please input valid value !"
                                        Mask="<0..999999g>" IncludeLiterals="DecimalSymbol" />
                                    </PropertiesTextEdit>
                                    <HeaderStyle HorizontalAlign="Center"></HeaderStyle>
                                    <CellStyle Font-Names="Tahoma" Font-Size="8pt">
                                    </CellStyle>
                                </dx:GridViewDataTextColumn>
                                <dx:GridViewDataTextColumn VisibleIndex="78" Caption="31" Width="60px" FieldName="DeliveryD31" HeaderStyle-HorizontalAlign="Center">
                                    <PropertiesTextEdit DisplayFormatString="{0:n0}">
                                        <MaskSettings ErrorText="Please input valid value !"
                                        Mask="<0..999999g>" IncludeLiterals="DecimalSymbol" />
                                    </PropertiesTextEdit>
                                    <HeaderStyle HorizontalAlign="Center"></HeaderStyle>
                                    <CellStyle Font-Names="Tahoma" Font-Size="8pt">
                                    </CellStyle>
                                </dx:GridViewDataTextColumn>
                                <dx:GridViewDataTextColumn VisibleIndex="79" Caption="31" Width="0px" FieldName="DeliveryD31Old" HeaderStyle-HorizontalAlign="Center">
                                    <PropertiesTextEdit DisplayFormatString="{0:n0}">
                                        <MaskSettings ErrorText="Please input valid value !"
                                        Mask="<0..999999g>" IncludeLiterals="DecimalSymbol" />
                                    </PropertiesTextEdit>
                                    <HeaderStyle HorizontalAlign="Center"></HeaderStyle>
                                    <CellStyle Font-Names="Tahoma" Font-Size="8pt">
                                    </CellStyle>
                                </dx:GridViewDataTextColumn>
                                <dx:GridViewDataTextColumn VisibleIndex="80" Caption="Header" Width="0px" FieldName="Header" HeaderStyle-HorizontalAlign="Center">                                   
                                    <HeaderStyle HorizontalAlign="Center"></HeaderStyle>
                                    <CellStyle Font-Names="Tahoma" Font-Size="8pt">
                                    </CellStyle>
                                </dx:GridViewDataTextColumn>
                            </Columns>
                            <HeaderStyle HorizontalAlign="Center" />
                        </dx:GridViewBandColumn>
                    </Columns>
                    <SettingsBehavior AllowSelectByRowClick="True" AllowSort="False"
                        ColumnResizeMode="Control" EnableRowHotTrack="True" />
                    <SettingsPager Visible="False" PageSize="15" Position="Top" 
                        Mode="ShowAllRecords">
                        <Summary Text="Page {0} of {1} [{2} record(s)]"
                                  AllPagesText="Page {0} of {1} " />
                        <Summary AllPagesText="Page {0} of {1} " Text="Page {0} of {1} [{2} record(s)]"></Summary>
                    </SettingsPager>
                    <SettingsEditing Mode="Batch" NewItemRowPosition="Bottom">
                        <BatchEditSettings ShowConfirmOnLosingChanges="False" />
                    </SettingsEditing>
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

        <tr>
            <td valign="top" align="left" colspan="6">
                <dx:ASPxButton ID="btnSubMenu" runat="server" Text="SUB MENU"
                    Font-Names="Tahoma" Width="85px" Font-Size="8pt">
                </dx:ASPxButton>
            </td>
            <td valign="top" align="right" style="width: 50px;">
            </td>       
            <td valign="top" align="right" style="width: 50px;">
                <dx:ASPxButton ID="btnDownload" runat="server" Text="DOWNLOAD"
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

    <dx:ASPxCallback ID="ButtonApprove" runat="server" ClientInstanceName="ButtonApprove">
        <ClientSideEvents EndCallback="function(s, e) {
	        txtUser1.SetText(s.cpUser1);
            txtUser2.SetText(s.cpUser2);
            txtUser3.SetText(s.cpUser3);
            txtUser4.SetText(s.cpUser4);
            txtUser5.SetText(s.cpUser5);
            txtUser6.SetText(s.cpUser6);
            txtUser7.SetText(s.cpUser7);
            txtUser8.SetText(s.cpUser8);

            txtDate1.SetText(s.cpDate1);
            txtDate2.SetText(s.cpDate2);
            txtDate3.SetText(s.cpDate3);
            txtDate4.SetText(s.cpDate4);
            txtDate5.SetText(s.cpDate5);
            txtDate6.SetText(s.cpDate6);
            txtDate7.SetText(s.cpDate7);
            txtDate8.SetText(s.cpDate8);

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
    <dx:ASPxCallback ID="ButtonPartNo" runat="server" ClientInstanceName="ButtonPartNo">
        <ClientSideEvents EndCallback="function(s, e) {
	         if (s.cpDelivery != '') 
            {
                txtDelivery.SetText(s.cpDelivery);
                txtCommercial.SetText(s.cpCommercial);
                txtShip.SetText(s.cpShip);
                txtRemarks.SetText(s.cpRemarks);
            } else {
                lblInfo.SetText('');
            }
            delete s.cpDelivery;
        }" />
    </dx:ASPxCallback>
</asp:Content>

