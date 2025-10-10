<%@ Page Title="" Language="vb" AutoEventWireup="false" MasterPageFile="~/SiteContent.Master" CodeBehind="POList.aspx.vb" Inherits="AffiliateSystem.POList" %>
<%@ Register Assembly="DevExpress.Web.v14.1, Version=14.1.6.0, Culture=neutral, PublicKeyToken=b88d1754d700e49a" namespace="DevExpress.Web.ASPxGridView.Export" TagPrefix="dx" %>
<%@ Register Assembly="DevExpress.Web.v14.1, Version=14.1.6.0, Culture=neutral, PublicKeyToken=b88d1754d700e49a" namespace="DevExpress.Web.ASPxGridView" TagPrefix="dx" %>
<%@ Register assembly="DevExpress.Web.v14.1, Version=14.1.6.0, Culture=neutral, PublicKeyToken=b88d1754d700e49a" namespace="DevExpress.Web.ASPxEditors" tagprefix="dx" %>
<%@ Register Assembly="DevExpress.Web.v14.1, Version=14.1.6.0, Culture=neutral, PublicKeyToken=b88d1754d700e49a" namespace="DevExpress.Web.ASPxGlobalEvents" TagPrefix="dx" %>
<%@ Register Assembly="DevExpress.Web.v14.1, Version=14.1.6.0, Culture=neutral, PublicKeyToken=b88d1754d700e49a" namespace="DevExpress.Web.ASPxPanel" TagPrefix="dx" %>
<%@ Register Assembly="DevExpress.Web.v14.1, Version=14.1.6.0, Culture=neutral, PublicKeyToken=b88d1754d700e49a" namespace="DevExpress.Web.ASPxRoundPanel" TagPrefix="dx" %>

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
        height = height - (height * 52 / 100)
        grid.SetHeight(height);
    }

//    function OnBatchEditStartEditing(s, e) {
//        currentColumnName = e.focusedColumn.fieldName;

//        if (currentColumnName == "DetailPage" || currentColumnName == "RevisePage" || currentColumnName == "Period" || currentColumnName == "PONo" || currentColumnName == "CommercialCls"
//            || currentColumnName == "ShipCls" || currentColumnName == "CurrAff" || currentColumnName  == "POMarking"
//            || currentColumnName == "AmountAff" || currentColumnName == "EntryDate" || currentColumnName == "EntryUser" || currentColumnName == "POStatus1"
//            || currentColumnName == "POStatus2" || currentColumnName == "POStatus3" || currentColumnName == "POStatus4" || currentColumnName == "POStatus5" || currentColumnName == "POStatus6" || currentColumnName == "POStatus7" || currentColumnName == "POStatus8") {
//            e.cancel = true;
//        }

//        currentEditableVisibleIndex = e.visibleIndex;
//    }
</script>
</asp:Content>

<asp:Content ID="Content2" ContentPlaceHolderID="Content" runat="server">
    <table style="width:100%;">
        <tr>
            <td width="50%">
                <table style="border-width: 1pt thin thin thin; border-style: ridge; border-color:#9598A1; width:100%; height: 120px;">
                    <tr>
                        <td colspan="8" height="30">
                            <table id="Table1">
                                <tr>
                                    <td style="width:5px;"></td>
                                    <td align="left" valign="middle" height="25px" width="130px">
                                        <dx:ASPxLabel ID="ASPxLabel1" runat="server" Text="PERIOD"
                                            Font-Names="Tahoma" Font-Size="8pt">
                                        </dx:ASPxLabel>
                                    </td>
                                    <td align="left" valign="middle" height="25px" width="203px">
                                        <table>
                                            <tr>
                                                <td>
                                                    <dx:ASPxTimeEdit ID="dtPeriodFrom" runat="server" ClientInstanceName="dtPeriodFrom"
                                                        DisplayFormatString="MMM yyyy" EditFormat="Custom"
                                                        EditFormatString="MMM yyyy" Font-Names="Tahoma" Font-Size="8pt"
                                                        Width="100px">
                                                        <ClientSideEvents ValueChanged="function(s, e) {
	                                                        grid.SetFocusedRowIndex(-1);
                                                            grid.PerformCallback('kosong');
	                                                        lblInfo.SetText('');
                                                        }" />
                                                    </dx:ASPxTimeEdit>
                                                </td>
                                                <td align="left" valign="middle" height="25px" width="3px"> ~
                                                </td>
                                                <td>
                                                    <dx:ASPxTimeEdit ID="dtPeriodTo" runat="server" ClientInstanceName="dtPeriodTo"
                                                        DisplayFormatString="MMM yyyy" EditFormat="Custom"
                                                        EditFormatString="MMM yyyy" Font-Names="Tahoma" Font-Size="8pt"
                                                        Width="100px">
                                                        <ClientSideEvents ValueChanged="function(s, e) {
	                                                        grid.SetFocusedRowIndex(-1);
                                                            grid.PerformCallback('kosong');
	                                                        lblInfo.SetText('');
                                                        }" />
                                                    </dx:ASPxTimeEdit>
                                                </td>
                                            </tr>
                                        </table>                                        
                                    </td>
                                    <td style="width:25px;"></td>
                                    <td align="left" valign="middle" height="25px" width="80px">
                                        <dx:ASPxLabel ID="ASPxLabel5" runat="server" Text="PO NO."
                                            Font-Names="Tahoma" Font-Size="8pt">
                                        </dx:ASPxLabel>
                                    </td>
                                    <td align="left" valign="middle" style="height:25px; width:170px;">
                                        <dx:ASPxTextBox ID="txtPONo" runat="server" Width="170px"
                                            ClientInstanceName="txtPONo" Font-Names="Tahoma"
                                            MaxLength="20" onkeypress="return singlequote(event)" Height="25px" Font-Size="8pt">
                                            <ClientSideEvents LostFocus="function(s, e) { lblInfo.SetText(&quot;&quot;); }" />
                                        </dx:ASPxTextBox>
                                    </td>
                                    <td style="width:5px;"></td>
                                </tr>
                                <tr>
                                    <td style="width:5px;"></td>
                                    <td align="left" valign="middle" height="25px" width="130px">
                                        <dx:ASPxLabel ID="ASPxLabel2" runat="server" Text="AFFILIATE APPROVAL"
                                            Font-Names="Tahoma" Font-Size="8pt">
                                        </dx:ASPxLabel>
                                    </td>
                                    <td align="left" valign="middle" height="25px" width="203px">
                                        <table>
                                            <tr>
                                                <td>
                                                    <dx:ASPxRadioButton ID="rdrAff1" ClientInstanceName="rdrAff1" runat="server" Text="ALL" GroupName="Affiliate" Font-Names="Tahoma" Font-Size="8pt">
                                                        <ClientSideEvents LostFocus="function(s, e) {
                                                        grid.SetFocusedRowIndex(-1);
                                                        grid.PerformCallback('kosong');
                                                        lblInfo.SetText('');
                                                        }" />
                                                    </dx:ASPxRadioButton>
                                                </td>
                                                <td>
                                                    <dx:ASPxRadioButton ID="rdrAff2" ClientInstanceName="rdrAff2" runat="server" Text="YES" GroupName="Affiliate" Font-Names="Tahoma" Font-Size="8pt">
                                                        <ClientSideEvents LostFocus="function(s, e) {
                                                        grid.SetFocusedRowIndex(-1);
                                                        grid.PerformCallback('kosong');
                                                        lblInfo.SetText('');
                                                        }" />
                                                    </dx:ASPxRadioButton>
                                                </td>
                                                <td>
                                                    <dx:ASPxRadioButton ID="rdrAff3" ClientInstanceName="rdrAff3" runat="server" Text="NO" GroupName="Affiliate" Font-Names="Tahoma" Font-Size="8pt">
                                                        <ClientSideEvents LostFocus="function(s, e) {
                                                        grid.SetFocusedRowIndex(-1);
                                                        grid.PerformCallback('kosong');
                                                        lblInfo.SetText('');
                                                        }" />
                                                    </dx:ASPxRadioButton>
                                                </td>
                                            </tr>
                                        </table> 
                                    </td>
                                    <td style="width:25px;"></td>
                                    <td align="left" valign="middle" height="25px" width="80px"></td>
                                    <td align="left" valign="middle" style="height:25px; width:170px;"></td>
                                    <td style="width:5px;"></td>
                                </tr>
                                <tr>
                                    <td style="width:5px;"></td>
                                    <td align="left" valign="middle" height="25px" width="130px">
                                        <dx:ASPxLabel ID="ASPxLabel4" runat="server" Text="COMMERCIAL"
                                            Font-Names="Tahoma" Font-Size="8pt">
                                        </dx:ASPxLabel>
                                    </td>
                                    <td align="left" valign="middle" height="25px" width="203px">
                                        <table>
                                            <tr>
                                                <td>
                                                    <dx:ASPxRadioButton ID="rdrCom1" ClientInstanceName="rdrCom1" runat="server" Text="ALL" GroupName="Commercial"  Font-Names="Tahoma" Font-Size="8pt">
                                                        <ClientSideEvents LostFocus="function(s, e) {
                                                        grid.SetFocusedRowIndex(-1);
                                                        grid.PerformCallback('kosong');
                                                        lblInfo.SetText('');
                                                        }" />
                                                    </dx:ASPxRadioButton>
                                                </td>
                                                <td>
                                                    <dx:ASPxRadioButton ID="rdrCom2" ClientInstanceName="rdrCom2" runat="server" Text="YES" GroupName="Commercial"  Font-Names="Tahoma" Font-Size="8pt">
                                                        <ClientSideEvents LostFocus="function(s, e) {
                                                        grid.SetFocusedRowIndex(-1);
                                                        grid.PerformCallback('kosong');
                                                        lblInfo.SetText('');
                                                        }" />
                                                    </dx:ASPxRadioButton>
                                                </td>
                                                <td>
                                                    <dx:ASPxRadioButton ID="rdrCom3" ClientInstanceName="rdrCom3" runat="server" Text="NO" GroupName="Commercial"  Font-Names="Tahoma" Font-Size="8pt">
                                                        <ClientSideEvents LostFocus="function(s, e) {
                                                        grid.SetFocusedRowIndex(-1);
                                                        grid.PerformCallback('kosong');
                                                        lblInfo.SetText('');
                                                        }" />
                                                    </dx:ASPxRadioButton>
                                                </td>
                                            </tr>
                                        </table>
                                    </td>
                                    <td style="width:25px;"></td>
                                    <td align="left" valign="middle" height="25px" width="80px"></td>
                                    <td align="left" valign="middle" style="height:25px; width:170px;">
                                        <table style="width:100%;">
                                            <tr>
                                                <td align="right" valign="middle" style="height:25px; width:90px;">
                                                    <dx:ASPxButton ID="btnRefresh" runat="server" Text="SEARCH"
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
                                                        <ClientSideEvents Click="function(s, e) {
                                                            txtPartCode.SetText('');
                                                            txtPartName.SetText('');
                                                            lblInfo.SetText('');
                                                            grid.SetFocusedRowIndex(-1);
                                                            grid.PerformCallback('kosong');
                                                        }" />                                   
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
            <td valign="top" width="40%" align="left">
                <dx:ASPxRoundPanel ID="ASPxRoundPanel1" runat="server" HeaderText="PO STATUS" ShowCollapseButton="true"
                    View="GroupBox" Width="100%" Font-Size="8pt" Font-Names="Tahoma">
                    <PanelCollection>
                        <dx:PanelContent runat="server">
                            <table id="Table2">
                                <tr>
                                    <td align="left" height="20px" valign="middle" width="50%">
                                        <dx:ASPxLabel ID="ASPxLabel8" runat="server" Font-Names="Tahoma" Font-Size="8pt"
                                            ForeColor="#003366" Text="(1) AFFILIATE ENTRY" Width="100%">
                                        </dx:ASPxLabel>
                                    </td>
                                    <td align="left" height="20px" valign="middle" width="50%">
                                        <dx:ASPxLabel ID="ASPxLabel14" runat="server" Font-Names="Tahoma" Font-Size="8pt"
                                            ForeColor="#003366" Text="(5) SUPPLIER PENDING (PARTIAL)" Width="100%">
                                        </dx:ASPxLabel>
                                    </td>
                                </tr>
                                <tr>
                                    <td align="left" height="20px" valign="middle" width="50%">
                                        <dx:ASPxLabel ID="ASPxLabel11" runat="server" Font-Names="Tahoma" Font-Size="8pt"
                                            ForeColor="#003366" Text="(2) AFFILIATE APPROVAL" Width="100%">
                                        </dx:ASPxLabel>
                                    </td>
                                    <td align="left" height="20px" valign="middle" width="50%">
                                        <dx:ASPxLabel ID="ASPxLabel15" runat="server" Font-Names="Tahoma" Font-Size="8pt"
                                            ForeColor="#003366" Text="(6) SUPPLIER UNAPPROVE" Width="100%">
                                        </dx:ASPxLabel>
                                    </td>
                                </tr>
                                <tr>
                                    <td align="left" height="20px" valign="middle" width="50%">
                                        <dx:ASPxLabel ID="ASPxLabel12" runat="server" Font-Names="Tahoma" Font-Size="8pt"
                                            ForeColor="#003366" Text="(3) PASI SEND AFFILIATE PO TO SUPPLIER" 
                                            Width="100%">
                                        </dx:ASPxLabel>
                                    </td>
                                    <td align="left" height="20px" valign="middle" width="50%">
                                        <dx:ASPxLabel ID="ASPxLabel16" runat="server" Font-Names="Tahoma" Font-Size="8pt"
                                            ForeColor="#003366" Text="(7) PASI APPROVAL" Width="100%">
                                        </dx:ASPxLabel>
                                    </td>
                                </tr>
                                <tr>
                                    <td align="left" height="20px" valign="middle" width="50%">
                                        <dx:ASPxLabel ID="ASPxLabel13" runat="server" Font-Names="Tahoma" Font-Size="8pt"
                                            ForeColor="#003366" Text="(4) SUPPLIER APPROVAL" Width="100%">
                                        </dx:ASPxLabel>
                                    </td>
                                    <td align="left" height="20px" valign="middle" width="50%">
                                        <dx:ASPxLabel ID="ASPxLabel17" runat="server" Font-Names="Tahoma" Font-Size="8pt"
                                            ForeColor="#003366" Text="(8) AFFILIATE FINAL APPROVAL" Width="100%">
                                        </dx:ASPxLabel>
                                    </td>
                                </tr>
                            </table>
                        </dx:PanelContent>
                    </PanelCollection>
                </dx:ASPxRoundPanel>
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
            <td valign="top" align="right" style="width: 50px;">                                  
            </td>
            <td valign="top" align="right" style="width: 50px;">    
                <dx:ASPxButton ID="btnADD" runat="server" Text="CREATE PO"
                    Font-Names="Tahoma" Width="85px" AutoPostBack="False" Font-Size="8pt">                               
                </dx:ASPxButton>
            </td>
        </tr>
    </table>

    <div style="height:1px;"></div>

    <table style="width:100%;">
        <tr>
            <td colspan="8" align="left" valign="top" height="220">
                <%--Column : Grid--%>
                <dx:ASPxGridView ID="grid" runat="server" Width="100%"
                    Font-Names="Tahoma" KeyFieldName="AllowAccess;Period;PONo;POMarking;SupplierID"
                    AutoGenerateColumns="False"
                    ClientInstanceName="grid" Font-Size="8pt">
                    <ClientSideEvents Init="OnInit" 
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
                    }" RowClick="function(s, e) {
	                    lblInfo.SetText('');
                    }"/>
                    
                    <Columns>
                        <dx:GridViewDataTextColumn VisibleIndex="0" Caption=" " Width="50px" FieldName="DetailPage">
                            <DataItemTemplate>
                                <a id="clickElement" href="POEntry.aspx?id=<%#GetRowValue(Container)%>&t1=<%#GetAffiliateID(Container)%>&t2=<%#GetAffiliateName(Container)%>&t3=<%#GetPeriod(Container)%>&t4=<%#GetSupplierID(Container)%>&Session=~/PurchaseOrder/POList.aspx">
                                    <%# "DETAIL"%></a>
                            </DataItemTemplate>
                            <CellStyle Font-Names="Tahoma" Font-Size="8pt">
                            </CellStyle>
                        </dx:GridViewDataTextColumn>
                        <dx:GridViewDataTextColumn VisibleIndex="2" Caption="PERIOD" FieldName="Period" Width="70px"
                            HeaderStyle-HorizontalAlign="Center">
                            <PropertiesTextEdit DisplayFormatString="MMM yyyy">
                            </PropertiesTextEdit>                            

<HeaderStyle HorizontalAlign="Center"></HeaderStyle>

                            <CellStyle Font-Names="Tahoma" Font-Size="8pt">
                            </CellStyle>
                        </dx:GridViewDataTextColumn>
                        <dx:GridViewDataTextColumn VisibleIndex="3" Caption="PO NO." FieldName="PONo" Width="220px"
                            HeaderStyle-HorizontalAlign="Center">                            
<HeaderStyle HorizontalAlign="Center"></HeaderStyle>

                            <CellStyle Font-Names="Tahoma" Font-Size="8pt">
                            </CellStyle>
                        </dx:GridViewDataTextColumn>
                        <dx:GridViewDataTextColumn VisibleIndex="4" Caption="PO MARKING" FieldName="POMarking" Width="220px"
                            HeaderStyle-HorizontalAlign="Center">                            
<HeaderStyle HorizontalAlign="Center"></HeaderStyle>

                            <CellStyle Font-Names="Tahoma" Font-Size="8pt">
                            </CellStyle>
                        </dx:GridViewDataTextColumn>
                        <dx:GridViewDataTextColumn VisibleIndex="5" Caption="COMMERCIAL" FieldName="CommercialCls"
                            Width="95px" HeaderStyle-HorizontalAlign="Center" CellStyle-HorizontalAlign="Center">                           
<HeaderStyle HorizontalAlign="Center"></HeaderStyle>

                            <CellStyle Font-Names="Tahoma" Font-Size="8pt">
                            </CellStyle>
                        </dx:GridViewDataTextColumn>
                        <dx:GridViewDataTextColumn VisibleIndex="6" Caption="SHIP BY" FieldName="ShipCls"
                            Width="100px" HeaderStyle-HorizontalAlign="Center">                            
<HeaderStyle HorizontalAlign="Center"></HeaderStyle>

                            <CellStyle Font-Names="Tahoma" Font-Size="8pt">
                            </CellStyle>
                        </dx:GridViewDataTextColumn>
                        <dx:GridViewDataTextColumn Caption="CURR" FieldName="CurrAff" VisibleIndex="7" Width="50px" CellStyle-HorizontalAlign="Center" Visible="False">                            
                            <CellStyle Font-Names="Tahoma" Font-Size="8pt">
                            </CellStyle>
                        </dx:GridViewDataTextColumn>
                        <dx:GridViewDataTextColumn Caption="AMOUNT" FieldName="AmountAff" VisibleIndex="8" Width="150px" Visible="False">
                            <HeaderStyle HorizontalAlign="Center"></HeaderStyle>
                            <CellStyle Font-Names="Tahoma" Font-Size="8pt">
                            </CellStyle>
                            <PropertiesTextEdit DisplayFormatString="{0:n2}">
                            </PropertiesTextEdit>
                        </dx:GridViewDataTextColumn>
                        <dx:GridViewDataTextColumn VisibleIndex="9" Caption="CREATED DATE" FieldName="EntryDate"
                            HeaderStyle-HorizontalAlign="Center" Width="130px">
                            <PropertiesTextEdit DisplayFormatString="yyyy-MM-dd HH:mm:ss">
                            </PropertiesTextEdit>

<HeaderStyle HorizontalAlign="Center"></HeaderStyle>

                            <CellStyle Font-Names="Tahoma" Font-Size="8pt">
                            </CellStyle>
                        </dx:GridViewDataTextColumn>
                        <dx:GridViewDataTextColumn VisibleIndex="10" Caption="CREATED BY" FieldName="EntryUser"
                            HeaderStyle-HorizontalAlign="Center" Width="145px">
<HeaderStyle HorizontalAlign="Center"></HeaderStyle>

                            <CellStyle Font-Names="Tahoma" Font-Size="8pt">
                            </CellStyle>
                        </dx:GridViewDataTextColumn>
                        <dx:GridViewBandColumn Caption="PO STATUS" VisibleIndex="11">
                            <Columns>
                                <dx:GridViewDataCheckColumn Caption="1" FieldName="POStatus1" ReadOnly="True" VisibleIndex="1" Width="30px">
                                    <PropertiesCheckEdit ValueChecked="1" ValueType="System.String" ValueUnchecked="0">
                                        <DisplayImageChecked IconID="support_feature_16x16">
                                        </DisplayImageChecked>
                                        <DisplayImageUnchecked Width="0px">
                                        </DisplayImageUnchecked>
                                        <CheckBoxStyle>
                                        <Border BorderStyle="None" />
                                        </CheckBoxStyle>
                                    </PropertiesCheckEdit>
                                    <HeaderStyle HorizontalAlign="Center"></HeaderStyle>
                                </dx:GridViewDataCheckColumn>
                                <dx:GridViewDataCheckColumn Caption="2" FieldName="POStatus2" ReadOnly="True" VisibleIndex="3" Width="30px">
                                    <PropertiesCheckEdit ValueChecked="1" ValueType="System.String" ValueUnchecked="0">
                                        <DisplayImageChecked IconID="support_feature_16x16">
                                        </DisplayImageChecked>
                                        <DisplayImageUnchecked Width="0px">
                                        </DisplayImageUnchecked>
                                    </PropertiesCheckEdit>
                                    <HeaderStyle HorizontalAlign="Center"></HeaderStyle>
                                </dx:GridViewDataCheckColumn>
                                <dx:GridViewDataCheckColumn Caption="3" FieldName="POStatus3" ReadOnly="True" VisibleIndex="5" Width="30px">
                                    <PropertiesCheckEdit ValueChecked="1" ValueType="System.String" ValueUnchecked="0">
                                        <DisplayImageChecked IconID="support_feature_16x16">
                                        </DisplayImageChecked>
                                        <DisplayImageUnchecked Width="0px">
                                        </DisplayImageUnchecked>
                                    </PropertiesCheckEdit>
                                    <HeaderStyle HorizontalAlign="Center"></HeaderStyle>
                                </dx:GridViewDataCheckColumn>
                                <dx:GridViewDataCheckColumn Caption="4" FieldName="POStatus4" ReadOnly="True" VisibleIndex="7" Width="30px">
                                    <PropertiesCheckEdit ValueChecked="1" ValueType="System.String" ValueUnchecked="0">
                                        <DisplayImageChecked IconID="support_feature_16x16">
                                        </DisplayImageChecked>
                                        <DisplayImageUnchecked Width="0px">
                                        </DisplayImageUnchecked>
                                    </PropertiesCheckEdit>
                                </dx:GridViewDataCheckColumn>
                                <dx:GridViewDataCheckColumn Caption="5" FieldName="POStatus5" ReadOnly="True" VisibleIndex="9" Width="30px">
                                    <PropertiesCheckEdit ValueChecked="1" ValueType="System.String" ValueUnchecked="0">
                                        <DisplayImageChecked IconID="support_feature_16x16">
                                        </DisplayImageChecked>
                                        <DisplayImageUnchecked Width="0px">
                                        </DisplayImageUnchecked>
                                    </PropertiesCheckEdit>
                                    <HeaderStyle HorizontalAlign="Center"></HeaderStyle>
                                </dx:GridViewDataCheckColumn>
                                <dx:GridViewDataCheckColumn Caption="6" FieldName="POStatus6" ReadOnly="True" VisibleIndex="11" Width="30px">
                                    <PropertiesCheckEdit ValueChecked="1" ValueType="System.String" ValueUnchecked="0">
                                        <DisplayImageChecked IconID="support_feature_16x16">
                                        </DisplayImageChecked>
                                        <DisplayImageUnchecked Width="0px">
                                        </DisplayImageUnchecked>
                                    </PropertiesCheckEdit>
                                    <HeaderStyle HorizontalAlign="Center"></HeaderStyle>
                                </dx:GridViewDataCheckColumn>
                                <dx:GridViewDataCheckColumn Caption="7" FieldName="POStatus7" ReadOnly="True" VisibleIndex="13" Width="30px">
                                    <PropertiesCheckEdit ValueChecked="1" ValueType="System.String" ValueUnchecked="0">
                                        <DisplayImageChecked IconID="support_feature_16x16">
                                        </DisplayImageChecked>
                                        <DisplayImageUnchecked Width="0px">
                                        </DisplayImageUnchecked>
                                    </PropertiesCheckEdit>
                                    <HeaderStyle HorizontalAlign="Center"></HeaderStyle>
                                </dx:GridViewDataCheckColumn>
                                <dx:GridViewDataCheckColumn Caption="8" FieldName="POStatus8" ReadOnly="True" VisibleIndex="14" Width="30px">
                                    <PropertiesCheckEdit ValueChecked="1" ValueType="System.String" ValueUnchecked="0">
                                        <DisplayImageChecked IconID="support_feature_16x16">
                                        </DisplayImageChecked>
                                        <DisplayImageUnchecked Width="0px">
                                        </DisplayImageUnchecked>
                                    </PropertiesCheckEdit>
                                    <HeaderStyle HorizontalAlign="Center"></HeaderStyle>
                                </dx:GridViewDataCheckColumn>
                                <dx:GridViewDataTextColumn VisibleIndex="15" Caption="SupplierID" FieldName="SupplierID" Width="0px" HeaderStyle-HorizontalAlign="Center">                                    
<HeaderStyle HorizontalAlign="Center"></HeaderStyle>
                                </dx:GridViewDataTextColumn>
                            </Columns>
                            <HeaderStyle HorizontalAlign="Center"></HeaderStyle>
                        </dx:GridViewBandColumn>
                    </Columns>
                    <SettingsBehavior AllowSelectByRowClick="True" AllowSort="False" 
                        EnableRowHotTrack="True"/>
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
   
        <tr>
            <td valign="top" align="left">
                <dx:ASPxButton ID="btnSubMenu" runat="server" Text="SUB MENU"
                    Font-Names="Tahoma" Width="85px" Font-Size="8pt" AutoPostBack="False">
                </dx:ASPxButton>
            </td>
            <td valign="top" align="right">
                <dx:ASPxButton ID="btnEDI" runat="server" Text="GET E.D.I"
                    Font-Names="Tahoma" Width="85px" Font-Size="8pt" Visible="False">
                </dx:ASPxButton>
            </td>
            <td valign="top" align="right">
                <dx:ASPxButton ID="btnDownload" runat="server" Text="EXPORT TO EXCEL"
                    Font-Names="Tahoma" Width="85px" Font-Size="8pt" AutoPostBack="False">
                    <ClientSideEvents Click="function(s, e) {     
                        grid.PerformCallback('downloadSummary');
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

