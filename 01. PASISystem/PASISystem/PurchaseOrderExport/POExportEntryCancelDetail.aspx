<%@ Page Title="" Language="vb" AutoEventWireup="false" MasterPageFile="~/SiteContent.Master" CodeBehind="POExportEntryCancelDetail.aspx.vb" Inherits="PASISystem.POExportEntryCancelDetail" %>
<%@ Register Assembly="DevExpress.Web.v14.1, Version=14.1.6.0, Culture=neutral, PublicKeyToken=b88d1754d700e49a" namespace="DevExpress.Web.ASPxGridView.Export" TagPrefix="dx" %>
<%@ Register Assembly="DevExpress.Web.v14.1, Version=14.1.6.0, Culture=neutral, PublicKeyToken=b88d1754d700e49a" namespace="DevExpress.Web.ASPxGridView" TagPrefix="dx" %>
<%@ Register assembly="DevExpress.Web.v14.1, Version=14.1.6.0, Culture=neutral, PublicKeyToken=b88d1754d700e49a" namespace="DevExpress.Web.ASPxEditors" tagprefix="dx" %>
<%@ Register Assembly="DevExpress.Web.v14.1, Version=14.1.6.0, Culture=neutral, PublicKeyToken=b88d1754d700e49a" namespace="DevExpress.Web.ASPxGlobalEvents" TagPrefix="dx" %>
<%@ Register Assembly="DevExpress.Web.v14.1, Version=14.1.6.0, Culture=neutral, PublicKeyToken=b88d1754d700e49a" namespace="DevExpress.Web.ASPxPanel" TagPrefix="dx" %>
<%@ Register Assembly="DevExpress.Web.v14.1, Version=14.1.6.0, Culture=neutral, PublicKeyToken=b88d1754d700e49a" namespace="DevExpress.Web.ASPxRoundPanel" TagPrefix="dx" %>
<%@ Register Assembly="DevExpress.Web.v14.1, Version=14.1.6.0, Culture=neutral, PublicKeyToken=b88d1754d700e49a" Namespace="DevExpress.Web.ASPxCallback" TagPrefix="dx" %>
<%@ Register Assembly="DevExpress.Web.v14.1, Version=14.1.6.0, Culture=neutral, PublicKeyToken=b88d1754d700e49a" Namespace="DevExpress.Web.ASPxHiddenField" TagPrefix="dx" %>
<%@ Register assembly="DevExpress.Web.v14.1, Version=14.1.6.0, Culture=neutral, PublicKeyToken=b88d1754d700e49a" namespace="DevExpress.Web.ASPxUploadControl" tagprefix="dx" %>
<%@ Register assembly="DevExpress.Web.v14.1, Version=14.1.6.0, Culture=neutral, PublicKeyToken=b88d1754d700e49a" namespace="DevExpress.Web.ASPxHiddenField" tagprefix="dx1" %>

<asp:Content ID="Content1" ContentPlaceHolderID="head" runat="server">
    <style type="text/css">
        .dxeHLC, .dxeHC, .dxeHFC
        {
            display: none;
        }
        .style1
        {
            width: 170px;   
        }
        .style2
        {
            width: 120px;   
        }
        .style3
        {
            width: 250px;   
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
            height = height - (height * 65 / 100)
            grid.SetHeight(height);
        }

        function OnBatchEditStartEditing(s, e) {
            currentColumnName = e.focusedColumn.fieldName;

            if (currentColumnName == "NoUrut" || currentColumnName == "PartNo" || currentColumnName == "PartName" || currentColumnName == "UnitCls"
                || currentColumnName == "MOQ" || currentColumnName == "PONo" || currentColumnName == "QtyBox" || currentColumnName == "CancelReffQty"
                || currentColumnName == "PreviousForecast" || currentColumnName == "Variance" || currentColumnName == "VariancePercentage"
                || currentColumnName == "TotalPOQty" || currentColumnName == "AffiliateID" || currentColumnName == "SupplierID" || currentColumnName == "Period") {
                e.cancel = true;
           
            }

            currentEditableVisibleIndex = e.visibleIndex;
        }

        function OnBatchEditEndEditing(s, e) {
            currentEditableVisibleIndex = -1;

            window.setTimeout(function () {
                var Week1 = s.batchEditApi.GetCellValue(e.visibleIndex, "Week1");
                var PreviousForecast = s.batchEditApi.GetCellValue(e.visibleIndex, "PreviousForecast");

                s.batchEditApi.SetCellValue(e.visibleIndex, "Variance", parseInt(Week1) - parseInt(PreviousForecast));

                if (PreviousForecast == "0.00") {
                    s.batchEditApi.SetCellValue(e.visibleIndex, "VariancePercentage", parseInt(0));
                }
                else {
                    s.batchEditApi.SetCellValue(e.visibleIndex, "VariancePercentage", (parseInt(Week1) - (PreviousForecast) / (PreviousForecast)) * 100 / 100);

                    if (((parseInt(Week1) - (PreviousForecast) / (PreviousForecast)) * 100 / 100) > 30) {
                        s.GetRow(s.GetFocusedRowIndex()).style.backgroundColor = "magenta";
                    }
                    else {
                        s.GetRow(s.GetFocusedRowIndex()).style.backgroundColor = "lightyellow";
                    }
                }
            }, 10);
        }
    </script>
</asp:Content>

<asp:Content ID="Content2" ContentPlaceHolderID="Content" runat="server">
    <table style="width:100%;">
        <tr>
            <td colspan="16" width="100%">
                <table style="border-width: 1pt thin thin thin; border-style: ridge; border-color:#9598A1; width:100%; height: 70px;">
                    <tr>
                        <td colspan="8" height="20">
                            <table id="Table1" width="100%">
                                <tr>
                                    <td align="left" valign="middle" class="style1">
                                        <dx:ASPxLabel ID="ASPxLabel1" runat="server" Text="PERIOD"
                                            Font-Names="Tahoma" Font-Size="8pt" Width="100%">
                                        </dx:ASPxLabel>
                                    </td>
                                    <td align="left" valign="middle" colspan="2" class="style2">
                                        <dx:ASPxTimeEdit ID="dtPeriodFrom" runat="server"
                                            ClientInstanceName="dtPeriodFrom" DisplayFormatString="yyyy-MM" 
                                            EditFormat="Custom" EditFormatString="yyyy-MM" Width="75px" 
                                            Height="21px" ReadOnly="True">
                                            <ClientSideEvents DateChanged="function(s, e) {
	                                            gridPerformCallback('kosong');
                                            }" />
                                        </dx:ASPxTimeEdit>
                                    </td>
                                    <td width="15">&nbsp;</td>
                                    <td align="left" colspan="3" class="style3">
                                        <table style="width:100%;">
                                            <tr>
                                                <td width="50%">
                                                    <dx:ASPxLabel ID="ASPxLabel4" runat="server" Text="COMMERCIAL"
                                                        Font-Names="Tahoma" Font-Size="8pt" Width="100%">
                                                    </dx:ASPxLabel>
                                                </td>
                                                <td align="left" width="25%">
                                                    <dx:ASPxRadioButton ID="rdrCom1" ClientInstanceName="rdrCom1" runat="server" 
                                                        Text="YES" GroupName="Commercial" Font-Names="Tahoma" Font-Size="8pt" 
                                                        ReadOnly="True">
                                                        <ClientSideEvents LostFocus="function(s, e) {
                                                            lblInfo.SetText('');
                                                        }" />
                                                    </dx:ASPxRadioButton>
                                                </td>
                                                <td align="left" width="25%">
                                                    <dx:ASPxRadioButton ID="rdrCom2" ClientInstanceName="rdrCom2" runat="server" 
                                                        Text="NO" GroupName="Commercial" Font-Names="Tahoma" Font-Size="8pt" 
                                                        ReadOnly="True">
                                                        <ClientSideEvents LostFocus="function(s, e) {
                                                            lblInfo.SetText('');                                                         
                                                        }" />
                                                    </dx:ASPxRadioButton>
                                                </td>
                                            </tr>
                                        </table>
                                    </td>
                                    <td width="15">&nbsp;</td>
                                    <td align="left" colspan="3" class="style3">
                                        <table style="width:100%;">
                                            <tr>
                                                <td align="left" valign="middle" colspan="2" width="50%">
                                                    <dx:ASPxLabel ID="ASPxLabel32" runat="server" Text="DELIVERY LOCATION"
                                                        Font-Names="Tahoma" Font-Size="8pt" width="100%">
                                                    </dx:ASPxLabel>
                                                </td>
                                                <td align="left" valign="middle" colspan="2" width="50%">
                                                    <dx:ASPxComboBox ID="cboDelLoc" width="100%" runat="server" 
                                                        Font-Size="8pt" Font-Names="Tahoma" TextFormatString="{0}" 
                                                        ClientInstanceName="cboDelLoc" TabIndex="3" ReadOnly="True">
                                                        <ClientSideEvents SelectedIndexChanged="function(s, e) {
	                                                        txtDelLoc.SetText(cboDelLoc.GetSelectedItem().GetColumnText(1));
	                                                        lblInfo.SetText('');
                                                        }" />
                                                        <LoadingPanelStyle ImageSpacing="5px">
                                                        </LoadingPanelStyle>
                                                    </dx:ASPxComboBox>
                                                </td>
                                            </tr>
                                        </table>
                                    </td>
                                    <td align="left" valign="middle" colspan="3" class="style3">
                                        <dx:ASPxTextBox ID="txtDelLoc" runat="server" Width="100%" 
                                            ClientInstanceName="txtDelLoc" Font-Names="Tahoma" Font-Size="8pt"
                                            ReadOnly="True" MaxLength="100" Height="20px">
                                            <readonlystyle backcolor="#CCCCCC">
                                            </readonlystyle>
                                        </dx:ASPxTextBox>
                                    </td>
                                    <td>&nbsp;</td>
                                </tr>

                                <tr>
                                    <td align="left" valign="middle" class="style1">
                                        <dx:ASPxLabel ID="ASPxLabel5" runat="server" Text="PO MONTHLY /EMERGENCY"
                                            Font-Names="Tahoma" Font-Size="8pt" width="100%">
                                        </dx:ASPxLabel>
                                    </td>
                                    <td align="left" colspan="3" class="style2">
                                        <table style="width:100%;">
                                            <tr>
                                                <td align="left" valign="middle" height="20px" width="50%">
                                                    <dx:ASPxRadioButton ID="rdMonthly" runat="server" 
                                                        ClientInstanceName="rdMonthly" Text="M" GroupName="ME" ReadOnly="True">
                                                    </dx:ASPxRadioButton>
                                                </td>
                                                <td align="left" valign="middle" height="20px" width="50%">
                                                    <dx:ASPxRadioButton ID="rdEmergency" runat="server" 
                                                        ClientInstanceName="rdEmergency" Text="E" GroupName="ME" ReadOnly="True">
                                                    </dx:ASPxRadioButton>
                                                </td>
                                            </tr>
                                        </table>
                                    </td>
                                    <td align="left" colspan="3" class="style3">
                                        <table style="width:100%;">
                                            <tr>
                                                <td width="50%">
                                                    <dx:ASPxLabel ID="ASPxLabel2" runat="server" Text="SHIP BY"
                                                        Font-Names="Tahoma" Font-Size="8pt" Width="100%">
                                                    </dx:ASPxLabel>
                                                </td>
                                                <td width="25%">
                                                    <dx:ASPxRadioButton ID="rdrShipBy2" ClientInstanceName="rdrShipBy2" runat="server" 
                                                        Text="BOAT" GroupName="POEm" Font-Names="Tahoma" Font-Size="8pt" 
                                                        ReadOnly="True">
                                                        <ClientSideEvents LostFocus="function(s, e) {                                                        
                                                            lblInfo.SetText('');
                                                        }" />
                                                    </dx:ASPxRadioButton>
                                                </td>   
                                                <td width="25%">
                                                    <dx:ASPxRadioButton ID="rdrShipBy3" ClientInstanceName="rdrShipBy3" runat="server" 
                                                        Text="AIR" GroupName="POEm" Font-Names="Tahoma" Font-Size="8pt" 
                                                        ReadOnly="True">
                                                        <ClientSideEvents LostFocus="function(s, e) {
                                                            lblInfo.SetText('');                                                         
                                                        }" />
                                                    </dx:ASPxRadioButton>
                                                </td>
                                            </tr>
                                        </table>
                                    </td>
                                    <td>&nbsp;</td>
                                    <td>&nbsp;</td>
                                </tr>
                                        
                                <tr>
                                    <td align="left" valign="middle" class="style1">
                                        <dx:ASPxLabel ID="ASPxLabel20" runat="server" Text="AFFILIATE CODE/NAME"
                                            Font-Names="Tahoma" Font-Size="8pt">
                                        </dx:ASPxLabel>
                                    </td>
                                    <td align="left" valign="middle" colspan="2" class="style2">
                                        <dx:ASPxComboBox ID="cboAffiliate" width="100%" runat="server" 
                                            Font-Size="8pt" Font-Names="Tahoma" TextFormatString="{0}" 
                                            ClientInstanceName="cboAffiliate" TabIndex="3" ReadOnly="True">
                                            <ClientSideEvents SelectedIndexChanged="function(s, e) {
	                                            txtAffiliate.SetText(cboAffiliate.GetSelectedItem().GetColumnText(1));
                                                txtconsignee.SetText(cboAffiliate.GetSelectedItem().GetColumnText(2));
	                                            lblInfo.SetText('');
                                            }" />
                                            <LoadingPanelStyle ImageSpacing="5px">
                                            </LoadingPanelStyle>
                                        </dx:ASPxComboBox>
                                    </td>
                                    <td>&nbsp;</td>
                                    <td align="left" colspan="3" class="style3">
                                        <dx:ASPxTextBox ID="txtAffiliate" runat="server" Width="100%" 
                                            ClientInstanceName="txtAffiliate" Font-Names="Tahoma" Font-Size="8pt"
                                            ReadOnly="True" MaxLength="100" Height="20px">
                                            <readonlystyle backcolor="#CCCCCC">
                                            </readonlystyle>
                                        </dx:ASPxTextBox>
                                    </td>
                                    <td>&nbsp;</td>
                                    <td align="left" valign="middle" class="style3">
                                        <dx:ASPxTextBox ID="txtconsignee" runat="server" Width="0px" 
                                            ClientInstanceName="txtconsignee" ForeColor="White">
                                            <Border BorderStyle="None" />
                                        </dx:ASPxTextBox>
                                    </td>
                                </tr>
                                
                                <tr>
                                    <td align="left" valign="middle" class="style1">
                                        <dx:ASPxLabel ID="ASPxLabel34" runat="server" Text="ORDER NO"
                                            Font-Names="Tahoma" Font-Size="8pt">
                                        </dx:ASPxLabel>
                                    </td>
                                    <td align="left" valign="middle" colspan="2" class="style2">
                                        <dx:ASPxTextBox ID="txtOrderNo" runat="server" Width="100%" 
                                            ClientInstanceName="txtOrderNo" ReadOnly="True">
                                        </dx:ASPxTextBox>
                                    </td>
                                    <td>&nbsp;</td>
                                    <td align="left" colspan="3" class="style3">
                                        <table style="width:100%;">
                                            <tr>
                                                <td align="left" valign="middle" width="50%">
                                                    <dx:ASPxLabel ID="ASPxLabel35" runat="server" Text="ETD VENDOR"
                                                        Font-Names="Tahoma" Font-Size="8pt" Width="100%">
                                                    </dx:ASPxLabel>
                                                </td>
                                                <td align="right" valign="middle" width="50%">
                                                    <dx:ASPxDateEdit ID="dtETDVendor" runat="server" Width="100%" 
                                                        ClientInstanceName="dtETDVendor" DisplayFormatString="yyyy-MM-dd" 
                                                        EditFormat="Custom" EditFormatString="yyyy-MM-dd" ReadOnly="True">
                                                    </dx:ASPxDateEdit>
                                                </td>
                                            </tr>
                                        </table>
                                    </td>
                                    <td>&nbsp;</td>
                                    <td align="left" colspan="3" class="style3">
                                        <table style="width:100%;">
                                            <tr>
                                                <td align="left" width="50%">
                                                    <dx:ASPxLabel ID="ASPxLabel36" runat="server" Text="ETD PORT"
                                                        Font-Names="Tahoma" Font-Size="8pt" Width="100%">
                                                    </dx:ASPxLabel>
                                                </td>
                                                <td align="right" valign="middle" width="50%">
                                                    <dx:ASPxDateEdit ID="dtETDPort" runat="server" Width="100%" 
                                                        ClientInstanceName="dtETDPort" DisplayFormatString="yyyy-MM-dd" 
                                                        EditFormat="Custom" EditFormatString="yyyy-MM-dd" ReadOnly="True">
                                                    </dx:ASPxDateEdit>
                                                </td>
                                            </tr>
                                        </table>
                                    </td>
                                    <td align="left" colspan="3" class="style3">
                                        <table style="width:100%;">
                                            <tr>
                                                <td align="left" valign="middle" width="50%">
                                                    <dx:ASPxLabel ID="ASPxLabel37" runat="server" Text="ETA PORT"
                                                        Font-Names="Tahoma" Font-Size="8pt" Width="100%">
                                                    </dx:ASPxLabel>
                                                </td>
                                                <td align="right" valign="middle" width="50%">
                                                    <dx:ASPxDateEdit ID="dtETAPort" runat="server" Width="100%" 
                                                        ClientInstanceName="dtETAPort" DisplayFormatString="yyyy-MM-dd" 
                                                        EditFormat="Custom" EditFormatString="yyyy-MM-dd" ReadOnly="True">
                                                    </dx:ASPxDateEdit>
                                                </td>
                                            </tr>
                                        </table>
                                    </td>
                                    <td align="left" colspan="3" class="style3">
                                        <table style="width:100%;">
                                            <tr>
                                                <td align="left" valign="middle" width="50%">
                                                    <dx:ASPxLabel ID="ASPxLabel38" runat="server" Text="ETA FACTORY"
                                                        Font-Names="Tahoma" Font-Size="8pt" Width="100%">
                                                    </dx:ASPxLabel>
                                                </td>
                                                <td align="right" valign="middle" width="50%">
                                                    <dx:ASPxDateEdit ID="dtETAFactory" runat="server" Width="100%" 
                                                        ClientInstanceName="dtETAFactory" DisplayFormatString="yyyy-MM-dd" 
                                                        EditFormat="Custom" EditFormatString="yyyy-MM-dd" ReadOnly="True">
                                                    </dx:ASPxDateEdit>
                                                </td>
                                            </tr>
                                        </table>
                                    </td>
                                    <td align="right" valign="middle">&nbsp;</td>

                                </tr>
                                                                                 
                                <tr>
                                    <td align="left" valign="middle" class="style1">
                                        <dx:ASPxLabel ID="ASPxLabel39" runat="server" Text="ORIGINAL O/NO"
                                            Font-Names="Tahoma" Font-Size="8pt">
                                        </dx:ASPxLabel>
                                    </td>
                                    <td align="left" valign="middle" colspan="2" class="style2">
                                            <dx:ASPxTextBox ID="txtpono" runat="server" Width="100%" 
                                            ClientInstanceName="txtpono" ReadOnly="True">
                                        </dx:ASPxTextBox>
                                    </td>
                                </tr>                               
                            </table>
                        </td>
                    </tr>
                </table>
            </td> 
        </tr>

        <tr>
            <td colspan="16" height="15">
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

    <div style="height: 1px;"></div>
                
    <table style="width: 100%;">
        <tr>
            <td colspan="16" align="left" valign="top" height="220">
                <%--Column : Grid--%>
                <dx:ASPxGridView ID="grid" runat="server" Width="100%"
                    Font-Names="Tahoma" KeyFieldName="NoUrut;PONo;PartNo;AffiliateID;SupplierID"
                    AutoGenerateColumns="False"
                    ClientInstanceName="grid" Font-Size="8pt">
                    <ClientSideEvents Init="OnInit" 
                        CallbackError="function(s, e) {e.handled = true;}" 
                        BatchEditEndEditing="OnBatchEditEndEditing" EndCallback="function(s, e) {
                        var pMsg = s.cpMessage;
                        if (pMsg != '') {
                            if (pMsg.substring(1,2) == '1') {
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
                    }" BatchEditStartEditing="OnBatchEditStartEditing" />                    
                    <Columns>
                        <dx:GridViewDataCheckColumn FieldName="AllowAccess" Name="AllowAccess" 
                            VisibleIndex="0" Width="30px"
                            Caption=" ">
                            <HeaderStyle HorizontalAlign="Center" VerticalAlign="Middle" />
                            <HeaderCaptionTemplate>
                            </HeaderCaptionTemplate>
                            <PropertiesCheckEdit ValueChecked="1" ValueType="System.Int32" ValueUnchecked="0">
                            </PropertiesCheckEdit>
                        </dx:GridViewDataCheckColumn>
                        <dx:GridViewDataTextColumn VisibleIndex="1" Caption="NO." FieldName="NoUrut" 
                            Name="NoUrut"
                            Width="30px" HeaderStyle-HorizontalAlign="Center">                            
                            <HeaderStyle HorizontalAlign="Center" Wrap="True"></HeaderStyle>
                            <CellStyle Font-Names="Tahoma" Font-Size="8pt" HorizontalAlign="Left">
                            </CellStyle>
                        </dx:GridViewDataTextColumn>
                        <dx:GridViewDataTextColumn VisibleIndex="2" Caption="PART NO." 
                            FieldName="PartNo" Width="100px" 
                        Name="PartNo" HeaderStyle-HorizontalAlign="Center">                            
                            <HeaderStyle HorizontalAlign="Center" Wrap="True"></HeaderStyle>
                            <CellStyle Font-Names="Tahoma" Font-Size="8pt">
                            </CellStyle>
                            </dx:GridViewDataTextColumn>
                        <dx:GridViewDataTextColumn VisibleIndex="3" Caption="PART NAME" 
                            FieldName="PartName" Width="180px" 
                        Name="PartName" HeaderStyle-HorizontalAlign="Center">                            
                            <HeaderStyle HorizontalAlign="Center"></HeaderStyle>
                            <CellStyle Font-Names="Tahoma" Font-Size="8pt">
                            </CellStyle>
                        </dx:GridViewDataTextColumn>
                        <dx:GridViewDataTextColumn VisibleIndex="4" Caption="UOM" Width="40px" 
                            HeaderStyle-HorizontalAlign="Center" CellStyle-HorizontalAlign="Center" 
                            FieldName="UOM" Name="UOM">                            
                            <HeaderStyle HorizontalAlign="Center"></HeaderStyle>
                            <CellStyle Font-Names="Tahoma" Font-Size="8pt">
                            </CellStyle>
                        </dx:GridViewDataTextColumn>
                        <dx:GridViewDataTextColumn VisibleIndex="5" Caption="MOQ" FieldName="MOQ" Name="MOQ"
                            Width="50px" HeaderStyle-HorizontalAlign="Center">                            
                            <PropertiesTextEdit DisplayFormatString="{0:n0}">
                            </PropertiesTextEdit>
                            <HeaderStyle HorizontalAlign="Center"></HeaderStyle>
                            <CellStyle Font-Names="Tahoma" Font-Size="8pt">
                            </CellStyle>
                        </dx:GridViewDataTextColumn>
                        <dx:GridViewDataTextColumn VisibleIndex="6" Caption="QTY BOX" 
                            FieldName="QtyBox" Name="QtyBox"
                            Width="75px" HeaderStyle-HorizontalAlign="Center">                            
                            <PropertiesTextEdit DisplayFormatString="{0:n0}">
                            </PropertiesTextEdit>
                            <HeaderStyle HorizontalAlign="Center"></HeaderStyle>
                            <CellStyle Font-Names="Tahoma" Font-Size="8pt">
                            </CellStyle>
                        </dx:GridViewDataTextColumn>
                        <dx:GridViewDataTextColumn VisibleIndex="9" Caption="CANCEL QTY *" 
                            FieldName="Week1" Name="Week1" Width="75px" 
                            HeaderStyle-HorizontalAlign="Center">                            
                            <PropertiesTextEdit DisplayFormatString="{0:n0}" MaxLength="18">
                            </PropertiesTextEdit>
                            <HeaderStyle Font-Names="Tahoma" Font-Size="8pt" Font-Underline="False" 
                                Wrap="True" HorizontalAlign="Center"/>
                            <CellStyle Font-Names="Tahoma" Font-Size="8pt">
                            </CellStyle>
                            <PropertiesTextEdit DisplayFormatString="{0:n0}" MaxLength="18">
                            <MaskSettings ErrorText="Please input valid value !"
                            Mask="<0..999999g>" IncludeLiterals="DecimalSymbol" />
                            </PropertiesTextEdit>
                        </dx:GridViewDataTextColumn>
                        <dx:GridViewDataTextColumn VisibleIndex="10" Caption="PREVIOUS FORECAST N" 
                            Width="75px" HeaderStyle-HorizontalAlign="Center" 
                            FieldName="PreviousForecast" Name="PreviousForecast">                            
                            <PropertiesTextEdit DisplayFormatString="{0:n0}">
                            </PropertiesTextEdit>
                            <HeaderStyle Font-Names="Tahoma" Font-Size="8pt" Font-Underline="False" 
                                Wrap="True" HorizontalAlign="Center"/>
                            <CellStyle Font-Names="Tahoma" Font-Size="8pt">
                            </CellStyle>
                        </dx:GridViewDataTextColumn>
                        <dx:GridViewDataTextColumn VisibleIndex="11" Caption="VARIANCE" Width="80px"
                            HeaderStyle-HorizontalAlign="Center" CellStyle-HorizontalAlign="Center" 
                            FieldName="Variance" Name="Variance">
                            <PropertiesTextEdit DisplayFormatString="{0:n0}">
                            </PropertiesTextEdit>                            
                            <HeaderStyle HorizontalAlign="Center"></HeaderStyle>
                            <CellStyle Font-Names="Tahoma" Font-Size="8pt" HorizontalAlign="Right">
                            </CellStyle>
                        </dx:GridViewDataTextColumn>
                        <dx:GridViewDataTextColumn VisibleIndex="12" Caption="VARIANCE %" Width="75px"
                            FieldName="VariancePercentage" Name="VariancePercentage">
                            <HeaderStyle HorizontalAlign="Center"></HeaderStyle>
                            <PropertiesTextEdit DisplayFormatString="{0:n2}">
                            </PropertiesTextEdit>
                            <CellStyle Font-Names="Tahoma" Font-Size="8pt" HorizontalAlign="Right">
                            </CellStyle>
                        </dx:GridViewDataTextColumn>
                         <dx:GridViewDataTextColumn VisibleIndex="13" Caption="FORECAST N+1" 
                            FieldName="Forecast1" Name="Forecast1" Width="75px" 
                            HeaderStyle-HorizontalAlign="Center">                            
                            <PropertiesTextEdit DisplayFormatString="{0:n0}" MaxLength="18">
                            </PropertiesTextEdit>
                            <HeaderStyle Font-Names="Tahoma" Font-Size="8pt" Font-Underline="False" 
                                Wrap="True" HorizontalAlign="Center"/>
                            <CellStyle Font-Names="Tahoma" Font-Size="8pt">
                            </CellStyle>
                            <PropertiesTextEdit DisplayFormatString="{0:n0}" MaxLength="18">
                            <MaskSettings ErrorText="Please input valid value !"
                            Mask="<0..999999g>" IncludeLiterals="DecimalSymbol" />
                            </PropertiesTextEdit>
                        </dx:GridViewDataTextColumn>
                        <dx:GridViewDataTextColumn VisibleIndex="14" Caption="FORECAST N+2" 
                            Width="75px" HeaderStyle-HorizontalAlign="Center" FieldName="Forecast2" 
                            Name="Forecast2">
                            <HeaderStyle Font-Names="Tahoma" Font-Size="8pt" Font-Underline="False" 
                                Wrap="True" HorizontalAlign="Center"/>
                            <CellStyle Font-Names="Tahoma" Font-Size="8pt">
                            </CellStyle>
                            <PropertiesTextEdit DisplayFormatString="{0:n0}" MaxLength="18">
                            <MaskSettings ErrorText="Please input valid value !"
                            Mask="<0..999999g>" IncludeLiterals="DecimalSymbol" />
                            </PropertiesTextEdit>
                        </dx:GridViewDataTextColumn>
                        <dx:GridViewDataTextColumn VisibleIndex="15" Caption="FORECAST N+3" 
                            Width="75px" HeaderStyle-HorizontalAlign="Center" FieldName="Forecast3" 
                            Name="Forecast3">
                            <HeaderStyle Font-Names="Tahoma" Font-Size="8pt" Font-Underline="False" 
                                Wrap="True" HorizontalAlign="Center"/>
                            <CellStyle Font-Names="Tahoma" Font-Size="8pt">
                            </CellStyle>
                            <PropertiesTextEdit DisplayFormatString="{0:n0}" MaxLength="18">
                            <MaskSettings ErrorText="Please input valid value !"
                            Mask="<0..999999g>" IncludeLiterals="DecimalSymbol" />
                            </PropertiesTextEdit>
                        </dx:GridViewDataTextColumn>
                        <dx:GridViewDataTextColumn Caption="AFFILIATE CODE" FieldName="AffiliateID" 
                        Name="AffiliateID"
                            VisibleIndex="16" Width="200px" Visible="False">
                        </dx:GridViewDataTextColumn>
                        <dx:GridViewDataTextColumn Caption="SUPPLIER CODE" FieldName="SupplierID" 
                        Name="SupplierID"
                            VisibleIndex="17" Width="80px">
                        </dx:GridViewDataTextColumn>
                        <dx:GridViewDataTextColumn FieldName="PONo" Width="100px" Caption="PO NO." 
                        Name="PONo"
                            VisibleIndex="20" Visible="False">
                            <HeaderStyle HorizontalAlign="Center"></HeaderStyle>
                            <CellStyle Font-Names="Tahoma" Font-Size="8pt"></CellStyle>
                        </dx:GridViewDataTextColumn>
                        <dx:GridViewDataTextColumn Caption="AdaData" FieldName="AdaData" 
                            Name="AdaData" VisibleIndex="23" Width="0px">
                            <HeaderStyle Font-Names="Tahoma" Font-Size="8pt" Font-Underline="False" 
                                Wrap="True" HorizontalAlign="Center"/>
                            <CellStyle Font-Names="Tahoma" Font-Size="8pt">
                            </CellStyle>
                        </dx:GridViewDataTextColumn>
                        <dx:GridViewDataTextColumn Caption="ErrorStatus" FieldName="ErrorStatus" 
                            VisibleIndex="24" Width="100px" Visible="False">
                        </dx:GridViewDataTextColumn>
                        <dx:GridViewDataTextColumn Caption="PO QTY" FieldName="CancelReffQty" 
                            VisibleIndex="8" Width="75px">
                            <HeaderStyle HorizontalAlign="Center" />
                            <PropertiesTextEdit DisplayFormatString="{0:n0}" MaxLength="18">
                            </PropertiesTextEdit>
                        </dx:GridViewDataTextColumn>
                        <dx:GridViewDataTextColumn Caption="CANCEL REFF PO" FieldName="CancelReffPONo" 
                            VisibleIndex="7" Width="100px">
                        </dx:GridViewDataTextColumn>
                    </Columns>
                    <SettingsPager Visible="False" Position="Top">
                        <Summary Text="Page {0} of {1} [{2} record(s)]" AllPagesText="Page {0} of {1} " />
                        <Summary AllPagesText="Page {0} of {1} " Text="Page {0} of {1} [{2} record(s)]">
                        </Summary>
                    </SettingsPager>
                    <SettingsEditing Mode="Batch">
                    </SettingsEditing>
                    <Settings ShowHorizontalScrollBar="True" ShowVerticalScrollBar="True" ShowGroupButtons="False"
                        ShowStatusBar="Hidden" VerticalScrollableHeight="250" />

                    <Settings ShowGroupButtons="False" ShowVerticalScrollBar="True" ShowHorizontalScrollBar="True" 
                        VerticalScrollableHeight="250" ShowStatusBar="Hidden"></Settings>
                    <Styles>
                                <Header BackColor="#FFD2A6" Font-Names="Verdana" Font-Size="8pt"></Header>
                                <Row BackColor="#FFFFE1" Font-Names="Verdana" Font-Size="8pt" Wrap="False"></Row>
                                <RowHotTrack BackColor="#E8EFFD" Font-Names="Verdana" Font-Size="8pt" Wrap="False"></RowHotTrack>
                                <SelectedRow Wrap="False">
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
                
    <div style="height: 8px;"></div>

    <table style="width: 100%;">
        <tr>
            <td align="left">
                <dx:ASPxButton ID="btnSubMenu" runat="server" Text="SUB MENU"
                    Font-Names="Tahoma" Width="85px" Font-Size="8pt" 
                    HorizontalAlign="Center" VerticalAlign="Bottom" AutoPostBack="False">
                </dx:ASPxButton>
            </td>
            <td>
                <dx:ASPxTextBox ID="txtHeijunka" runat="server" ClientInstanceName="txtHeijunka" 
                    Width="0px" BackColor="White" ForeColor="White">
                    <Border BorderColor="White" />
                </dx:ASPxTextBox>                          
            </td>
            <td>
                <dx:ASPxTextBox ID="txtMode" runat="server" ClientInstanceName="txtMode" 
                    Width="0px" BackColor="White" ForeColor="White">
                    <Border BorderColor="White" />
                </dx:ASPxTextBox>                          
            </td>
            <td align="right">
                <dx:ASPxTextBox ID="tampung" runat="server" ClientInstanceName="tampung" 
                    Width="0px" BackColor="White" ForeColor="White">
                    <Border BorderColor="White" />
                </dx:ASPxTextBox>                          
            </td>
            <td align="right">
                &nbsp;
            </td>
            <td align="right" width="85px">
                &nbsp;
            </td>
            <td align="right">
                &nbsp;
            </td>
            <td align="right" width="85px">
                <dx:ASPxButton ID="btnApprove" runat="server" Text="RECOVERY PO CANCEL"
                    Font-Names="Tahoma"
                    Width="80px" AutoPostBack="False" Font-Size="8pt" 
                    ClientInstanceName="btnApprove" VerticalAlign="Bottom">
                    <ClientSideEvents Click="function(s, e) {
			        if (cboAffiliate.GetText() == &quot;&quot;) {
                        lblInfo.GetMainElement().style.color = 'Red';
                        lblInfo.SetText(&quot;[6011] Please Select Affiliate first!&quot;);
                        cboAffiliateCode.Focus();
                        e.ProcessOnServer = false;
                        return false;
                    }

			        if (txtpono.GetText() == &quot;&quot;) {
				        lblInfo.GetMainElement().style.color = 'Red';
                        lblInfo.SetText(&quot;[6011] Please Input Order No first!&quot;);
                        txtpono.Focus();
                        e.ProcessOnServer = false;
                        return false;
                    }

			        if (grid.GetVisibleRowsOnPage() == 0){
        		        lblInfo.GetMainElement().style.color = 'Red';
	    		        llblInfo.SetText('[6013] No data to submit!');
        		        e.processOnServer = false;
        		        return false;
			        }

                    var millisecondsToWait = 100;

                    setTimeout(function() {ASPxCallback1.PerformCallback('recoverypocancel');
                    }, millisecondsToWait);	

                    setTimeout(function() {grid.PerformCallback('gridloadupdate');
                    }, millisecondsToWait);	

                    btnRecover.SetEnabled(false);
	
                    }"/> 
                </dx:ASPxButton>
            </td>
        </tr>
        <tr>
            <td valign="top" align="left">
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
            <td align="right" style="width:80px;">                                   
            </td>
        </tr>
    </table>
                    
    <dx:ASPxGlobalEvents ID="ge" runat="server" >
        <ClientSideEvents ControlsInitialized="function(s, e) {
	        OnControlsInitializedSplitter();
	        OnControlsInitializedGrid();
        }" />
    </dx:ASPxGlobalEvents>

    <dx:ASPxCallback ID="ASPxCallback1" runat="server" ClientInstanceName="ASPxCallback1">
        <ClientSideEvents EndCallback="function(s, e) {            
            var pMsg = s.cpMessage;

            if (pMsg != '') {
                if (pMsg.substring(1,2) == '1') {
                    lblInfo.GetMainElement().style.color = 'Blue';
                } else {
                    lblInfo.GetMainElement().style.color = 'Red';
                }
                lblInfo.SetText(pMsg);
            } else {
                lblInfo.SetText('');
            }
        
            delete s.cpMessage;

		    txtError.SetText(s.cpJumlahError);
        }" />       
    </dx:ASPxCallback>
</asp:Content>

