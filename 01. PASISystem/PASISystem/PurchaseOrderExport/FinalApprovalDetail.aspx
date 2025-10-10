<%@ Page Title="" Language="vb" AutoEventWireup="false" MasterPageFile="~/SiteContent.Master" CodeBehind="FinalApprovalDetail.aspx.vb" Inherits="PASISystem.FinalApprovalDetail" %>
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
                            <table id="Table1" align="left">
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
                                                    <dx:ASPxTextBox ID="txtperiod" runat="server" Width="200px" Font-Names="Tahoma"
                                                        Font-Size="8pt" BackColor="#CCCCCC" ReadOnly="True" MaxLength="20"
                                                        ClientInstanceName="txtperiod">
                                                    </dx:ASPxTextBox>
                                    </td>

                                    <td style="width:5px;"></td>

                                </tr>
                                <!-- ROW 3 -->
                                <tr>
                                    <td style="width:5px;"></td>
                                    <td align="left" valign="middle" >
                                        <dx:ASPxLabel ID="ASPxLabel1" runat="server" Text="ORDER NO"
                                            Font-Names="Tahoma" Font-Size="8pt">
                                        </dx:ASPxLabel>
                                    </td>
                                    <td align="left" valign="middle" >
                                                    <dx:ASPxTextBox ID="txtorderno" runat="server" Width="200px" Font-Names="Tahoma"
                                                        Font-Size="8pt" BackColor="#CCCCCC" ReadOnly="True" MaxLength="20"
                                                        ClientInstanceName="txtorderno">
                                                    </dx:ASPxTextBox>
                                    </td>

                                    <td style="width:5px;"></td>

                                </tr>
                                <!-- ROW 4 -->
                                <tr>
                                    <td style="width:5px;"></td>
                                    <td align="left" valign="middle" >
                                        <dx:ASPxLabel ID="ASPxLabel18" runat="server" Text="ORIGINAL ORDER NO"
                                            Font-Names="Tahoma" Font-Size="8pt">
                                        </dx:ASPxLabel>
                                    </td>
                                    <td align="left" valign="middle" >
                                                    <dx:ASPxTextBox ID="txtoriginal" runat="server" Width="200px" Font-Names="Tahoma"
                                                        Font-Size="8pt" BackColor="#CCCCCC" ReadOnly="True" MaxLength="20"
                                                        ClientInstanceName="txtoriginal">
                                                    </dx:ASPxTextBox>
                                    </td>

                                    <td style="width:5px;"></td>

                                </tr>
                                <!-- ROW 5 -->
                                <tr>
                                    <td style="width:5px;"></td>
                                    <td align="left" valign="middle" >
                                        <dx:ASPxLabel ID="ASPxLabel23" runat="server" Text="AFFILIATE"
                                            Font-Names="Tahoma" Font-Size="8pt">
                                        </dx:ASPxLabel>
                                    </td>
                                    <td align="left" valign="middle" >
                                                    <dx:ASPxTextBox ID="txtaffiliate" runat="server" Width="200px" Font-Names="Tahoma"
                                                        Font-Size="8pt" BackColor="#CCCCCC" ReadOnly="True" MaxLength="20"
                                                        ClientInstanceName="txtaffiliate">
                                                    </dx:ASPxTextBox>
                                    </td>

                                    <td style="width:5px;">
                                                    <dx:ASPxTextBox ID="txtaffname" runat="server" Width="350px" Font-Names="Tahoma"
                                                        Font-Size="8pt" BackColor="#CCCCCC" ReadOnly="True" MaxLength="20"
                                                        ClientInstanceName="txtaffname">
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
                                        <dx:ASPxLabel ID="ASPxLabel24" runat="server" Text="SUPPLIER"
                                            Font-Names="Tahoma" Font-Size="8pt">
                                        </dx:ASPxLabel>
                                    </td>
                                    <td align="left" valign="middle" >
                                                    <dx:ASPxTextBox ID="txtsupplier" runat="server" Width="200px" Font-Names="Tahoma"
                                                        Font-Size="8pt" BackColor="#CCCCCC" ReadOnly="True" MaxLength="20"
                                                        ClientInstanceName="txtsupplier">
                                                    </dx:ASPxTextBox>
                                                </td>

                                    <td style="width:5px;">
                                                    <dx:ASPxTextBox ID="txtsuppname" runat="server" Width="350px" Font-Names="Tahoma"
                                                        Font-Size="8pt" BackColor="#CCCCCC" ReadOnly="True" MaxLength="20"
                                                        ClientInstanceName="txtsuppname">
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
                    Font-Names="Tahoma" KeyFieldName="NoUrut;PONo;PartNo;AffiliateID;SupplierID"
                    AutoGenerateColumns="False"
                    ClientInstanceName="grid" Font-Size="8pt">
                    <ClientSideEvents Init="OnInit" 
                        CallbackError="function(s, e) {e.handled = true;}" 
                        BatchEditEndEditing="OnBatchEditEndEditing" EndCallback="function(s, e) {
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
                    }" BatchEditStartEditing="OnBatchEditStartEditing" />                    
                    <Columns>
                        <dx:GridViewDataCheckColumn FieldName="AllowAccess" Name="AllowAccess" 
                            VisibleIndex="0" Width="0px"
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
                        <dx:GridViewDataTextColumn VisibleIndex="6" Caption="UOM" Width="40px" 
                            HeaderStyle-HorizontalAlign="Center" CellStyle-HorizontalAlign="Center" 
                            FieldName="UOM" Name="UOM">                            
                            <HeaderStyle HorizontalAlign="Center"></HeaderStyle>
                            <CellStyle Font-Names="Tahoma" Font-Size="8pt">
                            </CellStyle>
                        </dx:GridViewDataTextColumn>
                        <dx:GridViewDataTextColumn VisibleIndex="7" Caption="MOQ" FieldName="MOQ" Name="MOQ"
                            Width="50px" HeaderStyle-HorizontalAlign="Center">                            
                            <PropertiesTextEdit DisplayFormatString="{0:n0}">
                            </PropertiesTextEdit>
                            <HeaderStyle HorizontalAlign="Center"></HeaderStyle>
                            <CellStyle Font-Names="Tahoma" Font-Size="8pt">
                            </CellStyle>
                        </dx:GridViewDataTextColumn>
                        <dx:GridViewDataTextColumn VisibleIndex="8" Caption="QTY BOX" 
                            FieldName="QtyBox" Name="QtyBox"
                            Width="75px" HeaderStyle-HorizontalAlign="Center">                            
                            <PropertiesTextEdit DisplayFormatString="{0:n0}">
                            </PropertiesTextEdit>
                            <HeaderStyle HorizontalAlign="Center"></HeaderStyle>
                            <CellStyle Font-Names="Tahoma" Font-Size="8pt">
                            </CellStyle>
                        </dx:GridViewDataTextColumn>
                        <dx:GridViewDataTextColumn VisibleIndex="9" Caption="TOTAL FIRM QTY *" 
                            FieldName="Week1" Name="Week1" Width="100px" 
                            HeaderStyle-HorizontalAlign="Center">                            
                            <PropertiesTextEdit DisplayFormatString="{0:n0}" MaxLength="18">
<MaskSettings Mask="&lt;0..999999g&gt;" IncludeLiterals="DecimalSymbol" ErrorText="Please input valid value !"></MaskSettings>
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
                            Width="100px" HeaderStyle-HorizontalAlign="Center" 
                            FieldName="PreviousForecast" Name="PreviousForecast">                            
                            <PropertiesTextEdit DisplayFormatString="{0:n0}">
                            </PropertiesTextEdit>
                            <HeaderStyle Font-Names="Tahoma" Font-Size="8pt" Font-Underline="False" 
                                Wrap="True" HorizontalAlign="Center"/>
                            <CellStyle Font-Names="Tahoma" Font-Size="8pt">
                            </CellStyle>
                        </dx:GridViewDataTextColumn>
                        <dx:GridViewDataTextColumn VisibleIndex="11" Caption="VARIANCE" Width="100px"
                            HeaderStyle-HorizontalAlign="Center" CellStyle-HorizontalAlign="Center" 
                            FieldName="Variance" Name="Variance">
                            <PropertiesTextEdit DisplayFormatString="{0:n0}">
                            </PropertiesTextEdit>                            
                            <HeaderStyle HorizontalAlign="Center"></HeaderStyle>
                            <CellStyle Font-Names="Tahoma" Font-Size="8pt" HorizontalAlign="Right">
                            </CellStyle>
                        </dx:GridViewDataTextColumn>
                        <dx:GridViewDataTextColumn VisibleIndex="12" Caption="VARIANCE %" Width="100px"
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
<MaskSettings Mask="&lt;0..999999g&gt;" IncludeLiterals="DecimalSymbol" ErrorText="Please input valid value !"></MaskSettings>
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
                            VisibleIndex="17" Width="100px">
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
                        <dx:GridViewBandColumn Caption="BOX NO" VisibleIndex="4">
                            <Columns>
                                <dx:GridViewDataTextColumn Caption="FROM" FieldName="box1" VisibleIndex="2">
                                    <HeaderStyle HorizontalAlign="Center" />
                                </dx:GridViewDataTextColumn>
                                <dx:GridViewDataTextColumn Caption="TO" FieldName="box2" VisibleIndex="3">
                                    <HeaderStyle HorizontalAlign="Center" />
                                </dx:GridViewDataTextColumn>
                            </Columns>
                            <HeaderStyle HorizontalAlign="Center" />
                        </dx:GridViewBandColumn>
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

    <div style="height:1px;"></div>
    
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
                &nbsp;</td>
        </tr>
    </table>
                    
    <dx:ASPxGlobalEvents ID="ge" runat="server" >
        <ClientSideEvents ControlsInitialized="function(s, e) {
	        OnControlsInitializedSplitter();
	        OnControlsInitializedGrid();
        }" />
    </dx:ASPxGlobalEvents>
</asp:Content>
