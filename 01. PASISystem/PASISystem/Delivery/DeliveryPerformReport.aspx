<%@ Page Title="" Language="vb" AutoEventWireup="false" MasterPageFile="~/SiteContent.Master" CodeBehind="DeliveryPerformReport.aspx.vb" Inherits="PASISystem.DeliveryPerformReport" %>
<%@ Register assembly="DevExpress.Web.v14.1, Version=14.1.6.0, Culture=neutral, PublicKeyToken=b88d1754d700e49a" namespace="DevExpress.Web.ASPxEditors" tagprefix="dx1" %>
<%@ Register assembly="DevExpress.Web.v14.1, Version=14.1.6.0, Culture=neutral, PublicKeyToken=b88d1754d700e49a" namespace="DevExpress.Web.ASPxGridView" tagprefix="dx" %>
<%@ Register assembly="DevExpress.Web.v14.1, Version=14.1.6.0, Culture=neutral, PublicKeyToken=b88d1754d700e49a" namespace="DevExpress.Web.ASPxGridView.Export" tagprefix="dx2" %>
<%@ Register Assembly="DevExpress.Web.v14.1, Version=14.1.6.0, Culture=neutral, PublicKeyToken=b88d1754d700e49a" namespace="DevExpress.Web.ASPxGlobalEvents" TagPrefix="dx" %>
<%@ Register Assembly="DevExpress.Web.v14.1, Version=14.1.6.0, Culture=neutral, PublicKeyToken=b88d1754d700e49a" Namespace="DevExpress.Web.ASPxGridView.Export" TagPrefix="dx" %>
<asp:Content ID="Content1" ContentPlaceHolderID="head" runat="server">
</asp:Content>
<asp:Content ID="Content2" ContentPlaceHolderID="Content" runat="server">
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
        height = height - (height * 58 / 100)
        grid.SetHeight(height);
    }
</script>
    <table style="width:100%;" width="100%">
        <tr>
            <td align="left" style="width:20px;">
                <dx1:ASPxLabel ID="ASPxLabel1" runat="server" Font-Names="Tahoma" Font-Size="8pt"
                    Text="PERIOD">
                </dx1:ASPxLabel>
            </td>
            <td align="left" style="width:20px;">
                <dx1:ASPxDateEdit ID="dtPeriod" runat="server" ClientInstanceName="dtPeriod" 
                    Font-Size="8pt" EditFormat="Custom" EditFormatString="MMM yyyy">
                    <ClientSideEvents DateChanged="function(s, e) {
	grid.PerformCallback('kosong');
}" />
                </dx1:ASPxDateEdit>
            </td>
            <td align="left" style="width:20px;">
                
            </td>
            <td align="left">
                
            </td>
            <td align="left"style="width:20px;">
                
            </td>            
        </tr>
        <tr>
            <td align="left">
                <dx1:ASPxLabel ID="ASPxLabel2" runat="server" Font-Names="Tahoma" Font-Size="8pt"
                    Text="SUPPLIER GROUP">
                </dx1:ASPxLabel>
            </td>
            <td align="left">
                <dx1:ASPxComboBox ID="cboSupplierGroup" runat="server" ClientInstanceName="cboSupplierGroup"
                    Font-Names="Tahoma" TextFormatString="{0}" Font-Size="8pt">
                    <ClientSideEvents SelectedIndexChanged="function(s, e) {
	                    txtSupplierGroup.SetText(cboSupplierGroup.GetSelectedItem().GetColumnText(1));
                                    }" />
                </dx1:ASPxComboBox>
            </td>
            <td align="left">
                <dx1:ASPxTextBox ID="txtSupplierGroup" runat="server" Font-Names="Tahoma" 
                    Font-Size="8pt" Width="170px" BackColor="Silver" 
                    ClientInstanceName="txtSupplierGroup" ReadOnly="True">
                </dx1:ASPxTextBox>
            </td>
            <td align="left">
                &nbsp;</td>
            <td align="left">
                &nbsp;</td>            
        </tr>
        <tr>
            <td align="left">
                <dx1:ASPxLabel ID="ASPxLabel3" runat="server" Font-Names="Tahoma" Font-Size="8pt"
                    Text="PARTCODE/NAME">
                </dx1:ASPxLabel>
            </td>
            <td align="left">
                <dx1:ASPxComboBox ID="cboPartCode" runat="server" ClientInstanceName="cboPartCode"
                    Font-Names="Tahoma" TextFormatString="{0}" Font-Size="8pt">
                    <ClientSideEvents SelectedIndexChanged="function(s, e) {
                        txtPartCode.SetText(cboPartCode.GetSelectedItem().GetColumnText(1));
                                    }" />
                </dx1:ASPxComboBox>
            </td>
            <td align="left">
                <dx1:ASPxTextBox ID="txtPartCode" runat="server" Font-Names="Tahoma" 
                    Font-Size="8pt" Width="170px" BackColor="Silver" 
                    ClientInstanceName="txtPartCode" ReadOnly="True">
                </dx1:ASPxTextBox>
            </td>
            <td align="right">
                
            </td>
            <td>
                
            </td>
        </tr>  
        <tr>
            <td align="left">
                <dx1:ASPxLabel ID="ASPxLabel4" runat="server" Font-Names="Tahoma" Font-Size="8pt"
                    Text="AFFILIATE CODE">
                </dx1:ASPxLabel>
            </td>
            <td align="left">
                <dx1:ASPxComboBox ID="cboAffiliateCode" runat="server" ClientInstanceName="cboAffiliateCode"
                    Font-Names="Tahoma" TextFormatString="{0}" Font-Size="8pt">
                    <ClientSideEvents SelectedIndexChanged="function(s, e) {
                        txtAffiliateCode.SetText(cboAffiliateCode.GetSelectedItem().GetColumnText(1));
                                    }" />
                </dx1:ASPxComboBox>
            </td>
            <td align="left">
                <dx1:ASPxTextBox ID="txtAffiliateCode" runat="server" Font-Names="Tahoma" 
                    Font-Size="8pt" Width="170px" BackColor="Silver" 
                    ClientInstanceName="txtAffiliateCode" ReadOnly="True">
                </dx1:ASPxTextBox>
            </td>
            <td align="right">
                <dx1:ASPxButton ID="btnSearch" runat="server" Text="SEARCH" Width="85px" Font-Names="Tahoma"
                    Font-Size="8pt" ClientInstanceName="btnSearch" AutoPostBack="False">
                    <ClientSideEvents Click="function(s, e) {                                                 
	                            grid.PerformCallback('gridload');
	                            lblerrmessage.SetText('');                                                                                        
                            }" />
                </dx1:ASPxButton>
            </td>
            <td>
                <dx1:ASPxButton ID="btnClear" runat="server" Text="CLEAR" Width="85px" Font-Names="Tahoma"
                    Font-Size="8pt">                    
                    <ClientSideEvents Click="function(s, e) {                                                 
	                            grid.PerformCallback('kosong');
	                            lblerrmessage.SetText('');
                            }" />
                </dx1:ASPxButton>
            </td>
        </tr>     
    </table>
    <br />
            <table align="left" width="100%">
                <tr align="left">
                    <td width="100%" height="16px" style="border-top-style: solid; border-top-width: thin;
                        border-top-color: #808080; border-bottom-style: solid; border-bottom-width: thin;
                        border-bottom-color: #808080" align="left">
                        <dx1:ASPxLabel ID="lblerrmessage" runat="server" Font-Names="Tahoma" Font-Size="8pt"
                            Text="ERROR MESSAGE" Width="100%" ClientInstanceName="lblerrmessage">
                        </dx1:ASPxLabel>
                    </td>
                </tr>                
            </table>
        <br />
    <br />
    <table style="width:100%;" width="100%">
        <tr>
            <td align="left" class="style1" colspan = "2">
                <dx:ASPxGridView ID="grid" runat="server" AutoGenerateColumns="False" 
                    Width="100%" KeyFieldName="PerformanceCls;PlanActual;SupplierGroupCode;PartNo;ItemNo" ClientInstanceName="grid">
                    <ClientSideEvents EndCallback="function(s, e) {
						var pMsg = s.cpMessage;
                        if (pMsg != '') {
                            if (pMsg.substring(1,5) == '1001' || pMsg.substring(1,5) == '1002' || pMsg.substring(1,5) == '1003' || pMsg.substring(1,5) == '2001') {
                                lblerrmessage.GetMainElement().style.color = 'Blue';
                            } else {
                                lblerrmessage.GetMainElement().style.color = 'Red';
                            }
                            lblerrmessage.SetText(pMsg);
                        } else {
                            lblerrmessage.SetText('');
                        }
                        delete s.cpMessage;
}" RowClick="function(s, e) {
	lblerrmessage.SetText('');
}" Init="function(s, e) {
	
}" />
                    <Columns>                           
                        <dx:GridViewDataTextColumn Caption="NO" FieldName="No" Name="No" 
                            VisibleIndex="3" Width="30px">
                            <HeaderStyle Font-Names="Tahoma" Font-Size="8pt" Font-Underline="False" 
                                Wrap="True" HorizontalAlign="Center" />
                            <CellStyle Font-Names="Tahoma" Font-Size="8pt">
                            </CellStyle>
                        </dx:GridViewDataTextColumn>
                        <dx:GridViewDataTextColumn Caption="PERFORMANCE CLS" FieldName="PerformanceCls" 
                            Name="PerformanceCls" VisibleIndex="4" Width="100px">
                            <HeaderStyle Font-Names="Tahoma" Font-Size="8pt" Font-Underline="False" 
                                Wrap="True" HorizontalAlign="Center"/>
                            <CellStyle Font-Names="Tahoma" Font-Size="8pt">
                            </CellStyle>
                        </dx:GridViewDataTextColumn>
                        <dx:GridViewDataTextColumn Caption="PLAN/ACTUAL" 
                            FieldName="PlanActual" Name="PlanActual" VisibleIndex="5" 
                            Width="100px">
                            <HeaderStyle Font-Names="Tahoma" Font-Size="8pt" Font-Underline="False" 
                                Wrap="True" HorizontalAlign="Center"/>
                            <CellStyle Font-Names="Tahoma" Font-Size="8pt">
                            </CellStyle>
                        </dx:GridViewDataTextColumn>
                        <dx:GridViewDataTextColumn Caption="SUPPLIER GROUP" 
                            FieldName="SupplierGroupCode" Name="SupplierGroupCode" VisibleIndex="6" 
                            Width="120px">
                            <HeaderStyle Font-Names="Tahoma" Font-Size="8pt" Font-Underline="False" 
                                Wrap="True" HorizontalAlign="Center"/>
                            <CellStyle Font-Names="Tahoma" Font-Size="8pt">
                            </CellStyle>
                        </dx:GridViewDataTextColumn>
                        <dx:GridViewDataTextColumn Caption="PART NO" 
                            FieldName="PartNo" Name="PartNo" VisibleIndex="7" 
                            Width="120px">
                            <HeaderStyle Font-Names="Tahoma" Font-Size="8pt" Font-Underline="False" 
                                Wrap="True" HorizontalAlign="Center"/>
                            <CellStyle Font-Names="Tahoma" Font-Size="8pt">
                            </CellStyle>
                        </dx:GridViewDataTextColumn>
                        <dx:GridViewDataTextColumn Caption="ITEM NO" 
                            FieldName="ItemNo" Name="ItemNo" VisibleIndex="8" 
                            Width="100px">
                            <HeaderStyle Font-Names="Tahoma" Font-Size="8pt" Font-Underline="False" 
                                Wrap="True" HorizontalAlign="Center"/>
                            <CellStyle Font-Names="Tahoma" Font-Size="8pt">
                            </CellStyle>
                        </dx:GridViewDataTextColumn>
                        <dx:GridViewBandColumn Caption="MONTH" VisibleIndex="9">
                            <Columns>
                            <dx:GridViewDataTextColumn Caption="Jan-2015" FieldName="Month01" Name="Month01" VisibleIndex="1" Width="70px">
                                <PropertiesTextEdit DisplayFormatString="{0}">
                                </PropertiesTextEdit>
                                <HeaderStyle Font-Names="Tahoma" Font-Size="8pt" HorizontalAlign="Center" VerticalAlign="Middle"
                                    Wrap="True" />
                                <CellStyle Font-Names="Tahoma" Font-Size="8pt">
                                </CellStyle>
                            </dx:GridViewDataTextColumn>
                            <dx:GridViewDataTextColumn Caption="Feb-2015" FieldName="Month02" Name="Month02" VisibleIndex="2" Width="70px">
                                <PropertiesTextEdit DisplayFormatString="{0}">
                                </PropertiesTextEdit>
                                <HeaderStyle Font-Names="Tahoma" Font-Size="8pt" HorizontalAlign="Center" VerticalAlign="Middle"
                                    Wrap="True" />
                                <CellStyle Font-Names="Tahoma" Font-Size="8pt">
                                </CellStyle>
                            </dx:GridViewDataTextColumn>
                            <dx:GridViewDataTextColumn Caption="Mar-2015" FieldName="Month03" Name="Month03" VisibleIndex="3" Width="70px">
                                <PropertiesTextEdit DisplayFormatString="{0}">
                                </PropertiesTextEdit>
                                <HeaderStyle Font-Names="Tahoma" Font-Size="8pt" HorizontalAlign="Center" VerticalAlign="Middle"
                                    Wrap="True" />
                                <CellStyle Font-Names="Tahoma" Font-Size="8pt">
                                </CellStyle>
                            </dx:GridViewDataTextColumn>
                            <dx:GridViewDataTextColumn Caption="Apr-2015" FieldName="Month04" Name="Month04" VisibleIndex="4" Width="70px">
                                <PropertiesTextEdit DisplayFormatString="{0}">
                                </PropertiesTextEdit>
                                <HeaderStyle Font-Names="Tahoma" Font-Size="8pt" HorizontalAlign="Center" VerticalAlign="Middle"
                                    Wrap="True" />
                                <CellStyle Font-Names="Tahoma" Font-Size="8pt">
                                </CellStyle>
                            </dx:GridViewDataTextColumn>
                            <dx:GridViewDataTextColumn Caption="May-2015" FieldName="Month05" Name="Month05" VisibleIndex="5" Width="70px">
                                <PropertiesTextEdit DisplayFormatString="{0}">
                                </PropertiesTextEdit>
                                <HeaderStyle Font-Names="Tahoma" Font-Size="8pt" HorizontalAlign="Center" VerticalAlign="Middle"
                                    Wrap="True" />
                                <CellStyle Font-Names="Tahoma" Font-Size="8pt">
                                </CellStyle>
                            </dx:GridViewDataTextColumn>
                            <dx:GridViewDataTextColumn Caption="Jun-2015" FieldName="Month06" Name="Month06" VisibleIndex="6" Width="70px">
                                <PropertiesTextEdit DisplayFormatString="{0}">
                                </PropertiesTextEdit>
                                <HeaderStyle Font-Names="Tahoma" Font-Size="8pt" HorizontalAlign="Center" VerticalAlign="Middle"
                                    Wrap="True" />
                                <CellStyle Font-Names="Tahoma" Font-Size="8pt">
                                </CellStyle>
                            </dx:GridViewDataTextColumn>
                            <dx:GridViewDataTextColumn Caption="Jul-2015" FieldName="Month07" Name="Month07" VisibleIndex="7" Width="70px">
                                <PropertiesTextEdit DisplayFormatString="{0}">
                                </PropertiesTextEdit>
                                <HeaderStyle Font-Names="Tahoma" Font-Size="8pt" HorizontalAlign="Center" VerticalAlign="Middle"
                                    Wrap="True" />
                                <CellStyle Font-Names="Tahoma" Font-Size="8pt">
                                </CellStyle>
                            </dx:GridViewDataTextColumn>
                            <dx:GridViewDataTextColumn Caption="Aug-2015" FieldName="Month08" Name="Month08" VisibleIndex="8" Width="70px">
                                <PropertiesTextEdit DisplayFormatString="{0}">
                                </PropertiesTextEdit>
                                <HeaderStyle Font-Names="Tahoma" Font-Size="8pt" HorizontalAlign="Center" VerticalAlign="Middle"
                                    Wrap="True" />
                                <CellStyle Font-Names="Tahoma" Font-Size="8pt">
                                </CellStyle>
                            </dx:GridViewDataTextColumn>
                            <dx:GridViewDataTextColumn Caption="Sep-2015" FieldName="Month09" Name="Month09" VisibleIndex="9" Width="70px">
                                <PropertiesTextEdit DisplayFormatString="{0}">
                                </PropertiesTextEdit>
                                <HeaderStyle Font-Names="Tahoma" Font-Size="8pt" HorizontalAlign="Center" VerticalAlign="Middle"
                                    Wrap="True" />
                                <CellStyle Font-Names="Tahoma" Font-Size="8pt">
                                </CellStyle>
                            </dx:GridViewDataTextColumn>
                            <dx:GridViewDataTextColumn Caption="Oct-2015" FieldName="Month10" Name="Month10" VisibleIndex="10" Width="70px">
                                <PropertiesTextEdit DisplayFormatString="{0}">
                                </PropertiesTextEdit>
                                <HeaderStyle Font-Names="Tahoma" Font-Size="8pt" HorizontalAlign="Center" VerticalAlign="Middle"
                                    Wrap="True" />
                                <CellStyle Font-Names="Tahoma" Font-Size="8pt">
                                </CellStyle>
                            </dx:GridViewDataTextColumn>
                            <dx:GridViewDataTextColumn Caption="Nov-2015" FieldName="Month11" Name="Month11" VisibleIndex="11" Width="70px">
                                <PropertiesTextEdit DisplayFormatString="{0}">
                                </PropertiesTextEdit>
                                <HeaderStyle Font-Names="Tahoma" Font-Size="8pt" HorizontalAlign="Center" VerticalAlign="Middle"
                                    Wrap="True" />
                                <CellStyle Font-Names="Tahoma" Font-Size="8pt">
                                </CellStyle>
                            </dx:GridViewDataTextColumn>
                            <dx:GridViewDataTextColumn Caption="Dec-2015" FieldName="Month12" Name="Month12" VisibleIndex="12" Width="70px">
                                <PropertiesTextEdit DisplayFormatString="{0}">
                                </PropertiesTextEdit>
                                <HeaderStyle Font-Names="Tahoma" Font-Size="8pt" HorizontalAlign="Center" VerticalAlign="Middle"
                                    Wrap="True" />
                                <CellStyle Font-Names="Tahoma" Font-Size="8pt">
                                </CellStyle>
                            </dx:GridViewDataTextColumn>                            
                            </Columns>
                            <HeaderStyle HorizontalAlign="Center" />
                        </dx:GridViewBandColumn>
                        <dx:GridViewDataTextColumn Caption="ON TIME QTY" FieldName="OntimeQty" 
                            Name="OntimeQty" VisibleIndex="10" Width="70px">
                            <HeaderStyle Font-Names="Tahoma" Font-Size="8pt" HorizontalAlign="Center" VerticalAlign="Middle"
                                Wrap="True" />
                            <CellStyle Font-Names="Tahoma" Font-Size="8pt">
                            </CellStyle>
                        </dx:GridViewDataTextColumn>
                        <dx:GridViewDataTextColumn Caption="PLANNED QTY" FieldName="PlannedQty" 
                            Name="PlannedQty" VisibleIndex="11" Width="70px">
                            <HeaderStyle Font-Names="Tahoma" Font-Size="8pt" HorizontalAlign="Center" VerticalAlign="Middle"
                                Wrap="True" />
                            <CellStyle Font-Names="Tahoma" Font-Size="8pt">
                            </CellStyle>
                        </dx:GridViewDataTextColumn>
                        <dx:GridViewDataTextColumn Caption="ON TIME LINE" FieldName="OntimeLine" 
                            Name="OntimeLine" VisibleIndex="12" Width="70px">
                            <HeaderStyle Font-Names="Tahoma" Font-Size="8pt" HorizontalAlign="Center" VerticalAlign="Middle"
                                Wrap="True" />
                            <CellStyle Font-Names="Tahoma" Font-Size="8pt">
                            </CellStyle>
                        </dx:GridViewDataTextColumn>
                        <dx:GridViewDataTextColumn Caption="PLANNED LINE" FieldName="PlannedLine" 
                            Name="PlannedLine" VisibleIndex="13" Width="70px">
                            <HeaderStyle Font-Names="Tahoma" Font-Size="8pt" HorizontalAlign="Center" VerticalAlign="Middle"
                                Wrap="True" />
                            <CellStyle Font-Names="Tahoma" Font-Size="8pt">
                            </CellStyle>
                        </dx:GridViewDataTextColumn>
                        <dx:GridViewDataTextColumn Caption="CUMM QTY" FieldName="CummQtySub" 
                            Name="CummQtySub" VisibleIndex="14" Width="70px">
                            <HeaderStyle Font-Names="Tahoma" Font-Size="8pt" HorizontalAlign="Center" VerticalAlign="Middle"
                                Wrap="True" />
                            <CellStyle Font-Names="Tahoma" Font-Size="8pt">
                            </CellStyle>
                        </dx:GridViewDataTextColumn>
                        <dx:GridViewDataTextColumn Caption="AQ (%)" FieldName="AQ" Name="AQ" 
                            VisibleIndex="15" Width="70px">
                            <HeaderStyle Font-Names="Tahoma" Font-Size="8pt" HorizontalAlign="Center" VerticalAlign="Middle"
                                Wrap="True" />
                            <CellStyle Font-Names="Tahoma" Font-Size="8pt">
                            </CellStyle>
                        </dx:GridViewDataTextColumn>
                        <dx:GridViewDataTextColumn Caption="AL (%)" FieldName="AL" Name="AL" VisibleIndex="19" Width="70px">
                            <HeaderStyle Font-Names="Tahoma" Font-Size="8pt" HorizontalAlign="Center" VerticalAlign="Middle"
                                Wrap="True" />
                            <CellStyle Font-Names="Tahoma" Font-Size="8pt">
                            </CellStyle>
                        </dx:GridViewDataTextColumn>
                        <dx:GridViewDataTextColumn Caption="CQ (%)" FieldName="CQ" Name="CQ" VisibleIndex="20" Width="70px">
                            <HeaderStyle Font-Names="Tahoma" Font-Size="8pt" HorizontalAlign="Center" VerticalAlign="Middle"
                                Wrap="True" />
                            <CellStyle Font-Names="Tahoma" Font-Size="8pt">
                            </CellStyle>
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
                    <Styles>
                        <SelectedRow ForeColor="Black">
                        </SelectedRow>
                    </Styles>
                    <StylesEditors ButtonEditCellSpacing="0">
                        <ProgressBar Height="21px">
                        </ProgressBar>
                    </StylesEditors>
                </dx:ASPxGridView>
                <dx:ASPxGridViewExporter ID="gridExport" runat="server" GridViewID="grid" >
                </dx:ASPxGridViewExporter>
                <br />
            </td>
        </tr>
        <tr>
            <td align="left">
                <dx1:ASPxButton ID="btnSubmenu" runat="server" Text="SUB MENU" Width="90px" Font-Names="Tahoma"
                    Font-Size="8pt">
                </dx1:ASPxButton>
            </td>
            <td align="right" width="50px">
                <dx1:ASPxButton ID="btnDownload" runat="server" Text="DOWNLOAD" Width="90px" Font-Names="Tahoma"
                    Font-Size="8pt" AutoPostBack="False">
                    <ClientSideEvents Click="function(s, e) {
    
}" />
                </dx1:ASPxButton>
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
