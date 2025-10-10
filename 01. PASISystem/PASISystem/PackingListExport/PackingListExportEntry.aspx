<%@ Page Title="" Language="vb" AutoEventWireup="false" MasterPageFile="~/SiteContent.Master" CodeBehind="PackingListExportEntry.aspx.vb" Inherits="PASISystem.PackingListExportEntry" %>
<%@ Register Assembly="DevExpress.Web.v14.1, Version=14.1.6.0, Culture=neutral, PublicKeyToken=b88d1754d700e49a"
    Namespace="DevExpress.Web.ASPxEditors" TagPrefix="dx1" %>
<%@ Register Assembly="DevExpress.Web.v14.1, Version=14.1.6.0, Culture=neutral, PublicKeyToken=b88d1754d700e49a"
    Namespace="DevExpress.Web.ASPxGridView" TagPrefix="dx" %>
<%@ Register Assembly="DevExpress.Web.v14.1, Version=14.1.6.0, Culture=neutral, PublicKeyToken=b88d1754d700e49a"
    Namespace="DevExpress.Web.ASPxGridView.Export" TagPrefix="dx" %>
<%@ Register Assembly="DevExpress.Web.v14.1, Version=14.1.6.0, Culture=neutral, PublicKeyToken=b88d1754d700e49a"
    Namespace="DevExpress.Web.ASPxGlobalEvents" TagPrefix="dx" %>
<%@ Register Assembly="DevExpress.Web.v14.1, Version=14.1.6.0, Culture=neutral, PublicKeyToken=b88d1754d700e49a"
    Namespace="DevExpress.Web.ASPxHiddenField" TagPrefix="dx" %>
<%@ Register Assembly="DevExpress.Web.v14.1, Version=14.1.6.0, Culture=neutral, PublicKeyToken=b88d1754d700e49a"
    Namespace="DevExpress.Web.ASPxRoundPanel" TagPrefix="dx" %>
<%@ Register Assembly="DevExpress.Web.v14.1, Version=14.1.6.0, Culture=neutral, PublicKeyToken=b88d1754d700e49a"
    Namespace="DevExpress.Web.ASPxPanel" TagPrefix="dx" %>

<asp:Content ID="Content1" ContentPlaceHolderID="head" runat="server">
    <style type="text/css">
        .dxeBase
        {
            font: 12px Tahoma, Geneva, sans-serif;
        }
        
        .dxeBase
        {
            font: 12px Tahoma, Geneva, sans-serif;
        }
        
        .style1
        {
            height: 331px;
        }
    </style>
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
    <table style="width: 100%;" width="100%">
        <tr>
            <td align="left">
                <dx1:ASPxComboBox ID="cbopart2" runat="server" ClientInstanceName="cbopart" Font-Names="Tahoma"
                    TextFormatString="{0}" Font-Size="8pt" Width="100px">
                    <ClientSideEvents SelectedIndexChanged="function(s, e) {txtpart.SetText(cbopart.GetSelectedItem().GetColumnText(1));}" />
                </dx1:ASPxComboBox>
            </td>
            <td align="left">
                &nbsp;</td>
            <td align="left">
                &nbsp;</td>
            <td align="left">
                &nbsp;</td>
            <td align="left">
                &nbsp;</td>
            <td align="left" rowspan="6" valign="top" width="50%">
                <dx:ASPxRoundPanel ID="ASPxRoundPanel1" runat="server" 
                    HeaderText="FILTER FORWARDER RECEIVING" ShowCollapseButton="true"
                    View="GroupBox" Width="100%" Font-Size="8pt" Font-Names="Tahoma">
                    <PanelCollection>
                        <dx:panelcontent ID="PanelContent1" runat="server">
                            <table id="Table2">
                                <tr>
                                    <td align="left" height="20px" valign="middle" width="50%">
                                       
                                        <dx1:ASPxLabel ID="ASPxLabel13" runat="server" Font-Names="Tahoma" 
                                            Font-Size="8pt" Text="SUPPLIER CODE/NAME">
                                        </dx1:ASPxLabel>
                                       
                                    </td>
                                    <td align="left" height="20px" valign="middle" width="50%">
                                        
                                        <dx1:ASPxComboBox ID="cbopart1" runat="server" ClientInstanceName="cbopart" 
                                            Font-Names="Tahoma" Font-Size="8pt" TextFormatString="{0}">
                                            <ClientSideEvents SelectedIndexChanged="function(s, e) {txtpart.SetText(cbopart.GetSelectedItem().GetColumnText(1));}" />
                                        </dx1:ASPxComboBox>
                                        
                                    </td>
                                    <td align="left" height="20px" valign="middle" width="50%">
                                        <dx1:ASPxTextBox ID="txtpart0" runat="server" BackColor="Silver" 
                                            ClientInstanceName="txtpart" Font-Names="Tahoma" Font-Size="8pt" 
                                            ReadOnly="True" Width="170px">
                                            <ClientSideEvents TextChanged="function(s, e) {lblerrmessage.SetText('');}" />
                                        </dx1:ASPxTextBox>
                                    </td>
                                </tr>
                                <tr>
                                    <td align="left" height="20px" valign="middle" width="50%">
                                        
                                        <dx1:ASPxLabel ID="ASPxLabel8" runat="server" Font-Names="Tahoma" 
                                            Font-Size="8pt" Text="PART CODE/NAME">
                                        </dx1:ASPxLabel>
                                        
                                    </td>
                                    <td align="left" height="20px" valign="middle" width="50%">
                                      
                                        <dx1:ASPxComboBox ID="cbopart" runat="server" ClientInstanceName="cbopart" 
                                            Font-Names="Tahoma" Font-Size="8pt" TextFormatString="{0}">
                                            <ClientSideEvents SelectedIndexChanged="function(s, e) {txtpart.SetText(cbopart.GetSelectedItem().GetColumnText(1));}" />
                                        </dx1:ASPxComboBox>
                                      
                                    </td>
                                    <td align="left" height="20px" valign="middle" width="50%">
                                        <dx1:ASPxTextBox ID="txtpart" runat="server" BackColor="Silver" 
                                            ClientInstanceName="txtpart" Font-Names="Tahoma" Font-Size="8pt" 
                                            ReadOnly="True" Width="170px">
                                            <ClientSideEvents TextChanged="function(s, e) {lblerrmessage.SetText('');}" />
                                        </dx1:ASPxTextBox>
                                    </td>
                                </tr>
                                <tr>
                                    <td align="left" height="20px" valign="middle" width="50%">
                                        
                                        <dx1:ASPxLabel ID="ASPxLabel9" runat="server" Font-Names="Tahoma" 
                                            Font-Size="8pt" Text="PO NO.">
                                        </dx1:ASPxLabel>
                                        
                                    </td>
                                    <td align="left" height="20px" valign="middle" width="50%">
                                       
                                        <dx1:ASPxTextBox ID="txtpono" runat="server" ClientInstanceName="txtpono" 
                                            Font-Names="Tahoma" Font-Size="8pt" Width="170px">
                                        </dx1:ASPxTextBox>
                                       
                                    </td>
                                    <td align="left" height="20px" valign="middle" width="50%">
                                        &nbsp;</td>
                                </tr>
                                <tr>
                                    <td align="left" height="20px" valign="middle" width="50%">
                                        &nbsp;</td>
                                    <td align="left" height="20px" valign="middle" width="50%">
                                        &nbsp;</td>
                                    <td align="left" height="20px" valign="middle" width="50%">
                                        &nbsp;</td>
                                </tr>
                                <tr>
                                    <td align="left" height="20px" valign="middle" width="50%">
                                        &nbsp;</td>
                                    <td align="left" height="20px" valign="middle" width="50%">
                                        &nbsp;</td>
                                    <td align="left" height="20px" valign="middle" width="50%">
                                        &nbsp;</td>
                                </tr>
                                <tr>
                                    <td align="left" height="20px" valign="middle" width="50%">
                                        &nbsp;</td>
                                    <td align="left" height="20px" valign="middle" width="50%">
                                        &nbsp;</td>
                                    <td align="left" height="20px" valign="middle" width="50%">
                                        &nbsp;</td>
                                </tr>
                            </table>
                        </dx:panelcontent>
                    </PanelCollection>
                </dx:ASPxRoundPanel>
            </td>
        </tr>
        <tr>
            <td align="left">
                <dx1:ASPxLabel ID="ASPxLabel1" runat="server" Font-Names="Tahoma" Font-Size="8pt"
                    Text="PACKING LIST DATE (FILTER)">
                </dx1:ASPxLabel>
            </td>
            <td align="left">
                <dx1:ASPxCheckBox ID="checkbox1" runat="server" CheckState="Unchecked" ClientInstanceName="checkbox1"
                    Text=" ">
                </dx1:ASPxCheckBox>
            </td>
            <td align="left">
                <dx1:ASPxDateEdit ID="dtfrom" runat="server" ClientInstanceName="dtfrom" Font-Size="8pt"
                    EditFormat="Custom" EditFormatString="dd MMM yyyy">
                </dx1:ASPxDateEdit>
            </td>
            <td align="left">
                <dx1:ASPxLabel ID="ASPxLabel6" runat="server" Font-Names="Tahoma" Font-Size="8pt"
                    Text="~">
                </dx1:ASPxLabel>
                &nbsp;
            </td>
            <td align="left">
                <dx1:ASPxDateEdit ID="dtto" runat="server" ClientInstanceName="dtto" Font-Size="8pt"
                    EditFormat="Custom" EditFormatString="dd MMM yyyy">
                </dx1:ASPxDateEdit>
            </td>
        </tr>
        <tr>
            <td align="left">
                <dx1:ASPxLabel ID="ASPxLabel7" runat="server" Font-Names="Tahoma" Font-Size="8pt"
                    Text="AFFILIATE CODE/NAME">
                </dx1:ASPxLabel>
            </td>
            <td align="left">
                &nbsp;
            </td>
            <td align="left">
                <dx1:ASPxComboBox ID="cboaffiliate" runat="server" ClientInstanceName="cboaffiliate"
                    Font-Names="Tahoma" TextFormatString="{0}" Font-Size="8pt">
                    <ClientSideEvents SelectedIndexChanged="function(s, e) {txtaffiliate.SetText(cboaffiliate.GetSelectedItem().GetColumnText(1));}" />
                </dx1:ASPxComboBox>
            </td>
            <td align="left">
                &nbsp;
            </td>
            <td align="left">
                &nbsp;<dx1:ASPxTextBox ID="txtaffiliate" runat="server" Font-Names="Tahoma" Font-Size="8pt"
                    Width="170px" BackColor="Silver" ClientInstanceName="txtaffiliate" ReadOnly="True">
                    <ClientSideEvents TextChanged="function(s, e) {	lblerrmessage.SetText('');}" />
                </dx1:ASPxTextBox>
            </td>
        </tr>
        <tr>
            <td align="left">
                <dx1:ASPxLabel ID="ASPxLabel3" runat="server" Font-Names="Tahoma" Font-Size="8pt"
                    Text="DELIVERY LOCATION">
                </dx1:ASPxLabel>
            </td>
            <td align="left">
                &nbsp;
            </td>
            <td align="left">
                <dx1:ASPxComboBox ID="cboaffiliate0" runat="server" ClientInstanceName="cboaffiliate"
                    Font-Names="Tahoma" TextFormatString="{0}" Font-Size="8pt">
                    <ClientSideEvents SelectedIndexChanged="function(s, e) {txtaffiliate.SetText(cboaffiliate.GetSelectedItem().GetColumnText(1));}" />
                </dx1:ASPxComboBox>
            </td>
            <td align="left">
                &nbsp;
            </td>
            <td align="left">
                &nbsp;<dx1:ASPxTextBox ID="txtaffiliate0" runat="server" Font-Names="Tahoma" Font-Size="8pt"
                    Width="170px" BackColor="Silver" ClientInstanceName="txtaffiliate" 
                    ReadOnly="True">
                    <ClientSideEvents TextChanged="function(s, e) {	lblerrmessage.SetText('');}" />
                </dx1:ASPxTextBox>
            </td>
        </tr>
        <tr>
            <td align="left">
                <dx1:ASPxLabel ID="ASPxLabel4" runat="server" Font-Names="Tahoma" Font-Size="8pt"
                    Text="PACKING LIST NO. (INVOICE NO.)">
                </dx1:ASPxLabel>
            </td>
            <td align="left">
                &nbsp;
            </td>
            <td align="left">
                <dx1:ASPxComboBox ID="cbopart0" runat="server" ClientInstanceName="cbopart" Font-Names="Tahoma"
                    TextFormatString="{0}" Font-Size="8pt">
                    <ClientSideEvents SelectedIndexChanged="function(s, e) {txtpart.SetText(cbopart.GetSelectedItem().GetColumnText(1));}" />
                </dx1:ASPxComboBox>
            </td>
            <td align="left">
                &nbsp;
            </td>
            <td align="left">
                &nbsp;</td>
        </tr>
        <tr>
            <td align="left">
                <dx1:ASPxLabel ID="ASPxLabel5" runat="server" Font-Names="Tahoma" Font-Size="8pt"
                    Text="PACKING LIST DATE">
                </dx1:ASPxLabel>
            </td>
            <td align="left">
                <dx1:ASPxCheckBox ID="checkbox2" runat="server" CheckState="Unchecked" ClientInstanceName="checkbox2"
                    Text=" ">
                </dx1:ASPxCheckBox>
            </td>
            <td align="left">
                <dx1:ASPxDateEdit ID="dt1" runat="server" ClientInstanceName="dt1" Font-Names="Tahoma"
                    Font-Size="8pt" EditFormat="Custom" EditFormatString="dd MMM yyyy">
                </dx1:ASPxDateEdit>
            </td>
            <td align="left">
                &nbsp;</td>
            <td align="left" width="2%">
                <table style="width: 100%;">
                    <tr>
                        <td>
                            <dx1:ASPxButton ID="btnsearch" runat="server" Text="SEARCH" Width="85px" Font-Names="Tahoma"
                                Font-Size="8pt" ClientInstanceName="btnsearch" AutoPostBack="False">
                                <ClientSideEvents Click="function(s, e) {
                                            var pDateFrom = dtfrom.GetText();
	                                        var pDateTo = dtto.GetText();
                                            var pAffiliate = cboaffiliate.GetText();
                                            var pPlandate = dt1.GetText();
                                            var pSupplierDeliver = rbdeliver.GetValue();
                                            var pKanban = rbkanban.GetValue();
                                            var pPono = txtpono.GetText();
                                            var psj = txtsj.GetText();
                                            var pRemaining = rbreceiving.GetValue();
                                            var pPartcode = cbopart.GetText();
                                            
                                            
	                                        grid.PerformCallback('gridload' + '|' + pPlandate + '|' + pSupplierDeliver + '|' + pRemaining + '|' + psj + '|' + pDateFrom + '|' + pDateTo + '|' + pAffiliate + '|' + pPartcode + '|' + pPono + '|' + pKanban);
	                                        lblerrmessage.SetText('');

                                            var pMsg = s.cpMessage;
                                            if (pMsg != '') {
                                                if (pMsg.substring(1,5) == '1001' || pMsg.substring(1,5) == '1002' || pMsg.substring(1,5) == '1003') {
                                                    lblerrmessage.GetMainElement().style.color = 'Blue';  
                                                } else {
                                                    lblerrmessage.GetMainElement().style.color = 'Red';
                                                }
                                                    lblerrmessage.SetText(pMsg);
                                                } else {
                                                    lblerrmessage.SetText('');
                                             }
                                        }" />
                            </dx1:ASPxButton>
                        </td>
                        <td>
                            <dx1:ASPxButton ID="btnclear" runat="server" Text="CLEAR" Width="85px" Font-Names="Tahoma"
                                Font-Size="8pt">
                            </dx1:ASPxButton>
                        </td>
                    </tr>
                </table>
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
    <table width="100%">
        <tr>
            <td align="right" width="93%">
                &nbsp;</td>
            <td align="right">
                &nbsp;</td>
        </tr>
    </table>
    <table style="width: 100%;" width="100%">
        <tr>
            <td align="left" class="style1" colspan="3">
                <dx:ASPxGridView ID="grid" runat="server" AutoGenerateColumns="False" 
                    Width="100%" KeyFieldName="colpono;colsuppliercode;colpartno;H_KANBANORDER;H_POORDER;H_AFFILIATEORDER;HSupSJ;HPasiSJ" ClientInstanceName="grid">
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
	dtfrom.SetText(s.cpdtfrom);
    dtto.SetText(s.cpdtto); 
	dt1.SetText(s.cpdt1);
	rbdeliver.SetValue(s.cpdeliver);
	rbreceiving.SetValue(s.cpreceive);
	rbkanban.SetValue(s.cpkanban);

}" CallbackError="function(s, e) {
	e.handled = true;
}" />

                    <Columns>   
                        <dx:GridViewDataHyperLinkColumn Caption=" " FieldName="coldetail" 
                            Name="coldetail" VisibleIndex="1" Width="65px">
                            <PropertiesHyperLinkEdit TextField="coldetailname">
                            </PropertiesHyperLinkEdit>
                            <HeaderStyle Font-Names="Tahoma" Font-Size="8pt" Font-Underline="False" 
                                Wrap="True" />
                            <CellStyle Font-Names="Tahoma" Font-Size="8pt" VerticalAlign="Middle">
                            </CellStyle>
                        </dx:GridViewDataHyperLinkColumn> 
                        <dx:GridViewDataCheckColumn Caption=" " FieldName="Act" 
                            Name="Act" VisibleIndex="2" Width="30px">
                                <PropertiesCheckEdit ValueChecked="1" ValueType="System.Int32" 
                                    ValueUnchecked="0">
                                </PropertiesCheckEdit>
                        </dx:GridViewDataCheckColumn>
                        <dx:GridViewDataTextColumn Caption="NO" FieldName="colno" Name="colno" 
                            VisibleIndex="3" Width="30px">
                            <HeaderStyle Font-Names="Tahoma" Font-Size="8pt" Font-Underline="False" 
                                Wrap="True" HorizontalAlign="Center" />
                            <CellStyle Font-Names="Tahoma" Font-Size="8pt">
                            </CellStyle>
                        </dx:GridViewDataTextColumn>
                        <dx:GridViewDataTextColumn Caption="PERIOD" FieldName="colperiod" 
                            Name="colperiod" VisibleIndex="4" Width="60px">
                            <HeaderStyle Font-Names="Tahoma" Font-Size="8pt" Font-Underline="False" 
                                Wrap="True" HorizontalAlign="Center"/>
                            <CellStyle Font-Names="Tahoma" Font-Size="8pt">
                            </CellStyle>
                        </dx:GridViewDataTextColumn>
                        <dx:GridViewDataTextColumn Caption="PO NO." FieldName="colpono" Name="colpono" 
                            VisibleIndex="9" Width="80px">
                            <HeaderStyle Font-Names="Tahoma" Font-Size="8pt" Font-Underline="False" 
                                Wrap="True" HorizontalAlign="Center"/>
                            <CellStyle Font-Names="Tahoma" Font-Size="8pt">
                            </CellStyle>
                        </dx:GridViewDataTextColumn>
                        <dx:GridViewDataTextColumn Caption="PART NO." FieldName="colpartno" 
                            Name="colpartno" VisibleIndex="19" Width="90px">
                            <HeaderStyle Font-Names="Tahoma" Font-Size="8pt" Font-Underline="False" 
                                Wrap="True" HorizontalAlign="Center"/>
                            <CellStyle Font-Names="Tahoma" Font-Size="8pt">
                            </CellStyle>
                        </dx:GridViewDataTextColumn>
                        <dx:GridViewDataTextColumn Caption="PART NAME" 
                            FieldName="colpartname" Name="colpartname" VisibleIndex="20" Width="140px">
                            <HeaderStyle Font-Names="Tahoma" Font-Size="8pt" Font-Underline="False" 
                                Wrap="True" HorizontalAlign="Center"/>
                            <CellStyle Font-Names="Tahoma" Font-Size="8pt">
                            </CellStyle>
                        </dx:GridViewDataTextColumn>
                        <dx:GridViewDataTextColumn Caption="UOM" FieldName="coluom" Name="coluom" 
                            VisibleIndex="22" Width="40px">
                            <HeaderStyle Font-Names="Tahoma" Font-Size="8pt" Font-Underline="False" 
                                Wrap="True" HorizontalAlign="Center"/>
                            <CellStyle Font-Names="Tahoma" Font-Size="8pt" HorizontalAlign="Right">
                            </CellStyle>
                        </dx:GridViewDataTextColumn>
                        <dx:GridViewDataTextColumn Caption="QTY/BOX" FieldName="coldeliveryqty" 
                            Name="coldeliveryqty" VisibleIndex="23" Width="70px">
                            <PropertiesTextEdit DisplayFormatString="{0:n0}">
                            </PropertiesTextEdit>
                            <HeaderStyle Font-Names="Tahoma" Font-Size="8pt" Font-Underline="False" 
                                Wrap="True" HorizontalAlign="Center"/>
                            <CellStyle Font-Names="Tahoma" Font-Size="8pt" HorizontalAlign="Right">
                            </CellStyle>
                        </dx:GridViewDataTextColumn>
                        <dx:GridViewDataTextColumn Caption="GOOD RECEIVEING QTY" FieldName="colreceiveqty" 
                            Name="colreceiveqty" VisibleIndex="24" Width="70px">
                            <HeaderStyle Font-Names="Tahoma" Font-Size="8pt" Font-Underline="False" 
                                Wrap="True" HorizontalAlign="Center"/>
                            <CellStyle Font-Names="Tahoma" Font-Size="8pt" HorizontalAlign="Right">
                            </CellStyle>
                        </dx:GridViewDataTextColumn>
                        <dx:GridViewDataTextColumn Caption="PREVIOUS PACKING LIST QTY" FieldName="coldefect" 
                            Name="coldefect" VisibleIndex="25" Width="70px">
                            <HeaderStyle Font-Names="Tahoma" Font-Size="8pt" Font-Underline="False" 
                                Wrap="True" HorizontalAlign="Center"/>
                            <CellStyle Font-Names="Tahoma" Font-Size="8pt" HorizontalAlign="Right">
                            </CellStyle>
                        </dx:GridViewDataTextColumn>
                        <dx:GridViewDataTextColumn Caption="PACKING LIST QTY" 
                            FieldName="colremaining" Name="colremaining" VisibleIndex="26" 
                            Width="70px">
                            <HeaderStyle Font-Names="Tahoma" Font-Size="8pt" Font-Underline="False" 
                                Wrap="True" HorizontalAlign="Center"/>
                            <CellStyle Font-Names="Tahoma" Font-Size="8pt" HorizontalAlign="Right">
                            </CellStyle>
                        </dx:GridViewDataTextColumn>
                        <dx:GridViewDataTextColumn Caption="BOX QTY" FieldName="colpasideliveryqty" 
                            Name="colpasideliveryqty" VisibleIndex="27" Width="80px">
                            <HeaderStyle Font-Names="Tahoma" Font-Size="8pt" Font-Underline="False" 
                                Wrap="True" HorizontalAlign="Center"/>
                            <CellStyle Font-Names="Tahoma" Font-Size="8pt" HorizontalAlign="Right">
                            </CellStyle>
                        </dx:GridViewDataTextColumn>
                        <dx:GridViewDataTextColumn Caption="LABEL NO." FieldName="coldeliverydate" 
                            Name="coldeliverydate" VisibleIndex="28" Width="80px">
                            <HeaderStyle Font-Names="Tahoma" Font-Size="8pt" Font-Underline="False" 
                                Wrap="True" HorizontalAlign="Center"/>
                            <CellStyle Font-Names="Tahoma" Font-Size="8pt">
                            </CellStyle>
                        </dx:GridViewDataTextColumn>
                        <dx:GridViewDataTextColumn Caption="NETT WEIGHT (KGS)" FieldName="coldeliveryby" 
                            Name="coldeliveryby" VisibleIndex="29" Width="100px">
                            <HeaderStyle Font-Names="Tahoma" Font-Size="8pt" Font-Underline="False" 
                                Wrap="True" HorizontalAlign="Center"/>
                            <CellStyle Font-Names="Tahoma" Font-Size="8pt">
                            </CellStyle>
                        </dx:GridViewDataTextColumn>   
                        <dx:GridViewDataTextColumn Caption="HS CODE" VisibleIndex="21">
                        </dx:GridViewDataTextColumn>
                        <dx:GridViewDataTextColumn Caption="SUPPLIER CODE" VisibleIndex="30">
                        </dx:GridViewDataTextColumn>
                        <dx:GridViewDataTextColumn Caption="SUPPLIER NAME" VisibleIndex="31">
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
                        <SelectedRow ForeColor="Black">
                        </SelectedRow>
                    </Styles>
                    <StylesEditors ButtonEditCellSpacing="0">
                        <ProgressBar Height="21px">
                        </ProgressBar>
                    </StylesEditors>
                </dx:ASPxGridView>
                <br />
            </td>
        </tr>
        <tr>
            <td align="left">
                <dx1:ASPxButton ID="btnsubmenu" runat="server" Text="SUB MENU" Width="90px" Font-Names="Tahoma"
                    Font-Size="8pt">
                </dx1:ASPxButton>
            </td>
            <td align="right">
                <dx1:ASPxButton ID="btnPrint" runat="server" Text="PRINT" Width="90px" Font-Names="Tahoma"
                    Font-Size="8pt">
                </dx1:ASPxButton>
            </td>
            <td align="right" width="50px">
                <dx1:ASPxButton ID="btnPrint0" runat="server" Text="PRINT DETAIL" Width="90px" Font-Names="Tahoma"
                    Font-Size="8pt">
                </dx1:ASPxButton>
            </td>
        </tr>
    </table>
    <dx:ASPxGlobalEvents ID="ge" runat="server">
        <ClientSideEvents ControlsInitialized="function(s, e) {
	        OnControlsInitializedSplitter();
	        OnControlsInitializedGrid();
        }" />
    </dx:ASPxGlobalEvents>
    <dx:ASPxHiddenField ID="HF" runat="server" ClientInstanceName="HF">
    </dx:ASPxHiddenField>
</asp:Content>
