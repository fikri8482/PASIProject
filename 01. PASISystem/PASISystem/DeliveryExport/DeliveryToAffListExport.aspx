<%@ Page Title="" Language="vb" AutoEventWireup="false" MasterPageFile="~/SiteContent.Master"CodeBehind="DeliveryToAffListExport.aspx.vb" Inherits="PASISystem.DeliveryToAffListExport" %>

<%@ Register Assembly="DevExpress.Web.v14.1, Version=14.1.6.0, Culture=neutral, PublicKeyToken=b88d1754d700e49a"
    Namespace="DevExpress.Web.ASPxEditors" TagPrefix="dx1" %>
<%@ Register Assembly="DevExpress.Web.v14.1, Version=14.1.6.0, Culture=neutral, PublicKeyToken=b88d1754d700e49a"
    Namespace="DevExpress.Web.ASPxGridView" TagPrefix="dx" %>
<%@ Register Assembly="DevExpress.Web.v14.1, Version=14.1.6.0, Culture=neutral, PublicKeyToken=b88d1754d700e49a"
    Namespace="DevExpress.Web.ASPxGridView.Export" TagPrefix="dx2" %>
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
        
        .style2
        {
            width: 564px;
        }
        .style3
        {
            width: 564px;
            height: 81px;
        }
        .style4
        {
            height: 81px;
        }
        .style5
        {
            width: 564px;
            height: 22px;
        }
        .style6
        {
            height: 22px;
        }
    </style>
</asp:Content>
<asp:Content ID="Content2" ContentPlaceHolderID="Content" runat="server">
    <dx:ASPxGlobalEvents ID="ge" runat="server">
        <ClientSideEvents ControlsInitialized="function(s, e) {
	OnControlsInitializedSplitter();
	OnControlsInitializedGrid();
}" />
    </dx:ASPxGlobalEvents>
    <dx:ASPxHiddenField ID="HF" runat="server" ClientInstanceName="HF">
    </dx:ASPxHiddenField>

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
            <td align="left" class="style3">
                <dx1:ASPxLabel ID="ASPxLabel1" runat="server" Font-Names="Tahoma" Font-Size="8pt"
                    Text="SUPPLIER PLAN DELIVERY DATE (UNTIL)">
                </dx1:ASPxLabel>
            </td>
            <td align="left" class="style4">
                <dx1:ASPxCheckBox ID="checkbox1" runat="server" CheckState="Unchecked" ClientInstanceName="checkbox1"
                    Text=" ">
                </dx1:ASPxCheckBox>
            </td>
            <td align="left" width="180px" class="style4">
                <dx1:ASPxDateEdit ID="dt1" runat="server" ClientInstanceName="dt1" Font-Names="Tahoma"
                    Font-Size="8pt" EditFormat="Custom" EditFormatString="dd MMM yyyy">
                </dx1:ASPxDateEdit>
            </td>
            <td align="left" width="10px" class="style4">
            </td>
            <td align="left" class="style4" width="50px">
            </td>
            <td align="left" class="style4" width="207px">
                <dx1:ASPxLabel ID="ASPxLabel8" runat="server" Font-Names="Tahoma" Font-Size="8pt"
                    Text="PART CODE/NAME">
                </dx1:ASPxLabel>
            </td>
            <td align="left" class="style4">
                <dx1:ASPxTextBox ID="txtstatus" runat="server" ClientInstanceName="txtstatus" 
                    Width="170px" Visible="False">
                </dx1:ASPxTextBox>
                <dx1:ASPxComboBox ID="cbopart" runat="server" ClientInstanceName="cbopart" Font-Names="Tahoma"
                    TextFormatString="{0}" Font-Size="8pt">
                    <ClientSideEvents SelectedIndexChanged="function(s, e) {
	                                            txtpart.SetText(cbopart.GetSelectedItem().GetColumnText(1));
                                                }" />
                </dx1:ASPxComboBox>
            </td>
            <td align="left" class="style4">
                <dx1:ASPxTextBox ID="txtpart" runat="server" Font-Names="Tahoma" Font-Size="8pt"
                    Width="170px" BackColor="Silver" ClientInstanceName="txtpart" ReadOnly="True">
                    <ClientSideEvents TextChanged="function(s, e) {
	lblerrmessage.SetText('');
}" />
                </dx1:ASPxTextBox>
            </td>
            <td align="left" rowspan="5" xml:lang="150px">
                <dx:ASPxRoundPanel ID="ASPxRoundPanel1" runat="server" HeaderText="STATUS" Height="80px"
                    ShowCollapseButton="true" View="GroupBox" Width="100%">
                    <PanelCollection>
                        <dx:PanelContent ID="PanelContent1" runat="server">
                            <table style="width: 100%;">
                                <tr>
                                    <td>
                                        <dx1:ASPxLabel ID="ASPxLabel13" runat="server" Font-Names="Tahoma" Font-Size="8pt"
                                            Text="(7). DELIVERY BY SUPPLIER" Width="145px">
                                        </dx1:ASPxLabel>
                                    </td>
                                </tr>
                                <tr>
                                    <td>
                                        <dx1:ASPxLabel ID="ASPxLabel14" runat="server" Font-Names="Tahoma" Font-Size="8pt"
                                            Text="(8). RECEIVE BY FORWARDER" Width="200px">
                                        </dx1:ASPxLabel>
                                    </td>
                                </tr>
                                <tr>
                                    <td>
                                        <dx1:ASPxLabel ID="ASPxLabel15" runat="server" Font-Names="Tahoma" Font-Size="8pt"
                                            Text="(9). SHIPPING INSTRUCTION" Width="145px">
                                        </dx1:ASPxLabel>
                                    </td>
                                </tr>
                                  <tr>
                                    <td>
                                        <dx1:ASPxLabel ID="ASPxLabel18" runat="server" Font-Names="Tahoma" Font-Size="8pt"
                                            Text="(10). RECEIVE TALLY DATA" Width="200px">
                                        </dx1:ASPxLabel>
                                    </td>
                                </tr>
                            </table>
                        </dx:PanelContent>
                    </PanelCollection>
                </dx:ASPxRoundPanel>
                <br />
                <br />
                <br />
                <br />
                <br />
                <br />
                <br />
                <br />
                <br />
                <br />
                <br />
            </td>
        </tr>
        <tr>
            <td align="left" class="style5">
                <dx1:ASPxLabel ID="ASPxLabel2" runat="server" Font-Names="Tahoma" Font-Size="8pt"
                    Text="SUPPLIER ALREADY DELIVER">
                </dx1:ASPxLabel>
            </td>
            <td align="left" class="style6">
                &nbsp;
            </td>
            <td align="left" width="180px" class="style6">
                <dx1:ASPxRadioButtonList ID="rbdeliver" runat="server" ClientInstanceName="rbdeliver"
                    Font-Names="Tahoma" Font-Size="8pt" Height="16px" RepeatDirection="Horizontal">
                    <Items>
                        <dx1:ListEditItem Text="ALL" Value="ALL" Selected="True" />
                        <dx1:ListEditItem Text="YES" Value="YES" />
                        <dx1:ListEditItem Text="NO" Value="NO" />
                    </Items>
                    <Border BorderStyle="None" />
                </dx1:ASPxRadioButtonList>
            </td>
            <td align="left" width="10px" class="style6">
                &nbsp;
            </td>
            <td align="left" width="50px" class="style6">
                &nbsp;
            </td>
            <td align="left" width="207px" class="style6">
                <dx1:ASPxLabel ID="ASPxLabel12" runat="server" Font-Names="Tahoma" Font-Size="8pt"
                    Text="AFFILIATE CODE/NAME">
                </dx1:ASPxLabel>
            </td>
            <td align="left" class="style6">
                <dx1:ASPxComboBox ID="cboaffiliate" runat="server" ClientInstanceName="cboaffiliate"
                    Font-Names="Tahoma" TextFormatString="{0}" Font-Size="8pt">
                    <ClientSideEvents SelectedIndexChanged="function(s, e) {
	                                            txtaffiliate.SetText(cboaffiliate.GetSelectedItem().GetColumnText(1));
                                                }" />
                </dx1:ASPxComboBox>
            </td>
            <td align="left" class="style6">
                <dx1:ASPxTextBox ID="txtaffiliate" runat="server" Font-Names="Tahoma" Font-Size="8pt"
                    Width="170px" BackColor="Silver" ClientInstanceName="txtaffiliate" ReadOnly="True">
                    <ClientSideEvents TextChanged="function(s, e) {
	lblerrmessage.SetText('');
}" />
                </dx1:ASPxTextBox>
            </td>
        </tr>
        <tr>
            <td align="left" class="style2">
                <dx1:ASPxLabel ID="ASPxLabel3" runat="server" Font-Names="Tahoma" Font-Size="8pt"
                    Text="REMAINING RECEIVING QTY">
                </dx1:ASPxLabel>
            </td>
            <td align="left">
                &nbsp;
            </td>
            <td align="left" width="180px">
                <dx1:ASPxRadioButtonList ID="rbreceiving" runat="server" ClientInstanceName="rbreceiving"
                    Font-Names="Tahoma" Font-Size="8pt" RepeatDirection="Horizontal">
                    <Items>
                        <dx1:ListEditItem Text="ALL" Value="ALL" Selected="True" />
                        <dx1:ListEditItem Text="YES" Value="YES" />
                        <dx1:ListEditItem Text="NO" Value="NO" />
                    </Items>
                    <Border BorderStyle="None" />
                </dx1:ASPxRadioButtonList>
            </td>
            <td align="left" width="10px">
                &nbsp;
            </td>
            <td align="left" width="50px">
                &nbsp;
            </td>
            <td align="left" width="207px">
                <dx1:ASPxLabel ID="ASPxLabel9" runat="server" Font-Names="Tahoma" Font-Size="8pt"
                    Text="ORDER NO.">
                </dx1:ASPxLabel>
            </td>
            <td align="left">
                <dx1:ASPxTextBox ID="txtorderno" runat="server" Font-Names="Tahoma" Font-Size="8pt"
                    Width="170px" ClientInstanceName="txtorderno">
                </dx1:ASPxTextBox>
            </td>
            <td align="left">
                &nbsp;
                </td>
        </tr>
        <tr>
            <td align="left" class="style2">
                <dx1:ASPxLabel ID="ASPxLabel16" runat="server" Font-Names="Tahoma" Font-Size="8pt"
                    Text="DIFFERENT DELIVERY AND RECEIVING QTY">
                </dx1:ASPxLabel>
            </td>
            <td align="left">
                &nbsp;
            </td>
            <td align="left" width="180px">
                <dx1:ASPxRadioButtonList ID="rbdiff" runat="server" ClientInstanceName="rbdiff" Font-Names="Tahoma"
                    Font-Size="8pt" RepeatDirection="Horizontal">
                    <Items>
                        <dx1:ListEditItem Text="ALL" Value="ALL" Selected="True" />
                        <dx1:ListEditItem Text="YES" Value="YES" />
                        <dx1:ListEditItem Text="NO" Value="NO" />
                    </Items>
                    <Border BorderStyle="None" />
                </dx1:ASPxRadioButtonList>
            </td>
            <td align="left" width="10px">
                &nbsp;
            </td>
            <td align="left" width="50px">
                &nbsp;
            </td>
            <td align="left" width="207px">
                <dx1:ASPxLabel ID="ASPxLabel7" runat="server" Font-Names="Tahoma" Font-Size="8pt"
                    Text="SUPPLIER CODE/NAME">
                </dx1:ASPxLabel>
            </td>
            <td align="left">
                <dx1:ASPxComboBox ID="cbosupplier" runat="server" ClientInstanceName="cbosupplier"
                    Font-Names="Tahoma" TextFormatString="{0}" Font-Size="8pt">
                    <ClientSideEvents SelectedIndexChanged="function(s, e) {
	                                            txtsupplier.SetText(cbosupplier.GetSelectedItem().GetColumnText(1));
                                                }" />
                </dx1:ASPxComboBox>
            </td>
            <td align="left">
                <dx1:ASPxTextBox ID="txtsupplier" runat="server" Font-Names="Tahoma" Font-Size="8pt"
                    Width="170px" BackColor="Silver" ClientInstanceName="txtsupplier" ReadOnly="True">
                    <ClientSideEvents TextChanged="function(s, e) {
	lblerrmessage.SetText('');
}" />
                </dx1:ASPxTextBox>
            </td>
        </tr>
        <tr>
            <td align="left" class="style2">
                <dx1:ASPxLabel ID="ASPxLabel17" runat="server" Font-Names="Tahoma" Font-Size="8pt"
                    Text="ALREADY SHIPPING">
                </dx1:ASPxLabel>
            </td>
            <td align="left">
                &nbsp;
            </td>
            <td align="left" width="180px">
                <dx1:ASPxRadioButtonList ID="rbshipping" runat="server" ClientInstanceName="rbshipping"
                    Font-Names="Tahoma" Font-Size="8pt" RepeatDirection="Horizontal">
                    <Items>
                        <dx1:ListEditItem Text="ALL" Value="ALL" />
                        <dx1:ListEditItem Text="YES" Value="YES" />
                        <dx1:ListEditItem Text="NO" Value="NO" Selected="True" />
                    </Items>
                    <Border BorderStyle="None" />
                </dx1:ASPxRadioButtonList>
            </td>
            <td align="left" width="10px">
                &nbsp;
            </td>
            <td align="left" width="50px">
                &nbsp;
            </td>
            <td align="left" width="207px">
                &nbsp;
            </td>
            <td align="left">
                &nbsp;
                </td>
            <td align="left">
                &nbsp;
                </td>
        </tr>
        <tr>
            <td align="left" class="style2">
                <dx1:ASPxLabel ID="ASPxLabel5" runat="server" Font-Names="Tahoma" Font-Size="8pt"
                    Text="RECEIVED DATE">
                </dx1:ASPxLabel>
            </td>
            <td align="left">
                <dx1:ASPxCheckBox ID="checkbox2" runat="server" CheckState="Unchecked" ClientInstanceName="checkbox2"
                    Text=" ">
                </dx1:ASPxCheckBox>
            </td>
            <td align="left" width="180px">
                <dx1:ASPxDateEdit ID="dtfrom" runat="server" ClientInstanceName="dtfrom" Font-Size="8pt"
                    EditFormat="Custom" EditFormatString="dd MMM yyyy">
                </dx1:ASPxDateEdit>
            </td>
            <td align="left" width="10px">
                <dx1:ASPxLabel ID="ASPxLabel6" runat="server" Font-Names="Tahoma" Font-Size="8pt"
                    Text="~">
                </dx1:ASPxLabel>
            </td>
            <td align="left" width="50px" colspan="2" style="width: 414px">
                <dx1:ASPxDateEdit ID="dtto" runat="server" ClientInstanceName="dtto" Font-Size="8pt"
                    EditFormat="Custom" EditFormatString="dd MMM yyyy">
                </dx1:ASPxDateEdit>
            </td>
            <td align="left">
                &nbsp;
                <dx1:ASPxButton ID="btndeliver" runat="server" Text="UPLOAD" Width="90px" Font-Names="Tahoma"
                    Font-Size="8pt" AutoPostBack="False" Visible="False">
                </dx1:ASPxButton>
            </td>
            <td align="left">
                &nbsp;
                </td>
            <td align="left">
                <table style="width: 100%;">
                    <tr>
                        <td width="90px">
                            <dx1:ASPxButton ID="btnsearch" runat="server" Text="SEARCH" Width="85px" Font-Names="Tahoma"
                                Font-Size="8pt" ClientInstanceName="btnsearch" AutoPostBack="False">
                                <ClientSideEvents Click="function(s, e) {
                                            var pDateFrom = dtfrom.GetText();
	                                        var pDateTo = dtto.GetText();
                                            var pSupplier = cbosupplier.GetText();
                                            var pPlandate = dt1.GetText();
                                            var pOrderNo = txtorderno.GetText();
                                            var psj = ''
                                            var pRemaining = rbreceiving.GetValue();
                                            var pPartcode = cbopart.GetText();
                                         
                                            
	                                        grid.PerformCallback('gridload' + '|' + pPlandate + '|' + pSupplier + '|' + pRemaining + '|' + psj + '|' + pDateFrom + '|' + pDateTo + '|' + pPartcode + '|' + pOrderNo);

	                                        lblerrmessage.SetText('');

                                            var pMsg = s.cpMessage;
                                             if (pMsg) {
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
                <dx1:ASPxTextBox ID="txtsupplier8" runat="server" Width="20px" Font-Names="Verdana"
                    Font-Size="8pt" BackColor="#FF66CC" Height="16px">
                    <ClientSideEvents TextChanged="function(s, e) {
	                    lblerrmessage.SetText('');
                    }" />
                </dx1:ASPxTextBox>
            </td>
            <td align="right">
                <dx1:ASPxLabel ID="ASPxLabel11" runat="server" Font-Names="Tahoma" Font-Size="8pt"
                    Text=": DIFFERENCE">
                </dx1:ASPxLabel>
            </td>
        </tr>
    </table>

    <br />
    <table style="width:100%;">
        <tr>
            <td colspan="3">
                <dx:ASPxGridView ID="grid" runat="server" Width="100%"
                    KeyFieldName="H_ORDERNO;H_SJRECEIVING;H_SJ;H_PARTNO;URUT;colgood;coldefect;LabelNo;colno" 
                    ClientInstanceName="grid" AutoGenerateColumns="False">
                    <ClientSideEvents EndCallback="function(s, e) {
						var pMsg = s.cpMessage;
                         if (pMsg) {
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
	rbreceiving.SetValue(s.cpreceive);

}" CallbackError="function(s, e) {
	e.handled = true;
}" />
                    <Columns>
                        <dx:GridViewDataHyperLinkColumn Caption=" " FieldName="coldetail" Name="coldetail"
                            VisibleIndex="3" Width="65px">
                            <PropertiesHyperLinkEdit TextField="coldetailname">
                            </PropertiesHyperLinkEdit>
                            <HeaderStyle Font-Names="Tahoma" Font-Size="8pt" Font-Underline="False" Wrap="True" />
                            <CellStyle Font-Names="Tahoma" Font-Size="8pt" VerticalAlign="Middle">
                            </CellStyle>
                        </dx:GridViewDataHyperLinkColumn>
                        <dx:GridViewDataTextColumn Caption="NO" FieldName="colno" Name="colno" VisibleIndex="4"
                            Width="30px">
                            <HeaderStyle Font-Names="Tahoma" Font-Size="8pt" Font-Underline="False" Wrap="True"
                                HorizontalAlign="Center" />
                            <CellStyle Font-Names="Tahoma" Font-Size="8pt">
                            </CellStyle>
                        </dx:GridViewDataTextColumn>
                        <dx:GridViewDataTextColumn Caption="PERIOD" FieldName="colperiod" Name="colperiod"
                            VisibleIndex="6" Width="60px">
                            <HeaderStyle Font-Names="Tahoma" Font-Size="8pt" Font-Underline="False" Wrap="True"
                                HorizontalAlign="Center" />
                            <CellStyle Font-Names="Tahoma" Font-Size="8pt">
                            </CellStyle>
                        </dx:GridViewDataTextColumn>
                        <dx:GridViewDataTextColumn Caption="AFFILIATE CODE" FieldName="colaffiliatecode"
                            Name="colaffiliatecode" VisibleIndex="7" Width="70px">
                            <HeaderStyle Font-Names="Tahoma" Font-Size="8pt" Font-Underline="False" Wrap="True"
                                HorizontalAlign="Center" />
                            <CellStyle Font-Names="Tahoma" Font-Size="8pt">
                            </CellStyle>
                        </dx:GridViewDataTextColumn>
                        <dx:GridViewDataTextColumn Caption="AFFILIATE NAME" FieldName="colaffiliatename"
                            Name="colaffiliatename" VisibleIndex="8" Width="0px">
                            <HeaderStyle Font-Names="Tahoma" Font-Size="8pt" Font-Underline="False" Wrap="True"
                                HorizontalAlign="Center" />
                            <CellStyle Font-Names="Tahoma" Font-Size="8pt">
                            </CellStyle>
                        </dx:GridViewDataTextColumn>
                        <dx:GridViewDataTextColumn Caption="DELIVERY LOCATION CODE" FieldName="coldeliverylocationcode"
                            Name="coldeliverylocationcode" VisibleIndex="10" Width="120px">
                            <HeaderStyle Font-Names="Tahoma" Font-Size="8pt" Font-Underline="False" Wrap="True"
                                HorizontalAlign="Center" />
                            <CellStyle Font-Names="Tahoma" Font-Size="8pt">
                            </CellStyle>
                        </dx:GridViewDataTextColumn>
                        <dx:GridViewDataTextColumn Caption="DELIVERY LOCATION NAME" FieldName="coldeliverylocationname"
                            Name="coldeliverylocationname" VisibleIndex="11" Width="0px">
                            <HeaderStyle Font-Names="Tahoma" Font-Size="8pt" Font-Underline="False" Wrap="True"
                                HorizontalAlign="Center" />
                            <CellStyle Font-Names="Tahoma" Font-Size="8pt">
                            </CellStyle>
                        </dx:GridViewDataTextColumn>
                        <dx:GridViewDataTextColumn Caption="ORDER NO." FieldName="colorderno" Name="colorderno"
                            VisibleIndex="12" Width="80px">
                            <HeaderStyle Font-Names="Tahoma" Font-Size="8pt" Font-Underline="False" Wrap="True"
                                HorizontalAlign="Center" />
                            <CellStyle Font-Names="Tahoma" Font-Size="8pt">
                            </CellStyle>
                        </dx:GridViewDataTextColumn>
                        <dx:GridViewDataTextColumn Caption="SUPPLIER CODE" FieldName="colsuppliercode" Name="colsuppliercode"
                            VisibleIndex="9" Width="70px">
                            <HeaderStyle Font-Names="Tahoma" Font-Size="8pt" Font-Underline="False" Wrap="True"
                                HorizontalAlign="Center" />
                            <CellStyle Font-Names="Tahoma" Font-Size="8pt">
                            </CellStyle>
                        </dx:GridViewDataTextColumn>
                        <dx:GridViewDataTextColumn Caption="SUPPLIER NAME" FieldName="colsuppliername" Name="colsuppliername"
                            VisibleIndex="13" Width="0px">
                            <HeaderStyle Font-Names="Tahoma" Font-Size="8pt" Font-Underline="False" Wrap="True"
                                HorizontalAlign="Center" />
                            <CellStyle Font-Names="Tahoma" Font-Size="8pt" HorizontalAlign="Center">
                            </CellStyle>
                        </dx:GridViewDataTextColumn>
                        <dx:GridViewDataTextColumn Caption="SUPPLIER PLAN DELIVERY DATE" FieldName="colplandeldate"
                            Name="colplandeldate" VisibleIndex="14" Width="0px">
                            <HeaderStyle Font-Names="Tahoma" Font-Size="8pt" Font-Underline="False" Wrap="True"
                                HorizontalAlign="Center" />
                            <CellStyle Font-Names="Tahoma" Font-Size="8pt">
                            </CellStyle>
                        </dx:GridViewDataTextColumn>
                        <dx:GridViewDataTextColumn Caption="RECEIVE DATE" FieldName="coldeldate"
                            Name="coldeldate" VisibleIndex="15" Width="80px">
                            <HeaderStyle Font-Names="Tahoma" Font-Size="8pt" Font-Underline="False" Wrap="True"
                                HorizontalAlign="Center" />
                            <CellStyle Font-Names="Tahoma" Font-Size="8pt">
                            </CellStyle>
                        </dx:GridViewDataTextColumn>
                        <dx:GridViewDataTextColumn Caption="SURAT JALAN NO." FieldName="colsj" Name="colsj"
                            VisibleIndex="16" Width="180px">
                            <HeaderStyle Font-Names="Tahoma" Font-Size="8pt" Font-Underline="False" Wrap="True"
                                HorizontalAlign="Center" />
                            <CellStyle Font-Names="Tahoma" Font-Size="8pt">
                            </CellStyle>
                        </dx:GridViewDataTextColumn>
                        <dx:GridViewDataTextColumn Caption="PART NO." FieldName="colpartno" Name="colpartno"
                            VisibleIndex="17" Width="0px">
                            <HeaderStyle Font-Names="Tahoma" Font-Size="8pt" Font-Underline="False" Wrap="True"
                                HorizontalAlign="Center" />
                            <CellStyle Font-Names="Tahoma" Font-Size="8pt" HorizontalAlign="Left">
                            </CellStyle>
                            <HeaderStyle HorizontalAlign="Center" Wrap="True" Font-Names="Tahoma" Font-Size="8pt"
                                Font-Underline="False"></HeaderStyle>
                            <CellStyle Font-Names="Tahoma" Font-Size="8pt">
                            </CellStyle>
                        </dx:GridViewDataTextColumn>
                        <dx:GridViewDataTextColumn Caption="PART NAME" FieldName="colpartname" Name="colpartname"
                            VisibleIndex="18" Width="0px">
                            <HeaderStyle Font-Names="Tahoma" Font-Size="8pt" Font-Underline="False" Wrap="True"
                                HorizontalAlign="Center" />
                            <CellStyle Font-Names="Tahoma" Font-Size="8pt" HorizontalAlign="Left">
                            </CellStyle>
                        </dx:GridViewDataTextColumn>
                        <dx:GridViewDataTextColumn Caption="UOM" FieldName="coluom" Name="coluom" VisibleIndex="20"
                            Width="0px">
                            <HeaderStyle Font-Names="Tahoma" Font-Size="8pt" Font-Underline="False" Wrap="True"
                                HorizontalAlign="Center" />
                            <CellStyle Font-Names="Tahoma" Font-Size="8pt" HorizontalAlign="Right">
                            </CellStyle>
                        </dx:GridViewDataTextColumn>
                        <dx:GridViewDataTextColumn Caption="SUPPLIER DELIVERY QTY" FieldName="coldeliveryqty"
                            Name="coldeliveryqty" VisibleIndex="21" Width="0px">
                            <PropertiesTextEdit DisplayFormatString="{0:n0}">
                            </PropertiesTextEdit>
                            <HeaderStyle Font-Names="Tahoma" Font-Size="8pt" Font-Underline="False" Wrap="True"
                                HorizontalAlign="Center" />
                            <CellStyle Font-Names="Tahoma" Font-Size="8pt" HorizontalAlign="Right">
                            </CellStyle>
                        </dx:GridViewDataTextColumn>
                        <dx:GridViewDataTextColumn Caption="FWD GOOD RECEIVEING QTY" 
                            FieldName="colgood" Name="colgood"
                            VisibleIndex="22" Width="0px">
                            <HeaderStyle Font-Names="Tahoma" Font-Size="8pt" Font-Underline="False" Wrap="True"
                                HorizontalAlign="Center" />
                            <CellStyle Font-Names="Tahoma" Font-Size="8pt">
                            </CellStyle>
                        </dx:GridViewDataTextColumn>
                        <dx:GridViewDataTextColumn Caption="FWD DEFECT RECEIVEING QTY" FieldName="coldefect"
                            Name="coldefect" VisibleIndex="23" Width="0px">
                            <HeaderStyle Font-Names="Tahoma" Font-Size="8pt" Font-Underline="False" Wrap="True"
                                HorizontalAlign="Center" />
                            <CellStyle Font-Names="Tahoma" Font-Size="8pt" HorizontalAlign="Right">
                            </CellStyle>
                        </dx:GridViewDataTextColumn>
                        <dx:GridViewDataTextColumn Caption="REMAINING RECEIVE QTY" FieldName="colremaining"
                            Name="colremaining" VisibleIndex="24" Width="0px">
                            <HeaderStyle Font-Names="Tahoma" Font-Size="8pt" Font-Underline="False" Wrap="True"
                                HorizontalAlign="Center" />
                            <CellStyle Font-Names="Tahoma" Font-Size="8pt" HorizontalAlign="Right">
                            </CellStyle>
                        </dx:GridViewDataTextColumn>
                        <dx:GridViewDataTextColumn Caption="RECEIVED DATE" FieldName="colreceivedate" Name="colreceivedate"
                            VisibleIndex="25" Width="0px">
                        </dx:GridViewDataTextColumn>
                        <dx:GridViewDataTextColumn FieldName="H_SJRECEIVING" VisibleIndex="27" Width="0px">
                        </dx:GridViewDataTextColumn>
                        <dx:GridViewDataTextColumn FieldName="H_SJ" VisibleIndex="28" Width="0px">
                        </dx:GridViewDataTextColumn>
                        <dx:GridViewDataTextColumn FieldName="H_ORDERNO" VisibleIndex="29" Width="0px">
                        </dx:GridViewDataTextColumn>
                        <dx:GridViewDataTextColumn FieldName="H_PARTNO" VisibleIndex="30" Width="0px">
                        </dx:GridViewDataTextColumn>
                        <dx:GridViewDataTextColumn FieldName="URUT" VisibleIndex="31" Width="0px">
                        </dx:GridViewDataTextColumn>
                        <dx:GridViewBandColumn Caption="STATUS" VisibleIndex="26">
                            <Columns>
                                <dx:GridViewDataTextColumn Caption="(7)" FieldName="S7" VisibleIndex="0" Width="60px">
                                    <CellStyle HorizontalAlign="Center" VerticalAlign="Middle" Wrap="True">
                                    </CellStyle>
                                </dx:GridViewDataTextColumn>
                                <dx:GridViewDataTextColumn Caption="(8)" FieldName="S8" VisibleIndex="1" Width="60px">
                                    <CellStyle Wrap="True">
                                    </CellStyle>
                                </dx:GridViewDataTextColumn>
                                <dx:GridViewDataTextColumn Caption="(9)" FieldName="S9" VisibleIndex="2" Width="60px">
                                    <CellStyle Wrap="True">
                                    </CellStyle>
                                </dx:GridViewDataTextColumn>
                                 <dx:GridViewDataTextColumn Caption="(10)" FieldName="S10" VisibleIndex="3" Width="60px">
                                    <CellStyle Wrap="True">
                                    </CellStyle>
                                </dx:GridViewDataTextColumn>
                            </Columns>
                            <HeaderStyle HorizontalAlign="Center" VerticalAlign="Middle" />
                        </dx:GridViewBandColumn>
                        <dx:GridViewDataCheckColumn Caption="  " FieldName="ACT" Name="ACT" 
                            VisibleIndex="5" Width="30px" >
                            <PropertiesCheckEdit ValueChecked="1" ValueType="System.Int32" ValueUnchecked="0">
                            </PropertiesCheckEdit>
                        </dx:GridViewDataCheckColumn>
                        <dx:GridViewDataTextColumn FieldName="FWDID" VisibleIndex="32" Width="0px">
                        </dx:GridViewDataTextColumn>
                        <dx:GridViewDataTextColumn Caption="BOX NO." FieldName="LabelNo" 
                            VisibleIndex="19" Width="0px">
                            <HeaderStyle HorizontalAlign="Center" />
                        </dx:GridViewDataTextColumn>
                        <dx:GridViewDataHyperLinkColumn FieldName="coldetail2" VisibleIndex="2" Caption=" " 
                            Width="110px">
                            <PropertiesHyperLinkEdit TextField="coldetailname2">
                            </PropertiesHyperLinkEdit>
                        </dx:GridViewDataHyperLinkColumn>
                    </Columns>
                    <SettingsPager Visible="False" Position="Top">
                        <Summary Text="Page {0} of {1} [{2} record(s)]" AllPagesText="Page {0} of {1} " />
                        <Summary AllPagesText="Page {0} of {1} " Text="Page {0} of {1} [{2} record(s)]">
                        </Summary>
                    </SettingsPager>
                    <SettingsEditing Mode="Batch">
                    </SettingsEditing>
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
        <tr>
            <td>
               </td>
            <td>
                </td>
            <td>
                </td>
        </tr>
        <tr>
            <td align="left">
                <dx1:ASPxButton ID="btnsubmenu" runat="server" Text="SUB MENU" Width="90px" Font-Names="Tahoma"
                    Font-Size="8pt">
                </dx1:ASPxButton>
                </td>
            <td align="right">
                <dx1:ASPxButton ID="btnExcel" runat="server" AutoPostBack="False" ClientInstanceName="btnExcel"
                    Text="EXCEL">
                     <ClientSideEvents Click="function(s, e) {
                        HF.Set('hfTest', '2');
	                    grid.UpdateEdit();
                        grid.PerformCallback('gridExcel');               
                    }" />
                </dx1:ASPxButton>
               </td>
<%--            <td align="right">
                <dx1:ASPxButton ID="btnsend" runat="server" AutoPostBack="False" ClientInstanceName="btnsend"
                    Text="SEND GOOD RECEIVING">
                     <ClientSideEvents Click="function(s, e) {
                        HF.Set('hfTest', '1');
	                    grid.UpdateEdit();	                    
                    }" />
                </dx1:ASPxButton>
               </td>--%>
            <td width="90">
                <dx1:ASPxButton ID="btnshipping" runat="server" Text="SHIPPING INSTRUCTION" Width="90px"
                    Font-Names="Tahoma" Font-Size="8pt" AutoPostBack="False" ClientInstanceName="btnshipping">
                    <ClientSideEvents Click="function(s, e) {
                        HF.Set('hfTest', '0');
	                    grid.UpdateEdit();	 
                        grid.PerformCallback('gridload');                   
                    }" />
                </dx1:ASPxButton>
                </td>
        </tr>
    </table>

</asp:Content>
