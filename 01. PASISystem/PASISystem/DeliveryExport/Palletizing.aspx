<%@ Page Title="" Language="vb" AutoEventWireup="false" MasterPageFile="~/SiteContent.Master" CodeBehind="Palletizing.aspx.vb" Inherits="PASISystem.Palletizing" %>

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
        
        .style1
        {
            width: 122px;
        }
        .style2
        {
            width: 34px;
        }
        .style3
        {
            width: 123px;
        }
        .style4
        {
            width: 213px;
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

    <table style="width:100%;">
        <tr>
            <td>
                <table style="border-width: 1pt thin thin thin; border-style: ridge; border-color:#9598A1; width:100%;">
                    <tr>
                        <td height="30">
                            <table id="Table1">
                                <tr>
                                    <td style="width:5px;"></td>
                                    <td align="left" valign="middle" height="25px" class="style4">
                                        <dx1:ASPxLabel ID="ASPxLabel12" runat="server" Text="AFFILIATE">
                                        </dx1:ASPxLabel>
                                    </td>
                                    <td align="left" valign="middle" height="25px" width="15px">
                                        &nbsp;</td>
                                    <td align="left" valign="middle" height="25px" width="150px">
                <dx1:ASPxComboBox ID="cboaffiliate" runat="server" ClientInstanceName="cboaffiliate"
                    Font-Names="Tahoma" TextFormatString="{0}" Font-Size="8pt">
                    <ClientSideEvents SelectedIndexChanged="function(s, e) {
	                                            txtaffiliate.SetText(cboaffiliate.GetSelectedItem().GetColumnText(1));
                                                }" />
                </dx1:ASPxComboBox>
                                    </td>
                                    <td align="left" valign="middle" style="height:25px; width:10px;"></td>
                                    <td align="left" valign="middle" height="25px" width="120px">
                <dx1:ASPxTextBox ID="txtaffiliate" runat="server" Font-Names="Tahoma" Font-Size="8pt"
                    Width="170px" BackColor="Silver" ClientInstanceName="txtaffiliate" ReadOnly="True">
                    <ClientSideEvents TextChanged="function(s, e) {
	lblerrmessage.SetText('');
}" />
                </dx1:ASPxTextBox>
                                    </td>
                                    <td align="left" valign="middle" height="25px" class="style2">
                                        &nbsp;</td>
                                    <td align="left" valign="middle" height="25px" width="10px"></td>
                                    <td align="left" valign="middle" height="25px" class="style1">
                                        <dx1:ASPxLabel ID="ASPxLabel18" runat="server" Text="ORDER NO">
                                        </dx1:ASPxLabel>
                                    </td>
                                    <td align="left" valign="middle" height="25px" class="style3">
                <dx1:ASPxTextBox ID="txtorderno" runat="server" Font-Names="Tahoma" Font-Size="8pt"
                    Width="170px" ClientInstanceName="txtorderno">
                </dx1:ASPxTextBox>
                                    </td>
                                </tr>
                                <tr>
                                    <td style="width:5px;"></td>
                                    <td align="left" valign="middle" height="25px" class="style4">
                                        <dx1:ASPxLabel ID="ASPxLabel13" runat="server" Text="SUPPLIER">
                                        </dx1:ASPxLabel>
                                    </td>
                                    <td align="left" valign="middle" height="25px" width="15px">
                                        &nbsp;</td>
                                    <td align="left" valign="middle" height="25px" width="150px">
                <dx1:ASPxComboBox ID="cbosupplier" runat="server" ClientInstanceName="cbosupplier"
                    Font-Names="Tahoma" TextFormatString="{0}" Font-Size="8pt">
                    <ClientSideEvents SelectedIndexChanged="function(s, e) {
	                                            txtsupplier.SetText(cbosupplier.GetSelectedItem().GetColumnText(1));
                                                }" />
                </dx1:ASPxComboBox>
                                    </td>
                                    <td align="left" valign="middle" height="25px" width="10px"></td>
                                    <td align="left" valign="middle" height="25px" width="120px">
                <dx1:ASPxTextBox ID="txtsupplier" runat="server" Font-Names="Tahoma" Font-Size="8pt"
                    Width="170px" BackColor="Silver" ClientInstanceName="txtsupplier" ReadOnly="True">
                    <ClientSideEvents TextChanged="function(s, e) {
	lblerrmessage.SetText('');
}" />
                </dx1:ASPxTextBox>
                                    </td>
                                    <td align="left" valign="middle" height="25px" class="style2">
                                        &nbsp;</td>
                                    <td align="left" valign="middle" height="25px" width="10px"></td>
                                    <td align="left" valign="middle" height="25px" class="style1">
                                        <dx1:ASPxLabel ID="ASPxLabel19" runat="server" Text="PALLET NO">
                                        </dx1:ASPxLabel>
                                    </td>
                                    <td align="left" valign="middle" height="25px" class="style3">
                <dx1:ASPxTextBox ID="txtPalletNo" runat="server" Font-Names="Tahoma" Font-Size="8pt"
                    Width="170px" ClientInstanceName="txtPalletNo">
                </dx1:ASPxTextBox>
                                    </td>
                                    <td align="right" valign="middle" style="height:25px; width:100px;">
                                        <table style="width:100%;">
                                            <tr>
                                                <td align="right" valign="middle" style="height:25px; width:90px;">
                                                    &nbsp;</td>
                                                <td align="right" valign="middle" style="height:25px; width:90px;">
                                                    &nbsp;</td>
                                            </tr>
                                        </table>                                        
                                    </td>
                                    <td align="right" valign="middle" style="height:25px; width:100px;">
                                        &nbsp;</td>
                                </tr>                                
                                <tr>
                                    <td style="width:5px;">&nbsp;</td>
                                    <td align="left" valign="middle" height="25px" class="style4">
                                        <dx1:ASPxLabel ID="ASPxLabel14" runat="server" Text="PASI RECEIVING DATE">
                                        </dx1:ASPxLabel>
                                    </td>
                                    <td align="right" valign="middle" height="25px" width="15px">
                <dx1:ASPxCheckBox ID="checkboxdt" runat="server" CheckState="Checked" ClientInstanceName="checkboxdt"
                    Text=" ">
                     <ClientSideEvents CheckedChanged="function (s, e) {
                                                                if (s.GetChecked()==true) {
                                                                    dtReceivingDate.SetEnabled(true);
                                                                } else {
                                                                    dtReceivingDate.SetEnabled(false);
                                                                }
                                                          }" />
                </dx1:ASPxCheckBox>
                                    </td>
                                    <td align="left" valign="middle" height="25px" width="150px">
                <dx1:ASPxDateEdit ID="dtReceivingDate" runat="server" ClientInstanceName="dtReceivingDate" Font-Size="8pt"
                    EditFormat="Custom" EditFormatString="dd MMM yyyy">
                </dx1:ASPxDateEdit>
                                    </td>
                                    <td align="left" valign="middle" height="25px" width="10px">&nbsp;</td>
                                    <td align="left" valign="middle" height="25px" width="120px">
                                        &nbsp;</td>
                                    <td align="left" valign="middle" height="25px" class="style2">
                                        &nbsp;</td>
                                    <td align="left" valign="middle" height="25px" width="10px">&nbsp;</td>
                                    <td align="left" valign="middle" height="25px" class="style1">
                                        <dx1:ASPxLabel ID="ASPxLabel20" runat="server" Text="BOX NO">
                                        </dx1:ASPxLabel>
                                    </td>
                                    <td align="left" valign="middle" height="25px" class="style3">
                <dx1:ASPxTextBox ID="txtBoxNo" runat="server" Font-Names="Tahoma" Font-Size="8pt"
                    Width="170px" ClientInstanceName="txtBoxNo">
                </dx1:ASPxTextBox>
                                    </td>
                                    <td align="right" valign="middle" style="height:25px; width:100px;">
                                        &nbsp;</td>
                                    <td align="right" valign="middle" style="height:25px; width:100px;">
                                        &nbsp;</td>
                                </tr>                                
                                <tr>
                                    <td style="width:5px;">&nbsp;</td>
                                    <td align="left" valign="middle" height="25px" class="style4">
                                        <dx1:ASPxLabel ID="ASPxLabel15" runat="server" Text="INVOICE NO">
                                        </dx1:ASPxLabel>
                                    </td>
                                    <td align="left" valign="middle" height="25px" width="15px">
                                        &nbsp;</td>
                                    <td align="left" valign="middle" height="25px" width="150px">
                <dx1:ASPxTextBox ID="txtInvoiceNo" runat="server" Font-Names="Tahoma" Font-Size="8pt"
                    Width="170px" ClientInstanceName="txtInvoiceNo">
                </dx1:ASPxTextBox>
                                    </td>
                                    <td align="left" valign="middle" height="25px" width="10px">&nbsp;</td>
                                    <td align="left" valign="middle" height="25px" width="120px">
                                        &nbsp;</td>
                                    <td align="left" valign="middle" height="25px" class="style2">
                                        &nbsp;</td>
                                    <td align="left" valign="middle" height="25px" width="10px">&nbsp;</td>
                                    <td align="left" valign="middle" height="25px" class="style1">&nbsp;</td>
                                    <td align="left" valign="middle" height="25px" class="style3">&nbsp;</td>
                                    <td align="right" valign="middle" style="height:25px; width:100px;">
                                        &nbsp;</td>
                                    <td align="right" valign="middle" style="height:25px; width:100px;">
                                        &nbsp;</td>
                                </tr>                                
                                <tr>
                                    <td style="width:5px;">&nbsp;</td>
                                    <td align="left" valign="middle" height="25px" class="style4">
                                        <dx1:ASPxLabel ID="ASPxLabel17" runat="server" Text="STATUS ">
                                        </dx1:ASPxLabel>
                                    </td>
                                    <td align="left" valign="middle" height="25px" width="15px">
                                        &nbsp;</td>
                                    <td align="left" valign="middle" height="25px" width="150px">
                <dx1:ASPxRadioButtonList ID="rbshipping" runat="server" ClientInstanceName="rbshipping"
                    Font-Names="Tahoma" Font-Size="8pt" RepeatDirection="Horizontal">
                    <Items>
                        <dx1:ListEditItem Text="ALREADY SHIPPING" Value="ALREADY" />
                        <dx1:ListEditItem Text="ON PROGRESS" Value="PROGRESS" />
                    </Items>
                    <Border BorderStyle="None" />
                </dx1:ASPxRadioButtonList>
                                    </td>
                                    <td align="left" valign="middle" height="25px" width="10px">&nbsp;</td>
                                    <td align="left" valign="middle" height="25px" width="120px">
                                        &nbsp;</td>
                                    <td align="left" valign="middle" height="25px" class="style2">
                                        &nbsp;</td>
                                    <td align="left" valign="middle" height="25px" width="10px">&nbsp;</td>
                                    <td align="left" valign="middle" height="25px" class="style1">&nbsp;</td>
                                    <td align="left" valign="middle" height="25px" class="style3">&nbsp;</td>
                                    <td align="right" valign="middle" style="height:25px; width:100px;">
                                        &nbsp;</td>
                                    <td align="right" valign="middle" style="height:25px; width:100px;">
                                        &nbsp;</td>
                                </tr>                                
                                <tr>
                                    <td style="width:5px;">&nbsp;</td>
                                    <td align="left" valign="middle" height="25px" class="style4">
                                        &nbsp;</td>
                                    <td align="left" valign="middle" height="25px" width="15px">
                                        &nbsp;</td>
                                    <td align="left" valign="middle" height="25px" width="150px">
                                        &nbsp;</td>
                                    <td align="left" valign="middle" height="25px" width="10px">&nbsp;</td>
                                    <td align="left" valign="middle" height="25px" width="120px">
                                        &nbsp;</td>
                                    <td align="left" valign="middle" height="25px" class="style2">
                                        &nbsp;</td>
                                    <td align="left" valign="middle" height="25px" width="10px">&nbsp;</td>
                                    <td align="left" valign="middle" height="25px" class="style1">&nbsp;</td>
                                    <td align="left" valign="middle" height="25px" class="style3">&nbsp;</td>
                                    <td align="right" valign="middle" style="height:25px; width:100px;">
                            <dx1:ASPxButton ID="btnsearch" runat="server" Text="SEARCH" Width="85px" Font-Names="Tahoma"
                                Font-Size="8pt" ClientInstanceName="btnsearch" AutoPostBack="False">
                                <ClientSideEvents Click="function(s, e) {
                                            var pAffiliate = cboaffiliate.GetText();
                                            var pSupplier = cbosupplier.GetText();
                                            var pReceivingDate = dtReceivingDate.GetText();
                                            var pInvoice = txtInvoiceNo.GetText();
                                            var pshipping = rbshipping.GetValue();
                                            var pOrderNo = txtorderno.GetText();
                                            var pPaletNo = txtPalletNo.GetText();
                                            var pBoxNo = txtBoxNo.GetText();                                            
                                            
                                          
	                                        grid.PerformCallback('gridload' + '|' + pAffiliate + '|' + pSupplier + '|' + pReceivingDate + '|' + pInvoice + '|' + pOrderNo + '|' + pPaletNo + '|' + pBoxNo + '|' + pshipping);

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
                                    <td align="right" valign="middle" style="height:25px; width:100px;">
                            <dx1:ASPxButton ID="btnclear" runat="server" Text="CLEAR" Width="85px" Font-Names="Tahoma"
                                Font-Size="8pt">
                     
                            </dx1:ASPxButton>
                                    </td>
                                </tr>                                
                                <tr>
                                    <td style="width:5px;">&nbsp;</td>
                                    <td align="left" valign="middle" height="25px" class="style4">
                                        &nbsp;</td>
                                    <td align="left" valign="middle" height="25px" width="15px">
                                        &nbsp;</td>
                                    <td align="left" valign="middle" height="25px" width="150px">
                                        &nbsp;</td>
                                    <td align="left" valign="middle" height="25px" width="10px">&nbsp;</td>
                                    <td align="left" valign="middle" height="25px" width="120px">
                                        &nbsp;</td>
                                    <td align="left" valign="middle" height="25px" class="style2">
                                        &nbsp;</td>
                                    <td align="left" valign="middle" height="25px" width="10px">&nbsp;</td>
                                    <td align="left" valign="middle" height="25px" class="style1">&nbsp;</td>
                                    <td align="left" valign="middle" height="25px" class="style3">&nbsp;</td>
                                    <td align="right" valign="middle" style="height:25px; width:100px;">
                                        &nbsp;</td>
                                    <td align="right" valign="middle" style="height:25px; width:100px;">
                                        &nbsp;</td>
                                </tr>                                
                                </table>
                        </td>
                    </tr>
                </table>
            </td>
        </tr>

        </table>
                    
     <script type="text/javascript">
         function OnInit(s, e) {
             AdjustSizeGrid();
         }
         function OnControlsInitializedGrid(s, e) {
             ASPxClientUtils.AttachEventToElement(window, "resize", function (evt) {
                 AdjustSizeGrid();
             });
         }
         function btndeleteClick(s, e) {
             var r = confirm("Are you sure want to Delete ?");
             if (r == true) {
                 grid.PerformCallback('Delete|');
             }
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

    <br />
    <table style="width:100%;">
        <tr>
            <td colspan="4">
                <dx:ASPxGridView ID="grid" runat="server" Width="100%"
                    KeyFieldName="colperiod;colpartno;colaffiliatecode;colHForwarder;colinvoiveno;colHsuratjalan;colorderno;colLabelNo;colpallet" 
                    ClientInstanceName="grid" AutoGenerateColumns="False">
                    <ClientSideEvents EndCallback="function(s, e) {
						var pMsg = s.cpMessage;
                        if( pMsg ) {
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
	dtReceivingDate.SetText(s.cpdtReceivingDate);

}" CallbackError="function(s, e) {
	e.handled = true;
}" />
  <Columns>
           <dx:GridViewCommandColumn ShowSelectCheckbox="True"
                    ShowClearFilterButton="true" VisibleIndex="0" SelectAllCheckboxMode="Page" 
                    Width="30px" FixedStyle="Left" />
                        <dx:GridViewDataTextColumn Caption="PERIOD" FieldName="colperiod" Name="colperiod"
                            VisibleIndex="1" Width="60px">
                            <HeaderStyle Font-Names="Tahoma" Font-Size="8pt" Font-Underline="False" Wrap="True"
                                HorizontalAlign="Center" />
                            <CellStyle Font-Names="Tahoma" Font-Size="8pt">
                            </CellStyle>
                        </dx:GridViewDataTextColumn>
                        <dx:GridViewDataTextColumn Caption="AFFILIATE CODE" FieldName="colaffiliatecode"
                            Name="colaffiliatecode" VisibleIndex="2" Width="70px">
                            <HeaderStyle Font-Names="Tahoma" Font-Size="8pt" Font-Underline="False" Wrap="True"
                                HorizontalAlign="Center" />
                            <CellStyle Font-Names="Tahoma" Font-Size="8pt">
                            </CellStyle>
                        </dx:GridViewDataTextColumn>
                        <dx:GridViewDataTextColumn Caption="AFFILIATE NAME" FieldName="colaffiliatename"
                            Name="colaffiliatename" VisibleIndex="3" Width="0px">
                            <HeaderStyle Font-Names="Tahoma" Font-Size="8pt" Font-Underline="False" Wrap="True"
                                HorizontalAlign="Center" />
                            <CellStyle Font-Names="Tahoma" Font-Size="8pt">
                            </CellStyle>
                        </dx:GridViewDataTextColumn>
                         <dx:GridViewDataTextColumn Caption="SUPPLIER CODE" FieldName="colsuppliercode" Name="colsuppliercode"
                            VisibleIndex="4" Width="70px">
                            <HeaderStyle Font-Names="Tahoma" Font-Size="8pt" Font-Underline="False" Wrap="True"
                                HorizontalAlign="Center" />
                            <CellStyle Font-Names="Tahoma" Font-Size="8pt">
                            </CellStyle>
                        </dx:GridViewDataTextColumn>
                            <dx:GridViewDataTextColumn Caption="SUPPLIER NAME" FieldName="colsuppliername"
                            Name="colsuppliername" VisibleIndex="5" Width="0px">
                        </dx:GridViewDataTextColumn>
                        <dx:GridViewDataTextColumn Caption="ORDER NO." FieldName="colorderno" 
                            Name="colorderno" VisibleIndex="6"
                            Width="80px">
                        </dx:GridViewDataTextColumn>
                        <dx:GridViewDataTextColumn Caption="INVOICE NO" FieldName="colinvoiveno" Name="colinvoiveno"
                            VisibleIndex="7" Width="180px">
                            <HeaderStyle Font-Names="Tahoma" Font-Size="8pt" Font-Underline="False" Wrap="True"
                                HorizontalAlign="Center" />
                            <CellStyle Font-Names="Tahoma" Font-Size="8pt" HorizontalAlign="Center">
                            </CellStyle>
                        </dx:GridViewDataTextColumn>
                        <dx:GridViewDataTextColumn Caption="PALLET NO" FieldName="colpallet"
                            Name="colpallet" VisibleIndex="8" Width="120px">
                            <HeaderStyle Font-Names="Tahoma" Font-Size="8pt" Font-Underline="False" Wrap="True"
                                HorizontalAlign="Center" />
                            <CellStyle Font-Names="Tahoma" Font-Size="8pt">
                            </CellStyle>
                        </dx:GridViewDataTextColumn>
                        <dx:GridViewDataTextColumn Caption="PART NO." FieldName="colpartno"
                            Name="colpartno" VisibleIndex="10" Width="120px">
                            <HeaderStyle Font-Names="Tahoma" Font-Size="8pt" Font-Underline="False" Wrap="True"
                                HorizontalAlign="Center" />
                            <CellStyle Font-Names="Tahoma" Font-Size="8pt">
                            </CellStyle>
                            </dx:GridViewDataTextColumn>
                       <%--<dx:GridViewCommandColumn Caption="  " FieldName="ACT" Name="ACT" 
                      "Left" SelectAllCheckboxMode="Page">
                            <PropertiesCheckEdit ValueChecked="1" ValueType="System.Int32" ValueUnchecked="0">
                            </PropertiesCheckEGridViewDataCheckColumn     </dx:GridViewCommandColumn>--%>
                        <dx:GridViewDataTextColumn FieldName="colLabelNo" Name="colLabelNo" VisibleIndex="9" Width="120px" 
                            Caption="LABEL NO.">
                            <HeaderStyle HorizontalAlign="Center" />
                        </dx:GridViewDataTextColumn>
                        <dx:GridViewDataTextColumn Caption="QTY" FieldName="colqty" 
                            VisibleIndex="12" Name="colqty" Width="80px">
                        </dx:GridViewDataTextColumn>
                        <dx:GridViewDataTextColumn Caption="LOCATION" FieldName="collocation"
                            Name="collocation" VisibleIndex="13" Width="80px">
                            <HeaderStyle Font-Names="Tahoma" Font-Size="8pt" Font-Underline="False" Wrap="True"
                                HorizontalAlign="Center" />
                            <CellStyle Font-Names="Tahoma" Font-Size="8pt">
                            </CellStyle>
                        </dx:GridViewDataTextColumn>
                        <dx:GridViewDataTextColumn Caption="PALLET TYPE" FieldName="colpallettype" Name="colpallettype"
                            VisibleIndex="14" Width="80px">
                            <HeaderStyle Font-Names="Tahoma" Font-Size="8pt" Font-Underline="False" Wrap="True"
                                HorizontalAlign="Center" />
                            <CellStyle Font-Names="Tahoma" Font-Size="8pt" HorizontalAlign="Left">
                            </CellStyle>
                        </dx:GridViewDataTextColumn>
                        <dx:GridViewDataTextColumn FieldName="colHForwarder" Name="colHForwarder" VisibleIndex="9" Width="120px" Visible = "false" 
                            >
                            <HeaderStyle HorizontalAlign="Center" />
                        </dx:GridViewDataTextColumn>
                        <dx:GridViewDataTextColumn FieldName="colHsuratjalan" Name="colHsuratjalan" VisibleIndex="9" Width="120px" Visible = "false"
                            >
                            <HeaderStyle HorizontalAlign="Center" />
                        </dx:GridViewDataTextColumn>
                    </Columns>
                    <SettingsPager Visible="False" Position="Top">
                        <Summary Text="Page {0} of {1} [{2} record(s)]" AllPagesText="Page {0} of {1} " />
                        <Summary AllPagesText="Page {0} of {1} " Text="Page {0} of {1} [{2} record(s)]">
                        </Summary>
                    </SettingsPager>
                     <Settings ShowGroupButtons="False" ShowVerticalScrollBar="True" ShowHorizontalScrollBar="True" 
                        VerticalScrollableHeight="250" ShowStatusBar="Hidden"></Settings>
                    <Styles>
                                <Header BackColor="#FFD2A6" Font-Names="Verdana" Font-Size="8pt"></Header>
                                <Row BackColor="#FFFFE1" ForeColor = "Black" Font-Names="Verdana" Font-Size="8pt" Wrap="False"></Row>
                                <RowHotTrack BackColor="#E8EFFD" ForeColor = "Black" Font-Names="Verdana" Font-Size="8pt" Wrap="False"></RowHotTrack>
                                <SelectedRow BackColor="#ccffcc" ForeColor = "Black" Wrap="False">
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
            <td colspan="2">
                <dx1:ASPxButton ID="btnprint" runat="server" Text="PRINT" Width="90px"
                    Font-Names="Tahoma" Font-Size="8pt" AutoPostBack="False" Visible = "false"
                    ClientInstanceName="btnprint">
                </dx1:ASPxButton>

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
                &nbsp;</td>
            <td align="right">
                <dx1:ASPxButton ID="btndelete" runat="server" Text="DELETE" Width="90px"
                    Font-Names="Tahoma" Font-Size="8pt" AutoPostBack="False" 
                    ClientInstanceName="btndelete">  
                    <ClientSideEvents Click="btndeleteClick"/>
                    <Paddings Padding="2px" />
                </dx1:ASPxButton>
               </td>
            <td width="90">
            <dx1:ASPxButton ID="btnExcel" runat="server" AutoPostBack="False" ClientInstanceName="btnExcel" Width="90px"
                    Font-Names="Tahoma" Font-Size="8pt" Text="EXCEL">
                     <ClientSideEvents Click="function(s, e) {
                    
                        grid.PerformCallback('gridExcel');               
                    }" />
                     <Paddings Padding="2px" />
                </dx1:ASPxButton>

                </td>
        </tr>
    </table>

</asp:Content>

