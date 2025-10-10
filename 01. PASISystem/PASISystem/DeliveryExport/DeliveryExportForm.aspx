<%@ Page Title="" Language="vb" AutoEventWireup="false" MasterPageFile="~/SiteContent.Master" CodeBehind="DeliveryExportForm.aspx.vb" Inherits="PASISystem.DeliveryExportForm" %>

<%@ Register Assembly="DevExpress.Web.v14.1, Version=14.1.6.0, Culture=neutral, PublicKeyToken=b88d1754d700e49a" Namespace="DevExpress.Web.ASPxCallback" TagPrefix="dx" %>
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
        .style7
        {
            height: 17px;
        }
        .style8
        {
            width: 5px;
            height: 25px;
        }
        .style9
        {
            width: 213px;
            height: 25px;
        }
        .style10
        {
            height: 25px;
        }
        .style11
        {
            width: 10px;
            height: 25px;
        }
        .style12
        {
            width: 34px;
            height: 25px;
        }
        .style13
        {
            width: 123px;
            height: 25px;
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
                                    <td class="style8"></td>
                                    <td align="left" valign="middle" class="style9">
                                        &nbsp;</td>
                                    <td align="left" valign="middle" width="15px">
                                        </td>
                                    <td align="left" valign="middle" width="150px" class="style10">
                                        &nbsp;</td>
                                    <td align="left" valign="middle" class="style11"></td>
                                    <td align="left" valign="middle" width="120px" class="style10">
                                        &nbsp;</td>
                                    <td align="left" valign="middle" class="style12">
                                        </td>
                                    <td align="left" valign="middle" width="10px" class="style10"></td>
                                    <td align="left" valign="middle" class="style10">
                                        </td>
                                    <td align="left" valign="middle" class="style13">
                                        </td>
                                </tr>
                                <tr>
                                    <td style="width:5px;"></td>
                                    <td align="left" valign="middle" height="25px" class="style4">
                                        &nbsp;</td>
                                    <td align="left" valign="middle" height="25px" width="15px">
                                        &nbsp;</td>
                                    <td align="left" valign="middle" height="25px" width="150px">
                                        <dx1:ASPxComboBox runat="server" Width="100px" ClientInstanceName="cboCreate" 
                                            ID="cboCreate">
                                              <ClientSideEvents SelectedIndexChanged="function(s, e) {
	                                                                if (cboCreate.GetText() == 'CREATE') {
	                                                                checkboxdt.SetEnabled(false);
                                                                    dtFDeliveryDateFrom.SetEnabled(false);
                                                                    dtFDeliveryDateEnd.SetEnabled(false);
                                                                    cboFaffiliate.SetEnabled(false);
                                                                    cboFForwarder.SetEnabled(false);
                                                                    txtFaffiliate.SetEnabled(false);
                                                                    txtFForwarder.SetEnabled(false);
                                                                    txtFSJNo.SetText('')
                                                                    txtFSJNo.SetEnabled(false);
                                                                    txtFContainerNo.SetText('')
                                                                    txtFContainerNo.SetEnabled(false);
                                                                } else {
                                                                    checkboxdt.SetEnabled(true);
                                                                    txtFContainerNo.SetEnabled(true);
                                                                    cboFaffiliate.SetEnabled(true);
                                                                    cboFForwarder.SetEnabled(true);
                                                                    txtFaffiliate.SetEnabled(true);
                                                                    txtFForwarder.SetEnabled(true);
                                                                    txtFSJNo.SetEnabled(true);
                                                                    dtFDeliveryDateFrom.SetEnabled(true);
                                                                    dtFDeliveryDateEnd.SetEnabled(true);
                                                       }
                                            }" />
                                            </dx1:ASPxComboBox>

                                    </td>
                                    <td align="left" valign="middle" height="25px" width="10px">&nbsp;</td>
                                    <td align="left" valign="middle" height="25px" width="120px">
                                        &nbsp;</td>
                                    <td align="left" valign="middle" height="25px" class="style2">
                                        &nbsp;</td>
                                    <td align="left" valign="middle" height="25px" width="10px"></td>
                                    <td align="left" valign="middle" class="style1" colspan="4" rowspan="11">
                <dx:ASPxRoundPanel ID="ASPxRoundPanel1" runat="server" HeaderText="" Height="318px"
                    ShowCollapseButton="true" View="GroupBox" Width="100%" style="margin-top: 0px">
                    <PanelCollection>
                        <dx:PanelContent ID="PanelContent1" runat="server">
                            <table style="width: 100%;">
                                <tr>
                                    <td>
                                        <dx1:ASPxLabel ID="ASPxLabel23" runat="server" Font-Names="Tahoma" Font-Size="8pt"
                                            Text="SURAT JALAN" Width="145px">
                                        </dx1:ASPxLabel>
                                    </td>
                                    <td>
                                        &nbsp;</td>
                                    <td>
                                        <dx1:ASPxTextBox ID="txtSJNo" runat="server" 
                                            ClientInstanceName="txtSJNo" Font-Names="Tahoma" Font-Size="8pt" 
                                            Width="170px" TabIndex="9">
                                        </dx1:ASPxTextBox>
                                    </td>
                                </tr>
                                <tr>
                                    <td class="style7">
                                        <dx1:ASPxLabel ID="ASPxLabel33" runat="server" Font-Names="Tahoma" 
                                            Font-Size="8pt" Text="CONTAINER" Width="145px">
                                        </dx1:ASPxLabel>
                                    </td>
                                    <td class="style7">
                                        &nbsp;</td>
                                    <td class="style7">
                                        <dx1:ASPxComboBox ID="cboContainer" runat="server" 
                                            ClientInstanceName="cboContainer" Font-Names="Tahoma" Font-Size="8pt" 
                                            TextFormatString="{0}">
                                        </dx1:ASPxComboBox>
                                    </td>
                                </tr>
                                <tr>
                                    <td class="style7">
                                        <dx1:ASPxLabel ID="ASPxLabel32" runat="server" Font-Names="Tahoma" 
                                            Font-Size="8pt" Text="DELIVERY DATE" Width="145px">
                                        </dx1:ASPxLabel>
                                    </td>
                                    <td class="style7">
                                        &nbsp;</td>
                                    <td class="style7">
                                        <dx1:ASPxDateEdit ID="dtDeliveryDate" runat="server" 
                                            ClientInstanceName="dtDeliveryDate" EditFormat="Custom" 
                                            EditFormatString="dd MMM yyyy" Font-Size="8pt" TabIndex="17">
                                        </dx1:ASPxDateEdit>
                                    </td>
                                </tr>
                                <tr>
                                    <td class="style7">
                                        <dx1:ASPxLabel ID="ASPxLabel26" runat="server" Font-Names="Tahoma" 
                                            Font-Size="8pt" Text="PIC" Width="145px">
                                        </dx1:ASPxLabel>
                                    </td>
                                    <td class="style7">
                                        &nbsp;</td>
                                    <td class="style7">
                                        <dx1:ASPxTextBox ID="txtPIC" runat="server" ClientInstanceName="txtPIC" 
                                            Font-Names="Tahoma" Font-Size="8pt" Width="170px" TabIndex="11">
                                        </dx1:ASPxTextBox>
                                    </td>
                                </tr>
                                  <tr>
                                      <td class="style7">
                                          <dx1:ASPxLabel ID="ASPxLabel27" runat="server" Font-Names="Tahoma" 
                                              Font-Size="8pt" Text="JENIS ARMADA" Width="145px" >
                                          </dx1:ASPxLabel>
                                      </td>
                                      <td class="style7">
                                          &nbsp;</td>
                                      <td class="style7">
                                          <dx1:ASPxTextBox ID="txtJenisArmada" runat="server" 
                                              ClientInstanceName="txtJenisArmada" Font-Names="Tahoma" Font-Size="8pt" 
                                              TabIndex="13" Width="170px">
                                          </dx1:ASPxTextBox>


                                      </td>
                                </tr>
                                <tr>
                                    <td class="style7">
                                        <dx1:ASPxLabel ID="ASPxLabel28" runat="server" Font-Names="Tahoma" 
                                            Font-Size="8pt" Text="DRIVER NAME" Width="145px" >
                                        </dx1:ASPxLabel>
                                    </td>
                                    <td class="style7">
                                        &nbsp;</td>
                                    <td class="style7">
                                        <dx1:ASPxTextBox ID="txtDriverName" runat="server" 
                                            ClientInstanceName="txtDriverName" Font-Names="Tahoma" Font-Size="8pt" TabIndex="13"
                                            Width="170px">
                                        </dx1:ASPxTextBox>
                                    </td>
                                </tr>
                                <tr>
                                    <td class="style7">
                                        <dx1:ASPxLabel ID="ASPxLabel29" runat="server" Font-Names="Tahoma" 
                                            Font-Size="8pt" Text="DRIVER CONTACT" Width="145px">
                                        </dx1:ASPxLabel>
                                    </td>
                                    <td class="style7">
                                    </td>
                                    <td class="style7">
                                        <dx1:ASPxTextBox ID="txtDriverContact" runat="server" 
                                            ClientInstanceName="txtDriverContact" Font-Names="Tahoma" Font-Size="8pt" 
                                            TabIndex="14" Width="170px">
                                        </dx1:ASPxTextBox>
                                    </td>
                                </tr>
                                <tr>
                                    <td class="style7">
                                        <dx1:ASPxLabel ID="ASPxLabel30" runat="server" Font-Names="Tahoma" 
                                            Font-Size="8pt" Text="NO POL" Width="145px">
                                        </dx1:ASPxLabel>
                                    </td>
                                    <td class="style7">
                                        &nbsp;</td>
                                    <td class="style7">
                                        <dx1:ASPxTextBox ID="txtNoPol" runat="server" ClientInstanceName="txtNoPol" 
                                            Font-Names="Tahoma" Font-Size="8pt" TabIndex="15" Width="170px">
                                        </dx1:ASPxTextBox>
                                    </td>
                                </tr>
                                <tr>
                                    <td class="style7">
                                        <dx1:ASPxLabel ID="ASPxLabel31" runat="server" Font-Names="Tahoma" 
                                            Font-Size="8pt" Text="TOTAL BOX" Width="145px">
                                        </dx1:ASPxLabel>
                                    </td>
                                    <td class="style7">
                                        &nbsp;</td>
                                    <td class="style7">
                                        <dx1:ASPxTextBox ID="txtTotalBox" runat="server" 
                                            ClientInstanceName="txtTotalBox" Font-Names="Tahoma" Font-Size="8pt" 
                                            TabIndex="16" Width="170px">
                                        </dx1:ASPxTextBox>
                                    </td>
                                </tr>
                                <tr>
                                    <td class="style7">
                                        <dx1:ASPxLabel ID="ASPxLabel25" runat="server" Font-Names="Tahoma" 
                                            Font-Size="8pt" Text="TOTAL PALLET" Width="145px">
                                        </dx1:ASPxLabel>
                                    </td>
                                    <td class="style7">
                                        &nbsp;</td>
                                    <td class="style7">
                                        <dx1:ASPxTextBox ID="txtTotalPallet" runat="server" 
                                            ClientInstanceName="txtTotalPallet" Font-Names="Tahoma" Font-Size="8pt" 
                                            TabIndex="18" Width="170px">
                                        </dx1:ASPxTextBox>
                                    </td>
                                </tr>
                                  <tr>
                                    <td>
                                        &nbsp;</td>
                                      <td>
                                          &nbsp;</td>
                                      <td>
                                          &nbsp;</td>
                                </tr>
                                <tr>
                                    <td>
                                        &nbsp;</td>
                                    <td>
                                        &nbsp;</td>
                                    <td align="right">
                                        <dx1:ASPxButton ID="btnsave" runat="server" AutoPostBack="False" 
                                            ClientInstanceName="btnsave" Font-Names="Tahoma" Font-Size="8pt" Text="SAVE" TabIndex="19"
                                            Width="90px">
                                            <ClientSideEvents Click="function(s, e) {
                                                validasisubmit();
                                                 if (valid !== 'false') 
                                                     {
                                                      up_Insert();
                                                     }
                                            }" />
                                            <Paddings Padding="2px" />
                                        </dx1:ASPxButton>
                                    </td>
                                </tr>
                            </table>
                        </dx:PanelContent>
                    </PanelCollection>
                </dx:ASPxRoundPanel>
                                    </td>
                                </tr>                                
                                <tr>
                                    <td style="width:5px;">&nbsp;</td>
                                    <td align="left" valign="middle" height="25px" class="style4">
                                        <dx1:ASPxLabel ID="ASPxLabel17" runat="server" Text="STUFFING DATE">
                                        </dx1:ASPxLabel>
                                    </td>
                                    <td align="left" valign="middle" height="25px" width="15px">
                <dx1:ASPxCheckBox ID="checkboxdt" runat="server" CheckState="Checked" ClientInstanceName="checkboxdt"
                    Text=" ">
                    <ClientSideEvents CheckedChanged="function (s, e) {
                                                                if (s.GetChecked()==true) {
                                                                    dtFDeliveryDateFrom.SetEnabled(true);
                                                                    dtFDeliveryDateEnd.SetEnabled(true);
                                                                } else {
                                                                    dtFDeliveryDateFrom.SetEnabled(false);
                                                                    dtFDeliveryDateEnd.SetEnabled(false);
                                                                }
                                                          }" />
                </dx1:ASPxCheckBox>
                                    </td>
                                    <td align="left" valign="middle" height="25px" width="150px">
                <dx1:ASPxDateEdit ID="dtFDeliveryDateFrom" runat="server" ClientInstanceName="dtFDeliveryDateFrom" Font-Size="8pt"
                    EditFormat="Custom" EditFormatString="dd MMM yyyy" TabIndex="5">
                </dx1:ASPxDateEdit>
                                    </td>
                                    <td align="left" valign="middle" height="25px" width="10px">
                                        <dx1:ASPxLabel ID="ASPxLabel22" runat="server" Text="   TO  " 
                                            Font-Size="X-Small">
                                        </dx1:ASPxLabel>
                                    </td>
                                    <td align="left" valign="middle" height="25px" width="120px">
                <dx1:ASPxDateEdit ID="dtFDeliveryDateEnd" runat="server" ClientInstanceName="dtFDeliveryDateEnd" Font-Size="8pt"
                    EditFormat="Custom" EditFormatString="dd MMM yyyy" TabIndex="6">
                </dx1:ASPxDateEdit>

                                    </td>
                                    <td align="left" valign="middle" height="25px" class="style2">
                                        &nbsp;</td>
                                    <td align="left" valign="middle" height="25px" width="10px">&nbsp;</td>
                                </tr>                                
                                <tr>
                                    <td style="width:5px;">&nbsp;</td>
                                    <td align="left" valign="middle" height="25px" class="style4">
                                        <dx1:ASPxLabel ID="ASPxLabel12" runat="server" Text="AFFILIATE">
                                        </dx1:ASPxLabel>
                                    </td>
                                    <td align="left" valign="middle" height="25px" width="15px">
                                        &nbsp;</td>
                                    <td align="left" valign="middle" height="25px" width="150px">
                <dx1:ASPxComboBox ID="cboFaffiliate" runat="server" ClientInstanceName="cboFaffiliate"
                    Font-Names="Tahoma" TextFormatString="{0}" Font-Size="8pt">
                    <ClientSideEvents SelectedIndexChanged="function(s, e) {
	                                            txtFaffiliate.SetText(cboFaffiliate.GetSelectedItem().GetColumnText(1));
                                                }" />
                </dx1:ASPxComboBox>
                                    </td>
                                    <td align="left" valign="middle" height="25px" width="10px">&nbsp;</td>
                                    <td align="left" valign="middle" height="25px" width="120px">
                <dx1:ASPxTextBox ID="txtFaffiliate" runat="server" Font-Names="Tahoma" Font-Size="8pt"
                    Width="170px" BackColor="Silver" ClientInstanceName="txtFaffiliate" ReadOnly="True">
                    <ClientSideEvents TextChanged="function(s, e) {
	lblerrmessage.SetText('');
}" />
                </dx1:ASPxTextBox>
                                    </td>
                                    <td align="left" valign="middle" height="25px" class="style2">
                                        &nbsp;</td>
                                    <td align="left" valign="middle" height="25px" width="10px">&nbsp;</td>
                                </tr>                                
                                <tr>
                                    <td style="width:5px;">&nbsp;</td>
                                    <td align="left" valign="middle" height="25px" class="style4">
                                        <dx1:ASPxLabel ID="ASPxLabel13" runat="server" Text="FORWARDER">
                                        </dx1:ASPxLabel>
                                    </td>
                                    <td align="left" valign="middle" height="25px" width="15px">
                                        &nbsp;</td>
                                    <td align="left" valign="middle" height="25px" width="150px">
                <dx1:ASPxComboBox ID="cboFForwarder" runat="server" ClientInstanceName="cboFForwarder"
                    Font-Names="Tahoma" TextFormatString="{0}" Font-Size="8pt" TabIndex="1">
                    <ClientSideEvents SelectedIndexChanged="function(s, e) {
	                                            txtFForwarder.SetText(cboFForwarder.GetSelectedItem().GetColumnText(1));
                                                }" />
                </dx1:ASPxComboBox>
                                    </td>
                                    <td align="left" valign="middle" height="25px" width="10px">&nbsp;</td>
                                    <td align="left" valign="middle" height="25px" width="120px">
                <dx1:ASPxTextBox ID="txtFForwarder" runat="server" Font-Names="Tahoma" Font-Size="8pt"
                    Width="170px" BackColor="Silver" ClientInstanceName="txtFForwarder" ReadOnly="True">
                    <ClientSideEvents TextChanged="function(s, e) {
	lblerrmessage.SetText('');
}" />
                </dx1:ASPxTextBox>
                                    </td>
                                    <td align="left" valign="middle" height="25px" class="style2">
                                        &nbsp;</td>
                                    <td align="left" valign="middle" height="25px" width="10px">&nbsp;</td>
                                </tr>                                
                                <tr>
                                    <td style="width:5px;">&nbsp;</td>
                                    <td align="left" valign="middle" height="25px" class="style4">
                                        <dx1:ASPxLabel ID="ASPxLabel14" runat="server" Text="SURAT JALAN NO.">
                                        </dx1:ASPxLabel>
                                    </td>
                                    <td align="right" valign="middle" height="25px" width="15px">
                                        &nbsp;</td>
                                    <td align="left" valign="middle" height="25px" width="150px">
                <dx1:ASPxTextBox ID="txtFSJNo" runat="server" Font-Names="Tahoma" Font-Size="8pt"
                    Width="170px" ClientInstanceName="txtFSJNo" TabIndex="2">
                </dx1:ASPxTextBox>
                                    </td>
                                    <td align="center" valign="middle" height="25px" width="10px">
                                        &nbsp;</td>
                                    <td align="left" valign="middle" height="25px" width="120px">
                                        &nbsp;</td>
                                    <td align="left" valign="middle" height="25px" class="style2">
                                        &nbsp;</td>
                                    <td align="left" valign="middle" height="25px" width="10px">&nbsp;</td>
                                </tr>                                
                                <tr>
                                    <td style="width:5px;">&nbsp;</td>
                                    <td align="left" valign="middle" height="25px" class="style4">
                                        <dx1:ASPxLabel ID="ASPxLabel15" runat="server" Text="CONTAINER NO">
                                        </dx1:ASPxLabel>
                                    </td>
                                    <td align="left" valign="middle" height="25px" width="15px">
                                        &nbsp;</td>
                                    <td align="left" valign="middle" height="25px" width="150px">
                <dx1:ASPxTextBox ID="txtFContainerNo" runat="server" Font-Names="Tahoma" Font-Size="8pt"
                    Width="170px" ClientInstanceName="txtFContainerNo" TabIndex="3">
                </dx1:ASPxTextBox>
                                    </td>
                                    <td align="left" valign="middle" height="25px" width="10px">&nbsp;</td>
                                    <td align="left" valign="middle" height="25px" width="120px">
                                        &nbsp;</td>
                                    <td align="left" valign="middle" height="25px" class="style2">
                                        &nbsp;</td>
                                    <td align="left" valign="middle" height="25px" width="10px">&nbsp;</td>
                                </tr>                                
                                <tr>
                                    <td style="width:5px;">&nbsp;</td>
                                    <td align="left" valign="middle" height="25px" class="style4">
                                        &nbsp;</td>
                                    <td align="left" valign="middle" height="25px" width="15px">
                                        &nbsp;</td>
                                    <td align="right" valign="middle" height="25px" width="150px">
                            <dx1:ASPxButton ID="btnsearch" runat="server" Text="SEARCH" Width="85px" Font-Names="Tahoma"
                                Font-Size="8pt" ClientInstanceName="btnsearch" AutoPostBack="False" TabIndex="7">
                                <ClientSideEvents Click="function(s, e) {
                                            var pAffiliate = cboFaffiliate.GetText();
                                            var pForwarder = cboFForwarder.GetText();
                                            var pSJNo = txtFSJNo.GetText();
                                            var pContainer = txtFContainerNo.GetText();
                                            var pDeliveryDateFrom = dtFDeliveryDateFrom.GetText();
                                            var pDeliveryDateEnd = dtFDeliveryDateEnd.GetText();
                                          
	                                        grid.PerformCallback('gridload' + '|' + pAffiliate + '|' + pForwarder + '|' + pSJNo + '|' + pContainer + '|' + pDeliveryDateFrom + '|' + pDeliveryDateEnd);


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
                                    <td align="left" valign="middle" height="25px" width="10px">&nbsp;</td>
                                    <td align="left" valign="middle" height="25px" width="120px">
                            <dx1:ASPxButton ID="btnclear" runat="server" Text="CLEAR" Width="85px" Font-Names="Tahoma"
                                Font-Size="8pt" TabIndex="8">
                     
                            </dx1:ASPxButton>
                                    </td>
                                    <td align="left" valign="middle" height="25px" class="style2">
                                        &nbsp;</td>
                                    <td align="left" valign="middle" height="25px" width="10px">&nbsp;</td>
                                </tr>                                
                                <tr>
                                    <td style="width:5px;">&nbsp;</td>
                                    <td align="left" valign="middle" height="25px" class="style4">
                                        &nbsp;</td>
                                    <td align="left" valign="middle" height="25px"  width="15px">
                                        &nbsp;</td>
                                    <td align="left" valign="middle" height="25px" width="150px">
                                        &nbsp;</td>
                                    <td align="left" valign="middle" height="25px" width="10px">&nbsp;</td>
                                    <td align="left" valign="middle" height="25px" width="120px">
                                        &nbsp;</td>
                                    <td align="left" valign="middle" height="25px" class="style2">
                                        &nbsp;</td>
                                    <td align="left" valign="middle" height="25px" width="10px">&nbsp;</td>
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
                                </tr>                                
                                <tr>
                                    <td style="width:5px;">&nbsp;</td>
                                    <td align="left" valign="middle" height="25px" class="style4">
                                        &nbsp;</td>
                                    <td align="left" valign="middle" height="7px" width="15px">
                                        &nbsp;</td>
                                    <td align="left" valign="middle" height="25px" width="150px">
                                        &nbsp;</td>
                                    <td align="left" valign="middle" height="25px" width="10px">&nbsp;</td>
                                    <td align="left" valign="middle" height="25px" width="120px">
                                        &nbsp;</td>
                                    <td align="left" valign="middle" height="25px" class="style2">
                                        &nbsp;</td>
                                    <td align="left" valign="middle" height="25px" width="10px">&nbsp;</td>
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

         function numbersonly(e) {
             var unicode = e.charCode ? e.charCode : e.keyCode
             if (unicode != 8) { //if the key isn't the backspace key (which we should allow)
                 if (unicode < 45 || unicode > 57) //if not a number
                                      return false //disable key press
             }
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
         var valid = '';
         function validasisubmit() {
             valid = ''
             lblerrmessage.GetMainElement().style.color = 'Red';



             if (txtTotalPallet.GetText() == "") {
                 lblerrmessage.SetText("[7020] Please Input Total Pallet first!");
                 txtTotalPallet.Focus();
                 valid = 'false';
             }

             if (txtTotalBox.GetText() == "") {
                 lblerrmessage.SetText("[7019] Please Input Total Box first!");
                 txtTotalBox.Focus();
                 valid = 'false';
             }


             if (txtNoPol.GetText() == "") {
                 lblerrmessage.SetText("[7018] Please Input No. Pol first!");
                 txtNoPol.Focus();
                 valid = 'false';
             }

             if (txtDriverContact.GetText() == "") {
                 lblerrmessage.SetText("[7017] Please Input Driver Contact first!");
                 txtDriverContact.Focus();
                 valid = 'false';
             }


             if (txtDriverName.GetText() == "") {
                 lblerrmessage.SetText("[7016] Please Input Driver Name first!");
                 txtDriverName.Focus();
                 valid = 'false';
             }

             if (txtJenisArmada.GetText() == "") {
                 lblerrmessage.SetText("[7015] Please Input Jenis Armada first!");
                 txtJenisArmada.Focus();
                 valid = 'false';
             }

             if (txtPIC.GetText() == "") {
                 lblerrmessage.SetText("[7014] Please Input PIC first!");
                 txtPIC.Focus();
                 valid = 'false';
             }

             if (dtDeliveryDate.GetText() == "") {
                 lblerrmessage.SetText("[7013] Please Input Delivery Date first!");
                 dtDeliveryDate.Focus();
                 valid = 'false';
             }

             if (txtSJNo.GetText() == "") {
                 lblerrmessage.SetText("[7012] Please Input Surat Jalan first!");
                 txtSJNo.Focus();
                 valid = 'false';
             }

             
                 
         }
         function up_Insert() {
             var pIsUpdate = '';
             var pSJNo = txtSJNo.GetText();
             var pDelivery = dtDeliveryDate.GetText();
             var pPIC = txtPIC.GetText();
             var pJenisArmada = txtJenisArmada.GetText();
             var pDriverName = txtDriverName.GetText();
             var pDriverContact = txtDriverContact.GetText();
             var pNoPol = txtNoPol.GetText();
             var pTotalBox = txtTotalBox.GetText();
             var pTotalPallet = txtTotalPallet.GetText();

             SaveSubmit.PerformCallback('save|' + pIsUpdate + '|' + pSJNo + '|' + pDelivery + '|' + pPIC + '|' + pJenisArmada + '|' + pDriverName + '|' + pDriverContact + '|' + pNoPol + '|' + pTotalBox + '|' + pTotalPallet );
         }

         function OnGridFocusedRowChangedDeliveryExport() {
             grid.GetRowValues(grid.GetFocusedRowIndex(), 'colInvoice;colForwarder;colAffiliate;colSuratJalan;colDeliveryDate;colPIC;colJenisArmada;colDriverName;colDriverContact;colNoPol;colSumTotalBox;colTotalPallet', OnGetRowValuesDeliveryExport);
         }

         function OnGetRowValuesDeliveryExport(values) {

             txtSJNo.SetText(values[3]);
             dtDeliveryDate.SetText(Date());
             txtPIC.SetText(values[5]);
             txtJenisArmada.SetText(values[6]);
             txtDriverName.SetText(values[7]);
             txtDriverContact.SetText(values[8]);
             txtNoPol.SetText(values[9]);
             txtTotalBox.SetText(values[10]);
             txtTotalPallet.SetText(values[11]);
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
    <table style="width:98%;">
        <tr>
            <td colspan="4">
                <dx:ASPxGridView ID="grid" runat="server" Width="100%"
                    KeyFieldName="colSuratJalan;colForwarder;colAffiliate;colInvoice;colContainer;colpallet;colOrder;colPart;colBoxFrom;colBoxTo;colTotalBox;colQTY;NoUrut"
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
	dtFDeliveryDateFrom.SetText(s.cpdtFDeliveryDateFrom);
    dtFDeliveryDateEnd.SetText(s.cpdtFDeliveryDateEnd);
    dtDeliveryDate.SetText(s.cpdtDeliveryDate);

}" CallbackError="function(s, e) {
	e.handled = true;
}" 
   RowDblClick="function(s, e) {OnGridFocusedRowChangedDeliveryExport();}" 
   />
  <Columns>
           <dx:GridViewCommandColumn ShowSelectCheckbox="True"
                    ShowClearFilterButton="true" VisibleIndex="0" SelectAllCheckboxMode="Page" 
                    Width="30px" FixedStyle="Left" Name="ACT" />
                        <dx:GridViewDataTextColumn Caption="INVOICE NO" 
                               FieldName="colInvoice" Name="colInvoice"
                               VisibleIndex="1" Width="0px">
                            <HeaderStyle Font-Names="Tahoma" Font-Size="8pt" Font-Underline="False" Wrap="True"
                                HorizontalAlign="Center" />
                            <CellStyle Font-Names="Tahoma" Font-Size="8pt">
                            </CellStyle>
                        </dx:GridViewDataTextColumn>
                        <dx:GridViewDataTextColumn Caption="FORWARDER" FieldName="colForwarder"
                            Name="colForwarder" VisibleIndex="2" Width="0px">
                            <HeaderStyle Font-Names="Tahoma" Font-Size="8pt" Font-Underline="False" Wrap="True"
                                HorizontalAlign="Center" />
                            <CellStyle Font-Names="Tahoma" Font-Size="8pt">
                            </CellStyle>
                        </dx:GridViewDataTextColumn>
                        <dx:GridViewDataTextColumn Caption="AFFILIATE" FieldName="colAffiliate"
                            Name="colAffiliate" VisibleIndex="3" Width="0px">
                            <HeaderStyle Font-Names="Tahoma" Font-Size="8pt" Font-Underline="False" Wrap="True"
                                HorizontalAlign="Center" />
                            <CellStyle Font-Names="Tahoma" Font-Size="8pt">
                            </CellStyle>
                        </dx:GridViewDataTextColumn>
                        <dx:GridViewDataTextColumn Caption="CONTAINER" 
                            FieldName="colContainer" Name="colContainer"
                            VisibleIndex="4" Width="100px" FixedStyle="Left">
                            <HeaderStyle Font-Names="Tahoma" Font-Size="8pt" Font-Underline="False" Wrap="True"
                                HorizontalAlign="Center" />
                            <CellStyle Font-Names="Tahoma" Font-Size="8pt">
                            </CellStyle>
                        </dx:GridViewDataTextColumn>
                        <dx:GridViewDataTextColumn Caption="DELIVERY DATE" 
               FieldName="colDeliveryDate" Name="colDeliveryDate"
                            VisibleIndex="16" Width="129px">
                            <HeaderStyle Font-Names="Tahoma" Font-Size="8pt" Font-Underline="False" Wrap="True"
                                HorizontalAlign="Center" />
                            <CellStyle Font-Names="Tahoma" Font-Size="8pt">
                            </CellStyle>
                        </dx:GridViewDataTextColumn>
                        <dx:GridViewDataTextColumn Caption="PALLET NO" FieldName="colpallet"
                            Name="colpallet" Width="100px" VisibleIndex="5" 
               FixedStyle="Left">
                             <HeaderStyle Font-Names="Tahoma" Font-Size="8pt" Font-Underline="False" Wrap="True"
                                HorizontalAlign="Center" />
                            <CellStyle Font-Names="Tahoma" Font-Size="8pt">
                            </CellStyle>
                        </dx:GridViewDataTextColumn>
                        <dx:GridViewDataTextColumn Caption="ORDER NO" FieldName="colOrder" 
                            VisibleIndex="11" Name="colOrder" Width="120px">
                            <HeaderStyle Font-Names="Tahoma" Font-Size="8pt" Font-Underline="False" 
                                HorizontalAlign="Center" Wrap="True" />
                            <CellStyle Font-Names="Tahoma" Font-Size="8pt" HorizontalAlign="Right">
                            </CellStyle>
                        </dx:GridViewDataTextColumn>
                        <dx:GridViewDataTextColumn Caption="PART NO" FieldName="colPart" 
                            Name="colPart" VisibleIndex="12"
                            Width="100px">
                              <HeaderStyle Font-Names="Tahoma" Font-Size="8pt" Font-Underline="False" Wrap="True"
                                HorizontalAlign="Center" />
                            <CellStyle Font-Names="Tahoma" Font-Size="8pt">
                            </CellStyle>
                        </dx:GridViewDataTextColumn>
                        <dx:GridViewDataTextColumn Caption="BOX NO FROM" FieldName="colBoxFrom"
                            Name="colBoxFrom" VisibleIndex="13" Width="100px">
                            <HeaderStyle Font-Names="Tahoma" Font-Size="8pt" Font-Underline="False" Wrap="True"
                                HorizontalAlign="Center" />
                            <CellStyle Font-Names="Tahoma" Font-Size="8pt">
                            </CellStyle>
                        </dx:GridViewDataTextColumn>
                         <dx:GridViewDataTextColumn FieldName="colBoxTo" 
                            Name="colBoxTo" VisibleIndex="14" Width="100px" 
                            Caption="BOX NO TO">
                             <HeaderStyle Font-Names="Tahoma" Font-Size="8pt" Font-Underline="False" Wrap="True"
                                HorizontalAlign="Center" />
                            <CellStyle Font-Names="Tahoma" Font-Size="8pt">
                            </CellStyle>
                        </dx:GridViewDataTextColumn>
                        <dx:GridViewDataTextColumn Caption="LENGTH" FieldName="colLength"
                            Name="colLength" VisibleIndex="6" Width="90px">
                            <HeaderStyle Font-Names="Tahoma" Font-Size="8pt" Font-Underline="False" Wrap="True"
                                HorizontalAlign="Center" />
                            <CellStyle Font-Names="Tahoma" Font-Size="8pt" HorizontalAlign="Right">
                            </CellStyle>
                        </dx:GridViewDataTextColumn>
                         <dx:GridViewDataTextColumn FieldName="colWidth" 
                            Name="colWidth" VisibleIndex="7" Width="100px" 
               Caption="WIDTH" >
                            <HeaderStyle Font-Names="Tahoma" Font-Size="8pt" Font-Underline="False" Wrap="True"
                                HorizontalAlign="Center" />
                            <CellStyle Font-Names="Tahoma" Font-Size="8pt">
                            </CellStyle>
                        </dx:GridViewDataTextColumn>
                        <dx:GridViewDataTextColumn Caption="HEIGHT" FieldName="colHeight"
                            Name="colHeight" VisibleIndex="8" Width="100px">
                            <HeaderStyle Font-Names="Tahoma" Font-Size="8pt" Font-Underline="False" Wrap="True"
                                HorizontalAlign="Center" />
                            <CellStyle Font-Names="Tahoma" Font-Size="8pt" HorizontalAlign="Right">
                            </CellStyle>
                        </dx:GridViewDataTextColumn>
                        <dx:GridViewDataTextColumn FieldName="colM3" Name="colM3" 
                             VisibleIndex="9" Width="100px" Caption="M3">
                             <HeaderStyle Font-Names="Tahoma" Font-Size="8pt" Font-Underline="False" Wrap="True"
                                HorizontalAlign="Center" />
                            <CellStyle Font-Names="Tahoma" Font-Size="8pt" HorizontalAlign="Right">
                            </CellStyle>
                        </dx:GridViewDataTextColumn>
                        <dx:GridViewDataTextColumn Caption="HEIGHT PALLET" 
                            FieldName="colHeightPallet" Name="colHeightPallet"
                            VisibleIndex="10" Width="100px">
                            <HeaderStyle Font-Names="Tahoma" Font-Size="8pt" Font-Underline="False" Wrap="True"
                                HorizontalAlign="Center" />
                            <CellStyle Font-Names="Tahoma" Font-Size="8pt" HorizontalAlign="Right">
                            </CellStyle>
                        </dx:GridViewDataTextColumn>
                        <dx:GridViewDataTextColumn Caption="TOTAL BOX" FieldName="colTotalBox" 
                            Name="colTotalBox" VisibleIndex="22" Width="100px">
                            <HeaderStyle Font-Names="Tahoma" Font-Size="8pt" Font-Underline="False" Wrap="True"
                                HorizontalAlign="Center" />
                            <CellStyle Font-Names="Tahoma" Font-Size="8pt" HorizontalAlign="Right">
                            </CellStyle>
                        </dx:GridViewDataTextColumn>
                    <dx:GridViewDataTextColumn Caption="NoUrut" FieldName="NoUrut" 
               Name="NoUrut" VisibleIndex="27" Width="0px">
                <HeaderStyle Font-Names="Tahoma" Font-Size="8pt" Font-Underline="False" Wrap="True"
                                HorizontalAlign="Center" />
                            <CellStyle Font-Names="Tahoma" Font-Size="8pt">
                            </CellStyle>
           </dx:GridViewDataTextColumn>
                    <dx:GridViewDataTextColumn Caption="SURAT JALAN" 
               FieldName="colSuratJalan" Name="colSuratJalan" VisibleIndex="15" Width="150px">
                <HeaderStyle Font-Names="Tahoma" Font-Size="8pt" Font-Underline="False" Wrap="True"
                                HorizontalAlign="Center" />
                            <CellStyle Font-Names="Tahoma" Font-Size="8pt">
                            </CellStyle>
           </dx:GridViewDataTextColumn>
           <dx:GridViewDataTextColumn Caption="PIC" FieldName="colPIC" Name="colPIC" 
               VisibleIndex="17" Width="120px">
                <HeaderStyle Font-Names="Tahoma" Font-Size="8pt" Font-Underline="False" Wrap="True"
                                HorizontalAlign="Center" />
                            <CellStyle Font-Names="Tahoma" Font-Size="8pt">
                            </CellStyle>
           </dx:GridViewDataTextColumn>
           <dx:GridViewDataTextColumn Caption="JENIS ARMADA" FieldName="colJenisArmada" 
               Name="colJenisArmada" VisibleIndex="18" Width="120px">
                <HeaderStyle Font-Names="Tahoma" Font-Size="8pt" Font-Underline="False" Wrap="True"
                                HorizontalAlign="Center" />
                            <CellStyle Font-Names="Tahoma" Font-Size="8pt">
                            </CellStyle>
           </dx:GridViewDataTextColumn>
           <dx:GridViewDataTextColumn Caption="DRIVER NAME" FieldName="colDriverName" 
               Name="colDriverName" VisibleIndex="19" Width="120px">
                <HeaderStyle Font-Names="Tahoma" Font-Size="8pt" Font-Underline="False" Wrap="True"
                                HorizontalAlign="Center" />
                            <CellStyle Font-Names="Tahoma" Font-Size="8pt">
                            </CellStyle>
           </dx:GridViewDataTextColumn>
           <dx:GridViewDataTextColumn Caption="DRIVER CONTACT" 
               FieldName="colDriverContact" Name="colDriverContact" VisibleIndex="20" 
               Width="120px">
                <HeaderStyle Font-Names="Tahoma" Font-Size="8pt" Font-Underline="False" Wrap="True"
                                HorizontalAlign="Center" />
                            <CellStyle Font-Names="Tahoma" Font-Size="8pt">
                            </CellStyle>
           </dx:GridViewDataTextColumn>
           <dx:GridViewDataTextColumn Caption="NO POL." FieldName="colNoPol" 
               Name="colNoPol" VisibleIndex="21" Width="120px">
                <HeaderStyle Font-Names="Tahoma" Font-Size="8pt" Font-Underline="False" Wrap="True"
                                HorizontalAlign="Center" />
                            <CellStyle Font-Names="Tahoma" Font-Size="8pt">
                            </CellStyle>
           </dx:GridViewDataTextColumn>
           <dx:GridViewDataTextColumn Caption="TOTAL PALLET" FieldName="colTotalPallet" 
               Name="colTotalPallet" VisibleIndex="24" Width="100px">
                <HeaderStyle Font-Names="Tahoma" Font-Size="8pt" Font-Underline="False" Wrap="True"
                                HorizontalAlign="Center" />
                            <CellStyle Font-Names="Tahoma" Font-Size="8pt">
                            </CellStyle>
           </dx:GridViewDataTextColumn>
           <dx:GridViewDataTextColumn Caption="QTY" FieldName="colQTY" Name="colQTY" 
               VisibleIndex="25" Width="100px">
                <HeaderStyle Font-Names="Tahoma" Font-Size="8pt" Font-Underline="False" Wrap="True"
                                HorizontalAlign="Center" />
                            <CellStyle Font-Names="Tahoma" Font-Size="8pt">
                            </CellStyle>
           </dx:GridViewDataTextColumn>
           <dx:GridViewDataTextColumn Caption="TOTAL QTY" FieldName="colTotalQty" 
               Name="colTotalQty" VisibleIndex="26" Width="100px">
                <HeaderStyle Font-Names="Tahoma" Font-Size="8pt" Font-Underline="False" Wrap="True"
                                HorizontalAlign="Center" />
                            <CellStyle Font-Names="Tahoma" Font-Size="8pt">
                            </CellStyle>
           </dx:GridViewDataTextColumn>
                    <dx:GridViewDataTextColumn Caption="SUM TOTAL BOX" 
               FieldName="colSumTotalBox" Name="colSumTotalBox" VisibleIndex="23">
                <HeaderStyle Font-Names="Tahoma" Font-Size="8pt" Font-Underline="False" Wrap="True"
                                HorizontalAlign="Center" />
                            <CellStyle Font-Names="Tahoma" Font-Size="8pt">
                            </CellStyle>
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
                &nbsp;</td>
            <td >
                &nbsp;</td>
            <td>
                </td>
        </tr>
        <tr>
            <td align="left">
                <dx1:ASPxButton ID="btnsubmenu" runat="server" Text="SUB MENU" Width="90px" Font-Names="Tahoma"
                    Font-Size="8pt" TabIndex="20">
                </dx1:ASPxButton>
                </td>
            <td align="right" width="90">
                <dx1:ASPxButton ID="btnprint" runat="server" Text="PRINT" Width="90px"
                    Font-Names="Tahoma" Font-Size="8pt" AutoPostBack="False"
                    ClientInstanceName="btnprint" TabIndex="856">
                    <Paddings Padding="2px" />
                </dx1:ASPxButton>

                </td>
                <td align="left" width="90">
                <dx1:ASPxButton ID="btndelete" runat="server" Text="DELETE" Width="90px"
                    Font-Names="Tahoma" Font-Size="8pt" AutoPostBack="False" TabIndex="21"
                    ClientInstanceName="btndelete">  
                    <ClientSideEvents Click="btndeleteClick"/>
                    <Paddings Padding="2px" />
                </dx1:ASPxButton>
                </td>
            <td align="left" width="90">
            <dx1:ASPxButton ID="btnExcel" runat="server" AutoPostBack="False" ClientInstanceName="btnExcel" Width="90px"
                    Font-Names="Tahoma" Font-Size="8pt" Text="EXCEL" TabIndex="22">
                     <ClientSideEvents Click="function(s, e) {
                    
                        grid.PerformCallback('gridExcel');               
                    }" />
                     <Paddings Padding="2px" />
                </dx1:ASPxButton>

               </td>
            
        </tr>
    </table>
    <div style="height:8px;"></div>  
    <dx:ASPxCallback ID="SaveSubmit" runat="server" ClientInstanceName = "SaveSubmit">
        <ClientSideEvents EndCallback="function(s, e) {
            var pMsg = s.cpMessage;        
            if (pMsg) {
                if (s.cpType == 'error'){
                    lblerrmessage.GetMainElement().style.color = 'Red';
                }
                else if (s.cpType == 'info'){
                    lblerrmessage.GetMainElement().style.color = 'Blue';
                }
                else {
                    lblerrmessage.GetMainElement().style.color = 'Red';
                }
                lblerrmessage.SetText(pMsg);
        }}" />
    </dx:ASPxCallback>
    </asp:Content>

