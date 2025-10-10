<%@ Page Title="" Language="vb" AutoEventWireup="false" MasterPageFile="~/SiteContent.Master" CodeBehind="TallyDataForm.aspx.vb" Inherits="PASISystem.TallyDataForm" %>

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
                                        <dx1:ASPxLabel ID="ASPxLabel12" runat="server" Text="AFFILIATE">
                                        </dx1:ASPxLabel>
                                    </td>
                                    <td align="left" valign="middle" width="15px">
                                        </td>
                                    <td align="left" valign="middle" width="150px" class="style10">
                <dx1:ASPxComboBox ID="cboFaffiliate" runat="server" ClientInstanceName="cboFaffiliate"
                    Font-Names="Tahoma" TextFormatString="{0}" Font-Size="8pt">
                    <ClientSideEvents SelectedIndexChanged="function(s, e) {
	                                            txtFaffiliate.SetText(cboFaffiliate.GetSelectedItem().GetColumnText(1));
                                                }" />
                </dx1:ASPxComboBox>
                                    </td>
                                    <td align="left" valign="middle" class="style11"></td>
                                    <td align="left" valign="middle" width="120px" class="style10">
                <dx1:ASPxTextBox ID="txtFaffiliate" runat="server" Font-Names="Tahoma" Font-Size="8pt"
                    Width="170px" BackColor="Silver" ClientInstanceName="txtFaffiliate" ReadOnly="True">
                    <ClientSideEvents TextChanged="function(s, e) {
	lblerrmessage.SetText('');
}" />
                </dx1:ASPxTextBox>
                                    </td>
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
                                    <td align="left" valign="middle" height="25px" width="10px"></td>
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
                                            Text="INVOICE NO *" Width="145px">
                                        </dx1:ASPxLabel>
                                    </td>
                                    <td>
                                        &nbsp;</td>
                                    <td>
                                        <dx1:ASPxTextBox ID="txtInvoiceNo" runat="server" 
                                            ClientInstanceName="txtInvoiceNo" Font-Names="Tahoma" Font-Size="8pt" 
                                            Width="170px" TabIndex="9">
                                        </dx1:ASPxTextBox>
                                    </td>
                                </tr>
                                <tr>
                                    <td>
                                        <dx1:ASPxLabel ID="ASPxLabel101" runat="server" Font-Names="Tahoma" Font-Size="8pt"
                                            Text="BL/AWB NO" Width="145px">
                                        </dx1:ASPxLabel>
                                    </td>
                                    <td>
                                        &nbsp;</td>
                                    <td>
                                        <dx1:ASPxTextBox ID="txtBLAWBNo" runat="server" 
                                            ClientInstanceName="txtBLAWBNo" Font-Names="Tahoma" Font-Size="8pt" 
                                            Width="170px" TabIndex="9">
                                        </dx1:ASPxTextBox>
                                    </td>
                                </tr>
                                <tr>
                                    <td class="style7">
                                        <dx1:ASPxLabel ID="ASPxLabel102" runat="server" Font-Names="Tahoma" 
                                            Font-Size="8pt" Text="BL/AWB DATE" Width="145px">
                                        </dx1:ASPxLabel>
                                    </td>
                                    <td class="style7">
                                        &nbsp;</td>
                                    <td class="style7">
                                        <dx1:ASPxDateEdit ID="dtBLAWBDate" runat="server" 
                                            ClientInstanceName="dtBLAWBDate" EditFormat="Custom" TabIndex="17"
                                            EditFormatString="dd MMM yyyy" Font-Size="8pt">
                                        </dx1:ASPxDateEdit>
                                    </td>
                                </tr>
                                 <tr>
                                    <td>
                                        <dx1:ASPxLabel ID="ASPxLabel1" runat="server" Font-Names="Tahoma" Font-Size="8pt"
                                            Text="PEB NO" Width="145px">
                                        </dx1:ASPxLabel>
                                    </td>
                                    <td>
                                        &nbsp;</td>
                                    <td>
                                        <dx1:ASPxTextBox ID="txtPEBNo" runat="server" 
                                            ClientInstanceName="txtPEBNo" Font-Names="Tahoma" Font-Size="8pt" 
                                            Width="170px" TabIndex="9">
                                        </dx1:ASPxTextBox>
                                    </td>
                                </tr>
                                <tr>
                                    <td class="style7">
                                        <dx1:ASPxLabel ID="ASPxLabel2" runat="server" Font-Names="Tahoma" 
                                            Font-Size="8pt" Text="PEB DATE" Width="145px">
                                        </dx1:ASPxLabel>
                                    </td>
                                    <td class="style7">
                                        &nbsp;</td>
                                    <td class="style7">
                                        <dx1:ASPxDateEdit ID="dtPEBDate" runat="server" 
                                            ClientInstanceName="dtPEBDate" EditFormat="Custom" TabIndex="17"
                                            EditFormatString="dd MMM yyyy" Font-Size="8pt">
                                        </dx1:ASPxDateEdit>
                                    </td>
                                </tr>
                                <tr>
                                    <td>
                                        <dx1:ASPxLabel ID="ASPxLabel3" runat="server" Font-Names="Tahoma" Font-Size="8pt"
                                            Text="TYPE" Width="145px">
                                        </dx1:ASPxLabel>
                                    </td>
                                    <td>
                                        &nbsp;</td>
                                    <td>
                                        <dx1:ASPxTextBox ID="txtType" runat="server" 
                                            ClientInstanceName="txtType" Font-Names="Tahoma" Font-Size="8pt" 
                                            Width="170px" TabIndex="9">
                                        </dx1:ASPxTextBox>
                                    </td>
                                </tr>
                                <tr>
                                    <td class="style7">
                                        <dx1:ASPxLabel ID="ASPxLabel24" runat="server" Font-Names="Tahoma" Font-Size="8pt"
                                            Text="CONTAINER NO" Width="200px">
                                        </dx1:ASPxLabel>
                                    </td>
                                    <td class="style7">
                                        &nbsp;</td>
                                    <td class="style7">
                                        <dx1:ASPxTextBox ID="txtContainerNo" runat="server" 
                                            ClientInstanceName="txtContainerNo" Font-Names="Tahoma" Font-Size="8pt" 
                                            Width="170px" TabIndex="10">
                                        </dx1:ASPxTextBox>
                                    </td>
                                </tr>
                                 <tr>
                                    <td class="style7">
                                        <dx1:ASPxLabel ID="ASPxLabel29" runat="server" Font-Names="Tahoma" 
                                            Font-Size="8pt" Text="SIZE CONTAINER" Width="145px" >
                                        </dx1:ASPxLabel>
                                    </td>
                                    <td class="style7">
                                        &nbsp;</td>
                                    <td class="style7">
                                        <dx1:ASPxTextBox ID="txtSizeContainer" runat="server" 
                                            ClientInstanceName="txtSizeContainer" Font-Names="Tahoma" Font-Size="8pt" TabIndex="14"
                                            Width="170px" >
                                        </dx1:ASPxTextBox>
                                    </td>
                                </tr>
                                <tr>
                                    <td class="style7">
                                        <dx1:ASPxLabel ID="ASPxLabel26" runat="server" Font-Names="Tahoma" 
                                            Font-Size="8pt" Text="SEAL NO. " Width="145px">
                                        </dx1:ASPxLabel>
                                    </td>
                                    <td class="style7">
                                        &nbsp;</td>
                                    <td class="style7">
                                        <dx1:ASPxTextBox ID="txtSealNo" runat="server" ClientInstanceName="txtSealNo" 
                                            Font-Names="Tahoma" Font-Size="8pt" Width="170px" TabIndex="11">
                                        </dx1:ASPxTextBox>
                                    </td>
                                </tr>
                                  <tr>
                                      <td class="style7">
                                          <dx1:ASPxLabel ID="ASPxLabel27" runat="server" Font-Names="Tahoma" 
                                              Font-Size="8pt" Text="TARE" Width="145px" >
                                          </dx1:ASPxLabel>
                                      </td>
                                      <td class="style7">
                                          &nbsp;</td>
                                      <td class="style7">
                                          <dx1:ASPxTextBox ID="txtTare" runat="server" ClientInstanceName="txtTare" HorizontalAlign="Right"
                                              Font-Names="Tahoma" Font-Size="8pt" Width="170px" TabIndex="12" DisplayFormatString="#,##0.0000" MaskSettings-Mask="<0..9999999999999g>.<0..9999g>" MaskSettings-IncludeLiterals="DecimalSymbol" >
                                           
                                          </dx1:ASPxTextBox>


                                      </td>
                                </tr>
                                <tr>
                                    <td class="style7">
                                        <dx1:ASPxLabel ID="ASPxLabel28" runat="server" Font-Names="Tahoma" 
                                            Font-Size="8pt" Text="VESSEL NAME/VOYAGE" Width="145px" >
                                        </dx1:ASPxLabel>
                                    </td>
                                    <td class="style7">
                                        &nbsp;</td>
                                    <td class="style7">
                                        <dx1:ASPxTextBox ID="txtVesselNo" runat="server" 
                                            ClientInstanceName="txtVesselNo" Font-Names="Tahoma" Font-Size="8pt" TabIndex="13"
                                            Width="170px">
                                        </dx1:ASPxTextBox>
                                    </td>
                                </tr>
                               
                                <tr>
                                    <td class="style7">
                                        <dx1:ASPxLabel ID="ASPxLabel30" runat="server" Font-Names="Tahoma" 
                                            Font-Size="8pt" Text="SHIPPING LINE" Width="145px">
                                        </dx1:ASPxLabel>
                                    </td>
                                    <td class="style7">
                                    </td>
                                    <td class="style7">
                                        <dx1:ASPxTextBox ID="txtShippingLine" runat="server" 
                                            ClientInstanceName="txtShippingLine" Font-Names="Tahoma" Font-Size="8pt" TabIndex="15"
                                            Width="170px" >
                                        </dx1:ASPxTextBox>
                                    </td>
                                </tr>
                               <%-- <tr>
                                    <td class="style7">
                                        <dx1:ASPxLabel ID="ASPxLabel31" runat="server" Font-Names="Tahoma" 
                                            Font-Size="8pt" Text="VESSEL NAME" Width="145px" >
                                        </dx1:ASPxLabel>
                                    </td>
                                    <td class="style7">
                                        &nbsp;</td>
                                    <td class="style7">
                                        <dx1:ASPxTextBox ID="txtVesselName" runat="server" 
                                            ClientInstanceName="txtVesselName" Font-Names="Tahoma" Font-Size="8pt" TabIndex="16"
                                            Width="170px">
                                        </dx1:ASPxTextBox>
                                    </td>
                                </tr>--%>
                                <tr>
                                    <td class="style7">
                                        <dx1:ASPxLabel ID="ASPxLabel32" runat="server" Font-Names="Tahoma" 
                                            Font-Size="8pt" Text="STUFFING DATE" Width="145px">
                                        </dx1:ASPxLabel>
                                    </td>
                                    <td class="style7">
                                        &nbsp;</td>
                                    <td class="style7">
                                        <dx1:ASPxDateEdit ID="dtStuffingDate" runat="server" 
                                            ClientInstanceName="dtStuffingDate" EditFormat="Custom" TabIndex="17"
                                            EditFormatString="dd MMM yyyy" Font-Size="8pt">
                                        </dx1:ASPxDateEdit>
                                    </td>
                                </tr>
                                <tr>
                                    <td class="style7">
                                        <dx1:ASPxLabel ID="ASPxLabel25" runat="server" Font-Names="Tahoma" 
                                            Font-Size="8pt" Text="PORT OF LOADING" Width="145px">
                                        </dx1:ASPxLabel>
                                    </td>
                                    <td class="style7">
                                        &nbsp;</td>
                                    <td class="style7">
                                        <dx1:ASPxTextBox ID="txtLocation" runat="server" 
                                            ClientInstanceName="txtLocation" Font-Names="Tahoma" Font-Size="8pt" TabIndex="18"
                                            Width="170px">
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
                                        <dx1:ASPxLabel ID="ASPxLabel14" runat="server" Text="INVOICE">
                                        </dx1:ASPxLabel>
                                    </td>
                                    <td align="left" valign="middle" height="25px" width="15px">
                                        &nbsp;</td>
                                    <td align="left" valign="middle" height="25px" width="150px">
                <dx1:ASPxTextBox ID="txtFInvoiceNo" runat="server" Font-Names="Tahoma" Font-Size="8pt"
                    Width="170px" ClientInstanceName="txtFInvoiceNo" TabIndex="2">
                </dx1:ASPxTextBox>
                                    </td>
                                    <td align="left" valign="middle" height="25px" width="10px">
                                        &nbsp;</td>
                                    <td align="left" valign="middle" height="25px" width="120px">
                                        <dx1:ASPxComboBox runat="server" Width="100px" ClientInstanceName="cboCreate" 
                                            ID="cboCreate">
                                              <ClientSideEvents SelectedIndexChanged="function(s, e) {
	                                                                if (cboCreate.GetText() == 'CREATE') {
	                                                                checkboxdt.SetEnabled(false);
                                                                    txtFContainerNo.SetText('')
                                                                    txtFContainerNo.SetEnabled(false);
                                                                    txtFPalletNo.SetText('')
                                                                    txtFPalletNo.SetEnabled(false);
                                                                    dtFStuffingDateFrom.SetEnabled(false);
                                                                    dtFStuffingDateEnd.SetEnabled(false);
                                                                } else {
                                                                    checkboxdt.SetEnabled(true);
                                                                    txtFContainerNo.SetEnabled(true);
                                                                    txtFPalletNo.SetEnabled(true);
                                                                    dtFStuffingDateFrom.SetEnabled(true);
                                                                    dtFStuffingDateEnd.SetEnabled(true);
                                                       }
                                            }" />
                                            </dx1:ASPxComboBox>

                                    </td>
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
                                        <dx1:ASPxLabel ID="ASPxLabel21" runat="server" Text="PALLET NO">
                                        </dx1:ASPxLabel>
                                    </td>
                                    <td align="left" valign="middle" height="25px" width="15px">
                                        &nbsp;</td>
                                    <td align="left" valign="middle" height="25px" width="150px">
                <dx1:ASPxTextBox ID="txtFPalletNo" runat="server" Font-Names="Tahoma" Font-Size="8pt"
                    Width="170px" ClientInstanceName="txtFPalletNo" TabIndex="4">
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
                                        <dx1:ASPxLabel ID="ASPxLabel17" runat="server" Text="STUFFING DATE">
                                        </dx1:ASPxLabel>
                                    </td>
                                    <td align="right" valign="middle" height="25px" width="15px">
                <dx1:ASPxCheckBox ID="checkboxdt" runat="server" CheckState="Checked" ClientInstanceName="checkboxdt"
                    Text=" ">
                    <ClientSideEvents CheckedChanged="function (s, e) {
                                                                if (s.GetChecked()==true) {
                                                                    dtFStuffingDateFrom.SetEnabled(true);
                                                                    dtFStuffingDateEnd.SetEnabled(true);
                                                                } else {
                                                                    dtFStuffingDateFrom.SetEnabled(false);
                                                                    dtFStuffingDateEnd.SetEnabled(false);
                                                                }
                                                          }" />
                </dx1:ASPxCheckBox>
                                    </td>
                                    <td align="left" valign="middle" height="25px" width="150px">
                <dx1:ASPxDateEdit ID="dtFStuffingDateFrom" runat="server" ClientInstanceName="dtFStuffingDateFrom" Font-Size="8pt"
                    EditFormat="Custom" EditFormatString="dd MMM yyyy" TabIndex="5">
                </dx1:ASPxDateEdit>
                                    </td>
                                    <td align="center" valign="middle" height="25px" width="10px">
                                        <dx1:ASPxLabel ID="ASPxLabel22" runat="server" Text="   TO  " 
                                            Font-Size="X-Small">
                                        </dx1:ASPxLabel>
                                    </td>
                                    <td align="left" valign="middle" height="25px" width="120px">
                <dx1:ASPxDateEdit ID="dtFStuffingDateEnd" runat="server" ClientInstanceName="dtFStuffingDateEnd" Font-Size="8pt"
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
                                    <td align="right" valign="middle" height="25px" width="150px">
                            <dx1:ASPxButton ID="btnsearch" runat="server" Text="SEARCH" Width="85px" Font-Names="Tahoma"
                                Font-Size="8pt" ClientInstanceName="btnsearch" AutoPostBack="False" TabIndex="7">
                                <ClientSideEvents Click="function(s, e) {
                                            var pAffiliate = cboFaffiliate.GetText();
                                            var pForwarder = cboFForwarder.GetText();
                                            var pInvoice = txtFInvoiceNo.GetText();
                                            var pContainer = txtFContainerNo.GetText();
                                            var pPaletNo = txtFPalletNo.GetText();
                                            var pStuffingDateFrom = dtFStuffingDateFrom.GetText();
                                            var pStuffingDateEnd = dtFStuffingDateEnd.GetText();
                                          
	                                        grid.PerformCallback('gridload' + '|' + pAffiliate + '|' + pForwarder + '|' + pInvoice + '|' + pContainer + '|' + pPaletNo + '|' + pStuffingDateFrom + '|' + pStuffingDateEnd);


	                                        lblerrmessage.SetText('');

txtInvoiceNo.SetText('');
             txtContainerNo.SetText('');
             txtSealNo.SetText('');
             txtTare.SetText('');
             txtVesselNo.SetText('');
             txtSizeContainer.SetText('');
             txtShippingLine.SetText('');
             txtLocation.SetText('');
             txtBLAWBNo.SetText('');
             txtPEBNo.SetText('');
             txtType.SetText('');

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

             //             if (txtLocation.GetText() == "") {
             //                 lblerrmessage.SetText("[9999] Please Input Port of Loading first!");
             //                 txtLocation.Focus();
             //                 valid = 'false';
             //             }

             //             if (dtStuffingDate.GetText() == "") {
             //                 lblerrmessage.SetText("[9999] Please Input Stuffing Date first!");
             //                 dtStuffingDate.Focus();
             //                 valid = 'false';
             //             }


             //             if (txtShippingLine.GetText() == "") {
             //                 lblerrmessage.SetText("[7017] Please Input Shipping Line first!");
             //                 txtShippingLine.Focus();
             //                 valid = 'false';
             //             }


             //             if (txtVesselNo.GetText() == "") {
             //                 lblerrmessage.SetText("[9999] Please Input Vessel Name / VOY first!");
             //                 txtVesselNo.Focus();
             //                 valid = 'false';
             //             }

             //             if (txtTare.GetText() == "") {
             //                 lblerrmessage.SetText("[9999] Please Input Tare first!");
             //                 txtTare.Focus();
             //                 valid = 'false';
             //             }

             //             if (txtSealNo.GetText() == "") {
             //                 lblerrmessage.SetText("[9999] Please Input Seal No first!");
             //                 txtSealNo.Focus();
             //                 valid = 'false';
             //             }

             //             if (txtSizeContainer.GetText() == "") {
             //                 lblerrmessage.SetText("[9999] Please Input Size Container first!");
             //                 txtSizeContainer.Focus();
             //                 valid = 'false';
             //             }

             if (txtContainerNo.GetText() == "") {
                 lblerrmessage.SetText("[9999] Please Input Container No first!");
                 txtContainerNo.Focus();
                 valid = 'false';
             }

             //             if (txtType.GetText() == "") {
             //                 lblerrmessage.SetText("[9999] Please Input Type first!");
             //                 txtType.Focus();
             //                 valid = 'false';
             //             }

             //             if (dtPEBDate.GetText() == "") {
             //                 lblerrmessage.SetText("[9999] Please Input PEB Date first!");
             //                 dtPEBDate.Focus();
             //                 valid = 'false';
             //             }

             //             if (txtPEBNo.GetText() == "") {
             //                 lblerrmessage.SetText("[9999] Please Input PEB No first!");
             //                 txtPEBNo.Focus();
             //                 valid = 'false';
             //            }

             //             if (dtBLAWBDate.GetText() == "") {
             //                 lblerrmessage.SetText("[9999] Please Input BL/AWB Date first!");
             //                 dtBLAWBDate.Focus();
             //                 valid = 'false';
             //            }

             //             if (txtBLAWBNo.GetText() == "") {
             //                 lblerrmessage.SetText("[9999] Please Input BL/AWB No first!");
             //                 txtBLAWBNo.Focus();
             //                 valid = 'false';
             //             }

             if (txtInvoiceNo.GetText() == "") {
                 lblerrmessage.SetText("[9999] Please Input Invoice No first!");
                 txtInvoiceNo.Focus();
                 valid = 'false';
             }

//             if (cboFForwarder.GetText() == "" || cboFForwarder.GetText() == "== ALL ==") {
//                 lblerrmessage.SetText("[7012] Please Input Forwarder first!");
//                 cboFForwarder.Focus();
//                 valid = 'false';
//             }

//             if (cboFaffiliate.GetText() == "" || cboFaffiliate.GetText() == "== ALL ==") {
//                 lblerrmessage.SetText("[7011] Please Input Affiliate first!");
//                 cboFaffiliate.Focus();
//                 valid = 'false';
//             }
         }
         function up_Insert() {
             var pIsUpdate = '';
             var pInvoiceNo = txtInvoiceNo.GetText();
             var pForwarder = cboFForwarder.GetText();
             var pAffiliate = cboFaffiliate.GetText();
             var pContainerNo = txtContainerNo.GetText();
             var pSealNo = txtSealNo.GetText();
             var pTare = txtTare.GetText();
             var pVesselNo = txtVesselNo.GetText();
             var pSizeContainer = txtSizeContainer.GetText();
             var pShippingLine = txtShippingLine.GetText();
             var pVesselName = "";
             var pStuffingDate = dtStuffingDate.GetText();
             var pLocation = txtLocation.GetText();

             var pBLAWBNo = txtBLAWBNo.GetText();
             var pBLAWBDate = dtBLAWBDate.GetText();
             var pPEBNo = txtPEBNo.GetText();
             var pPEBDate = dtPEBDate.GetText();
             var pType = txtType.GetText();


             SaveSubmit.PerformCallback('save|' + pIsUpdate + '|' + pInvoiceNo + '|' + pForwarder + '|' + pAffiliate + '|' + pContainerNo + '|' + pSealNo + '|' + pTare + '|' + pVesselNo + '|' + pSizeContainer + '|' + pShippingLine + '|' + pVesselName + '|' + pStuffingDate + '|' + pLocation + '|' + pBLAWBNo + '|' + pBLAWBDate + '|' + pPEBNo + '|' + pPEBDate + '|' + pType);

             var pAffiliate1 = cboFaffiliate.GetText();
             var pForwarder1 = cboFForwarder.GetText();
             var pInvoice1 = txtFInvoiceNo.GetText();
             var pContainer1 = txtFContainerNo.GetText();
             var pPaletNo1 = txtFPalletNo.GetText();
             var pStuffingDateFrom1 = dtFStuffingDateFrom.GetText();
             var pStuffingDateEnd1 = dtFStuffingDateEnd.GetText();

             cboCreate.SetText('SEARCH');

             grid.PerformCallback('gridload' + '|' + pAffiliate1 + '|' + pForwarder1 + '|' + pInvoice1 + '|' + pContainer1 + '|' + pPaletNo1 + '|' + pStuffingDateFrom1 + '|' + pStuffingDateEnd1);

        
         }

         function OnGridFocusedRowChangedTallyData() {
             grid.GetRowValues(grid.GetFocusedRowIndex(), 'colAffiliate;colForwarder;colInvoice;colContainer;colSeal;colTare;colVesselNo;colContainerSize;colShippingLine;colVesselName;colStuffing;colDestination;colBLAWBNo;colBLAWBDate;colPEBNo;colPEBDate;colType', OnGetRowValuesTallyData);
         }

         function OnGetRowValuesTallyData(values) {
             cboFaffiliate.SetText(values[0]);
             cboFForwarder.SetText(values[1]);
             txtFaffiliate.SetText(values[0]);
             txtFForwarder.SetText(values[1]);
             txtInvoiceNo.SetText(values[2].trim());
             txtContainerNo.SetText(values[3].trim());
             txtSealNo.SetText(values[4].trim());
             txtTare.SetText(values[5].toString());
             txtVesselNo.SetText(values[6].trim());
             txtSizeContainer.SetText(values[7].trim());
             txtShippingLine.SetText(values[8].trim());
             //             txtVesselName.SetText(values[9].trim());
             dtStuffingDate.SetText(values[10]);
             txtLocation.SetText(values[11].trim());

             txtBLAWBNo.SetText(values[12].toString());
             dtBLAWBDate.SetText(values[13]);
             txtPEBNo.SetText(values[14].toString());
             dtPEBDate.SetText(values[15]);
             txtType.SetText(values[16].toString());

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
                    KeyFieldName="colInvoice;colForwarder;colAffiliate;colpallet;colOrder;colPart;colBoxFrom;colBoxTo;colContainer;colSeal;NoUrut;colLength;colWidth;colHeight;colM3;colHeightPallet;colTotalBox"
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
	dtFStuffingDateFrom.SetText(s.cpdtFStuffingDateFrom);
    dtFStuffingDateEnd.SetText(s.cpdtFStuffingDateEnd);
    dtStuffingDate.SetText(s.cpdtStuffingDate);
    dtBLAWBDate.SetText(s.cpdtBLAWBDate);
    dtPEBDate.SetText(s.cpdtPEBDate);

}" CallbackError="function(s, e) {
	e.handled = true;
}" 
   RowDblClick="function(s, e) {OnGridFocusedRowChangedTallyData();}" 
   />
  <Columns>
           <dx:GridViewCommandColumn ShowSelectCheckbox="True"
                    ShowClearFilterButton="true" VisibleIndex="0" SelectAllCheckboxMode="Page" 
                    Width="30px" FixedStyle="Left" Name="ACT" />
                        <dx:GridViewDataTextColumn Caption="INVOICE NO" 
                               FieldName="colInvoice" Name="colInvoice"
                               VisibleIndex="1" Width="90px" FixedStyle="Left">
                            <HeaderStyle Font-Names="Tahoma" Font-Size="8pt" Font-Underline="False" Wrap="True"
                                HorizontalAlign="Center" />
                            <CellStyle Font-Names="Tahoma" Font-Size="8pt">
                            </CellStyle>
                        </dx:GridViewDataTextColumn>
                        <dx:GridViewDataTextColumn Caption="FORWARDER" FieldName="colForwarder"
                            Name="colForwarder" VisibleIndex="2" Width="80px" 
                            FixedStyle="Left">
                            <HeaderStyle Font-Names="Tahoma" Font-Size="8pt" Font-Underline="False" Wrap="True"
                                HorizontalAlign="Center" />
                            <CellStyle Font-Names="Tahoma" Font-Size="8pt">
                            </CellStyle>
                        </dx:GridViewDataTextColumn>
                        <dx:GridViewDataTextColumn Caption="AFFILIATE" FieldName="colAffiliate"
                            Name="colAffiliate" VisibleIndex="3" Width="70px" 
                            FixedStyle="Left">
                            <HeaderStyle Font-Names="Tahoma" Font-Size="8pt" Font-Underline="False" Wrap="True"
                                HorizontalAlign="Center" />
                            <CellStyle Font-Names="Tahoma" Font-Size="8pt">
                            </CellStyle>
                        </dx:GridViewDataTextColumn>
                        <dx:GridViewDataTextColumn Caption="CONTAINER" 
                            FieldName="colContainer" Name="colContainer"
                            VisibleIndex="9" Width="100px">
                            <HeaderStyle Font-Names="Tahoma" Font-Size="8pt" Font-Underline="False" Wrap="True"
                                HorizontalAlign="Center" />
                            <CellStyle Font-Names="Tahoma" Font-Size="8pt">
                            </CellStyle>
                        </dx:GridViewDataTextColumn>
                        <dx:GridViewDataTextColumn Caption="SEAL NO" FieldName="colSeal"
                            Name="colSeal" Width="100px" VisibleIndex="10">
                             <HeaderStyle Font-Names="Tahoma" Font-Size="8pt" Font-Underline="False" Wrap="True"
                                HorizontalAlign="Center" />
                            <CellStyle Font-Names="Tahoma" Font-Size="8pt">
                            </CellStyle>
                        </dx:GridViewDataTextColumn>
                        <dx:GridViewDataTextColumn Caption="TARE" FieldName="colTare" 
                            Name="colTare" VisibleIndex="11"
                            Width="80px">
                             <HeaderStyle Font-Names="Tahoma" Font-Size="8pt" Font-Underline="False" Wrap="True"
                                HorizontalAlign="Center" />
                            <CellStyle Font-Names="Tahoma" Font-Size="8pt" HorizontalAlign="Right">
                            </CellStyle>
                            <PropertiesTextEdit DisplayFormatString="#,##0.0000">
                           <%-- <MaskSettings IncludeLiterals = "DecimalSymbol" Mask = "<0..9999999999999g>.<0..9999g>" />--%>
                            </PropertiesTextEdit>
                        </dx:GridViewDataTextColumn>
                        <dx:GridViewDataTextColumn Caption="GROSS" FieldName="colGross" Name="colGross"
                            VisibleIndex="12" Width="80px">
                            <HeaderStyle Font-Names="Tahoma" Font-Size="8pt" Font-Underline="False" Wrap="True"
                                HorizontalAlign="Center" />
                            <CellStyle Font-Names="Tahoma" Font-Size="8pt" HorizontalAlign="Right">
                            </CellStyle>
                        </dx:GridViewDataTextColumn>
                        <dx:GridViewDataTextColumn Caption="VESSEL NO" FieldName="colVesselNo"
                            Name="colVesselNo" VisibleIndex="13" Width="150px">
                            <HeaderStyle Font-Names="Tahoma" Font-Size="8pt" Font-Underline="False" Wrap="True"
                                HorizontalAlign="Center" />
                            <CellStyle Font-Names="Tahoma" Font-Size="8pt">
                            </CellStyle>
                        </dx:GridViewDataTextColumn>
                         <dx:GridViewDataTextColumn FieldName="colVesselName" 
                            Name="colVesselName" VisibleIndex="14" Width="180px" 
                            Caption="VESSEL NAME">
                             <HeaderStyle Font-Names="Tahoma" Font-Size="8pt" Font-Underline="False" Wrap="True"
                                HorizontalAlign="Center" />
                            <CellStyle Font-Names="Tahoma" Font-Size="8pt">
                            </CellStyle>
                        </dx:GridViewDataTextColumn>
                         <dx:GridViewDataTextColumn FieldName="colContainerSize" 
                            Name="colContainerSize" VisibleIndex="15" Width="100px" 
               Caption="CONTAINER SIZE" >
                            <HeaderStyle Font-Names="Tahoma" Font-Size="8pt" Font-Underline="False" Wrap="True"
                                HorizontalAlign="Center" />
                            <CellStyle Font-Names="Tahoma" Font-Size="8pt">
                            </CellStyle>
                        </dx:GridViewDataTextColumn>
                        <dx:GridViewDataTextColumn FieldName="colETD" Name="colETD" 
                             VisibleIndex="17" Width="140px" Caption="ETD PORT">
                             <HeaderStyle Font-Names="Tahoma" Font-Size="8pt" Font-Underline="False" Wrap="True"
                                HorizontalAlign="Center" />
                            <CellStyle Font-Names="Tahoma" Font-Size="8pt">
                            </CellStyle>
                        </dx:GridViewDataTextColumn>
                        <dx:GridViewDataTextColumn Caption="SHIPPING LINE" FieldName="colShippingLine"
                            Name="colShippingLine" VisibleIndex="18" Width="110px">
                            <HeaderStyle Font-Names="Tahoma" Font-Size="8pt" Font-Underline="False" Wrap="True"
                                HorizontalAlign="Center" />
                            <CellStyle Font-Names="Tahoma" Font-Size="8pt">
                            </CellStyle>
                        </dx:GridViewDataTextColumn>
                        <dx:GridViewDataTextColumn Caption="DESTINATION PORT" FieldName="colDestination" 
                            VisibleIndex="19" Name="colDestination" Width="120px">
                            <CellStyle Font-Names="Tahoma" Font-Size="8pt">
                            </CellStyle>
                        </dx:GridViewDataTextColumn>
                        <dx:GridViewDataTextColumn Caption="STUFFING DATE" FieldName="colStuffing"
                            Name="colStuffing" VisibleIndex="16" Width="129px">
                            <HeaderStyle Font-Names="Tahoma" Font-Size="8pt" Font-Underline="False" Wrap="True"
                                HorizontalAlign="Center" />
                            <CellStyle Font-Names="Tahoma" Font-Size="8pt">
                            </CellStyle>
                        </dx:GridViewDataTextColumn>

                          <dx:GridViewDataTextColumn Caption="BL/AWB NO" FieldName="colBLAWBNo" 
                            VisibleIndex="27" Name="colBLAWBNo" Width="120px">
                            <CellStyle Font-Names="Tahoma" Font-Size="8pt">
                            </CellStyle>
                        </dx:GridViewDataTextColumn>
                        <dx:GridViewDataTextColumn Caption="BL/AWB DATE" FieldName="colBLAWBDate"
                            Name="colBLAWBDate" VisibleIndex="28" Width="129px">
                            <HeaderStyle Font-Names="Tahoma" Font-Size="8pt" Font-Underline="False" Wrap="True"
                                HorizontalAlign="Center" />
                            <CellStyle Font-Names="Tahoma" Font-Size="8pt">
                            </CellStyle>
                        </dx:GridViewDataTextColumn>
                            <dx:GridViewDataTextColumn Caption="PEB NO" FieldName="colPEBNo" 
                            VisibleIndex="29" Name="colPEBNo" Width="120px">
                            <CellStyle Font-Names="Tahoma" Font-Size="8pt">
                            </CellStyle>
                        </dx:GridViewDataTextColumn>
                        <dx:GridViewDataTextColumn Caption="PEB DATE" FieldName="colPEBDate"
                            Name="colPEBDate" VisibleIndex="30" Width="129px">
                            <HeaderStyle Font-Names="Tahoma" Font-Size="8pt" Font-Underline="False" Wrap="True"
                                HorizontalAlign="Center" />
                            <CellStyle Font-Names="Tahoma" Font-Size="8pt">
                            </CellStyle>
                        </dx:GridViewDataTextColumn>
                         <dx:GridViewDataTextColumn Caption="TYPE" FieldName="colType" 
                            VisibleIndex="31" Name="colType" Width="120px">
                            <CellStyle Font-Names="Tahoma" Font-Size="8pt">
                            </CellStyle>
                        </dx:GridViewDataTextColumn>


                        <dx:GridViewDataTextColumn Caption="PALLET NO" 
                            FieldName="colpallet" Name="colpallet"
                            VisibleIndex="4" Width="120px" FixedStyle="Left">
                            <HeaderStyle Font-Names="Tahoma" Font-Size="8pt" Font-Underline="False" Wrap="True"
                                HorizontalAlign="Center" />
                        </dx:GridViewDataTextColumn>
                        <dx:GridViewDataTextColumn Caption="ORDER NO" FieldName="colOrder" 
                            Name="colOrder" VisibleIndex="5" Width="160px">
                             <HeaderStyle Font-Names="Tahoma" Font-Size="8pt" Font-Underline="False" Wrap="True"
                                HorizontalAlign="Center" />
                            <CellStyle Font-Names="Tahoma" Font-Size="8pt">
                            </CellStyle>
                       </dx:GridViewDataTextColumn>
                       <dx:GridViewDataTextColumn Caption="PART NO" FieldName="colPart" Name="colPart" 
                            VisibleIndex="6" Width="100px">
                            <HeaderStyle Font-Names="Tahoma" Font-Size="8pt" Font-Underline="False" Wrap="True"
                                HorizontalAlign="Center" />
                            <CellStyle Font-Names="Tahoma" Font-Size="8pt">
                            </CellStyle>
                       </dx:GridViewDataTextColumn>
                       <dx:GridViewDataTextColumn Caption="BOX NO FROM" FieldName="colBoxFrom" 
                            Name="colBoxFrom" VisibleIndex="7" Width="100px">
                            <HeaderStyle Font-Names="Tahoma" Font-Size="8pt" Font-Underline="False" Wrap="True"
                                HorizontalAlign="Center" />
                            <CellStyle Font-Names="Tahoma" Font-Size="8pt">
                            </CellStyle>
                       </dx:GridViewDataTextColumn>
                       <dx:GridViewDataTextColumn Caption="BOX NO TO" FieldName="colBoxTo" 
                            Name="colBoxTo" VisibleIndex="8" Width="100px">
                            <HeaderStyle Font-Names="Tahoma" Font-Size="8pt" Font-Underline="False" Wrap="True"
                                HorizontalAlign="Center" />
                            <CellStyle Font-Names="Tahoma" Font-Size="8pt">
                            </CellStyle>
                       </dx:GridViewDataTextColumn>
                       <dx:GridViewDataTextColumn Caption="LENGTH" FieldName="colLength" 
                            Name="colLength" VisibleIndex="20" Width="100px">
                            <HeaderStyle Font-Names="Tahoma" Font-Size="8pt" Font-Underline="False" Wrap="True"
                                HorizontalAlign="Center" />
                            <CellStyle Font-Names="Tahoma" Font-Size="8pt" HorizontalAlign="Right">
                            </CellStyle>
                       </dx:GridViewDataTextColumn>
                       <dx:GridViewDataTextColumn Caption="WIDTH" FieldName="colWidth" Name="colWidth" 
                            VisibleIndex="21" Width="100px">
                            <HeaderStyle Font-Names="Tahoma" Font-Size="8pt" Font-Underline="False" Wrap="True"
                                HorizontalAlign="Center" />
                            <CellStyle Font-Names="Tahoma" Font-Size="8pt" HorizontalAlign="Right">
                                        </CellStyle>
                       </dx:GridViewDataTextColumn>
                       <dx:GridViewDataTextColumn Caption="HEIGHT" FieldName="colHeight" 
                            Name="colHeight" VisibleIndex="22" Width="100px">
                            <HeaderStyle Font-Names="Tahoma" Font-Size="8pt" Font-Underline="False" Wrap="True"
                                HorizontalAlign="Center" />
                            <CellStyle Font-Names="Tahoma" Font-Size="8pt" HorizontalAlign="Right">
                            </CellStyle>
                       </dx:GridViewDataTextColumn>
                       <dx:GridViewDataTextColumn Caption="M3" FieldName="colM3" Name="colM3" 
                            VisibleIndex="23" Width="100px">
                            <HeaderStyle Font-Names="Tahoma" Font-Size="8pt" Font-Underline="False" Wrap="True"
                                HorizontalAlign="Center" />
                            <CellStyle Font-Names="Tahoma" Font-Size="8pt" HorizontalAlign="Right">
                            </CellStyle>
                        </dx:GridViewDataTextColumn>
                        <dx:GridViewDataTextColumn Caption="HEIGHT PALLET" FieldName="colHeightPallet" 
                            Name="colHeightPallet" VisibleIndex="24" Width="100px">
                            <HeaderStyle Font-Names="Tahoma" Font-Size="8pt" Font-Underline="False" Wrap="True"
                                HorizontalAlign="Center" />
                            <CellStyle Font-Names="Tahoma" Font-Size="8pt" HorizontalAlign="Right">
                            </CellStyle>
                        </dx:GridViewDataTextColumn>
                        <dx:GridViewDataTextColumn Caption="TOTAL BOX" FieldName="colTotalBox" 
                            Name="colTotalBox" VisibleIndex="25" Width="100px">
                            <HeaderStyle Font-Names="Tahoma" Font-Size="8pt" Font-Underline="False" Wrap="True"
                                HorizontalAlign="Center" />
                            <CellStyle Font-Names="Tahoma" Font-Size="8pt" HorizontalAlign="Right">
                            </CellStyle>
                        </dx:GridViewDataTextColumn>
                    <dx:GridViewDataTextColumn Caption="NoUrut" FieldName="NoUrut" 
               Name="NoUrut" Visible="False" VisibleIndex="26" Width="0px">
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
                    Font-Size="8pt" TabIndex="20">
                </dx1:ASPxButton>
                </td>
            <td align="right">
                &nbsp;</td>
            <td align="right">
                <dx1:ASPxButton ID="btndelete" runat="server" Text="DELETE" Width="90px"
                    Font-Names="Tahoma" Font-Size="8pt" AutoPostBack="False" TabIndex="21"
                    ClientInstanceName="btndelete">  
                    <ClientSideEvents Click="btndeleteClick"/>
                    <Paddings Padding="2px" />
                </dx1:ASPxButton>
               </td>
            <td width="90">
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

