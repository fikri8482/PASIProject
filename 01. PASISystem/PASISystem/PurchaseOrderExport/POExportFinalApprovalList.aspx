<%@ Page Title="" Language="vb" AutoEventWireup="false" MasterPageFile="~/SiteContent.Master" CodeBehind="POExportFinalApprovalList.aspx.vb" Inherits="PASISystem.POExportFinalApprovalList" %>
<%@ Register Assembly="DevExpress.Web.v14.1, Version=14.1.6.0, Culture=neutral, PublicKeyToken=b88d1754d700e49a"
    Namespace="DevExpress.Web.ASPxHiddenField" TagPrefix="dx" %>
<%@ Register Assembly="DevExpress.Web.v14.1, Version=14.1.6.0, Culture=neutral, PublicKeyToken=b88d1754d700e49a" namespace="DevExpress.Web.ASPxGridView.Export" TagPrefix="dx" %>
<%@ Register Assembly="DevExpress.Web.v14.1, Version=14.1.6.0, Culture=neutral, PublicKeyToken=b88d1754d700e49a" namespace="DevExpress.Web.ASPxGridView" TagPrefix="dx" %>
<%@ Register assembly="DevExpress.Web.v14.1, Version=14.1.6.0, Culture=neutral, PublicKeyToken=b88d1754d700e49a" namespace="DevExpress.Web.ASPxEditors" tagprefix="dx" %>
<%@ Register Assembly="DevExpress.Web.v14.1, Version=14.1.6.0, Culture=neutral, PublicKeyToken=b88d1754d700e49a" namespace="DevExpress.Web.ASPxGlobalEvents" TagPrefix="dx" %>
<%@ Register Assembly="DevExpress.Web.v14.1, Version=14.1.6.0, Culture=neutral, PublicKeyToken=b88d1754d700e49a" namespace="DevExpress.Web.ASPxPanel" TagPrefix="dx" %>
<%@ Register Assembly="DevExpress.Web.v14.1, Version=14.1.6.0, Culture=neutral, PublicKeyToken=b88d1754d700e49a" namespace="DevExpress.Web.ASPxRoundPanel" TagPrefix="dx" %>

<%@ Register assembly="DevExpress.Web.v14.1, Version=14.1.6.0, Culture=neutral, PublicKeyToken=b88d1754d700e49a" namespace="DevExpress.Web.ASPxGridView.Export" tagprefix="dx1" %>

<asp:Content ID="Content1" ContentPlaceHolderID="head" runat="server">
    <style type="text/css">
        .dxeHLC, .dxeHC, .dxeHFC
        {
            display: none;
        }
        .style2
        {
            width: 7px;
        }
        .style3
        {
            width: 180px;
        }
        .style4
        {
            height: 18px;
            width: 180px;
        }
        .style5
        {
            width: 90px;
        }
        .style6
        {
            height: 200px;
        }
        .style7
        {
            width: 50px;
        }
        </style>
<script type="text/javascript">
    function OnAllCheckedChanged(s, e) {
        if (s.GetValue() == -1) s.SetValue(1);
        for (var i = 0; i < grid.GetVisibleRowsOnPage(); i++) {
            grid.batchEditApi.SetCellValue(i, "Act", s.GetValue());
        }
    }

    function OnUpdateClick(s, e) {
        Grid.PerformCallback("Update");
    }

    function OnCancelClick(s, e) {
        Grid.PerformCallback("Cancel");
    }

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
        height = height - (height * 70 / 100)
        grid.SetHeight(height);
    }


    function OnBatchEditStartEditing(s, e) {
        currentColumnName = e.focusedColumn.fieldName;

        if (currentColumnName == "NoUrut" || currentColumnName == "coldetail" || currentColumnName == "Period" || currentColumnName == "OrderNo" || currentColumnName == "AffiliateID" || currentColumnName == "SupplierID" || currentColumnName == "PONo" || currentColumnName == "EmergencyCls" || currentColumnName == "CommercialCls"
            || currentColumnName == "ErrorStatus" || currentColumnName == "ShipCls" || currentColumnName == "POStatus1" || currentColumnName == "POStatus2"
            || currentColumnName == "POStatus3" || currentColumnName == "POStatus4" || currentColumnName == "POStatus5" || currentColumnName == "POStatus6") {
            e.cancel = true;
        }

        currentEditableVisibleIndex = e.visibleIndex;
    }
</script>
</asp:Content>

<asp:Content ID="Content2" ContentPlaceHolderID="Content" runat="server">
                <table style="border-width: 1pt thin thin thin; border-style: ridge; border-color:#9598A1; width:100%; height: 120px;">
                    <tr>
                        <td colspan="8" height="30">
                            <table id="Table1" style = "width:100%;">
                                <tr>
                                    <td align="left" valign="middle" class="style5">
                                        <dx:ASPxLabel ID="ASPxLabel20" runat="server" Text="AFFILIATE CODE/NAME"
                                            Font-Names="Tahoma" Font-Size="8pt" Width="150px">
                                        </dx:ASPxLabel>
                                    </td>
                                    <td align="left" valign="middle" height="18px" width="110px">
                                        <table style="width:100%;">
                                            <tr>
                                                <td>
                                        <dx:ASPxComboBox ID="cboAffiliate" width="110px" runat="server" Font-Size="8pt" 
                                                        Font-Names="Tahoma" TextFormatString="{0}" ClientInstanceName="cboAffiliate" 
                                                        TabIndex="3">
                                            <ClientSideEvents SelectedIndexChanged="function(s, e) {
	                                            txtAffiliate.SetText(cboAffiliate.GetSelectedItem().GetColumnText(1));                               
	                                            grid.PerformCallback('kosong');	
                                                lblInfo.SetText('');	
                                            }" />
                                            <LoadingPanelStyle ImageSpacing="5px">
                                            </LoadingPanelStyle>
                                        </dx:ASPxComboBox>
                                                </td>
                                                <td>
                                        <dx:ASPxTextBox ID="txtAffiliate" runat="server" Width="240px" 
                                            ClientInstanceName="txtAffiliate" Font-Names="Tahoma" Font-Size="8pt"
                                            ReadOnly="True" MaxLength="100" Height="18px">
                                            <readonlystyle backcolor="#CCCCCC">
                                            </readonlystyle>
                                        </dx:ASPxTextBox>
                                                </td>
                                            </tr>
                                        </table>
                                    </td>
                                    <td>
                                    </td>

                                    <td class="style7">
                                        &nbsp;</td>

                                </tr>

                                <tr>
                                    <td align="left" valign="middle" class="style5">
                                        <dx:ASPxLabel ID="ASPxLabel1" runat="server" Text="SUPPLIER CODE/NAME"
                                            Font-Names="Tahoma" Font-Size="8pt" Width="120px">
                                        </dx:ASPxLabel>
                                    </td>
                                    <td align="left" valign="middle" height="18px" width="110px">
                                        <table style="width:100%;">
                                            <tr>
                                                <td>
                                        <dx:ASPxComboBox ID="cboSupplierCode" width="110px" runat="server" Font-Size="8pt" 
                                                        Font-Names="Tahoma" TextFormatString="{0}" ClientInstanceName="cboSupplierCode" 
                                                        TabIndex="3">
                                            <ClientSideEvents SelectedIndexChanged="function(s, e) {
	                                            txtSupplierName.SetText(cboSupplierCode.GetSelectedItem().GetColumnText(1));                               
	                                            grid.PerformCallback('kosong');	
                                                lblInfo.SetText('');	
                                            }" />
                                            <LoadingPanelStyle ImageSpacing="5px">
                                            </LoadingPanelStyle>
                                        </dx:ASPxComboBox>
                                                </td>
                                                <td>
                                        <dx:ASPxTextBox ID="txtSupplierName" runat="server" Width="240px" 
                                            ClientInstanceName="txtSupplierName" Font-Names="Tahoma" Font-Size="8pt"
                                            ReadOnly="True" MaxLength="100" Height="18px">
                                            <readonlystyle backcolor="#CCCCCC">
                                            </readonlystyle>
                                        </dx:ASPxTextBox>
                                                </td>
                                            </tr>
                                        </table>
                                    </td>
                                    <td>
                                    </td>

                                    <td class="style7">
                                        &nbsp;</td>

                                </tr>
                                <tr>
                                    <td align="left" valign="middle" height="18px" class="style5">
                                        <dx:ASPxLabel ID="ASPxLabel3" runat="server" Text="PO MONTHLY / EMERGENCY"
                                            Font-Names="Tahoma" Font-Size="8pt" Width="150px">
                                        </dx:ASPxLabel>
                                    </td>
                                    <td align="left" valign="middle" height="18px" width="203px">
                                        <table>
                                            <tr>
                                                <td>

                                                    <dx:ASPxRadioButton ID="rdrEAll" ClientInstanceName="rdrEAll" runat="server" 
                                                        Text="ALL" GroupName="Emergency" Font-Names="Tahoma" Font-Size="8pt" 
                                                        TabIndex="7">
                                                        <ClientSideEvents LostFocus="function(s, e) {
                                                        grid.SetFocusedRowIndex(-1);
                                                        grid.PerformCallback('kosong');
                                                        lblInfo.SetText('');
                                                        }" />
                                                    </dx:ASPxRadioButton>
                                                </td>
                                                <td>
                                                    <dx:ASPxRadioButton ID="rdrEM" ClientInstanceName="rdrEM" 
                                                        runat="server" Text="MONTHLY" GroupName="Emergency" Font-Names="Tahoma" 
                                                        Font-Size="8pt" TabIndex="8">
                                                        <ClientSideEvents LostFocus="function(s, e) {
                                                        grid.SetFocusedRowIndex(-1);
                                                        grid.PerformCallback('kosong');
                                                        lblInfo.SetText('');
                                                        }" />
                                                    </dx:ASPxRadioButton>
                                                </td>
                                                <td>
                                                    <dx:ASPxRadioButton ID="rdrEE" ClientInstanceName="rdrEE" 
                                                        runat="server" Text="EMERGENCY" GroupName="Emergency" Font-Names="Tahoma" 
                                                        Font-Size="8pt" TabIndex="9">
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
                                    <td align="left" valign="middle" style="height:18px; width:180px;">&nbsp;</td>
                                    <td align="left" valign="middle" class="style7">&nbsp;</td>
                                </tr>
                                <tr>
                                    <td align="left" valign="middle" height="18px" class="style5">
                                        <dx:ASPxLabel ID="ASPxLabel11" runat="server" Text="PASI APPROVAL"
                                            Font-Names="Tahoma" Font-Size="8pt">
                                        </dx:ASPxLabel>
                                    </td>
                                    <td align="left" valign="middle" height="18px" width="203px">
                                        <table>
                                            <tr>
                                                <td>
                                                    <dx:ASPxRadioButton ID="rdAppALL" ClientInstanceName="rdAppALL" runat="server" 
                                                        Text="ALL" GroupName="Approval"  Font-Names="Tahoma" Font-Size="8pt" 
                                                        TabIndex="13">
                                                        <ClientSideEvents LostFocus="function(s, e) {
                                                        grid.SetFocusedRowIndex(-1);
                                                        grid.PerformCallback('kosong');
                                                        lblInfo.SetText('');
                                                        }" />
                                                    </dx:ASPxRadioButton>
                                                </td>
                                                <td>
                                                    <dx:ASPxRadioButton ID="rdAppYES" ClientInstanceName="rdAppYES" runat="server" 
                                                        Text="YES" GroupName="Approval"  Font-Names="Tahoma" Font-Size="8pt" 
                                                        TabIndex="14">
                                                        <ClientSideEvents LostFocus="function(s, e) {
                                                        grid.SetFocusedRowIndex(-1);
                                                        grid.PerformCallback('kosong');
                                                        lblInfo.SetText('');
                                                        }" />
                                                    </dx:ASPxRadioButton>
                                                </td>
                                                <td>
                                                    <dx:ASPxRadioButton ID="rdAppNO" ClientInstanceName="rdAppNO" runat="server" 
                                                        Text="NO" GroupName="Approval"  Font-Names="Tahoma" Font-Size="8pt" 
                                                        TabIndex="15">
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
                                    <td align="left" valign="middle" style="height:18px; width:180px;">
                                        <table style="width:100%;">
                                            <tr>
                                                <td align="right" valign="middle" style="height:18px; width:90px;">
                                                    &nbsp;</td>
                                            </tr>
                                        </table>
                                    </td>
                                    <td align="left" valign="middle" class="style7">
                                        &nbsp;</td>
                                </tr>
                                <tr>
                                    <td align="left" valign="middle" height="18px" class="style5">
                                        <dx:ASPxLabel ID="ASPxLabel4" runat="server" Text="COMMERCIAL"
                                            Font-Names="Tahoma" Font-Size="8pt">
                                        </dx:ASPxLabel>
                                    </td>
                                    <td align="left" valign="middle" height="18px" width="203px">
                                        <table>
                                            <tr>
                                                <td>
                                                    <dx:ASPxRadioButton ID="rdrCom1" ClientInstanceName="rdrCom1" runat="server" 
                                                        Text="ALL" GroupName="Commercial"  Font-Names="Tahoma" Font-Size="8pt" 
                                                        TabIndex="13">
                                                        <ClientSideEvents LostFocus="function(s, e) {
                                                        grid.SetFocusedRowIndex(-1);
                                                        grid.PerformCallback('kosong');
                                                        lblInfo.SetText('');
                                                        }" />
                                                    </dx:ASPxRadioButton>
                                                </td>
                                                <td>
                                                    <dx:ASPxRadioButton ID="rdrCom2" ClientInstanceName="rdrCom2" runat="server" 
                                                        Text="YES" GroupName="Commercial"  Font-Names="Tahoma" Font-Size="8pt" 
                                                        TabIndex="14">
                                                        <ClientSideEvents LostFocus="function(s, e) {
                                                        grid.SetFocusedRowIndex(-1);
                                                        grid.PerformCallback('kosong');
                                                        lblInfo.SetText('');
                                                        }" />
                                                    </dx:ASPxRadioButton>
                                                </td>
                                                <td>
                                                    <dx:ASPxRadioButton ID="rdrCom3" ClientInstanceName="rdrCom3" runat="server" 
                                                        Text="NO" GroupName="Commercial"  Font-Names="Tahoma" Font-Size="8pt" 
                                                        TabIndex="15">
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
                                    <td align="left" valign="middle" style="height:18px; width:180px;">
                                        <table style="width:100%;">
                                            <tr>
                                                <td align="right" valign="middle" style="height:18px; width:90px;">
                                                    &nbsp;</td>
                                            </tr>
                                        </table>
                                    </td>
                                    <td align="left" valign="middle" class="style7">
                                        &nbsp;</td>
                                </tr>
                                <tr>
                                    <td align="left" valign="middle" class="style5">
                                        <dx:ASPxLabel ID="ASPxLabel2" runat="server" Text="ORDER NO"
                                            Font-Names="Tahoma" Font-Size="8pt" Width="150px">
                                        </dx:ASPxLabel>
                                    </td>
                                    <td align="left" valign="middle" height="18px" width="110px">
                                        <table style="width:100%;">
                                            <tr>
                                                <td>
                                                    <dx:ASPxTextBox ID="txtOrderNo" runat="server" Width="110px" 
                                                        ClientInstanceName="txtOrderNo">
                                                    </dx:ASPxTextBox>
                                                </td>
                                                <td>
                                                    &nbsp;</td>
                                            </tr>
                                        </table>
                                    </td>
                                    <td align="right">
                                        <dx:ASPxLabel ID="ASPxLabel21" runat="server" BackColor="#99CCFF" Text=":" 
                                            Width="30px">
                                        </dx:ASPxLabel>
                                    </td>

                                    <td style="margin-left: 40px" align="center" class="style7">
                                        <dx:ASPxLabel ID="ASPxLabel22" runat="server" Text="By PASI" Width="50px">
                                        </dx:ASPxLabel>
                                    </td>

                                </tr>
                                <tr>
                                    <td align="left" valign="middle" colspan="4" width="100%">
                                        <dx:ASPxRoundPanel ID="ASPxRoundPanel2" runat="server" 
            HeaderText="PO STATUS" ShowCollapseButton="true" 
            View="GroupBox" Width="100%" Height="40px">
            <ContentPaddings PaddingLeft="5px" PaddingRight="5px" />
<ContentPaddings PaddingLeft="5px" PaddingRight="5px"></ContentPaddings>
            <PanelCollection>
                <dx:PanelContent ID="PanelContent1" runat="server">
               <table style="width:100%;">
                    <tr>
                        <td class="style2">
                            <table style="width:100%;">
                                <tr>
                                    <td>
                                        <dx:ASPxButton ID="btn1" runat="server" AutoPostBack="False" 
                                            ClientInstanceName="btn1" EncodeHtml="False" Height="65px" 
                                            Text="(1) &lt;/br&gt; UPLOADED" Width="110px" EnableDefaultAppearance="False"
                                            EnableTheming="True" BackColor="#99CCFF" Theme="MetropolisBlue">
                                            <ClientSideEvents Click="function(s, e) {
grid.PerformCallback('POStatus1');
}" />
                                        </dx:ASPxButton>
                                    </td>
                                    <td align="center">
                                        <dx:ASPxLabel ID="ASPxLabel10" runat="server" EncodeHtml="False" 
                                            Text="&amp;rarr;" Width="20px">
                                        </dx:ASPxLabel>
                                    </td>
                                    <td>
                            
                                        <dx:ASPxButton ID="btn2" runat="server" AutoPostBack="False" 
                                            ClientInstanceName="btn2" Text="(2) </br> CHECK & SEND </br> TO SUPP"
                                            Width="110px" Height="65px" Wrap="True" EncodeHtml="False">
                                            <ClientSideEvents Click="function(s, e) {
grid.PerformCallback('POStatus2');
}" />
                                        </dx:ASPxButton>
                            
                                    </td>
                                    <td align="center">
                                        <dx:ASPxLabel ID="ASPxLabel5" runat="server" EncodeHtml="False" 
                                            Text="&amp;rarr;" Width="20px">
                                        </dx:ASPxLabel>
                                    </td>
                                    <td>
                            
                                        <dx:ASPxButton ID="btn3" runat="server" AutoPostBack="False" 
                                            ClientInstanceName="btn3" 
                                            
                                            Text="(3) (4) (5) &lt;/br&gt; SUPP APPROVE &lt;/br&gt; (FULL/PARTIAL/ &lt;/br&gt; UNAPPROVE)" 
                                            Width="110px" EncodeHtml="False" Height="65px" BackColor="#99CCFF" 
                                            Theme="MetropolisBlue">
                                            <ClientSideEvents Click="function(s, e) {
grid.PerformCallback('POStatus3');
}" />
                                        </dx:ASPxButton>
                            
                                    </td>
                                    <td align="center">
                                        <dx:ASPxLabel ID="ASPxLabel6" runat="server" EncodeHtml="False" 
                                            Text="&amp;rarr;" Width="20px">
                                        </dx:ASPxLabel>
                                    </td>
                                    <td>
                                        <dx:ASPxButton ID="btn4" runat="server" AutoPostBack="False" 
                                            ClientInstanceName="btn4" Text="(6) </br> PASI FINAL </br> APPROVE" Width="110px" 
                                            EncodeHtml="False" Height="65px">
                                            <ClientSideEvents Click="function(s, e) {
grid.PerformCallback('POStatus4');
}" />
                                        </dx:ASPxButton>
                                    </td>
                                    <td align="center">
                                        <dx:ASPxLabel ID="ASPxLabel7" runat="server" EncodeHtml="False" 
                                            Text="&amp;rarr;" Width="20px">
                                        </dx:ASPxLabel>  
                                    </td>
                                    <td>
                                        <dx:ASPxButton ID="btn5" runat="server" AutoPostBack="False" 
                                            ClientInstanceName="btn5" 
                                            Text="(7) &lt;/br&gt; DELIVERY BY &lt;/br&gt; SUPP" Width="110px" Height="65px" 
                                            EncodeHtml="False" BackColor="#99CCFF" Theme="MetropolisBlue" 
                                            >
                                            <ClientSideEvents Click="function(s, e) {
grid.PerformCallback('POStatus5');
}" />
                                        </dx:ASPxButton>
                                    </td>
                                    <td align="center">
                                        <dx:ASPxLabel ID="ASPxLabel8" runat="server" EncodeHtml="False" 
                                            Text="&amp;rarr;" Width="20px">
                                        </dx:ASPxLabel>
                                    </td>
                                    <td>
                                        <dx:ASPxButton ID="btn6" runat="server" AutoPostBack="False" 
                                            ClientInstanceName="btn6" Text="(8) &lt;/br&gt; RECEIVE BY FWD" 
                                            Width="110px" Height="65px" EncodeHtml="False" BackColor="#99CCFF" 
                                            Theme="MetropolisBlue">
                                            <ClientSideEvents Click="function(s, e) {
grid.PerformCallback('POStatus6');
}" />
                                        </dx:ASPxButton>
                                    </td>
                                    <td align="center">
                                        <dx:ASPxLabel ID="ASPxLabel9" runat="server" EncodeHtml="False" 
                                            Text="&amp;rarr;" Width="20px">
                                        </dx:ASPxLabel>
                                    </td>
                                    <td>
                                        <dx:ASPxButton ID="btn7" runat="server" AutoPostBack="False" 
                                            ClientInstanceName="btn7" 
                                            Text="(9) &lt;/br&gt; SHIPPING &lt;/br&gt; INSTRUCTION &lt;/br&gt; BY PASI" 
                                            Width="110px" Height="65px" EncodeHtml="False">
                                            <ClientSideEvents Click="function(s, e) {
grid.PerformCallback('POStatus7');
}" />
                                        </dx:ASPxButton>
                                    </td>
                                </tr>
                            </table>
                        </td>
                    </tr>
                </table>
                </dx:PanelContent>
            </PanelCollection>
        </dx:ASPxRoundPanel>
                                    </td>

                                </tr>
                                </table>
                        </td>
                    </tr>
                </table>

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

    <table style="width:100%;">
        <tr>
            <td align="left" class="style6" colspan="5">
                <%--Column : Grid--%>
                <dx:ASPxGridView ID="grid" runat="server" Width="100%"
                    Font-Names="Tahoma" KeyFieldName="NoUrut;Period;OrderNomor;AffiliateID;SupplierID;ForwarderID"
                    AutoGenerateColumns="False"
                    ClientInstanceName="grid" Font-Size="8pt">
                    <ClientSideEvents Init="OnInit" FocusedRowChanged="function(s, e) {
	                    OnGridFocusedRowChanged();
                    }" EndCallback="function(s, e) {
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
                        <dx:GridViewDataTextColumn Caption="NO." VisibleIndex="0" Width="30px" 
                            FieldName="NoUrut">
                        </dx:GridViewDataTextColumn>
                        <dx:GridViewDataCheckColumn FieldName="cols" Name="cols" VisibleIndex="1" Width="40px"
                            Caption=" ">
                            <HeaderStyle HorizontalAlign="Center" VerticalAlign="Middle" />
                            <HeaderCaptionTemplate>
                                <dx:ASPxCheckBox ID="chkAll" runat="server" ClientInstanceName="chkAll" ClientSideEvents-CheckedChanged="OnAllCheckedChanged"
                               ValueType="System.String" ValueChecked="1" ValueUnchecked="0">
                                </dx:ASPxCheckBox>
                            </HeaderCaptionTemplate>
                            <PropertiesCheckEdit ValueChecked="1" ValueType="System.Int32" 
                                ValueUnchecked="0">
                            </PropertiesCheckEdit>
                        </dx:GridViewDataCheckColumn>
                        <dx:GridViewDataHyperLinkColumn Caption=" " FieldName="coldetail" Name="coldetail"
                            VisibleIndex="2" Width="65px">
                            <PropertiesHyperLinkEdit TextField="DetailPage">
                            </PropertiesHyperLinkEdit>
                            <HeaderStyle Font-Names="Tahoma" Font-Size="8pt" Font-Underline="False" Wrap="True" />
                            <CellStyle Font-Names="Tahoma" Font-Size="8pt" VerticalAlign="Middle">
                            </CellStyle>
                        </dx:GridViewDataHyperLinkColumn>
                        <dx:GridViewDataTextColumn Caption="AFFILIATE" FieldName="AffiliateID" 
                            VisibleIndex="3">
                        </dx:GridViewDataTextColumn>
                         <dx:GridViewDataTextColumn VisibleIndex="4" Caption="SUPPLIER" 
                            FieldName="SupplierID" Width="120px" HeaderStyle-HorizontalAlign="Center">
                            <HeaderStyle HorizontalAlign="Center"></HeaderStyle>
                        </dx:GridViewDataTextColumn>
                        <dx:GridViewDataTextColumn VisibleIndex="5" Caption="ORDER NO." 
                            FieldName="PONo" Width="150px"
                            HeaderStyle-HorizontalAlign="Center">
                            <HeaderStyle HorizontalAlign="Center"></HeaderStyle>
                            <CellStyle Font-Names="Tahoma" Font-Size="8pt">
                            </CellStyle>
                        </dx:GridViewDataTextColumn>
                        <dx:GridViewDataTextColumn VisibleIndex="6" Caption="PERIOD" FieldName="Period" Width="70px"
                            HeaderStyle-HorizontalAlign="Center">
                            <PropertiesTextEdit DisplayFormatString="yyyy-MM">
                            </PropertiesTextEdit>
                            <HeaderStyle HorizontalAlign="Center"></HeaderStyle>
                            <CellStyle Font-Names="Tahoma" Font-Size="8pt">
                            </CellStyle>
                        </dx:GridViewDataTextColumn>
                        <dx:GridViewDataTextColumn VisibleIndex="7" Caption="PO MONTHLY / EMERGENCY" FieldName="EmergencyCls"
                            Width="93px" HeaderStyle-HorizontalAlign="Center" 
                            CellStyle-HorizontalAlign="Center">
                            <HeaderStyle HorizontalAlign="Center" Wrap="True"></HeaderStyle>
                            <CellStyle Font-Names="Tahoma" Font-Size="8pt">
                            </CellStyle>
                        </dx:GridViewDataTextColumn>
                        <dx:GridViewDataTextColumn VisibleIndex="8" Caption="COMMERCIAL" FieldName="CommercialCls"
                            Width="95px" HeaderStyle-HorizontalAlign="Center" 
                            CellStyle-HorizontalAlign="Center">
                            <HeaderStyle HorizontalAlign="Center"></HeaderStyle>
                            <CellStyle Font-Names="Tahoma" Font-Size="8pt">
                            </CellStyle>
                        </dx:GridViewDataTextColumn>
                        <dx:GridViewDataTextColumn VisibleIndex="9" Caption="SHIP BY (BOAT / AIR)" FieldName="ShipCls"
                            Width="90px" HeaderStyle-HorizontalAlign="Center">
                            <HeaderStyle HorizontalAlign="Center" Wrap="True"></HeaderStyle>
                            <CellStyle Font-Names="Tahoma" Font-Size="8pt">
                            </CellStyle>
                        </dx:GridViewDataTextColumn>
                        <dx:GridViewDataTextColumn VisibleIndex="10" Caption="ERROR STATUS" FieldName="ErrorStatus"
                            HeaderStyle-HorizontalAlign="Center" Width="60px">
                            <HeaderStyle HorizontalAlign="Center" Wrap="True"></HeaderStyle>
                            <CellStyle Font-Names="Tahoma" Font-Size="8pt">
                            </CellStyle>
                            <HeaderStyle HorizontalAlign="Center"></HeaderStyle>
                        </dx:GridViewDataTextColumn>
                        <dx:GridViewBandColumn Caption="PO STATUS" VisibleIndex="11">
                            <Columns>
                        <dx:GridViewDataTextColumn Caption="Part NO" FieldName="PartNo" VisibleIndex="29" 
                            Width="0px" CellStyle-HorizontalAlign="Center" Visible="False">
                            <HeaderStyle HorizontalAlign="Center"></HeaderStyle>
                            <CellStyle Font-Names="Tahoma" Font-Size="8pt">
                            </CellStyle>
                        </dx:GridViewDataTextColumn>
                        <dx:GridViewDataTextColumn VisibleIndex="23" Caption="PASISendToSupplierCls" 
                            FieldName="PASISendToSupplierCls" Width="0px" HeaderStyle-HorizontalAlign="Center">
                            <HeaderStyle HorizontalAlign="Center"></HeaderStyle>
                        </dx:GridViewDataTextColumn>
                        <dx:GridViewDataTextColumn VisibleIndex="24" Caption="SupplierApprovalCls" 
                            FieldName="SupplierApprovalCls" Width="0px" HeaderStyle-HorizontalAlign="Center">
                            <HeaderStyle HorizontalAlign="Center"></HeaderStyle>
                        </dx:GridViewDataTextColumn>
                        <dx:GridViewDataTextColumn Caption="1" FieldName="POStatus1" ReadOnly="True" 
                            VisibleIndex="1" Width="60px">
                            <cellstyle wrap="True">
                            </cellstyle>
                            <HeaderStyle HorizontalAlign="Center" />
                        </dx:GridViewDataTextColumn>
                        <dx:GridViewDataHyperLinkColumn Caption="2" FieldName="GOTOPOStatus2" Name="GOTOPOStatus2"
                            VisibleIndex="2" Width="0px">
                            <PropertiesHyperLinkEdit TextField="detailGOTO">
                            </PropertiesHyperLinkEdit>
                            <HeaderStyle Font-Names="Tahoma" HorizontalAlign="Center" Font-Size="8pt" Font-Underline="False" Wrap="True" />
                            <CellStyle Font-Names="Tahoma" Font-Size="8pt" VerticalAlign="Middle">
                            </CellStyle>
                        </dx:GridViewDataHyperLinkColumn>
                        <dx:GridViewDataTextColumn Caption="2" FieldName="POStatus2" ReadOnly="True" 
                            VisibleIndex="3" Width="60px">
                            <cellstyle wrap="True">
                            </cellstyle>
                            <HeaderStyle HorizontalAlign="Center" />
                        </dx:GridViewDataTextColumn>
                        <dx:GridViewDataTextColumn Caption="3" FieldName="POStatus3" ReadOnly="True" 
                            VisibleIndex="4" Width="60px">
                            <cellstyle wrap="True">
                            </cellstyle>
                            <HeaderStyle HorizontalAlign="Center" />
                        </dx:GridViewDataTextColumn>
                        <dx:GridViewDataTextColumn Caption="4" FieldName="POStatus4" ReadOnly="True" 
                            VisibleIndex="5" Width="60px">
                            <cellstyle wrap="True">
                            </cellstyle>
                            <HeaderStyle HorizontalAlign="Center" />
                        </dx:GridViewDataTextColumn>
                        <dx:GridViewDataTextColumn Caption="5" FieldName="POStatus5" ReadOnly="True" 
                            VisibleIndex="6" Width="60px">
                            <cellstyle wrap="True">
                            </cellstyle>
                            <HeaderStyle HorizontalAlign="Center" />
                        </dx:GridViewDataTextColumn>
                        <dx:GridViewDataHyperLinkColumn Caption="6" FieldName="GOTOPOStatus6" Name="GOTOPOStatus6"
                            VisibleIndex="7" Width="0px">
                            <PropertiesHyperLinkEdit TextField="detailGOTO">
                            </PropertiesHyperLinkEdit>
                            <HeaderStyle Font-Names="Tahoma" HorizontalAlign="Center" Font-Size="8pt" Font-Underline="False" Wrap="True" />
                            <CellStyle Font-Names="Tahoma" Font-Size="8pt" VerticalAlign="Middle">
                            </CellStyle>
                        </dx:GridViewDataHyperLinkColumn>
                        <dx:GridViewDataTextColumn Caption="6" FieldName="POStatus6" ReadOnly="True" 
                            VisibleIndex="8" Width="60px">
                            <cellstyle wrap="True">
                            </cellstyle>
                            <HeaderStyle HorizontalAlign="Center" />
                        </dx:GridViewDataTextColumn>
                        <dx:GridViewDataHyperLinkColumn Caption="7" FieldName="GOTOPOStatus7" Name="GOTOPOStatus7"
                            VisibleIndex="9" Width="0px">
                            <PropertiesHyperLinkEdit TextField="detailGOTO">
                            </PropertiesHyperLinkEdit>
                            <HeaderStyle Font-Names="Tahoma" HorizontalAlign="Center" Font-Size="8pt" Font-Underline="False" Wrap="True" />
                            <CellStyle Font-Names="Tahoma" Font-Size="8pt" VerticalAlign="Middle">
                            </CellStyle>
                        </dx:GridViewDataHyperLinkColumn>
                        <dx:GridViewDataTextColumn Caption="7" FieldName="POStatus7" ReadOnly="True" 
                            VisibleIndex="10" Width="60px">
                            <cellstyle wrap="True">
                            </cellstyle>
                            <HeaderStyle HorizontalAlign="Center" />
                        </dx:GridViewDataTextColumn>
                       <dx:GridViewDataHyperLinkColumn Caption="8" FieldName="GOTOPOStatus8" Name="GOTOPOStatus8"
                            VisibleIndex="11" Width="0px">
                            <PropertiesHyperLinkEdit TextField="detailGOTO">
                            </PropertiesHyperLinkEdit>
                            <HeaderStyle Font-Names="Tahoma" HorizontalAlign="Center" Font-Size="8pt" Font-Underline="False" Wrap="True" />
                            <CellStyle Font-Names="Tahoma" Font-Size="8pt" VerticalAlign="Middle">
                            </CellStyle>
                        </dx:GridViewDataHyperLinkColumn>
                        <dx:GridViewDataTextColumn Caption="8" FieldName="POStatus8" ReadOnly="True" 
                            VisibleIndex="12" Width="60px">
                            <cellstyle wrap="True">
                            </cellstyle>
                            <HeaderStyle HorizontalAlign="Center" />
                        </dx:GridViewDataTextColumn>
                        <dx:GridViewDataHyperLinkColumn Caption="9" FieldName="GOTOPOStatus9" Name="GOTOPOStatus9"
                            VisibleIndex="13" Width="0px">
                            <PropertiesHyperLinkEdit TextField="detailGOTO">
                            </PropertiesHyperLinkEdit>
                            <HeaderStyle Font-Names="Tahoma" HorizontalAlign="Center" Font-Size="8pt" Font-Underline="False" Wrap="True" />
                            <CellStyle Font-Names="Tahoma" Font-Size="8pt" VerticalAlign="Middle">
                            </CellStyle>
                        </dx:GridViewDataHyperLinkColumn>
                        <dx:GridViewDataTextColumn Caption="9" FieldName="POStatus9" ReadOnly="True" 
                            VisibleIndex="14" Width="60px">
                            <cellstyle wrap="True">
                            </cellstyle>
                            <HeaderStyle HorizontalAlign="Center" />
                        </dx:GridViewDataTextColumn>
                            </Columns>
                            <HeaderStyle HorizontalAlign="Center"></HeaderStyle>
                        </dx:GridViewBandColumn>
                        <dx:GridViewDataTextColumn Caption="Forwarder" FieldName="ForwarderID" 
                            VisibleIndex="25" Width="0px">
                        </dx:GridViewDataTextColumn>
                         <dx:GridViewDataTextColumn Caption="OrderNomor" FieldName="OrderNomor" 
                            VisibleIndex="26" Width="0px">
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
                        ShowStatusBar="Hidden" VerticalScrollableHeight="210" />

                    <Settings ShowGroupButtons="False" ShowVerticalScrollBar="True" ShowHorizontalScrollBar="True" 
                        VerticalScrollableHeight="210" ShowStatusBar="Hidden"></Settings>
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

    <table style="width: 100%;">
        <tr>
            <td align="left">
                <dx:ASPxButton ID="btnSubMenu" runat="server" Text="SUB MENU"
                    Font-Names="Tahoma" Width="85px" Font-Size="8pt" HorizontalAlign="Center" 
                    TabIndex="21">
                </dx:ASPxButton>
            </td>
            <td>
                &nbsp;</td>
            <td>
                &nbsp;</td>
            <td>
                &nbsp;
            </td>
            <td>
                &nbsp;</td>
            <td>
                &nbsp;</td>
            <td>
                &nbsp;
                </td>
            <td align="right">
                <dx:ASPxButton ID="btnApprove" 
                runat="server" Text="FINAL APPROVAL"
                    Font-Names="Tahoma"
                    Width="80px" AutoPostBack="False" Font-Size="8pt" 
                    ClientInstanceName="btnApprove" HorizontalAlign="Right" TabIndex="20">
                 <ClientSideEvents Click="function(s, e) {   
                    grid.UpdateEdit();
                    grid.PerformCallback('load');
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
    <dx:ASPxHiddenField ID="HF" runat="server" ClientInstanceName="HF">
    </dx:ASPxHiddenField>
    </asp:Content>



