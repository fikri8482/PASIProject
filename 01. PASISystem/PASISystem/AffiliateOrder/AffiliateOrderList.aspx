<%@ Page Title="" Language="vb" AutoEventWireup="false" MasterPageFile="~/SiteContent.Master"
    CodeBehind="AffiliateOrderList.aspx.vb" Inherits="PASISystem.AffiliateOrderList" %>

<%@ Register Assembly="DevExpress.Web.v14.1, Version=14.1.6.0, Culture=neutral, PublicKeyToken=b88d1754d700e49a"
    Namespace="DevExpress.Web.ASPxGridView.Export" TagPrefix="dx" %>
<%@ Register Assembly="DevExpress.Web.v14.1, Version=14.1.6.0, Culture=neutral, PublicKeyToken=b88d1754d700e49a"
    Namespace="DevExpress.Web.ASPxGridView" TagPrefix="dx" %>
<%@ Register Assembly="DevExpress.Web.v14.1, Version=14.1.6.0, Culture=neutral, PublicKeyToken=b88d1754d700e49a"
    Namespace="DevExpress.Web.ASPxEditors" TagPrefix="dx" %>
<%@ Register Assembly="DevExpress.Web.v14.1, Version=14.1.6.0, Culture=neutral, PublicKeyToken=b88d1754d700e49a"
    Namespace="DevExpress.Web.ASPxGlobalEvents" TagPrefix="dx" %>
<%@ Register Assembly="DevExpress.Web.v14.1, Version=14.1.6.0, Culture=neutral, PublicKeyToken=b88d1754d700e49a"
    Namespace="DevExpress.Web.ASPxRoundPanel" TagPrefix="dx" %>
<%@ Register Assembly="DevExpress.Web.v14.1, Version=14.1.6.0, Culture=neutral, PublicKeyToken=b88d1754d700e49a"
    Namespace="DevExpress.Web.ASPxPanel" TagPrefix="dx" %>
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
            height = height - (height * 60 / 100)
            grid.SetHeight(height);
        }

        function OnBatchEditStartEditing(s, e) {
            currentColumnName = e.focusedColumn.fieldName;

            if (currentColumnName == "NoUrut" || currentColumnName == "Period" || currentColumnName == "AffiliateID" || currentColumnName == "AffiliateName"
            || currentColumnName == "PONo" || currentColumnName == "POMarking" || currentColumnName == "CommercialCls" || currentColumnName == "SupplierID" || currentColumnName == "SupplierName"
            || currentColumnName == "ShipCls" || currentColumnName == "EntryDate" || currentColumnName == "EntryUser" || currentColumnName == "POStatus1"
            || currentColumnName == "POStatus2" || currentColumnName == "POStatus3" || currentColumnName == "POStatus4" || currentColumnName == "POStatus5"
            || currentColumnName == "POStatus6" || currentColumnName == "POStatus7" || currentColumnName == "POStatus8") {
                e.cancel = true;
            }

            currentEditableVisibleIndex = e.visibleIndex;
        }
        function SelectedIndexChangedAff() {
            txtAffiliateName.SetText(cboAffiliateCode.GetSelectedItem().GetColumnText(1));
            lblInfo.SetText('');
        }
        function SelectedIndexChangedSupp() {
            txtSupplierName.SetText(cboSupplierCode.GetSelectedItem().GetColumnText(1));
            lblInfo.SetText('');
        }
    </script>
</asp:Content>
<asp:Content ID="Content2" ContentPlaceHolderID="Content" runat="server">
    <table style="width: 100%;">
        <tr>
            <td valign="top" width="60%">
                <table style="border-width: 1pt thin thin thin; border-style: ridge; border-color: #9598A1;
                    width: 100%;">
                    <tr>
                        <td height="30">
                            <table id="Table1">
                                <tr>
                                    <td style="width: 5px;">
                                        &nbsp;
                                    </td>
                                    <td align="left" valign="middle" width="120px">
                                        <dx:aspxlabel id="ASPxLabel3" runat="server" text="PERIOD" font-names="Tahoma" font-size="8pt"
                                            width="100%">
                                        </dx:aspxlabel>
                                    </td>
                                    <td align="left" valign="middle" width="150px">
                                        <dx:aspxtimeedit id="dtPeriodFrom" runat="server" clientinstancename="dtPeriodFrom"
                                            displayformatstring="MMM yyyy" editformat="Custom" editformatstring="MMM yyyy"
                                            width="150px" font-names="Tahoma" font-size="8pt">
                                        </dx:aspxtimeedit>
                                    </td>
                                    <td align="left" valign="middle" width="10px">
                                        <dx:aspxlabel id="ASPxLabel4" runat="server" text="~" font-names="Tahoma" font-size="8pt"
                                            width="10px">
                                        </dx:aspxlabel>
                                    </td>
                                    <td align="left" valign="middle" width="150px">
                                        <dx:aspxtimeedit id="dtPeriodTo" runat="server" clientinstancename="dtPeriodTo" displayformatstring="MMM yyyy"
                                            editformat="Custom" editformatstring="MMM yyyy" width="150px" font-names="Tahoma"
                                            font-size="8pt">
                                        </dx:aspxtimeedit>
                                    </td>
                                    <td align="left" valign="middle" width="50px">
                                        <dx:aspxlabel id="ASPxLabel5" runat="server" text="PO NO." font-names="Tahoma" font-size="8pt"
                                            width="50px">
                                        </dx:aspxlabel>
                                    </td>
                                    <td align="left" valign="middle" width="165px">
                                        <dx:aspxtextbox id="txtPONo" runat="server" width="150px" height="20px" clientinstancename="txtPONo"
                                            font-names="Tahoma" font-size="8pt" maxlength="10" tabindex="3">
                                            <ClientSideEvents TextChanged="function(s, e) {
                                                grid.SetFocusedRowIndex(-1);                                                
	                                            lblErrMsg.SetText('');
                                            }" />
                                        </dx:aspxtextbox>
                                    </td>
                                </tr>
                                <tr>
                                    <td style="width: 5px;">
                                    </td>
                                    <td align="left" valign="middle" width="120px">
                                        <dx:aspxlabel id="ASPxLabel1" runat="server" text="AFFILIATE CODE/NAME" font-names="Tahoma"
                                            font-size="8pt" width="150px">
                                        </dx:aspxlabel>
                                    </td>
                                    <td align="left" valign="middle" width="150px">
                                        <dx:aspxcombobox id="cboAffiliateCode" runat="server" textformatstring="{0}" dropdownstyle="DropDown"
                                            width="150px" clientinstancename="cboAffiliateCode" incrementalfilteringmode="StartsWith"
                                            font-names="Tahoma" font-size="8pt" maxlength="10" tabindex="4">
                                            <ClientSideEvents SelectedIndexChanged="function(s, e) {SelectedIndexChangedAff();}"
                                                LostFocus="function(s, e) { lblInfo.SetText(&quot;&quot;); }" />
                                        </dx:aspxcombobox>
                                    </td>
                                    <td align="left" valign="middle" colspan="4" width="360px">
                                        <dx:aspxtextbox id="txtAffiliateName" runat="server" clientinstancename="txtAffiliateName"
                                            width="370px" height="20px" readonly="True" tabindex="5" font-names="Tahoma"
                                            font-size="8pt">
                                            <ReadOnlyStyle BackColor="#CCCCCC">
                                            </ReadOnlyStyle>
                                        </dx:aspxtextbox>
                                    </td>
                                </tr>
                                <tr>
                                    <td style="width: 5px;">
                                    </td>
                                    <td align="left" valign="middle" width="80px">
                                        <dx:aspxlabel id="ASPxLabel2" runat="server" text="SUPPLIER CODE/NAME" font-names="Tahoma"
                                            font-size="8pt" width="150px">
                                        </dx:aspxlabel>
                                    </td>
                                    <td align="left" valign="middle" width="150px">
                                        <dx:aspxcombobox id="cboSupplierCode" runat="server" textformatstring="{0}" dropdownstyle="DropDown"
                                            height="20px" width="150px" clientinstancename="cboSupplierCode" incrementalfilteringmode="StartsWith"
                                            font-names="Tahoma" font-size="8pt" maxlength="10" tabindex="6">
                                            <ClientSideEvents SelectedIndexChanged="function(s, e) {SelectedIndexChangedSupp();}"
                                                LostFocus="function(s, e) { lblInfo.SetText(&quot;&quot;); }" />
                                        </dx:aspxcombobox>
                                    </td>
                                    <td align="left" valign="middle" width="360px" colspan="4">
                                        <dx:aspxtextbox id="txtSupplierName" runat="server" clientinstancename="txtSupplierName"
                                            width="370px" height="20px" readonly="True" tabindex="7" font-names="Tahoma"
                                            font-size="8pt">
                                            <ReadOnlyStyle BackColor="#CCCCCC">
                                            </ReadOnlyStyle>
                                        </dx:aspxtextbox>
                                    </td>
                                </tr>
                                <tr>
                                    <td style="width: 5px;">
                                        &nbsp;
                                    </td>
                                    <td align="left" valign="middle" width="80px">
                                        <dx:aspxlabel id="ASPxLabel6" runat="server" text="SEND TO SUPPLIER" font-names="Tahoma"
                                            font-size="8pt" width="150px">
                                        </dx:aspxlabel>
                                    </td>
                                    <td align="left" valign="middle" width="150px">
                                        <dx:aspxradiobuttonlist id="rblSendTo" runat="server" repeatdirection="Horizontal"
                                            width="150px" clientinstancename="rblSendTo" selectedindex="0" tabindex="8" font-names="Tahoma"
                                            font-size="8pt">
                                            <RadioButtonStyle HorizontalAlign="Left">
                                            </RadioButtonStyle>
                                            <Items>
                                                <dx:ListEditItem Text="ALL" Value="2" Selected="True" />
                                                <dx:ListEditItem Text="YES" Value="1" />
                                                <dx:ListEditItem Text="NO" Value="0" />
                                            </Items>
                                            <Border BorderStyle="None"></Border>
                                        </dx:aspxradiobuttonlist>
                                    </td>
                                    <td align="left" valign="middle" width="80px" colspan="4">
                                    </td>
                                </tr>
                                <tr>
                                    <td style="width: 5px;">
                                        &nbsp;
                                    </td>
                                    <td align="left" valign="middle" width="80px">
                                        <dx:aspxlabel id="ASPxLabel7" runat="server" text="COMMERCIAL" font-names="Tahoma"
                                            font-size="8pt" width="100%">
                                        </dx:aspxlabel>
                                    </td>
                                    <td align="left" valign="middle" width="150px">
                                        <dx:aspxradiobuttonlist id="rblCommercial" runat="server" repeatdirection="Horizontal"
                                            width="150px" clientinstancename="rblCommercial" selectedindex="0" tabindex="9"
                                            font-names="Tahoma" font-size="8pt">
                                            <RadioButtonStyle HorizontalAlign="Left">
                                            </RadioButtonStyle>
                                            <Items>
                                                <dx:ListEditItem Text="ALL" Value="2" Selected="True" />
                                                <dx:ListEditItem Text="YES" Value="1" />
                                                <dx:ListEditItem Text="NO" Value="0" />
                                            </Items>
                                            <Border BorderStyle="None"></Border>
                                        </dx:aspxradiobuttonlist>
                                    </td>
                                    <td align="left" valign="middle" width="80px" colspan="2">
                                        &nbsp;
                                    </td>
                                    <td align="left" valign="middle" width="120px" colspan="2">
                                        <table style="width: 100%;">
                                            <tr>
                                                <td align="right" valign="middle">
                                                    <dx:aspxbutton id="btnRefresh" runat="server" text="SEARCH" font-names="Tahoma" width="85px"
                                                        autopostback="False" font-size="8pt" clientinstancename="btnRefresh" tabindex="10">
                                                        <ClientSideEvents Click="function(s, e) {
                                                            var pDateFrom = dtPeriodFrom.GetText();
                                                            var pDateTo = dtPeriodTo.GetText();
                                                            var pPONo = txtPONo.GetText();
                                                            var pAffCode= cboAffiliateCode.GetText();
                                                            var pSuppCode = cboSupplierCode.GetText();
                                                            var pSendTo = rblSendTo.GetValue();
                                                            var pComm = rblCommercial.GetValue();
                                                            
                                                            grid.PerformCallback('load' + '|' + pDateFrom + '|' + pDateTo+ '|' + pPONo+ '|' + pAffCode + '|' + pSuppCode + '|' + pSendTo + '|' + pComm);
                                                            lblInfo.SetText('');
                                                        }" />
                                                    </dx:aspxbutton>
                                                </td>
                                                <td align="right" valign="middle">
                                                    <dx:aspxbutton id="btnClear" runat="server" text="CLEAR" font-names="Tahoma" width="85px"
                                                        autopostback="False" font-size="8pt" clientinstancename="btnClear" tabindex="11">
                                                        <ClientSideEvents Click="function(s, e) {
                                                            txtPONo.SetText('');
                                                            lblInfo.SetText('');
                                                            grid.SetFocusedRowIndex(-1);                                                            
                                                        }" />
                                                    </dx:aspxbutton>
                                                </td>
                                            </tr>
                                        </table>
                                    </td>
                                </tr>
                            </table>
                        </td>
                    </tr>
                </table>
            </td>
            <td valign="top" width="40%" align="left">
                <dx:aspxroundpanel id="ASPxRoundPanel1" runat="server" headertext="PO STATUS" showcollapsebutton="true"
                    view="GroupBox" width="100%" font-size="8pt" font-names="Tahoma">
                    <PanelCollection>
                        <dx:PanelContent ID="PanelContent1" runat="server">
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
                </dx:aspxroundpanel>
            </td>
        </tr>
    </table>
    <div style="height: 1px;">
    </div>
    <table style="width: 100%; height: 15px;">
        <tr>
            <td colspan="8" height="15">
                <%--error message--%>
                <table id="tblMsg" style="border-width: thin; border-style: inset hidden ridge hidden;
                    border-color: #9598A1; width: 100%; height: 15px;">
                    <tr>
                        <td align="left" valign="middle" height="15px">
                            <dx:aspxlabel id="lblInfo" runat="server" text="[lblinfo]" font-names="Tahoma" clientinstancename="lblInfo"
                                font-bold="True" font-italic="True" font-size="8pt">
                            </dx:aspxlabel>
                        </td>
                    </tr>
                </table>
            </td>
        </tr>
    </table>
    <div style="height: 1px;">
    </div>
    <table style="width: 100%;">
        <tr>
            <td valign="top" align="right" style="width: 50px;">
            </td>
            <td valign="top" align="right" style="width: 50px;">
                &nbsp;
            </td>
        </tr>
        <tr>
            <td colspan="8" align="left" valign="top" height="220">
                <%--Column : Grid--%>
                <dx:aspxgridview id="grid" runat="server" width="100%" font-names="Tahoma" keyfieldname="NoUrut;AffiliateID;POMarking"
                    autogeneratecolumns="False" clientinstancename="grid" font-size="8pt" tabindex="12">
                    <ClientSideEvents Init="OnInit" FocusedRowChanged="function(s, e) {OnGridFocusedRowChanged();}"
                        EndCallback="function(s, e) {
                        var pMsg = s.cpMessage;
                        if (pMsg != '') {
                            if (pMsg.substring(1,5) == '1001' || pMsg.substring(1,5) == '1002' || pMsg.substring(1,5) == '1003' || pMsg.substring(1,5) == '2001' || pMsg.substring(1,5) == '1008') {
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
	                    lblInfo.SetText('');}" BatchEditStartEditing="OnBatchEditStartEditing" />
                    <Columns>
                        <dx:GridViewDataTextColumn VisibleIndex="0" Caption=" " Width="60px" FieldName="DetailPage">
                            <DataItemTemplate>
                                <a id="clickElement" href="AffiliateOrderDetail.aspx?id=<%#GetRowValue(Container)%>&t1=<%#GetAffiliateID(Container)%>&t2=<%#GetSupplierID(Container)%>&t3=<%#GetPeriod(Container)%>&Session=~/AffiliateOrder/AffiliateOrderList.aspx">
                                    <%# "DETAIL"%></a>
                            </DataItemTemplate>
                        </dx:GridViewDataTextColumn>    
                        <dx:GridViewDataTextColumn VisibleIndex="1" Caption="NO" FieldName="NoUrut" Width="30px"
                            HeaderStyle-HorizontalAlign="Center">
                            <HeaderStyle HorizontalAlign="Center"></HeaderStyle>
                        </dx:GridViewDataTextColumn>
                        <dx:GridViewDataTextColumn VisibleIndex="2" Caption="PERIOD" FieldName="Period" Width="80px"
                            HeaderStyle-HorizontalAlign="Center">
                            <PropertiesTextEdit DisplayFormatString="MMM yyyy">
                            </PropertiesTextEdit>
                            <HeaderStyle HorizontalAlign="Center"></HeaderStyle>
                        </dx:GridViewDataTextColumn>
                        <dx:GridViewDataTextColumn VisibleIndex="3" Caption="AFFILIATE CODE" FieldName="AffiliateID"
                            Width="75px" HeaderStyle-HorizontalAlign="Center">
                            <HeaderStyle HorizontalAlign="Center"></HeaderStyle>
                        </dx:GridViewDataTextColumn>
                        <dx:GridViewDataTextColumn VisibleIndex="4" Caption="AFFILIATE NAME" FieldName="AffiliateName"
                            Width="200px" HeaderStyle-HorizontalAlign="Center">
                            <HeaderStyle HorizontalAlign="Center"></HeaderStyle>
                        </dx:GridViewDataTextColumn>
                        <dx:GridViewDataTextColumn VisibleIndex="5" Caption="PO NO." FieldName="PONo" Width="100px"
                            HeaderStyle-HorizontalAlign="Center">
                            <HeaderStyle HorizontalAlign="Center"></HeaderStyle>
                        </dx:GridViewDataTextColumn>
                        <dx:GridViewDataTextColumn VisibleIndex="6" Caption="PO MARKING" FieldName="POMarking" Width="100px"
                            HeaderStyle-HorizontalAlign="Center">
                            <HeaderStyle HorizontalAlign="Center"></HeaderStyle>
                        </dx:GridViewDataTextColumn>
                        <dx:GridViewDataTextColumn VisibleIndex="7" Caption="COMMERCIAL" FieldName="CommercialCls"
                            Width="100px" HeaderStyle-HorizontalAlign="Center">
                            <HeaderStyle HorizontalAlign="Center"></HeaderStyle>
                        </dx:GridViewDataTextColumn>
                        <dx:GridViewDataTextColumn VisibleIndex="8" Caption="SUPPLIER CODE" FieldName="SupplierID"
                            Width="100px" HeaderStyle-HorizontalAlign="Center">
                            <HeaderStyle HorizontalAlign="Center"></HeaderStyle>
                        </dx:GridViewDataTextColumn>
                        <dx:GridViewDataTextColumn VisibleIndex="9" Caption="SUPPLIER NAME" FieldName="SupplierName"
                            Width="100px" HeaderStyle-HorizontalAlign="Center">
                            <HeaderStyle HorizontalAlign="Center"></HeaderStyle>
                        </dx:GridViewDataTextColumn>
                        <dx:GridViewDataTextColumn VisibleIndex="10" Caption="SHIP BY" FieldName="ShipCls"
                            Width="75px" HeaderStyle-HorizontalAlign="Center">
                            <HeaderStyle HorizontalAlign="Center"></HeaderStyle>
                        </dx:GridViewDataTextColumn>                        
                        <dx:GridViewDataTextColumn VisibleIndex="13" Caption="CREATED DATE" FieldName="EntryDate"
                            HeaderStyle-HorizontalAlign="Center" Width="140px">
                            <PropertiesTextEdit DisplayFormatString="yyyy-MM-dd HH:mm:ss">
                            </PropertiesTextEdit>
                            <HeaderStyle HorizontalAlign="Center"></HeaderStyle>
                        </dx:GridViewDataTextColumn>
                        <dx:GridViewDataTextColumn VisibleIndex="14" Caption="CREATED USER" FieldName="EntryUser"
                            HeaderStyle-HorizontalAlign="Center">
                            <HeaderStyle HorizontalAlign="Center"></HeaderStyle>
                        </dx:GridViewDataTextColumn>
                        <dx:GridViewBandColumn Caption="PO STATUS" VisibleIndex="15">
                            <Columns>
                                <dx:GridViewDataCheckColumn Caption="1" FieldName="POStatus1" ReadOnly="True" VisibleIndex="1"
                                    Width="30px">
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
                                <dx:GridViewDataCheckColumn Caption="2" FieldName="POStatus2" ReadOnly="True" VisibleIndex="3"
                                    Width="30px">
                                    <PropertiesCheckEdit ValueChecked="1" ValueType="System.String" ValueUnchecked="0">
                                        
<DisplayImageChecked IconID="support_feature_16x16">
                                        </DisplayImageChecked>
                                        

<DisplayImageUnchecked Width="0px">
                                        </DisplayImageUnchecked>
                                    
</PropertiesCheckEdit>
                                    <HeaderStyle HorizontalAlign="Center"></HeaderStyle>
                                </dx:GridViewDataCheckColumn>
                                <dx:GridViewDataCheckColumn Caption="3" FieldName="POStatus3" ReadOnly="True" VisibleIndex="5"
                                    Width="30px">
                                    <PropertiesCheckEdit ValueChecked="1" ValueType="System.String" ValueUnchecked="0">
                                        
<DisplayImageChecked IconID="support_feature_16x16">
                                        </DisplayImageChecked>
                                        

<DisplayImageUnchecked Width="0px">
                                        </DisplayImageUnchecked>
                                    
</PropertiesCheckEdit>
                                    <HeaderStyle HorizontalAlign="Center"></HeaderStyle>
                                </dx:GridViewDataCheckColumn>
                                <dx:GridViewDataCheckColumn Caption="4" FieldName="POStatus4" ReadOnly="True" VisibleIndex="7"
                                    Width="30px">
                                    <PropertiesCheckEdit ValueChecked="1" ValueType="System.String" ValueUnchecked="0">
                                        
<DisplayImageChecked IconID="support_feature_16x16">
                                        </DisplayImageChecked>
                                        

<DisplayImageUnchecked Width="0px">
                                        </DisplayImageUnchecked>
                                    
</PropertiesCheckEdit>
                                </dx:GridViewDataCheckColumn>
                                <dx:GridViewDataCheckColumn Caption="5" FieldName="POStatus5" ReadOnly="True" VisibleIndex="9"
                                    Width="30px">
                                    <PropertiesCheckEdit ValueChecked="1" ValueType="System.String" ValueUnchecked="0">
                                        
<DisplayImageChecked IconID="support_feature_16x16">
                                        </DisplayImageChecked>
                                        

<DisplayImageUnchecked Width="0px">
                                        </DisplayImageUnchecked>
                                    
</PropertiesCheckEdit>
                                    <HeaderStyle HorizontalAlign="Center"></HeaderStyle>
                                </dx:GridViewDataCheckColumn>
                                <dx:GridViewDataCheckColumn Caption="6" FieldName="POStatus6" ReadOnly="True" VisibleIndex="11"
                                    Width="30px">
                                    <PropertiesCheckEdit ValueChecked="1" ValueType="System.String" ValueUnchecked="0">
                                        
<DisplayImageChecked IconID="support_feature_16x16">
                                        </DisplayImageChecked>
                                        

<DisplayImageUnchecked Width="0px">
                                        </DisplayImageUnchecked>
                                    
</PropertiesCheckEdit>
                                    <HeaderStyle HorizontalAlign="Center"></HeaderStyle>
                                </dx:GridViewDataCheckColumn>
                                <dx:GridViewDataCheckColumn Caption="7" FieldName="POStatus7" ReadOnly="True" VisibleIndex="13"
                                    Width="30px">
                                    <PropertiesCheckEdit ValueChecked="1" ValueType="System.String" ValueUnchecked="0">
                                        
<DisplayImageChecked IconID="support_feature_16x16">
                                        </DisplayImageChecked>
                                        

<DisplayImageUnchecked Width="0px">
                                        </DisplayImageUnchecked>
                                    
</PropertiesCheckEdit>
                                    <HeaderStyle HorizontalAlign="Center"></HeaderStyle>
                                </dx:GridViewDataCheckColumn>
                                <dx:GridViewDataCheckColumn Caption="8" FieldName="POStatus8" ReadOnly="True" VisibleIndex="14"
                                    Width="30px">
                                    <PropertiesCheckEdit ValueChecked="1" ValueType="System.String" ValueUnchecked="0">
                                        
<DisplayImageChecked IconID="support_feature_16x16">
                                        </DisplayImageChecked>
                                        

<DisplayImageUnchecked Width="0px">
                                        </DisplayImageUnchecked>
                                    
</PropertiesCheckEdit>
                                    <HeaderStyle HorizontalAlign="Center"></HeaderStyle>
                                </dx:GridViewDataCheckColumn>
                            </Columns>
                            <HeaderStyle HorizontalAlign="Center"></HeaderStyle>
                        </dx:GridViewBandColumn>
                    </Columns>                    
                    <SettingsBehavior AllowSort="False" AllowSelectByRowClick="True" ColumnResizeMode="Control" 
                        EnableRowHotTrack="True">
                    </SettingsBehavior>
                    <SettingsPager Visible="False" PageSize="13" Position="Top">
                        <Summary Text="Page {0} of {1} [{2} record(s)]" AllPagesText="Page {0} of {1} " />                        
                    </SettingsPager>
                    <SettingsEditing Mode="Batch" NewItemRowPosition="Bottom">
                        <BatchEditSettings ShowConfirmOnLosingChanges="False" />
                    </SettingsEditing>
                    <Settings ShowGroupButtons="False" ShowVerticalScrollBar="True" ShowHorizontalScrollBar="True" 
                        VerticalScrollableHeight="190" ShowStatusBar="Hidden">
                    </Settings>
                    <SettingsCommandButton EditButton-ButtonType="Link">
                        <EditButton Text="Detail">
                        </EditButton>
                    </SettingsCommandButton>
                    <StylesEditors ButtonEditCellSpacing="0">
                        <ProgressBar Height="21px">
                        </ProgressBar>
                    </StylesEditors>
                </dx:aspxgridview>
            </td>
        </tr>
    </table>
    <div style="height: 8px;">
    </div>
    <table id="button" style="width: 100%;">
        <tr>
            <td valign="top" align="left">
                <dx:aspxbutton id="btnSubMenu" runat="server" text="SUB MENU" font-names="Tahoma"
                    width="85px" font-size="8pt" tabindex="13">
                </dx:aspxbutton>
            </td>
        </tr>
    </table>
    <dx:aspxglobalevents id="ge" runat="server">
        <ClientSideEvents ControlsInitialized="function(s, e) {
	        OnControlsInitializedSplitter();
	        OnControlsInitializedGrid();
        }" />
    </dx:aspxglobalevents>
</asp:Content>
