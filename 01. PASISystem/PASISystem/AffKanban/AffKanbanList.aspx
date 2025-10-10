<%@ Page Title="" Language="vb" AutoEventWireup="false" MasterPageFile="~/SiteContent.Master"
    CodeBehind="AffKanbanList.aspx.vb" Inherits="PASISystem.AffKanbanList" %>

<%@ Register Assembly="DevExpress.Web.v14.1, Version=14.1.6.0, Culture=neutral, PublicKeyToken=b88d1754d700e49a"
    Namespace="DevExpress.Web.ASPxEditors" TagPrefix="dx" %>
<%@ Register Assembly="DevExpress.Web.v14.1, Version=14.1.6.0, Culture=neutral, PublicKeyToken=b88d1754d700e49a"
    Namespace="DevExpress.Web.ASPxPager" TagPrefix="dx" %>
<%@ Register Assembly="DevExpress.Web.v14.1, Version=14.1.6.0, Culture=neutral, PublicKeyToken=b88d1754d700e49a"
    Namespace="DevExpress.Web.ASPxGridView" TagPrefix="dx" %>
<%@ Register Assembly="DevExpress.Web.v14.1, Version=14.1.6.0, Culture=neutral, PublicKeyToken=b88d1754d700e49a"
    Namespace="DevExpress.Web.ASPxRoundPanel" TagPrefix="dx" %>
<%@ Register Assembly="DevExpress.Web.v14.1, Version=14.1.6.0, Culture=neutral, PublicKeyToken=b88d1754d700e49a"
    Namespace="DevExpress.Web.ASPxPanel" TagPrefix="dx" %>

<asp:Content ID="Content1" ContentPlaceHolderID="head" runat="server">
    <style type="text/css">
        .style6
        {
            width: 137px;
        }
        .style11
        {
            width: 84px;
        }
        .style12
        {
            width: 173px;
        }
        
        .dxflEmptyItem
        {
            height: 21px;
        }
        
        .style25
        {
            width: 1001px;
            height: 20px;
        }
        .style26
        {
            width: 708px;
        }
        .style28
        {
            width: 2px;
        }
        .style29
        {
            width: 179px;
        }
    </style>
</asp:Content>
<asp:Content ID="Content2" ContentPlaceHolderID="Content" runat="server">
    <script language="javascript" type="text/javascript">
        function OnUpdateClick(s, e) {
            Grid.PerformCallback("Update");
        }

        function OnCancelClick(s, e) {
            Grid.PerformCallback("Cancel");
        }

        function OnBatchEditStartEditing(s, e) {
            currentColumnName = e.focusedColumn.fieldName;
            if (currentColumnName == "cols" ) {

                e.cancel = false;
            } else {

                e.cancel = true;
            }


            currentEditableVisibleIndex = e.visibleIndex;
        }
    </script>
    <table align="center" width="100%">
        <tr>
            <td align="left" class="style26" width="100%">
                <table style="border: thin solid #808080; width: 100%;" width="100%">
                    <tr>
                        <td class="style11" align="left" height="26px">
                            <dx:ASPxLabel ID="ASPxLabel1" runat="server" Font-Names="Tahoma" Font-Size="8pt"
                                Text="KANBAN DATE">
                            </dx:ASPxLabel>
                        </td>
                        <td class="style6" rowspan="1" height="26px">
                            <dx:ASPxDateEdit ID="dt1" runat="server" Font-Names="Tahoma" Font-Size="8pt" EditFormat="Custom"
                                EditFormatString="dd MMM yyyy" ClientInstanceName="dt1">
                                <ClientSideEvents ValueChanged="function(s, e) {
	                                            var pDateFrom = dt1.GetText();
	                                            var pAffiliate = cboaffiliate.GetValue();
	                                            var pDateTo = dt2.GetText();
                                                var pSupplier = cbosupplier.GetText();
                                                var pLocation = cbolocation.GetText();
                                                
												cbokanbanno.PerformCallback(pDateFrom + '|' + pDateTo + '|' + pSupplier + '|' + pLocation + '|' + pAffiliate);
}" />
                            </dx:ASPxDateEdit>
                        </td>
                        <td align="center" class="style28" height="26px">
                            <dx:ASPxLabel ID="ASPxLabel3" runat="server" Font-Names="Tahoma" Font-Size="8pt"
                                Text="~">
                            </dx:ASPxLabel>
                        </td>
                        <td class="style12" align="left" height="26px">
                            <dx:ASPxDateEdit ID="dt2" runat="server" Font-Names="Tahoma" Font-Size="8pt" EditFormat="Custom"
                                EditFormatString="dd MMM yyyy" ClientInstanceName="dt2">
                                <ClientSideEvents ValueChanged="function(s, e) {
		                                        var pDateFrom = dt1.GetText();
	                                            var pAffiliate = cboaffiliate.Getvalue();
	                                            var pDateTo = dt2.GetText();
                                                var pSupplier = cbosupplier.GetText();
                                                var pLocation = cbolocation.GetText();
												cbokanbanno.PerformCallback(pDateFrom + '|' + pDateTo + '|' + pSupplier + '|' + pLocation + '|' + pAffiliate);
}" />
                            </dx:ASPxDateEdit>
                        </td>
                        <td align="left" class="style29" height="26px">
                        </td>
                        <td align="left" class="style12" height="26px">
                            &nbsp;
                        </td>
                    </tr>
                    <tr>
                        <td class="style11" align="left" height="26px">
                            <dx:ASPxLabel ID="ASPxLabel14" runat="server" Font-Names="Tahoma" Font-Size="8pt"
                                Text="AFFILIATE CODE/NAME" Width="137px">
                            </dx:ASPxLabel>
                        </td>
                        <td class="style6" align="left" height="26px">
                            <dx:ASPxComboBox ID="cboaffiliate" runat="server" ClientInstanceName="cboaffiliate"
                                Font-Names="Tahoma" TextFormatString="{0}" Font-Size="8pt">
                                <ClientSideEvents SelectedIndexChanged="function(s, e) {
	                                            txtaffiliate.SetText(cboaffiliate.GetSelectedItem().GetColumnText(1));
                                                var pDateFrom = dt1.GetText();
	                                            var pAffiliate = cboaffiliate.GetValue().toString();
	                                            var pDateTo = dt2.GetText();
                                                var pSupplier = cbosupplier.GetText();
                                                var pLocation = cbolocation.GetText();
												cbolocation.PerformCallback(cboaffiliate.GetValue().toString());
												cbokanbanno.PerformCallback(pDateFrom + '|' + pDateTo + '|' + pSupplier + '|' + pLocation + '|' + pAffiliate);

												
                                                }" />
                            </dx:ASPxComboBox>
                        </td>
                        <td align="center" class="style28" height="26px">
                            &nbsp;</td>
                        <td align="left" class="style12" height="26px">
                            <dx:ASPxTextBox ID="txtaffiliate" runat="server" Width="230px" Font-Names="Tahoma"
                                Font-Size="8pt" BackColor="#CCCCCC" ReadOnly="True" Height="16px" 
                                ClientInstanceName="txtaffiliate">
                                <ClientSideEvents TextChanged="function(s, e) {
	                                lblerrmessage.SetText('');
                                }" />
                            </dx:ASPxTextBox>
                        </td>
                        <td align="right" class="style29" height="26px">
                            &nbsp;</td>
                        <td align="left" class="style12" height="26px">
                            &nbsp;</td>
                    </tr>
                    <tr>
                        <td class="style11" align="left" height="26px">
                            <dx:ASPxLabel ID="ASPxLabel2" runat="server" Font-Names="Tahoma" Font-Size="8pt"
                                Text="DELIVERY LOCATION" Width="137px">
                            </dx:ASPxLabel>
                        </td>
                        <td class="style6" rowspan="1" align="left" height="26px">
                            <dx:ASPxComboBox ID="cbolocation" runat="server" ClientInstanceName="cbolocation"
                                Font-Names="Tahoma" TextFormatString="{0}" Font-Size="8pt">
                                <ClientSideEvents SelectedIndexChanged="function(s, e) {
												txtlocation.SetText(cbolocation.GetSelectedItem().GetColumnText(1));
                                                
												var pDateFrom = dt1.GetText();
												var pAffiliate = cboaffiliate.GetText();
	                                            var pDateTo = dt2.GetText();
                                                var pSupplier = cbosupplier.GetText();
                                                var pLocation = cbolocation.GetValue().toString();
													cbokanbanno.PerformCallback(pDateFrom + '|' + pDateTo + '|' + pSupplier + '|' + pLocation + '|' + pAffiliate);
                                                }" />
                            </dx:ASPxComboBox>
                        </td>
                        <td align="center" class="style28" height="26px">
                            &nbsp;
                        </td>
                        <td align="left" class="style12" height="26px">

                            <dx:ASPxTextBox ID="txtlocation" runat="server" Width="230px" Font-Names="Tahoma"
                                Font-Size="8pt" BackColor="#CCCCCC" ReadOnly="True" Height="16px" 
                                ClientInstanceName="txtlocation">
                                <ClientSideEvents TextChanged="function(s, e) {
	                                lblerrmessage.SetText('');
                                }" />
                            </dx:ASPxTextBox>

                        </td>
                        <td align="right" class="style29" height="26px">
                            &nbsp;
                        </td>
                        <td align="left" class="style12" height="26px">
                            &nbsp;
                        </td>
                    </tr>
                    <tr>
                        <td class="style11" align="left" height="26px">
                            <dx:ASPxLabel ID="ASPxLabel13" runat="server" Font-Names="Tahoma" Font-Size="8pt"
                                Text="SUPPLIER CODE/NAME" Width="137px">
                            </dx:ASPxLabel>
                        </td>
                        <td class="style6" align="left" height="26px">
                            <dx:ASPxComboBox ID="cbosupplier" runat="server" ClientInstanceName="cbosupplier"
                                Font-Names="Tahoma" TextFormatString="{0}" Font-Size="8pt">
                                <ClientSideEvents SelectedIndexChanged="function(s, e) {
	                                            txtsupplier.SetText(cbosupplier.GetSelectedItem().GetColumnText(1));
	var pDateFrom = dt1.GetText();
	var pAffiliate = cboaffiliate.Getvalue();
	                                            var pDateTo = dt2.GetText();
                                                var pSupplier = cbosupplier.GetText();
                                                var pLocation = cbolocation.GetText();
													cbokanbanno.PerformCallback(pDateFrom + '|' + pDateTo + '|' + pSupplier + '|' + pLocation + '|' + pAffiliate);
                                                }" />
                            </dx:ASPxComboBox>
                        </td>
                        <td align="center" class="style28" height="26px">
                            &nbsp;
                        </td>
                        <td align="left" class="style12" height="26px">
                            <dx:ASPxTextBox ID="txtsupplier" runat="server" Width="230px" Font-Names="Tahoma"
                                Font-Size="8pt" BackColor="#CCCCCC" ReadOnly="True" Height="16px" 
                                ClientInstanceName="txtsupplier">
                                <ClientSideEvents TextChanged="function(s, e) {
	                                lblerrmessage.SetText('');
                                }" />
                            </dx:ASPxTextBox>

                        </td>
                        <td align="right" class="style29" height="26px" width="85px">
                            <dx:ASPxButton ID="btnsearch" runat="server" Text="SEARCH" Width="85px" Font-Names="Tahoma"
                                Font-Size="8pt" ClientInstanceName="btnsearch" AutoPostBack="False">
                                <ClientSideEvents Click="function(s, e) {
	                                        var pDateFrom = dt1.GetText();
	                                        var pDateTo = dt2.GetText();
                                            var pSupplier = cbosupplier.GetText();
                                                            
	                                        grid.PerformCallback('gridload' + '|' + pDateFrom + '|' + pDateTo + '|' + pSupplier);
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
                            </dx:ASPxButton>
                        </td>
                        <td align="left" class="style12" height="26px" width="85px">
                            <dx:ASPxButton ID="btnclear" runat="server" Text="CLEAR" Width="85px" Font-Names="Tahoma"
                                Font-Size="8pt">
                            </dx:ASPxButton>
                        </td>
                    </tr>
                    <tr>
                        <td class="style11" align="left" height="26px">
                            <dx:ASPxLabel ID="ASPxLabel15" runat="server" Font-Names="Tahoma" 
                                Font-Size="8pt" Text="KANBAN NO." Width="137px">
                            </dx:ASPxLabel>
                        </td>
                        <td class="style6" align="left" height="26px">
                            <dx:ASPxComboBox ID="cbokanbanno" runat="server" 
                                ClientInstanceName="cbokanbanno" Font-Names="Tahoma" Font-Size="8pt" 
                                TextFormatString="{0}">
                            </dx:ASPxComboBox>
                        </td>
                        <td align="center" class="style28" height="26px">
                            &nbsp;</td>
                        <td align="left" class="style12" height="26px">
                            &nbsp;</td>
                        <td align="right" class="style29" height="26px" width="85px">
                            &nbsp;</td>
                        <td align="left" class="style12" height="26px" width="85px">
                            &nbsp;</td>
                    </tr>
                </table>
            </td>
            <td width="100%">
                <dx:ASPxRoundPanel ID="ASPxRoundPanel1" runat="server" HeaderText="KANBAN STATUS"
                    ShowCollapseButton="true" View="GroupBox" Width="100%">
                    <PanelCollection>
                        <dx:PanelContent runat="server">
                            <table width="100%">
                                <tr>
                                    <td align="left" colspan="1" height="17px" rowspan="1">
                                        <dx:ASPxLabel ID="ASPxLabel12" runat="server" Font-Names="Tahoma" 
                                            Font-Size="8pt" Text="(1) AFFILIATE ENTRY">
                                        </dx:ASPxLabel>
                                    </td>
                                </tr>
                                <tr>
                                    <td align="left" colspan="1" height="17px">
                                        <dx:ASPxLabel ID="ASPxLabel10" runat="server" Font-Names="Tahoma" Font-Size="8pt"
                                            Text="(2) AFFILIATE APPROVAL">
                                        </dx:ASPxLabel>
                                    </td>
                                </tr>
                                <tr>
                                    <td align="left" colspan="1" height="17px" rowspan="1">
                                        <dx:ASPxLabel ID="ASPxLabel11" runat="server" Font-Names="Tahoma" Font-Size="8pt"
                                            Text="(3) SUPPLIER APPROVAL">
                                        </dx:ASPxLabel>
                                    </td>
                                </tr>
                                <tr>
                                    <td align="left" colspan="1" height="17px">
                                        &nbsp;</td>
                                </tr>
                            </table>
                        </dx:PanelContent>
                    </PanelCollection>
                </dx:ASPxRoundPanel>
            </td>
        </tr>
    </table>
    <table width="100%">
        <td align="left" bgcolor="White" class="style25" width="100%">
            <table align="left" width="100%">
                <tr align="left">
                    <td width="100%" height="16px" style="border-top-style: solid; border-top-width: thin;
                        border-top-color: #808080; border-bottom-style: solid; border-bottom-width: thin;
                        border-bottom-color: #808080" align="left">
                        <dx:ASPxLabel ID="lblerrmessage" runat="server" Font-Names="Tahoma" Font-Size="8pt"
                            Text="ERROR MESSAGE" Width="100%" ClientInstanceName="lblerrmessage">
                        </dx:ASPxLabel>
                    </td>
                </tr>
            </table>
        </td>
    </table>
    <table style="width: 100%; height: 195px;" align="left">
        <tr>
            <td align="right" width="100%">
                <dx:ASPxButton ID="btncreatekanban" runat="server" Text="CREATE NEW KANBAN" Width="90px"
                    Font-Names="Tahoma" Font-Size="8pt" Enabled="False" Visible="False">
                </dx:ASPxButton>
            </td>
        </tr>
        <tr>
            <td align="left" width="100%" height="0px">
                <dx:ASPxGridView ID="grid" runat="server" AutoGenerateColumns="False" Width="100%"
                    KeyFieldName="colno;colkanbandate" ClientInstanceName="grid">
<ClientSideEvents RowClick="function(s, e) {
	                    lblerrmessage.SetText(&#39;&#39;);}" 
                        BatchEditStartEditing="OnBatchEditStartEditing" EndCallback="function(s, e) {
  
                        var pMsg = s.cpMessage;
                        if (pMsg != &#39;&#39;) {
                            if (pMsg.substring(1,5) == &#39;1001&#39; || pMsg.substring(1,5) == &#39;1002&#39; || pMsg.substring(1,5) == &#39;1003&#39; || pMsg.substring(1,5) == &#39;2001&#39;) {
                                lblerrmessage.GetMainElement().style.color = &#39;Blue&#39;;
                            } else {
                                lblerrmessage.GetMainElement().style.color = &#39;Red&#39;;
                            }
                            lblerrmessage.SetText(pMsg);
                        } else {
                            lblerrmessage.SetText(&#39;&#39;);
                        }
                        delete s.cpMessage;
                        }" CallbackError="function(s, e) {e.handled = true;}" Init="function(s, e) {
	dt1.SetText(s.cpdt1);
    dt2.SetText(s.cpdt2);
}"></ClientSideEvents>

<Settings ShowGroupButtons="False" ShowVerticalScrollBar="True" ShowHorizontalScrollBar="True" 
                        VerticalScrollableHeight="250" ShowStatusBar="Hidden"></Settings>

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
	                    lblerrmessage.SetText('');}" Init="function(s, e) {
	dt1.SetText(s.cpdt1);
    dt2.SetText(s.cpdt2); 
}" BatchEditStartEditing="OnBatchEditStartEditing" 
                        CallbackError="function(s, e) {e.handled = true;}" />
                    <Columns>
                        <dx:GridViewDataHyperLinkColumn Caption="  " FieldName="coldetailurl" Name="coldetail"
                            VisibleIndex="0" Width="60px">
                            <PropertiesHyperLinkEdit TextField="coldetailname">
                            </PropertiesHyperLinkEdit>
                            <HeaderStyle Font-Names="Tahoma" Font-Size="8pt" HorizontalAlign="Center" VerticalAlign="Middle"
                                Wrap="True" />
                            <CellStyle Font-Names="Tahoma" Font-Size="8pt" Font-Underline="True" VerticalAlign="Middle">
                            </CellStyle>
                        </dx:GridViewDataHyperLinkColumn>
                        <dx:GridViewDataCheckColumn FieldName="cols" Name="cols" VisibleIndex="1" 
                            Width="30px" Caption=" ">
                            <HeaderStyle HorizontalAlign="Center" VerticalAlign="Middle" />
                            <PropertiesCheckEdit ValueChecked="1" ValueType="System.Int32" ValueUnchecked="0">
                            </PropertiesCheckEdit>
                        </dx:GridViewDataCheckColumn>
                        <dx:GridViewDataTextColumn Caption="NO" FieldName="colno" Name="colno" VisibleIndex="2"
                            Width="30px">
                            <HeaderStyle Font-Names="Tahoma" Font-Size="8pt" HorizontalAlign="Center" VerticalAlign="Middle"
                                Wrap="True" />
                            <CellStyle Font-Names="Tahoma" Font-Size="8pt" HorizontalAlign="Center" VerticalAlign="Middle">
                            </CellStyle>
                        </dx:GridViewDataTextColumn>
                        <dx:GridViewDataTextColumn Caption="SUPPLIER CODE" FieldName="colsuppliercode" Name="colsuppliercode"
                            VisibleIndex="4" Width="100px">
                            <HeaderStyle Font-Names="Tahoma" Font-Size="8pt" HorizontalAlign="Center" VerticalAlign="Middle"
                                Wrap="True" />
                            <CellStyle Font-Names="Tahoma" Font-Size="8pt">
                            </CellStyle>
                        </dx:GridViewDataTextColumn>
                        <dx:GridViewDataTextColumn Caption="SUPPLIER NAME" FieldName="colsuppliername" Name="colsuppliername"
                            VisibleIndex="5" Width="120px">
                            <HeaderStyle Font-Names="Tahoma" Font-Size="8pt" HorizontalAlign="Center" VerticalAlign="Middle"
                                Wrap="True" />
                            <CellStyle Font-Names="Tahoma" Font-Size="8pt">
                            </CellStyle>
                        </dx:GridViewDataTextColumn>
                        <dx:GridViewDataTextColumn Caption="KANBAN DATE" FieldName="colkanbandate" Name="colkanbandate"
                            VisibleIndex="8" Width="85px">
                            <HeaderStyle Font-Names="Tahoma" Font-Size="8pt" HorizontalAlign="Center" VerticalAlign="Middle"
                                Wrap="True" />
                            <CellStyle Font-Names="Tahoma" Font-Size="8pt">
                            </CellStyle>
                        </dx:GridViewDataTextColumn>
                        <dx:GridViewDataTextColumn Caption="CREATED DATE" FieldName="colcreateddate" Name="colcreateddate"
                            VisibleIndex="9" Width="130px">
                            <HeaderStyle Font-Names="Tahoma" Font-Size="8pt" HorizontalAlign="Center" VerticalAlign="Middle"
                                Wrap="True" />
                            <CellStyle Font-Names="Tahoma" Font-Size="8pt">
                            </CellStyle>
                        </dx:GridViewDataTextColumn>
                        <dx:GridViewDataTextColumn Caption="CREATED BY" FieldName="colcreatedby" Name="colcreatedby"
                            VisibleIndex="10" Width="100px">
                            <HeaderStyle Font-Names="Tahoma" Font-Size="8pt" HorizontalAlign="Center" VerticalAlign="Middle"
                                Wrap="True" />
                            <CellStyle Font-Names="Tahoma" Font-Size="8pt">
                            </CellStyle>
                        </dx:GridViewDataTextColumn>
                        <dx:GridViewDataTextColumn Caption=" KANBAN NO." FieldName="colkanbanno" Name="colkanbanno"
                            VisibleIndex="3" Width="100px">
                        </dx:GridViewDataTextColumn>
                        
                        <dx:GridViewBandColumn Caption="KANBAN STATUS" Name="colHkanban" 
                            VisibleIndex="11">
                            <Columns>
                                <dx:GridViewDataCheckColumn Caption="(1)" FieldName="colkanbanstatus1" 
                                    Name="colkanbanstatus1" VisibleIndex="1" Width="30px">
                                    <PropertiesCheckEdit ValueChecked="1" ValueType="System.String" 
                                        ValueUnchecked="0">
                                        
<DisplayImageChecked IconID="support_feature_16x16">
                                        </DisplayImageChecked>
                                        

<DisplayImageUnchecked Width="0px">
                                        </DisplayImageUnchecked>
                                    
</PropertiesCheckEdit>
                                </dx:GridViewDataCheckColumn>
                                <dx:GridViewDataCheckColumn Caption="(2)" FieldName="colkanbanstatus2" 
                                    Name="colkanbanstatus2" ToolTip="(2)" VisibleIndex="2" Width="30px">
                                    <PropertiesCheckEdit ValueChecked="1" ValueType="System.String" 
                                        ValueUnchecked="0">
                                        
<DisplayImageChecked IconID="support_feature_16x16">
                                        </DisplayImageChecked>
                                        

<DisplayImageUnchecked Width="0px">
                                        </DisplayImageUnchecked>
                                    
</PropertiesCheckEdit>
                                </dx:GridViewDataCheckColumn>
                                <dx:GridViewDataCheckColumn Caption="(3)" FieldName="colkanbanstatus3" 
                                    Name="colkanbanstatus3" VisibleIndex="3" Width="30px">
                                    <PropertiesCheckEdit ValueChecked="1" ValueType="System.String" 
                                        ValueUnchecked="0">
                                        
<DisplayImageChecked IconID="support_feature_16x16">
                                        </DisplayImageChecked>
                                        

<DisplayImageUnchecked Width="0px">
                                        </DisplayImageUnchecked>
                                    
</PropertiesCheckEdit>
                                </dx:GridViewDataCheckColumn>
                            </Columns>
                            <HeaderStyle Font-Names="Tahoma" Font-Size="8pt" HorizontalAlign="Center" VerticalAlign="Middle"
                                Wrap="True" />
                        </dx:GridViewBandColumn>
                        <dx:GridViewDataTextColumn Caption="DELIVERY LOCATION CODE" 
                            FieldName="coldeliverycode" Name="coldeliverycode" VisibleIndex="6" 
                            Width="100px">
                            <HeaderStyle Font-Names="Tahoma" Font-Size="8pt" VerticalAlign="Middle" 
                                Wrap="True" />
                            <CellStyle Font-Names="Tahoma" Font-Size="8pt">
                            </CellStyle>
                        </dx:GridViewDataTextColumn>
                        <dx:GridViewDataTextColumn Caption="DELIVERY LOCATION NAME" 
                            FieldName="coldeliveryname" Name="coldeliveryname" VisibleIndex="7" 
                            Width="120px">
                            <HeaderStyle Font-Names="Tahoma" Font-Size="8pt" VerticalAlign="Middle" 
                                Wrap="True" />
                            <CellStyle Font-Names="Tahoma" Font-Size="8pt">
                            </CellStyle>
                        </dx:GridViewDataTextColumn>
                    </Columns>
                    
                    <SettingsPager Mode="ShowAllRecords" Visible="False">
                    </SettingsPager>
                    
                    <SettingsEditing Mode="Batch" NewItemRowPosition="Bottom">
                        <BatchEditSettings ShowConfirmOnLosingChanges="False" />
                        <BatchEditSettings ShowConfirmOnLosingChanges="False"></BatchEditSettings>
                    </SettingsEditing>
                    <Settings ShowHorizontalScrollBar="True" ShowVerticalScrollBar="True" ShowGroupButtons="False"
                        ShowStatusBar="Hidden" VerticalScrollableHeight="250" />

                    <StylesEditors ButtonEditCellSpacing="0">
                        <ProgressBar Height="21px">
                        </ProgressBar>
                    </StylesEditors>
                </dx:ASPxGridView>
            </td>
        </tr>
    </table>
    <table style="width: 100%;" width="100%">
        <tr>
            <td align="left">
                <dx:ASPxButton ID="btnsubmenu" runat="server" Text="SUB MENU" Width="90px" Font-Names="Tahoma"
                    Font-Size="8pt">
                </dx:ASPxButton>
            </td>
            <td align="right">
                <dx:ASPxButton ID="btnprintcard" runat="server" Text="PRINT KANBAN CARD" Width="90px"
                    Font-Names="Tahoma" Font-Size="8pt" AutoPostBack="False">
                    <ClientSideEvents Click="function(s, e) {
	grid.UpdateEdit();
	grid.PerformCallback('PrintCard');
	grid.CancelEdit();
}" />
                </dx:ASPxButton>
                &nbsp;<dx:ASPxButton ID="btnprintcycle" runat="server" Text="PRINT KANBAN CYCLE"
                    Width="90px" Font-Names="Tahoma" Font-Size="8pt" AutoPostBack="False">
                    <ClientSideEvents Click="function(s, e) {
	grid.UpdateEdit();
	grid.PerformCallback('PrintCycle');
	grid.CancelEdit();
}" />
                </dx:ASPxButton>
            </td>
        </tr>
    </table>
    <br />
</asp:Content>
