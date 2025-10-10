<%@ Page Title="" Language="vb" AutoEventWireup="false" MasterPageFile="~/SiteContent.Master"
    CodeBehind="AffKanbanCreate.aspx.vb" Inherits="PASISystem.AffKanbanCreate" %>

<%@ Register Assembly="DevExpress.Web.v14.1, Version=14.1.6.0, Culture=neutral, PublicKeyToken=b88d1754d700e49a"
    Namespace="DevExpress.Web.ASPxEditors" TagPrefix="dx" %>
<%@ Register Assembly="DevExpress.Web.v14.1, Version=14.1.6.0, Culture=neutral, PublicKeyToken=b88d1754d700e49a"
    Namespace="DevExpress.Web.ASPxPager" TagPrefix="dx" %>
<%@ Register Assembly="DevExpress.Web.v14.1, Version=14.1.6.0, Culture=neutral, PublicKeyToken=b88d1754d700e49a"
    Namespace="DevExpress.Web.ASPxGridView" TagPrefix="dx" %>
<%@ Register Assembly="DevExpress.Web.v14.1, Version=14.1.6.0, Culture=neutral, PublicKeyToken=b88d1754d700e49a"
    Namespace="DevExpress.Web.ASPxPanel" TagPrefix="dx" %>
<%@ Register Assembly="DevExpress.Web.v14.1, Version=14.1.6.0, Culture=neutral, PublicKeyToken=b88d1754d700e49a"
    Namespace="DevExpress.Web.ASPxRoundPanel" TagPrefix="dx" %>
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
        .style34
        {
            width: 188px;
        }
        .style35
        {
            width: 145px;
        }
        .style36
        {
            width: 36px;
        }
        .style43
        {
            width: 5%;
        }
        .style44
        {
            width: 4%;
        }
        .style45
        {
            width: 14px;
        }
        .style40
        {
            width: 6%;
        }
        .style47
        {
            width: 84px;
            height: 7px;
        }
        .style48
        {
            width: 137px;
            height: 7px;
        }
        .style49
        {
            width: 173px;
            height: 7px;
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

        var currentColumnName;
        function OnBatchEditStartEditing(s, e) {
            currentColumnName = e.focusedColumn.fieldName;
            if (currentColumnName == "colno" || currentColumnName == "coldescription" || currentColumnName == "colpartno" || currentColumnName == "colpono"
            || currentColumnName == "coluom" || currentColumnName == "colqty" || currentColumnName == "colmoq" || currentColumnName == "colpoqty"
            || currentColumnName == "colremainingpo" || currentColumnName == "colremainingsupplier" || currentColumnName == "coldeliveryqty" || currentColumnName == "colbox"
            ) {
                
                e.cancel = true;
            } else {

                e.cancel = true;
            }


            currentEditableVisibleIndex = e.visibleIndex;
        }

        function OnBatchEditEndEditing(s, e) {
            window.setTimeout(function () {
                //                if (s.batchEditApi.GetCellValue(e.visibleIndex, "cols") = "1") {
//                                alert('diceklist')
                var kanbanqty = s.batchEditApi.GetCellValue(e.visibleIndex, "colkanbanqty");
                var cycle1 = s.batchEditApi.GetCellValue(e.visibleIndex, "colcycle1");
                var cycle2 = s.batchEditApi.GetCellValue(e.visibleIndex, "colcycle2");
                var cycle3 = s.batchEditApi.GetCellValue(e.visibleIndex, "colcycle3");
                var cycle4 = s.batchEditApi.GetCellValue(e.visibleIndex, "colcycle4");

                    if (currentcolumnname = "colkanbanqty") {
//                        alert(currentcolumnname);
                        var num = kanbanqty / 4;
                        n = Number(num.toString().match(/^\d+(?:\.\d{0,0})?/))
                        s.batchEditApi.SetCellValue(e.visibleIndex, "colcycle1", n);
                        s.batchEditApi.SetCellValue(e.visibleIndex, "colcycle2", n);
                        s.batchEditApi.SetCellValue(e.visibleIndex, "colcycle3", n);
                        s.batchEditApi.SetCellValue(e.visibleIndex, "colcycle4", n + (kanbanqty % 4));
                    }
            }, 10);
        } 

    </script>
    <table align="center" width="100%">
        <tr>
            <td align="left" class="style26" width="100%">
                <table style="border: 0.1px solid #808080;" width="100%">
                    <tr>
                        <td class="style11" align="left" height="25px" width="80%">
                            <dx:ASPxLabel ID="ASPxLabel2" runat="server" Font-Names="Tahoma" Font-Size="8pt"
                                Text="KANBAN DATE" Width="100px">
                            </dx:ASPxLabel>
                        </td>
                        <td class="style6" rowspan="1" height="25px" width="80%">
                            <dx:ASPxTextBox ID="dtkanban" runat="server" Width="170px" Font-Names="Tahoma"
                                Font-Size="8pt" BackColor="#CCCCCC" ReadOnly="True" ClientInstanceName="dtkanban"
                                CssPostfix="DevEx" Height="16px">
                            </dx:ASPxTextBox>
                        </td>
                        <td class="style12" align="left" height="25px" width="80%">
                            <table style="width:100%;">
                                <tr>
                                    <td>
                            <dx:ASPxLabel ID="ASPxLabel14" runat="server" Font-Names="Tahoma" Font-Size="8pt"
                                Text="AFFILIATE CODE/NAME" Width="137px">
                            </dx:ASPxLabel>
                                    </td>
                                    <td>

                            <dx:ASPxTextBox ID="txtaffiliatecode" runat="server" Width="90px" Font-Names="Tahoma"
                                Font-Size="8pt" BackColor="#CCCCCC" ReadOnly="True" Height="16px" 
                                ClientInstanceName="txtaffiliatecode">
                                <ClientSideEvents TextChanged="function(s, e) {
	                                lblerrmessage.SetText('');
                                }" />
                            </dx:ASPxTextBox>

                                    </td>
                                </tr>
                            </table>
                        </td>
                        <td class="style12" align="left" height="25px" width="80%">

                            <dx:ASPxTextBox ID="txtaffiliatename" runat="server" Width="250px" Font-Names="Tahoma"
                                Font-Size="8pt" BackColor="#CCCCCC" ReadOnly="True" Height="16px" 
                                ClientInstanceName="txtaffiliatename">
                                <ClientSideEvents TextChanged="function(s, e) {
	                                lblerrmessage.SetText('');
                                }" />
                            </dx:ASPxTextBox>

                            </td>
                    </tr>
                    <tr>
                        <td class="style11" align="left" height="25px" width="80%">
                            <dx:ASPxLabel ID="ASPxLabel3" runat="server" Font-Names="Tahoma" Font-Size="8pt"
                                Text="SUPPLIER CODE/NAME" Width="137px">
                            </dx:ASPxLabel>
                        </td>
                        <td class="style6" rowspan="1" align="left" height="25px" width="80%">

                            <dx:ASPxTextBox ID="cbosupplier" runat="server" Width="170px" Font-Names="Tahoma"
                                Font-Size="8pt" BackColor="#CCCCCC" ReadOnly="True" Height="16px" 
                                ClientInstanceName="cbosupplier">
                                <ClientSideEvents TextChanged="function(s, e) {
	                                lblerrmessage.SetText('');
                                }" />
                            </dx:ASPxTextBox>

                        </td>
                        <td align="left" class="style12" height="25px" width="80%">
                            <dx:ASPxTextBox ID="txtsuppliername" runat="server" Width="250px" Font-Names="Tahoma"
                                Font-Size="8pt" BackColor="#CCCCCC" ReadOnly="True" ClientInstanceName="txtsuppliername"
                                CssPostfix="DevEx" Height="16px">
                            </dx:ASPxTextBox>
                        </td>
                        <td align="left" class="style12" height="25px" width="80%">
                            &nbsp;</td>
                    </tr>
                    <tr>
                        <td class="style47" align="left" height="26px" width="80%">
                            <dx:ASPxLabel ID="ASPxLabel13" runat="server" Font-Names="Tahoma" Font-Size="8pt"
                                Text="DELIVERY LOCATION" Width="137px">
                            </dx:ASPxLabel>
                        </td>
                        <td class="style48" align="left" height="26px" width="80%">

                            <dx:ASPxTextBox ID="cbolocation" runat="server" Width="170px" Font-Names="Tahoma"
                                Font-Size="8pt" BackColor="#CCCCCC" ReadOnly="True" Height="16px" 
                                ClientInstanceName="cbolocation">
                                <ClientSideEvents TextChanged="function(s, e) {
	                                lblerrmessage.SetText('');
                                }" />
                            </dx:ASPxTextBox>

                        </td>
                        <td align="left" class="style49" height="26px" width="80%">

                            <dx:ASPxTextBox ID="txtlocation" runat="server" Width="250px" Font-Names="Tahoma"
                                Font-Size="8pt" BackColor="#CCCCCC" ReadOnly="True" Height="16px" 
                                ClientInstanceName="txtlocation">
                                <ClientSideEvents TextChanged="function(s, e) {
	                                lblerrmessage.SetText('');
                                }" />
                            </dx:ASPxTextBox>

                            </td>
                        <td align="right" class="style49" height="26px" width="80%">
                            <dx:ASPxButton ID="btnsearch" runat="server" Text="SEARCH" Width="85px" Font-Names="Tahoma"
                                Font-Size="8pt" AutoPostBack="False" ClientInstanceName="btnsearch">
                                <ClientSideEvents Click="function(s, e) {
											var pDate = dtkanban.GetText();
                                            var pSupplier = cbosupplier.GetText();
                                                            
	                                        Grid.PerformCallback('gridload' + '|' + pDate + '|' + pSupplier);
	                                        lblerrmessage.SetText('');
}" />
                            </dx:ASPxButton>
                                    &nbsp;<dx:ASPxButton ID="btnclear" runat="server" Text="CLEAR" Width="85px" Font-Names="Tahoma"
                                Font-Size="8pt" AutoPostBack="False" ClientInstanceName="btnclear">
                            </dx:ASPxButton>
                        </td>
                    </tr>
                </table>
            </td>
            <td align="left" class="style26" width="100%">
                <dx:ASPxRoundPanel ID="ASPxRoundPanel2" runat="server" HeaderText="KANBAN STATUS"
                    Height="80px" ShowCollapseButton="true" View="GroupBox" Width="100%">
                    <PanelCollection>
                        <dx:PanelContent runat="server">
                            <table border="0" style="width: 100%;">
                                <tr>
                                    <td class="style35">
                                        <dx:ASPxLabel ID="ASPxLabel4" runat="server" Font-Names="Tahoma" 
                                            Font-Size="8pt" Text="(1) AFFILIATE ENTRY" Width="137px">
                                        </dx:ASPxLabel>
                                    </td>
                                    <td class="style36">
                                        <dx:ASPxTextBox ID="txtaffiliateentrydate" runat="server" BackColor="#CCCCCC" 
                                            ClientInstanceName="txtaffiliateentrydate" Font-Names="Tahoma" Font-Size="8pt" 
                                            Height="16px" ReadOnly="True" Width="150px">
                                            <ClientSideEvents TextChanged="function(s, e) {
	                                			Grid.SetFocusedRowIndex(-1);
                                                Grid.PerformCallback('kosong');
	                                            lblerrmessage.SetText('');
                                }" />
                                        </dx:ASPxTextBox>
                                    </td>
                                    <td class="style34">
                                        <dx:ASPxTextBox ID="txtaffiliateentryname" runat="server" BackColor="#CCCCCC" 
                                            ClientInstanceName="txtaffiliateentryname" Font-Names="Tahoma" Font-Size="8pt" 
                                            Height="16px" ReadOnly="True" Width="120px">
                                            <ClientSideEvents TextChanged="function(s, e) {
	                                            Grid.SetFocusedRowIndex(-1);
                                                Grid.PerformCallback('kosong');
	                                            lblerrmessage.SetText('');
                                }" />
                                        </dx:ASPxTextBox>
                                    </td>
                                </tr>
                                <tr>
                                    <td class="style35">
                                        <dx:ASPxLabel ID="ASPxLabel7" runat="server" Font-Names="Tahoma" 
                                            Font-Size="8pt" Text="(2) AFFILIATE APPROVAL" Width="145px">
                                        </dx:ASPxLabel>
                                    </td>
                                    <td class="style36">
                                        <dx:ASPxTextBox ID="txtaffiliateappdate" runat="server" BackColor="#CCCCCC" 
                                            ClientInstanceName="txtaffiliateappdate" Font-Names="Tahoma" Font-Size="8pt" 
                                            Height="16px" ReadOnly="True" Width="150px">
                                            <ClientSideEvents TextChanged="function(s, e) {
	                                			Grid.SetFocusedRowIndex(-1);
                                                Grid.PerformCallback('kosong');
	                                            lblerrmessage.SetText('');
                                }" />
                                        </dx:ASPxTextBox>
                                    </td>
                                    <td class="style34">
                                        <dx:ASPxTextBox ID="txtaffiliateappname" runat="server" BackColor="#CCCCCC" 
                                            ClientInstanceName="txtaffiliateappname" Font-Names="Tahoma" Font-Size="8pt" 
                                            Height="16px" ReadOnly="True" Width="120px">
                                            <ClientSideEvents TextChanged="function(s, e) {
	                                            Grid.SetFocusedRowIndex(-1);
                                                Grid.PerformCallback('kosong');
	                                            lblerrmessage.SetText('');
                                }" />
                                        </dx:ASPxTextBox>
                                    </td>
                                </tr>
                                <tr>
                                    <td class="style35">
                                        <dx:ASPxLabel ID="ASPxLabel5" runat="server" Font-Names="Tahoma" 
                                            Font-Size="8pt" Text="(3) SUPPLIER APPROVAL" Width="145px">
                                        </dx:ASPxLabel>
                                    </td>
                                    <td class="style36">
                                        <dx:ASPxTextBox ID="txtsupplierapprovaldate" runat="server" BackColor="#CCCCCC" 
                                            ClientInstanceName="txtsupplierapprovaldate" Font-Names="Tahoma" 
                                            Font-Size="8pt" Height="16px" ReadOnly="True" Width="150px">
                                            <ClientSideEvents TextChanged="function(s, e) {
	                                Grid.SetFocusedRowIndex(-1);
                                                Grid.PerformCallback('kosong');
	                                            lblerrmessage.SetText('');
                                }" />
                                        </dx:ASPxTextBox>
                                    </td>
                                    <td class="style34">
                                        <dx:ASPxTextBox ID="txtsupplierapprovalname" runat="server" BackColor="#CCCCCC" 
                                            ClientInstanceName="txtsupplierapprovalname" Font-Names="Tahoma" 
                                            Font-Size="8pt" Height="16px" ReadOnly="True" Width="120px">
                                            <ClientSideEvents TextChanged="function(s, e) {
	                                Grid.SetFocusedRowIndex(-1);
                                                Grid.PerformCallback('kosong');
	                                            lblerrmessage.SetText('');
                                }" />
                                        </dx:ASPxTextBox>
                                    </td>
                                </tr>
                            </table>
                        </dx:PanelContent>
                    </PanelCollection>
                </dx:ASPxRoundPanel>
            </td>
        </tr>
    </table>
    <table width="100%">
        <tr>
            <td align="left" bgcolor="White" class="style25" style="border-style: none;" width="100%">
                <table style="width: 100%;" border="0" cellpadding="0" cellspacing="0" width="100%">
                    <tr>
                        <td width="100%" height="0">
                            <table width="100%">
                                <tr>
                                    <td align="left" bgcolor="White" class="style25" style="border-width: thin; border-style: inset hidden ridge hidden;"
                                        width="100%" height="16px">
                                        <table style="width: 100%;" width="100%">
                                            <tr>
                                                <td width="100%">
                                                    <dx:ASPxLabel ID="lblerrmessage" runat="server" Font-Names="Tahoma" Font-Size="8pt"
                                                        Text="ERROR MESSAGE" Width="100%" ClientInstanceName="lblerrmessage">
                                                    </dx:ASPxLabel>
                                                </td>
                                            </tr>
                                        </table>
                                    </td>
                                </tr>
                            </table>
                            <table style="width: 100%;" border="0" cellpadding="0" cellspacing="0" width="100%">
                                <tr>
                                    <td height="16px">
                                        &nbsp;</td>
                                    <td height="16px">
                                        &nbsp;</td>
                                    <td height="16px">
                                        &nbsp;</td>
                                    <td height="16px">
                                        &nbsp;</td>
                                    <td height="16px">
                                        &nbsp;</td>
                                    <td height="16px">
                                        &nbsp;</td>
                                    <td height="16px">
                                        &nbsp;</td>
                                    <td height="16px" colspan="9">
                                        &nbsp;</td>
                                    <td class="style40" height="16px">
                                        <table style="width:100%;">
                                            <tr>
                                                <td align="right">
                                                    <dx:ASPxTextBox ID="txtsupplier8" runat="server" Width="20px" Font-Names="Tahoma"
                                                        Font-Size="8pt" BackColor="#FF66CC" Height="16px">
                                                        <ClientSideEvents TextChanged="function(s, e) {
	                                lblerrmessage.SetText('');
                                }" />
                                                    </dx:ASPxTextBox>
                                                </td>
                                                <td>
                                                    <dx:ASPxLabel ID="ASPxLabel6" runat="server" Font-Names="Tahoma" Font-Size="8pt"
                                                        Text=": SUPPLIER DELIVERY CAPACITY" Width="170px">
                                                    </dx:ASPxLabel>
                                                </td>
                                            </tr>
                                        </table>
                                    </td>
                                </tr>
                                <tr>
                                    <td height="16px">
                                        <dx:ASPxTextBox ID="txtkanban5" runat="server" Width="25px" Font-Names="Tahoma"
                                            Font-Size="5pt" BackColor="White" ReadOnly="True" Height="16px">
                                            <ClientSideEvents TextChanged="function(s, e) {
	                                lblerrmessage.SetText('');
                                }" />
                                            <Border BorderStyle="None" />
                                        </dx:ASPxTextBox>
                                    </td>
                                    <td height="16px">
                                        <dx:ASPxTextBox ID="txtkanban6" runat="server" Width="25px" Font-Names="Tahoma"
                                            Font-Size="5pt" BackColor="White" ReadOnly="True" Height="16px">
                                            <ClientSideEvents TextChanged="function(s, e) {
	                                lblerrmessage.SetText('');
                                }" />
                                            <Border BorderStyle="None" />
                                        </dx:ASPxTextBox>
                                    </td>
                                    <td height="16px">
                                        <dx:ASPxTextBox ID="txtkanban7" runat="server" Width="65px" Font-Names="Tahoma"
                                            Font-Size="5pt" BackColor="White" ReadOnly="True" Height="16px">
                                            <ClientSideEvents TextChanged="function(s, e) {
	                                lblerrmessage.SetText('');
                                }" />
                                            <Border BorderStyle="None" />
                                        </dx:ASPxTextBox>
                                    </td>
                                    <td height="16px">
                                        <dx:ASPxTextBox ID="txtkanban8" runat="server" Width="87px" Font-Names="Tahoma"
                                            Font-Size="5pt" BackColor="White" ReadOnly="True" Height="16px">
                                            <ClientSideEvents TextChanged="function(s, e) {
	                                lblerrmessage.SetText('');
                                }" />
                                            <Border BorderStyle="None" />
                                        </dx:ASPxTextBox>
                                    </td>
                                    <td height="16px">
                                        <dx:ASPxTextBox ID="txtkanban9" runat="server" Width="65px" Font-Names="Tahoma"
                                            Font-Size="5pt" BackColor="White" ReadOnly="True" Height="16px">
                                            <ClientSideEvents TextChanged="function(s, e) {
	                                lblerrmessage.SetText('');
                                }" />
                                            <Border BorderStyle="None" />
                                        </dx:ASPxTextBox>
                                    </td>
                                    <td height="16px">
                                        <dx:ASPxTextBox ID="txtkanban10" runat="server" Width="38px" Font-Names="Tahoma"
                                            Font-Size="5pt" BackColor="White" ReadOnly="True" Height="16px">
                                            <ClientSideEvents TextChanged="function(s, e) {
	                                lblerrmessage.SetText('');
                                }" />
                                            <Border BorderStyle="None" />
                                        </dx:ASPxTextBox>
                                    </td>
                                    <td height="16px">
                                        <dx:ASPxTextBox ID="txtkanban11" runat="server" Width="38px" Font-Names="Tahoma"
                                            Font-Size="5pt" BackColor="White" ReadOnly="True" Height="16px">
                                            <ClientSideEvents TextChanged="function(s, e) {
	                                lblerrmessage.SetText('');
                                }" />
                                            <Border BorderStyle="None" />
                                        </dx:ASPxTextBox>
                                    </td>
                                    <td height="16px">
                                        <dx:ASPxTextBox ID="txtkanban12" runat="server" Width="40px" Font-Names="Tahoma"
                                            Font-Size="5pt" BackColor="White" ReadOnly="True" Height="16px">
                                            <ClientSideEvents TextChanged="function(s, e) {
	                                lblerrmessage.SetText('');
                                }" />
                                            <Border BorderStyle="None" />
                                        </dx:ASPxTextBox>
                                    </td>
                                    <td height="16px">
                                        &nbsp;</td>
                                    <td height="16px">
                                        &nbsp;</td>
                                    <td height="16px">
                                        &nbsp;</td>
                                    <td height="16px" style="border-style: solid none none solid; border-width: 0.1px;
                                        border-color: #808080;" width="119px">
                                        <dx:ASPxTextBox ID="txtsupplierapprovalname0" runat="server" Width="119px" Font-Names="Tahoma"
                                            Font-Size="8pt" BackColor="#FFCC99" ReadOnly="True" Height="16px" 
                                            Text="KANBAN NO.">
                                            <ClientSideEvents TextChanged="function(s, e) {
	                                lblerrmessage.SetText('');
                                }" />
                                            <Border BorderStyle="None" />
                                            <BorderLeft BorderColor="#CCCCCC" BorderStyle="Groove" BorderWidth="1px" />
                                            <BorderTop BorderColor="#CCCCCC" BorderStyle="Groove" BorderWidth="1px" />
                                            <BorderRight BorderColor="#CCCCCC" BorderStyle="Groove" BorderWidth="1px" />
                                            <BorderBottom BorderColor="#CCCCCC" BorderStyle="Groove" BorderWidth="1px" />
                                        </dx:ASPxTextBox>
                                    </td>
                                    <td class="style43" height="16px" style="border-style: solid none none solid; border-width: 0.1px;
                                        border-color: #808080;">
                                        <dx:ASPxTextBox ID="txtkanban1" runat="server" Width="80px" Font-Names="Tahoma"
                                            Font-Size="8pt" BackColor="White" Height="16px" MaxLength="10"
                                            ClientInstanceName="txtkanban1" HorizontalAlign="Center" ReadOnly="True">
                                            <ClientSideEvents TextChanged="function(s, e) {
	                                lblerrmessage.SetText('');
                                }" />
                                           <Border BorderStyle="None" />
                                            <BorderLeft BorderColor="#CCCCCC" BorderStyle="Groove" BorderWidth="1px" />
                                            <BorderTop BorderColor="#CCCCCC" BorderStyle="Groove" BorderWidth="1px" />
                                            <BorderRight BorderColor="#CCCCCC" BorderStyle="Groove" BorderWidth="1px" />
                                            <BorderBottom BorderColor="#CCCCCC" BorderStyle="Groove" BorderWidth="1px" />
                                        </dx:ASPxTextBox>
                                    </td>
                                    <td class="style44" height="15px" style="border-style: solid none none solid; border-width: 0.1px;
                                        border-color: #808080;">
                                        <dx:ASPxTextBox ID="txtkanban2" runat="server" Width="80px" Font-Names="Tahoma"
                                            Font-Size="8pt" BackColor="White" Height="15px" MaxLength="10"
                                            ClientInstanceName="txtkanban2" HorizontalAlign="Center" ReadOnly="True">
                                            <ClientSideEvents TextChanged="function(s, e) {
	                                lblerrmessage.SetText('');
                                }" />
                                           <Border BorderStyle="None" />
                                            <BorderLeft BorderColor="#CCCCCC" BorderStyle="Groove" BorderWidth="1px" />
                                            <BorderTop BorderColor="#CCCCCC" BorderStyle="Groove" BorderWidth="1px" />
                                            <BorderRight BorderColor="#CCCCCC" BorderStyle="Groove" BorderWidth="1px" />
                                            <BorderBottom BorderColor="#CCCCCC" BorderStyle="Groove" BorderWidth="1px" />
                                        </dx:ASPxTextBox>
                                    </td>
                                    <td class="style45" height="15px" style="border-style: solid none none solid; border-width: 0.1px;
                                        border-color: #808080;">
                                        <dx:ASPxTextBox ID="txtkanban3" runat="server" Width="80px" Font-Names="Tahoma"
                                            Font-Size="8pt" BackColor="White" Height="15px" MaxLength="10"
                                            ClientInstanceName="txtkanban3" HorizontalAlign="Center" ReadOnly="True">
                                            <ClientSideEvents TextChanged="function(s, e) {
	                                lblerrmessage.SetText('');
                                }" />
                                            <Border BorderStyle="None" />
                                            <BorderLeft BorderColor="#CCCCCC" BorderStyle="Groove" BorderWidth="1px" />
                                            <BorderTop BorderColor="#CCCCCC" BorderStyle="Groove" BorderWidth="1px" />
                                            <BorderRight BorderColor="#CCCCCC" BorderStyle="Groove" BorderWidth="1px" />
                                            <BorderBottom BorderColor="#CCCCCC" BorderStyle="Groove" BorderWidth="1px" />
                                        </dx:ASPxTextBox>
                                    </td>
                                    <td class="style40" height="15px" style="border-style: solid solid none solid; border-width: 0.1px;
                                        border-color: #808080;">
                                        <dx:ASPxTextBox ID="txtkanban4" runat="server" Width="80px" Font-Names="Tahoma"
                                            Font-Size="8pt" BackColor="White" Height="15px" MaxLength="10"
                                            ClientInstanceName="txtkanban4" HorizontalAlign="Center" ReadOnly="True">
                                            <ClientSideEvents TextChanged="function(s, e) {
	                                lblerrmessage.SetText('');
                                }" />
                                            <Border BorderStyle="None" />
                                            <BorderLeft BorderColor="#CCCCCC" BorderStyle="Groove" BorderWidth="1px" />
                                            <BorderTop BorderColor="#CCCCCC" BorderStyle="Groove" BorderWidth="1px" />
                                            <BorderRight BorderColor="#CCCCCC" BorderStyle="Groove" BorderWidth="1px" />
                                            <BorderBottom BorderColor="#CCCCCC" BorderStyle="Groove" BorderWidth="1px" />
                                        </dx:ASPxTextBox>
                                    </td>
                                    <td class="style40" height="16px">
                                        &nbsp;
                                    </td>
                                </tr>
                                <tr>
                                    <td height="0">
                                        &nbsp;
                                    </td>
                                    <td height="0">
                                        &nbsp;
                                    </td>
                                    <td height="0">
                                        &nbsp;
                                    </td>
                                    <td height="0">
                                        &nbsp;
                                    </td>
                                    <td height="0">
                                        &nbsp;
                                    </td>
                                    <td height="0">
                                        &nbsp;
                                    </td>
                                    <td height="0">
                                        &nbsp;
                                    </td>
                                    <td height="0">
                                        &nbsp;
                                    </td>
                                    <td height="0">
                                        &nbsp;
                                    </td>
                                    <td height="0">
                                        &nbsp;&nbsp;
                                    </td>
                                    <td height="0">
                                        &nbsp;
                                    </td>
                                    <td height="0" style="border-style: solid none solid solid; border-width: 0.1px;
                                        border-color: #808080;">
                                        <dx:ASPxTextBox ID="txtsupplierapprovalname1" runat="server" Width="119px" Font-Names="Tahoma"
                                            Font-Size="8pt" BackColor="#FFCC99" ReadOnly="True" Height="16px" 
                                            Text="TIME">
                                            <ClientSideEvents TextChanged="function(s, e) {
	                                lblerrmessage.SetText('');
                                }" />
                                            <Border BorderStyle="None" />
                                            <BorderLeft BorderColor="#CCCCCC" BorderStyle="Groove" BorderWidth="1px" />
                                            <BorderTop BorderColor="#CCCCCC" BorderStyle="Groove" BorderWidth="1px" />
                                            <BorderRight BorderColor="#CCCCCC" BorderStyle="Groove" BorderWidth="1px" />
                                            <BorderBottom BorderColor="#CCCCCC" BorderStyle="Groove" BorderWidth="1px" />
                                        </dx:ASPxTextBox>
                                    </td>
                                    <td class="style43" height="0" style="border-style: solid none solid solid; border-width: 0.1px;
                                        border-color: #808080;">
                                        <dx:ASPxTextBox ID="txttime1" runat="server" Width="80px" Font-Names="Tahoma" Font-Size="8pt"
                                            BackColor="White" Height="16px" ClientInstanceName="txttime1" 
                                            HorizontalAlign="Center" MaxLength="5" ReadOnly="True">
                                            <ClientSideEvents TextChanged="function(s, e) {
	                                lblerrmessage.SetText('');
                                }" />
                                            <Border BorderStyle="None" />
                                            <BorderLeft BorderColor="#CCCCCC" BorderStyle="Groove" BorderWidth="1px" />
                                            <BorderTop BorderColor="#CCCCCC" BorderStyle="Groove" BorderWidth="1px" />
                                            <BorderRight BorderColor="#CCCCCC" BorderStyle="Groove" BorderWidth="1px" />
                                            <BorderBottom BorderColor="#CCCCCC" BorderStyle="Groove" BorderWidth="1px" />
                                        </dx:ASPxTextBox>
                                    </td>
                                    <td class="style44" height="0" style="border-style: solid none solid solid; border-width: 0.1px;
                                        border-color: #808080;">
                                        <dx:ASPxTextBox ID="txttime2" runat="server" Width="80px" Font-Names="Tahoma" Font-Size="8pt"
                                            BackColor="White" Height="15px" ClientInstanceName="txttime2" 
                                            HorizontalAlign="Center" MaxLength="5" ReadOnly="True">
                                            <ClientSideEvents TextChanged="function(s, e) {
	                                lblerrmessage.SetText('');
                                }" />
                                            <Border BorderStyle="None" />
                                            <BorderLeft BorderColor="#CCCCCC" BorderStyle="Groove" BorderWidth="1px" />
                                            <BorderTop BorderColor="#CCCCCC" BorderStyle="Groove" BorderWidth="1px" />
                                            <BorderRight BorderColor="#CCCCCC" BorderStyle="Groove" BorderWidth="1px" />
                                            <BorderBottom BorderColor="#CCCCCC" BorderStyle="Groove" BorderWidth="1px" />
                                        </dx:ASPxTextBox>
                                    </td>
                                    <td class="style45" height="0" style="border-style: solid none solid solid; border-width: 0.1px;
                                        border-color: #808080;">
                                        <dx:ASPxTextBox ID="txttime3" runat="server" Width="80px" Font-Names="Tahoma" Font-Size="8pt"
                                            BackColor="White" Height="15px" ClientInstanceName="txttime3" 
                                            HorizontalAlign="Center" MaxLength="5" ReadOnly="True">
                                            <ClientSideEvents TextChanged="function(s, e) {
	                                lblerrmessage.SetText('');
                                }" />
                                            <Border BorderStyle="None" />
                                            <BorderLeft BorderColor="#CCCCCC" BorderStyle="Groove" BorderWidth="1px" />
                                            <BorderTop BorderColor="#CCCCCC" BorderStyle="Groove" BorderWidth="1px" />
                                            <BorderRight BorderColor="#CCCCCC" BorderStyle="Groove" BorderWidth="1px" />
                                            <BorderBottom BorderColor="#CCCCCC" BorderStyle="Groove" BorderWidth="1px" />
                                        </dx:ASPxTextBox>
                                    </td>
                                    <td class="style40" height="0" style="border: 0.1px solid #808080">
                                        <dx:ASPxTextBox ID="txttime4" runat="server" Width="80px" Font-Names="Tahoma" Font-Size="8pt"
                                            BackColor="White" Height="15px" ClientInstanceName="txttime4" 
                                            HorizontalAlign="Center" MaxLength="5" ReadOnly="True">
                                            <ClientSideEvents TextChanged="function(s, e) {
	                                lblerrmessage.SetText('');
                                }" />
                                            <Border BorderStyle="None" />
                                            <BorderLeft BorderColor="#CCCCCC" BorderStyle="Groove" BorderWidth="1px" />
                                            <BorderTop BorderColor="#CCCCCC" BorderStyle="Groove" BorderWidth="1px" />
                                            <BorderRight BorderColor="#CCCCCC" BorderStyle="Groove" BorderWidth="1px" />
                                            <BorderBottom BorderColor="#CCCCCC" BorderStyle="Groove" BorderWidth="1px" />
                                        </dx:ASPxTextBox>
                                    </td>
                                    <td class="style40" height="0">
                                        &nbsp;
                                    </td>
                                </tr>
                            </table>
                        </td>
                    </tr>
                </table>
            </td>
        </tr>
    </table>
    <table style="width: 100%;" border="0" cellpadding="0" cellspacing="0" width="100%">
        <tr>
            <td align="left" height="0" valign="top">
                <dx:ASPxGridView ID="Grid" runat="server" AutoGenerateColumns="False" Width="100%"
                    Style="margin-top: 0px" KeyFieldName="colpartno;colpono;colno" ClientInstanceName="Grid"
                    Font-Names="Tahoma">
                    <ClientSideEvents EndCallback="function(s, e) {
                                      Grid.UpdateEdit();
                                      Grid.CancelEdit();
                                      
                                      txtaffiliateentrydate.SetText(s.cpAEDate);
                                      txtaffiliateentryname.SetText(s.cpAEName); 
                                      txtaffiliateappdate.SetText(s.cpAADate);
                                      txtaffiliateappname.SetText(s.cpAAName);
                                      txtsupplierapprovaldate.SetText(s.cpASDate);
                                      txtsupplierapprovalname.SetText(s.cpASName);
                                      txtkanban1.SetText(s.cpKanban1);
                                      txttime1.SetText(s.cpTime1);
                                      txtkanban2.SetText(s.cpKanban2);
                                      txttime2.SetText(s.cpTime2);
                                      txtkanban3.SetText(s.cpKanban3);
                                      txttime3.SetText(s.cpTime3);
                                      txtkanban4.SetText(s.cpKanban4);
                                      txttime4.SetText(s.cpTime4);
                                      txtaffiliatecode.SetText(s.cpaffcode);
                                      txtaffiliatename.SetText(s.cpaffname);
                                      cbolocation.SetText(s.locationcode);
                                      txtlocation.SetText(s.locationname);

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

                                      }" CallbackError="function(s, e) {e.handled = true;}" BatchEditEndEditing="OnBatchEditEndEditing"
                        BatchEditStartEditing="OnBatchEditStartEditing" FocusedRowChanged="function(s, e) {

                                        }"  />
                    <Columns>
                        <dx:GridViewDataCheckColumn FieldName="cols" Name="cols" VisibleIndex="0" Width="25px"
                            Caption=" ">
                            <HeaderStyle HorizontalAlign="Center" VerticalAlign="Middle" />
                            <PropertiesCheckEdit ValueChecked="1" ValueType="System.Int32" ValueUnchecked="0">
                            </PropertiesCheckEdit>
                        </dx:GridViewDataCheckColumn>
                        <dx:GridViewDataTextColumn Caption="NO" FieldName="colno" Name="colno" VisibleIndex="1"
                            Width="25px">
                            <HeaderStyle Font-Names="Tahoma" HorizontalAlign="Center" VerticalAlign="Middle"
                                Wrap="True" Font-Size="8pt" />
                            <CellStyle Font-Names="Tahoma" HorizontalAlign="Center" VerticalAlign="Middle" Font-Size="8pt">
                            </CellStyle>
                        </dx:GridViewDataTextColumn>
                        <dx:GridViewDataTextColumn Caption="PART NO" FieldName="colpartno" Name="colpartno"
                            VisibleIndex="2" Width="80px">
                            <HeaderStyle Font-Names="Tahoma" HorizontalAlign="Center" VerticalAlign="Middle"
                                Wrap="True" Font-Size="8pt" />
                            <CellStyle Font-Names="Tahoma" Font-Size="8pt">
                            </CellStyle>
                        </dx:GridViewDataTextColumn>
                        <dx:GridViewDataTextColumn Caption="DESCRIPTION" FieldName="coldescription" Name="coldescription"
                            VisibleIndex="3" Width="120px">
                            <HeaderStyle Font-Names="Tahoma" HorizontalAlign="Center" VerticalAlign="Middle"
                                Wrap="True" Font-Size="8pt" />
                            <CellStyle Font-Names="Tahoma" Font-Size="8pt">
                            </CellStyle>
                        </dx:GridViewDataTextColumn>
                        <dx:GridViewDataTextColumn Caption="PO NO" FieldName="colpono" Name="colpono" VisibleIndex="4"
                            Width="80px">
                            <HeaderStyle Font-Names="Tahoma" HorizontalAlign="Center" VerticalAlign="Middle"
                                Wrap="True" Font-Size="8pt" />
                            <CellStyle Font-Names="Tahoma" Font-Size="8pt">
                            </CellStyle>
                        </dx:GridViewDataTextColumn>
                        <dx:GridViewDataTextColumn Caption="UOM" FieldName="coluom" Name="coluom" VisibleIndex="5"
                            Width="40px">
                            <PropertiesTextEdit DisplayFormatString="{0:n0}">
                            </PropertiesTextEdit>
                            <HeaderStyle Font-Names="Tahoma" HorizontalAlign="Center" VerticalAlign="Middle"
                                Wrap="True" Font-Size="8pt" />
                            <CellStyle Font-Names="Tahoma" Font-Size="8pt">
                            </CellStyle>
                        </dx:GridViewDataTextColumn>
                        <dx:GridViewDataTextColumn Caption="MOQ" FieldName="colmoq" Name="colmoq" VisibleIndex="6"
                            Width="50px">
                            <PropertiesTextEdit DisplayFormatString="{0:n0}">
                                <MaskSettings IncludeLiterals="DecimalSymbol" />
                            </PropertiesTextEdit>
                            <HeaderStyle Font-Names="Tahoma" HorizontalAlign="Center" VerticalAlign="Middle"
                                Wrap="True" Font-Size="8pt" />
                            <CellStyle Font-Names="Tahoma" Font-Size="8pt">
                            </CellStyle>
                        </dx:GridViewDataTextColumn>
                        <dx:GridViewDataTextColumn Caption="QTY/ BOX" FieldName="colqty" Name="colqty" VisibleIndex="7"
                            Width="50px">
                            <PropertiesTextEdit DisplayFormatString="{0:n0}">
                                <MaskSettings IncludeLiterals="DecimalSymbol" />
                            </PropertiesTextEdit>
                            <HeaderStyle Font-Names="Tahoma" HorizontalAlign="Center" VerticalAlign="Middle"
                                Wrap="True" Font-Size="8pt" />
                            <CellStyle Font-Names="Tahoma" Font-Size="8pt">
                            </CellStyle>
                        </dx:GridViewDataTextColumn>
                        <dx:GridViewDataTextColumn Caption="PO QTY" FieldName="colpoqty" Name="colpoqty"
                            VisibleIndex="8" Width="50px">
                            <PropertiesTextEdit DisplayFormatString="{0:n0}">
                                <MaskSettings IncludeLiterals="DecimalSymbol" />
                            </PropertiesTextEdit>
                            <HeaderStyle Font-Names="Tahoma" HorizontalAlign="Center" VerticalAlign="Middle"
                                Wrap="True" Font-Size="8pt" />
                            <CellStyle Font-Names="Tahoma" Font-Size="8pt">
                            </CellStyle>
                        </dx:GridViewDataTextColumn>
                        <dx:GridViewDataTextColumn Caption="REMAINING SUPPLIER DAILY DELIVERY CAPACITY" FieldName="colremainingsupplier"
                            Name="colremainingsupplier" VisibleIndex="10" Width="75px">
                            <PropertiesTextEdit DisplayFormatString="{0:n0}">
                                <MaskSettings IncludeLiterals="DecimalSymbol" />
                            </PropertiesTextEdit>
                            <HeaderStyle Font-Names="Tahoma" HorizontalAlign="Center" VerticalAlign="Middle"
                                Wrap="True" Font-Size="8pt" />
                            <CellStyle Font-Names="Tahoma" Font-Size="8pt">
                            </CellStyle>
                        </dx:GridViewDataTextColumn>
                        <dx:GridViewDataTextColumn Caption="REMAINING PO QTY" FieldName="colremainingpo"
                            Name="colremainingpo" VisibleIndex="9" Width="65px">
                            <PropertiesTextEdit DisplayFormatString="{0:n0}">
                                <MaskSettings IncludeLiterals="DecimalSymbol" />
                            </PropertiesTextEdit>
                            <HeaderStyle Font-Names="Tahoma" HorizontalAlign="Center" VerticalAlign="Middle"
                                Wrap="True" Font-Size="8pt" />
                            <CellStyle Font-Names="Tahoma" Font-Size="8pt">
                            </CellStyle>
                        </dx:GridViewDataTextColumn>
                        <dx:GridViewDataTextColumn Caption="DELIVERY SCHEDULLE QTY" FieldName="coldeliveryqty"
                            Name="coldeliveryqty" VisibleIndex="11" Width="70px">
                            <PropertiesTextEdit DisplayFormatString="{0:n0}">
                                <MaskSettings IncludeLiterals="DecimalSymbol" />
                            </PropertiesTextEdit>
                            <HeaderStyle Font-Names="Tahoma" HorizontalAlign="Center" VerticalAlign="Middle"
                                Wrap="True" Font-Size="8pt" />
                            <CellStyle Font-Names="Tahoma" Font-Size="8pt">
                            </CellStyle>
                        </dx:GridViewDataTextColumn>
                        <dx:GridViewDataTextColumn Caption="KANBAN QTY" FieldName="colkanbanqty" Name="colkanbanqty"
                            VisibleIndex="12" Width="60px">
                            <PropertiesTextEdit DisplayFormatString="{0:n0}">
                                <MaskSettings ErrorText="Please input valid value !" Mask="<0..9999999999999g>.<00..99>"
                                    IncludeLiterals="DecimalSymbol" />
                            </PropertiesTextEdit>
                            <HeaderStyle Font-Names="Tahoma" HorizontalAlign="Center" VerticalAlign="Middle"
                                Wrap="True" Font-Size="8pt" />
                            <CellStyle Font-Names="Tahoma" Font-Size="8pt">
                            </CellStyle>
                        </dx:GridViewDataTextColumn>
                        <dx:GridViewDataTextColumn Caption="CYCLE 1" FieldName="colcycle1" Name="colcycle1"
                            VisibleIndex="13" Width="62px">
                            <PropertiesTextEdit DisplayFormatString="{0:n0}">
                                <MaskSettings ErrorText="Please input valid value !" Mask="<0..9999999999999g>.<00..99>"
                                    IncludeLiterals="DecimalSymbol" />
                            </PropertiesTextEdit>
                            <HeaderStyle Font-Names="Tahoma" HorizontalAlign="Center" VerticalAlign="Middle"
                                Wrap="True" Font-Size="8pt" />
                            <CellStyle Font-Names="Tahoma" Font-Size="8pt">
                            </CellStyle>
                        </dx:GridViewDataTextColumn>
                        <dx:GridViewDataTextColumn Caption="CYCLE 2" FieldName="colcycle2" Name="colcycle2"
                            VisibleIndex="14" Width="62px">
                            <PropertiesTextEdit DisplayFormatString="{0:n0}">
                                <MaskSettings ErrorText="Please input valid value !" Mask="<0..9999999999999g>.<00..99>"
                                    IncludeLiterals="DecimalSymbol" />
                            </PropertiesTextEdit>
                            <HeaderStyle Font-Names="Tahoma" HorizontalAlign="Center" VerticalAlign="Middle"
                                Wrap="True" Font-Size="8pt" />
                            <CellStyle Font-Names="Tahoma" Font-Size="8pt">
                            </CellStyle>
                        </dx:GridViewDataTextColumn>
                        <dx:GridViewDataTextColumn Caption="CYCLE 3" FieldName="colcycle3" Name="colcycle3"
                            VisibleIndex="15" Width="62px">
                            <PropertiesTextEdit DisplayFormatString="{0:n0}">
                                <MaskSettings ErrorText="Please input valid value !" Mask="<0..9999999999999g>.<00..99>"
                                    IncludeLiterals="DecimalSymbol" />
                            </PropertiesTextEdit>
                            <HeaderStyle Font-Names="Tahoma" HorizontalAlign="Center" VerticalAlign="Middle"
                                Wrap="True" Font-Size="8pt" />
                            <CellStyle Font-Names="Tahoma" Font-Size="8pt">
                            </CellStyle>
                        </dx:GridViewDataTextColumn>
                        <dx:GridViewDataTextColumn Caption="BOX QTY" FieldName="colbox" Name="colbox" VisibleIndex="17"
                            Width="41px">
                            <PropertiesTextEdit DisplayFormatString="{0:n0}">
                                <MaskSettings IncludeLiterals="DecimalSymbol" />
                            </PropertiesTextEdit>
                            <HeaderStyle Font-Names="Tahoma" HorizontalAlign="Center" VerticalAlign="Middle"
                                Wrap="True" Font-Size="8pt" />
                            <CellStyle Font-Names="Tahoma" Font-Size="8pt">
                            </CellStyle>
                        </dx:GridViewDataTextColumn>
                        <dx:GridViewDataTextColumn Caption="CYCLE 4" FieldName="colcycle4" Name="colcycle4"
                            VisibleIndex="16" Width="62px">
                            <PropertiesTextEdit DisplayFormatString="{0:n0}">
                                <MaskSettings ErrorText="Please input valid value !" Mask="<0..9999999999999g>.<00..99>"
                                    IncludeLiterals="DecimalSymbol" />
                            </PropertiesTextEdit>
                            <HeaderStyle Font-Names="Tahoma" HorizontalAlign="Center" VerticalAlign="Middle"
                                Wrap="True" Font-Size="8pt" />
                            <CellStyle Font-Names="Tahoma" Font-Size="8pt">
                            </CellStyle>
                        </dx:GridViewDataTextColumn>
                        <dx:GridViewDataTextColumn Caption="uom code" FieldName="coluomcode" 
                            Name="coluomcode" VisibleIndex="18" Width="0pt">
                        </dx:GridViewDataTextColumn>
                        <dx:GridViewDataTextColumn Caption="idx" FieldName="idx" Name="idx" 
                            VisibleIndex="19" Width="0px">
                        </dx:GridViewDataTextColumn>
                        <dx:GridViewDataTextColumn FieldName="kanbanno1" Name="kanbanno1" 
                            VisibleIndex="20" Width="0px">
                        </dx:GridViewDataTextColumn>
                        <dx:GridViewDataTextColumn Caption="kanbantime" FieldName="kanbantime1" 
                            Name="kanbantime1" VisibleIndex="24" Width="0px">
                        </dx:GridViewDataTextColumn>
                        <dx:GridViewDataTextColumn FieldName="kanbanno2" Name="kanbanno2" 
                            VisibleIndex="21" Width="0px">
                        </dx:GridViewDataTextColumn>
                        <dx:GridViewDataTextColumn FieldName="kanbanno3" Name="kanbanno3" 
                            VisibleIndex="22" Width="0px">
                        </dx:GridViewDataTextColumn>
                        <dx:GridViewDataTextColumn FieldName="kanbanno4" Name="kanbanno4" 
                            VisibleIndex="23" Width="0px">
                        </dx:GridViewDataTextColumn>
                        <dx:GridViewDataTextColumn FieldName="kanbantime2" Name="kanbantime2" 
                            VisibleIndex="25" Width="0px">
                        </dx:GridViewDataTextColumn>
                        <dx:GridViewDataTextColumn FieldName="kanbantime3" Name="kanbantime3" 
                            VisibleIndex="26" Width="0px">
                        </dx:GridViewDataTextColumn>
                        <dx:GridViewDataTextColumn FieldName="kanbantime4" Name="kanbantime4" 
                            VisibleIndex="27" Width="0px">
                        </dx:GridViewDataTextColumn>
                    </Columns>
                    <SettingsPager Visible="False">
                    </SettingsPager>
                    <SettingsEditing Mode="Batch" NewItemRowPosition="Bottom">
                        <BatchEditSettings ShowConfirmOnLosingChanges="False" />
                        <BatchEditSettings ShowConfirmOnLosingChanges="False"></BatchEditSettings>
                    </SettingsEditing>
                    <Settings ShowHorizontalScrollBar="True" ShowVerticalScrollBar="True" ShowGroupButtons="False"
                        ShowStatusBar="Hidden" VerticalScrollableHeight="135" />
                    <Settings ShowGroupButtons="False" ShowVerticalScrollBar="True" ShowHorizontalScrollBar="True"
                        VerticalScrollableHeight="135" ShowStatusBar="Hidden"></Settings>
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
        </tr>
        <tr>
            <td>
                <table style="width: 100%;" width="100%">
                    <tr>
                        <td align="left" height="0">
                            <dx:ASPxButton ID="btnsubmenu" runat="server" Width="90px" Font-Names="Tahoma" Font-Size="8pt"
                                Text="BACK">
                            </dx:ASPxButton>
                        </td>
                        <td align="right" height="0">
                            <dx:ASPxButton ID="btnprintcard" runat="server" Text="PRINT KANBAN CARD" Width="90px"
                                Font-Names="Tahoma" Font-Size="8pt" ClientInstanceName="btnprintcard">
                            </dx:ASPxButton>
                            &nbsp;<dx:ASPxButton ID="btnprintcycle" runat="server" Text="PRINT KANBAN CYCLE"
                                Width="90px" Font-Names="Tahoma" Font-Size="8pt" ClientInstanceName="btnprintcycle">
                            </dx:ASPxButton>
                            &nbsp;</td>
                    </tr>
                </table>
            </td>
        </tr>
    </table>
    <asp:SqlDataSource ID="SupplierCode" runat="server" ConnectionString="<%$ ConnectionStrings:KonString %>"
        ProviderName="<%$ ConnectionStrings:KonString.ProviderName %>" SelectCommand="SELECT supplierid AS Supplier_Code , suppliername AS Supplier_Name FROM MS_Supplier ">
    </asp:SqlDataSource>
    <br />
</asp:Content>
