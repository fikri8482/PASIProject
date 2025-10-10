<%@ Page Title="" Language="vb" AutoEventWireup="false" MasterPageFile="~/SiteContent.Master"
    CodeBehind="KanbanCreate.aspx.vb" Inherits="AffiliateSystem.KanbanCreate" %>

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
<%@ Register assembly="DevExpress.Web.v14.1, Version=14.1.6.0, Culture=neutral, PublicKeyToken=b88d1754d700e49a" namespace="DevExpress.Web.ASPxCallback" tagprefix="dx" %>
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

        function OnAllCheckedChanged(s, e) {
            if (s.GetValue() == -1) s.SetValue(1);
            for (var i = 0; i < Grid.GetVisibleRowsOnPage(); i++) {
                Grid.batchEditApi.SetCellValue(i, "cols", s.GetValue());
            }
        }

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
                
                e.cancel = false;
            }


            currentEditableVisibleIndex = e.visibleIndex;
        }

        function OnBatchEditEndEditing(s, e) {
            window.setTimeout(function () {
                var kanbanqty = s.batchEditApi.GetCellValue(e.visibleIndex, "colkanbanqty");
                var cycle1 = s.batchEditApi.GetCellValue(e.visibleIndex, "colcycle1");
                var cycle2 = s.batchEditApi.GetCellValue(e.visibleIndex, "colcycle2");
                var cycle3 = s.batchEditApi.GetCellValue(e.visibleIndex, "colcycle3");
                var cycle4 = s.batchEditApi.GetCellValue(e.visibleIndex, "colcycle4");
                var qtybox = s.batchEditApi.GetCellValue(e.visibleIndex, "colqty");
                //                                alert(currentColumnName)
                if (currentColumnName == "cols" || currentColumnName == "colbox" || currentColumnName == "colcycle1" || currentColumnName == "colcycle2" || currentColumnName == "colcycle3" || currentColumnName == "colcycle4") {
                    var ntotal = parseInt(cycle1) + parseInt(cycle2) + parseInt(cycle3) + parseInt(cycle4);
                    s.batchEditApi.SetCellValue(e.visibleIndex, "colkanbanqty", ntotal);
                } else {
                    if (currentcolumnname = "colkanbanqty") {
                        var n1 = Math.ceil(Math.floor(kanbanqty / 4) / qtybox) * qtybox;

                        if (kanbanqty - (Math.ceil(Math.floor(kanbanqty / 4) / qtybox) * qtybox) > 0) {
                            var n2 = Math.ceil(Math.floor(kanbanqty / 4) / qtybox) * qtybox;
                        } else {
                            var n2 = 0;
                        }

                        if ((n1 + n2) > kanbanqty) {
                            var ncycle2 = ((n1 + n2) - kanbanqty);
                        } else {
                            var ncycle2 = n2;
                        }

                        if (kanbanqty - ((Math.ceil(Math.floor(kanbanqty / 4) / qtybox) * qtybox) * 2) > 0) {
                            var n3 = Math.ceil(Math.floor(kanbanqty / 4) / qtybox) * qtybox;
                        } else {
                            var n3 = 0;
                        }

                        if ((n1 + n2 + n3) > kanbanqty) {
                            var ncycle3 = ((n1 + n2 + n3) - kanbanqty);
                        } else {
                            var ncycle3 = n3;
                        }

                        if (kanbanqty - ((Math.ceil(Math.floor(kanbanqty / 4) / qtybox) * qtybox) * 3) > 0) {
                            var n4 = kanbanqty - (n1 * 3);
                        } else {
                            var n4 = 0;
                        }

                        s.batchEditApi.SetCellValue(e.visibleIndex, "colcycle1", n1);
                        s.batchEditApi.SetCellValue(e.visibleIndex, "colcycle2", ncycle2);
                        s.batchEditApi.SetCellValue(e.visibleIndex, "colcycle3", ncycle3);
                        s.batchEditApi.SetCellValue(e.visibleIndex, "colcycle4", n4);
                        //                        alert('selesai');
                    }
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
                            <dx:ASPxLabel ID="ASPxLabel1" runat="server" Font-Names="Tahoma" Font-Size="8pt"
                                Text="PERIOD">
                            </dx:ASPxLabel>
                        </td>
                        <td class="style6" rowspan="1" height="25px" width="80%">
                            <dx:ASPxTimeEdit ID="dt1" runat="server" ClientInstanceName="dt1"
                                DisplayFormatString="MMM yyyy" EditFormat="Custom" 
                                EditFormatString="MMM yyyy" Font-Names="Tahoma" Font-Size="8pt">
                                <ClientSideEvents ButtonClick="function(s, e) {
	Grid.PerformCallback('change' + '|' + '2014-04-04' + '|' + '' + '|' + '')
}" />
<ClientSideEvents ButtonClick="function(s, e) {
	Grid.PerformCallback('change' + '|' + '2014-04-04' + '|' + '' + '|' + '')
	lblerrmessage.SetText('');
}"></ClientSideEvents>
                            </dx:ASPxTimeEdit>
                        </td>
                        <td class="style12" align="left" height="25px" width="80%">
                            <table style="width:100%;">
                                <tr>
                                    <td>
                            <dx:ASPxLabel ID="ASPxLabel2" runat="server" Font-Names="Tahoma" Font-Size="8pt"
                                Text="KANBAN DATE" Width="90px">
                            </dx:ASPxLabel>
                                    </td>
                                    <td>
                            <dx:ASPxDateEdit ID="dtkanban" runat="server" Font-Names="Tahoma" Font-Size="8pt"
                                EditFormat="Custom" EditFormatString="dd MMM yyyy" ClientInstanceName="dtkanban">
                                <clientsideevents ValueChanged="function(s, e) {
	Grid.PerformCallback('change' + '|' + '2014-04-04' + '|' + '' + '|' + '')
	lblerrmessage.SetText('');
}" />
                            </dx:ASPxDateEdit>
                                    </td>
                                </tr>
                            </table>
                        </td>
                        <td class="style12" align="left" height="25px" width="80%">
                            <dx:ASPxComboBox ID="cboseqno" runat="server" ClientInstanceName="cboseqno"
                                Font-Names="Tahoma" TextFormatString="{0}" Font-Size="8pt" Width="70px" 
                                Visible="False">
                            </dx:ASPxComboBox>
                        </td>
                    </tr>
                    <tr>
                        <td class="style11" align="left" height="25px" width="80%">
                            <dx:ASPxLabel ID="ASPxLabel3" runat="server" Font-Names="Tahoma" Font-Size="8pt"
                                Text="SUPPLIER CODE/NAME" Width="137px">
                            </dx:ASPxLabel>
                        </td>
                        <td class="style6" rowspan="1" align="left" height="25px" width="80%">
                            <dx:ASPxComboBox ID="cbosupplier" runat="server" ClientInstanceName="cbosupplier"
                                Font-Names="Tahoma" TextFormatString="{0}" Font-Size="8pt">
                                <ClientSideEvents SelectedIndexChanged="function(s, e) {
	                txtsuppliername.SetText(cbosupplier.GetSelectedItem().GetColumnText(1));
                     txtaddress.SetText(cbosupplier.GetSelectedItem().GetColumnText(2));
                    }" />
<ClientSideEvents SelectedIndexChanged="function(s, e) {
	                 txtsuppliername.SetText(cbosupplier.GetSelectedItem().GetColumnText(1));
					 Grid.PerformCallback('change' + '|' + '2014-04-04' + '|' + '' + '|' + '')
					 lblerrmessage.SetText('');
                    }"></ClientSideEvents>
                            </dx:ASPxComboBox>
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
                            <dx:ASPxComboBox ID="cbolocation" runat="server" ClientInstanceName="cbolocation"
                                Font-Names="Tahoma" TextFormatString="{0}" Font-Size="8pt">
                                <ClientSideEvents SelectedIndexChanged="function(s, e) {
	                                            txtlocation.SetText(cbolocation.GetSelectedItem().GetColumnText(1));
                                                }" />
<ClientSideEvents SelectedIndexChanged="function(s, e) {
	                                            txtlocation.SetText(cbolocation.GetSelectedItem().GetColumnText(1));
												Grid.PerformCallback('change' + '|' + '2014-04-04' + '|' + '' + '|' + '')
												lblerrmessage.SetText('');
                                                }"></ClientSideEvents>
                            </dx:ASPxComboBox>
                        </td>
                        <td align="left" class="style49" height="26px" width="80%">

                            <dx:ASPxTextBox ID="txtlocation" runat="server" Width="170px" Font-Names="Tahoma"
                                Font-Size="8pt" BackColor="#CCCCCC" ReadOnly="True" Height="16px" 
                                ClientInstanceName="txtlocation">
                                <ClientSideEvents TextChanged="function(s, e) {
	                                lblerrmessage.SetText('');
                                }" />
                            </dx:ASPxTextBox>

                            </td>
                        <td align="left" class="style49" height="26px" width="80%">
                            <table style="width:100%;">
                                <tr>
                                    <td>
                            <dx:ASPxButton ID="btnsearch" runat="server" Text="SEARCH" Width="85px" Font-Names="Tahoma"
                                Font-Size="8pt" AutoPostBack="False" ClientInstanceName="btnsearch">
                                <ClientSideEvents Click="function(s, e) {
											var pDate = dtkanban.GetText();
                                            var pSupplier = cbosupplier.GetText();
											var pSeq = cboseq.GetText();
                                                            
	                                        Grid.PerformCallback('gridload' + '|' + pDate + '|' + pSupplier + '|' + pSeq);
	                                        lblerrmessage.SetText('');
}" />
                            </dx:ASPxButton>
                                    </td>
                                    <td>
                            <dx:ASPxButton ID="btnclear" runat="server" Text="CLEAR" Width="85px" Font-Names="Tahoma"
                                Font-Size="8pt" AutoPostBack="False" ClientInstanceName="btnclear">
                            </dx:ASPxButton>
                                    </td>
                                </tr>
                            </table>
                        </td>
                    </tr>
                    <tr>
                        <td class="style47" align="left" height="26px" width="80%">
                            <dx:ASPxLabel ID="ASPxLabel14" runat="server" Font-Names="Tahoma" Font-Size="8pt"
                                Text="CYCLE" Width="137px">
                            </dx:ASPxLabel>
                        </td>
                        <td class="style48" align="left" height="26px" width="80%">
                            <dx:ASPxComboBox ID="cboseq" runat="server" ClientInstanceName="cboseq"
                                Font-Names="Tahoma" TextFormatString="{0}" Font-Size="8pt">
                                <ClientSideEvents SelectedIndexChanged="function(s, e) {
		Grid.PerformCallback('change' + '|' + '2014-04-04' + '|' + '' + '|' + '')
		lblerrmessage.SetText('');
                                                }" />
                            </dx:ASPxComboBox>
                        </td>
                        <td align="left" class="style49" height="26px" width="80%">

                            <table style="width:100%;">
                                <tr>
                                    <td>
                            <dx:ASPxLabel ID="ASPxLabel15" runat="server" Font-Names="Tahoma" Font-Size="8pt"
                                Text="TYPE" Width="90px">
                            </dx:ASPxLabel>
                                    </td>
                                    <td>
                            <dx:ASPxComboBox ID="cbotype" runat="server" ClientInstanceName="cbotype"
                                Font-Names="Tahoma" TextFormatString="{0}" Font-Size="8pt">
                                <ClientSideEvents SelectedIndexChanged="function(s, e) {
		Grid.PerformCallback('change' + '|' + '2014-04-04' + '|' + '' + '|' + '')
		lblerrmessage.SetText('');
                                                }" />
                            </dx:ASPxComboBox>
                                    </td>
                                </tr>
                            </table>
                        </td>
                        <td align="left" class="style49" height="26px" width="80%">
                            &nbsp;</td>
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
                                            Font-Size="8pt" Text="(1) AFFILIATE CONFIRMATION" Width="137px">
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
                                            Font-Size="8pt" Text="(2) SEND TO SUPPLIER" Width="145px">
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
                            <table id="tblMsg" style="border-width: thin; border-style: inset hidden ridge hidden; border-color:#9598A1; width:100%; height: 15px;">
                                <tr>
                                    <td align="left" valign="middle" height="15px">
                                        <dx:ASPxLabel ID="lblerrmessage" runat="server" Text="[lblinfo]" Font-Names="Tahoma" 
                                            ClientInstanceName="lblerrmessage" Font-Bold="True" Font-Italic="True" 
                                            Font-Size="8pt">
                                        </dx:ASPxLabel>
                                    </td>
                                </tr>         
                            </table>
                            <table style="width: 100%;" border="0" cellpadding="0" cellspacing="0" width="100%">
                                <tr>
                                    <td height="16px" colspan="13">
                                        &nbsp;</td>
                                    <td class="style40" height="16px" align="right">
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
                                                            <td align="left" width="170">
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
                                    <td height="16px" style="border-style: solid none none solid; border-width: 0.1px;
                                        border-color: #808080;" width="119px">
                                        <dx:ASPxTextBox ID="txtsupplierapprovalname0" runat="server" Width="119px" Font-Names="Tahoma"
                                            Font-Size="8pt" BackColor="#FFCC99" ReadOnly="True" Height="16px" 
                                            Text="KANBAN NO.">
                                            <ClientSideEvents TextChanged="function(s, e) {
	                                lblerrmessage.SetText('');
                                }" />
                                            <Border BorderStyle="Groove" BorderColor="Silver" BorderWidth="1px" />
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
                                            ClientInstanceName="txtkanban1" HorizontalAlign="Center">
                                            <ClientSideEvents TextChanged="function(s, e) {
	                                lblerrmessage.SetText('');
                                }" />
                                            <Border BorderStyle="Groove" BorderColor="Silver" BorderWidth="1px" />
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
                                            ClientInstanceName="txtkanban2" HorizontalAlign="Center">
                                            <ClientSideEvents TextChanged="function(s, e) {
	                                lblerrmessage.SetText('');
                                }" />
                                            <Border BorderStyle="Groove" BorderColor="Silver" BorderWidth="1px" />
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
                                            ClientInstanceName="txtkanban3" HorizontalAlign="Center">
                                            <ClientSideEvents TextChanged="function(s, e) {
	                                lblerrmessage.SetText('');
                                }" />
                                            <Border BorderStyle="Groove" BorderColor="Silver" BorderWidth="1px" />
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
                                            ClientInstanceName="txtkanban4" HorizontalAlign="Center">
                                            <ClientSideEvents TextChanged="function(s, e) {
	                                lblerrmessage.SetText('');
                                }" />
                                            <Border BorderStyle="Groove" BorderColor="Silver" BorderWidth="1px" />
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
                                    <td height="0" align="right">
                                                    &nbsp;</td>
                                    <td height="0" colspan="4">
                                                    &nbsp;</td>
                                    <td height="0" style="border-style: solid none solid solid; border-width: 0.1px;
                                        border-color: #808080;">
                                        <dx:ASPxTextBox ID="txtsupplierapprovalname1" runat="server" Width="119px" Font-Names="Tahoma"
                                            Font-Size="8pt" BackColor="#FFCC99" ReadOnly="True" Height="16px" 
                                            Text="TIME">
                                            <ClientSideEvents TextChanged="function(s, e) {
	                                lblerrmessage.SetText('');
                                }" />
                                            <Border BorderStyle="Groove" BorderColor="Silver" BorderWidth="1px" />
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
                                            HorizontalAlign="Center" MaxLength="5">
                                            <ClientSideEvents TextChanged="function(s, e) {
	                                lblerrmessage.SetText('');
                                }" />
                                            <Border BorderStyle="Groove" BorderColor="Silver" BorderWidth="1px" />
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
                                            HorizontalAlign="Center" MaxLength="5">
                                            <ClientSideEvents TextChanged="function(s, e) {
	                                lblerrmessage.SetText('');
                                }" />
                                            <Border BorderStyle="Groove" BorderColor="Silver" BorderWidth="1px" />
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
                                            HorizontalAlign="Center" MaxLength="5">
                                            <ClientSideEvents TextChanged="function(s, e) {
	                                lblerrmessage.SetText('');
                                }" />
                                            <Border BorderStyle="Groove" BorderColor="Silver" BorderWidth="1px" />
                                            <BorderLeft BorderColor="#CCCCCC" BorderStyle="Groove" BorderWidth="1px" />
                                            <BorderTop BorderColor="#CCCCCC" BorderStyle="Groove" BorderWidth="1px" />
                                            <BorderRight BorderColor="#CCCCCC" BorderStyle="Groove" BorderWidth="1px" />
                                            <BorderBottom BorderColor="#CCCCCC" BorderStyle="Groove" BorderWidth="1px" />
                                        </dx:ASPxTextBox>
                                    </td>
                                    <td class="style40" height="0" style="border: 0.1px solid #808080">
                                        <dx:ASPxTextBox ID="txttime4" runat="server" Width="80px" Font-Names="Tahoma" Font-Size="8pt"
                                            BackColor="White" Height="15px" ClientInstanceName="txttime4" 
                                            HorizontalAlign="Center" MaxLength="5">
                                            <ClientSideEvents TextChanged="function(s, e) {
	                                lblerrmessage.SetText('');
                                }" />
                                            <Border BorderStyle="Groove" BorderColor="Silver" BorderWidth="1px" />
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
                                      Grid.CancelEdit();
                                      txtaffiliateentrydate.SetText(s.cpAEDate);
                                      txtaffiliateentryname.SetText(s.cpAEName); 
                                      txtaffiliateappdate.SetText(s.cpAADate);
                                      txtaffiliateappname.SetText(s.cpAAName);
                                      txtkanban1.SetText(s.cpKanban1);
                                      txttime1.SetText(s.cpTime1);
                                      txtkanban2.SetText(s.cpKanban2);
                                      txttime2.SetText(s.cpTime2);
                                      txtkanban3.SetText(s.cpKanban3);
                                      txttime3.SetText(s.cpTime3);
                                      txtkanban4.SetText(s.cpKanban4);
                                      txttime4.SetText(s.cpTime4);
									  btnApprove.SetText(s.cpButton);

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
                    <Settings ShowHorizontalScrollBar="True" ShowVerticalScrollBar="True" ShowGroupButtons="False"
                        ShowStatusBar="Hidden" VerticalScrollableHeight="135" />
<ClientSideEvents FocusedRowChanged="function(s, e) {

                                        }" BatchEditStartEditing="OnBatchEditStartEditing" 
                        BatchEditEndEditing="OnBatchEditEndEditing" EndCallback="function(s, e) {
                                      Grid.CancelEdit();
                                      txtaffiliateentrydate.SetText(s.cpAEDate);
                                      txtaffiliateentryname.SetText(s.cpAEName); 
                                      txtaffiliateappdate.SetText(s.cpAADate);
                                      txtaffiliateappname.SetText(s.cpAAName);
                                      txtkanban1.SetText(s.cpKanban1);
                                      txttime1.SetText(s.cpTime1);
                                      txtkanban2.SetText(s.cpKanban2);
                                      txttime2.SetText(s.cpTime2);
                                      txtkanban3.SetText(s.cpKanban3);
                                      txttime3.SetText(s.cpTime3);
                                      txtkanban4.SetText(s.cpKanban4);
                                      txttime4.SetText(s.cpTime4);

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

                                      }" CallbackError="function(s, e) {e.handled = true;}"></ClientSideEvents>
                    <Columns>
                        <dx:GridViewDataCheckColumn FieldName="cols" Name="cols" VisibleIndex="0" Width="25px"
                            Caption=" ">
                            <HeaderStyle HorizontalAlign="Center" VerticalAlign="Middle" />
                            <HeaderCaptionTemplate>
                                <dx:ASPxCheckBox ID="chkAll" runat="server" ClientInstanceName="chkAll" ClientSideEvents-CheckedChanged="OnAllCheckedChanged"
                                ValueType="System.Int32" ValueChecked="1" ValueUnchecked="0">
                                </dx:ASPxCheckBox>
                            </HeaderCaptionTemplate>
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
                        <dx:GridViewDataTextColumn Caption="PO MONTHLY TOTAL QTY" FieldName="colpoqty" Name="colpoqty"
                            VisibleIndex="8" Width="50px">
                            <PropertiesTextEdit DisplayFormatString="{0:n0}">
                                <MaskSettings IncludeLiterals="DecimalSymbol" />
                            </PropertiesTextEdit>
                            <HeaderStyle Font-Names="Tahoma" HorizontalAlign="Center" VerticalAlign="Middle"
                                Wrap="True" Font-Size="8pt" />
                            <CellStyle Font-Names="Tahoma" Font-Size="8pt">
                            </CellStyle>
                        </dx:GridViewDataTextColumn>
                        <dx:GridViewDataTextColumn Caption="REMAINING SUPLLIER DELIVERY CAPACITY OVER LIMIT" FieldName="colremainingsupplier"
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
                                <MaskSettings ErrorText="Please input valid value !" Mask="<0..9999999999999g>"
                                    IncludeLiterals="DecimalSymbol" />
                                <ValidationSettings ErrorDisplayMode="None"></ValidationSettings>
                            </PropertiesTextEdit>
                            <HeaderStyle Font-Names="Tahoma" HorizontalAlign="Center" VerticalAlign="Middle"
                                Wrap="True" Font-Size="8pt" />
                            <CellStyle Font-Names="Tahoma" Font-Size="8pt">
                            </CellStyle>
                        </dx:GridViewDataTextColumn>
                        <dx:GridViewDataTextColumn Caption="CYCLE 1" FieldName="colcycle1" Name="colcycle1"
                            VisibleIndex="13" Width="62px">
                            <PropertiesTextEdit DisplayFormatString="{0:n0}">
                                <MaskSettings ErrorText="Please input valid value !" Mask="<0..9999999999999g>"
                                    IncludeLiterals="DecimalSymbol" />
                                <ValidationSettings ErrorDisplayMode="None"></ValidationSettings>
                            </PropertiesTextEdit>
                            <HeaderStyle Font-Names="Tahoma" HorizontalAlign="Center" VerticalAlign="Middle"
                                Wrap="True" Font-Size="8pt" />
                            <CellStyle Font-Names="Tahoma" Font-Size="8pt">
                            </CellStyle>
                        </dx:GridViewDataTextColumn>
                        <dx:GridViewDataTextColumn Caption="CYCLE 2" FieldName="colcycle2" Name="colcycle2"
                            VisibleIndex="14" Width="62px">
                            <PropertiesTextEdit DisplayFormatString="{0:n0}">
                                <MaskSettings ErrorText="Please input valid value !" Mask="<0..9999999999999g>"
                                    IncludeLiterals="DecimalSymbol" />
                                <ValidationSettings ErrorDisplayMode="None"></ValidationSettings>
                            </PropertiesTextEdit>
                            <HeaderStyle Font-Names="Tahoma" HorizontalAlign="Center" VerticalAlign="Middle"
                                Wrap="True" Font-Size="8pt" />
                            <CellStyle Font-Names="Tahoma" Font-Size="8pt">
                            </CellStyle>
                        </dx:GridViewDataTextColumn>
                        <dx:GridViewDataTextColumn Caption="CYCLE 3" FieldName="colcycle3" Name="colcycle3"
                            VisibleIndex="15" Width="62px">
                            <PropertiesTextEdit DisplayFormatString="{0:n0}">
                                <MaskSettings ErrorText="Please input valid value !" Mask="<0..9999999999999g>"
                                    IncludeLiterals="DecimalSymbol" />
                                <ValidationSettings ErrorDisplayMode="None"></ValidationSettings>
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
                                <ValidationSettings ErrorDisplayMode="None"></ValidationSettings>
                            </PropertiesTextEdit>
                            <HeaderStyle Font-Names="Tahoma" HorizontalAlign="Center" VerticalAlign="Middle"
                                Wrap="True" Font-Size="8pt" />
                            <CellStyle Font-Names="Tahoma" Font-Size="8pt">
                            </CellStyle>
                        </dx:GridViewDataTextColumn>
                        <dx:GridViewDataTextColumn Caption="CYCLE 4" FieldName="colcycle4" Name="colcycle4"
                            VisibleIndex="16" Width="62px">
                            <PropertiesTextEdit DisplayFormatString="{0:n0}">
                                <MaskSettings ErrorText="Please input valid value !" Mask="<0..9999999999999g>"
                                    IncludeLiterals="DecimalSymbol" />
                                <ValidationSettings ErrorDisplayMode="None"></ValidationSettings>
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
                    <SettingsPager Mode="ShowAllRecords" Visible="False">
                    </SettingsPager>
                    <SettingsEditing Mode="Batch" NewItemRowPosition="Bottom">
                        <BatchEditSettings ShowConfirmOnLosingChanges="False" />
                        <BatchEditSettings ShowConfirmOnLosingChanges="False"></BatchEditSettings>
                    </SettingsEditing>
                    <Settings ShowGroupButtons="False" ShowVerticalScrollBar="True" ShowHorizontalScrollBar="True"
                        VerticalScrollableHeight="180" ShowStatusBar="Hidden"></Settings>
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
                                Text="SUB MENU">
                            </dx:ASPxButton>
                        </td>
                        <td align="right" height="0">
                            <dx:ASPxButton ID="btnprintcard" runat="server" Text="PRINT KANBAN CARD" Width="90px"
                                Font-Names="Tahoma" Font-Size="8pt" ClientInstanceName="btnprintcard">
                            </dx:ASPxButton>
                            &nbsp;<dx:ASPxButton ID="btnprintcycle" runat="server" Text="PRINT KANBAN CYCLE"
                                Width="90px" Font-Names="Tahoma" Font-Size="8pt" ClientInstanceName="btnprintcycle">
                            </dx:ASPxButton>
                            &nbsp;<dx:ASPxButton ID="btnApprove" runat="server" Text="SEND TO SUPPLIER" Width="85px" Font-Names="Tahoma"
                                Font-Size="8pt" AutoPostBack="False" ClientInstanceName="btnApprove">
                                <ClientSideEvents Click="function(s, e) { 
                                    Approve.PerformCallback();
                                                }" />
                            </dx:ASPxButton>
                            <%--&nbsp;<dx:ASPxButton ID="btndelete" runat="server" Text="DELETE" Width="85px" Font-Names="Tahoma"
                                Font-Size="8pt" AutoPostBack="False" ClientInstanceName="btndelete">
                            </dx:ASPxButton>--%>
                            <%--&nbsp;<dx:ASPxButton ID="btnsubmit" runat="server" Text="SAVE" Width="85px" Font-Names="Tahoma"
                                Font-Size="8pt" AutoPostBack="False" ClientInstanceName="btnsubmit">
                                <ClientSideEvents Click="function(s, e) {
									Grid.UpdateEdit();
	                                var pDate = dtkanban.GetText();
                                    var pSupplier = cbosupplier.GetText();
									var pSeq = cboseq.GetText();	
	                                Grid.PerformCallback('save' + '|' + pDate + '|' + pSupplier + '|' + pSeq);
									
									Grid.CancelEdit();
                                }" />
                            </dx:ASPxButton>--%>
                        </td>
                    </tr>
                </table>
            </td>
        </tr>
    </table>
    <asp:SqlDataSource ID="SupplierCode" runat="server" ConnectionString="<%$ ConnectionStrings:KonString %>"
        ProviderName="<%$ ConnectionStrings:KonString.ProviderName %>" SelectCommand="SELECT supplierid AS Supplier_Code , suppliername AS Supplier_Name FROM MS_Supplier ">
    </asp:SqlDataSource>

    <dx:ASPxCallback ID="Approve" runat="server" ClientInstanceName="Approve">
<ClientSideEvents EndCallback="function(s, e) {            
			txtaffiliateentrydate.SetText(s.cpAEDate);
			txtaffiliateentryname.SetText(s.cpAEName); 
			txtaffiliateappdate.SetText(s.cpAADate);
			txtaffiliateappname.SetText(s.cpAAName);
			txtkanban1.SetText(s.cpKanban1);
			txttime1.SetText(s.cpTime1);
			txtkanban2.SetText(s.cpKanban2);
			txttime2.SetText(s.cpTime2);
			txtkanban3.SetText(s.cpKanban3);
			txttime3.SetText(s.cpTime3);
			txtkanban4.SetText(s.cpKanban4);
			txttime4.SetText(s.cpTime4);

            var pMsg = s.cpMessage;
            if (pMsg != '') {
                if (pMsg.substring(1,5) == '1006' || pMsg.substring(1,5) == '1009' ) {
                    lblerrmessage.GetMainElement().style.color = 'Blue';
                } else {
                    lblerrmessage.GetMainElement().style.color = 'Red';
                }
                lblerrmessage.SetText(pMsg);
            } else {
                lblerrmessage.SetText('');
            }

            if (s.cpButton == 'APPROVE') {
                btnApprove.SetEnabled(true);
            } else {
                btnApprove.SetEnabled(false);
            }

            delete s.cpMessage;
        }"></ClientSideEvents>
    </dx:ASPxCallback>
    <br />
</asp:Content>
