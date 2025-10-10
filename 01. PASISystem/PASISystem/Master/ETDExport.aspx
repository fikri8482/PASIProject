<%@ Page Title="" Language="vb" AutoEventWireup="false" MasterPageFile="~/SiteContent.Master" CodeBehind="ETDExport.aspx.vb" Inherits="PASISystem.ETDExport" %>
<%@ Register Assembly="DevExpress.Web.v14.1, Version=14.1.6.0, Culture=neutral, PublicKeyToken=b88d1754d700e49a" namespace="DevExpress.Web.ASPxGridView.Export" TagPrefix="dx" %>
<%@ Register Assembly="DevExpress.Web.v14.1, Version=14.1.6.0, Culture=neutral, PublicKeyToken=b88d1754d700e49a" namespace="DevExpress.Web.ASPxGridView" TagPrefix="dx" %>
<%@ Register assembly="DevExpress.Web.v14.1, Version=14.1.6.0, Culture=neutral, PublicKeyToken=b88d1754d700e49a" namespace="DevExpress.Web.ASPxEditors" tagprefix="dx" %>
<%@ Register Assembly="DevExpress.Web.v14.1, Version=14.1.6.0, Culture=neutral, PublicKeyToken=b88d1754d700e49a" namespace="DevExpress.Web.ASPxGlobalEvents" TagPrefix="dx" %>
<%@ Register Assembly="DevExpress.Web.v14.1, Version=14.1.6.0, Culture=neutral, PublicKeyToken=b88d1754d700e49a" Namespace="DevExpress.Web.ASPxCallback" TagPrefix="dx" %>
<%@ Register Assembly="DevExpress.Web.v14.1, Version=14.1.6.0, Culture=neutral, PublicKeyToken=b88d1754d700e49a" Namespace="DevExpress.Web.ASPxHiddenField" TagPrefix="dx" %>
<%@ Register assembly="DevExpress.Web.v14.1, Version=14.1.6.0, Culture=neutral, PublicKeyToken=b88d1754d700e49a" namespace="DevExpress.Web.ASPxEditors" tagprefix="dx1" %>

<asp:Content ID="Content1" ContentPlaceHolderID="head" runat="server">
    <style type="text/css">
        .dxeHLC, .dxeHC, .dxeHFC
        {
            display: none;
        }
        .style7
        {
            height: 25px;
            width: 401px;
        }
        .style12
        {
            width: 119px;
        }
        .style18
        {
            width: 112px;
            height: 25px;
        }
        .style21
        {
            height: 25px;
            width: 80px;
        }
        #Table1
        {
            width: 986px;
            margin-left: 0px;
        }
        .style23
        {
            width: 266px;
            height: 25px;
        }
        .style24
        {
            width: 986px;
        }
        .style36
        {
            width: 95px;
        }
        .style49
        {
            width: 714px;
        }
        .style50
        {
            width: 300px;
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
    function numbersonly(e) {
        var unicode = e.charCode ? e.charCode : e.keyCode
        if (unicode != 8) { //if the key isn't the backspace key (which we should allow)
            if (unicode < 45 || unicode > 57) //if not a number
                return false //disable key press
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

    function OnGridFocusedRowChanged() {
        grid.GetRowValues(grid.GetFocusedRowIndex(), "Week;ETDVendor;ETDPort;ETAPort;ETAFactory;ETAForwarder;DeleteCls", OnGetRowValues);
    }

    function OnGetRowValues(values) {
        var d = new Date(values[1]);
        var e = new Date(values[2]);
        var f = new Date(values[3]);
        var g = new Date(values[4]);
        var h = new Date(values[5]);

        if (d.getFullYear() == "1970") {
            var d = new Date();
        }
        var day1 = d.getDate().toString();
        var month1 = (1 + d.getMonth()).toString();
        var year1 = d.getFullYear();
        var date1 = year1 + '/' + month1 + '/' + day1;

        if (e.getFullYear() == "1970") {
            var e = new Date();
        }
        var day2 = e.getDate().toString();
        var month2 = (1 + e.getMonth()).toString();
        var year2 = e.getFullYear();
        var date2 = year2 + '/' + month2 + '/' + day2;

        if (f.getFullYear() == "1970") {
            var f = new Date();
        }
        var day3 = f.getDate().toString();
        var month3 = (1 + f.getMonth()).toString();
        var year3 = f.getFullYear();
        var date3 = year3 + '/' + month3 + '/' + day3;

        if (g.getFullYear() == "1970") {
            var g = new Date();
        }
        var day4 = g.getDate().toString();
        var month4 = (1 + g.getMonth()).toString();
        var year4 = g.getFullYear();
        var date4 = year4 + '/' + month4 + '/' + day4;

        if (h.getFullYear() == "1970") {
            var h = new Date();
        }
        var day5 = h.getDate().toString();
        var month5 = (1 + h.getMonth()).toString();
        var year5 = h.getFullYear();
        var date5 = year5 + '/' + month5 + '/' + day5;

        if (values[0] != "" && values[0] != null && values[0] != "null") {
            cboWeek.SetText(values[0]);

            dtETDVendor.SetDate(d);
            dtETDPort.SetDate(e);
            dtETAPort.SetDate(f);
            dtETAFactory.SetDate(g);
            dtETAForwarder.SetDate(h);
            var vDeleteCls = values[6];

            cboWeek.GetInputElement().setAttribute('style', 'background:#CCCCCC;foreground:#CCCCCC;');
            cboWeek.GetInputElement().readOnly = true;
            cboWeek.SetEnabled(false);

            if (vDeleteCls == "1") {
                btnSubmit.SetEnabled(false);
                btnUpload.SetEnabled(false);
                btnDownload.SetEnabled(false);
                btnClear.SetEnabled(false);
                btnDelete.SetText("RECOVERY");
                HF.Set('DeleteCls', '1');
            } else {
                btnSubmit.SetEnabled(true);
                btnUpload.SetEnabled(true);
                btnDownload.SetEnabled(true);
                btnClear.SetEnabled(true);
                btnDelete.SetText("DELETE");
                HF.Set('DeleteCls', '0');
            }
//            DtPeriod.GetInputElement().setAttribute('style', 'background:#CCCCCC;foreground:#CCCCCC;');
//            DtPeriod.GetInputElement().readOnly = true;
//            DtPeriod.SetEnabled(false);

        }
    }

    function up_Clear() {
        
//        cboAffiliate2.GetInputElement().setAttribute('style', 'background:#FFFFFF;');
//        cboAffiliate2.GetInputElement().readOnly = false;
//        cboAffiliate2.SetEnabled(true);

//        cboSupplier2.GetInputElement().setAttribute('style', 'background:#FFFFFF;');
//        cboSupplier2.GetInputElement().readOnly = false;
//        cboSupplier2.SetEnabled(true);

        cboWeek.GetInputElement().setAttribute('style', 'background:#FFFFFF;');
        cboWeek.GetInputElement().readOnly = false;
        cboWeek.SetEnabled(true);
//        DtPeriod.GetInputElement().setAttribute('style', 'background:#FFFFFF;');
//        DtPeriod.GetInputElement().readOnly = false;
//        DtPeriod.SetEnabled(true);
    }

    function up_delete() {

        if (cboAffiliate.GetText() == '') {
            lblInfo.GetMainElement().style.color = 'Red';
            lblInfo.SetText('[6011] Please Select Affiliate Code first!');
            e.processOnServer = false;
            return;
        }
        if (cboSupplier.GetText() == '') {
            lblInfo.GetMainElement().style.color = 'Red';
            lblInfo.SetText('[6011] Please Select Supplier Code first!');
            e.processOnServer = false;
            return;
        }

        if (HF.Get('DeleteCls') == "0") {
            var msg = confirm('Are you sure want to delete this data ?');
            if (msg == false) {
                e.processOnServer = false;
                return;
            }
        } else {
            var msg = confirm('Are you sure want to recovery this data ?');
            if (msg == false) {
                e.processOnServer = false;
                return;
            }
        }

        var pWeek = cboWeek.GetValue();

        grid.PerformCallback('delete|' + pWeek);


    }
</script>
</asp:Content>

<asp:Content ID="Content2" ContentPlaceHolderID="Content" runat="server">
    <table style="width:100%;">
        <tr>
            <td>
                <table style="border-left: thin ridge #9598A1; border-right: thin ridge #9598A1; border-top: 1pt ridge #9598A1; border-bottom: thin ridge #9598A1; width:100%;">
                    <tr>
                        <td colspan="8" height="30" class="style24">
                            <table id="Table1" >
                                <tr>
                                    <td align="left" valign="middle" class="style23">
                <dx1:ASPxLabel ID="ASPxLabel69" runat="server" Text="PERIOD" 
                    Font-Names="Tahoma" Font-Size="8pt" ForeColor="#002060">
                </dx1:ASPxLabel>
                                    </td>
                                    <td align="left" valign="middle" class="style18">       
                <dx1:ASPxTimeEdit ID="DtPeriod" runat="server"
                DisplayFormatString="MMM yyyy" EditFormat="Custom" 
                                EditFormatString="MMM yyyy" Font-Names="Tahoma" 
                    Font-Size="8pt" ForeColor="#002060" Width="120px" ClientInstanceName="DtPeriod" TabIndex="1">                    
                </dx1:ASPxTimeEdit>
                                    </td>
                                    <td align="left" valign="middle" class="style7">
                                        <table style="width:100%;">
                                            <tr>
                                                <td>
                <dx1:ASPxLabel ID="ASPxLabel74" runat="server" Text="CUT OFF" 
                    Font-Names="Tahoma" Font-Size="8pt" ForeColor="#002060">
                </dx1:ASPxLabel>
                                                </td>
                                                <td>
                                                    <dx:ASPxDateEdit ID="dtcutoff" runat="server" ClientInstanceName="dtcutoff" 
                                                        DisplayFormatString="dd MMM yyyy" EditFormat="Custom" 
                                                        EditFormatString="dd MMM yyyy" Font-Names="Tahoma" Font-Size="8pt">
                                                        <calendarproperties>
<fastnavproperties enabled="False"></fastnavproperties>
</calendarproperties>
                                                    </dx:ASPxDateEdit>
                                                </td>
                                            </tr>
                                        </table>
                                    </td>
                                    <td align="left" valign="middle" class="style7">
                                        &nbsp;</td>
                                    <td align="left" valign="middle" class="style7">
                                        &nbsp;</td>
                                    <td align="left" valign="middle" class="style7">
                                        &nbsp;</td>
                                    <td align="left" valign="middle" class="style21"></td> 
                                </tr>
                                <tr>
                                    <td align="left" valign="middle" class="style23">
                                        <dx:ASPxLabel ID="ASPxLabel70" runat="server" Font-Names="Tahoma" 
                                            Font-Size="8pt" Text="AFFILIATE CODE" Width="100%">
                                        </dx:ASPxLabel>
                                    </td>
                                    <td align="left" valign="middle" class="style18">       
                                        <dx:ASPxComboBox ID="cboAffiliate" runat="server" 
                                            ClientInstanceName="cboAffiliate" Width="120px"
                                            Font-Size="8pt" 
                                            Font-Names="Tahoma" TextFormatString="{0}" TabIndex="2">
                                            <ClientSideEvents SelectedIndexChanged="function(s, e) {
	txtAffiliate.SetText(cboAffiliate.GetSelectedItem().GetColumnText(1));
	lblInfo.SetText('');
}" CallbackError="function(s, e) {
                                            
}" />
                                            <LoadingPanelStyle ImageSpacing="5px">
                                            </LoadingPanelStyle>
                                        </dx:ASPxComboBox>  
                                    </td>
                                    <td align="left" valign="middle" class="style7">
                                        <dx:ASPxTextBox ID="txtAffiliate" runat="server" Width="400px" Height="20px"
                                            ClientInstanceName="txtAffiliate" Font-Names="Tahoma"
                                            Font-Size="8pt" MaxLength="100" BackColor="#CCCCCC" ReadOnly="True">
                                            <ClientSideEvents TextChanged="function(s, e) {
	lblInfo.SetText('');
}" />
                                        </dx:ASPxTextBox>
                                    </td>
                                    <td align="left" valign="middle" class="style7">
                                                    <dx:ASPxButton ID="btnRefresh" runat="server" Text="SEARCH"
                                                        Font-Names="Tahoma" Width="85px" AutoPostBack="False" Font-Size="8pt" 
                                                        TabIndex="3" >
                                                        <%--<ClientSideEvents Click="function(s, e) {
	if (DtPeriod.GetText() == &quot;&quot;) {
            lblInfo.GetMainElement().style.color = 'Red';
            lblInfo.SetText(&quot;[6011] Please Select Period!&quot;);
            DtPeriod.Focus();
            e.ProcessOnServer = false;
            return false;
        }
    
    if (cboAffiliate.GetText() == &quot;&quot;) {
            lblInfo.GetMainElement().style.color = 'Red';
            lblInfo.SetText(&quot;[6011] Please Select Affiliate Code first!&quot;);
            cboAffiliate.Focus();
            e.ProcessOnServer = false;
            return false;
        }

	if (cboSupplier.GetText() == &quot;&quot;) {
            lblInfo.GetMainElement().style.color = 'Red';
            lblInfo.SetText(&quot;[6011] Please Select Supplier Code first!&quot;);
            cboSupplier.Focus();
            e.ProcessOnServer = false;
            return false;
        }



	grid.PerformCallback('load');


}" />--%>
                                                    </dx:ASPxButton>
                                    </td>
                                    <td align="left" valign="middle" class="style7">
                                        &nbsp;</td>
                                    <td align="left" valign="middle" class="style7">
                                        &nbsp;</td>
                                    <td align="left" valign="middle" class="style21"></td> 
                                </tr>
                                <tr>
                                    <td align="left" valign="middle" class="style23">
                                        <dx:ASPxLabel ID="ASPxLabel1" runat="server" Text="SUPPLIER CODE"
                                            Font-Names="Tahoma" Font-Size="8pt" Width="100%">
                                        </dx:ASPxLabel>
                                    </td>
                                    <td align="left" valign="middle" class="style18">       
                                        <dx:ASPxComboBox ID="cboSupplier" runat="server" 
                                            ClientInstanceName="cboSupplier" Font-Names="Tahoma" Font-Size="8pt" 
                                            TabIndex="1" TextFormatString="{0}" Width="120px">
                                            <ClientSideEvents SelectedIndexChanged="function(s, e) {
	txtSupplier.SetText(cboSupplier.GetSelectedItem().GetColumnText(1));
	lblInfo.SetText('');
}" BeginCallback="function(s, e) {
	lblInfo.SetText('');
}" />
                                            <LoadingPanelStyle ImageSpacing="5px">
                                            </LoadingPanelStyle>
                                        </dx:ASPxComboBox>
                                    </td>
                                    <td align="left" valign="middle" class="style7">
                                        <dx:ASPxTextBox ID="txtSupplier" runat="server" BackColor="#CCCCCC" 
                                            ClientInstanceName="txtSupplier" Font-Names="Tahoma" Font-Size="8pt" 
                                            Height="20px" MaxLength="100" ReadOnly="True" Width="400px">
                                        </dx:ASPxTextBox>
                                    </td>                                    
                                    <td align="left" valign="middle" class="style7">
                                                    &nbsp;</td>
                                    <td align="left" valign="middle" class="style7">
                                        &nbsp;</td>
                                    <td align="left" valign="middle" class="style7">
                                        &nbsp;</td>
                                    <td align="left" valign="middle" class="style21">&nbsp;</td> 
                                </tr>
                                </table>
                        </td>
                    </tr>
                </table>
            </td>
        </tr>
    </table>

    <div style="height:1px;"></div>

    <table style="width:100%; height: 15px;">
        <tr>
            <td colspan="8" height="15">
                <%--error message--%>
                <table id="tblMsg" style="border-width: thin; border-style: inset hidden ridge hidden; border-color:#9598A1; width:100%; height: 15px;">
                    <tr>
                        <td align="left" valign="middle" height="15px">
                            <dx:ASPxLabel ID="lblInfo" runat="server" Font-Names="Tahoma" 
                                ClientInstanceName="lblInfo" Font-Bold="True" Font-Italic="True" 
                                Font-Size="8pt">
                            </dx:ASPxLabel>
                        </td>
                    </tr>         
                </table>
            </td>            
        </tr>
    </table>

    <div style="height:1px;"></div>

    <table style="width:100%;">
        <tr>
            <tr>
                <td align="right">
                    &nbsp
                </td>
                <td align="right">
                    <dx:ASPxImage ID="ASPxImage1" runat="server" ShowLoadingImage="true" ImageUrl="~/Images/fuchsia.jpg"
                        Height="15px" Width="15px">
                    </dx:ASPxImage>
                    <dx:ASPxLabel ID="ASPxLabel20" runat="server" Text=" : DELETE DATA" Font-Names="Tahoma"
                        ClientInstanceName="difference" Font-Bold="True" Font-Size="8pt">
                    </dx:ASPxLabel>
                </td>
            </tr>
            <td colspan="2" align="left" valign="top" height="100px">
                <%--Column : Grid--%>
                <dx:ASPxGridView ID="grid" runat="server" Width="100%"
                    Font-Names="Tahoma" KeyFieldName="RowNumber"
                    AutoGenerateColumns="False"
                    ClientInstanceName="grid" Font-Size="8pt">

<SettingsBehavior AllowSort="False" AllowSelectByRowClick="True" ColumnResizeMode="Control" 
                        EnableRowHotTrack="True"></SettingsBehavior>

<Settings ShowGroupButtons="False" ShowVerticalScrollBar="True" ShowHorizontalScrollBar="True" 
                        VerticalScrollableHeight="190" ShowStatusBar="Hidden"></Settings>

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
                        grid.SetFocusedRowIndex(-1);
                        
                    }" RowClick="function(s, e) {
	                     
                         delete s.cpMessage;
                         lblInfo.SetText('');  
                       
                    }" />
                    <Columns>
                        <dx:GridViewDataTextColumn VisibleIndex="0" Caption="NO" FieldName="RowNumber" Width="35px" HeaderStyle-HorizontalAlign="Center">
                            <HeaderStyle HorizontalAlign="Center" Font-Names="Tahoma" Font-Size="8pt"></HeaderStyle>
                            <CellStyle Font-Names="Tahoma" Font-Size="8pt">
                            </CellStyle>
                            <HeaderStyle BackColor="#FFD2A6" ForeColor="Black" />
                        </dx:GridViewDataTextColumn>
                        <dx:GridViewDataTextColumn Caption="AFFILIATE CODE" FieldName="AffiliateID" 
                            VisibleIndex="1" Width="110px">
                            <HeaderStyle Font-Names="Tahoma" Font-Size="8pt" HorizontalAlign="Center" />
                            <CellStyle Font-Names="Tahoma" Font-Size="8pt">
                            </CellStyle>
                            <HeaderStyle BackColor="#FFD2A6" ForeColor="Black" />
                        </dx:GridViewDataTextColumn>
                        <dx:GridViewDataTextColumn Caption="SUPPLIER CODE" FieldName="SupplierID" 
                            VisibleIndex="2" Width="110px">
                            <HeaderStyle Font-Names="Tahoma" Font-Size="8pt" HorizontalAlign="Center" />
                            <CellStyle Font-Names="Tahoma" Font-Size="8pt">
                            </CellStyle>
                            <HeaderStyle BackColor="#FFD2A6" ForeColor="Black" />
                        </dx:GridViewDataTextColumn>
                        <dx:GridViewDataTextColumn Caption="WEEK" FieldName="Week" 
                            VisibleIndex="3" Width="60px">
                            <HeaderStyle Font-Names="Tahoma" Font-Size="8pt" HorizontalAlign="Center" />
                            <CellStyle Font-Names="Tahoma" Font-Size="8pt">
                            </CellStyle>
                            <HeaderStyle BackColor="#FFD2A6" ForeColor="Black" />
                        </dx:GridViewDataTextColumn>
                        <dx:GridViewDataDateColumn Caption="E.T.D VENDOR" FieldName="ETDVendor" 
                            VisibleIndex="4" Width="110px">
                            <PropertiesDateEdit DisplayFormatString="d MMM yyyy">
                            </PropertiesDateEdit>
                            <HeaderStyle Font-Names="Tahoma" Font-Size="8pt" HorizontalAlign="Center" />
                            <CellStyle Font-Names="Tahoma" Font-Size="8pt">
                            </CellStyle>
                            <HeaderStyle BackColor="#FFD2A6" ForeColor="Black" />
                        </dx:GridViewDataDateColumn>
                        <dx:GridViewDataDateColumn Caption="E.T.A FORWARDER" FieldName="ETAForwarder" 
                            VisibleIndex="4" Width="110px">
                            <PropertiesDateEdit DisplayFormatString="d MMM yyyy">
                            </PropertiesDateEdit>
                            <HeaderStyle Font-Names="Tahoma" Font-Size="8pt" HorizontalAlign="Center" />
                            <CellStyle Font-Names="Tahoma" Font-Size="8pt">
                            </CellStyle>
                            <HeaderStyle BackColor="#FFD2A6" ForeColor="Black" />
                        </dx:GridViewDataDateColumn>
                        <dx:GridViewDataDateColumn Caption="E.T.D PORT" FieldName="ETDPort" 
                            VisibleIndex="5" Width="110px">
                            <PropertiesDateEdit DisplayFormatString="d MMM yyyy">
                            </PropertiesDateEdit>
                            <HeaderStyle Font-Names="Tahoma" Font-Size="8pt" HorizontalAlign="Center" />
                            <CellStyle Font-Names="Tahoma" Font-Size="8pt">
                            </CellStyle>
                            <HeaderStyle BackColor="#FFD2A6" ForeColor="Black" />
                        </dx:GridViewDataDateColumn>
                        <dx:GridViewDataDateColumn Caption="E.T.A PORT" FieldName="ETAPort" 
                            VisibleIndex="6" Width="110px">
                            <PropertiesDateEdit DisplayFormatString="d MMM yyyy">
                            </PropertiesDateEdit>
                            <HeaderStyle Font-Names="Tahoma" Font-Size="8pt" HorizontalAlign="Center" />
                            <CellStyle Font-Names="Tahoma" Font-Size="8pt">
                            </CellStyle>
                            <HeaderStyle BackColor="#FFD2A6" ForeColor="Black" />
                        </dx:GridViewDataDateColumn>
                        <dx:GridViewDataDateColumn Caption="E.T.A FACTORY" FieldName="ETAFactory" 
                            VisibleIndex="7" Width="110px">
                            <PropertiesDateEdit DisplayFormatString="d MMM yyyy">
                            </PropertiesDateEdit>
                            <HeaderStyle Font-Names="Tahoma" Font-Size="8pt" HorizontalAlign="Center" />
                            <CellStyle Font-Names="Tahoma" Font-Size="8pt">
                            </CellStyle>
                            <HeaderStyle BackColor="#FFD2A6" ForeColor="Black" />
                        </dx:GridViewDataDateColumn>
                        <dx:GridViewDataTextColumn Caption="AffiliateName" FieldName="AffiliateName" 
                            VisibleIndex="8" Width="0px">
                            <HeaderStyle Font-Names="Tahoma" Font-Size="8pt" HorizontalAlign="Center" />
                            <CellStyle Font-Names="Tahoma" Font-Size="8pt">
                            </CellStyle>
                            <HeaderStyle BackColor="#FFD2A6" ForeColor="Black" />
                        </dx:GridViewDataTextColumn>
                        <dx:GridViewDataTextColumn Caption="SupplierName" FieldName="SupplierName" 
                            VisibleIndex="9" Width="0px">
                            <HeaderStyle Font-Names="Tahoma" Font-Size="8pt" HorizontalAlign="Center" />
                            <CellStyle Font-Names="Tahoma" Font-Size="8pt">
                            </CellStyle>
                            <HeaderStyle BackColor="#FFD2A6" ForeColor="Black" />
                        </dx:GridViewDataTextColumn>
                        <dx:GridViewDataTextColumn VisibleIndex="9" Caption="REGISTER DATE" 
                            FieldName="EntryDate" Width="150px" HeaderStyle-HorizontalAlign="Center">
                        </dx:GridViewDataTextColumn>
                         <dx:GridViewDataTextColumn VisibleIndex="10" Caption="REGISTER USER" 
                            FieldName="EntryUser" Width="100px" HeaderStyle-HorizontalAlign="Center">
                        </dx:GridViewDataTextColumn>
                         <dx:GridViewDataTextColumn VisibleIndex="11" Caption="UPDATE DATE" 
                            FieldName="UpdateDate" Width="150px" HeaderStyle-HorizontalAlign="Center">
                        </dx:GridViewDataTextColumn>
                         <dx:GridViewDataTextColumn VisibleIndex="12" Caption="UPDATE USER" 
                            FieldName="UpdateUser" Width="100px" HeaderStyle-HorizontalAlign="Center">
                        </dx:GridViewDataTextColumn>
                        <dx:GridViewDataTextColumn VisibleIndex="13" Caption="DeleteCls" 
                            FieldName="DeleteCls" Width="0px" HeaderStyle-HorizontalAlign="Center">
                        </dx:GridViewDataTextColumn>
                    </Columns>
                    <SettingsBehavior AllowSelectByRowClick="True" AllowSort="False"
                        ColumnResizeMode="Control" EnableRowHotTrack="True" />

                    <SettingsPager Visible="False" PageSize="14" 
                        NumericButtonCount="10" AlwaysShowPager="True" mode="ShowAllRecords">
                        <Summary Text="Page {0} of {1} [{2} record(s)]" 
                            AllPagesText="Page {0} of {1} " />
<Summary AllPagesText="Page {0} of {1} " Text="Page {0} of {1} [{2} record(s)]"></Summary>
                    </SettingsPager>
                    <Settings ShowHorizontalScrollBar="True" ShowVerticalScrollBar="True"
                        ShowGroupButtons="False" ShowStatusBar="Hidden"
                        VerticalScrollableHeight="190" />

                    <SettingsCommandButton EditButton-ButtonType="Link">
                        <EditButton Text="Detail">
                        </EditButton>
                    </SettingsCommandButton>
                    <Styles>
                        <Row ForeColor="Black">
                        </Row>
                        <RowHotTrack ForeColor="Black">
                        </RowHotTrack>
                        <PreviewRow ForeColor="Black">
                        </PreviewRow>
                        <SelectedRow ForeColor="Black">
                        </SelectedRow>
                        <FocusedRow ForeColor="Black">
                        </FocusedRow>
                    </Styles>
                    <StylesEditors ButtonEditCellSpacing="0">
                        <ProgressBar Height="21px">
                        </ProgressBar>
                    </StylesEditors>
                </dx:ASPxGridView>
            </td>
        </tr>
    </table>

    <div style="height:8px;"></div>
    
    <table style="width:100%;">
        <tr>
            <td height="50">
                <!-- INPUT AREA -->
                <table id="tbl1" 
                style="border-width: 1pt thin thin thin; border-style: ridge; border-color:#9598A1; width:100%; height:50px;">
                    <tr>
                        <td valign="top" bgcolor="#FFD2A6" 
                            width="110px">
                            <dx:ASPxLabel ID="ASPxLabel72" runat="server" Font-Names="Tahoma" 
                                Font-Size="8pt" Text="Week" Width="110px">
                            </dx:ASPxLabel>
                        </td>
                        <td valign="top" bgcolor="#FFD2A6" width="100px">
                            <dx:ASPxLabel ID="ASPxLabel73" runat="server" Font-Names="Tahoma" 
                                Font-Size="8pt" Text="E.T.D VENDOR" Width="100px">
                            </dx:ASPxLabel>
                        </td>
                        <td valign="top" bgcolor="#FFD2A6" width="100px">
                            <dx:ASPxLabel ID="ASPxLabel4" runat="server" Font-Names="Tahoma" 
                                Font-Size="8pt" Text="E.T.A FORWARDER" Width="100px">
                            </dx:ASPxLabel>
                        </td>                 
                        <td valign="top" bgcolor="#FFD2A6" width="100px">
                            <dx:ASPxLabel ID="ASPxLabel55" runat="server" Font-Names="Tahoma" 
                                Font-Size="8pt" Text="E.T.D PORT" Width="100px">
                            </dx:ASPxLabel>
                        </td>                 
                        <td valign="top" bgcolor="#FFD2A6" width="100px">
                            <dx:ASPxLabel ID="ASPxLabel2" runat="server" Font-Names="Tahoma" 
                                Font-Size="8pt" Text="E.T.A PORT" Width="100px">
                            </dx:ASPxLabel>
                        </td>
                        <td valign="top" bgcolor="#FFD2A6" width="100px">
                            <dx:ASPxLabel ID="ASPxLabel3" runat="server" Font-Names="Tahoma" 
                                Font-Size="8pt" Text="E.T.A FACTORY" Width="100px">
                            </dx:ASPxLabel>
                        </td>               
                    </tr>
                    <tr>
                        <td valign="top" width="110px">
                            <dx:ASPxComboBox ID="cboWeek" runat="server" 
                                ClientInstanceName="cboWeek" Font-Names="Tahoma" Font-Size="8pt" 
                                TabIndex="13" TextFormatString="{0}" Width="110px">
                                <ClientSideEvents SelectedIndexChanged="function(s, e) {	                                
	                                lblInfo.SetText('');	
                                }" />
                                <LoadingPanelStyle ImageSpacing="5px">
                                </LoadingPanelStyle>
                            </dx:ASPxComboBox>
                        </td>
                        <td valign="top" width="100px">
                            <dx:ASPxDateEdit ID="dtETDVendor" runat="server" 
                                ClientInstanceName="dtETDVendor" EditFormatString="dd MMM yyyy" 
                                Font-Names="Tahoma" Font-Size="8pt" Height="21px" TabIndex="5" 
                                Width="100px" DisplayFormatString="dd MMM yyyy">
                            </dx:ASPxDateEdit>
                        </td>
                        <td valign="top" width="100px">
                            <dx:ASPxDateEdit ID="dtETAForwarder" runat="server" ClientInstanceName="dtETAForwarder" 
                                EditFormatString="dd MMM yyyy" Font-Names="Tahoma" Font-Size="8pt" 
                                Height="21px" TabIndex="5" Width="100px" DisplayFormatString="dd MMM yyyy">
                            </dx:ASPxDateEdit>
                        </td>
                        <td valign="top" width="100px">
                            <dx:ASPxDateEdit ID="dtETDPort" runat="server" ClientInstanceName="dtETDPort" 
                                EditFormatString="dd MMM yyyy" Font-Names="Tahoma" Font-Size="8pt" 
                                Height="21px" TabIndex="5" Width="100px" DisplayFormatString="dd MMM yyyy">
                            </dx:ASPxDateEdit>
                        </td>                 
                        <td valign="top" width="100px">
                            <dx:ASPxDateEdit ID="dtETAPort" runat="server" ClientInstanceName="dtETAPort" 
                                EditFormatString="dd MMM yyyy" Font-Names="Tahoma" Font-Size="8pt" 
                                Height="21px" TabIndex="5" Width="100px" DisplayFormatString="dd MMM yyyy">
                            </dx:ASPxDateEdit>
                        </td>  
                        <td valign="top" width="100px">
                            <dx:ASPxDateEdit ID="dtETAFactory" runat="server" ClientInstanceName="dtETAFactory" 
                                EditFormatString="dd MMM yyyy" Font-Names="Tahoma" Font-Size="8pt" 
                                Height="21px" TabIndex="5" Width="100px" DisplayFormatString="dd MMM yyyy">
                            </dx:ASPxDateEdit>
                        </td>                  
                    </tr>
                </table>
                </td>
        </tr>
    </table> 
    
    <div style="height:8px;"></div>

    <table id="button" style=" width:100%;">
        <tr>
            <td valign="top" align="left" class="style49">
                <dx:ASPxButton ID="btnSubMenu" runat="server" Text="SUB MENU"
                    Font-Names="Tahoma" Width="85px" Font-Size="8pt" TabIndex="9" 
                    ClientInstanceName="btnSubMenu">
                </dx:ASPxButton>
            </td>
            <td valign="top" align="right" style="width: 0px;" bgcolor="White">      
                &nbsp;</td>
            <td valign="top" align="right" style="width: 50px;">
                <dx:ASPxButton ID="btnUpload" runat="server" Text="UPLOAD" ClientInstanceName="btnUpload"
                    Font-Names="Tahoma" Width="85px" AutoPostBack="False" Font-Size="8pt" >
                </dx:ASPxButton>
            </td>
            <td valign="top" align="right" style="width: 50px;">
                <dx:ASPxButton ID="btnDownload" runat="server" Text="DOWNLOAD" ClientInstanceName="btnDownload"
                    Font-Names="Tahoma" Width="85px" AutoPostBack="False" Font-Size="8pt" >
                    <ClientSideEvents Click="function(s, e) {
    if (DtPeriod.GetText() == &quot;&quot;) {
            lblInfo.GetMainElement().style.color = 'Red';
            lblInfo.SetText(&quot;[6011] Please Select Period!&quot;);
            DtPeriod.Focus();
            e.ProcessOnServer = false;
            return false;
        }

	if (cboAffiliate.GetText() == &quot;&quot;) {
            lblInfo.GetMainElement().style.color = 'Red';
            lblInfo.SetText(&quot;[6011] Please Select Affiliate Code first!&quot;);
            cboAffiliate.Focus();
            e.ProcessOnServer = false;
            return false;
        }

grid.PerformCallback('downloadSummary');
	}" />
                </dx:ASPxButton>
            </td>                   
            <td valign="top" align="right" style="width: 50px;">                                   
                <dx:ASPxButton ID="btnClear" runat="server" Text="CLEAR"                    
                    Font-Names="Tahoma"
                    Width="90px" AutoPostBack="False" Font-Size="8pt" TabIndex="8" 
                    ClientInstanceName="btnClear">
                    <clientsideevents click="function(s, e) {
                        cboAffiliate.SetText('');
                        txtAffiliate.SetText('');
                        cboSupplier.SetText('');
                        txtSupplier.SetText('');
                        cboWeek.SetText('');
                        dtcutoff.SetDate(new Date());
                        dtETDVendor.SetDate(new Date());
                        dtETAForwarder.SetDate(new Date());
                        dtETDPort.SetDate(new Date());
                        dtETAPort.SetDate(new Date());
                        dtETAFactory.SetDate(new Date());
                        DtPeriod.SetDate(new Date());
        
                        grid.PerformCallback('loadeventchange');
                        lblInfo.SetText('');

                        DtPeriod.GetInputElement().setAttribute('style', 'background:#FFFFFF;');
                        DtPeriod.GetInputElement().readOnly = false;
                        DtPeriod.SetEnabled(true);


}" GotFocus="function(s, e) {
lblInfo.SetText('');
}" LostFocus="function(s, e) {
	lblInfo.SetText('');
}" />
                </dx:ASPxButton>
            </td>
            <td valign="top" align="right" style="width: 90px;">                                   
                <dx:ASPxButton ID="btnDelete" runat="server" Text="DELETE" Font-Names="Tahoma" Width="90px"
                    AutoPostBack="False" Font-Size="8pt" TabIndex="18" ClientInstanceName="btnDelete">
                    <ClientSideEvents Click="function(s, e) {
                        
                        up_delete();
                       
                        cboWeek.SetText('');
                        dtcutoff.SetDate(new Date());
                        dtETDVendor.SetDate(new Date());
                        dtETAForwarder.SetDate(new Date());
                        dtETDPort.SetDate(new Date());
                        dtETAPort.SetDate(new Date());
                        dtETAFactory.SetDate(new Date());
                        
                        up_Clear();    


                    }" />
                </dx:ASPxButton>
            </td>
            <td valign="top" align="left" style="width: 90px; margin-left: 40px;">                                   
                <dx:ASPxButton ID="btnSubmit" runat="server" Text="SAVE"                    
                    Font-Names="Tahoma"
                    Width="90px" AutoPostBack="False" Font-Size="8pt" TabIndex="7" 
                    ClientInstanceName="btnSubmit">
                    <ClientSideEvents Click="function(s, e) {
        if (cboAffiliate.GetText() == '') {
            lblInfo.GelblInfotMainElement().style.color = 'Red';
            lblInfo.SetText('[6011] Please Select Affiliate Code first!');
            e.processOnServer = false;
            return;
        }

        if (cboSupplier.GetText() == '') {
            lblInfo.GetMainElement().style.color = 'Red';
            lblInfo.SetText('[6011] Please Select Supplier Code first!');
            e.processOnServer = false;
            return;
        }



		var pIsUpdate = '';

        var pAffiliateID = cboAffiliate.GetText();
        var pSupplierID = cboSupplier.GetText();
		var pPeriod = DtPeriod.GetValue();
        var pWeek = cboWeek.GetValue();
        var pETDVendor = dtETDVendor.GetValue();
        var pETAForwarder = dtETAForwarder.GetValue();
        var pETDPort = dtETDPort.GetValue();
        var pETAPort = dtETAPort.GetValue();
        var pETAFactory = dtETAFactory.GetValue();


        if (grid.GetFocusedRowIndex() == -1) {
            pIsUpdate = 'new';
            grid.PerformCallback('save|' + pIsUpdate + '|' + pPeriod + '|' + pAffiliateID + '|' + pSupplierID + '|' + pWeek + '|' + pETDVendor + '|' + pETDPort + '|' + pETAPort + '|' + pETAFactory + '|' + pETAForwarder);

            } else {
                pIsUpdate = 'edit';        
                grid.PerformCallback('save|' + pIsUpdate + '|' + pPeriod + '|' + pAffiliateID + '|' + pSupplierID + '|' + pWeek + '|' + pETDVendor + '|' + pETDPort + '|' + pETAPort + '|' + pETAFactory + '|' + pETAForwarder);
    	        
        }

                        cboWeek.SetText('');
                        dtETAForwarder.SetDate(new Date());
                        dtETDVendor.SetDate(new Date());
                        dtETDPort.SetDate(new Date());
                        dtETAPort.SetDate(new Date());
                        dtETAFactory.SetDate(new Date());
                        
                        up_Clear();                                              
                                            

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
