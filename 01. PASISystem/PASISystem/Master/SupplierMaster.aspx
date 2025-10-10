<%@ Page Title="" Language="vb" AutoEventWireup="false" MasterPageFile="~/SiteContent.Master" CodeBehind="SupplierMaster.aspx.vb" Inherits="PASISystem.SupplierMaster" %>
<%@ Register Assembly="DevExpress.Web.v14.1, Version=14.1.6.0, Culture=neutral, PublicKeyToken=b88d1754d700e49a" namespace="DevExpress.Web.ASPxGridView.Export" TagPrefix="dx" %>
<%@ Register Assembly="DevExpress.Web.v14.1, Version=14.1.6.0, Culture=neutral, PublicKeyToken=b88d1754d700e49a" namespace="DevExpress.Web.ASPxGridView" TagPrefix="dx" %>
<%@ Register assembly="DevExpress.Web.v14.1, Version=14.1.6.0, Culture=neutral, PublicKeyToken=b88d1754d700e49a" namespace="DevExpress.Web.ASPxEditors" tagprefix="dx" %>
<%@ Register Assembly="DevExpress.Web.v14.1, Version=14.1.6.0, Culture=neutral, PublicKeyToken=b88d1754d700e49a" namespace="DevExpress.Web.ASPxGlobalEvents" TagPrefix="dx" %>

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
        height = height - (height * 45 / 100)
        grid.SetHeight(height);
    }

    function OnBatchEditStartEditing(s, e) {
        currentColumnName = e.focusedColumn.fieldName;

        if (currentColumnName == "NoUrut" || currentColumnName == "SupplierID" || currentColumnName == "SupplierName" || currentColumnName == "Address"
            || currentColumnName == "SupplierType" || currentColumnName == "Overseas"
            || currentColumnName == "City" || currentColumnName == "PostalCode" || currentColumnName == "Phone1" || currentColumnName == "Phone2"
            || currentColumnName == "Fax" || currentColumnName == "NPWP" || currentColumnName == "DetailPage") {
            e.cancel = true;
        }

        currentEditableVisibleIndex = e.visibleIndex;
    }
</script>
</asp:Content>

<asp:Content ID="Content2" ContentPlaceHolderID="Content" runat="server">
    <table style="width:100%;">
        <tr>
            <td colspan="4">
                <table style="border-width: 1pt thin thin thin; border-style: ridge; border-color:#9598A1; width:100%;">
                    <tr>
                        <td height="30">
                            <table id="Table1">
                                <tr>
                                    <td style="width:5px;"></td>
                                    <td align="left" valign="middle" height="25px" width="120px">
                                        <dx:ASPxLabel ID="ASPxLabel1" runat="server" Text="SUPPLIER ID"
                                            Font-Names="Verdana" Font-Size="8pt">
                                        </dx:ASPxLabel>
                                    </td>
                                    <td align="left" valign="middle" height="25px" width="150px">
                                        <dx:ASPxTextBox ID="txtSupplierCode" runat="server" Width="100%" Height="20px"
                                            ClientInstanceName="txtSupplierCode" Font-Names="Verdana"
                                            Font-Size="8pt" MaxLength="10">
                                            <ClientSideEvents TextChanged="function(s, e) {
                                                grid.SetFocusedRowIndex(-1);
                                                grid.PerformCallback('kosong');
	                                            lblErrMsg.SetText('');
                                            }" />
                                        </dx:ASPxTextBox>
                                    </td>
                                    <td align="left" valign="middle" style="height:25px; width:10px;"></td>
                                    <td align="left" valign="middle" height="25px" width="150px"></td> 
                                </tr>
                                <tr>
                                    <td style="width:5px;"></td>
                                    <td align="left" valign="middle" height="25px" width="80px">
                                        <dx:ASPxLabel ID="ASPxLabel2" runat="server" Text="SUPPLIER NAME"
                                            Font-Names="Verdana" Font-Size="8pt">
                                        </dx:ASPxLabel>
                                    </td>
                                    <td align="left" valign="middle" height="25px" width="150px">
                                        <dx:ASPxTextBox ID="txtSupplierName" runat="server" Width="100%" Height="20px"
                                            ClientInstanceName="txtSupplierName" Font-Names="Verdana"
                                            Font-Size="8pt" MaxLength="50">
                                            <ClientSideEvents TextChanged="function(s, e) {
                                                grid.SetFocusedRowIndex(-1);
                                                grid.PerformCallback('kosong');
	                                            lblErrMsg.SetText('');
                                            }" />
                                        </dx:ASPxTextBox>
                                    </td>
                                    <td align="left" valign="middle" height="25px" width="10px"></td>
                                    <td align="left" valign="middle" height="25px" width="150px"></td> 
                                </tr>
                                <tr>
                                    <td style="width:5px;"></td>
                                    <td align="left" valign="middle" height="25px" width="120px">
                                        <dx:ASPxLabel ID="ASPxLabel3" runat="server" Text="SUPPLIER TYPE"
                                            Font-Names="Verdana" Font-Size="8pt">
                                        </dx:ASPxLabel>
                                    </td>
                                    <td align="left" valign="middle" height="25px" width="150px">       
                                        <dx:ASPxComboBox ID="cbSupplierType" runat="server" TextFormatString="{0}" 
                                            DropDownStyle="DropDown" Height="20px" Width="100%" MaxLength="1"
                                            IncrementalFilteringMode="StartsWith" Font-Names="Verdana" 
                                            Font-Size="8pt">
                                        </dx:ASPxComboBox>                                        
                                    </td>
                                    <td align="left" valign="middle" height="25px" width="10px"></td>
                                    <td align="right" valign="middle" style="height:25px; width:100px;">
                                        <table style="width:100%;">
                                            <tr>
                                                <td align="right" valign="middle" style="height:25px; width:90px;">
                                                    <dx:ASPxButton ID="btnRefresh" runat="server" Text="SEARCH"
                                                        Font-Names="Verdana" Width="85px" AutoPostBack="False" Font-Size="8pt" >
                                                        <ClientSideEvents Click="function(s, e) {
                                                            grid.PerformCallback('load');
                                                            lblInfo.SetText('');
                                                        }" />
                                                    </dx:ASPxButton>
                                                </td>
                                                <td align="right" valign="middle" style="height:25px; width:90px;">
                                                    <dx:ASPxButton ID="btnClear" runat="server" Text="CLEAR"
                                                        Font-Names="Verdana" Width="85px" AutoPostBack="False" Font-Size="8pt">
                                                        <ClientSideEvents Click="function(s, e) {
                                                            txtSupplierCode.SetText('');
                                                            txtSupplierName.SetText('');
                                                            lblInfo.SetText('');
                                                            grid.SetFocusedRowIndex(-1);
                                                            grid.PerformCallback('kosong');
                                                        }" />                                   
                                                    </dx:ASPxButton>
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
        </tr>

        <tr>
            <td height="15" colspan="4">
            <%--error message--%>
                <table id="tblMsg" style="border-width: thin; border-style: inset hidden ridge hidden; border-color:#9598A1; width:100%; height: 15px;">
                    <tr>
                        <td align="left" valign="middle" height="15px">
                            <dx:ASPxLabel ID="lblInfo" runat="server" Text="[lblinfo]" Font-Names="Verdana" 
                                ClientInstanceName="lblInfo" Font-Bold="True" Font-Italic="True" 
                                Font-Size="8pt">
                            </dx:ASPxLabel>
                        </td>
                    </tr>         
                </table>
            </td>            
        </tr>
        
        <tr>
            <td valign="top" align="right" style="width: 50px;">                                  
            </td>
            <td valign="top" align="right" style="width: 50px;" colspan="3">    
                <dx:ASPxButton ID="btnADD" runat="server" Text="ADD SUPPLIER"
                    Font-Names="Verdana" Width="85px" AutoPostBack="False" Font-Size="8pt">                               
                </dx:ASPxButton>
            </td>
        </tr>

        <tr>
            <td align="left" valign="top" height="220" colspan="4">
                <%--Column : Grid--%>
                <dx:ASPxGridView ID="grid" runat="server" Width="100%"
                    Font-Names="Verdana" KeyFieldName="SupplierID"
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
                        <dx:GridViewDataTextColumn VisibleIndex="0" Caption=" " Width="60px" FieldName="DetailPage" CellStyle-HorizontalAlign="Center">
                            <DataItemTemplate>
                                <a id="clickElement" href="SupplierMasterDetail.aspx?id=<%#GetRowValue(Container)%>&t1=<%#GetAffiliateID()%>&t2=<%#GetAffiliateName()%>&Session=~/Master/SupplierMaster.aspx"><%# "DETAIL"%></a>
                            </DataItemTemplate>

<CellStyle HorizontalAlign="Center"></CellStyle>
                        </dx:GridViewDataTextColumn>
                        <dx:GridViewDataTextColumn VisibleIndex="1" Caption="No" FieldName="NoUrut" Width="0px" HeaderStyle-HorizontalAlign="Center">
<HeaderStyle HorizontalAlign="Center"></HeaderStyle>
                        </dx:GridViewDataTextColumn>
                        <dx:GridViewDataTextColumn VisibleIndex="2" Caption="SUPPLIER ID" FieldName="SupplierID" Width="70px" HeaderStyle-HorizontalAlign="Center">
<HeaderStyle HorizontalAlign="Center"></HeaderStyle>
                        </dx:GridViewDataTextColumn>
                        <dx:GridViewDataTextColumn VisibleIndex="3" Caption="SUPPLIER NAME" FieldName="SupplierName" Width="210px" HeaderStyle-HorizontalAlign="Center">
<HeaderStyle HorizontalAlign="Center"></HeaderStyle>
                        </dx:GridViewDataTextColumn>
                        <dx:GridViewDataTextColumn VisibleIndex="4" Caption="SUPPLIER TYPE" FieldName="SupplierType" Width="130px" HeaderStyle-HorizontalAlign="Center">
<HeaderStyle HorizontalAlign="Center"></HeaderStyle>
                        </dx:GridViewDataTextColumn>
                        <dx:GridViewDataTextColumn VisibleIndex="5" Caption="SUPPLIER CODE" FieldName="Overseas" Width="70px" HeaderStyle-HorizontalAlign="Center">
<HeaderStyle HorizontalAlign="Center"></HeaderStyle>                        
                        </dx:GridViewDataTextColumn>
                        <dx:GridViewDataTextColumn VisibleIndex="5" Caption="PREFIX LABEL" FieldName="LabelCode" Width="70px" HeaderStyle-HorizontalAlign="Center">
<HeaderStyle HorizontalAlign="Center"></HeaderStyle>                        
                        </dx:GridViewDataTextColumn>
                        <dx:GridViewDataTextColumn VisibleIndex="6" Caption="ADDRESS" FieldName="Address" Width="250px" HeaderStyle-HorizontalAlign="Center">
<HeaderStyle HorizontalAlign="Center"></HeaderStyle>
                        </dx:GridViewDataTextColumn>
                        <dx:GridViewDataTextColumn VisibleIndex="7" Caption="CITY" FieldName="City" Width="100px" HeaderStyle-HorizontalAlign="Center">
<HeaderStyle HorizontalAlign="Center"></HeaderStyle>
                        </dx:GridViewDataTextColumn>
                        <dx:GridViewDataTextColumn VisibleIndex="8" Caption="POSTAL CODE" FieldName="PostalCode" Width="60px" HeaderStyle-HorizontalAlign="Center">
<HeaderStyle HorizontalAlign="Center"></HeaderStyle>
                        </dx:GridViewDataTextColumn>
                        <dx:GridViewDataTextColumn VisibleIndex="9" Caption="PHONE 1" FieldName="Phone1" Width="100px" HeaderStyle-HorizontalAlign="Center">
<HeaderStyle HorizontalAlign="Center"></HeaderStyle>
                        </dx:GridViewDataTextColumn>
                        <dx:GridViewDataTextColumn VisibleIndex="10" Caption="PHONE 2" FieldName="Phone2" Width="100px" HeaderStyle-HorizontalAlign="Center">
<HeaderStyle HorizontalAlign="Center"></HeaderStyle>
                        </dx:GridViewDataTextColumn>
                        <dx:GridViewDataTextColumn VisibleIndex="11" Caption="FAX" FieldName="Fax" Width="100px" HeaderStyle-HorizontalAlign="Center">
<HeaderStyle HorizontalAlign="Center"></HeaderStyle>
                        </dx:GridViewDataTextColumn>
                        <dx:GridViewDataTextColumn VisibleIndex="12" Caption="NPWP" FieldName="NPWP" Width="140px" HeaderStyle-HorizontalAlign="Center">
<HeaderStyle HorizontalAlign="Center"></HeaderStyle>
                        </dx:GridViewDataTextColumn>                        
                    </Columns>
                    <SettingsBehavior AllowSelectByRowClick="True" AllowSort="False"
                        ColumnResizeMode="Control" EnableRowHotTrack="True" />
                    <SettingsPager PageSize="10" Position="Top">
                        <Summary Text="Page {0} of {1} [{2} record(s)]"
                                  AllPagesText="Page {0} of {1} " />
                        <Summary AllPagesText="Page {0} of {1} " Text="Page {0} of {1} [{2} record(s)]"></Summary>
                    </SettingsPager>
                    <SettingsEditing Mode="Batch" NewItemRowPosition="Bottom">
                        <BatchEditSettings ShowConfirmOnLosingChanges="False" />
                    </SettingsEditing>
                    <Settings ShowHorizontalScrollBar="True" ShowVerticalScrollBar="True"
                        ShowGroupButtons="False" ShowStatusBar="Hidden"
                        VerticalScrollableHeight="190" />
                    <SettingsCommandButton EditButton-ButtonType="Link">
                        <EditButton Text="Detail">
                        </EditButton>
                    </SettingsCommandButton>
                    <Styles>
                        <SelectedRow ForeColor="Black">
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
            <td valign="top" align="left" colspan="2">
                <dx:ASPxButton ID="btnSubMenu" runat="server" Text="SUB MENU"
                    Font-Names="Verdana" Width="85px" Font-Size="8pt">
                </dx:ASPxButton>
            </td>
            <td valign="top" align="right" style="width: 50px;">
                <dx:ASPxButton ID="btnUpload" runat="server" Text="UPLOAD" ClientInstanceName="btnUpload"
                    Font-Names="Tahoma" Width="85px" AutoPostBack="False" Font-Size="8pt" >
                    <ClientSideEvents Click="function(s, e) {grid.PerformCallback('save');}" />
                </dx:ASPxButton>
            </td>
            <td valign="top" align="right" style="width: 50px;">
                <dx:ASPxButton ID="btnDownload" runat="server" Text="DOWNLOAD" ClientInstanceName="btnDownload"
                    Font-Names="Tahoma" Width="85px" AutoPostBack="False" Font-Size="8pt" >
                    <ClientSideEvents Click="function(s, e) {grid.PerformCallback('downloadSummary');}" />
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
</asp:Content>

