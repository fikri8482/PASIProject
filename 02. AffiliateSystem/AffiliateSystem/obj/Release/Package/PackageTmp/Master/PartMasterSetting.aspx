<%@ Page Title="" Language="vb" AutoEventWireup="false" MasterPageFile="~/SiteContent.Master" CodeBehind="PartMasterSetting.aspx.vb" Inherits="AffiliateSystem.PartMasterSetting" %>
<%@ Register Assembly="DevExpress.Web.v14.1, Version=14.1.6.0, Culture=neutral, PublicKeyToken=b88d1754d700e49a" namespace="DevExpress.Web.ASPxGridView.Export" TagPrefix="dx" %>
<%@ Register Assembly="DevExpress.Web.v14.1, Version=14.1.6.0, Culture=neutral, PublicKeyToken=b88d1754d700e49a" namespace="DevExpress.Web.ASPxGridView" TagPrefix="dx" %>
<%@ Register assembly="DevExpress.Web.v14.1, Version=14.1.6.0, Culture=neutral, PublicKeyToken=b88d1754d700e49a" namespace="DevExpress.Web.ASPxEditors" tagprefix="dx" %>
<%@ Register Assembly="DevExpress.Web.v14.1, Version=14.1.6.0, Culture=neutral, PublicKeyToken=b88d1754d700e49a" namespace="DevExpress.Web.ASPxGlobalEvents" TagPrefix="dx" %>
<%@ Register Assembly="DevExpress.Web.v14.1, Version=14.1.6.0, Culture=neutral, PublicKeyToken=b88d1754d700e49a" Namespace="DevExpress.Web.ASPxCallback" TagPrefix="dx" %>

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
        .style18
        {
            width: 112px;
            height: 25px;
        }
        .style20
        {
            width: 100px;
            height: 25px;
        }
        .style21
        {
            height: 25px;
            width: 80px;
        }
        #Table1
        {
            width: 100%;
            margin-left: 0px;
        }
        .style23
        {
            width: 94px;
            height: 25px;
        }
        .style24
        {
            width: 100%;
        }
        .style25
        {
            width: 100%;
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

    function OnBatchEditStartEditing(s, e) {

        currentColumnName = e.focusedColumn.fieldName;

        if (currentColumnName == "NoUrut" || currentColumnName == "PartNo" || currentColumnName == "PartName") {

            e.cancel = true;
        }

        currentEditableVisibleIndex = e.visibleIndex;
    }

    function validasearch() {
        lblInfo.GetMainElement().style.color = 'Red';
        if (cboPartNo.GetText() == "") {
            lblInfo.SetText("[6011] Please Select Part No. first!");
            cboPartNo.Focus();
            e.ProcessOnServer = false;
            return false;

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
        height = height - (height * 35 / 100)
        grid.SetHeight(height);
       
    }

    function OnGridFocusedRowChanged() {
        grid.GetRowValues(grid.GetFocusedRowIndex(), 'PartNo', OnGetRowValues);
    }

</script>
</asp:Content>

<asp:Content ID="Content2" ContentPlaceHolderID="Content" runat="server">
    <table style="width:100%;">
        <tr>
            <td>
                <table style="border-left: thin ridge #9598A1; border-right: thin ridge #9598A1; border-top: 1pt ridge #9598A1; border-bottom: thin ridge #9598A1;" 
                    class="style25">
                    <tr>
                        <td class="style24">
                            <table id="Table1" >
                                <tr>
                                    <td align="left" valign="middle" class="style23">
                                        <dx:ASPxLabel ID="ASPxLabel4" runat="server" Text="PART NO."
                                            Font-Names="Verdana" Font-Size="8pt">
                                        </dx:ASPxLabel>
                                    </td>
                                    <td align="left" valign="middle" class="style18">       
                                        <%--<dx:ASPxComboBox ID="cboPartNo" runat="server" TextFormatString="{0}" 
                                            DropDownStyle="DropDown" Height="20px" Width="100%" MaxLength="1"
                                            IncrementalFilteringMode="StartsWith" Font-Names="Verdana" 
                                            Font-Size="8pt">
                                        <ClientSideEvents SelectedIndexChanged="function(s, e) {
	                                        txtPartNo.SetText(cboPartNo.GetSelectedItem().GetColumnText(1));	                                        
	                                        grid.PerformCallback('kosong');
	                                        lblErrMsg.SetText('');	
                                        }" />
                                        </dx:ASPxComboBox>--%>
                                        <dx:ASPxComboBox ID="cboPartNo" runat="server" 
                                            ClientInstanceName="cboPartNo" Width="110px"
                                            Font-Size="8pt" 
                                            Font-Names="Verdana" TextFormatString="{0}">
                                            <ClientSideEvents SelectedIndexChanged="function(s, e) {
	                                            txtPartNo.SetText(cboPartNo.GetSelectedItem().GetColumnText(1));	                              
	                                            lblInfo.SetText('');
                                            }" />
                                            <LoadingPanelStyle ImageSpacing="5px">
                                            </LoadingPanelStyle>
                                        </dx:ASPxComboBox>
                                    </td>
                                    <td align="left" valign="middle" class="style7">
                                        <dx:ASPxTextBox ID="txtPartNo" runat="server" Width="400px" Height="20px"
                                            ClientInstanceName="txtPartNo" Font-Names="Verdana"
                                            Font-Size="8pt" MaxLength="100" BackColor="#CCCCCC" ReadOnly="True" 
                                            style="margin-right: 31px">
                                        </dx:ASPxTextBox>
                                    </td>
                                    <td align="left" valign="middle" class="style20">
                                        &nbsp;</td>
                                    <td align="left" valign="middle" class="style20">
                                        &nbsp;</td>
                                    <td align="left" valign="middle" class="style21">
                                                    <dx:ASPxButton ID="btnRefresh" runat="server" Text="SEARCH"
                                                        Font-Names="Verdana" Width="85px" AutoPostBack="False" Font-Size="8pt" 
                                                        TabIndex="8" >
                                                        <ClientSideEvents Click="function(s, e) {
validasearch(); 
                                                            grid.PerformCallback('load');
                                                            lblInfo.SetText('');
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

    <div style="height:1px;"></div>

    <table style="width:100%; height: 15px;">
        <tr>
            <td colspan="8" height="15">
                <%--error message--%>
                <table id="tblMsg" style="border-width: thin; border-style: inset hidden ridge hidden; border-color:#9598A1; width:100%; height: 15px;">
                    <tr>
                        <td align="left" valign="middle" height="15px">
                            <dx:ASPxLabel ID="lblInfo" runat="server" Font-Names="Verdana" 
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
            <td colspan="2" align="left" valign="top" height="100px">
                <%--Column : Grid--%>
                <dx:ASPxGridView ID="grid" runat="server" Width="100%"
                    Font-Names="Verdana" KeyFieldName="PartNo"
                    AutoGenerateColumns="False"
                    ClientInstanceName="grid" Font-Size="8pt">
                    <ClientSideEvents Init="function(s, e) {
	OnInit
}" BatchEditStartEditing="function(s, e) {
	OnBatchEditStartEditing
}" CallbackError="function(s, e) {
	e.handled = true;
}" EndCallback="function(s, e) {
	 var pMsg = s.cpMessage;        
                        if (pMsg != '') {
                            if (s.cpType == 'error'){
                                lblInfo.GetMainElement().style.color = 'Blue';
                            }
                            else if (s.cpType == 'info'){
                                lblInfo.GetMainElement().style.color = 'Blue';
                            }
                            else {
                                lblInfo.GetMainElement().style.color = 'Red';
                            }
        
                            lblInfo.SetText(pMsg);
}
}" RowClick="function(s, e) {
	lblInfo.SetText('');
}" />
<ClientSideEvents RowClick="function(s, e) {
	lblInfo.SetText(&#39;&#39;);
}" BatchEditStartEditing="function(s, e) {
	OnBatchEditStartEditing
}" EndCallback="function(s, e) {
	 var pMsg = s.cpMessage;        
                        if (pMsg != &#39;&#39;) {
                            if (s.cpType == &#39;error&#39;){
                                lblInfo.GetMainElement().style.color = &#39;Blue&#39;;
                            }
                            else if (s.cpType == &#39;info&#39;){
                                lblInfo.GetMainElement().style.color = &#39;Blue&#39;;
                            }
                            else {
                                lblInfo.GetMainElement().style.color = &#39;Red&#39;;
                            }
        
                            lblInfo.SetText(pMsg);
}
}" CallbackError="function(s, e) {
	e.handled = true;
}" Init="OnInit"></ClientSideEvents>
                    <Columns>
                        <dx:GridViewDataTextColumn VisibleIndex="0" Caption="NO" FieldName="NoUrut" Width="30px" HeaderStyle-HorizontalAlign="Center">
<HeaderStyle HorizontalAlign="Center"></HeaderStyle>

                            <CellStyle Font-Names="Verdana" Font-Size="8pt">
                            </CellStyle>
                        </dx:GridViewDataTextColumn>
                        <dx:GridViewDataTextColumn VisibleIndex="1" Caption="PART NO." 
                            FieldName="PartNo" Width="110px" HeaderStyle-HorizontalAlign="Center">
<HeaderStyle HorizontalAlign="Center"></HeaderStyle>

                            <CellStyle Font-Names="Verdana" Font-Size="8pt">
                            </CellStyle>
                        </dx:GridViewDataTextColumn>
                        <dx:GridViewDataTextColumn VisibleIndex="2" Caption="PART NAME" FieldName="PartName" Width="210px" HeaderStyle-HorizontalAlign="Center">
<HeaderStyle HorizontalAlign="Center"></HeaderStyle>

                            <CellStyle Font-Names="Verdana" Font-Size="8pt">
                            </CellStyle>
                        </dx:GridViewDataTextColumn>
                        <dx:GridViewDataTextColumn VisibleIndex="3" Caption="MOQ" 
                            FieldName="MOQ" Width="110px" HeaderStyle-HorizontalAlign="Center">
                            <PropertiesTextEdit DisplayFormatString="{0:n0}">
                            </PropertiesTextEdit>

<HeaderStyle HorizontalAlign="Center"></HeaderStyle>

                            <CellStyle Font-Names="Verdana" Font-Size="8pt">
                            </CellStyle>
                        </dx:GridViewDataTextColumn>
                        <dx:GridViewDataCheckColumn Caption="SHOW" VisibleIndex="4" FieldName="ShowCls">
                            <PropertiesCheckEdit ValueChecked="1" ValueType="System.String" 
                                ValueUnchecked="0">
                            </PropertiesCheckEdit>
                        </dx:GridViewDataCheckColumn>
                        <dx:GridViewDataTextColumn VisibleIndex="5" Caption="LOCATION" FieldName="LocationID" Width="210px" HeaderStyle-HorizontalAlign="Left">
<HeaderStyle HorizontalAlign="Left"></HeaderStyle>

                            <CellStyle Font-Names="Verdana" Font-Size="8pt">
                            </CellStyle>
                        </dx:GridViewDataTextColumn>
                    </Columns>
                    <SettingsBehavior AllowSelectByRowClick="True" AllowSort="False"
                        ColumnResizeMode="Control" EnableRowHotTrack="True" />

<SettingsBehavior AllowSort="False" AllowSelectByRowClick="True" ColumnResizeMode="Control" 
                        EnableRowHotTrack="True"></SettingsBehavior>

                    <SettingsPager Visible="False" PageSize="14" 
                        NumericButtonCount="10" AlwaysShowPager="True" mode="ShowAllRecords">
                        <Summary Text="Page {0} of {1} [{2} record(s)]" 
                            AllPagesText="Page {0} of {1} " />
<Summary AllPagesText="Page {0} of {1} " Text="Page {0} of {1} [{2} record(s)]"></Summary>
                    </SettingsPager>
                    <SettingsEditing Mode="Batch" NewItemRowPosition="Bottom" 
                        EditFormColumnCount="10">
                        <BatchEditSettings ShowConfirmOnLosingChanges="False" />
<BatchEditSettings ShowConfirmOnLosingChanges="False"></BatchEditSettings>
                    </SettingsEditing>
                    <Settings ShowHorizontalScrollBar="True" ShowVerticalScrollBar="True"
                        ShowGroupButtons="False" ShowStatusBar="Hidden"
                        VerticalScrollableHeight="190" />

<Settings ShowGroupButtons="False" ShowVerticalScrollBar="True" ShowHorizontalScrollBar="True" 
                        VerticalScrollableHeight="190" ShowStatusBar="Hidden"></Settings>

                    <SettingsCommandButton EditButton-ButtonType="Link">
                        <EditButton Text="Detail">
                        </EditButton>
                    </SettingsCommandButton>
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
            <td>
    <table id="table2" style=" width:100%;">
        <tr>
            <td valign="top" align="left">
                <dx:ASPxButton ID="btnSubMenu" runat="server" Text="SUB MENU"
                    Font-Names="Verdana" Width="85px" Font-Size="8pt" TabIndex="20">
                </dx:ASPxButton>
            </td>
            <td valign="top" align="right" style="width: 50px;">                                   
                <dx:ASPxButton ID="btnUpload" runat="server" Text="UPLOAD"                    
                    Font-Names="Verdana"
                    Width="90px" AutoPostBack="False" Font-Size="8pt" TabIndex="19">
                </dx:ASPxButton>
            </td>

            <td valign="top" align="right" style="width: 50px;">                                   
                <dx:ASPxButton ID="btnClear" runat="server" Text="CLEAR"                    
                    Font-Names="Verdana"
                    Width="90px" AutoPostBack="False" Font-Size="8pt" TabIndex="19">
                    <ClientSideEvents Click="function(s, e) {
                    cboPartNo.SetText('== ALL ==');
                    txtPartNo.SetText('== ALL ==');
grid.PerformCallback('kosong');
lblInfo.SetText('');

}" />
                </dx:ASPxButton>
            </td>
            <td valign="top" align="right" style="width: 50px;">                                   
                <dx:ASPxButton ID="btnExcel" runat="server" Text="EXCEL"                    
                    Font-Names="Verdana"
                    Width="90px" AutoPostBack="False" Font-Size="8pt" TabIndex="17">
                    <ClientSideEvents Click="function(s, e) {
                        grid.PerformCallback('downloadSummary');
                       }" />
                </dx:ASPxButton>
            </td>
            <td valign="top" align="right" style="width: 50px;">                                   
                <dx:ASPxButton ID="btnSubmit" runat="server" Text="SAVE"                    
                    Font-Names="Verdana"
                    Width="90px" AutoPostBack="False" Font-Size="8pt" TabIndex="17">
                    <ClientSideEvents Click="function(s, e) {
                        grid.UpdateEdit();
                        grid.PerformCallback('load');

                       }" />
                </dx:ASPxButton>
            </td>
        </tr>
    </table>
     </td>
        </tr>
    </table>
    
                    
    <dx:ASPxGlobalEvents ID="ge" runat="server" >
        <ClientSideEvents ControlsInitialized="function(s, e) {
	        OnControlsInitializedSplitter();
	        OnControlsInitializedGrid();
        }" />
    </dx:ASPxGlobalEvents>

    <dx:ASPxCallback ID="AffiliateSubmit" runat="server" ClientInstanceName = "AffiliateSubmit">
        <ClientSideEvents EndCallback="function(s, e) {
            var pMsg = s.cpMessage;        
            if (pMsg != '') {
                if (s.cpType == 'error'){
                    lblInfo.GetMainElement().style.color = 'Red';
                }
                else if (s.cpType == 'info'){
                    lblInfo.GetMainElement().style.color = 'Blue';
                }
                else {
                    lblInfo.GetMainElement().style.color = 'Red';
                }
        
                lblInfo.SetText(pMsg);

                if (s.cpFunction == 'delete'){
                    if (s.cpType != 'error'){
                        clear2();
                    }
                }else if(s.cpFunction == 'insert'){

                    clear2();

                }
                
            } else {
                lblInfo.SetText('');
            }  
        }" />
    </dx:ASPxCallback>
</asp:Content>

