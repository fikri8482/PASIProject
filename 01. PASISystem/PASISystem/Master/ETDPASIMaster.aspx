<%@ Page Title="" Language="vb" AutoEventWireup="false" MasterPageFile="~/SiteContent.Master" CodeBehind="ETDPASIMaster.aspx.vb" Inherits="PASISystem.ETDPASIMaster" %>
<%@ Register Assembly="DevExpress.Web.v14.1, Version=14.1.6.0, Culture=neutral, PublicKeyToken=b88d1754d700e49a" namespace="DevExpress.Web.ASPxGridView.Export" TagPrefix="dx" %>
<%@ Register Assembly="DevExpress.Web.v14.1, Version=14.1.6.0, Culture=neutral, PublicKeyToken=b88d1754d700e49a" namespace="DevExpress.Web.ASPxGridView" TagPrefix="dx" %>
<%@ Register assembly="DevExpress.Web.v14.1, Version=14.1.6.0, Culture=neutral, PublicKeyToken=b88d1754d700e49a" namespace="DevExpress.Web.ASPxEditors" tagprefix="dx" %>
<%@ Register Assembly="DevExpress.Web.v14.1, Version=14.1.6.0, Culture=neutral, PublicKeyToken=b88d1754d700e49a" namespace="DevExpress.Web.ASPxGlobalEvents" TagPrefix="dx" %>
<%@ Register Assembly="DevExpress.Web.v14.1, Version=14.1.6.0, Culture=neutral, PublicKeyToken=b88d1754d700e49a" Namespace="DevExpress.Web.ASPxCallback" TagPrefix="dx" %>
<%@ Register Assembly="DevExpress.Web.v14.1, Version=14.1.6.0, Culture=neutral, PublicKeyToken=b88d1754d700e49a" Namespace="DevExpress.Web.ASPxUploadControl" TagPrefix="dx" %>
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
        .style13
        {
            width: 140px;
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
            width: 377px;
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

    function OnBatchEditStartEditing(s, e) {
        currentColumnName = e.focusedColumn.fieldName;

        if (currentColumnName == "NoUrut" || currentColumnName == "ETAAffiliate" || currentColumnName == "ETDPASI") {

    } else {
            e.cancel = true;
        currentEditableVisibleIndex = e.visibleIndex;
        }
    }

    function OnGridFocusedRowChanged() {
        grid.GetRowValues(grid.GetFocusedRowIndex(), "ETAAffiliate;ETDPASI;AffiliateID;AffiliateName", OnGetRowValues);
    }

    function OnGetRowValues(values) {
        if (values[0] != "" && values[0] != null && values[0] != "null") {
            
            dtAffiliate.SetText(values[0]);
            dtPASI.SetText(values[1]);
            cboAffiliate2.SetText(values[2]);
            txtAffiliate2.SetText(values[3]);

            txtMode.SetText('update');
           
            cboAffiliate2.GetInputElement().setAttribute('style', 'background:#CCCCCC;foreground:#CCCCCC;');
            cboAffiliate2.GetInputElement().readOnly = true;
            cboAffiliate2.SetEnabled(false);

            dtAffiliate.GetInputElement().setAttribute('style', 'background:#CCCCCC;foreground:#CCCCCC;');
            dtAffiliate.GetInputElement().readOnly = true;
            dtAffiliate.SetEnabled(false);

            lblInfo.SetText('');
            if (cboAffiliate2.GetText() == '') {
                lblInfo.GetMainElement().style.color = 'Red';
                lblInfo.SetText('[6011] Please Select Affiliate Code first!');
                e.processOnServer = false;
                return;
            }
  
        }
    }

    function validasearch() {

                if (DtPeriod.GetText() == "") {
                lblInfo.GetMainElement().style.color = 'Red';
                    lblInfo.SetText("[6011] Please Select Period!");
                    DtPeriod.Focus();
                    e.ProcessOnServer = false;
                    return false;
                }
                
                else if (cboAffiliate.GetText() == "") {
                lblInfo.GetMainElement().style.color = 'Red';
                    lblInfo.SetText("[6011] Please Select Affiliate Code first!");
                    cboAffiliate.Focus();
                    e.ProcessOnServer = false;
                    return false;
                }

                else {
                    lblInfo.SetText('');
                }

        }


    function afterinsert() {
        
        cboAffiliate2.GetInputElement().readOnly = true;
        cboAffiliate2.SetEnabled(false);

        dtAffiliate.GetInputElement().readOnly = true;
        dtAffiliate.SetEnabled(false);
    
    }

    function up_Insert() {
        lblInfo.SetText('');
        if (cboAffiliate2.GetText() == '') {
            lblInfo.GetMainElement().style.color = 'Red';
            lblInfo.SetText('[6011] Please Select Affiliate Code first!');
            e.processOnServer = false;
            return;
        }
        if (dtAffiliate.GetText() == '') {
            lblInfo.GetMainElement().style.color = 'Red';
            lblInfo.SetText('[6011] Please Input E.T.A Affiliate!');
            e.processOnServer = false;
            return;
        }
        
        var pIsUpdate = '';
        
        var pAffiliateID = cboAffiliate2.GetText();
     
        var pStartDate = dtAffiliate.GetValue();
      
        var pEndDate = dtPASI.GetValue();
        
        var vStartDate = pStartDate.getMonth() + '/' + pStartDate.getDate() + '/' + pStartDate.getFullYear();
       
        var vEndDate = pEndDate.getMonth() + '/' + pEndDate.getDate() + '/' + pEndDate.getFullYear();
       
        grid.PerformCallback('save|' + pIsUpdate + '|' + pAffiliateID + '|' + vStartDate + '|' + vEndDate);

    }

    function up_delete() {

        if (cboAffiliate2.GetText() == '') {
            lblInfo.GetMainElement().style.color = 'Red';
            lblInfo.SetText('[6011] Please Select Affiliate Code first!');
            e.processOnServer = false;
            return;
        }

        if (dtAffiliate.GetText() == '') {
            lblInfo.GetMainElement().style.color = 'Red';
            lblInfo.SetText('[6011] Please Input E.T.A Affiliate!');
            e.processOnServer = false;
            return;
        }

        if (grid.GetFocusedRowIndex() == -1) {
            lblInfo.GetMainElement().style.color = 'Red';
            lblInfo.SetText("[6011] Please select the data first!");
            e.ProcessOnServer = false;
            return false;
        }

        var msg = confirm('Are you sure want to delete this data ?');
        if (msg == false) {
            e.processOnServer = false;
            return;
        }

        var pAffiliateID = cboAffiliate2.GetText();

        var pStartDate = dtAffiliate.GetValue();
        var vStartDate = pStartDate.getMonth() + '/' + pStartDate.getDate() + '/' + pStartDate.getFullYear();

        grid.PerformCallback('delete|' + pAffiliateID + '|' + vStartDate);


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
                                        <%--<dx:ASPxComboBox ID="cboPartNo" runat="server" TextFormatString="{0}" 
                                            DropDownStyle="DropDown" Height="20px" Width="100%" MaxLength="1"
                                            IncrementalFilteringMode="StartsWith" Font-Names="Tahoma" 
                                            Font-Size="8pt">
                                        <ClientSideEvents SelectedIndexChanged="function(s, e) {
	                                        txtPartNo.SetText(cboPartNo.GetSelectedItem().GetColumnText(1));	                                        
	                                        grid.PerformCallback('kosong');
	                                        lblErrMsg.SetText('');	
                                        }" />
                                        </dx:ASPxComboBox>--%>
                <dx1:ASPxTimeEdit ID="DtPeriod" runat="server"
                DisplayFormatString="MMM yyyy" EditFormat="Custom" 
                                EditFormatString="MMM yyyy" Font-Names="Tahoma" 
                    Font-Size="8pt" ForeColor="#002060" Width="120px" ClientInstanceName="DtPeriod" TabIndex="1">
                </dx1:ASPxTimeEdit>
                                    </td>
                                    <td align="left" valign="middle" class="style7">
                                        &nbsp;</td>
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
                                        <dx:ASPxLabel ID="ASPxLabel1" runat="server" Text="AFFILIATE CODE"
                                            Font-Names="Tahoma" Font-Size="8pt" Width="100%">
                                        </dx:ASPxLabel>
                                    </td>
                                    <td align="left" valign="middle" class="style18">       
                                        <%--<dx:ASPxComboBox ID="cboAffiliate" runat="server" TextFormatString="{0}" 
                                            DropDownStyle="DropDown" Height="20px" Width="100%" MaxLength="1"
                                            IncrementalFilteringMode="StartsWith" Font-Names="Tahoma" 
                                            Font-Size="8pt">
                                        <ClientSideEvents SelectedIndexChanged="function(s, e) {
	                                        txtAffiliate.SetText(cboAffiliate.GetSelectedItem().GetColumnText(1));	                                        
	                                        grid.PerformCallback('kosong');
	                                        lblErrMsg.SetText('');	
                                        }" />
                                        </dx:ASPxComboBox>  --%>
                                        <dx:ASPxComboBox ID="cboAffiliate" runat="server" 
                                            ClientInstanceName="cboAffiliate" Width="120px"
                                            Font-Size="8pt" 
                                            Font-Names="Tahoma" TextFormatString="{0}" TabIndex="2">
                                            <ClientSideEvents SelectedIndexChanged="function(s, e) {
	                                            txtAffiliate.SetText(cboAffiliate.GetSelectedItem().GetColumnText(1));                              
	                                            lblInfo.SetText('');
cboAffiliate2.GetInputElement().setAttribute('style', 'background:#FFFFFF;foreground:#FFFFFF;');
                        cboAffiliate2.GetInputElement().readOnly = false;
                        cboAffiliate2.SetEnabled(true);

                        dtAffiliate.GetInputElement().setAttribute('style', 'background:#FFFFFF;foreground:#FFFFFF;');
                        dtAffiliate.GetInputElement().readOnly = false;
                        dtAffiliate.SetEnabled(true);
                                            }" CallbackError="function(s, e) {

}" ValueChanged="function(s, e) {

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
                                                        <ClientSideEvents Click="function(s, e) {
               
                lblInfo.SetText('');
                validasearch();                              
				grid.PerformCallback('load');
                cboAffiliate2.GetInputElement().setAttribute('style', 'background:#FFFFFF;foreground:#FFFFFF;');
                        cboAffiliate2.GetInputElement().readOnly = false;
                        cboAffiliate2.SetEnabled(true);

                        dtAffiliate.GetInputElement().setAttribute('style', 'background:#FFFFFF;foreground:#FFFFFF;');
                        dtAffiliate.GetInputElement().readOnly = false;
                        dtAffiliate.SetEnabled(true);
                             
                         
                                                        }" />
                                                    </dx:ASPxButton>
                                    </td>
                                    <td align="left" valign="middle" class="style7">
                                        &nbsp;</td>
                                    <td align="left" valign="middle" class="style7">
                                        &nbsp;</td>
                                    <td align="left" valign="middle" class="style21"></td> 
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
            <td colspan="2" align="left" valign="top" height="100px">
                <%--Column : Grid--%>
                <dx:ASPxGridView ID="grid" runat="server" Width="100%"
                    Font-Names="Tahoma" KeyFieldName="AffiliateID;ETAAffiliate"
                    AutoGenerateColumns="False"
                    ClientInstanceName="grid" Font-Size="8pt" ForeColor="Black">

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
                           
                    }" RowClick="function(s, e) {
	                     
                         delete s.cpMessage;
                         lblInfo.SetText('');  
                       
                    }" />
                    <Columns>
                        <dx:GridViewDataTextColumn VisibleIndex="0" Caption="NO" FieldName="NoUrut" Width="30px" HeaderStyle-HorizontalAlign="Center">
<HeaderStyle HorizontalAlign="Center" Font-Names="Tahoma" Font-Size="8pt"></HeaderStyle>
                            <CellStyle Font-Names="Tahoma" Font-Size="8pt">
                            </CellStyle>
                            <HeaderStyle BackColor="#FFD2A6" ForeColor="Black" />
                        </dx:GridViewDataTextColumn>
                        <dx:GridViewDataDateColumn Caption="E.T.A AFFILIATE" FieldName="ETAAffiliate" 
                            VisibleIndex="2" Width="110px">
                            <PropertiesDateEdit DisplayFormatString="d MMM yyyy">
                            </PropertiesDateEdit>
                            <HeaderStyle Font-Names="Tahoma" Font-Size="8pt" HorizontalAlign="Center" />
                            <CellStyle Font-Names="Tahoma" Font-Size="8pt">
                            </CellStyle>
                            <HeaderStyle BackColor="#FFD2A6" ForeColor="Black" />
                        </dx:GridViewDataDateColumn>
                        <dx:GridViewDataDateColumn Caption="E.T.D PASI" FieldName="ETDPASI" 
                            VisibleIndex="3" Width="110px">
                            <PropertiesDateEdit DisplayFormatString="d MMM yyyy">
                            </PropertiesDateEdit>
                            <HeaderStyle Font-Names="Tahoma" Font-Size="8pt" HorizontalAlign="Center" />
                            <CellStyle Font-Names="Tahoma" Font-Size="8pt">
                            </CellStyle>
                            <HeaderStyle BackColor="#FFD2A6" ForeColor="Black" />
                        </dx:GridViewDataDateColumn>
                         <dx:GridViewDataTextColumn VisibleIndex="4" Caption="AFFILIATE CODE" 
                            FieldName="AffiliateID" Width="0px" HeaderStyle-HorizontalAlign="Center">
<HeaderStyle HorizontalAlign="Center" Font-Names="Tahoma" Font-Size="8pt"></HeaderStyle>
                            <CellStyle Font-Names="Tahoma" Font-Size="8pt">
                            </CellStyle>
                        </dx:GridViewDataTextColumn>
                        <dx:GridViewDataTextColumn Caption="AFFILIATE NAME" FieldName="AffiliateName" 
                            VisibleIndex="5" Width="0px">
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
                        <FilterRowMenuItem ForeColor="Black">
                        </FilterRowMenuItem>
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
                        <td valign="top" bgcolor="#FFD2A6" width="110px">
                            <dx:ASPxLabel ID="ASPxLabel56" runat="server" Text="AFFILIATE CODE"
                                Font-Names="Tahoma" Font-Size="8pt" Width="100px">
                            </dx:ASPxLabel>
                        </td>
                        
                        <td valign="top" bgcolor="#FFD2A6" align="center" class="style50">
                                    <dx:ASPxLabel ID="ASPxLabel5" runat="server" Text="AFFILIATE NAME" 
                                Font-Names="Tahoma" Font-Size="8pt" Width="380px">
                            </dx:ASPxLabel>
                        </td>
                        
                        <td valign="top" class="style12" style="width: 110px;" bgcolor="#FFD2A6">
                            <dx:ASPxLabel ID="ASPxLabel55" runat="server" Text="E.T.A AFFILIATE"
                                Font-Names="Tahoma" Font-Size="8pt" Width="110px">
                            </dx:ASPxLabel>
                        </td> 
                        <td valign="top" class="style13" style="width: 110px;" bgcolor="#FFD2A6">
                            <dx:ASPxLabel ID="ASPxLabel3" runat="server" Text="E.T.D PASI"
                                Font-Names="Tahoma" Font-Size="8pt" Width="100px">
                            </dx:ASPxLabel>
                        </td>   
                        <td valign="top" class="style36" bgcolor="#FFD2A6">
                            &nbsp;</td>                 
                        <td valign="top" class="style36" bgcolor="#FFD2A6">
                            &nbsp;</td>                 
                        <td valign="top" class="style36" bgcolor="#FFD2A6">
                            &nbsp;</td>                 
                        <td valign="top" class="style36" bgcolor="#FFD2A6">
                            &nbsp;</td>                 
                        <td valign="top" class="style36" bgcolor="#FFD2A6">
                            &nbsp;</td>                 
                    </tr>
                    <tr>
                        <td valign="top" align="left" width="110px">
                            <dx:ASPxComboBox ID="cboAffiliate2" runat="server"
                                ClientInstanceName="cboAffiliate2" Width="100px"
                                Font-Size="8pt" 
                                Font-Names="Tahoma" TextFormatString="{0}" TabIndex="4">
                                <ClientSideEvents SelectedIndexChanged="function(s, e) {
                                cboAffiliate.SetText(cboAffiliate2.GetSelectedItem().GetColumnText(0));
	                                txtAffiliate.SetText(cboAffiliate2.GetSelectedItem().GetColumnText(1));
                                    txtAffiliate2.SetText(cboAffiliate2.GetSelectedItem().GetColumnText(1));	                                
	                                lblInfo.SetText('');	
                                }" />
                                <LoadingPanelStyle ImageSpacing="5px">
                                </LoadingPanelStyle>
                            </dx:ASPxComboBox>
                        </td>
                        
                        <td valign="top" class="style50">
                            <dx:ASPxTextBox ID="txtAffiliate2" runat="server" Width="450px" Height="20px"
                                ClientInstanceName="txtAffiliate2"                                 
                                Font-Names="Tahoma" 
                                Font-Size="8pt" MaxLength="50" BackColor="#CCCCCC" ReadOnly="True">
                            </dx:ASPxTextBox>
                        </td>
                        
                        <td valign="top" class="style12" style="width: 110px;">
                            <dx:ASPxDateEdit ID="dtAffiliate" runat="server" 
                                ClientInstanceName="dtAffiliate" Height="21px" 
                                Width="100px"
                                EditFormatString="dd MMM yyyy"
                                Font-Names="Tahoma" Font-Size="8pt" TabIndex="5" >
                            </dx:ASPxDateEdit>
                        </td> 
                        <td valign="top" class="style13" style="width: 110px;">
                            <dx:ASPxDateEdit ID="dtPASI" runat="server" ClientInstanceName="dtPASI"
                            EditFormatString="dd MMM yyyy"
                            Font-Names="Tahoma" Font-Size="8pt" Width="100px" TabIndex="6">
                            </dx:ASPxDateEdit>
                        </td>   
                        <td valign="top" class="style36">
                            &nbsp;</td>                 
                        <td valign="top" class="style36">
                            &nbsp;</td>                 
                        <td valign="top" class="style36">
                            &nbsp;</td>                 
                        <td valign="top" class="style36">
                            &nbsp;</td>                 
                        <td valign="top" class="style36">
                            &nbsp;</td>                 
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
                <dx:ASPxLabel ID="txtMode" runat="server" ClientInstanceName="txtMode" 
                    Width="0px" BackColor="White" ForeColor="White" Visible="true">
                    <Border BorderColor="White" />
                </dx:ASPxLabel>
            </td>
            <td align="left" valign="middle" height="20px" width="380px">
                &nbsp;</td>
            <td valign="top" align="right" style="width: 50px;">
                <dx:ASPxButton ID="btnUpload" runat="server" Text="UPLOAD" ClientInstanceName="btnUpload"
                    Font-Names="Tahoma" Width="85px" AutoPostBack="False" Font-Size="8pt" >                    
                </dx:ASPxButton>
            </td>
            <td valign="top" align="right" style="width: 50px;">
                <dx:ASPxButton ID="btnDownload" runat="server" Text="DOWNLOAD" ClientInstanceName="btnDownload"
                    Font-Names="Tahoma" Width="85px" AutoPostBack="False" Font-Size="8pt" >
                    <ClientSideEvents Click="function(s, e) {grid.PerformCallback('downloadSummary');}" />
                </dx:ASPxButton>
            </td>       
            <td valign="top" align="right" style="width: 50px;">                                   
                <dx:ASPxButton ID="btnClear" runat="server" Text="CLEAR"                    
                    Font-Names="Tahoma"
                    Width="90px" AutoPostBack="False" Font-Size="8pt" TabIndex="8" 
                    ClientInstanceName="btnClear">
                    <clientsideevents click="function(s, e) {
                    
                    lblInfo.SetText('');	
                    dtAffiliate.SetDate(new Date());
                    dtPASI.SetDate(new Date());

                    cboAffiliate.SetText('');
                    txtAffiliate.SetText('');
        
                    cboAffiliate2.SetText('');
                    txtAffiliate2.SetText('');
        
                    txtMode.SetText('new');
        
                    grid.PerformCallback('kosong');
                    lblInfo.SetText('');


                        cboAffiliate2.GetInputElement().setAttribute('style', 'background:#FFFFFF;foreground:#FFFFFF;');
                        cboAffiliate2.GetInputElement().readOnly = false;
                        cboAffiliate2.SetEnabled(true);

                        dtAffiliate.GetInputElement().setAttribute('style', 'background:#FFFFFF;foreground:#FFFFFF;');
                        dtAffiliate.GetInputElement().readOnly = false;
                        dtAffiliate.SetEnabled(true);

                        lblInfo.SetText('');

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
                        grid.PerformCallback('load');
                        
                        dtAffiliate.SetDate(new Date());
                        dtPASI.SetDate(new Date());

                        cboAffiliate.SetText('');
                        txtAffiliate.SetText('');
        
                        cboAffiliate2.SetText('');
                        txtAffiliate2.SetText('');

                        lblInfo.SetText('');
                        cboAffiliate2.GetInputElement().setAttribute('style', 'background:#FFFFFF;foreground:#FFFFFF;');
                        cboAffiliate2.GetInputElement().readOnly = false;
                        cboAffiliate2.SetEnabled(true);
            
                        dtAffiliate.GetInputElement().setAttribute('style', 'background:#FFFFFF;foreground:#FFFFFF;');
                        dtAffiliate.GetInputElement().readOnly = false;
                        dtAffiliate.SetEnabled(true);


                    }" />
                </dx:ASPxButton>
            </td>
            <td valign="top" align="left" style="width: 90px;">                                   
                <dx:ASPxButton ID="btnSubmit" runat="server" Text="SAVE"                    
                    Font-Names="Tahoma"
                    Width="90px" AutoPostBack="False" Font-Size="8pt" TabIndex="7" 
                    ClientInstanceName="btnSubmit">
                    <ClientSideEvents Click="function(s, e) {

                        up_Insert();
                            grid.PerformCallback('load');
                         
                        dtAffiliate.SetDate(new Date());
                        dtPASI.SetDate(new Date());
                        cboAffiliate2.SetText('');
                        txtAffiliate2.SetText('');
                        txtMode.SetText('new');
              
                        cboAffiliate2.GetInputElement().setAttribute('style', 'background:#FFFFFF;foreground:#FFFFFF;');
                        cboAffiliate2.GetInputElement().readOnly = false;
                        cboAffiliate2.SetEnabled(true);

                        dtAffiliate.GetInputElement().setAttribute('style', 'background:#FFFFFF;foreground:#FFFFFF;');
                        dtAffiliate.GetInputElement().readOnly = false;
                        dtAffiliate.SetEnabled(true);
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

                    }
                }else if(s.cpFunction == 'insert'){
             
                }
            } else {
                lblInfo.SetText('');
            }
             delete s.cpMessage;
                        
                        delete s.cpSearch; 
        }" />
    </dx:ASPxCallback>
</asp:Content>

