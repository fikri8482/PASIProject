<%@ Page Title="" Language="vb" AutoEventWireup="false" MasterPageFile="~/SiteContent.Master" CodeBehind="PartConversion.aspx.vb" Inherits="PASISystem.PartConversion" %>
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
        height = height - (height * 53 / 100)
        grid.SetHeight(height);
    }

    function OnBatchEditStartEditing(s, e) {
        currentColumnName = e.focusedColumn.fieldName;

        if (currentColumnName == "NoUrut" || currentColumnName == "FinishGoodNo" || currentColumnName == "FinishGoodName" || currentColumnName == "FGUnitCls"
            || currentColumnName == "FGQty" || currentColumnName == "PartNo"
            || currentColumnName == "PartName" || currentColumnName == "PartUnitCls" || currentColumnName == "PartQty") {
            e.cancel = true;
        }

        currentEditableVisibleIndex = e.visibleIndex;
    }

    function OnGridFocusedRowChanged() {
        grid.GetRowValues(grid.GetFocusedRowIndex(), "FinishGoodNo;FinishGoodName;FGUnitCls;FGQty;PartNo;PartName;PartUnitCls;PartQty", OnGetRowValues);
    }
    function OnGetRowValues(values) {
        if (values[0] != "" && values[0] != null && values[0] != "null") {

            cboFGNo2.SetText(values[0]);
            txtFGNo2.SetText(values[1]);
            txtFGUnit.SetText(values[2]);
            txtFGQty.SetText(values[3]);
            cboPartNo2.SetText(values[4]);
            txtPartNo2.SetText(values[5]);
            txtPartUnit.SetText(values[6]);
            txtPartQty.SetText(values[7]);
            txtModee.SetText('update');

            lblInfo.SetText('');
            cboFGNo2.GetInputElement().setAttribute('style', 'background:#CCCCCC;foreground:#CCCCCC;');
            cboFGNo2.GetInputElement().readOnly = true;
            cboFGNo2.SetEnabled(false);

            cboPartNo2.GetInputElement().setAttribute('style', 'background:#CCCCCC;foreground:#CCCCCC;');
            cboPartNo2.GetInputElement().readOnly = true;
            cboPartNo2.SetEnabled(false);          
           
        }
    }

    function up_delete() {
        if (cboPartNo2.GetText() == "") {
            lblInfo.GetMainElement().style.color = 'Red';
            lblInfo.SetText("[6011] Please select the data first!");
            e.ProcessOnServer = false;
            return false;
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

        var pPartNo = cboPartNo2.GetText();
        var pFGNo = cboFGNo2.GetText();

        grid.PerformCallback('delete|' + pFGNo + '|' + pPartNo);
    }

    function validasubmit() {
        lblInfo.GetMainElement().style.color = 'Red';
        if (cboPartNo2.GetText() == "") {
            lblInfo.SetText("[6011] Please Select Part No. first!");
            cboPartNo2.Focus();
            e.ProcessOnServer = false;
            return false;
        }

        if (cboFGNo2.GetText() == "") {
            lblInfo.SetText("[6011] Please Select Finish Good No. first!");
            cboFGNo2.Focus();
            e.ProcessOnServer = false;
            return false;
        }       
    }

    function up_Insert() {
        var pIsUpdate = '';
        var pPartID = cboPartNo2.GetSelectedItem().GetColumnText(0);
        var pFGID = cboFGNo2.GetSelectedItem().GetColumnText(0);
        var pFGQty = txtFGQty.GetValue();
        var pPartQty = txtPartQty.GetValue();

        grid.PerformCallback('save|' + pIsUpdate + '|' + pPartID + '|' + pFGID + '|' + pFGQty + '|' + pPartQty);
    }

    function singlequote(e) {
        var unicode = e.charCode ? e.charCode : e.keyCode
        if (unicode == 39) {
            return false //disable key press
        }
    }

    function numbersonly(e) {
        var unicode = e.charCode ? e.charCode : e.keyCode
        if (unicode != 8) { //if the key isn't the backspace key (which we should allow)
            if (unicode < 45 || unicode > 57) //if not a number
                return false //disable key press
        }
    }
</script>
</asp:Content>

<asp:Content ID="Content2" ContentPlaceHolderID="Content" runat="server">
    <table style="width:100%;">
        <tr>
            <td>
                <table style="border-width: 1pt thin thin thin; border-style: ridge; border-color:#9598A1; width:100%;">
                    <tr>
                        <td colspan="8" height="30">
                            <table id="Table1">
                                <tr>
                                    <td style="width:5px;"></td>
                                    <td align="left" valign="middle" height="25px" width="120px">
                                        <dx:ASPxLabel ID="ASPxLabel4" runat="server" Text="FINISH GOOD NO."
                                            Font-Names="Verdana" Font-Size="8pt">
                                        </dx:ASPxLabel>
                                    </td>
                                    <td align="left" valign="middle" height="25px" width="150px">       
                                        <dx:ASPxComboBox ID="cboFGNo" runat="server" 
                                            ClientInstanceName="cboFGNo" Width="100%"
                                            Font-Size="8pt" 
                                            Font-Names="Verdana" TextFormatString="{0}">
                                            <ClientSideEvents SelectedIndexChanged="function(s, e) {
	                                            txtFGNo.SetText(cboFGNo.GetSelectedItem().GetColumnText(1));
                                                grid.PerformCallback('kosong');	                                
	                                            lblInfo.SetText('');	
                                            }" />
                                            <LoadingPanelStyle ImageSpacing="5px">
                                            </LoadingPanelStyle>
                                        </dx:ASPxComboBox>
                                    </td>
                                    <td align="left" valign="middle" style="height:25px; width:200px;">
                                        <dx:ASPxTextBox ID="txtFGNo" runat="server" Width="100%" Height="20px"
                                            ClientInstanceName="txtFGNo" Font-Names="Verdana"
                                            Font-Size="8pt" MaxLength="100" BackColor="#CCCCCC" ReadOnly="True">
                                        </dx:ASPxTextBox>
                                    </td>
                                    <td align="left" valign="middle" height="25px" width="150px"></td> 
                                </tr>
                                <tr>
                                    <td style="width:5px;"></td>
                                    <td align="left" valign="middle" height="25px" width="120px">
                                        <dx:ASPxLabel ID="ASPxLabel1" runat="server" Text="PART NO."
                                            Font-Names="Verdana" Font-Size="8pt">
                                        </dx:ASPxLabel>
                                    </td>
                                    <td align="left" valign="middle" height="25px" width="150px">       
                                        <dx:ASPxComboBox ID="cboPartNo" runat="server" 
                                            ClientInstanceName="cboPartNo" Width="100%"
                                            Font-Size="8pt" 
                                            Font-Names="Verdana" TextFormatString="{0}">
                                            <ClientSideEvents SelectedIndexChanged="function(s, e) {
	                                            txtPartNo.SetText(cboPartNo.GetSelectedItem().GetColumnText(1));
                                                grid.PerformCallback('kosong');	                                
	                                            lblInfo.SetText('');	
                                            }" />
                                            <LoadingPanelStyle ImageSpacing="5px">
                                            </LoadingPanelStyle>
                                        </dx:ASPxComboBox>  
                                    </td>
                                    <td align="left" valign="middle" style="height:25px; width:200px;">
                                        <dx:ASPxTextBox ID="txtPartNo" runat="server" Width="100%" Height="20px"
                                            ClientInstanceName="txtPartNo" Font-Names="Verdana"
                                            Font-Size="8pt" MaxLength="100" BackColor="#CCCCCC" ReadOnly="True">
                                        </dx:ASPxTextBox>
                                    </td>
                                    <td align="right" valign="middle" style="height:25px; width:100px;">
                                        <table style="width:100%;">
                                            <tr>
                                                <td align="right" valign="middle" style="height:25px; width:90px;">
                                                    <dx:ASPxButton ID="btnRefresh" runat="server" Text="SEARCH"
                                                        Font-Names="Verdana" Width="85px" AutoPostBack="False" Font-Size="8pt" >
                                                        <ClientSideEvents Click="function(s, e) {
                                                            grid.PerformCallback('load');
															grid.SetFocusedRowIndex(-1);
                                                            lblInfo.SetText('');
                                                        }" />
                                                    </dx:ASPxButton>
                                                </td>
                                                <td align="right" valign="middle" style="height:25px; width:90px;">
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
    </table>

    <div style="height:1px;"></div>

    <table style="width:100%; height: 15px;">
        <tr>
            <td colspan="8" height="15">
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
    </table>

    <div style="height:1px;"></div>

    <table style="width:100%;">
        <tr>
            <td colspan="8" align="left" valign="top" height="220">
                <%--Column : Grid--%>
                <dx:ASPxGridView ID="grid" runat="server" Width="100%"
                    Font-Names="Verdana" KeyFieldName="FinishGoodNo;PartNo"
                    AutoGenerateColumns="False"
                    ClientInstanceName="grid" Font-Size="8pt">
                    <ClientSideEvents Init="OnInit" FocusedRowChanged="function(s, e) {
	                    OnGridFocusedRowChanged();
                    }" EndCallback="function(s, e) {
    grid.CancelEdit();                
    var pMsg = s.cpMessage;        
    if (pMsg != '') {
        if (pMsg.substring(1,5) == '6011' || pMsg.substring(1,5) == '1002' || pMsg.substring(1,5) == '2001' || pMsg.substring(1,5) == '1001'  || pMsg.substring(1,5) == '1003') {
            lblInfo.GetMainElement().style.color = 'Blue';
        } else {
            lblInfo.GetMainElement().style.color = 'Red';
        }
        
        lblInfo.SetText(pMsg);
    } else {
        lblInfo.SetText('');
    }    

    AdjustSizeGrid();
    delete s.cpMessage;
}" RowClick="function(s, e) {
	                    lblInfo.SetText('');
                    }" BatchEditStartEditing="OnBatchEditStartEditing" />
                    
                    <Columns>
                        <dx:GridViewDataTextColumn VisibleIndex="0" Caption="NO" FieldName="NoUrut" Width="30px" HeaderStyle-HorizontalAlign="Center">
<HeaderStyle HorizontalAlign="Center"></HeaderStyle>
                        </dx:GridViewDataTextColumn>
                        <dx:GridViewDataTextColumn VisibleIndex="1" Caption="FINISH GOOD NO." FieldName="FinishGoodNo" Width="100px" HeaderStyle-HorizontalAlign="Center">
<HeaderStyle HorizontalAlign="Center"></HeaderStyle>
                        </dx:GridViewDataTextColumn>
                        <dx:GridViewDataTextColumn VisibleIndex="2" Caption="FINISH GOOD NAME" FieldName="FinishGoodName" Width="210px" HeaderStyle-HorizontalAlign="Center">
<HeaderStyle HorizontalAlign="Center"></HeaderStyle>
                        </dx:GridViewDataTextColumn>
                        <dx:GridViewDataTextColumn VisibleIndex="3" Caption="FG UOM" FieldName="FGUnitCls" Width="70px" HeaderStyle-HorizontalAlign="Center">
<HeaderStyle HorizontalAlign="Center"></HeaderStyle>
                        </dx:GridViewDataTextColumn>
                        <dx:GridViewDataTextColumn VisibleIndex="4" Caption="FG QTY" FieldName="FGQty" Width="100px" HeaderStyle-HorizontalAlign="Center">
<HeaderStyle HorizontalAlign="Center"></HeaderStyle>
                        </dx:GridViewDataTextColumn>
                        <dx:GridViewDataTextColumn VisibleIndex="5" Caption="PART NO." FieldName="PartNo" Width="100px" HeaderStyle-HorizontalAlign="Center">
<HeaderStyle HorizontalAlign="Center"></HeaderStyle>
                        </dx:GridViewDataTextColumn>
                        <dx:GridViewDataTextColumn VisibleIndex="6" Caption="PART NAME" FieldName="PartName" Width="210px" HeaderStyle-HorizontalAlign="Center">
<HeaderStyle HorizontalAlign="Center"></HeaderStyle>
                        </dx:GridViewDataTextColumn>
                        <dx:GridViewDataTextColumn VisibleIndex="7" Caption="PART UOM" FieldName="PartUnitCls" Width="70px" HeaderStyle-HorizontalAlign="Center">
<HeaderStyle HorizontalAlign="Center"></HeaderStyle>
                        </dx:GridViewDataTextColumn>
                        <dx:GridViewDataTextColumn VisibleIndex="8" Caption="PART QTY" FieldName="PartQty" Width="100px" HeaderStyle-HorizontalAlign="Center">
<HeaderStyle HorizontalAlign="Center"></HeaderStyle>
                        </dx:GridViewDataTextColumn>
                    </Columns>
                    <SettingsBehavior AllowSelectByRowClick="True" AllowSort="False"
                        ColumnResizeMode="Control" EnableRowHotTrack="True" />
                    <SettingsPager Visible="False" PageSize="13" Position="Top">
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
                        <SelectedRow ForeColor="Black" Wrap="False">
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

    <div style="height:8px;"></div>
    
    <table style="width:100%;">
        <tr>
            <td colspan="8" height="70">
                <!-- INPUT AREA -->
                <table id="tbl1" 
                style="border-width: 1pt thin thin thin; border-style: ridge; border-color:#9598A1; width:100%; height:25px;background-color: #FFD2A6">
                    <tr>
                        <td valign="top"                             
                            style="width: 110px;">
                            <dx:ASPxLabel ID="ASPxLabel53" runat="server" Text="FINISH GOOD NO." 
                                Font-Names="Verdana" Font-Size="8pt" Width="110px">
                            </dx:ASPxLabel>
                        </td>
                        <td valign="top" 
                            style="width: 200px;">
                            <dx:ASPxLabel ID="ASPxLabel65" runat="server" Text="FINISH GOOD NAME" 
                                Font-Names="Verdana" Font-Size="8pt" Width="200px">
                            </dx:ASPxLabel>
                        </td>
                        <td valign="top"
                            style="width: 90px;">
                            <dx:ASPxLabel ID="ASPxLabel56" runat="server" Text="FG UOM"
                                Font-Names="Verdana" Font-Size="8pt" Width="90px">
                            </dx:ASPxLabel>
                        </td>
                        <td valign="top"
                            style="width: 70px;">
                            <dx:ASPxLabel ID="ASPxLabel57" runat="server" Text="FG QTY"
                                Font-Names="Verdana" Font-Size="8pt" Width="70px">
                            </dx:ASPxLabel>
                        </td>
                        <td valign="top"
                            style="width: 130px;">
                            <dx:ASPxLabel ID="ASPxLabel55" runat="server" Text="PART NO."
                                Font-Names="Verdana" Font-Size="8pt" Width="130px">
                            </dx:ASPxLabel>
                        </td> 
                        <td valign="top"
                            style="width: 180px;">
                            <dx:ASPxLabel ID="ASPxLabel3" runat="server" Text="PART NAME"
                                Font-Names="Verdana" Font-Size="8pt" Width="180px">
                            </dx:ASPxLabel>
                        </td>   
                        <td valign="top"
                            style="width: 90px;">
                            <dx:ASPxLabel ID="ASPxLabel5" runat="server" Text="PART UOM"
                                Font-Names="Verdana" Font-Size="8pt" Width="90px">
                            </dx:ASPxLabel>
                        </td>
                        <td valign="top"
                            style="width: 70px;">
                            <dx:ASPxLabel ID="ASPxLabel2" runat="server" Text="PART QTY"
                                Font-Names="Verdana" Font-Size="8pt" Width="70px">
                            </dx:ASPxLabel>
                        </td> 
                    </tr>
                </table>

                <table id="tbl2" 
                style="border-width: 1pt thin thin thin; border-style: ridge; border-color:#9598A1; width:100%; height:35px;">
                    <tr>
                        <td style="width: 110px;">                            
                            <dx:ASPxComboBox ID="cboFGNo2" runat="server" 
                                ClientInstanceName="cboFGNo2" Width="110px"
                                Font-Size="8pt" 
                                Font-Names="Verdana" TextFormatString="{0}">
                                <ClientSideEvents SelectedIndexChanged="function(s, e) {
	                                txtFGNo2.SetText(cboFGNo2.GetSelectedItem().GetColumnText(1));
                                    txtFGUnit.SetText(cboFGNo2.GetSelectedItem().GetColumnText(2)); 
	                                lblInfo.SetText('');	
                                }" />
                                <LoadingPanelStyle ImageSpacing="5px">
                                </LoadingPanelStyle>
                            </dx:ASPxComboBox>
                        </td>
                        <td style="width: 200px;">
                            <dx:ASPxTextBox ID="txtFGNo2" runat="server" Width="200px" Height="20px"
                                ClientInstanceName="txtFGNo2"                                 
                                Font-Names="Verdana" 
                                Font-Size="8pt" MaxLength="50" BackColor="#CCCCCC" ReadOnly="True">
                            </dx:ASPxTextBox>
                        </td>
                        <td style="width: 90px;">
                            <dx:ASPxTextBox ID="txtFGUnit" runat="server" Width="90px" Height="20px"
                                ClientInstanceName="txtFGUnit"                                 
                                Font-Names="Verdana" 
                                Font-Size="8pt" MaxLength="50" BackColor="#CCCCCC" ReadOnly="True">
                            </dx:ASPxTextBox>
                        </td>
                        <td style="width: 70px;">
                            <dx:ASPxTextBox ID="txtFGQty" runat="server" Width="70px" Height="20px"
                                ClientInstanceName="txtFGQty"                                 
                                Font-Names="Verdana" 
                                Font-Size="8pt" MaxLength="16" onkeypress="return numbersonly(event)"
                                HorizontalAlign="Right">
                                <ClientSideEvents TextChanged="function(s, e) {
	                                lblInfo.SetText('');
                                }" />
                            </dx:ASPxTextBox>
                        </td>
                        <td style="width: 130px;">
                            <dx:ASPxComboBox ID="cboPartNo2" runat="server"
                                ClientInstanceName="cboPartNo2" Width="130px"
                                Font-Size="8pt"
                                Font-Names="Verdana" TextFormatString="{0}">
                                <ClientSideEvents SelectedIndexChanged="function(s, e) {
	                                txtPartNo2.SetText(cboPartNo2.GetSelectedItem().GetColumnText(1));
                                    txtPartUnit.SetText(cboPartNo2.GetSelectedItem().GetColumnText(2)); 
	                                lblInfo.SetText('');	
                                }" />
                                <LoadingPanelStyle ImageSpacing="5px">
                                </LoadingPanelStyle>
                            </dx:ASPxComboBox>
                        </td>
                        <td style="width: 180px;">
                            <dx:ASPxTextBox ID="txtPartNo2" runat="server" Width="180px" Height="20px"
                                ClientInstanceName="txtPartNo2"                                 
                                Font-Names="Verdana" 
                                Font-Size="8pt" MaxLength="50" BackColor="#CCCCCC" ReadOnly="True">
                            </dx:ASPxTextBox>
                        </td>
                        <td style="width: 90px;">
                            <dx:ASPxTextBox ID="txtPartUnit" runat="server" Width="90px" Height="20px"
                                ClientInstanceName="txtPartUnit"                                 
                                Font-Names="Verdana" 
                                Font-Size="8pt" MaxLength="50" BackColor="#CCCCCC" ReadOnly="True">
                            </dx:ASPxTextBox>
                        </td>
                        <td style="width: 70px;">
                            <dx:ASPxTextBox ID="txtPartQty" runat="server" Width="70px" Height="20px"
                                ClientInstanceName="txtPartQty"                                 
                                Font-Names="Verdana" 
                                Font-Size="8pt" MaxLength="16" onkeypress="return numbersonly(event)"
                                HorizontalAlign="Right">
                                <ClientSideEvents TextChanged="function(s, e) {
	                                lblInfo.SetText('');
                                }" />
                            </dx:ASPxTextBox>
                        </td>                        
                    </tr>
                </table>
            </td>
        </tr>
    </table> 

    <div style="height:8px;"></div>

    <table id="button" style=" width:100%;">
        <tr>
            <td valign="top" align="left">
                <dx:ASPxButton ID="btnSubMenu" runat="server" Text="SUB MENU"
                    Font-Names="Verdana" Width="85px" Font-Size="8pt">
                </dx:ASPxButton>
            </td>
            <td valign="top" align="right" style="width: 50px;" bgcolor="White">
                <dx:ASPxTextBox ID="txtModee" runat="server" ClientInstanceName="txtModee" 
                    Width="0px" BackColor="White" ForeColor="White">
                    <Border BorderColor="White" />
                </dx:ASPxTextBox>      
            </td>
            <td valign="top" align="right" style="width: 50px;">                                   
                <dx:ASPxButton ID="btnClear" runat="server" Text="CLEAR"                    
                    Font-Names="Verdana"
                    Width="90px" AutoPostBack="False" Font-Size="8pt">                    
                </dx:ASPxButton>
            </td>
            <td valign="top" align="right" style="width:80px;">                                   
                <dx:ASPxButton ID="btnDelete" runat="server" Text="DELETE"                              
                    Font-Names="Verdana" Width="80px" AutoPostBack="False" Font-Size="8pt">
                    <ClientSideEvents Click="function(s, e) {
                        up_delete();
                        grid.PerformCallback('loadaftersubmit');

                        cboFGNo2.SetText('');
                        txtFGNo2.SetText('');
                        txtFGUnit.SetText('');
                        txtFGQty.SetText('0');

                        cboPartNo2.SetText('');                   
                        txtPartNo2.SetText('');
                        txtPartUnit.SetText('');
                        txtPartQty.SetText('0');
                        txtModee.SetText('new');
						grid.SetFocusedRowIndex(-1);
                    
                        cboFGNo2.GetInputElement().setAttribute('style', 'background:#FFFFFF;');
                        cboFGNo2.GetInputElement().readOnly = false;
                        cboFGNo2.SetEnabled(true);

                        cboPartNo2.GetInputElement().setAttribute('style', 'background:#FFFFFF;');
                        cboPartNo2.GetInputElement().readOnly = false;
                        cboPartNo2.SetEnabled(true);

                    }" />
                </dx:ASPxButton>
            </td>
            <td valign="top" align="right" style="width: 50px;">                                   
                <dx:ASPxButton ID="btnSubmit" runat="server" Text="SAVE"                    
                    Font-Names="Verdana"
                    Width="90px" AutoPostBack="False" Font-Size="8pt">
                    <ClientSideEvents Click="function(s, e) {
                        validasubmit();
                        up_Insert();
                       
                        grid.PerformCallback('loadaftersubmit');
                        
                        cboFGNo2.SetText('');
                        txtFGNo2.SetText('');
                        txtFGUnit.SetText('');
                        txtFGQty.SetText('0');

                        cboPartNo2.SetText('');                   
                        txtPartNo2.SetText('');
                        txtPartUnit.SetText('');
                        txtPartQty.SetText('0');
                        
						grid.SetFocusedRowIndex(-1);

                        cboFGNo2.GetInputElement().setAttribute('style', 'background:#FFFFFF;');
                        cboFGNo2.GetInputElement().readOnly = false;
                        cboFGNo2.SetEnabled(true);

                        cboPartNo2.GetInputElement().setAttribute('style', 'background:#FFFFFF;');
                        cboPartNo2.GetInputElement().readOnly = false;
                        cboPartNo2.SetEnabled(true);
                        txtModee.SetText('new');
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
                        clear();
                    }
                }else if(s.cpFunction == 'insert'){
                    clear();
                }
            } else {
                lblInfo.SetText('');
            }  
        }" />
    </dx:ASPxCallback>
</asp:Content>

