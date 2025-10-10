<%@ Page Title="" Language="vb" AutoEventWireup="false" MasterPageFile="~/SiteContent.Master" CodeBehind="Forwarder.aspx.vb" Inherits="PASISystem.Forwarder" %>
<%@ Register Assembly="DevExpress.Web.v14.1, Version=14.1.6.0, Culture=neutral, PublicKeyToken=b88d1754d700e49a"
    Namespace="DevExpress.Web.ASPxGridView.Export" TagPrefix="dx" %>
<%@ Register Assembly="DevExpress.Web.v14.1, Version=14.1.6.0, Culture=neutral, PublicKeyToken=b88d1754d700e49a"
    Namespace="DevExpress.Web.ASPxGridView" TagPrefix="dx" %>
<%@ Register Assembly="DevExpress.Web.v14.1, Version=14.1.6.0, Culture=neutral, PublicKeyToken=b88d1754d700e49a"
    Namespace="DevExpress.Web.ASPxEditors" TagPrefix="dx" %>
<%@ Register Assembly="DevExpress.Web.v14.1, Version=14.1.6.0, Culture=neutral, PublicKeyToken=b88d1754d700e49a"
    Namespace="DevExpress.Web.ASPxGlobalEvents" TagPrefix="dx" %>
<%@ Register Assembly="DevExpress.Web.v14.1, Version=14.1.6.0, Culture=neutral, PublicKeyToken=b88d1754d700e49a"
    Namespace="DevExpress.Web.ASPxCallback" TagPrefix="dx" %>
<asp:Content ID="Content1" ContentPlaceHolderID="head" runat="server">
    <style type="text/css">
        .dxeHLC, .dxeHC, .dxeHFC
        {
            display: none;
        }
        #Table1
        {
            width: 986px;
            margin-left: 0px;
        }
        .style1
        {
            width: 805px;
        }
    </style>

    <script language="javascript" type="text/javascript">
        function OnGridFocusedRowChanged() {
            grid.GetRowValues(grid.GetFocusedRowIndex(), 'ForwarderID;ForwarderName;Address;City;PostalCode;Phone1;Phone2;Fax;NPWP;DefaultCls;PORT', OnGetRowValues);
        }
        function OnGetRowValues(values) {
            if (values[0] != "" && values[0] != null && values[0] != "null") {
                txtForwarderCode.SetText(values[0]);
                txtForwarderName.SetText(values[1]);
                txtAddress.SetText(values[2]);
                txtCity.SetText(values[3]);
                txtPostalCode.SetText(values[4]);
                txtPhone1.SetText(values[5]);
                txtPhone2.SetText(values[6]);
                txtFax.SetText(values[7]);
                txtNPWP.SetText(values[8]);
                cboDefault.SetText(values[9]);
                txtPort.SetText(values[10]);
                
                txtForwarderCode.GetInputElement().setAttribute('style', 'background:#CCCCCC;');
                txtForwarderCode.GetInputElement().readOnly = true;

                lblInfo.SetText("");
            }

        }

        function singlequote(e) {
            var unicode = e.charCode ? e.charCode : e.keyCode
            if (unicode == 39) {
                return false //disable key press
            }
            if (unicode == 44) {
                return false //disable key press
            }
        }

    </script>

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
            height = height - (height * 52 / 100)
            grid.SetHeight(height);
        }

    </script>
</asp:Content>
<asp:Content ID="Content2" ContentPlaceHolderID="Content" runat="server">
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
                            <dx:ASPxLabel ID="lblInfo" runat="server" Font-Names="Tahoma" ClientInstanceName="lblInfo"
                                Font-Bold="True" Font-Italic="True" Font-Size="8pt">
                            </dx:ASPxLabel>
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
            <td colspan="2" align="left" valign="top" height="100px">
                <%--Column : Grid--%>
                <dx:ASPxGridView ID="grid" runat="server" Width="100%" Font-Names="Tahoma" KeyFieldName="ForwarderID"
                    AutoGenerateColumns="False" ClientInstanceName="grid" Font-Size="8pt">
                    <ClientSideEvents CallbackError="function(s, e) {e.handled = true;}" EndCallback="function(s, e) {
 
    if (s.cpMessage != '') {
        lblInfo.GetMainElement().style.color = 'Blue';
        lblInfo.SetText(s.cpMessage);
    } else {
        lblInfo.SetText('');
    }
    
}" Init="OnInit" RowClick="function(s, e) {
	lblInfo.SetText('');
}" FocusedRowChanged="function(s, e) { OnGridFocusedRowChanged(); lblInfo.SetText('');}" />
                    <Columns>
                        <dx:GridViewDataTextColumn VisibleIndex="0" Caption="NO" FieldName="RowNumber" Width="35px"
                            HeaderStyle-HorizontalAlign="Center">
                            <HeaderStyle HorizontalAlign="Center" Font-Names="Tahoma" Font-Size="8pt"></HeaderStyle>
                            <CellStyle Font-Names="Tahoma" Font-Size="8pt" HorizontalAlign="Left">
                            </CellStyle>
                        </dx:GridViewDataTextColumn>
                        <dx:GridViewDataTextColumn VisibleIndex="1" Caption="FORWARDER CODE" FieldName="ForwarderID" Width="120px"
                            HeaderStyle-HorizontalAlign="Center">
                            <HeaderStyle HorizontalAlign="Center" Font-Names="Tahoma" Font-Size="8pt"></HeaderStyle>
                            <CellStyle Font-Names="Tahoma" Font-Size="8pt">
                            </CellStyle>
                        </dx:GridViewDataTextColumn>
                        <dx:GridViewDataTextColumn VisibleIndex="2" Caption="FORWARDER NAME" FieldName="ForwarderName" Width="250px"
                            HeaderStyle-HorizontalAlign="Center">
                            <HeaderStyle HorizontalAlign="Center" Font-Names="Tahoma" Font-Size="8pt"></HeaderStyle>
                            <CellStyle Font-Names="Tahoma" Font-Size="8pt">
                            </CellStyle>
                        </dx:GridViewDataTextColumn>
                        <dx:GridViewDataTextColumn VisibleIndex="3" Caption="ADDRESS" FieldName="Address" Width="300px"
                            HeaderStyle-HorizontalAlign="Center">
                            <HeaderStyle HorizontalAlign="Center" Font-Names="Tahoma" Font-Size="8pt"></HeaderStyle>
                            <CellStyle Font-Names="Tahoma" Font-Size="8pt">
                            </CellStyle>
                        </dx:GridViewDataTextColumn>
                        <dx:GridViewDataTextColumn VisibleIndex="4" Caption="CITY" FieldName="City" Width="130px"
                            HeaderStyle-HorizontalAlign="Center">
                            <HeaderStyle HorizontalAlign="Center" Font-Names="Tahoma" Font-Size="8pt"></HeaderStyle>
                            <CellStyle Font-Names="Tahoma" Font-Size="8pt">
                            </CellStyle>
                        </dx:GridViewDataTextColumn>
                        <dx:GridViewDataTextColumn VisibleIndex="5" Caption="POSTAL CODE" FieldName="PostalCode" Width="100px"
                            HeaderStyle-HorizontalAlign="Center">
                            <HeaderStyle HorizontalAlign="Center" Font-Names="Tahoma" Font-Size="8pt"></HeaderStyle>
                            <CellStyle Font-Names="Tahoma" Font-Size="8pt">
                            </CellStyle>
                        </dx:GridViewDataTextColumn>
                        <dx:GridViewDataTextColumn VisibleIndex="6" Caption="PHONE 1" FieldName="Phone1" Width="150px"
                            HeaderStyle-HorizontalAlign="Center">
                            <HeaderStyle HorizontalAlign="Center" Font-Names="Tahoma" Font-Size="8pt"></HeaderStyle>
                            <CellStyle Font-Names="Tahoma" Font-Size="8pt">
                            </CellStyle>
                        </dx:GridViewDataTextColumn>
                        <dx:GridViewDataTextColumn VisibleIndex="7" Caption="PHONE 2" FieldName="Phone2" Width="150px"
                            HeaderStyle-HorizontalAlign="Center">
                            <HeaderStyle HorizontalAlign="Center" Font-Names="Tahoma" Font-Size="8pt"></HeaderStyle>
                            <CellStyle Font-Names="Tahoma" Font-Size="8pt">
                            </CellStyle>
                        </dx:GridViewDataTextColumn>
                        <dx:GridViewDataTextColumn VisibleIndex="8" Caption="FAX" FieldName="Fax" Width="150px"
                            HeaderStyle-HorizontalAlign="Center">
                            <HeaderStyle HorizontalAlign="Center" Font-Names="Tahoma" Font-Size="8pt"></HeaderStyle>
                            <CellStyle Font-Names="Tahoma" Font-Size="8pt">
                            </CellStyle>
                        </dx:GridViewDataTextColumn>
                        <dx:GridViewDataTextColumn VisibleIndex="9" Caption="NPWP" FieldName="NPWP" Width="150px"
                            HeaderStyle-HorizontalAlign="Center">
                            <HeaderStyle HorizontalAlign="Center" Font-Names="Tahoma" Font-Size="8pt"></HeaderStyle>
                            <CellStyle Font-Names="Tahoma" Font-Size="8pt">
                            </CellStyle>
                        </dx:GridViewDataTextColumn>
                        <dx:GridViewDataTextColumn VisibleIndex="10" Caption="PORT" FieldName="PORT" Width="150px"
                            HeaderStyle-HorizontalAlign="Center">
                            <HeaderStyle HorizontalAlign="Center" Font-Names="Tahoma" Font-Size="8pt"></HeaderStyle>
                            <CellStyle Font-Names="Tahoma" Font-Size="8pt">
                            </CellStyle>
                        </dx:GridViewDataTextColumn>
                        <dx:GridViewDataTextColumn VisibleIndex="11" Caption="DEFAULT CLS" FieldName="DefaultCls" Width="100px"
                            HeaderStyle-HorizontalAlign="Center">
                            <HeaderStyle HorizontalAlign="Center" Font-Names="Tahoma" Font-Size="8pt"></HeaderStyle>
                            <CellStyle Font-Names="Tahoma" Font-Size="8pt">
                            </CellStyle>
                        </dx:GridViewDataTextColumn>
                    </Columns>
                   <SettingsBehavior AllowSort="False" ColumnResizeMode="Control" 
                        EnableRowHotTrack="True" AllowFocusedRow="True" />
                    <SettingsPager Visible="False" Position="Top">
                        <Summary Text="Page {0} of {1} [{2} record(s)]" AllPagesText="Page {0} of {1} " />
                        <Summary AllPagesText="Page {0} of {1} " Text="Page {0} of {1} [{2} record(s)]">
                        </Summary>
                    </SettingsPager>
                    <Styles>
                        <Header BackColor="#FFD2A6" Font-Names="Verdana" Font-Size="8pt"></Header>
                        <Row BackColor="#FFFFE1" Font-Names="Verdana" Font-Size="8pt" Wrap="False"></Row>
                        <RowHotTrack BackColor="#E8EFFD" Font-Names="Verdana" Font-Size="8pt" Wrap="False"></RowHotTrack>
                        <SelectedRow Wrap="False">
                        </SelectedRow>
                    </Styles>
                   <Settings ShowGroupButtons="False" ShowVerticalScrollBar="True" 
                    ShowHorizontalScrollBar="True" 
                    ShowStatusBar="Hidden" VerticalScrollableHeight="210" /> 
                    <StylesEditors ButtonEditCellSpacing="0">
                        <ProgressBar Height="21px">
                        </ProgressBar>
                    </StylesEditors>
                </dx:ASPxGridView>
            </td>
        </tr>
    </table>
    <div style="height: 8px;">
    </div>
    <table style="width: 950px;">
        <tr>
            <td height="70" width="950">
                <!-- INPUT AREA -->
                <table id="Table1" style="border-width: 1pt thin thin thin; border-style: ridge;
                    border-color: #9598A1; width: 950px; height: 25px;">
                    <tr>
                    <td bgcolor="#FFD2A6" align="center" width="120px">
                        <dx:ASPxLabel ID="ASPxLabel56" runat="server" Text="FORWARDER CODE" Font-Names="Tahoma"
                            Font-Size="8pt" Width="120px">
                        </dx:ASPxLabel>
                    </td>
                    <td bgcolor="#FFD2A6" align="center" width="230px">
                        <dx:ASPxLabel ID="ASPxLabel5" runat="server" Text="FORWARDER NAME" Font-Names="Tahoma"
                            Font-Size="8pt" Width="230px">
                        </dx:ASPxLabel>
                    </td>
                    <td bgcolor="#FFD2A6" align="center" width="300px">
                        <dx:ASPxLabel ID="ASPxLabel57" runat="server" Text="ADDRESS" Font-Names="Tahoma"
                            Font-Size="8pt" Width="300px">
                        </dx:ASPxLabel>
                    </td>
                    <td bgcolor="#FFD2A6" align="center" width="150px">
                        <dx:ASPxLabel ID="ASPxLabel58" runat="server" Text="CITY" Font-Names="Tahoma" Font-Size="8pt"
                            Width="150px">
                        </dx:ASPxLabel>
                    </td>
                    <td bgcolor="#FFD2A6" align="center" width="150px">
                        <dx:ASPxLabel ID="ASPxLabel59" runat="server" Text="POSTAL CODE" Font-Names="Tahoma"
                            Font-Size="8pt" Width="150px">
                        </dx:ASPxLabel>
                    </td>
                    <td bgcolor="#FFD2A6" align="center" width="150px">                        
                    </td>
                    </tr>
                    <tr>
                        <td align="left" width="120px">
                            <dx:ASPxTextBox ID="txtForwarderCode" runat="server" Width="120px" 
                                Height="20px" ClientInstanceName="txtForwarderCode"
                                Font-Names="Tahoma" Font-Size="8pt" MaxLength="20" BackColor="White" onkeypress="return singlequote(event)">
                                <ClientSideEvents LostFocus="function(s, e) {
	lblInfo.SetText('');
}" />
                            </dx:ASPxTextBox>
                        </td>
                        <td align="left" width="230px">
                            <dx:ASPxTextBox ID="txtForwarderName" runat="server" Width="230px" Height="20px"
                                ClientInstanceName="txtForwarderName" Font-Names="Tahoma" 
                                Font-Size="8pt" MaxLength="100"
                                BackColor="White" TabIndex="1" onkeypress="return singlequote(event)">
                                <ClientSideEvents LostFocus="function(s, e) {
	lblInfo.SetText('');
}" />
                            </dx:ASPxTextBox>
                        </td>
                        <td align="left" width="300px">
                            <dx:ASPxTextBox ID="txtAddress" runat="server" Width="300px" Height="20px" ClientInstanceName="txtAddress"
                                Font-Names="Tahoma" Font-Size="8pt" MaxLength="150" BackColor="White" 
                                TabIndex="2" onkeypress="return singlequote(event)">
                                <ClientSideEvents LostFocus="function(s, e) {
	lblInfo.SetText('');
}" />
                            </dx:ASPxTextBox>
                        </td>
                        <td align="left" width="150px">
                            <dx:ASPxTextBox ID="txtCity" runat="server" Width="150px" Height="20px" ClientInstanceName="txtCity"
                                Font-Names="Tahoma" Font-Size="8pt" MaxLength="20" BackColor="White" onkeypress="return singlequote(event)" 
                                TabIndex="3">
                                <ClientSideEvents LostFocus="function(s, e) {
	lblInfo.SetText('');
}" />
                            </dx:ASPxTextBox>
                        </td>
                        <td align="left" width="150px">
                            <dx:ASPxTextBox ID="txtPostalCode" runat="server" Width="150px" Height="20px" ClientInstanceName="txtPostalCode"
                                onkeypress="return numbersonly(event)" Font-Names="Tahoma" Font-Size="8pt" MaxLength="15"
                                BackColor="White" TabIndex="4">
                                <ClientSideEvents LostFocus="function(s, e) {
	lblInfo.SetText('');
}" />
                            </dx:ASPxTextBox>
                        </td>
                        <td align="left" width="150px">                            
                        </td>
                    </tr>
                </table>
                </td>
                </tr>
        <tr>
            <td height="70" width="950">
                <!-- INPUT AREA -->
                <table id="Table2" style="border-width: 1pt thin thin thin; border-style: ridge;
                    border-color: #9598A1; width: 950px; height: 25px;">
                    <tr>
                        <td bgcolor="#FFD2A6" width="190px" align="center">
                            <dx:ASPxLabel ID="ASPxLabel1" runat="server" Text="PHONE 1" Font-Names="Tahoma" Font-Size="8pt"
                                Width="190px">
                            </dx:ASPxLabel>
                        </td>
                        <td bgcolor="#FFD2A6" width="190px" align="center">
                            <dx:ASPxLabel ID="ASPxLabel2" runat="server" Text="PHONE 2" Font-Names="Tahoma" Font-Size="8pt"
                                Width="190px">
                            </dx:ASPxLabel>
                        </td>
                        <td bgcolor="#FFD2A6" width="225px" align="center">
                            <dx:ASPxLabel ID="ASPxLabel3" runat="server" Text="FAX" Font-Names="Tahoma"
                                Font-Size="8pt" Width="225px">
                            </dx:ASPxLabel>
                        </td>
                        <td bgcolor="#FFD2A6" width="225px" align="center">
                            <dx:ASPxLabel ID="ASPxLabel4" runat="server" Text="NPWP" Font-Names="Tahoma" Font-Size="8pt"
                                Width="225px" Height="16px">
                            </dx:ASPxLabel>
                        </td>
                        <td bgcolor="#FFD2A6" width="225px" align="center">
                            <dx:ASPxLabel ID="ASPxLabel7" runat="server" Text="PORT" Font-Names="Tahoma" Font-Size="8pt"
                                Width="225px" Height="16px">
                            </dx:ASPxLabel>
                        </td>
                        <td bgcolor="#FFD2A6" width="100px" align="center">
                            <dx:ASPxLabel ID="ASPxLabel6" runat="server" Text="DEFAULT CLS" Font-Names="Tahoma"
                                Font-Size="8pt" Width="100px">
                            </dx:ASPxLabel>
                        </td>
                    </tr>
                    <tr>
                        <td width="200px">
                            <dx:ASPxTextBox ID="txtPhone1" runat="server" Width="200px" Height="20px" ClientInstanceName="txtPhone1"
                                Font-Names="Tahoma" onkeypress="return numbersonly(event)" Font-Size="8pt" MaxLength="20"
                                BackColor="White" TabIndex="5">
                                <ClientSideEvents LostFocus="function(s, e) {
	lblInfo.SetText('');
}" />
                            </dx:ASPxTextBox>
                        </td>
                        <td width="200px">
                            <dx:ASPxTextBox ID="txtPhone2" runat="server" Width="200px" Height="20px" ClientInstanceName="txtPhone2"
                                Font-Names="Tahoma" onkeypress="return numbersonly(event)" Font-Size="8pt" MaxLength="20"
                                BackColor="White" TabIndex="6">
                                <ClientSideEvents LostFocus="function(s, e) {
	lblInfo.SetText('');
}" />
                            </dx:ASPxTextBox>
                        </td>
                        <td width="225px">
                            <dx:ASPxTextBox ID="txtFax" runat="server" Width="225px" Height="20px" ClientInstanceName="txtFax"
                                Font-Names="Tahoma" onkeypress="return numbersonly(event)" Font-Size="8pt" MaxLength="20"
                                BackColor="White" TabIndex="7">
                                <ClientSideEvents LostFocus="function(s, e) {
	lblInfo.SetText('');
}" />
                            </dx:ASPxTextBox>
                        </td>
                        <td width="225px">
                            <dx:ASPxTextBox ID="txtNPWP" runat="server" Width="225px" Height="20px" ClientInstanceName="txtNPWP"
                                Font-Names="Tahoma" onkeypress="return numbersonly(event)" Font-Size="8pt" MaxLength="25"
                                BackColor="White" TabIndex="8">
                                <ClientSideEvents LostFocus="function(s, e) {
	lblInfo.SetText('');
}" />
                            </dx:ASPxTextBox>
                        </td>
                        <td width="225px">
                            <dx:ASPxTextBox ID="txtPort" runat="server" Width="225px" Height="20px" ClientInstanceName="txtPort"
                                Font-Names="Tahoma" Font-Size="8pt" MaxLength="50"
                                BackColor="White" TabIndex="8">
                                <ClientSideEvents LostFocus="function(s, e) {
	lblInfo.SetText('');
}" />
                            </dx:ASPxTextBox>
                        </td>
                        <td width="100px">
                            <dx:ASPxComboBox ID="cboDefault" runat="server" ClientInstanceName="cboDefault" 
                                Width="100px" TabIndex="9" onkeypress="return singlequote(event)">
                                <ClientSideEvents DropDown="function(s, e) {
	lblInfo.SetText('');
}" LostFocus="function(s, e) {
	lblInfo.SetText('');
}" />
                                <Items>
                                    <dx:ListEditItem Text="YES" Value="1" />
                                    <dx:ListEditItem Text="NO" Value="0" />
                                </Items>
                            </dx:ASPxComboBox>
                        </td>
                    </tr>
                </table>
            </td>
        </tr>
    </table>
    <div style="height: 8px;">
    </div>
    <table id="button" style="width: 100%;">
        <tr>
            <td valign="top" align="left" class="style1">
                <dx:ASPxButton ID="btnSubMenu" runat="server" Text="SUB MENU" Font-Names="Tahoma"
                    Width="85px" Font-Size="8pt" TabIndex="13" ClientInstanceName="btnSubMenu">
                </dx:ASPxButton>
            </td>
            <td valign="top" align="left">
               <%-- <dx:ASPxTextBox ID="txtMode" runat="server" ClientInstanceName="txtMode" 
                    Width="0px" BackColor="White" ForeColor="White">
                    <Border BorderColor="White" />
                </dx:ASPxTextBox>--%>
            </td>
            <td valign="top" align="right" style="width: 50px;">
                <dx:ASPxButton ID="btnClear" runat="server" Text="CLEAR" Font-Names="Tahoma" Width="90px"
                    AutoPostBack="False" Font-Size="8pt" TabIndex="12" 
                    ClientInstanceName="btnClear">
                    <ClientSideEvents Click="function(s, e) {
	txtForwarderCode.SetText('');
        txtForwarderName.SetText('');
        txtAddress.SetText('');
        txtCity.SetText('');
        txtPostalCode.SetText('');
        txtPhone1.SetText('');
        txtPhone2.SetText('');
        txtFax.SetText('');
        txtNPWP.SetText('');
        cboDefault.SetText('');
	grid.PerformCallback('');

    	grid.SetFocusedRowIndex(-1);
        
    lblInfo.SetText('');

	txtForwarderCode.GetInputElement().setAttribute('style', 'background:#FFFFFF;');
    txtForwarderCode.GetInputElement().readOnly = false;
}" />
                </dx:ASPxButton>
            </td>
            <td align="right" style="width: 80px;">
                <dx:ASPxButton ID="btnDelete" runat="server" Text="DELETE" Font-Names="Tahoma" Width="80px"
                    AutoPostBack="False" Font-Size="8pt" TabIndex="11" 
                    ClientInstanceName="btnDelete">
                    <ClientSideEvents Click="function(s, e) {
    if (grid.GetFocusedRowIndex() == -1) {
        lblInfo.GetMainElement().style.color = 'Red';
		lblInfo.SetText('[6010] Please select the data first!');
        e.processOnServer = false;
        return;
    } 
    
    var msg = confirm('Are you sure want to delete this data ?');                
    if (msg == false) {
        e.processOnServer = false;
        return;
    } 

var pForwarderCode = txtForwarderCode.GetText();
var pForwarderName = txtForwarderName.GetText(); 
var pAddress = txtAddress.GetText();
var pCity = txtCity.GetText();
var pPostalCode = txtPostalCode.GetText();
var pPhone1 = txtPhone1.GetText();
var pPhone2 = txtPhone2.GetText();
var pFax = txtFax.GetText();
var pNPWP = txtNPWP.GetText();
var pDefaultCls = cboDefault.GetText();

grid.PerformCallback('delete|'+pForwarderCode);     
        
    	txtForwarderCode.SetText('');
        txtForwarderName.SetText('');
        txtAddress.SetText('');
        txtCity.SetText('');
        txtPostalCode.SetText('');
        txtPhone1.SetText('');
        txtPhone2.SetText('');
        txtFax.SetText('');
        txtNPWP.SetText('');
        cboDefault.SetText('');
        txtMode.SetText('new');
    	grid.SetFocusedRowIndex(-1);
                
    txtForwarderCode.GetInputElement().setAttribute('style', 'background:#FFFFFF;');
    txtForwarderCode.GetInputElement().readOnly = false;        
    
}" />
                </dx:ASPxButton>
            </td>
            <td valign="top" align="right" style="width: 50px;">
                <dx:ASPxButton ID="btnSubmit" runat="server" Text="SAVE" Font-Names="Tahoma" Width="90px"
                    AutoPostBack="False" Font-Size="8pt" TabIndex="10" 
                    ClientInstanceName="btnSubmit">
                    <ClientSideEvents Click="function(s, e) {
if (txtForwarderCode.GetText() == ''){
        lblInfo.GetMainElement().style.color = 'Red';
		lblInfo.SetText('[6012] Please input Forwarder Code first!');
        txtForwarderCode.Focus();
		e.processOnServer = false;
        return;
	}

if (txtForwarderName.GetText() == ''){
        lblInfo.GetMainElement().style.color = 'Red';
        lblInfo.SetText('[6012] Please input Forwarder Name first!');
        txtForwarderName.Focus();
                e.ProcessOnServer = false;
                return;
    }

if (txtAddress.GetText() == ''){
        lblInfo.GetMainElement().style.color = 'Red';
		lblInfo.SetText('[6012] Please input Address first!');
        txtAddress.Focus();
		e.processOnServer = false;
        return;
	}

var pIsUpdate = '';
var pForwarderCode = txtForwarderCode.GetText();
var pForwarderName = txtForwarderName.GetText(); 
var pAddress = txtAddress.GetText();
var pCity = txtCity.GetText();
var pPostalCode = txtPostalCode.GetText();
var pPhone1 = txtPhone1.GetText();
var pPhone2 = txtPhone2.GetText();
var pFax = txtFax.GetText();
var pNPWP = txtNPWP.GetText();
var pDefaultCls = cboDefault.GetValue();

    if (grid.GetFocusedRowIndex() == -1) {
        pIsUpdate = 'new';
        grid.PerformCallback('save|'+pIsUpdate+'|'+pForwarderCode+'|'+pForwarderName+'|'+pAddress+'|'+pCity+'|'+pPostalCode+'|'+pPhone1+'|'+pPhone2+'|'+pFax+'|'+pNPWP+'|'+pDefaultCls);  

    } else {
        pIsUpdate = 'edit';        
        grid.PerformCallback('save|'+pIsUpdate+'|'+pForwarderCode+'|'+pForwarderName+'|'+pAddress+'|'+pCity+'|'+pPostalCode+'|'+pPhone1+'|'+pPhone2+'|'+pFax+'|'+pNPWP+'|'+pDefaultCls); 
    	grid.SetFocusedRowIndex(-1);
}
    
    	txtForwarderCode.SetText('');
        txtForwarderName.SetText('');
        txtAddress.SetText('');
        txtCity.SetText('');
        txtPostalCode.SetText('');
        txtPhone1.SetText('');
        txtPhone2.SetText('');
        txtFax.SetText('');
        txtNPWP.SetText('');
        cboDefault.SetText('');
        txtMode.SetText('new');
    	grid.SetFocusedRowIndex(-1);
        
    txtForwarderCode.GetInputElement().setAttribute('style', 'background:#FFFFFF;');
    txtForwarderCode.GetInputElement().readOnly = false;
}" />
                </dx:ASPxButton>
            </td>
        </tr>
    </table>
    <dx:ASPxGlobalEvents ID="ge" runat="server">
        <ClientSideEvents ControlsInitialized="function(s, e) {
	        OnControlsInitializedSplitter();
	        OnControlsInitializedGrid();
        }" />
    </dx:ASPxGlobalEvents>
</asp:Content>
