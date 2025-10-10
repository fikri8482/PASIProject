<%@ Page Title="" Language="vb" AutoEventWireup="false" MasterPageFile="~/SiteContent.Master" CodeBehind="PartMasterDetail.aspx.vb" Inherits="PASISystem.PartMasterDetail" %>
<%@ Register Assembly="DevExpress.Web.v14.1, Version=14.1.6.0, Culture=neutral, PublicKeyToken=b88d1754d700e49a" Namespace="DevExpress.Web.ASPxCallback" TagPrefix="dx" %>
<%@ Register Assembly="DevExpress.Web.v14.1, Version=14.1.6.0, Culture=neutral, PublicKeyToken=b88d1754d700e49a" Namespace="DevExpress.Web.ASPxGridView.Export" TagPrefix="dx" %>
<%@ Register Assembly="DevExpress.Web.v14.1, Version=14.1.6.0, Culture=neutral, PublicKeyToken=b88d1754d700e49a" Namespace="DevExpress.Web.ASPxGridView" TagPrefix="dx" %>
<%@ Register assembly="DevExpress.Web.v14.1, Version=14.1.6.0, Culture=neutral, PublicKeyToken=b88d1754d700e49a" namespace="DevExpress.Web.ASPxEditors" tagprefix="dx" %>
<%@ Register Assembly="DevExpress.Web.v14.1, Version=14.1.6.0, Culture=neutral, PublicKeyToken=b88d1754d700e49a" Namespace="DevExpress.Web.ASPxPanel" TagPrefix="dx" %>
<%@ Register Assembly="DevExpress.Web.v14.1, Version=14.1.6.0, Culture=neutral, PublicKeyToken=b88d1754d700e49a" Namespace="DevExpress.Web.ASPxPopupControl" TagPrefix="dx" %>
<%@ Register Assembly="DevExpress.Web.v14.1, Version=14.1.6.0, Culture=neutral, PublicKeyToken=b88d1754d700e49a" namespace="DevExpress.Web.ASPxGlobalEvents" TagPrefix="dx" %>
<%@ Register assembly="DevExpress.Web.v14.1, Version=14.1.6.0, Culture=neutral, PublicKeyToken=b88d1754d700e49a" namespace="DevExpress.Web.ASPxHiddenField" tagprefix="dx1" %>
<asp:Content ID="Content1" ContentPlaceHolderID="head" runat="server">
<style type="text/css">
.dxeHLC, .dxeHC, .dxeHFC
{
display: none;
}
</style> 

<script language="javascript" type="text/javascript">
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

    function clear() {
        txtPartID.SetText('');
        txtPartName.SetText('');
        txtCarMakerCode.SetText('');
        txtCarMakerName.SetText('');
        txtPartNameGroup.SetText('');
        txtHSCode.SetText('');
        cboUnit.SetText('PCS');        
        txtMaker.SetText('');
        txtProject.SetText('');
        rdrYes.SetChecked(true);

        txtPartID.GetInputElement().setAttribute('style', 'background:#FFFFFF;foreground:#FFFFFF;');
        txtPartID.GetInputElement().readOnly = false;
    }

    function clear() {
        txtPartID.GetInputElement().setAttribute('style', 'background:#CCCCCC;foreground:#CCCCCC;');
        txtPartID.GetInputElement().readOnly = true;
    }

    function validasubmit() {
        lblInfo.GetMainElement().style.color = 'Red';
        if (txtPartID.GetText() == "") {
            lblInfo.SetText("[6011] Please Input Part No. first!");
            txtPartID.Focus();
            e.ProcessOnServer = false;
            return false;
        }        

        if (txtPartName.GetText() == "") {
            lblInfo.SetText("[6011] Please Input Part Name first!");
            txtPartName.Focus();
            e.ProcessOnServer = false;
            return false;
        }
        
        if (cboUnit.GetText() == "" ) {
            lblInfo.SetText("[6011] Please Select Unit first!");
            txtAddress.Focus();
            e.ProcessOnServer = false;
            return false;
        }              
    }

    function up_Insert() {
        var pIsUpdate = '';
        var pPartID = txtPartID.GetText();
        var pPartName = txtPartName.GetText();
        var pCarMakerCode = txtCarMakerCode.GetText();
        var pCarMakerName = txtCarMakerName.GetText();
        var pPartNameGroup = txtPartNameGroup.GetText();
        var pHSCode = txtHSCode.GetText();
        var pUOM = cboUnit.GetSelectedItem().GetColumnText(0);
        var pKanban = '';
        var pFG = '';
        var pMaker = txtMaker.GetText();
        var pProject = txtProject.GetText();
        
        if (rdrYes.GetValue() == true) {
            pKanban = '1';
        } else {
            pKanban = '0';
        }

//        if (rdrFG.GetValue() == true) {
//            pFG = '1';
//        } else {
//            pFG = '2';
//        }

        PartMasterSubmit.PerformCallback('save|' + pIsUpdate + '|' + pPartID + '|' + pPartName + '|' + pCarMakerCode + '|' + pCarMakerName + '|' + pPartNameGroup + '|' + pHSCode + '|' + pFG + '|' + pUOM + '|' + pMaker + '|' + pProject + '|' + pKanban);
    }

    function up_delete() {
        if (txtPartID.GetText() == "") {
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

        var pGroupCode = txtPartID.GetText();
        PartMasterSubmit.PerformCallback('delete|' + pGroupCode);
    }

    function readonly() {
        txtPartID.GetInputElement().setAttribute('style', 'background:#CCCCCC;foreground:#CCCCCC;');
        txtPartID.GetInputElement().readOnly = true;
        lblInfo.SetText('');
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
    <table style="border-width: thin; border-style: inset hidden ridge hidden; border-color:#9598A1; width:100%; height: 15px;">
        <tr>
            <td colspan="8" height="15">
                <%--error message--%>
                <table id="info" style="width:100%; height: 15px;">
                    <tr>
                        <td align="left" valign="middle" style="height:15px;">
                            <dx:ASPxLabel ID="lblInfo" runat="server" Text="[lblinfo]" Font-Names="Tahoma"
                                ClientInstanceName="lblInfo" Font-Bold="True" Font-Italic="True" Font-Size="8pt" >
                            </dx:ASPxLabel>
                        </td>
                    </tr>         
                </table>
            </td>            
        </tr>
    </table>

    <div style="height:5px;"></div> 

    <table style="border-width: 1pt thin thin thin; border-style: ridge; border-color:#9598A1; width:100%; height: 470px;">
        <tr>
            <td>
                <table style="width:100%;">
                    <tr>
                        <td align="left" width="50px">&nbsp;</td>
                        <td align="left" width="150px">
                            <dx:ASPxLabel ID="ASPxLabel10" runat="server" Font-Names="Tahoma"
                                Font-Size="8pt" Text="PART NO. YAZAKI">
                            </dx:ASPxLabel>
                        </td>
                        <td align="left">
                            <dx:ASPxTextBox ID="txtPartID" runat="server" Width="300px" 
                                ClientInstanceName="txtPartID" Font-Names="Tahoma"
                                MaxLength="25" onkeypress="return singlequote(event)" Height="25px">
                                <ClientSideEvents LostFocus="function(s, e) { lblInfo.SetText(&quot;&quot;); }" 
                                    KeyDown=" function(s, e) {
                                        if(ASPxClientUtils.GetKeyCode(e.htmlEvent) ===  ASPxKey.Enter){
                                            lblInfo.SetText('');
                                            PartMasterSubmit.PerformCallback('load');                                            
                                        }
                                    }" 
                                />
                            </dx:ASPxTextBox>
                        </td>
                    </tr>
                    <tr>
                        <td align="left" width="50px">&nbsp;</td>
                        <td align="left" width="150px">
                            <dx:ASPxLabel ID="ASPxLabel11" runat="server" Font-Names="Tahoma"
                                Font-Size="8pt" Text="PART NAME YAZAKI">
                            </dx:ASPxLabel>
                        </td>
                        <td align="left">
                            <dx:ASPxTextBox ID="txtPartName" runat="server" Width="300px" 
                                ClientInstanceName="txtPartName" Font-Names="Tahoma"
                                MaxLength="100" onkeypress="return singlequote(event)" Height="25px">
                                <ClientSideEvents LostFocus="function(s, e) { lblInfo.SetText(&quot;&quot;); }" />
                            </dx:ASPxTextBox>
                        </td>
                    </tr>
                    <tr>
                        <td align="left" width="50px">&nbsp;</td>
                        <td align="left" width="150px">
                            <dx:ASPxLabel ID="ASPxLabel8" runat="server" Font-Names="Tahoma"
                                Font-Size="8pt" Text="PART NO. CAR MAKER">
                            </dx:ASPxLabel>
                        </td>
                        <td align="left">
                            <dx:ASPxTextBox ID="txtCarMakerCode" runat="server" Width="300px" 
                                ClientInstanceName="txtCarMakerCode" Font-Names="Tahoma"
                                MaxLength="25" onkeypress="return singlequote(event)" Height="25px">
                                <ClientSideEvents LostFocus="function(s, e) { lblInfo.SetText(&quot;&quot;); }" />
                            </dx:ASPxTextBox>
                        </td>
                    </tr>
                    <tr>
                        <td align="left" width="50px">&nbsp;</td>
                        <td align="left" width="150px">
                            <dx:ASPxLabel ID="ASPxLabel9" runat="server" Font-Names="Tahoma"
                                Font-Size="8pt" Text="PART NAME CAR MAKER">
                            </dx:ASPxLabel>
                        </td>
                        <td align="left">
                            <dx:ASPxTextBox ID="txtCarMakerName" runat="server" Width="300px" 
                                ClientInstanceName="txtCarMakerName" Font-Names="Tahoma"
                                MaxLength="100" onkeypress="return singlequote(event)" Height="25px">
                                <ClientSideEvents LostFocus="function(s, e) { lblInfo.SetText(&quot;&quot;); }" />
                            </dx:ASPxTextBox>
                        </td>
                    </tr>
                    <tr>
                        <td align="left" width="50px">&nbsp;</td>
                        <td align="left" width="150px">
                            <dx:ASPxLabel ID="ASPxLabel19" runat="server" Font-Names="Tahoma"
                                Font-Size="8pt" Text="PART GROUP NAME">
                            </dx:ASPxLabel>
                        </td>
                        <td align="left">
                            <dx:ASPxTextBox ID="txtPartNameGroup" runat="server" Width="300px" 
                                ClientInstanceName="txtPartNameGroup" Font-Names="Tahoma"
                                MaxLength="100" onkeypress="return singlequote(event)" Height="25px">
                                <ClientSideEvents LostFocus="function(s, e) { lblInfo.SetText(&quot;&quot;); }" />
                            </dx:ASPxTextBox>
                        </td>
                    </tr>
                    <tr>
                        <td align="left" width="50px">&nbsp;</td>
                        <td align="left" width="150px">
                            <dx:ASPxLabel ID="ASPxLabel20" runat="server" Font-Names="Tahoma"
                                Font-Size="8pt" Text="HS CODE">
                            </dx:ASPxLabel>
                        </td>
                        <td align="left">
                            <dx:ASPxTextBox ID="txtHSCode" runat="server" Width="300px" 
                                ClientInstanceName="txtHSCode" Font-Names="Tahoma"
                                MaxLength="100" onkeypress="return singlequote(event)" Height="25px">
                                <ClientSideEvents LostFocus="function(s, e) { lblInfo.SetText(&quot;&quot;); }" />
                            </dx:ASPxTextBox>
                        </td>
                    </tr>                    
                    <tr>
                        <td align="left" width="50px">&nbsp;</td>
                        <td align="left" width="150px">
                            <dx:ASPxLabel ID="ASPxLabel12" runat="server" Font-Names="Tahoma"
                                Font-Size="8pt" Text="UOM">
                            </dx:ASPxLabel>
                        </td>
                        <td align="left">
                            <dx:ASPxComboBox ID="cboUnit" ClientInstanceName="cboUnit"  runat="server" TextFormatString="{1}" 
                                DropDownStyle="DropDown" Height="20px" Width="90px" MaxLength="1"
                                IncrementalFilteringMode="StartsWith" Font-Names="Tahoma"
                                Font-Size="8pt">
                                <ClientSideEvents LostFocus="function(s, e) { lblInfo.SetText(&quot;&quot;); }" />
                            </dx:ASPxComboBox>  
                        </td>
                    </tr>
                    <tr>
                        <td align="left" width="50px">&nbsp;</td>
                        <td align="left" width="150px">
                            <dx:ASPxLabel ID="ASPxLabel16" runat="server" Font-Names="Tahoma"
                                Font-Size="8pt" Text="KANBAN CLS">
                            </dx:ASPxLabel>
                        </td>
                        <td align="left">
                            <table>
                                <tr>
                                    <td>
                                        <dx:ASPxRadioButton ID="rdrYes" ClientInstanceName="rdrYes" runat="server" Text="YES" GroupName="pasi" Font-Names="Tahoma" Font-Size="8pt">
                                            <ClientSideEvents LostFocus="function(s, e) { lblInfo.SetText(&quot;&quot;); }" />
                                        </dx:ASPxRadioButton>
                                    </td>
                                    <td>
                                        <dx:ASPxRadioButton ID="rdrNo" ClientInstanceName="rdrNo" runat="server" Text="NO" GroupName="pasi" Font-Names="Tahoma" Font-Size="8pt">
                                            <ClientSideEvents LostFocus="function(s, e) { lblInfo.SetText(&quot;&quot;); }" />
                                        </dx:ASPxRadioButton>
                                    </td>
                                </tr>
                            </table> 
                        </td>
                    </tr>
                    <tr>
                        <td align="left" width="50px">&nbsp;</td>
                        <td align="left" width="150px">
                            <dx:ASPxLabel ID="ASPxLabel6" runat="server" Font-Names="Tahoma"
                                Font-Size="8pt" Text="MAKER">
                            </dx:ASPxLabel>
                        </td>
                        <td align="left">                            
                            <dx:ASPxTextBox ID="txtMaker" runat="server" Width="200px" 
                                ClientInstanceName="txtMaker" Font-Names="Tahoma"
                                MaxLength="20" onkeypress="return singlequote(event)" Height="25px">
                                <ClientSideEvents LostFocus="function(s, e) { lblInfo.SetText(&quot;&quot;); }" />
                            </dx:ASPxTextBox>
                        </td>                        
                    </tr>
                    <tr>
                        <td align="left" width="50px">&nbsp;</td>
                        <td align="left" width="150px">
                            <dx:ASPxLabel ID="ASPxLabel7" runat="server" Font-Names="Tahoma"
                                Font-Size="8pt" Text="PROJECT">
                            </dx:ASPxLabel>
                        </td>
                        <td align="left">                            
                            <dx:ASPxTextBox ID="txtProject" runat="server" Width="200px" 
                                ClientInstanceName="txtProject" Font-Names="Tahoma"
                                MaxLength="30" onkeypress="return singlequote(event)" Height="25px">
                                <ClientSideEvents LostFocus="function(s, e) { lblInfo.SetText(&quot;&quot;); }" />
                            </dx:ASPxTextBox>
                        </td>
                    </tr>
                </table>
            </td>
        </tr>

    </table> 
    <div style="height:8px;"></div>      

    <%--Button--%> 
    <table id="button" style=" width:100%;" onclick="return button_onclick()">
        <tr>                        
            <td valign="top" align="left">
                <dx:ASPxButton ID="btnSubMenu" runat="server" Text="SUB MENU"                     
                    Font-Names="Tahoma"
                    Width="90px" Font-Size="8pt" UseSubmitBehavior="False">
                </dx:ASPxButton>   
                </td>                     
            
            <td valign="top" align="right" style="width: 50px;">                                  
            </td>
            <td valign="top" align="right" style="width: 50px;">                                   
                <dx:ASPxButton ID="btnClear" runat="server" Text="CLEAR"                    
                    Font-Names="Tahoma"
                    Width="90px" AutoPostBack="False" Font-Size="8pt" 
                    UseSubmitBehavior="False" ClientInstanceName="btnClear">
                </dx:ASPxButton>
            </td>
            <td align="right" style="width:80px;">                                   
                <dx:ASPxButton ID="btnDelete" runat="server" Text="DELETE"                              
                    Font-Names="Tahoma" Width="80px" AutoPostBack="False" Font-Size="8pt" 
                    UseSubmitBehavior="False" ClientInstanceName="btnDelete">
                    <ClientSideEvents Click="function(s, e) {
                        up_delete();
                    }" />
                </dx:ASPxButton>
            </td>
            <td valign="top" align="right" style="width: 50px;">                                   
                <dx:ASPxButton ID="btnSubmit" runat="server" Text="SAVE"                    
                    Font-Names="Tahoma"
                    Width="90px" AutoPostBack="False" Font-Size="8pt" 
                    UseSubmitBehavior="False" ClientInstanceName="btnSubmit">
                    <ClientSideEvents Click="function(s, e) {
                        validasubmit();
                        up_Insert();
                    }" />
                </dx:ASPxButton>
            </td>
        </tr>
    </table>
    <div style="height:8px;"></div>  
    
    <dx:ASPxGlobalEvents ID="ge" runat="server" >
        <ClientSideEvents ControlsInitialized="function(s, e) {
	    OnControlsInitializedSplitter();
	    OnControlsInitializedGrid();
    }" />
    </dx:ASPxGlobalEvents>

    <dx:ASPxCallback ID="PartMasterSubmit" runat="server" 
        ClientInstanceName = "PartMasterSubmit">
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
                    clear2();
                }
            } else {
                lblInfo.SetText('');
            }  

            if (s.cpKeyPress == 'ON')
            {
                txtPartID.SetText(s.cpPartNo);
                txtPartName.SetText(s.cpPartName);
                txtCarMakerCode.SetText(s.cpCarMakerCode);
                txtCarMakerName.SetText(s.cpCarMakerName);
                txtPartNameGroup.SetText(s.cpPartNameGroup);
                txtHSCode.SetText(s.cpHSCode);
                cboUnit.SetText(s.cpUOM);
                txtMaker.SetText(s.cpMaker);
                txtProject.SetText(s.cpProject);
                          
                if (s.cpKanbanCls == '1') {
                    rdrYes.SetChecked(true);
                }else {
                    rdrNo.SetChecked(true);
                }

                txtPartID.GetInputElement().setAttribute('style', 'background:#CCCCCC;foreground:#CCCCCC;');
                txtPartID.GetInputElement().readOnly = true;

                delete s.cpKeyPress
            }
        }" />
    </dx:ASPxCallback>
</asp:Content>
