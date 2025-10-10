<%@ Page Title="" Language="vb" AutoEventWireup="false" MasterPageFile="~/SiteContent.Master"
    CodeBehind="SupplierCapacity.aspx.vb" Inherits="PASISystem.SupplierCapacity" %>

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
<%@ Register Assembly="DevExpress.Web.v14.1, Version=14.1.6.0, Culture=neutral, PublicKeyToken=b88d1754d700e49a" Namespace="DevExpress.Web.ASPxHiddenField" TagPrefix="dx" %>
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

            if (currentColumnName == "NoUrut" || currentColumnName == "SupplierID" || currentColumnName == "SupplierName"
            || currentColumnName == "PartNo" || currentColumnName == "PartName"
            || currentColumnName == "DailyCapacity" || currentColumnName == "MontlyCapacity") {
                e.cancel = true;
            }

            currentEditableVisibleIndex = e.visibleIndex;
        }

        function OnGridFocusedRowChanged() {
            grid.GetRowValues(grid.GetFocusedRowIndex(), "SupplierID;SupplierName;PartNo;PartName;DailyCapacity;MontlyCapacity;DeleteCls", OnGetRowValues);
        }
        function OnGetRowValues(values) {
            if (values[0] != "" && values[0] != null && values[0] != "null") {

                cboSupplierCode2.SetText(values[0]);
                txtSupplierCode2.SetText(values[1]);
                cboPartNo2.SetText(values[2]);
                txtPartNo2.SetText(values[3]);
                txtDailyCapacity.SetText(values[4]);
                txtMontlyCapacity.SetText(values[5]);
                var vDeleteCls = values[6];

                lblInfo.SetText('');
                cboSupplierCode2.GetInputElement().setAttribute('style', 'background:#CCCCCC;foreground:#CCCCCC;');
                cboSupplierCode2.GetInputElement().readOnly = true;
                cboSupplierCode2.SetEnabled(false);

                cboPartNo2.GetInputElement().setAttribute('style', 'background:#CCCCCC;foreground:#CCCCCC;');
                cboPartNo2.GetInputElement().readOnly = true;
                cboPartNo2.SetEnabled(false);
                
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
        

            var pPartNo = cboPartNo2.GetText();
            var pSupplierID = cboSupplierCode2.GetText();

            grid.PerformCallback('delete|' + pSupplierID + '|' + pPartNo);
        }

        function validasubmit() {
            lblInfo.GetMainElement().style.color = 'Red';
            if (cboPartNo2.GetText() == "") {
                lblInfo.SetText("[6011] Please Select Part No. first!");
                cboPartNo2.Focus();
                e.ProcessOnServer = false;
                return false;
            }

            if (cboSupplierCode2.GetText() == "") {
                lblInfo.SetText("[6011] Please Select Supplier Code first!");
                cboSupplierCode2.Focus();
                e.ProcessOnServer = false;
                return false;
            }
        }

        function up_Insert() {
            var pIsUpdate = '';
            lblInfo.SetText('4');
            var pPartID = cboPartNo2.GetText();
            lblInfo.SetText('5');
            var pSupplierID = cboSupplierCode2.GetText();
            lblInfo.SetText('6');
            var pDailyQty = txtDailyCapacity.GetValue();
            lblInfo.SetText('7');
            var pMontly = txtMontlyCapacity.GetValue();
            lblInfo.SetText('8');
            grid.PerformCallback('save|' + pIsUpdate + '|' + pSupplierID + '|' + pPartID + '|' + pDailyQty + '|' + pMontly);
            lblInfo.SetText('9');
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
    <table style="width: 100%;">
        <tr>
            <td>
                <table style="border-width: 1pt thin thin thin; border-style: ridge; border-color: #9598A1;
                    width: 100%;">
                    <tr>
                        <td colspan="8" height="30">
                            <table id="Table1">
                                <tr>
                                    <td style="width: 5px;">
                                    </td>
                                    <td align="left" valign="middle" height="25px" width="120px">
                                        <dx:aspxlabel id="ASPxLabel4" runat="server" text="SUPPLIER CODE" font-names="Verdana"
                                            font-size="8pt">
                                        </dx:aspxlabel>
                                    </td>
                                    <td align="left" valign="middle" height="25px" width="150px">
                                        <dx:aspxcombobox id="cboSupplierCode" runat="server" clientinstancename="cboSupplierCode"
                                            width="100%" font-size="8pt" font-names="Verdana" textformatstring="{0}">
                                            <ClientSideEvents SelectedIndexChanged="function(s, e) {
	                                            txtSupplierCode.SetText(cboSupplierCode.GetSelectedItem().GetColumnText(1));
                                                grid.PerformCallback('kosong');	                                
	                                            lblInfo.SetText('');	
                                            }" />
                                            <LoadingPanelStyle ImageSpacing="5px">
                                            </LoadingPanelStyle>
                                        </dx:aspxcombobox>
                                    </td>
                                    <td align="left" valign="middle" style="height: 25px; width: 200px;">
                                        <dx:aspxtextbox id="txtSupplierCode" runat="server" width="100%" height="20px" clientinstancename="txtSupplierCode"
                                            font-names="Verdana" font-size="8pt" maxlength="100" backcolor="#CCCCCC" readonly="True">
                                        </dx:aspxtextbox>
                                    </td>
                                    <td align="left" valign="middle" height="25px" width="150px">
                                    </td>
                                </tr>
                                <tr>
                                    <td style="width: 5px;">
                                    </td>
                                    <td align="left" valign="middle" height="25px" width="120px">
                                        <dx:aspxlabel id="ASPxLabel1" runat="server" text="PART NO." font-names="Verdana"
                                            font-size="8pt">
                                        </dx:aspxlabel>
                                    </td>
                                    <td align="left" valign="middle" height="25px" width="150px">
                                        <dx:aspxcombobox id="cboPartNo" runat="server" clientinstancename="cboPartNo" width="100%"
                                            font-size="8pt" font-names="Verdana" textformatstring="{0}">
                                            <ClientSideEvents SelectedIndexChanged="function(s, e) {
	                                            txtPartNo.SetText(cboPartNo.GetSelectedItem().GetColumnText(1));
                                                grid.PerformCallback('kosong');	                                
	                                            lblInfo.SetText('');	
                                            }" />
                                            <LoadingPanelStyle ImageSpacing="5px">
                                            </LoadingPanelStyle>
                                        </dx:aspxcombobox>
                                    </td>
                                    <td align="left" valign="middle" style="height: 25px; width: 200px;">
                                        <dx:aspxtextbox id="txtPartNo" runat="server" width="100%" height="20px" clientinstancename="txtPartNo"
                                            font-names="Verdana" font-size="8pt" maxlength="100" backcolor="#CCCCCC" readonly="True">
                                        </dx:aspxtextbox>
                                    </td>
                                    <td align="right" valign="middle" style="height: 25px; width: 100px;">
                                        <table style="width: 100%;">
                                            <tr>
                                                <td align="right" valign="middle" style="height: 25px; width: 90px;">
                                                    <dx:aspxbutton id="btnRefresh" runat="server" text="SEARCH" font-names="Verdana"
                                                        width="85px" autopostback="False" font-size="8pt">
                                                        <ClientSideEvents Click="function(s, e) {
                                                            grid.PerformCallback('load');
                                                            lblInfo.SetText('');
                                                        }" />
                                                    </dx:aspxbutton>
                                                </td>
                                                <td align="right" valign="middle" style="height: 25px; width: 90px;">
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
                            <dx:aspxlabel id="lblInfo" runat="server" text="[lblinfo]" font-names="Verdana" clientinstancename="lblInfo"
                                font-bold="True" font-italic="True" font-size="8pt">
                            </dx:aspxlabel>
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
            <td align="right" colspan="6">
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
        <tr>
            <td colspan="8" align="left" valign="top" height="220">
                <%--Column : Grid--%>
                <dx:aspxgridview id="grid" runat="server" width="100%" font-names="Verdana" keyfieldname="SupplierID;PartNo"
                    autogeneratecolumns="False" clientinstancename="grid" font-size="8pt">
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
                        <dx:GridViewDataTextColumn VisibleIndex="0" Caption="No" FieldName="NoUrut" Width="30px" HeaderStyle-HorizontalAlign="Center">
<HeaderStyle HorizontalAlign="Center"></HeaderStyle>
                        </dx:GridViewDataTextColumn>
                        <dx:GridViewDataTextColumn VisibleIndex="1" Caption="SUPPLIER CODE" FieldName="SupplierID" Width="90px" HeaderStyle-HorizontalAlign="Center">
<HeaderStyle HorizontalAlign="Center"></HeaderStyle>
                        </dx:GridViewDataTextColumn>
                        <dx:GridViewDataTextColumn VisibleIndex="2" Caption="SUPPLIER NAME" FieldName="SupplierName" Width="210px" HeaderStyle-HorizontalAlign="Center">
<HeaderStyle HorizontalAlign="Center"></HeaderStyle>
                        </dx:GridViewDataTextColumn>                        
                        <dx:GridViewDataTextColumn VisibleIndex="3" Caption="PART NO." FieldName="PartNo" Width="90px" HeaderStyle-HorizontalAlign="Center">
<HeaderStyle HorizontalAlign="Center"></HeaderStyle>
                        </dx:GridViewDataTextColumn>
                        <dx:GridViewDataTextColumn VisibleIndex="4" Caption="PART NAME" FieldName="PartName" Width="210px" HeaderStyle-HorizontalAlign="Center">
<HeaderStyle HorizontalAlign="Center"></HeaderStyle>
                        </dx:GridViewDataTextColumn>
                        <dx:GridViewDataTextColumn VisibleIndex="5" Caption="DAILY DELIVERY CAPACITY" FieldName="DailyCapacity" Width="120px" HeaderStyle-HorizontalAlign="Center">
<HeaderStyle HorizontalAlign="Center"></HeaderStyle>
                        </dx:GridViewDataTextColumn>
                        <dx:GridViewDataTextColumn VisibleIndex="6" Caption="MONTLY PRODUCTION CAPACITY" FieldName="MontlyCapacity" Width="120px" HeaderStyle-HorizontalAlign="Center">
<HeaderStyle HorizontalAlign="Center"></HeaderStyle>
                        </dx:GridViewDataTextColumn>
                        <dx:GridViewDataTextColumn VisibleIndex="7" Caption="REGISTER DATE" 
                            FieldName="EntryDate" Width="150px" HeaderStyle-HorizontalAlign="Center">
                        </dx:GridViewDataTextColumn>
                         <dx:GridViewDataTextColumn VisibleIndex="8" Caption="REGISTER USER" 
                            FieldName="EntryUser" Width="100px" HeaderStyle-HorizontalAlign="Center">
                        </dx:GridViewDataTextColumn>
                         <dx:GridViewDataTextColumn VisibleIndex="9" Caption="UPDATE DATE" 
                            FieldName="UpdateDate" Width="150px" HeaderStyle-HorizontalAlign="Center">
                        </dx:GridViewDataTextColumn>
                         <dx:GridViewDataTextColumn VisibleIndex="10" Caption="UPDATE USER" 
                            FieldName="UpdateUser" Width="100px" HeaderStyle-HorizontalAlign="Center">
                        </dx:GridViewDataTextColumn>
                        <dx:GridViewDataTextColumn VisibleIndex="11" Caption="DeleteCls" 
                            FieldName="DeleteCls" Width="0px" HeaderStyle-HorizontalAlign="Center">
                        </dx:GridViewDataTextColumn>
                    </Columns>
                    <SettingsBehavior AllowSelectByRowClick="True" AllowSort="False"
                        ColumnResizeMode="Control" EnableRowHotTrack="True" />

<SettingsBehavior AllowSort="False" AllowSelectByRowClick="True" ColumnResizeMode="Control" 
                        EnableRowHotTrack="True"></SettingsBehavior>

                    <SettingsPager Visible="False" PageSize="13" Position="Top">
                        <Summary Text="Page {0} of {1} [{2} record(s)]"
                                  AllPagesText="Page {0} of {1} " />
                        <Summary AllPagesText="Page {0} of {1} " Text="Page {0} of {1} [{2} record(s)]"></Summary>
                    </SettingsPager>
                    <SettingsEditing Mode="Batch" NewItemRowPosition="Bottom">
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
                    <styles>
                        <selectedrow forecolor="Black">
                        </selectedrow>
                    </styles>
                    <StylesEditors ButtonEditCellSpacing="0">
                        <ProgressBar Height="21px">
                        </ProgressBar>
                    </StylesEditors>
                </dx:aspxgridview>
            </td>
        </tr>
    </table>
    <div style="height: 8px;">
    </div>
    <table style="width: 100%;">
        <tr>
            <td colspan="8" height="70">
                <!-- INPUT AREA -->
                <table id="tbl1" style="border-width: 1pt thin thin thin; border-style: ridge; border-color: #9598A1;
                    width: 100%; height: 25px; background-color: #FFD2A6">
                    <tr>
                        <td valign="top" style="width: 130px;">
                            <dx:aspxlabel id="ASPxLabel53" runat="server" text="SUPPLIER CODE" font-names="Verdana"
                                font-size="8pt" width="120px">
                            </dx:aspxlabel>
                        </td>
                        <td valign="top" style="width: 230px;">
                            <dx:aspxlabel id="ASPxLabel65" runat="server" text="SUPPLIER NAME" font-names="Verdana"
                                font-size="8pt" width="230px">
                            </dx:aspxlabel>
                        </td>
                        <td valign="top" style="width: 130px;">
                            <dx:aspxlabel id="ASPxLabel55" runat="server" text="PART NO." font-names="Verdana"
                                font-size="8pt" width="120px">
                            </dx:aspxlabel>
                        </td>
                        <td valign="top" style="width: 230px;">
                            <dx:aspxlabel id="ASPxLabel3" runat="server" text="PART NAME" font-names="Verdana"
                                font-size="8pt" width="230px">
                            </dx:aspxlabel>
                        </td>
                        <td valign="top" style="width: 130px;">
                            <dx:aspxlabel id="ASPxLabel5" runat="server" text="DAILY DELIVERY CAPACITY" font-names="Verdana"
                                font-size="8pt" width="120px">
                            </dx:aspxlabel>
                        </td>
                        <td valign="top" style="width: 130px;">
                            <dx:aspxlabel id="ASPxLabel2" runat="server" text="MONTLY PRODUCTION CAPACITY" font-names="Verdana"
                                font-size="8pt" width="120px">
                            </dx:aspxlabel>
                        </td>
                        <td>
                        </td>
                    </tr>
                </table>
                <table id="tbl2" style="border-width: 1pt thin thin thin; border-style: ridge; border-color: #9598A1;
                    width: 100%; height: 35px;">
                    <tr>
                        <td style="width: 130px;">
                            <dx:aspxcombobox id="cboSupplierCode2" runat="server" clientinstancename="cboSupplierCode2"
                                width="120px" font-size="8pt" font-names="Verdana" textformatstring="{0}">
                                <ClientSideEvents SelectedIndexChanged="function(s, e) {
	                                txtSupplierCode2.SetText(cboSupplierCode2.GetSelectedItem().GetColumnText(1));
	                                lblInfo.SetText('');	
                                }" />
                                <LoadingPanelStyle ImageSpacing="5px">
                                </LoadingPanelStyle>
                            </dx:aspxcombobox>
                        </td>
                        <td style="width: 230px;">
                            <dx:aspxtextbox id="txtSupplierCode2" runat="server" width="230px" height="20px"
                                clientinstancename="txtSupplierCode2" font-names="Verdana" font-size="8pt" maxlength="50"
                                backcolor="#CCCCCC" readonly="True">
                            </dx:aspxtextbox>
                        </td>
                        <td style="width: 130px;">
                            <dx:aspxcombobox id="cboPartNo2" runat="server" clientinstancename="cboPartNo2" width="120px"
                                font-size="8pt" font-names="Verdana" textformatstring="{0}">
                                <ClientSideEvents SelectedIndexChanged="function(s, e) {
	                                txtPartNo2.SetText(cboPartNo2.GetSelectedItem().GetColumnText(1));
	                                lblInfo.SetText('');	
                                }" />
                                <LoadingPanelStyle ImageSpacing="5px">
                                </LoadingPanelStyle>
                            </dx:aspxcombobox>
                        </td>
                        <td style="width: 230px;">
                            <dx:aspxtextbox id="txtPartNo2" runat="server" width="230px" height="20px" clientinstancename="txtPartNo2"
                                font-names="Verdana" font-size="8pt" maxlength="50" backcolor="#CCCCCC" readonly="True">
                            </dx:aspxtextbox>
                        </td>
                        <td style="width: 130px;">
                            <dx:aspxtextbox id="txtDailyCapacity" runat="server" width="120px" height="20px"
                                clientinstancename="txtDailyCapacity" font-names="Verdana" font-size="8pt" maxlength="10"
                                onkeypress="return numbersonly(event)" horizontalalign="Right" 
                                displayformatstring="{0:n2}">
                                <ClientSideEvents TextChanged="function(s, e) {
	                                lblInfo.SetText('');
                                }" />
                            </dx:aspxtextbox>
                        </td>
                        <td style="width: 130px;">
                            <dx:aspxtextbox id="txtMontlyCapacity" runat="server" width="120px" height="20px"
                                clientinstancename="txtMontlyCapacity" font-names="Verdana" displayformatstring="{0:n2}"
                                font-size="8pt" maxlength="10" onkeypress="return numbersonly(event)" 
                                horizontalalign="Right">
                                <ClientSideEvents TextChanged="function(s, e) {
	                                lblInfo.SetText('');
                                }" />
                            </dx:aspxtextbox>
                        </td>
                        <td>
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
            <td valign="top" align="left">
                <dx:aspxbutton id="btnSubMenu" runat="server" text="SUB MENU" font-names="Verdana"
                    width="85px" font-size="8pt">
                </dx:aspxbutton>
            </td>
            <td valign="top" align="right" style="width: 50px;">
                <dx:aspxtextbox id="txtMode" runat="server" clientinstancename="txtMode" width="0px"
                    backcolor="White" forecolor="White">
                    <Border BorderColor="White" />
                </dx:aspxtextbox>
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
            <td valign="top" align="right" style="width: 50px;">
                <dx:aspxbutton id="btnClear" runat="server" text="CLEAR" font-names="Verdana" width="90px"
                    autopostback="False" font-size="8pt" ClientInstanceName="btnClear">                    
                </dx:aspxbutton>
            </td>
            <td valign="top" align="right" style="width: 80px;">
                <dx:aspxbutton id="btnDelete" runat="server" text="DELETE" font-names="Verdana" width="80px"
                    autopostback="False" font-size="8pt" ClientInstanceName="btnDelete">
                    <ClientSideEvents Click="function(s, e) {
                        up_delete();
                        grid.PerformCallback('loadaftersubmit');

                        cboSupplierCode2.SetText('');
                        txtSupplierCode2.SetText('');

                        cboPartNo2.SetText('');                   
                        txtPartNo2.SetText('');

                        txtDailyCapacity.SetText('0');
                        txtMontlyCapacity.SetText('0');
                                           
                        cboSupplierCode2.GetInputElement().setAttribute('style', 'background:#FFFFFF;');
                        cboSupplierCode2.GetInputElement().readOnly = false;
                        cboSupplierCode2.SetEnabled(true);

                        cboPartNo2.GetInputElement().setAttribute('style', 'background:#FFFFFF;');
                        cboPartNo2.GetInputElement().readOnly = false;
                        cboPartNo2.SetEnabled(true);

                    }" />
                </dx:aspxbutton>
            </td>
            <td valign="top" align="right" style="width: 50px;">
                <dx:aspxbutton id="btnSubmit" runat="server" text="SAVE" font-names="Verdana" width="90px"
                    autopostback="False" font-size="8pt" ClientInstanceName="btnSubmit">
                    <ClientSideEvents Click="function(s, e) {
                        lblInfo.SetText('1');
                        validasubmit();
                        lblInfo.SetText('2');
                        up_Insert();
                        lblInfo.SetText('3');
                        grid.PerformCallback('loadaftersubmit');

                        cboSupplierCode2.SetText('');
                        txtSupplierCode2.SetText('');

                        cboPartNo2.SetText('');                   
                        txtPartNo2.SetText('');

                        txtDailyCapacity.SetText('0');
                        txtMontlyCapacity.SetText('0');
                                           
                        cboSupplierCode2.GetInputElement().setAttribute('style', 'background:#FFFFFF;');
                        cboSupplierCode2.GetInputElement().readOnly = false;
                        cboSupplierCode2.SetEnabled(true);

                        cboPartNo2.GetInputElement().setAttribute('style', 'background:#FFFFFF;');
                        cboPartNo2.GetInputElement().readOnly = false;
                        cboPartNo2.SetEnabled(true);

                        }" />
                </dx:aspxbutton>
            </td>
        </tr>
    </table>
    <dx:aspxglobalevents id="ge" runat="server">
        <ClientSideEvents ControlsInitialized="function(s, e) {
	        OnControlsInitializedSplitter();
	        OnControlsInitializedGrid();
        }" />
    </dx:aspxglobalevents>
    <dx:aspxcallback id="AffiliateSubmit" runat="server" clientinstancename="AffiliateSubmit">
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
    </dx:aspxcallback>
    <dx:ASPxHiddenField ID="HF" runat="server" ClientInstanceName="HF">
    </dx:ASPxHiddenField>
</asp:Content>
