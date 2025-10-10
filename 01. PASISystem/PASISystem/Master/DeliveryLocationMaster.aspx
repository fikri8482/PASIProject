<%@ Page Title="" Language="vb" AutoEventWireup="false" MasterPageFile="~/SiteContent.Master"
    CodeBehind="DeliveryLocationMaster.aspx.vb" Inherits="PASISystem.DeliveryLocationMaster" %>

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

            if (currentColumnName == "NoUrut" || currentColumnName == "AffiliateID" || currentColumnName == "DeliveryLocationCode") {

                e.cancel = true;
            }

            currentEditableVisibleIndex = e.visibleIndex;
        }

        function OnGridFocusedRowChanged() {
            grid.GetRowValues(grid.GetFocusedRowIndex(), "AffiliateID;DeliveryLocationCode;DeliveryLocationName;Address;City;PostalCode;Phone1;Phone2;Fax;NPWP;PODeliveryBy;DefaultCls;", OnGetRowValues);
        }

        function OnGetRowValues(values) {
            if (values[0] != "" && values[0] != null && values[0] != "null") {

                cboAffiliateCode.SetText(values[0]);
                txtDeliveryLoc.SetText(values[1]);
                txtDeliveryLocName.SetText(values[2]);
                txtAddress.SetText(values[3]);
                txtCity.SetText(values[4]);
                txtPostalCode.SetText(values[5]);
                txtPhone1.SetText(values[6]);
                txtPhone2.SetText(values[7]);
                txtFax.SetText(values[8]);
                txtNPWP.SetText(values[9]);
                cboPODelby.SetText(values[10]);
                cboDefault.SetText(values[11]);               

                cboAffiliateCode.GetInputElement().setAttribute('style', 'background:#CCCCCC;foreground:#CCCCCC;');
                cboAffiliateCode.GetInputElement().readOnly = true;
                cboAffiliateCode.SetEnabled(false);

                txtDeliveryLoc.GetInputElement().setAttribute('style', 'background:#CCCCCC;foreground:#CCCCCC;');
                txtDeliveryLoc.GetInputElement().readOnly = true;
                txtDeliveryLoc.SetEnabled(false);

            }
        }

        function up_delete() {
            if (cboAffiliateCode.GetText() == "") {
                lblInfo.GetMainElement().style.color = 'Red';
                lblInfo.SetText("[6011] Please Select the data first!");
                e.ProcessOnServer = false;
                return false;
            }

            if (txtDeliveryLoc.GetText() == "") {
                lblInfo.GetMainElement().style.color = 'Red';
                lblInfo.SetText("[6011] Please Select the data first!");
                e.ProcessOnServer = false;
                return false;
            }

            if (grid.GetFocusedRowIndex() == -1) {
                lblInfo.GetMainElement().style.color = 'Red';
                lblInfo.SetText("[6011] Please Select the data first!");
                e.ProcessOnServer = false;
                return false;
            }

            var msg = confirm('Are you sure want to delete this data ?');
            if (msg == false) {
                e.processOnServer = false;
                return;
            }
            var pAffiliateID = cboAffiliateCode.GetText();
            var pDeliveryLoc = txtDeliveryLoc.GetText();
            grid.PerformCallback('delete|' + pAffiliateID + ' |' + pDeliveryLoc);

        }

        function readonly() {
            cboAffiliateCode.GetInputElement().setAttribute('style', 'background:#CCCCCC;foreground:#CCCCCC;');
            cboAffiliateCode.GetInputElement().readOnly = true;
            lblInfo.SetText('');

            txtDeliveryLoc.GetInputElement().setAttribute('style', 'background:#CCCCCC;foreground:#CCCCCC;');
            txtDeliveryLoc.GetInputElement().readOnly = true;
            lblInfo.SetText('');
        }

        function validasubmit() {
            lblInfo.GetMainElement().style.color = 'Red';
            if (cboAffiliateCode.GetText() == "") {
                lblInfo.SetText("[6011] Please Input Affiliate Code first!");
                cboAffiliateCode.Focus();
                e.ProcessOnServer = false;
                return false;
            }


            lblInfo.GetMainElement().style.color = 'Red';
            if (txtDeliveryLoc.GetText() == "") {
                lblInfo.SetText("[6011] Please Input the Delivery Location first!");
                txtDeliveryLoc.Focus();
                e.ProcessOnServer = false;
                return false;
            }
        }

        function up_Insert() {
            var pIsUpdate = '';

            var pAffiliateID = cboAffiliateCode.GetText();
            var pDeliveryLoc = txtDeliveryLoc.GetText();
            var pDeliveryLocName = txtDeliveryLocName.GetText();
            var pAddress = txtAddress.GetText();
            var pCity = txtCity.GetText();
            var pPostalCode = txtPostalCode.GetText();
            var pPhone1 = txtPhone1.GetText();
            var pPhone2 = txtPhone2.GetText();
            var pFax = txtFax.GetText();
            var pNPWP = txtNPWP.GetText();
            var pPODeliveryby = cboPODelby.GetValue();
            var pDefaultCls = cboDefault.GetValue();

            grid.PerformCallback('save|' + pIsUpdate + '|' + pAffiliateID + '|' + pDeliveryLoc + '|' + pDeliveryLocName + '|' + pAddress + '|' + pCity + '|' + pPostalCode + '|' + pPhone1 + '|' + pPhone2 + '|' + pFax + '|' + pNPWP + '|' + pPODeliveryby + '|' + pDefaultCls);

            
            grid.PerformCallback('load');


            cboAffiliateCode.SetText('');
            txtDeliveryLoc.SetText('');
            txtDeliveryLocName.SetText('');
            txtAddress.SetText('');
            txtCity.SetText('');
            txtPostalCode.SetText('');
            txtPhone1.SetText('');
            txtPhone2.SetText('');
            txtFax.SetText('');
            txtNPWP.SetText('');
            cboPODelby.SetText('');
            cboDefault.SetText('');
            lblInfo.SetText('');

//            cboAffiliateCode.GetInputElement().readOnly = false;
//            cboAffiliateCode.SetEnabled(true);

//            txtDeliveryLoc.GetInputElement().readOnly = false;
//            txtDeliveryLoc.SetEnabled(true);
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
        <tr>
            <td colspan="2" align="left" valign="top" height="100px">
                <%--Column : Grid--%>
                <dx:ASPxGridView ID="grid" runat="server" Width="100%" Font-Names="Tahoma" KeyFieldName="AffiliateID;DeliveryLocationCode"
                    AutoGenerateColumns="False" ClientInstanceName="grid" Font-Size="8pt" 
                    ForeColor="Black">
                    <ClientSideEvents Init="OnInit" FocusedRowChanged="function(s, e) {
	                    OnGridFocusedRowChanged();
                    }" EndCallback="function(s, e) {
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
						delete s.cpMessage;
                         
                    }" RowClick="function(s, e) {
	                    lblInfo.SetText('');
                       
                    }" BatchEditStartEditing="OnBatchEditStartEditing" />
                    <Columns>
                        <dx:GridViewDataTextColumn VisibleIndex="0" Caption="NO" FieldName="NoUrut" Width="30px"
                            HeaderStyle-HorizontalAlign="Center">
                            <HeaderStyle HorizontalAlign="Center" Font-Names="Tahoma" Font-Size="8pt"></HeaderStyle>
                            <CellStyle Font-Names="Tahoma" Font-Size="8pt">
                            </CellStyle>
                        </dx:GridViewDataTextColumn>
                        <dx:GridViewDataTextColumn VisibleIndex="1" Caption="AFFILIATE CODE" FieldName="AffiliateID"
                            Width="100px" HeaderStyle-HorizontalAlign="Center">
                            <HeaderStyle HorizontalAlign="Center" Font-Names="Tahoma" Font-Size="8pt"></HeaderStyle>
                            <CellStyle Font-Names="Tahoma" Font-Size="8pt">
                            </CellStyle>
                        </dx:GridViewDataTextColumn>
                        <dx:GridViewDataTextColumn VisibleIndex="2" Caption="DELIVERY LOCATION CODE" FieldName="DeliveryLocationCode"
                            Width="100px" HeaderStyle-HorizontalAlign="Center">
                            <HeaderStyle HorizontalAlign="Center" Font-Names="Tahoma" Font-Size="8pt"></HeaderStyle>
                            <CellStyle Font-Names="Tahoma" Font-Size="8pt">
                            </CellStyle>
                        </dx:GridViewDataTextColumn>
                        <dx:GridViewDataTextColumn Caption="DELIVERY LOCATION NAME" VisibleIndex="3" FieldName="DeliveryLocationName"
                            Width="150px">
                        </dx:GridViewDataTextColumn>
                        <dx:GridViewDataTextColumn Caption="ADDRESS" VisibleIndex="4" FieldName="Address"
                            Width="250px">
                        </dx:GridViewDataTextColumn>
                        <dx:GridViewDataTextColumn Caption="CITY" VisibleIndex="5" FieldName="City" Width="100px">
                        </dx:GridViewDataTextColumn>
                        <dx:GridViewDataTextColumn Caption="POSTAL CODE" VisibleIndex="6" FieldName="PostalCode"
                            Width="100px">
                        </dx:GridViewDataTextColumn>
                        <dx:GridViewDataTextColumn Caption="PHONE 1" VisibleIndex="7" FieldName="Phone1"
                            Width="150px">
                        </dx:GridViewDataTextColumn>
                        <dx:GridViewDataTextColumn Caption="PHONE 2" VisibleIndex="8" FieldName="Phone2"
                            Width="150px">
                        </dx:GridViewDataTextColumn>
                        <dx:GridViewDataTextColumn Caption="FAX" VisibleIndex="9" FieldName="Fax" Width="150px">
                        </dx:GridViewDataTextColumn>
                        <dx:GridViewDataTextColumn Caption="NPWP" VisibleIndex="10" FieldName="NPWP" Width="150px">
                        </dx:GridViewDataTextColumn>
                        <dx:GridViewDataTextColumn Caption="PO DELIVERED BY" VisibleIndex="11" FieldName="PODeliveryBy"
                            Width="150px">
                        </dx:GridViewDataTextColumn>
                        <dx:GridViewDataTextColumn Caption="DEFAULT CLS" VisibleIndex="12" FieldName="DefaultCls"
                            Width="80px">
                        </dx:GridViewDataTextColumn>
                        <dx:GridViewDataTextColumn VisibleIndex="13" Caption="REGISTER DATE" 
                            FieldName="EntryDate" Width="150px" HeaderStyle-HorizontalAlign="Center">
                        </dx:GridViewDataTextColumn>
                         <dx:GridViewDataTextColumn VisibleIndex="14" Caption="REGISTER USER" 
                            FieldName="EntryUser" Width="100px" HeaderStyle-HorizontalAlign="Center">
                        </dx:GridViewDataTextColumn>
                         <dx:GridViewDataTextColumn VisibleIndex="15" Caption="UPDATE DATE" 
                            FieldName="UpdateDate" Width="150px" HeaderStyle-HorizontalAlign="Center">
                        </dx:GridViewDataTextColumn>
                         <dx:GridViewDataTextColumn VisibleIndex="16" Caption="UPDATE USER" 
                            FieldName="UpdateUser" Width="100px" HeaderStyle-HorizontalAlign="Center">
                        </dx:GridViewDataTextColumn>
                        <dx:GridViewDataTextColumn VisibleIndex="17" Caption="DeleteCls" 
                            FieldName="DeleteCls" Width="0px" HeaderStyle-HorizontalAlign="Center">
                        </dx:GridViewDataTextColumn>
                    </Columns>
                    <SettingsBehavior AllowSelectByRowClick="True" AllowSort="False" ColumnResizeMode="Control"
                        EnableRowHotTrack="True" />
                    <SettingsPager Visible="False" PageSize="14" NumericButtonCount="10" AlwaysShowPager="True"
                        Mode="ShowAllRecords">
                        <Summary Text="Page {0} of {1} [{2} record(s)]" AllPagesText="Page {0} of {1} " />
                        <Summary AllPagesText="Page {0} of {1} " Text="Page {0} of {1} [{2} record(s)]">
                        </Summary>
                    </SettingsPager>
                    <Settings ShowHorizontalScrollBar="True" ShowVerticalScrollBar="True" ShowGroupButtons="False"
                        ShowStatusBar="Hidden" VerticalScrollableHeight="190" />
                    <Settings ShowGroupButtons="False" ShowVerticalScrollBar="True" ShowHorizontalScrollBar="True"
                        VerticalScrollableHeight="190" ShowStatusBar="Hidden"></Settings>
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
    </table>
    <div style="height: 8px;">
    </div>
    <table style="width: 100%;">
        <tr>
            <td height="70">
                <!-- INPUT AREA -->
                <table id="Table1" style="border-width: 1pt thin thin thin; border-style: ridge;
                    border-color: #9598A1; width: 100%; height: 25px;" width="100%">
                    <tr>
                           <td bgcolor="#FFD2A6" align="center" width="110px">
                        
                        <dx:ASPxLabel ID="ASPxLabel61" runat="server" Text="AFFILIATE CODE" Font-Names="Tahoma"
                            Font-Size="8pt" Width="110px">
                        </dx:ASPxLabel>
                        </td>
                        <td bgcolor="#FFD2A6" align="center" width="150px">
                            <dx:ASPxLabel ID="ASPxLabel56" runat="server" Text="DELIVERY LOCATION CODE" Font-Names="Tahoma"
                                Font-Size="8pt" Width="110px">
                            </dx:ASPxLabel>
                        </td>
                        <td bgcolor="#FFD2A6" align="center" width="250px">
                            <dx:ASPxLabel ID="ASPxLabel5" runat="server" Text="DELIVERY LOCATION NAME" Font-Names="Tahoma"
                                Font-Size="8pt" Width="250px">
                            </dx:ASPxLabel>
                        </td>
                        <td bgcolor="#FFD2A6" align="center">
                            <dx:ASPxLabel ID="ASPxLabel57" runat="server" Text="ADDRESS" Font-Names="Tahoma"
                                Font-Size="8pt" Width="300px">
                            </dx:ASPxLabel>
                        </td>
                        <td bgcolor="#FFD2A6" align="center">
                            <dx:ASPxLabel ID="ASPxLabel58" runat="server" Text="CITY" Font-Names="Tahoma" Font-Size="8pt"
                                Width="150px">
                            </dx:ASPxLabel>
                        </td>
                        <td bgcolor="#FFD2A6" align="center">
                            <dx:ASPxLabel ID="ASPxLabel59" runat="server" Text="POSTAL CODE" Font-Names="Tahoma"
                                Font-Size="8pt" Width="150px">
                            </dx:ASPxLabel>
                        </td>
                    </tr>
                    <tr>
                        <td align="left" width="110px">
                            <dx:ASPxComboBox ID="cboAffiliateCode" runat="server" Width="100%" 
                                ClientInstanceName="cboAffiliateCode" TabIndex="1">
                            </dx:ASPxComboBox>
                        </td>
                        <td align="left" width="110px">
                            <dx:ASPxTextBox ID="txtDeliveryLoc" runat="server" Width="100%" Height="20px" ClientInstanceName="txtDeliveryLoc"
                                Font-Names="Tahoma" Font-Size="8pt" MaxLength="20" BackColor="White" TabIndex="2"
                                ReadOnly="False">
                            </dx:ASPxTextBox>
                        </td>
                        <td align="left" width="250px">
                            <dx:ASPxTextBox ID="txtDeliveryLocName" runat="server" Width="100%" Height="20px"
                                ClientInstanceName="txtDeliveryLocName" Font-Names="Tahoma" 
                                Font-Size="8pt" MaxLength="150"
                                BackColor="White" ReadOnly="False" TabIndex="3">
                            </dx:ASPxTextBox>
                        </td>
                        <td align="left" Width="300px">
                            <dx:ASPxTextBox ID="txtAddress" runat="server" Width="100%" Height="20px" ClientInstanceName="txtAddress"
                                Font-Names="Tahoma" Font-Size="8pt" MaxLength="150" BackColor="White" 
                                ReadOnly="False" TabIndex="4">
                            </dx:ASPxTextBox>
                        </td>
                        <td align="left" Width="150px">
                            <dx:ASPxTextBox ID="txtCity" runat="server" Width="100%" Height="20px" ClientInstanceName="txtCity"
                                Font-Names="Tahoma" Font-Size="8pt" MaxLength="20" BackColor="White" 
                                ReadOnly="False" TabIndex="5">
                            </dx:ASPxTextBox>
                        </td>
                        <td align="left" Width="150px">
                            <dx:ASPxTextBox ID="txtPostalCode" runat="server" Width="100%" Height="20px" ClientInstanceName="txtPostalCode"
                                onkeypress="return numbersonly(event)" Font-Names="Tahoma" Font-Size="8pt" MaxLength="15"
                                BackColor="White" ReadOnly="False" TabIndex="5">
                            </dx:ASPxTextBox>
                        </td>
                    </tr>                    
                </table>
            </td>
        </tr>
        <tr>
            <td height="70">
                <!-- INPUT AREA -->
                <table id="Table2" style="border-width: 1pt thin thin thin; border-style: ridge;
                    border-color: #9598A1; width: 100%; height: 25px;">
                    <tr>
                        <td bgcolor="#FFD2A6" width="150px" align="center">
                            <dx:ASPxLabel ID="ASPxLabel1" runat="server" Text="PHONE 1" Font-Names="Tahoma" Font-Size="8pt"
                                Width="150px">
                            </dx:ASPxLabel>
                        </td>
                        <td bgcolor="#FFD2A6" width="150px" align="center">
                            <dx:ASPxLabel ID="ASPxLabel2" runat="server" Text="PHONE 2" Font-Names="Tahoma" Font-Size="8pt"
                                Width="150px">
                            </dx:ASPxLabel>
                        </td>
                        <td bgcolor="#FFD2A6" width="150px" align="center">
                            <dx:ASPxLabel ID="ASPxLabel3" runat="server" Text="FAX" Font-Names="Tahoma"
                                Font-Size="8pt" Width="150px">
                            </dx:ASPxLabel>
                            
                        </td>
                        <td bgcolor="#FFD2A6" width="150px" align="center">
                            <dx:ASPxLabel ID="ASPxLabel4" runat="server" Text="NPWP" Font-Names="Tahoma" Font-Size="8pt"
                                Width="150px" Height="16px">
                            </dx:ASPxLabel>
                        </td>
                        <td bgcolor="#FFD2A6" width="150px" align="center">
                            <dx:ASPxLabel ID="ASPxLabel60" runat="server" Text="PO DELIVERED BY" Font-Names="Tahoma"
                                Font-Size="8pt" Width="150px" Height="16px">
                            </dx:ASPxLabel>
                        </td>
                        <td bgcolor="#FFD2A6" width="150px" align="center">
                            <dx:ASPxLabel ID="ASPxLabel6" runat="server" Text="DEFAULT CLS" Font-Names="Tahoma"
                                Font-Size="8pt" Width="150px">
                            </dx:ASPxLabel>
                        </td>
                    </tr>
                    <tr>
                        <td width="150px">
                            <dx:ASPxTextBox ID="txtPhone1" runat="server" Width="100%" Height="20px" ClientInstanceName="txtPhone1"
                                Font-Names="Tahoma" Font-Size="8pt" MaxLength="20" BackColor="White" 
                                ReadOnly="False" TabIndex="5">
                            </dx:ASPxTextBox>
                        </td>
                        <td width="150px">
                            <dx:ASPxTextBox ID="txtPhone2" runat="server" Width="100%" Height="20px" ClientInstanceName="txtPhone2"
                                Font-Names="Tahoma" Font-Size="8pt" MaxLength="20" BackColor="White" 
                                ReadOnly="False" TabIndex="5">
                            </dx:ASPxTextBox>
                        </td>
                        <td width="150px">
                            <dx:ASPxTextBox ID="txtFax" runat="server" Width="100%" Height="20px" ClientInstanceName="txtFax"
                                Font-Names="Tahoma" Font-Size="8pt" MaxLength="20" BackColor="White" 
                                ReadOnly="False" TabIndex="5">
                            </dx:ASPxTextBox>
                        </td>
                        <td width="150px">
                            <dx:ASPxTextBox ID="txtNPWP" runat="server" Width="100%" Height="20px" ClientInstanceName="txtNPWP"
                                Font-Names="Tahoma" Font-Size="8pt" MaxLength="20" BackColor="White" 
                                ReadOnly="False" TabIndex="5">
                            </dx:ASPxTextBox>
                        </td>
                        <td width="150px">
                            <dx:ASPxComboBox ID="cboPODelby" runat="server" ClientInstanceName="cboPODelby" 
                                Width="100%" TabIndex="5">
                                <Items>
                                    <dx:ListEditItem Text="SUPPLIER" Value="0" />
                                    <dx:ListEditItem Text="PASI" Value="1" />
                                </Items>
                            </dx:ASPxComboBox>
                        </td>
                        <td width="150px">
                            <dx:ASPxComboBox ID="cboDefault" runat="server" ClientInstanceName="cboDefault" 
                                Width="100%" TabIndex="5">
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
                    Width="85px" Font-Size="8pt" TabIndex="11" ClientInstanceName="btnSubMenu">
                </dx:ASPxButton>
            </td>            
            <td valign="top" align="right" style="width: 50px;">
                <dx:ASPxButton ID="btnUpload" runat="server" Text="UPLOAD" ClientInstanceName="btnUpload"
                    Font-Names="Tahoma" Width="85px" AutoPostBack="False" Font-Size="8pt" 
                    TabIndex="10" >
                    <ClientSideEvents Click="function(s, e) {grid.PerformCallback('save');}" />
                </dx:ASPxButton>
            </td>
            <td valign="top" align="right" style="width: 50px;">
                <dx:ASPxButton ID="btnDownload" runat="server" Text="DOWNLOAD" ClientInstanceName="btnDownload"
                    Font-Names="Tahoma" Width="85px" AutoPostBack="False" Font-Size="8pt" 
                    TabIndex="9" >
                    <ClientSideEvents Click="function(s, e) {grid.PerformCallback('downloadSummary');}" />
                </dx:ASPxButton>
            </td>       
            <td valign="top" align="right" style="width: 50px;">
                <dx:ASPxButton ID="btnClear" runat="server" Text="CLEAR" Font-Names="Tahoma" Width="90px"
                    AutoPostBack="False" Font-Size="8pt" TabIndex="8" 
                    ClientInstanceName="btnClear">
                    <ClientSideEvents Click="function(s, e) {

        cboAffiliateCode.SetText('');
        txtDeliveryLoc.SetText('');
        txtDeliveryLocName.SetText('');
        txtAddress.SetText('');
        txtCity.SetText('');
        txtPostalCode.SetText('');
        txtPhone1.SetText('');
        txtPhone2.SetText('');
        txtFax.SetText('');
        txtNPWP.SetText('');
        cboPODelby.SetText('');
        cboDefault.SetText('');        
        lblInfo.SetText('');

        cboAffiliateCode.GetInputElement().setAttribute('style', 'background:#FFFFFF;foreground:#FFFFFF;');
        cboAffiliateCode.GetInputElement().readOnly = false;
                        cboAffiliateCode.SetEnabled(true);

                     
            txtDeliveryLoc.GetInputElement().setAttribute('style', 'background:#FFFFFF;foreground:#FFFFFF;');
              txtDeliveryLoc.GetInputElement().readOnly = false;
                        txtDeliveryLoc.SetEnabled(true);

}" />
                </dx:ASPxButton>
            </td>
            <td align="right" style="width: 80px;">
                <dx:ASPxButton ID="btnDelete" runat="server" Text="DELETE" Font-Names="Tahoma" Width="80px"
                    AutoPostBack="False" Font-Size="8pt" TabIndex="7" 
                    ClientInstanceName="btnDelete">
                    <ClientSideEvents Click="function(s, e) {
                        up_delete();
                        grid.PerformCallback('load');

                        cboAffiliateCode.SetText('');
        txtDeliveryLoc.SetText('');
        txtDeliveryLocName.SetText('');
        txtAddress.SetText('');
        txtCity.SetText('');
        txtPostalCode.SetText('');
        txtPhone1.SetText('');
        txtPhone2.SetText('');
        txtFax.SetText('');
        txtNPWP.SetText('');
        cboPODelby.SetText('');
        cboDefault.SetText('');

        
        cboAffiliateCode.GetInputElement().setAttribute('style', 'background:#FFFFFF;foreground:#FFFFFF;');
        cboAffiliateCode.GetInputElement().readOnly = false;
                        cboAffiliateCode.SetEnabled(true);

                     
            txtDeliveryLoc.GetInputElement().setAttribute('style', 'background:#FFFFFF;foreground:#FFFFFF;');
              txtDeliveryLoc.GetInputElement().readOnly = false;
                        txtDeliveryLoc.SetEnabled(true);
                        
                    }" />
                </dx:ASPxButton>
            </td>
            <td valign="top" align="right" style="width: 50px;">
                <dx:ASPxButton ID="btnSubmit" runat="server" Text="SAVE" Font-Names="Verdana" Width="90px"
                    AutoPostBack="False" Font-Size="8pt" TabIndex="6">
                    <ClientSideEvents Click="function(s, e) {
                    grid.SetFocusedRowIndex(-1);
                        validasubmit();
                        up_Insert();

                        
                        grid.PerformCallback('load');
                        

                        grid.SetFocusedRowIndex(-1);



                        cboAffiliateCode.SetText('');
        txtDeliveryLoc.SetText('');
        txtDeliveryLocName.SetText('');
        txtAddress.SetText('');
        txtCity.SetText('');
        txtPostalCode.SetText('');
        txtPhone1.SetText('');
        txtPhone2.SetText('');
        txtFax.SetText('');
        txtNPWP.SetText('');
        cboPODelby.SetText('');
        cboDefault.SetText('');

        
        cboAffiliateCode.GetInputElement().setAttribute('style', 'background:#FFFFFF;foreground:#FFFFFF;');
        cboAffiliateCode.GetInputElement().readOnly = false;
                        cboAffiliateCode.SetEnabled(true);

                     
            txtDeliveryLoc.GetInputElement().setAttribute('style', 'background:#FFFFFF;foreground:#FFFFFF;');
              txtDeliveryLoc.GetInputElement().readOnly = false;
                        txtDeliveryLoc.SetEnabled(true);
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
    <dx:ASPxCallback ID="AffiliateSubmit" runat="server" ClientInstanceName="AffiliateSubmit">
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


                }
                
            } else {
                lblInfo.SetText('');
            }  
        }" />
    </dx:ASPxCallback>
</asp:Content>
