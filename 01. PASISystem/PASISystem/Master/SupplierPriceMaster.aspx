<%@ Page Title="" Language="vb" AutoEventWireup="false" MasterPageFile="~/SiteContent.Master"
    CodeBehind="SupplierPriceMaster.aspx.vb" Inherits="PASISystem.SupplierPriceMaster" %>

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
        #Table1
        {
            width: 986px;
            margin-left: 0px;
        }
        .style23
        {
            width: 94px;
            height: 25px;
        }
        .style18
        {
            width: 112px;
            height: 25px;
        }
        .style7
        {
            height: 25px;
            width: 401px;
        }
        .style20
        {
            width: 100px;
            height: 25px;
        }
        .style3
        {
            height: 25px;
            width: 22px;
        }
        .style21
        {
            height: 25px;
            width: 80px;
        }
        .style6
        {
            width: 94px;
        }
        .style19
        {
            width: 112px;
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

            if (currentColumnName == "NoUrut" || currentColumnName == "PartNo" || currentColumnName == "PartName" || currentColumnName == "SupplierID" || currentColumnName == "SupplierName"
            || currentColumnName == "StartDate" || currentColumnName == "EndDate" || currentColumnName == "EntryDate" || currentColumnName == "Currency" || currentColumnName == "Price" || currentColumnName == "PriceDesc"
            || currentColumnName == "PackingCls" || currentColumnName == "PackingDesc" || currentColumnName == "DeliveryID" || currentColumnName == "DeliveryDesc") {
                e.cancel = true;
            }

            currentEditableVisibleIndex = e.visibleIndex;
        }

        function OnGridFocusedRowChanged() {
            grid.GetRowValues(grid.GetFocusedRowIndex(), "NoUrut;PartNo;PartName;SupplierID;SupplierName;StartDate;EndDate;EntryDate;CurrCls;Price;PriceCls;PriceDesc;PackingCls;PackingDesc;DeliveryID;DeliveryDesc;DeleteCls", OnGetRowValues);
        }
        function OnGetRowValues(values) {
            if (values[0] != "" && values[0] != null && values[0] != "null") {

                cboPartNo2.SetText(values[1]);
                txtPartNo2.SetText(values[2]);
                cboSupplier2.SetText(values[3]);
                txtSupplier2.SetText(values[4]);
                dt4.SetText(values[5]);
                dt5.SetText(values[6]);
                dt6.SetText(values[7]);
                CboCurrency.SetText(values[8]);
                TxtPrice.SetText(values[9]);

                cboPriceCls.SetText(values[10]);
                txtPriceCls.SetText(values[11]);
                cboPacking.SetText(values[12]);
                txtPacking.SetText(values[13]);
                var descDelivery = values[14];
                HF.Set('hfTest', descDelivery)
                //HF.Set('HTcls') = values[14];
                cboDeliveryID.SetText(values[14]);
                txtDeliveryDesc.SetText(values[15]);
                var vDeleteCls = values[16];

                cboPartNo2.GetInputElement().setAttribute('style', 'background:#CCCCCC;foreground:#CCCCCC;');
                cboPartNo2.GetInputElement().readOnly = true;
                cboPartNo2.SetEnabled(false);

                cboSupplier2.GetInputElement().setAttribute('style', 'background:#CCCCCC;foreground:#CCCCCC;');
                cboSupplier2.GetInputElement().readOnly = true;
                cboSupplier2.SetEnabled(false);

                dt4.GetInputElement().setAttribute('style', 'background:#CCCCCC;foreground:#CCCCCC;');
                dt4.GetInputElement().readOnly = true;
                dt4.SetEnabled(false);

                CboCurrency.GetInputElement().setAttribute('style', 'background:#CCCCCC;foreground:#CCCCCC;');
                CboCurrency.GetInputElement().readOnly = true;
                CboCurrency.SetEnabled(false);

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
            var pAffiliateID = cboSupplier2.GetText();
            var pStartDate = dt4.GetValue();
            var pCurrency = CboCurrency.GetValue();
            var pPacking = cboPacking.GetText();
            var pDeliveryID = cboDeliveryID.GetText();

            var pub_year = pStartDate.getFullYear();
            var pub_month = pStartDate.getMonth() + 1;
            var pub_day = pStartDate.getDate();

            var vStartDate = pub_year + '-' + pub_month + '-' + pub_day;

            grid.PerformCallback('delete|' + pPartNo + '|' + pAffiliateID + '|' + vStartDate + '|' + pCurrency + '|' + pPacking + '|' + pDeliveryID);
        }

        function readonly() {
            txtPartID.GetInputElement().setAttribute('style', 'background:#CCCCCC;foreground:#CCCCCC;');
            txtPartID.GetInputElement().readOnly = true;
            lblInfo.SetText('');
        }

        function validasubmit() {
            lblInfo.GetMainElement().style.color = 'Red';
            if (cboPartNo2.GetText() == "") {
                lblInfo.SetText("[6011] Please Select Part No. first!");
                cboPartNo2.Focus();
                e.ProcessOnServer = false;
                return false;

            }
            lblInfo.GetMainElement().style.color = 'Red';
            if (cboSupplier2.GetText() == "") {
                lblInfo.SetText("[6011] Please Select Supplier first!");
                cboSupplier2.Focus();
                e.ProcessOnServer = false;
                return false;
            }

            lblInfo.GetMainElement().style.color = 'Red';
            if (CboCurrency.GetText() == "") {
                lblInfo.SetText("[6011] Please Select Currency first!");
                CboCurrency.Focus();
                e.ProcessOnServer = false;
                return false;
            }

            lblInfo.GetMainElement().style.color = 'Red';
            if (TxtPrice.GetText() == "") {
                lblInfo.SetText("[6011] Please Input the Price first!");
                TxtPrice.Focus();
                e.ProcessOnServer = false;
                return false;
            }
        }

        function validasearch() {
            lblInfo.GetMainElement().style.color = 'Red';
            if (cboPartNo.GetText() == "") {
                lblInfo.SetText("[6011] Please Select Part No. first!");
                cboPartNo.Focus();
                e.ProcessOnServer = false;
                return false;

            }
            lblInfo.GetMainElement().style.color = 'Red';
            if (cboSupplier.GetText() == "") {
                lblInfo.SetText("[6011] Please Select Supplier first!");
                cboSupplier.Focus();
                e.ProcessOnServer = false;
                return false;
            }

        }

        function up_Insert() {
            var pIsUpdate = '';
            var pPartNo = cboPartNo2.GetText();
            var pAffiliateID = cboSupplier2.GetText();
            var pStartDate = dt4.GetValue();
            var pEndDate = dt5.GetValue();
            var pEffectiveDate = dt6.GetValue();
            var pCurrency = CboCurrency.GetValue();
            var pPrice = TxtPrice.GetText();
            var pPacking = cboPacking.GetText();
            var pPriceCls = cboPriceCls.GetText();
            var pDeliveryID = cboDeliveryID.GetText();

            var pub_year;
            var pub_month;
            var pub_day;

            pub_year = pStartDate.getFullYear();
            pub_month = pStartDate.getMonth() + 1;
            pub_day = pStartDate.getDate();
            var vStartDate = pub_year + '-' + pub_month + '-' + pub_day;

            pub_year = pEndDate.getFullYear();
            pub_month = pEndDate.getMonth() + 1;
            pub_day = pEndDate.getDate();
            var vEndDate = pub_year + '-' + pub_month + '-' + pub_day;

            pub_year = pEffectiveDate.getFullYear();
            pub_month = pEffectiveDate.getMonth() + 1;
            pub_day = pEffectiveDate.getDate();
            var vEffectiveDate = pub_year + '-' + pub_month + '-' + pub_day;

            grid.PerformCallback('save|' + pIsUpdate + '|' + pPartNo + '|' + pAffiliateID + '|' + vStartDate + '|' + vEndDate + '|' + vEffectiveDate + '|' + pCurrency + '|' + pPrice + '|' + pPacking + '|' + pPriceCls + '|' + pDeliveryID);

        }

   
    </script>
</asp:Content>
<asp:Content ID="Content2" ContentPlaceHolderID="Content" runat="server">
    <table style="width: 100%;">
        <tr>
            <td colspan="2">
                <table style="border-width: 1pt thin thin thin; border-style: ridge; border-color: #9598A1;
                    width: 100%;">
                    <tr>
                        <td height="30">
                            <table id="Table1">
                                <tr>
                                    <td align="left" valign="middle" class="style23">
                                        <dx:ASPxLabel ID="ASPxLabel4" runat="server" Text="PART NO." Font-Names="Tahoma"
                                            Font-Size="8pt">
                                        </dx:ASPxLabel>
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
                                        <dx:ASPxComboBox ID="cboPartNo" runat="server" ClientInstanceName="cboPartNo" Width="110px"
                                            Font-Size="8pt" Font-Names="Tahoma" TextFormatString="{0}">
                                            <ClientSideEvents SelectedIndexChanged="function(s, e) {
	                                            txtPartNo.SetText(cboPartNo.GetSelectedItem().GetColumnText(1));
                                                grid.PerformCallback('kosong');
	                                            lblInfo.SetText('');	
                                            }" />
                                            <LoadingPanelStyle ImageSpacing="5px">
                                            </LoadingPanelStyle>
                                        </dx:ASPxComboBox>
                                    </td>
                                    <td align="left" valign="middle" class="style7">
                                        <dx:ASPxTextBox ID="txtPartNo" runat="server" Width="400px" Height="20px" ClientInstanceName="txtPartNo"
                                            Font-Names="Tahoma" Font-Size="8pt" MaxLength="100" BackColor="#CCCCCC" ReadOnly="True"
                                            Style="margin-right: 31px">
                                        </dx:ASPxTextBox>
                                    </td>
                                    <td align="left" valign="middle" class="style20">
                                        <dx:ASPxLabel ID="ASPxLabel66" runat="server" Text="START DATE" Font-Names="Tahoma"
                                            Font-Size="8pt">
                                        </dx:ASPxLabel>
                                    </td>
                                    <td align="left" valign="middle" class="style3">
                                        <dx:ASPxCheckBox ID="checkbox1" runat="server" ClientInstanceName="checkbox1" TabIndex="2">
                                        </dx:ASPxCheckBox>
                                    </td>
                                    <td align="left" valign="middle" class="style20">
                                        <dx:ASPxDateEdit ID="dt1" runat="server" ClientInstanceName="dt1" EditFormatString="dd MMM yyyy"
                                            Font-Names="Tahoma" Font-Size="8pt" Width="100px" TabIndex="3">
                                        </dx:ASPxDateEdit>
                                    </td>
                                    <td align="left" valign="middle" class="style21">
                                    </td>
                                </tr>
                                <tr>
                                    <td align="left" valign="middle" class="style23">
                                        <dx:ASPxLabel ID="ASPxLabel1" runat="server" Text="SUPPLIER CODE" Font-Names="Tahoma"
                                            Font-Size="8pt">
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
                                        <dx:ASPxComboBox ID="cboSupplier" runat="server" ClientInstanceName="cboSupplier"
                                            Width="110px" Font-Size="8pt" Font-Names="Tahoma" TextFormatString="{0}" TabIndex="1">
                                            <ClientSideEvents SelectedIndexChanged="function(s, e) {
	                                            txtSupplier.SetText(cboSupplier.GetSelectedItem().GetColumnText(1));
                                                grid.PerformCallback('up_kosong');                              
	                                            lblInfo.SetText('');	
                                            }" />
                                            <LoadingPanelStyle ImageSpacing="5px">
                                            </LoadingPanelStyle>
                                        </dx:ASPxComboBox>
                                    </td>
                                    <td align="left" valign="middle" class="style7">
                                        <dx:ASPxTextBox ID="txtSupplier" runat="server" Width="400px" Height="20px" ClientInstanceName="txtSupplier"
                                            Font-Names="Tahoma" Font-Size="8pt" MaxLength="100" BackColor="#CCCCCC" ReadOnly="True">
                                        </dx:ASPxTextBox>
                                    </td>
                                    <td align="left" valign="middle" class="style20">
                                        <dx:ASPxLabel ID="ASPxLabel67" runat="server" Text="END DATE" Font-Names="Tahoma"
                                            Font-Size="8pt">
                                        </dx:ASPxLabel>
                                    </td>
                                    <td align="left" valign="middle" class="style3">
                                        <dx:ASPxCheckBox ID="checkbox2" runat="server" ClientInstanceName="checkbox2" TabIndex="4">
                                        </dx:ASPxCheckBox>
                                    </td>
                                    <td align="left" valign="middle" class="style20">
                                        <dx:ASPxDateEdit ID="dt2" runat="server" ClientInstanceName="dt2" EditFormatString="dd MMM yyyy"
                                            Font-Names="Tahoma" Font-Size="8pt" Width="100px" TabIndex="5">
                                        </dx:ASPxDateEdit>
                                    </td>
                                    <td align="left" valign="middle" class="style21">
                                    </td>
                                </tr>
                                <tr>
                                    <td align="left" valign="middle" height="25px" class="style6">
                                        &nbsp;
                                    </td>
                                    <td align="left" valign="middle" height="25px" class="style19">
                                        <%--<dx:ASPxComboBox ID="cboSupplier" runat="server" TextFormatString="{0}" 
                                            DropDownStyle="DropDown" Height="20px" Width="100%" MaxLength="1"
                                            IncrementalFilteringMode="StartsWith" Font-Names="Tahoma" 
                                            Font-Size="8pt">
                                        <ClientSideEvents SelectedIndexChanged="function(s, e) {
	                                        txtSupplier.SetText(cboSupplier.GetSelectedItem().GetColumnText(1));	                                        
	                                        grid.PerformCallback('kosong');
	                                        lblErrMsg.SetText('');	
                                        }" />
                                        </dx:ASPxComboBox>   --%>
                                    </td>
                                    <td align="left" valign="middle" class="style7">
                                        &nbsp;
                                    </td>
                                    <td align="left" valign="middle" class="style20">
                                        <dx:ASPxLabel ID="ASPxLabel68" runat="server" Text="EFFECTIVE DATE" Font-Names="Tahoma"
                                            Font-Size="8pt">
                                        </dx:ASPxLabel>
                                    </td>
                                    <td align="left" valign="middle" class="style3">
                                        <dx:ASPxCheckBox ID="checkbox3" runat="server" ClientInstanceName="checkbox3" TabIndex="6">
                                        </dx:ASPxCheckBox>
                                    </td>
                                    <td align="left" valign="middle" class="style20">
                                        <dx:ASPxDateEdit ID="dt3" runat="server" ClientInstanceName="dt3" EditFormatString="dd MMM yyyy"
                                            Font-Names="Tahoma" Font-Size="8pt" Width="100px" TabIndex="7">
                                        </dx:ASPxDateEdit>
                                    </td>
                                    <td align="right" valign="middle" class="style21">
                                        <table style="width: 100%;">
                                            <tr>
                                                <td align="right" valign="middle" style="height: 25px; width: 90px;">
                                                    <dx:ASPxButton ID="btnRefresh" runat="server" Text="SEARCH" Font-Names="Tahoma" Width="85px"
                                                        AutoPostBack="False" Font-Size="8pt" TabIndex="8">
                                                        <ClientSideEvents Click="function(s, e) {
                                                        validasearch();                                                 
                                                        grid.PerformCallback('load');
                                                        cboPartNo2.SetText('');
                                                        txtPartNo2.SetText('');
                                                        cboSupplier2.SetText('');
                                                        txtSupplier2.SetText('');
                                                        CboCurrency.SetText('');
                                                        TxtPrice.SetText('');
                                                        cboPacking.SetText('');
                                                        txtPacking.SetText('');
                                                        cboPriceCls.SetText('');
                                                        txtPriceCls.SetText('');
                                                        cboDeliveryID.SetText('');
                                                        txtDeliveryDesc.SetText('');

                                                        cboSupplier2.GetInputElement().setAttribute('style', 'background:#FFFFFF;foreground:#FFFFFF;');
                                                        cboSupplier2.GetInputElement().readOnly = false;
                                                        cboSupplier2.SetEnabled(true);
            
                                                        cboPartNo2.GetInputElement().setAttribute('style', 'background:#FFFFFF;foreground:#FFFFFF;');
                                                        cboPartNo2.GetInputElement().readOnly = false;
                                                        cboPartNo2.SetEnabled(true);

                                                        dt4.GetInputElement().setAttribute('style', 'background:#FFFFFF;foreground:#FFFFFF;');
                                                        dt4.GetInputElement().readOnly = false;
                                                        dt4.SetEnabled(true);

                                                        CboCurrency.GetInputElement().setAttribute('style', 'background:#FFFFFF;foreground:#FFFFFF;');
                                                        CboCurrency.GetInputElement().readOnly = false;
                                                        CboCurrency.SetEnabled(true);
                                                        lblInfo.SetText('');
                                                        }" />
                                                    </dx:ASPxButton>
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
        <div style="height: 1px;">
        </div>
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
                <dx:ASPxGridView ID="grid" runat="server" Width="100%" Font-Names="Tahoma" KeyFieldName="PartNo;SupplierID;StartDate;EntryDate;CurrCls;PackingCls;DeliveryID;PriceCls"
                    AutoGenerateColumns="False" ClientInstanceName="grid" Font-Size="8pt">
                    <ClientSideEvents Init="OnInit" FocusedRowChanged="function(s, e) {OnGridFocusedRowChanged();}" 
                        EndCallback="function(s, e) {
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
                                lblInfo.SetText('');}  
                        }" 
                        RowClick="function(s, e) {delete s.cpMessage;}" 
                        BatchEditStartEditing="OnBatchEditStartEditing" />
                    <Columns>
                        <dx:GridViewDataTextColumn VisibleIndex="0" Caption="NO" FieldName="NoUrut" Width="30px"
                            HeaderStyle-HorizontalAlign="Center">
                            <HeaderStyle HorizontalAlign="Center" Font-Names="Tahoma" Font-Size="8pt"></HeaderStyle>
                            <CellStyle Font-Names="Tahoma" Font-Size="8pt">
                            </CellStyle>
                        </dx:GridViewDataTextColumn>
                        <dx:GridViewDataTextColumn VisibleIndex="1" Caption="PART NO." FieldName="PartNo"
                            Width="110px" HeaderStyle-HorizontalAlign="Center">
                            <HeaderStyle HorizontalAlign="Center" Font-Names="Tahoma" Font-Size="8pt"></HeaderStyle>
                            <CellStyle Font-Names="Tahoma" Font-Size="8pt">
                            </CellStyle>
                        </dx:GridViewDataTextColumn>
                        <dx:GridViewDataTextColumn VisibleIndex="2" Caption="PART NAME" FieldName="PartName"
                            Width="210px" HeaderStyle-HorizontalAlign="Center">
                            <HeaderStyle HorizontalAlign="Center" Font-Names="Tahoma" Font-Size="8pt"></HeaderStyle>
                            <CellStyle Font-Names="Tahoma" Font-Size="8pt">
                            </CellStyle>
                        </dx:GridViewDataTextColumn>
                        <dx:GridViewDataTextColumn VisibleIndex="3" Caption="SUPPLIER CODE" FieldName="SupplierID"
                            Width="110px" HeaderStyle-HorizontalAlign="Center">
                            <HeaderStyle HorizontalAlign="Center" Font-Names="Tahoma" Font-Size="8pt"></HeaderStyle>
                            <CellStyle Font-Names="Tahoma" Font-Size="8pt">
                            </CellStyle>
                        </dx:GridViewDataTextColumn>
                        <dx:GridViewDataTextColumn VisibleIndex="6" Caption="SUPPLIER NAME" FieldName="SupplierName"
                            Width="210px" HeaderStyle-HorizontalAlign="Center">
                            <HeaderStyle HorizontalAlign="Center" Font-Names="Tahoma" Font-Size="8pt"></HeaderStyle>
                            <CellStyle Font-Names="Tahoma" Font-Size="8pt">
                            </CellStyle>
                        </dx:GridViewDataTextColumn>
                        <dx:GridViewDataTextColumn VisibleIndex="4" Caption="PACKING CLS" 
                            FieldName="PackingCls" Width="0px" HeaderStyle-HorizontalAlign="Center">
<HeaderStyle HorizontalAlign="Center"></HeaderStyle>

                            <CellStyle Font-Names="Tahoma" Font-Size="8pt">
                            </CellStyle>
                        </dx:GridViewDataTextColumn>
                        <dx:GridViewDataTextColumn VisibleIndex="12" Caption="PACKING DESCRIPTION" 
                            FieldName="PackingDesc" Width="150px" 
                            HeaderStyle-HorizontalAlign="Center">
<HeaderStyle HorizontalAlign="Center"></HeaderStyle>

                            <CellStyle Font-Names="Tahoma" Font-Size="8pt">
                            </CellStyle>
                        </dx:GridViewDataTextColumn>
                        <dx:GridViewDataTextColumn VisibleIndex="5" Caption="PRICE CLS" 
                            FieldName="PriceCls" Width="0px" HeaderStyle-HorizontalAlign="Center">
<HeaderStyle HorizontalAlign="Center"></HeaderStyle>

                            <CellStyle Font-Names="Tahoma" Font-Size="8pt">
                            </CellStyle>
                        </dx:GridViewDataTextColumn>
                        <dx:GridViewDataTextColumn VisibleIndex="13" Caption="PRICE DESCRIPTION" 
                            FieldName="PriceDesc" Width="200px" 
                            HeaderStyle-HorizontalAlign="Center">
<HeaderStyle HorizontalAlign="Center"></HeaderStyle>

                            <CellStyle Font-Names="Tahoma" Font-Size="8pt">
                            </CellStyle>
                        </dx:GridViewDataTextColumn>
                        <dx:GridViewDataDateColumn Caption="START DATE" FieldName="StartDate" VisibleIndex="7"
                            Width="110px">
                            <PropertiesDateEdit DisplayFormatString="" EditFormatString="dd MMM yyyy">
                            </PropertiesDateEdit>
                            <HeaderStyle Font-Names="Tahoma" Font-Size="8pt" HorizontalAlign="Center" />
                            <CellStyle Font-Names="Tahoma" Font-Size="8pt">
                            </CellStyle>
                        </dx:GridViewDataDateColumn>
                        <dx:GridViewDataDateColumn Caption="END DATE" FieldName="EndDate" VisibleIndex="8"
                            Width="110px">
                            <PropertiesDateEdit DisplayFormatString="" EditFormatString="dd MMM yyyy">
                            </PropertiesDateEdit>
                            <HeaderStyle Font-Names="Tahoma" Font-Size="8pt" HorizontalAlign="Center" />
                            <CellStyle Font-Names="Tahoma" Font-Size="8pt">
                            </CellStyle>
                        </dx:GridViewDataDateColumn>
                        <dx:GridViewDataDateColumn Caption="EFFECTIVE DATE" FieldName="EffectiveDate" VisibleIndex="9"
                            Width="110px">
                            <PropertiesDateEdit DisplayFormatString="" EditFormatString="dd MMM yyyy">
                            </PropertiesDateEdit>
                            <HeaderStyle Font-Names="Tahoma" Font-Size="8pt" HorizontalAlign="Center" />
                            <CellStyle Font-Names="Tahoma" Font-Size="8pt">
                            </CellStyle>
                        </dx:GridViewDataDateColumn>
                        <dx:GridViewDataTextColumn VisibleIndex="10" Caption="CURR" FieldName="CurrCls" HeaderStyle-HorizontalAlign="Center"
                            Width="50px">
                            <HeaderStyle HorizontalAlign="Center" Font-Names="Tahoma" Font-Size="8pt"></HeaderStyle>
                            <CellStyle Font-Names="Tahoma" Font-Size="8pt">
                            </CellStyle>
                        </dx:GridViewDataTextColumn>
                        <dx:GridViewDataTextColumn Caption="PRICE" FieldName="Price" VisibleIndex="11" 
                            Width="130px">
                            <EditCellStyle HorizontalAlign="Right">
                            </EditCellStyle>
                            <PropertiesTextEdit DisplayFormatString="{0:n4}">
                                        <MaskSettings Mask="<0..9999999999999g>" IncludeLiterals="DecimalSymbol" />
                                        <ValidationSettings ErrorDisplayMode="None">
                                        </ValidationSettings>
                                    </PropertiesTextEdit>
                            <HeaderStyle Font-Names="Tahoma" Font-Size="8pt" HorizontalAlign="Center" />
                            <CellStyle Font-Names="Tahoma" Font-Size="8pt" HorizontalAlign="Right">
                            </CellStyle>
                        </dx:GridViewDataTextColumn>
                        <dx:GridViewDataTextColumn Caption="AFFILIATE CODE" FieldName="DeliveryID" 
                            VisibleIndex="14">
                        </dx:GridViewDataTextColumn>
                        <dx:GridViewDataTextColumn Caption="DELIVERY LOCATION NAME" 
                            FieldName="DeliveryDesc" VisibleIndex="15" Width="0px">
                        </dx:GridViewDataTextColumn>
                        <dx:GridViewDataTextColumn VisibleIndex="16" Caption="REGISTER DATE" 
                            FieldName="EntryDate" Width="150px" HeaderStyle-HorizontalAlign="Center">
                        </dx:GridViewDataTextColumn>
                         <dx:GridViewDataTextColumn VisibleIndex="17" Caption="REGISTER USER" 
                            FieldName="EntryUser" Width="100px" HeaderStyle-HorizontalAlign="Center">
                        </dx:GridViewDataTextColumn>
                         <dx:GridViewDataTextColumn VisibleIndex="18" Caption="UPDATE DATE" 
                            FieldName="UpdateDate" Width="150px" HeaderStyle-HorizontalAlign="Center">
                        </dx:GridViewDataTextColumn>
                         <dx:GridViewDataTextColumn VisibleIndex="19" Caption="UPDATE USER" 
                            FieldName="UpdateUser" Width="100px" HeaderStyle-HorizontalAlign="Center">
                        </dx:GridViewDataTextColumn>
                        <dx:GridViewDataTextColumn VisibleIndex="20" Caption="DeleteCls" 
                            FieldName="DeleteCls" Width="0px" HeaderStyle-HorizontalAlign="Center">
                        </dx:GridViewDataTextColumn>
                    </Columns>
                    <SettingsBehavior AllowSort="False" AllowSelectByRowClick="True" ColumnResizeMode="Control"
                        EnableRowHotTrack="True" AllowFocusedRow="True"></SettingsBehavior>
                    <SettingsPager Visible="True" PageSize="100" NumericButtonCount="10" AlwaysShowPager="True"
                        Mode="ShowAllRecords" Position="Top">
                        <Summary Text="Page {0} of {1} [{2} record(s)]" AllPagesText="Page {0} of {1} " />
                        <Summary AllPagesText="Page {0} of {1} " Text="Page {0} of {1} [{2} record(s)]">
                        </Summary>
                    </SettingsPager>
                    <SettingsEditing Mode="Batch" NewItemRowPosition="Bottom" EditFormColumnCount="10">
                        <BatchEditSettings ShowConfirmOnLosingChanges="False" />
                    </SettingsEditing>
                    <Settings ShowHorizontalScrollBar="True" ShowVerticalScrollBar="True" ShowGroupButtons="False"
                        ShowStatusBar="Hidden" VerticalScrollableHeight="190" />
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
            <td height="50">
                <!-- INPUT AREA -->
                <table id="tbl1" style="border-width: 1pt thin thin thin; border-style: ridge; border-color: #9598A1;
                    width: 100%; height: 50px;">
                    <tr>
                        <td valign="top" bgcolor="#FFD2A6" width="140px" align="center">
                            <dx:ASPxLabel ID="ASPxLabel53" runat="server" Text="PART NO." Font-Names="Tahoma"
                                Font-Size="8pt" Width="140px" Height="16px">
                            </dx:ASPxLabel>
                        </td>
                        <td valign="top" bgcolor="#FFD2A6" align="center" width="210px">
                            <dx:ASPxLabel ID="ASPxLabel8" runat="server" Text="PART NAME" Font-Names="Tahoma"
                                Font-Size="8pt" Width="210px">
                            </dx:ASPxLabel>
                        </td>
                        <td valign="top" bgcolor="#FFD2A6" align="center" width="110px">
                            <dx:ASPxLabel ID="ASPxLabel56" runat="server" Text="SUPPLIER CODE" Font-Names="Tahoma"
                                Font-Size="8pt" Width="110px">
                            </dx:ASPxLabel>
                        </td>
                        <td valign="top" bgcolor="#FFD2A6" width="210px" align="center">
                            <dx:ASPxLabel ID="ASPxLabel5" runat="server" Text="SUPPLIER NAME" Font-Names="Tahoma"
                                Font-Size="8pt" Width="210px">
                            </dx:ASPxLabel>
                        </td>
                       <td valign="top" bgcolor="#FFD2A6" width="110px" align="center">
                            <dx:ASPxLabel ID="ASPxLabel55" runat="server" Text="START DATE" Font-Names="Tahoma"
                                Font-Size="8pt" Width="110px">
                            </dx:ASPxLabel>
                        </td>
                        <td valign="top" bgcolor="#FFD2A6" width="110px" align="center">
                            <dx:ASPxLabel ID="ASPxLabel3" runat="server" Text="END DATE" Font-Names="Tahoma"
                                Font-Size="8pt" Width="110px">
                            </dx:ASPxLabel>
                        </td>
                        <td valign="top" bgcolor="#FFD2A6" width="110px" align="center">
                            <dx:ASPxLabel ID="ASPxLabel2" runat="server" Text="EFFECTIVE DATE" Font-Names="Tahoma"
                                Font-Size="8pt" Width="110px">
                            </dx:ASPxLabel>
                        </td>
                        <td valign="top" bgcolor="#FFD2A6" width="50px" align="center">
                            <dx:ASPxLabel ID="ASPxLabel6" runat="server" Text="CURR" Font-Names="Tahoma" Font-Size="8pt"
                                Width="50px">
                            </dx:ASPxLabel>
                        </td>
                        <td valign="top" bgcolor="#FFD2A6" width="110px" align="center">
                            <dx:ASPxLabel ID="ASPxLabel7" runat="server" Text="PRICE" Font-Names="Tahoma" Font-Size="8pt"
                                Width="130px">
                            </dx:ASPxLabel>
                        </td>
                    </tr>
                    <tr>
                        <td valign="top" width="140px" align="center">
                            <dx:ASPxComboBox ID="cboPartNo2" runat="server" ClientInstanceName="cboPartNo2" Width="140px"
                                Font-Size="8pt" Font-Names="Tahoma" TextFormatString="{0}" TabIndex="9">
                                <ClientSideEvents SelectedIndexChanged="function(s, e) {
	                                txtPartNo2.SetText(cboPartNo2.GetSelectedItem().GetColumnText(1));	                                
	                                lblInfo.SetText('');	
                                }" />
                                <LoadingPanelStyle ImageSpacing="5px">
                                </LoadingPanelStyle>
                            </dx:ASPxComboBox>
                        </td>
                       <td valign="top" width="210px" align="center">
                            <dx:ASPxTextBox ID="txtPartNo2" runat="server" Width="210px" Height="20px" ClientInstanceName="txtPartNo2"
                                Font-Names="Tahoma" Font-Size="8pt" MaxLength="50" BackColor="#CCCCCC" ReadOnly="True">
                            </dx:ASPxTextBox>
                        </td>
                        <td valign="top" width="110px">
                            <dx:ASPxComboBox ID="cboSupplier2" runat="server" ClientInstanceName="cboSupplier2"
                                Width="110px" Font-Size="8pt" Font-Names="Tahoma" TextFormatString="{0}" TabIndex="13">
                                <ClientSideEvents SelectedIndexChanged="function(s, e) {
	                                txtSupplier2.SetText(cboSupplier2.GetSelectedItem().GetColumnText(1));	                                
	                                lblInfo.SetText('');	
                                }" />
                                <LoadingPanelStyle ImageSpacing="5px">
                                </LoadingPanelStyle>
                            </dx:ASPxComboBox>
                        </td>
                        <td valign="top" width="210px">
                            <dx:ASPxTextBox ID="txtSupplier2" runat="server" Width="210px" Height="20px" ClientInstanceName="txtSupplier2"
                                Font-Names="Tahoma" Font-Size="8pt" MaxLength="50" BackColor="#CCCCCC" ReadOnly="True">
                            </dx:ASPxTextBox>
                        </td>
                        <td valign="top" width="110px">
                            <dx:ASPxDateEdit ID="dt4" runat="server" ClientInstanceName="dt4" Height="21px" Width="110px"
                                EditFormatString="dd MMM yyyy" Font-Names="Tahoma" Font-Size="8pt" TabIndex="10">
                            </dx:ASPxDateEdit>
                        </td>
                        <td valign="top" width="110px">
                            <dx:ASPxDateEdit ID="dt5" runat="server" ClientInstanceName="dt5" EditFormatString="dd MMM yyyy"
                                Font-Names="Tahoma" Font-Size="8pt" Width="110px" TabIndex="11">
                            </dx:ASPxDateEdit>
                        </td>
                        <td valign="top" width="110px" align="left">
                            <dx:ASPxDateEdit ID="dt6" runat="server" ClientInstanceName="dt6" EditFormatString="dd MMM yyyy"
                                Font-Names="Tahoma" Font-Size="8pt" Width="110px" TabIndex="12">
                            </dx:ASPxDateEdit>
                        </td>
                        <td valign="top" width="50px">
                            <dx:ASPxComboBox ID="CboCurrency" runat="server" Width="50px" ClientInstanceName="CboCurrency"
                                ValueType="System.String" TextFormatString="{1)" TabIndex="14">
                            </dx:ASPxComboBox>
                        </td>
                        <td valign="top" width="130px">
                            <dx:ASPxTextBox ID="TxtPrice" runat="server" ClientInstanceName="TxtPrice" Width="130px"
                                HorizontalAlign="Right" MaxLength="15" DisplayFormatString="{0:n4}" onkeypress="return numbersonly(event)"
                                TabIndex="15">
                            </dx:ASPxTextBox>
                        </td>
                    </tr>
                </table>
            </td>
        </tr>
        <tr>
            <td height="50">
                <!-- INPUT AREA -->
                <table id="Table2" style="border-width: 1pt thin thin thin; border-style: ridge; border-color: #9598A1;
                    width: 100%; height: 50px;">
                    <tr>
                        <td valign="top" bgcolor="#FFD2A6" width="140px" align="center">
                            <dx:ASPxLabel ID="ASPxLabel9" runat="server" Text="PACKING GROUP" Font-Names="Tahoma"
                                Font-Size="8pt" Width="140px" Height="16px">
                            </dx:ASPxLabel>
                        </td>
                        <td valign="top" bgcolor="#FFD2A6" align="center" width="210px">
                            <dx:ASPxLabel ID="ASPxLabel10" runat="server" Text="PACKING DESCRIPTION" Font-Names="Tahoma"
                                Font-Size="8pt" Width="210px">
                            </dx:ASPxLabel>
                        </td>
                        <td valign="top" bgcolor="#FFD2A6" align="center" width="110px">
                            <dx:ASPxLabel ID="ASPxLabel11" runat="server" Text="PRICE GROUP" Font-Names="Tahoma"
                                Font-Size="8pt" Width="110px">
                            </dx:ASPxLabel>
                        </td>
                        <td valign="top" bgcolor="#FFD2A6" width="210px" align="center">
                            <dx:ASPxLabel ID="ASPxLabel12" runat="server" Text="PRICE DECRIPTION" Font-Names="Tahoma"
                                Font-Size="8pt" Width="210px">
                            </dx:ASPxLabel>
                        </td>
                        <td valign="top" bgcolor="#FFD2A6" width="110px" align="center">
                            <dx:ASPxLabel ID="ASPxLabel13" runat="server" Text="AFFILIATE CODE" Font-Names="Tahoma"
                                Font-Size="8pt" Width="110px">
                            </dx:ASPxLabel>
                        </td>
                        <td valign="top" bgcolor="#FFD2A6" width="225px" align="center">
                            <dx:ASPxLabel ID="ASPxLabel14" runat="server" Text="AFFILIATE NAME" Font-Names="Tahoma"
                                Font-Size="8pt" Width="225px">
                            </dx:ASPxLabel>
                        </td>
                        <td valign="top" bgcolor="#FFD2A6" width="50px" align="center">
                            <dx:ASPxLabel ID="ASPxLabel16" runat="server" Text=" " Font-Names="Tahoma" Font-Size="8pt"
                                Width="50px">
                            </dx:ASPxLabel>
                        </td>
                        <td valign="top" bgcolor="#FFD2A6" width="130px" align="center">
                            <dx:ASPxLabel ID="ASPxLabel17" runat="server" Text=" " Font-Names="Tahoma" Font-Size="8pt"
                                Width="130px">
                            </dx:ASPxLabel>
                        </td>
                    </tr>
                    <tr>
                        <td valign="top" width="140px" align="center">
                            <dx:ASPxComboBox ID="cboPacking" runat="server" ClientInstanceName="cboPacking" Width="140px"
                                Font-Size="8pt" Font-Names="Tahoma" TextFormatString="{0}" TabIndex="9">
                                <ClientSideEvents SelectedIndexChanged="function(s, e) {
	                                txtPacking.SetText(cboPacking.GetSelectedItem().GetColumnText(1));	                                
	                                lblInfo.SetText('');	
                                }" />
                                <LoadingPanelStyle ImageSpacing="5px">
                                </LoadingPanelStyle>
                            </dx:ASPxComboBox>
                        </td>
                       <td valign="top" width="210px" align="center">
                            <dx:ASPxTextBox ID="txtPacking" runat="server" Width="210px" Height="20px" ClientInstanceName="txtPacking"
                                Font-Names="Tahoma" Font-Size="8pt" MaxLength="50" BackColor="#CCCCCC" ReadOnly="True">
                            </dx:ASPxTextBox>
                        </td>
                        <td valign="top" width="110px">
                            <dx:ASPxComboBox ID="cboPriceCls" runat="server" ClientInstanceName="cboPriceCls"
                                Width="110px" Font-Size="8pt" Font-Names="Tahoma" TextFormatString="{0}" 
                                TabIndex="13">
                                <ClientSideEvents SelectedIndexChanged="function(s, e) {
	                                txtPriceCls.SetText(cboPriceCls.GetSelectedItem().GetColumnText(1));	                                
	                                lblInfo.SetText('');	
                                }" />
                                <LoadingPanelStyle ImageSpacing="5px">
                                </LoadingPanelStyle>
                            </dx:ASPxComboBox>
                        </td>
                        <td valign="top" width="210px">
                            <dx:ASPxTextBox ID="txtPriceCls" runat="server" Width="210px" Height="20px" ClientInstanceName="txtPriceCls"
                                Font-Names="Tahoma" Font-Size="8pt" MaxLength="50" BackColor="#CCCCCC" 
                                ReadOnly="True">
                            </dx:ASPxTextBox>
                        </td>                        
                        <td valign="top" width="110px">
                            <dx:ASPxComboBox ID="cboDeliveryID" runat="server" ClientInstanceName="cboDeliveryID"
                                Width="110px" Font-Size="8pt" Font-Names="Tahoma" TextFormatString="{0}" 
                                TabIndex="13">
                                <ClientSideEvents SelectedIndexChanged="function(s, e) {
	                                txtDeliveryDesc.SetText(cboDeliveryID.GetSelectedItem().GetColumnText(1));	                                
	                                lblInfo.SetText('');	
                                }" />
                                <LoadingPanelStyle ImageSpacing="5px">
                                </LoadingPanelStyle>
                            </dx:ASPxComboBox>
                        </td>
                        <td valign="top" width="225px">
                            <dx:ASPxTextBox ID="txtDeliveryDesc" runat="server" Width="225px" Height="20px" ClientInstanceName="txtDeliveryDesc"
                                Font-Names="Tahoma" Font-Size="8pt" MaxLength="50" BackColor="#CCCCCC" 
                                ReadOnly="True">
                            </dx:ASPxTextBox>
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
                <dx:ASPxButton ID="btnSubMenu" runat="server" Text="SUB MENU" Font-Names="Tahoma"
                    Width="85px" Font-Size="8pt" TabIndex="20" ClientInstanceName="btnSubMenu">
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
            <td valign="top" align="right" style="width: 50px;">
                <dx:ASPxButton ID="btnClear" ClientInstanceName="btnClear" runat="server" Text="CLEAR"
                    Font-Names="Tahoma" Width="90px" AutoPostBack="False" Font-Size="8pt" TabIndex="19">
                    <ClientSideEvents Click="function(s, e) {
                        dt1.SetDate(new Date());
                        dt2.SetDate(new Date());
                        dt3.SetDate(new Date());
                        dt4.SetDate(new Date());
                        dt5.SetDate(new Date());
                        dt6.SetDate(new Date());

                        checkbox1.SetChecked(false);
                        checkbox2.SetChecked(false);
                        checkbox3.SetChecked(false);
        
                        cboPartNo.SetText('== ALL ==');
                        txtPartNo.SetText('== ALL ==');
                        cboSupplier.SetText('== ALL ==');
                        txtSupplier.SetText('== ALL ==');
                        cboPartNo2.SetText('');
                        txtPartNo2.SetText('');
                        cboSupplier2.SetText('');
                        txtSupplier2.SetText('');
                        CboCurrency.SetText('');
                        TxtPrice.SetText('');
                        cboPacking.SetText('');
                        txtPacking.SetText('');
                        cboPriceCls.SetText('');
                        txtPriceCls.SetText('');
                        cboDeliveryID.SetText('');
                        txtDeliveryDesc.SetText('');

                        cboSupplier2.GetInputElement().setAttribute('style', 'background:#FFFFFF;foreground:#FFFFFF;');
                        cboSupplier2.GetInputElement().readOnly = false;
                        cboSupplier2.SetEnabled(true);
            
                        cboPartNo2.GetInputElement().setAttribute('style', 'background:#FFFFFF;foreground:#FFFFFF;');
                        cboPartNo2.GetInputElement().readOnly = false;
                        cboPartNo2.SetEnabled(true);

                        dt4.GetInputElement().setAttribute('style', 'background:#FFFFFF;foreground:#FFFFFF;');
                        dt4.GetInputElement().readOnly = false;
                        dt4.SetEnabled(true);

                        CboCurrency.GetInputElement().setAttribute('style', 'background:#FFFFFF;foreground:#FFFFFF;');
                        CboCurrency.GetInputElement().readOnly = false;
                        CboCurrency.SetEnabled(true);

                        grid.PerformCallback('kosong');                        
                        lblInfo.SetText('');
                    }" />
                </dx:ASPxButton>
            </td>
            <td valign="top" align="right" style="width: 80px;">
                <dx:ASPxButton ID="btnDelete" runat="server" Text="DELETE" Font-Names="Tahoma" Width="80px"
                    AutoPostBack="False" Font-Size="8pt" TabIndex="18" ClientInstanceName="btnDelete">
                    <ClientSideEvents Click="function(s, e) {
                        
                        up_delete();

                        cboPartNo2.SetText('');
                        txtPartNo2.SetText('');
                        cboSupplier2.SetText('');
                        txtSupplier2.SetText('');
                        CboCurrency.SetText('');
                        TxtPrice.SetText('');
                        dt1.SetDate(new Date());
                        dt2.SetDate(new Date());
                        dt3.SetDate(new Date());
                        dt4.SetDate(new Date());
                        dt5.SetDate(new Date());
                        dt6.SetDate(new Date());
                        cboPacking.SetText('');
                        txtPacking.SetText('');
                        cboPriceCls.SetText('');
                        txtPriceCls.SetText('');
                        cboDeliveryID.SetText('');
                        txtDeliveryDesc.SetText('');

                        lblInfo.SetText('');

                        grid.PerformCallback('loadaftersubmit');

                                                cboSupplier2.GetInputElement().setAttribute('style', 'background:#FFFFFF;foreground:#FFFFFF;');
                        cboSupplier2.GetInputElement().readOnly = false;
                        cboSupplier2.SetEnabled(true);
            
                        cboPartNo2.GetInputElement().setAttribute('style', 'background:#FFFFFF;foreground:#FFFFFF;');
                        cboPartNo2.GetInputElement().readOnly = false;
                        cboPartNo2.SetEnabled(true);

                        dt4.GetInputElement().setAttribute('style', 'background:#FFFFFF;foreground:#FFFFFF;');
                        dt4.GetInputElement().readOnly = false;
                        dt4.SetEnabled(true);

                        CboCurrency.GetInputElement().setAttribute('style', 'background:#FFFFFF;foreground:#FFFFFF;');
                        CboCurrency.GetInputElement().readOnly = false;
                        CboCurrency.SetEnabled(true);
                    }" />
                </dx:ASPxButton>
            </td>
            <td valign="top" align="right" style="width: 50px;">
                <dx:ASPxButton ID="btnSubmit" runat="server" Text="SAVE" Font-Names="Tahoma" Width="90px"
                    AutoPostBack="False" Font-Size="8pt" TabIndex="17" ClientInstanceName="btnSubmit">
                    <ClientSideEvents Click="function(s, e) {
                        grid.SetFocusedRowIndex(-1);
                        validasubmit();
                        up_Insert();
                                             
                        cboPartNo2.SetText('');
                        txtPartNo2.SetText('');
                        cboSupplier2.SetText('');
                        txtSupplier2.SetText('');
                        CboCurrency.SetText('');
                        TxtPrice.SetText('');
                        dt1.SetDate(new Date());
                        dt2.SetDate(new Date());
                        dt3.SetDate(new Date());
                        dt4.SetDate(new Date());
                        dt5.SetDate(new Date());
                        dt6.SetDate(new Date());
                        cboPacking.SetText('');
                        txtPacking.SetText('');
                        cboPriceCls.SetText('');
                        txtPriceCls.SetText('');
                        cboDeliveryID.SetText('');
                        txtDeliveryDesc.SetText('');

                        grid.PerformCallback('loadaftersubmit');

                        cboSupplier2.GetInputElement().setAttribute('style', 'background:#FFFFFF;foreground:#FFFFFF;');
                        cboSupplier2.GetInputElement().readOnly = false;
                        cboSupplier2.SetEnabled(true);
            
                        cboPartNo2.GetInputElement().setAttribute('style', 'background:#FFFFFF;foreground:#FFFFFF;');
                        cboPartNo2.GetInputElement().readOnly = false;
                        cboPartNo2.SetEnabled(true);

                        dt4.GetInputElement().setAttribute('style', 'background:#FFFFFF;foreground:#FFFFFF;');
                        dt4.GetInputElement().readOnly = false;
                        dt4.SetEnabled(true);

                        CboCurrency.GetInputElement().setAttribute('style', 'background:#FFFFFF;foreground:#FFFFFF;');
                        CboCurrency.GetInputElement().readOnly = false;
                        CboCurrency.SetEnabled(true);
                
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
    <dx:ASPxHiddenField ID="HF" runat="server" ClientInstanceName="HF">
    </dx:ASPxHiddenField>
</asp:Content>
