<%@ Page Title="" Language="vb" AutoEventWireup="false" MasterPageFile="~/SiteContent.Master" CodeBehind="PartMapping.aspx.vb" Inherits="PASISystem.PartMapping" %>
<%@ Register Assembly="DevExpress.Web.v14.1, Version=14.1.6.0, Culture=neutral, PublicKeyToken=b88d1754d700e49a" Namespace="DevExpress.Web.ASPxGridView.Export" TagPrefix="dx" %>
<%@ Register Assembly="DevExpress.Web.v14.1, Version=14.1.6.0, Culture=neutral, PublicKeyToken=b88d1754d700e49a" Namespace="DevExpress.Web.ASPxGridView" TagPrefix="dx" %>
<%@ Register Assembly="DevExpress.Web.v14.1, Version=14.1.6.0, Culture=neutral, PublicKeyToken=b88d1754d700e49a" Namespace="DevExpress.Web.ASPxEditors" TagPrefix="dx" %>
<%@ Register Assembly="DevExpress.Web.v14.1, Version=14.1.6.0, Culture=neutral, PublicKeyToken=b88d1754d700e49a" Namespace="DevExpress.Web.ASPxGlobalEvents" TagPrefix="dx" %>
<%@ Register Assembly="DevExpress.Web.v14.1, Version=14.1.6.0, Culture=neutral, PublicKeyToken=b88d1754d700e49a" Namespace="DevExpress.Web.ASPxCallback" TagPrefix="dx" %>
<%@ Register Assembly="DevExpress.Web.v14.1, Version=14.1.6.0, Culture=neutral, PublicKeyToken=b88d1754d700e49a" Namespace="DevExpress.Web.ASPxHiddenField" TagPrefix="dx" %>

<asp:Content ID="Content1" ContentPlaceHolderID="head" runat="server">
    <style type="text/css">.dxeHLC, .dxeHC, .dxeHFC {display: none;}</style>
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
            height = height - (height * 58 / 100)
            grid.SetHeight(height);
        }

        function OnBatchEditStartEditing(s, e) {
            currentColumnName = e.focusedColumn.fieldName;

            if (currentColumnName == "NoUrut" || currentColumnName == "SupplierID" || currentColumnName == "SupplierName" ||
            currentColumnName == "PartNo" || currentColumnName == "PartName" || currentColumnName == "AffiliateID" || currentColumnName == "AffiliateName" ||
            currentColumnName == "Quota" || currentColumnName == "LocationID" || currentColumnName == "PackingCls" || currentColumnName == "PackingDesc" ||
            currentColumnName == "MOQ" || currentColumnName == "QtyBox" || currentColumnName == "BoxPallet" || currentColumnName == "NetWeight" || currentColumnName == "GrossWeight" ||
            currentColumnName == "Length" || currentColumnName == "Width" || currentColumnName == "Height") 
            {
                e.cancel = true;
            }

            currentEditableVisibleIndex = e.visibleIndex;
        }

        function OnGridFocusedRowChanged() {
            grid.GetRowValues(grid.GetFocusedRowIndex(), "PartNo;PartName;AffiliateID;AffiliateName;SupplierID;SupplierName;Quota;LocationID;PackingCls;PackingDesc;MOQ;QtyBox;BoxPallet;NetWeight;GrossWeight;Length;Width;Height;DeleteCls", OnGetRowValues);
        }

        function OnGetRowValues(values) {
            if (values[0] != "" && values[0] != null && values[0] != "null") {

                txtPartnoDetail.SetText(values[0]);
                txtPartNo2.SetText(values[1]);
                cboAffiliate2.SetText(values[2]);
                txtAffiliate2.SetText(values[3]);
                cboSupplier2.SetText(values[4]);
                txtSupplier2.SetText(values[5]);                
                txtQuota.SetText(values[6]);
                txtLocation.SetText(values[7]);
                cboPacking.SetText(values[8]);
                txtPacking.SetText(values[9]);
                txtMOQ.SetText(values[10]);
                txtQtyBox.SetText(values[11]);
                txtBoxPallet.SetText(values[12]);
                txtNetWeight.SetText(values[13]);
                txtGrossWeight.SetText(values[14]);
                txtLength.SetText(values[15]);
                txtWidth.SetText(values[16]);
                txtHeight.SetText(values[17]);
                var vDeleteCls = values[18];

                lblInfo.SetText('');

                txtPartnoDetail.GetInputElement().setAttribute('style', 'background:#CCCCCC;foreground:#CCCCCC;');
                txtPartnoDetail.GetInputElement().readOnly = true;
                txtPartnoDetail.SetEnabled(false);

                cboAffiliate2.GetInputElement().setAttribute('style', 'background:#CCCCCC;foreground:#CCCCCC;');
                cboAffiliate2.GetInputElement().readOnly = true;
                cboAffiliate2.SetEnabled(false);

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
            if (txtPartnoDetail.GetText() == "" || txtPartNo2.GetText() == "") {
                lblInfo.GetMainElement().style.color = 'Red';
                lblInfo.SetText("[6011] Please Input Part No. a valid first!");
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

            var pPartNo = txtPartnoDetail.GetText();
            var pAffiliateID = cboAffiliate2.GetText();
            var pSupplierID = cboSupplier2.GetText();

            grid.PerformCallback('delete|' + pPartNo + '|' + pAffiliateID + '|' + pSupplierID);
        }

        function readonly() {
            txtPartID.GetInputElement().setAttribute('style', 'background:#CCCCCC;foreground:#CCCCCC;');
            txtPartID.GetInputElement().readOnly = true;
            lblInfo.SetText('');
        }

        function validasubmit() {
        debugger
            lblInfo.GetMainElement().style.color = 'Red';
            if (txtPartnoDetail.GetText() == "" || txtPartNo2.GetText() == "") {
                lblInfo.SetText("[6011] Please Input Part No. a valid first!");
                txtPartnoDetail.Focus();
                e.ProcessOnServer = false;
                return false;
            }

            if (cboAffiliate2.GetText() == "") {
                lblInfo.SetText("[6011] Please Select Affiliate first!");
                cboAffiliate2.Focus();
                e.ProcessOnServer = false;
                return false;
            }

            if (cboSupplier2.GetText() == "") {
                lblInfo.SetText("[6011] Please Select Supplier first!");
                cboSupplier2.Focus();
                e.ProcessOnServer = false;
                return false;
            }
        }

        function up_Insert() {
            var pIsUpdate = '';
            var pPartID = txtPartnoDetail.GetText();
            var pAffiliateID = cboAffiliate2.GetSelectedItem().GetColumnText(0);
            var pSupplierID = cboSupplier2.GetSelectedItem().GetColumnText(0);
            var pQuota = txtQuota.GetText();
            var pLocation = txtLocation.GetText();
            var pPackingID = cboPacking.GetSelectedItem().GetColumnText(0);
            var pMOQ = txtMOQ.GetText();
            var pQtyBox = txtQtyBox.GetText();
            var pBoxPallet = txtBoxPallet.GetText();
            var pNetWeight = txtNetWeight.GetText();
            var pGrossWeight = txtGrossWeight.GetText();
            var pLength = txtLength.GetText();
            var pWidth = txtWidth.GetText();
            var pHeight = txtHeight.GetText();

            grid.PerformCallback('save|' + pIsUpdate + '|' + pPartID + '|' + pAffiliateID + '|' + pSupplierID + '|' + pQuota + '|' + pLocation + '|' + pPackingID + '|' + pMOQ + '|' + pQtyBox + '|' + pBoxPallet + '|' + pNetWeight + '|' + pGrossWeight + '|' + pLength + '|' + pWidth + '|' + pHeight);
        }

        function numbersonly(e) {
            var unicode = e.charCode ? e.charCode : e.keyCode
            if (unicode != 8) { //if the key isn't the backspace key (which we should allow)
                if (unicode < 45 || unicode > 57) //if not a number
                    return false //disable key press
            }
        }

        function partChanged(s, e) {
            PartNoCallBack.PerformCallback('Load|' + txtPartnoDetail.GetText());
        }

        function EndPartNoCallBack(s, e) {
            var pPartno = s.cpPartno;
            var pPartnames = s.cpPartnosNames;

            txtPartnoDetail.SetText(pPartno);
            txtPartNo2.SetText(pPartnames);
        }
    </script>
</asp:Content>

<asp:Content ID="Content2" ContentPlaceHolderID="Content" runat="server">
    <table style="width:100%;">
        <tr>
            <td>
                <table style="border-width: 1pt thin thin thin; border-style: ridge; border-color: #9598A1; width: 100%;">
                    <tr>
                        <td>
                            <table>
                                <tr>
                                    <td style="width: 5px;">
                                    </td>
                                    <td align="left" valign="middle" height="25px" width="120px">
                                        <dx:ASPxLabel ID="ASPxLabel4" runat="server" Text="PART NO." Font-Names="Verdana"
                                            Font-Size="8pt">
                                        </dx:ASPxLabel>
                                    </td>
                                    <td align="left" valign="middle" height="25px" width="150px">
                                        <dx:ASPxTextBox ID="txtPartNo" runat="server" Width="100%" Height="20px" ClientInstanceName="txtPartNo"
                                            Font-Names="Verdana" Font-Size="8pt" MaxLength="100">
                                            <ClientSideEvents LostFocus="function(s, e) {	                                            
                                                grid.PerformCallback('kosong');	                                
	                                            lblInfo.SetText('');	
                                            }" />
                                        </dx:ASPxTextBox>
                                    </td>
                                    <td align="left" valign="middle" style="height: 25px; width: 200px;">
                                    </td>
                                    <td align="left" valign="middle" height="25px" width="150px">
                                    </td>
                                </tr>
                                <tr>
                                    <td style="width: 5px;">
                                    </td>
                                    <td align="left" valign="middle" height="25px" width="80px">
                                        <dx:ASPxLabel ID="ASPxLabel2" runat="server" Text="AFFILIATE CODE" Font-Names="Verdana"
                                            Font-Size="8pt">
                                        </dx:ASPxLabel>
                                    </td>
                                    <td align="left" valign="middle" height="25px" width="150px">
                                        <dx:ASPxComboBox ID="cboAffiliate" runat="server" ClientInstanceName="cboAffiliate"
                                            Width="100%" Font-Size="8pt" Font-Names="Verdana" TextFormatString="{0}">
                                            <ClientSideEvents SelectedIndexChanged="function(s, e) {
	                                            txtAffiliate.SetText(cboAffiliate.GetSelectedItem().GetColumnText(1));
                                                grid.PerformCallback('kosong');	                                
	                                            lblInfo.SetText('');	
                                            }" />
                                            <LoadingPanelStyle ImageSpacing="5px">
                                            </LoadingPanelStyle>
                                        </dx:ASPxComboBox>
                                    </td>
                                    <td align="left" valign="middle" height="25px" width="200px">
                                        <dx:ASPxTextBox ID="txtAffiliate" runat="server" Width="100%" Height="20px" ClientInstanceName="txtAffiliate"
                                            Font-Names="Verdana" Font-Size="8pt" MaxLength="100" BackColor="#CCCCCC" ReadOnly="True">
                                        </dx:ASPxTextBox>
                                    </td>
                                    <td align="left" valign="middle" height="25px" width="150px">
                                    </td>
                                </tr>
                                <tr>
                                    <td style="width: 5px;">
                                    </td>
                                    <td align="left" valign="middle" height="25px" width="120px">
                                        <dx:ASPxLabel ID="ASPxLabel1" runat="server" Text="SUPPLIER CODE" Font-Names="Verdana"
                                            Font-Size="8pt">
                                        </dx:ASPxLabel>
                                    </td>
                                    <td align="left" valign="middle" height="25px" width="150px">
                                        <dx:ASPxComboBox ID="cboSupplier" runat="server" ClientInstanceName="cboSupplier"
                                            Width="100%" Font-Size="8pt" Font-Names="Verdana" TextFormatString="{0}">
                                            <ClientSideEvents SelectedIndexChanged="function(s, e) {
	                                            txtSupplier.SetText(cboSupplier.GetSelectedItem().GetColumnText(1));
                                                grid.PerformCallback('kosong');	                                
	                                            lblInfo.SetText('');	
                                            }" />
                                            <LoadingPanelStyle ImageSpacing="5px">
                                            </LoadingPanelStyle>
                                        </dx:ASPxComboBox>
                                    </td>
                                    <td align="left" valign="middle" style="height: 25px; width: 200px;">
                                        <dx:ASPxTextBox ID="txtSupplier" runat="server" Width="100%" Height="20px" ClientInstanceName="txtSupplier"
                                            Font-Names="Verdana" Font-Size="8pt" MaxLength="100" BackColor="#CCCCCC" ReadOnly="True">
                                        </dx:ASPxTextBox>
                                    </td>
                                    <td align="right" valign="middle" style="height: 25px; width: 100px;">
                                        <table style="width: 100%;">
                                            <tr>
                                                <td align="right" valign="middle" style="height: 25px; width: 90px;">
                                                    <dx:ASPxButton ID="btnRefresh" runat="server" Text="SEARCH" Font-Names="Verdana"
                                                        Width="85px" AutoPostBack="False" Font-Size="8pt">
                                                        <ClientSideEvents Click="function(s, e) {
                                                            grid.PerformCallback('load');
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
    </table>

    <div style="height: 1px;"></div>
    
    <table style="width: 100%; height: 15px;">
        <tr>
            <td>
                <table id="tblMsg" style="border-width: thin; border-style: inset hidden ridge hidden; border-color: #9598A1; width: 100%; height: 15px;">
                    <tr>
                        <td align="left" valign="middle" height="15px">
                            <dx:ASPxLabel ID="lblInfo" runat="server" Text="[lblinfo]" Font-Names="Verdana" ClientInstanceName="lblInfo" Font-Bold="True" Font-Italic="True" Font-Size="8pt">
                            </dx:ASPxLabel>
                        </td>
                    </tr>
                </table>
            </td>
        </tr>
    </table>
    
    <div style="height: 1px;"></div>
    
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
                <dx:ASPxGridView ID="grid" runat="server" Width="100%" Font-Names="Verdana" KeyFieldName="PartNo;AffiliateID;SupplierID" AutoGenerateColumns="False" ClientInstanceName="grid" Font-Size="8pt">
                    <ClientSideEvents 
                        Init="OnInit" 
                        FocusedRowChanged="function(s, e){
	                        OnGridFocusedRowChanged();
                        }"
                        EndCallback="function(s, e){
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
                        }" 
                        RowClick="function(s, e) {
	                        lblInfo.SetText('');
                        }" 
                        BatchEditStartEditing="OnBatchEditStartEditing" 
                    />
                    <Columns>
                        <dx:GridViewDataTextColumn VisibleIndex="0" Caption="NO" FieldName="NoUrut" Width="50px"
                            HeaderStyle-HorizontalAlign="Center">                            
                        </dx:GridViewDataTextColumn>
                        <dx:GridViewDataTextColumn VisibleIndex="1" Caption="PART NO." FieldName="PartNo"
                            Width="100px" HeaderStyle-HorizontalAlign="Center">                            
                        </dx:GridViewDataTextColumn>
                        <dx:GridViewDataTextColumn VisibleIndex="2" Caption="PART NAME" FieldName="PartName"
                            Width="210px" HeaderStyle-HorizontalAlign="Center">                            
                        </dx:GridViewDataTextColumn>
                        <dx:GridViewDataTextColumn VisibleIndex="3" Caption="AFFILIATE CODE" FieldName="AffiliateID"
                            Width="80px" HeaderStyle-HorizontalAlign="Center">                            
                        </dx:GridViewDataTextColumn>
                        <dx:GridViewDataTextColumn VisibleIndex="4" Caption="AFFILIATE NAME" FieldName="AffiliateName"
                            Width="0px" HeaderStyle-HorizontalAlign="Center">                            
                        </dx:GridViewDataTextColumn>
                        <dx:GridViewDataTextColumn VisibleIndex="5" Caption="SUPPLIER CODE" FieldName="SupplierID"
                            Width="80px" HeaderStyle-HorizontalAlign="Center">                            
                        </dx:GridViewDataTextColumn>
                        <dx:GridViewDataTextColumn VisibleIndex="6" Caption="SUPPLIER NAME" FieldName="SupplierName"
                            Width="0px" HeaderStyle-HorizontalAlign="Center">                            
                        </dx:GridViewDataTextColumn>
                        <dx:GridViewDataTextColumn VisibleIndex="7" Caption="QUOTA (%)" FieldName="Quota"
                            Width="50px" HeaderStyle-HorizontalAlign="Center">                            
                        </dx:GridViewDataTextColumn>
                        <dx:GridViewDataTextColumn VisibleIndex="8" Caption="LOCATION" FieldName="LocationID"
                            Width="100px" HeaderStyle-HorizontalAlign="Center">                            
                        </dx:GridViewDataTextColumn>
                        <dx:GridViewDataTextColumn Caption="PACKING GROUP" FieldName="PackingCls" 
                            VisibleIndex="9" Width="0px">
                        </dx:GridViewDataTextColumn>
                        <dx:GridViewDataTextColumn Caption="PACKING STYLE" 
                            FieldName="PackingDesc" VisibleIndex="10" Width="200px">
                        </dx:GridViewDataTextColumn>
                        <dx:GridViewDataTextColumn Caption="MOQ" FieldName="MOQ" VisibleIndex="11" 
                            Width="75px">                            
                        </dx:GridViewDataTextColumn>
                        <dx:GridViewDataTextColumn Caption="QTY/BOX" FieldName="QtyBox" 
                            VisibleIndex="12" Width="75px">                            
                        </dx:GridViewDataTextColumn>
                        <dx:GridViewDataTextColumn Caption="BOX/PALLET" FieldName="BoxPallet" 
                            VisibleIndex="13" Width="100px">                            
                        </dx:GridViewDataTextColumn>
                        <dx:GridViewDataTextColumn Caption="NET WEIGHT (GR)" FieldName="NetWeight" 
                            VisibleIndex="14" Width="75px">                            
                        </dx:GridViewDataTextColumn>
                        <dx:GridViewDataTextColumn Caption="GROSS WEIGHT (GR)" FieldName="GrossWeight" 
                            VisibleIndex="15" Width="75px">                            
                        </dx:GridViewDataTextColumn>
                        <dx:GridViewDataTextColumn Caption="LENGTH (MM)" FieldName="Length" 
                            VisibleIndex="16" Width="75px">                            
                        </dx:GridViewDataTextColumn>
                        <dx:GridViewDataTextColumn Caption="WIDTH (MM)" FieldName="Width" 
                            VisibleIndex="17" 
                            Width="75px">                            
                        </dx:GridViewDataTextColumn>
                        <dx:GridViewDataTextColumn Caption="HEIGHT (MM)" FieldName="Height" 
                            VisibleIndex="18" Width="75px">                            
                        </dx:GridViewDataTextColumn>
                        <dx:GridViewDataTextColumn VisibleIndex="19" Caption="REGISTER DATE" 
                            FieldName="EntryDate" Width="150px" HeaderStyle-HorizontalAlign="Center">
                        </dx:GridViewDataTextColumn>
                         <dx:GridViewDataTextColumn VisibleIndex="20" Caption="REGISTER USER" 
                            FieldName="EntryUser" Width="100px" HeaderStyle-HorizontalAlign="Center">
                        </dx:GridViewDataTextColumn>
                         <dx:GridViewDataTextColumn VisibleIndex="21" Caption="UPDATE DATE" 
                            FieldName="UpdateDate" Width="150px" HeaderStyle-HorizontalAlign="Center">
                        </dx:GridViewDataTextColumn>
                         <dx:GridViewDataTextColumn VisibleIndex="22" Caption="UPDATE USER" 
                            FieldName="UpdateUser" Width="100px" HeaderStyle-HorizontalAlign="Center">
                        </dx:GridViewDataTextColumn>
                        <dx:GridViewDataTextColumn VisibleIndex="23" Caption="DeleteCls" 
                            FieldName="DeleteCls" Width="0px" HeaderStyle-HorizontalAlign="Center">
                        </dx:GridViewDataTextColumn>
                    </Columns>
                    <SettingsBehavior AllowSelectByRowClick="True" AllowSort="False" ColumnResizeMode="Control"
                        EnableRowHotTrack="True" />
                    <SettingsPager Visible="False" PageSize="15" Position="Top">
                        <Summary Text="Page {0} of {1} [{2} record(s)]" AllPagesText="Page {0} of {1} " />
                        <Summary AllPagesText="Page {0} of {1} " Text="Page {0} of {1} [{2} record(s)]">
                        </Summary>
                    </SettingsPager>
                    <SettingsEditing Mode="Batch" NewItemRowPosition="Bottom">
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

    <div style="height: 1px;"></div>

    <table style="width:100%;">
        <tr>
            <td height="50">
                <!-- INPUT AREA -->
                <table style="border-width: 1pt thin thin thin; border-style: ridge; border-color: #9598A1; width: 100%; height: 25px; background-color: #FFD2A6">
                    <tr>
                        <td valign="top" style="width: 120px;">
                            <dx:ASPxLabel ID="ASPxLabel53" runat="server" Text="PART NO." Font-Names="Verdana"
                                Font-Size="8pt" Width="100%">
                            </dx:ASPxLabel>
                        </td>
                        <td valign="top" style="width: 200px;">
                            <dx:ASPxLabel ID="ASPxLabel65" runat="server" Text="PART NAME" Font-Names="Verdana"
                                Font-Size="8pt" Width="100%">
                            </dx:ASPxLabel>
                        </td>
                        <td valign="top" style="width: 120px;">
                            <dx:ASPxLabel ID="ASPxLabel56" runat="server" Text="AFFILIATE CODE" Font-Names="Verdana"
                                Font-Size="8pt" Width="100%">
                            </dx:ASPxLabel>
                        </td>
                        <td valign="top" style="width: 200px;">
                            <dx:ASPxLabel ID="ASPxLabel57" runat="server" Text="AFFILIATE NAME" Font-Names="Verdana"
                                Font-Size="8pt" Width="100%">
                            </dx:ASPxLabel>
                        </td>
                        <td valign="top" style="width: 120px;">
                            <dx:ASPxLabel ID="ASPxLabel55" runat="server" Text="SUPPLIER CODE" Font-Names="Verdana"
                                Font-Size="8pt" Width="100%">
                            </dx:ASPxLabel>
                        </td>
                        <td valign="top" style="width: 200px;">
                            <dx:ASPxLabel ID="ASPxLabel3" runat="server" Text="SUPPLIER NAME" Font-Names="Verdana"
                                Font-Size="8pt" Width="100%">
                            </dx:ASPxLabel>
                        </td>
                        <td valign="top" style="width: 100px;">
                            <dx:ASPxLabel ID="ASPxLabel5" runat="server" Text="QUOTA (%)" Font-Names="Verdana"
                                Font-Size="8pt" Width="100%">
                            </dx:ASPxLabel>
                        </td>
                        <td valign="top">
                            <dx:ASPxLabel ID="ASPxLabel6" runat="server" Text="LOCATION" Font-Names="Verdana"
                                Font-Size="8pt" Width="100%">
                            </dx:ASPxLabel>
                        </td>
                    </tr>
                </table>
                <table style="border-width: 1pt thin thin thin; border-style: ridge; border-color: #9598A1; width: 100%; height: 35px;">
                    <tr>
                        <td style="width: 120px;">
                            <dx:ASPxTextBox ID="txtPartnoDetail" runat="server" Width="100%" Height="20px" ClientInstanceName="txtPartnoDetail"
                                Font-Names="Verdana" Font-Size="8pt" MaxLength="18">
                                <ClientSideEvents TextChanged="partChanged" />
                            </dx:ASPxTextBox>
                        </td>
                        <td style="width: 200px;">
                            <dx:ASPxTextBox ID="txtPartNo2" runat="server" Width="100%" Height="20px" ClientInstanceName="txtPartNo2"
                                Font-Names="Verdana" Font-Size="8pt" MaxLength="50" BackColor="#CCCCCC" ReadOnly="True">
                            </dx:ASPxTextBox>
                        </td>
                        <td style="width: 120px;">
                            <dx:ASPxComboBox ID="cboAffiliate2" runat="server" ClientInstanceName="cboAffiliate2"
                                Width="100%" Font-Size="8pt" Font-Names="Verdana" TextFormatString="{0}">
                                <ClientSideEvents SelectedIndexChanged="function(s, e) {
	                                txtAffiliate2.SetText(cboAffiliate2.GetSelectedItem().GetColumnText(1));	                                
	                                lblInfo.SetText('');	
                                }" />
                                <LoadingPanelStyle ImageSpacing="5px">
                                </LoadingPanelStyle>
                            </dx:ASPxComboBox>
                        </td>
                        <td style="width: 200px;">
                            <dx:ASPxTextBox ID="txtAffiliate2" runat="server" Width="100%" Height="20px" ClientInstanceName="txtAffiliate2"
                                Font-Names="Verdana" Font-Size="8pt" MaxLength="50" BackColor="#CCCCCC" ReadOnly="True">
                            </dx:ASPxTextBox>
                        </td>
                        <td style="width: 120px;">
                            <dx:ASPxComboBox ID="cboSupplier2" runat="server" ClientInstanceName="cboSupplier2"
                                Width="100%" Font-Size="8pt" Font-Names="Verdana" TextFormatString="{0}">
                                <ClientSideEvents SelectedIndexChanged="function(s, e) {
	                                txtSupplier2.SetText(cboSupplier2.GetSelectedItem().GetColumnText(1));	                                
	                                lblInfo.SetText('');	
                                }" />
                                <LoadingPanelStyle ImageSpacing="5px">
                                </LoadingPanelStyle>
                            </dx:ASPxComboBox>
                        </td>
                        <td style="width: 200px;">
                            <dx:ASPxTextBox ID="txtSupplier2" runat="server" Width="100%" Height="20px" ClientInstanceName="txtSupplier2"
                                Font-Names="Verdana" Font-Size="8pt" MaxLength="50" BackColor="#CCCCCC" ReadOnly="True">
                            </dx:ASPxTextBox>
                        </td>
                        <td style="width: 100px;">
                            <dx:ASPxTextBox ID="txtQuota" runat="server" Width="100%" Height="20px" ClientInstanceName="txtQuota"
                                Font-Names="Verdana" Font-Size="8pt" MaxLength="6" 
                                onkeypress="return numbersonly(event)" HorizontalAlign="Right">
                                <ClientSideEvents TextChanged="function(s, e) {
	                                lblInfo.SetText('');
                                }" />
                            </dx:ASPxTextBox>
                        </td>
                        <td>
                            <dx:ASPxTextBox ID="txtLocation" runat="server" Width="100%" Height="20px" ClientInstanceName="txtLocation"
                                Font-Names="Verdana" Font-Size="8pt" MaxLength="25">
                            </dx:ASPxTextBox>
                        </td>
                    </tr>
                </table>
            </td>
        </tr>
        <tr>
            <td height="50">
                <table style="border-width: 1pt thin thin thin; border-style: ridge; border-color: #9598A1; width: 100%; height: 25px; background-color: #FFD2A6">
                    <tr>
                        <td valign="top" style="width: 120px;">
                            <dx:ASPxLabel ID="ASPxLabel7" runat="server" Text="PACKING GROUP" Font-Names="Verdana"
                                Font-Size="8pt" Width="100%">
                            </dx:ASPxLabel>
                        </td>
                        <td valign="top" style="width: 200px;">
                            <dx:ASPxLabel ID="ASPxLabel8" runat="server" Text="PACKING DESCRIPTION" Font-Names="Verdana"
                                Font-Size="8pt" Width="100%">
                            </dx:ASPxLabel>
                        </td>
                        <td valign="top" style="width: 100px;">
                            <dx:ASPxLabel ID="ASPxLabel9" runat="server" Text="MOQ" Font-Names="Verdana"
                                Font-Size="8pt" Width="100%">
                            </dx:ASPxLabel>
                        </td>
                        <td valign="top" style="width: 100px;">
                            <dx:ASPxLabel ID="ASPxLabel10" runat="server" Text="QTY/BOX" Font-Names="Verdana"
                                Font-Size="8pt" Width="100%">
                            </dx:ASPxLabel>
                        </td>
                        <td valign="top" style="width: 100px;">
                            <dx:ASPxLabel ID="ASPxLabel11" runat="server" Text="BOX/PALLET" Font-Names="Verdana"
                                Font-Size="8pt" Width="100%">
                            </dx:ASPxLabel>
                        </td>
                        <td valign="top" style="width: 100px;">
                            <dx:ASPxLabel ID="ASPxLabel12" runat="server" Text="NET WEIGHT (GR)" Font-Names="Verdana"
                                Font-Size="8pt" Width="100%">
                            </dx:ASPxLabel>
                        </td>
                        <td valign="top" style="width: 100px;">
                            <dx:ASPxLabel ID="ASPxLabel13" runat="server" Text="GROSS WEIGHT (GR)" Font-Names="Verdana"
                                Font-Size="8pt" Width="100%">
                            </dx:ASPxLabel>
                        </td>
                        <td valign="top" style="width: 100px;">
                            <dx:ASPxLabel ID="ASPxLabel15" runat="server" Text="LENGTH (MM)" Font-Names="Verdana"
                                Font-Size="8pt" Width="100%">
                            </dx:ASPxLabel>
                        </td>
                        <td valign="top" style="width: 100px;">
                            <dx:ASPxLabel ID="ASPxLabel14" runat="server" Text="WIDTH (MM)" Font-Names="Verdana"
                                Font-Size="8pt" Width="100%">
                            </dx:ASPxLabel>
                        </td>
                        <td valign="top" style="width: 100px;">
                            <dx:ASPxLabel ID="ASPxLabel16" runat="server" Text="HEIGHT (MM)" Font-Names="Verdana"
                                Font-Size="8pt" Width="100%">
                            </dx:ASPxLabel>
                        </td>
                        <td></td>
                    </tr>
                </table>
                <table style="border-width: 1pt thin thin thin; border-style: ridge; border-color: #9598A1; width: 100%; height: 35px;">
                    <tr>
                        <td style="width: 120px;">
                            <dx:ASPxComboBox ID="cboPacking" runat="server" ClientInstanceName="cboPacking" Width="100%"
                                Font-Size="8pt" Font-Names="Verdana" TextFormatString="{0}">
                                <ClientSideEvents SelectedIndexChanged="function(s, e) {
	                                txtPacking.SetText(cboPacking.GetSelectedItem().GetColumnText(1));	                                
	                                lblInfo.SetText('');	
                                }" />
                                <LoadingPanelStyle ImageSpacing="5px">
                                </LoadingPanelStyle>
                            </dx:ASPxComboBox>
                        </td>
                        <td style="width: 200px;">
                            <dx:ASPxTextBox ID="txtPacking" runat="server" Width="100%" Height="20px" ClientInstanceName="txtPacking"
                                Font-Names="Verdana" Font-Size="8pt" MaxLength="50" BackColor="#CCCCCC" 
                                ReadOnly="True">
                            </dx:ASPxTextBox>
                        </td>
                        <td style="width: 100px;">
                            <dx:ASPxTextBox ID="txtMOQ" runat="server" Width="100%" Height="20px" ClientInstanceName="txtMOQ"
                                Font-Names="Verdana" Font-Size="8pt" MaxLength="18" 
                                onkeypress="return numbersonly(event)" HorizontalAlign="Right">
                                <ClientSideEvents TextChanged="function(s, e) {
	                                lblInfo.SetText('');
                                }" />
                            </dx:ASPxTextBox>
                        </td>
                        <td style="width: 100px;">
                            <dx:ASPxTextBox ID="txtQtyBox" runat="server" Width="100%" Height="20px" ClientInstanceName="txtQtyBox"
                                Font-Names="Verdana" Font-Size="8pt" MaxLength="18" 
                                onkeypress="return numbersonly(event)" HorizontalAlign="Right">
                                <ClientSideEvents TextChanged="function(s, e) {
	                                lblInfo.SetText('');
                                }" />
                            </dx:ASPxTextBox>
                        </td>
                        <td style="width: 100px;">
                            <dx:ASPxTextBox ID="txtBoxPallet" runat="server" Width="100%" Height="20px" ClientInstanceName="txtBoxPallet"
                                Font-Names="Verdana" Font-Size="8pt" MaxLength="18" 
                                onkeypress="return numbersonly(event)" HorizontalAlign="Right">
                                <ClientSideEvents TextChanged="function(s, e) {
	                                lblInfo.SetText('');
                                }" />
                            </dx:ASPxTextBox>
                        </td>
                        <td style="width: 100px;">
                            <dx:ASPxTextBox ID="txtNetWeight" runat="server" Width="100%" Height="20px" ClientInstanceName="txtNetWeight"
                                Font-Names="Verdana" Font-Size="8pt" MaxLength="18" 
                                onkeypress="return numbersonly(event)" HorizontalAlign="Right">
                                <ClientSideEvents TextChanged="function(s, e) {
	                                lblInfo.SetText('');
                                }" />
                            </dx:ASPxTextBox>
                        </td>
                        <td style="width: 100px;">
                            <dx:ASPxTextBox ID="txtGrossWeight" runat="server" Width="100%" Height="20px" ClientInstanceName="txtGrossWeight"
                                Font-Names="Verdana" Font-Size="8pt" MaxLength="18" 
                                onkeypress="return numbersonly(event)" HorizontalAlign="Right">
                                <ClientSideEvents TextChanged="function(s, e) {
	                                lblInfo.SetText('');
                                }" />
                            </dx:ASPxTextBox>
                        </td>
                        <td style="width: 100px;">
                            <dx:ASPxTextBox ID="txtLength" runat="server" Width="100%" Height="20px" ClientInstanceName="txtLength"
                                Font-Names="Verdana" Font-Size="8pt" MaxLength="18" 
                                onkeypress="return numbersonly(event)" HorizontalAlign="Right">
                                <ClientSideEvents TextChanged="function(s, e) {
	                                lblInfo.SetText('');
                                }" />
                            </dx:ASPxTextBox>
                        </td>
                        <td style="width: 100px;">
                            <dx:ASPxTextBox ID="txtWidth" runat="server" Width="100%" Height="20px" ClientInstanceName="txtWidth"
                                Font-Names="Verdana" Font-Size="8pt" MaxLength="18" 
                                onkeypress="return numbersonly(event)" HorizontalAlign="Right">
                                <ClientSideEvents TextChanged="function(s, e) {
	                                lblInfo.SetText('');
                                }" />
                            </dx:ASPxTextBox>
                        </td>
                        <td style="width: 100px;">
                            <dx:ASPxTextBox ID="txtHeight" runat="server" Width="100%" Height="20px" ClientInstanceName="txtHeight"
                                Font-Names="Verdana" Font-Size="8pt" MaxLength="18" 
                                onkeypress="return numbersonly(event)" HorizontalAlign="Right">
                                <ClientSideEvents TextChanged="function(s, e) {
	                                lblInfo.SetText('');
                                }" />
                            </dx:ASPxTextBox>
                        </td>
                        <td>
                            <dx:ASPxCallback ID="PartNoCallBack" ClientInstanceName="PartNoCallBack" runat="server">
                            <ClientSideEvents EndCallback="EndPartNoCallBack" />
                            </dx:ASPxCallback>
                        </td>
                    </tr>
                </table>
            </td>
        </tr>
    </table>
    
    <div style="height: 1px;"></div>
    
    <table style="width:100%;">
        <tr>
            <td valign="top" align="left">
                <dx:ASPxButton ID="btnSubMenu" runat="server" Text="SUB MENU" Font-Names="Verdana"
                    Width="85px" Font-Size="8pt">
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
            <td valign="top" align="left" style="width: 50px;">
                <dx:ASPxButton ID="btnClear" runat="server" Text="CLEAR" Font-Names="Verdana" Width="90px"
                    AutoPostBack="False" Font-Size="8pt" ClientInstanceName="btnClear">
                </dx:ASPxButton>
            </td>
            <td valign="top" align="left" style="width: 80px;">
                <dx:ASPxButton ID="btnDelete" runat="server" Text="DELETE" Font-Names="Verdana" Width="80px"
                    AutoPostBack="False" Font-Size="8pt" ClientInstanceName="btnDelete">
                    <ClientSideEvents Click="function(s, e) {
                        up_delete();
                        
                        txtPartnoDetail.SetText('');
                        txtPartNo2.SetText('');
                        cboAffiliate2.SetText('');
                        txtAffiliate2.SetText('');                   
                        cboSupplier2.SetText('');
                        txtSupplier2.SetText('');
                        txtQuota.SetText('0');
                        txtLocation.SetText('');
                        cboPacking.SetText('');
                        txtPacking.SetText('');
                        txtMOQ.SetText('0');
                        txtQtyBox.SetText('0');
                        txtBoxPallet.SetText('0');
                        txtNetWeight.SetText('0');
                        txtGrossWeight.SetText('0');
                        txtLength.SetText('0');
                        txtWidth.SetText('0');
                        txtHeight.SetText('0');
                        
                        grid.PerformCallback('load');
                                            
                        txtPartnoDetail.GetInputElement().setAttribute('style', 'background:#FFFFFF;');
                        txtPartnoDetail.GetInputElement().readOnly = false;
                        txtPartnoDetail.SetEnabled(true);

                        cboAffiliate2.GetInputElement().setAttribute('style', 'background:#FFFFFF;');
                        cboAffiliate2.GetInputElement().readOnly = false;
                        cboAffiliate2.SetEnabled(true);
                    }" />
                </dx:ASPxButton>
            </td>
            <td valign="top" align="right" style="width: 50px;">
                <dx:ASPxButton ID="btnSubmit" runat="server" Text="SAVE" Font-Names="Verdana" Width="90px"
                    AutoPostBack="False" Font-Size="8pt" ClientInstanceName="btnSubmit">
                    <ClientSideEvents Click="function(s, e) {
                        validasubmit();
                        up_Insert();

                        txtPartnoDetail.SetText('');
                        txtPartNo2.SetText('');
                        cboAffiliate2.SetText('');
                        txtAffiliate2.SetText('');                   
                        cboSupplier2.SetText('');
                        txtSupplier2.SetText('');
                        txtQuota.SetText('0');
                        txtLocation.SetText('');
                        cboPacking.SetText('');
                        txtPacking.SetText('');
                        txtMOQ.SetText('0');
                        txtQtyBox.SetText('0');
                        txtBoxPallet.SetText('0');
                        txtNetWeight.SetText('0');
                        txtGrossWeight.SetText('0');
                        txtLength.SetText('0');
                        txtWidth.SetText('0');
                        txtHeight.SetText('0');
                                            
                        txtPartnoDetail.GetInputElement().setAttribute('style', 'background:#FFFFFF;');
                        txtPartnoDetail.GetInputElement().readOnly = false;
                        txtPartnoDetail.SetEnabled(true);

                        cboAffiliate2.GetInputElement().setAttribute('style', 'background:#FFFFFF;');
                        cboAffiliate2.GetInputElement().readOnly = false;
                        cboAffiliate2.SetEnabled(true);
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
    <dx:ASPxHiddenField ID="HF" runat="server" ClientInstanceName="HF">
    </dx:ASPxHiddenField>
</asp:Content>
