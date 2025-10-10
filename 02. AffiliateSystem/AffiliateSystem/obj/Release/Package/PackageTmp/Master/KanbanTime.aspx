<%@ Page Title="" Language="vb" AutoEventWireup="false" MasterPageFile="~/SiteContent.Master" CodeBehind="KanbanTime.aspx.vb" Inherits="AffiliateSystem.KanbanTime" %>
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
        .style3
        {
            height: 25px;
            width: 70px;
        }
        .style7
        {
            height: 25px;
            width: 70px;
        }
        .style18
        {
            width: 112px;
            height: 25px;
        }
        .style20
        {
            width: 70px;
            height: 25px;
        }
        .style21
        {
            height: 25px;
            width: 700px;
        }
        #Table1
        {
            width: 100%;
            margin-left: 0px;
        }
        .style24
        {
            width: 100%;
        }
        .style25
        {
            width: 100%;
            height: 475px;
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
    
    function up_delete() {

        if (Cycle1.GetValue() == "") {
            lblInfo.GetMainElement().style.color = 'Red';
            lblInfo.SetText("[6011] Please input the time 1 first!");
            e.ProcessOnServer = false;
            return false;
        }

        if (Cycle2.GetValue() == "") {
            lblInfo.GetMainElement().style.color = 'Red';
            lblInfo.SetText("[6011] Please input the time 2 first!");
            e.ProcessOnServer = false;
            return false;
        }

        if (Cycle3.GetValue() == "") {
            lblInfo.GetMainElement().style.color = 'Red';
            lblInfo.SetText("[6011] Please input the time 3 first!");
            e.ProcessOnServer = false;
            return false;
        }

        if (Cycle4.GetValue() == "") {
            lblInfo.GetMainElement().style.color = 'Red';
            lblInfo.SetText("[6011] Please input the time 4 first!");
            e.ProcessOnServer = false;
            return false;
        }


        var msg = confirm('Are you sure want to delete this data ?');
        if (msg == false) {
            e.processOnServer = false;
            return;
        }

        AffiliateSubmit.PerformCallback('delete|' );
        
    }

    function readonly() {
        txtPartID.GetInputElement().setAttribute('style', 'background:#CCCCCC;foreground:#CCCCCC;');
        txtPartID.GetInputElement().readOnly = true;
        lblInfo.SetText('');
    }


    function validasubmit() {

        lblInfo.GetMainElement().style.color = 'Red';
        if (Cycle1.GetValue() == "") {
            lblInfo.GetMainElement().style.color = 'Red';
            lblInfo.SetText("[6011] Please input the time 1 first!");
            e.ProcessOnServer = false;
            return false;
        }

        if (Cycle2.GetValue() == "") {
            lblInfo.GetMainElement().style.color = 'Red';
            lblInfo.SetText("[6011] Please input the time 2 first!");
            e.ProcessOnServer = false;
            return false;
        }

        if (Cycle3.GetValue() == "") {
            lblInfo.GetMainElement().style.color = 'Red';
            lblInfo.SetText("[6011] Please input the time 3 first!");
            e.ProcessOnServer = false;
            return false;
        }

        if (Cycle4.GetValue() == "") {
            lblInfo.GetMainElement().style.color = 'Red';
            lblInfo.SetText("[6011] Please input the time 4 first!");
            e.ProcessOnServer = false;
            return false;
        }

    }

    function up_Insert() {
   
        var pIsUpdate = '';

        var pTime1 = Cycle1.GetValue();
        var pTime2 = Cycle2.GetValue();
        var pTime3 = Cycle3.GetValue();
        var pTime4 = Cycle4.GetValue();

        var pTime5 = Cycle5.GetValue();
        var pTime6 = Cycle6.GetValue();
        var pTime7 = Cycle7.GetValue();
        var pTime8 = Cycle8.GetValue();

        var pTime9 = Cycle9.GetValue();
        var pTime10 = Cycle10.GetValue();
        var pTime11 = Cycle11.GetValue();
        var pTime12 = Cycle12.GetValue();

        var pTime13 = Cycle13.GetValue();
        var pTime14 = Cycle14.GetValue();
        var pTime15 = Cycle15.GetValue();
        var pTime16 = Cycle16.GetValue();

        var pTime17 = Cycle17.GetValue();
        var pTime18 = Cycle18.GetValue();
        var pTime19 = Cycle19.GetValue();
        var pTime20 = Cycle20.GetValue();

        AffiliateSubmit.PerformCallback('save|' + pIsUpdate + '|' + pTime1 + '|' + pTime2 + '|' + pTime3 + '|' + pTime4 + '|' + pTime5 + '|' + pTime6 + '|' + pTime7 + '|' + pTime8 + '|' + pTime9 + '|' + pTime10 + '|' + pTime11 + '|' + pTime12 + '|' + pTime13 + '|' + pTime14 + '|' + pTime15 + '|' + pTime16 + '|' + pTime17 + '|' + pTime18 + '|' + pTime19 + '|' + pTime20);
       
    }
</script>
</asp:Content>

<asp:Content ID="Content2" ContentPlaceHolderID="Content" runat="server">
    <table style="width:100%;">
        <tr>
            <td >
                <%--error message--%>
                <table id="tblMsg" style="border-width: thin; border-style: inset hidden ridge hidden; border-color:#9598A1; width:100%; height: 25px;">
                    <tr>
                        <td align="left" valign="top" >
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
    
    <table style="width:100%; height: 80%;">
        <tr>
            <td width="100px">
                &nbsp;</td>
            <td>
                <table style="width:100%;">
                    <tr>
                        <td align="left" width="150px">
                            &nbsp;</td>
                        <td align="left" width="70px">       
                            <dx:ASPxLabel ID="ASPxLabel69" runat="server" Text="CYCLE 1"
                                Font-Names="Verdana" Font-Size="8pt">
                            </dx:ASPxLabel>
                        </td>
                        <td align="left" width="150px">
                            &nbsp;</td>
                        <td align="left" width="70px">
                            <dx:ASPxLabel ID="ASPxLabel73" runat="server" Text="CYCLE 2"
                                Font-Names="Verdana" Font-Size="8pt">
                            </dx:ASPxLabel>
                        </td>
                        <td align="left" width="150px">
                            &nbsp;</td>
                        <td align="left" width="70px">
                            <dx:ASPxLabel ID="ASPxLabel71" runat="server" Text="CYCLE 3"
                                Font-Names="Verdana" Font-Size="8pt">
                            </dx:ASPxLabel>
                        </td>
                        <td align="left" width="150px">
                            &nbsp;</td>
                        <td align="left" width="70px">
                            <dx:ASPxLabel ID="ASPxLabel72" runat="server" Text="CYCLE 4"
                                Font-Names="Verdana" Font-Size="8pt">
                            </dx:ASPxLabel>
                        </td>
                        <td align="left" width="150px">
                            &nbsp;</td>
                    </tr>
                    <tr>
                        <td align="left" width="150px">
                            &nbsp;</td>
                        <td align="left" width="70px">       
                            <dx:ASPxTimeEdit ID="Cycle1" runat="server" ClientInstanceName="Cycle1" 
                                Width="70px" EditFormat="Custom" EditFormatString="HH:mm">
                            </dx:ASPxTimeEdit>
                        </td>
                        <td align="left" width="150px">
                            &nbsp;</td>
                        <td align="left" width="70px">
                            <dx:ASPxTimeEdit ID="Cycle2" runat="server" ClientInstanceName="Cycle2" 
                                Width="70px" EditFormat="Custom" EditFormatString="HH:mm">
                            </dx:ASPxTimeEdit>
                        </td>
                        <td align="left" width="150px">
                            &nbsp;</td>
                        <td align="left" width="70px">
                            <dx:ASPxTimeEdit ID="Cycle3" runat="server" ClientInstanceName="Cycle3" 
                                Width="70px" EditFormat="Custom" EditFormatString="HH:mm">
                            </dx:ASPxTimeEdit>
                        </td>
                        <td align="left" width="150px">
                            &nbsp;</td>
                        <td align="left" width="70px">
                            <dx:ASPxTimeEdit ID="Cycle4" runat="server" ClientInstanceName="Cycle4" 
                                Width="70px" EditFormat="Custom" EditFormatString="HH:mm">
                            </dx:ASPxTimeEdit>
                        </td> 
                        <td align="left" width="150px">
                            &nbsp;</td>
                    </tr>

                    <tr>
                        <td align="left" width="150px">
                            &nbsp;</td>
                        <td align="left" width="70px">       
                            <dx:ASPxLabel ID="ASPxLabel2" runat="server" Text="CYCLE 5"
                                Font-Names="Verdana" Font-Size="8pt">
                            </dx:ASPxLabel>
                        </td>
                        <td align="left" width="150px">
                            &nbsp;</td>
                        <td align="left" width="70px">
                            <dx:ASPxLabel ID="ASPxLabel3" runat="server" Text="CYCLE 6"
                                Font-Names="Verdana" Font-Size="8pt">
                            </dx:ASPxLabel>
                        </td>
                        <td align="left" width="150px">
                            &nbsp;</td>
                        <td align="left" width="70px">
                            <dx:ASPxLabel ID="ASPxLabel4" runat="server" Text="CYCLE 7"
                                Font-Names="Verdana" Font-Size="8pt">
                            </dx:ASPxLabel>
                        </td>
                        <td align="left" width="150px">
                            &nbsp;</td>
                        <td align="left" width="70px">
                            <dx:ASPxLabel ID="ASPxLabel5" runat="server" Text="CYCLE 8"
                                Font-Names="Verdana" Font-Size="8pt">
                            </dx:ASPxLabel>
                        </td>
                        <td align="left" width="150px">
                            &nbsp;</td>
                    </tr>
                    <tr>
                        <td align="left" width="150px">
                            &nbsp;</td>
                        <td align="left" width="70px">       
                            <dx:ASPxTimeEdit ID="Cycle5" runat="server" ClientInstanceName="Cycle5" 
                                Width="70px" EditFormat="Custom" EditFormatString="HH:mm">
                            </dx:ASPxTimeEdit>
                        </td>
                        <td align="left" width="150px">
                            &nbsp;</td>
                        <td align="left" width="70px">
                            <dx:ASPxTimeEdit ID="Cycle6" runat="server" ClientInstanceName="Cycle6" 
                                Width="70px" EditFormat="Custom" EditFormatString="HH:mm">
                            </dx:ASPxTimeEdit>
                        </td>
                        <td align="left" width="150px">
                            &nbsp;</td>
                        <td align="left" width="70px">
                            <dx:ASPxTimeEdit ID="Cycle7" runat="server" ClientInstanceName="Cycle7" 
                                Width="70px" EditFormat="Custom" EditFormatString="HH:mm">
                            </dx:ASPxTimeEdit>
                        </td>
                        <td align="left" width="150px">
                            &nbsp;</td>
                        <td align="left" width="70px">
                            <dx:ASPxTimeEdit ID="Cycle8" runat="server" ClientInstanceName="Cycle8" 
                                Width="70px" EditFormat="Custom" EditFormatString="HH:mm">
                            </dx:ASPxTimeEdit>
                        </td> 
                        <td align="left" width="150px">
                            &nbsp;</td>
                    </tr>

                    <tr>
                        <td align="left" width="150px">
                            &nbsp;</td>
                        <td align="left" width="70px">       
                            <dx:ASPxLabel ID="ASPxLabel6" runat="server" Text="CYCLE 9"
                                Font-Names="Verdana" Font-Size="8pt">
                            </dx:ASPxLabel>
                        </td>
                        <td align="left" width="150px">
                            &nbsp;</td>
                        <td align="left" width="70px">
                            <dx:ASPxLabel ID="ASPxLabel7" runat="server" Text="CYCLE 10"
                                Font-Names="Verdana" Font-Size="8pt">
                            </dx:ASPxLabel>
                        </td>
                        <td align="left" width="150px">
                            &nbsp;</td>
                        <td align="left" width="70px">
                            <dx:ASPxLabel ID="ASPxLabel8" runat="server" Text="CYCLE 11"
                                Font-Names="Verdana" Font-Size="8pt">
                            </dx:ASPxLabel>
                        </td>
                        <td align="left" width="150px">
                            &nbsp;</td>
                        <td align="left" width="70px">
                            <dx:ASPxLabel ID="ASPxLabel9" runat="server" Text="CYCLE 12"
                                Font-Names="Verdana" Font-Size="8pt">
                            </dx:ASPxLabel>
                        </td>
                        <td align="left" width="150px">
                            &nbsp;</td>
                    </tr>
                    <tr>
                        <td align="left" width="150px">
                            &nbsp;</td>
                        <td align="left" width="70px">       
                            <dx:ASPxTimeEdit ID="Cycle9" runat="server" ClientInstanceName="Cycle9" 
                                Width="70px" EditFormat="Custom" EditFormatString="HH:mm">
                            </dx:ASPxTimeEdit>
                        </td>
                        <td align="left" width="150px">
                            &nbsp;</td>
                        <td align="left" width="70px">
                            <dx:ASPxTimeEdit ID="Cycle10" runat="server" ClientInstanceName="Cycle10" 
                                Width="70px" EditFormat="Custom" EditFormatString="HH:mm">
                            </dx:ASPxTimeEdit>
                        </td>
                        <td align="left" width="150px">
                            &nbsp;</td>
                        <td align="left" width="70px">
                            <dx:ASPxTimeEdit ID="Cycle11" runat="server" ClientInstanceName="Cycle11" 
                                Width="70px" EditFormat="Custom" EditFormatString="HH:mm">
                            </dx:ASPxTimeEdit>
                        </td>
                        <td align="left" width="150px">
                            &nbsp;</td>
                        <td align="left" width="70px">
                            <dx:ASPxTimeEdit ID="Cycle12" runat="server" ClientInstanceName="Cycle12" 
                                Width="70px" EditFormat="Custom" EditFormatString="HH:mm">
                            </dx:ASPxTimeEdit>
                        </td> 
                        <td align="left" width="150px">
                            &nbsp;</td>
                    </tr>

                    <tr>
                        <td align="left" width="150px">
                            &nbsp;</td>
                        <td align="left" width="70px">       
                            <dx:ASPxLabel ID="ASPxLabel10" runat="server" Text="CYCLE 13"
                                Font-Names="Verdana" Font-Size="8pt">
                            </dx:ASPxLabel>
                        </td>
                        <td align="left" width="150px">
                            &nbsp;</td>
                        <td align="left" width="70px">
                            <dx:ASPxLabel ID="ASPxLabel11" runat="server" Text="CYCLE 14"
                                Font-Names="Verdana" Font-Size="8pt">
                            </dx:ASPxLabel>
                        </td>
                        <td align="left" width="150px">
                            &nbsp;</td>
                        <td align="left" width="70px">
                            <dx:ASPxLabel ID="ASPxLabel12" runat="server" Text="CYCLE 15"
                                Font-Names="Verdana" Font-Size="8pt">
                            </dx:ASPxLabel>
                        </td>
                        <td align="left" width="150px">
                            &nbsp;</td>
                        <td align="left" width="70px">
                            <dx:ASPxLabel ID="ASPxLabel13" runat="server" Text="CYCLE 16"
                                Font-Names="Verdana" Font-Size="8pt">
                            </dx:ASPxLabel>
                        </td>
                        <td align="left" width="150px">
                            &nbsp;</td>
                    </tr>
                    <tr>
                        <td align="left" width="150px">
                            &nbsp;</td>
                        <td align="left" width="70px">       
                            <dx:ASPxTimeEdit ID="Cycle13" runat="server" ClientInstanceName="Cycle13" 
                                Width="70px" EditFormat="Custom" EditFormatString="HH:mm">
                            </dx:ASPxTimeEdit>
                        </td>
                        <td align="left" width="150px">
                            &nbsp;</td>
                        <td align="left" width="70px">
                            <dx:ASPxTimeEdit ID="Cycle14" runat="server" ClientInstanceName="Cycle14" 
                                Width="70px" EditFormat="Custom" EditFormatString="HH:mm">
                            </dx:ASPxTimeEdit>
                        </td>
                        <td align="left" width="150px">
                            &nbsp;</td>
                        <td align="left" width="70px">
                            <dx:ASPxTimeEdit ID="Cycle15" runat="server" ClientInstanceName="Cycle15" 
                                Width="70px" EditFormat="Custom" EditFormatString="HH:mm">
                            </dx:ASPxTimeEdit>
                        </td>
                        <td align="left" width="150px">
                            &nbsp;</td>
                        <td align="left" width="70px">
                            <dx:ASPxTimeEdit ID="Cycle16" runat="server" ClientInstanceName="Cycle16" 
                                Width="70px" EditFormat="Custom" EditFormatString="HH:mm">
                            </dx:ASPxTimeEdit>
                        </td> 
                        <td align="left" width="150px">
                            &nbsp;</td>
                    </tr>

                    <tr>
                        <td align="left" width="150px">
                            &nbsp;</td>
                        <td align="left" width="70px">       
                            <dx:ASPxLabel ID="ASPxLabel14" runat="server" Text="CYCLE 17"
                                Font-Names="Verdana" Font-Size="8pt">
                            </dx:ASPxLabel>
                        </td>
                        <td align="left" width="150px">
                            &nbsp;</td>
                        <td align="left" width="70px">
                            <dx:ASPxLabel ID="ASPxLabel15" runat="server" Text="CYCLE 18"
                                Font-Names="Verdana" Font-Size="8pt">
                            </dx:ASPxLabel>
                        </td>
                        <td align="left" width="150px">
                            &nbsp;</td>
                        <td align="left" width="70px">
                            <dx:ASPxLabel ID="ASPxLabel16" runat="server" Text="CYCLE 19"
                                Font-Names="Verdana" Font-Size="8pt">
                            </dx:ASPxLabel>
                        </td>
                        <td align="left" width="150px">
                            &nbsp;</td>
                        <td align="left" width="70px">
                            <dx:ASPxLabel ID="ASPxLabel17" runat="server" Text="CYCLE 20"
                                Font-Names="Verdana" Font-Size="8pt">
                            </dx:ASPxLabel>
                        </td>
                        <td align="left" width="150px">
                            &nbsp;</td>
                    </tr>
                    <tr>
                        <td align="left" width="150px">
                            &nbsp;</td>
                        <td align="left" width="70px">       
                            <dx:ASPxTimeEdit ID="Cycle17" runat="server" ClientInstanceName="Cycle17" 
                                Width="70px" EditFormat="Custom" EditFormatString="HH:mm">
                            </dx:ASPxTimeEdit>
                        </td>
                        <td align="left" width="150px">
                            &nbsp;</td>
                        <td align="left" width="70px">
                            <dx:ASPxTimeEdit ID="Cycle18" runat="server" ClientInstanceName="Cycle18" 
                                Width="70px" EditFormat="Custom" EditFormatString="HH:mm">
                            </dx:ASPxTimeEdit>
                        </td>
                        <td align="left" width="150px">
                            &nbsp;</td>
                        <td align="left" width="70px">
                            <dx:ASPxTimeEdit ID="Cycle19" runat="server" ClientInstanceName="Cycle19" 
                                Width="70px" EditFormat="Custom" EditFormatString="HH:mm">
                            </dx:ASPxTimeEdit>
                        </td>
                        <td align="left" width="150px">
                            &nbsp;</td>
                        <td align="left" width="70px">
                            <dx:ASPxTimeEdit ID="Cycle20" runat="server" ClientInstanceName="Cycle20" 
                                Width="70px" EditFormat="Custom" EditFormatString="HH:mm">
                            </dx:ASPxTimeEdit>
                        </td> 
                        <td align="left" width="150px">
                            &nbsp;</td>
                    </tr>
                </table>
            </td>
            <td width="100px">
                &nbsp;</td>
        </tr>
    </table>

    <div style="height:8px;"></div>

    <table id="button" style=" width:100%;">
        <tr>
            <td valign="top" align="left">
                <dx:ASPxButton ID="btnSubMenu" runat="server" Text="SUB MENU"
                    Font-Names="Verdana" Width="85px" Font-Size="8pt" TabIndex="20">
                </dx:ASPxButton>
            </td>
            <td valign="top" align="right" style="width: 50px;">      
                &nbsp;</td>
            <td valign="top" align="right" style="width: 50px;">                                   
                &nbsp;</td>
            <td align="right" style="width:80px;">                                   
                &nbsp;</td>
            <td valign="top" align="right" style="width: 50px;">                                   
                <dx:ASPxButton ID="btnSubmit" runat="server" Text="SAVE"                    
                    Font-Names="Verdana"
                    Width="90px" AutoPostBack="False" Font-Size="8pt" TabIndex="17">
                    <ClientSideEvents Click="function(s, e) {
                         validasubmit();
                        up_Insert();
                    
                    
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
        }" />
    </dx:ASPxCallback>

   
    <dx:ASPxCallback ID="cbSetData" runat="server" ClientInstanceName="cbSetData">
        <ClientSideEvents CallbackComplete="function(s, e) {
	 			if (s.cpTime1) {
				Cycle1.SetText(s.cpTime1);
				}
                
	            if (s.cpTime2) {
				Cycle2.SetText(s.cpTime2);
				}

                if (s.cpTime3) {
				Cycle3.SetText(s.cpTime3);
				}

                if (s.cpTime4) {
				Cycle4.SetText(s.cpTime4);
				}
}" EndCallback="function(s, e) {
	if (s.cpTime1) {
				Cycle1.SetText(s.cpTime1);
				}
                
	            if (s.cpTime2) {
				Cycle2.SetText(s.cpTime2);
				}

                if (s.cpTime3) {
				Cycle3.SetText(s.cpTime3);
				}

                if (s.cpTime4) {
				Cycle4.SetText(s.cpTime4);
				}
}" />
    </dx:ASPxCallback>
</asp:Content>

