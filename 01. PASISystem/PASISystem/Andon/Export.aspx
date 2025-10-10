<%@ Page Title="" Language="vb" AutoEventWireup="false" MasterPageFile="~/SiteContent.Master" CodeBehind="Export.aspx.vb" Inherits="PASISystem.Export" %>

<%@ Register Assembly="DevExpress.Web.v14.1, Version=14.1.6.0, Culture=neutral, PublicKeyToken=b88d1754d700e49a"
    Namespace="DevExpress.Web.ASPxGridView.Export" TagPrefix="dx" %>
<%@ Register Assembly="DevExpress.Web.v14.1, Version=14.1.6.0, Culture=neutral, PublicKeyToken=b88d1754d700e49a"
    Namespace="DevExpress.Web.ASPxGridView" TagPrefix="dx" %>
<%@ Register Assembly="DevExpress.Web.v14.1, Version=14.1.6.0, Culture=neutral, PublicKeyToken=b88d1754d700e49a"
    Namespace="DevExpress.Web.ASPxEditors" TagPrefix="dx" %>
<%@ Register Assembly="DevExpress.Web.v14.1, Version=14.1.6.0, Culture=neutral, PublicKeyToken=b88d1754d700e49a"
    Namespace="DevExpress.Web.ASPxGlobalEvents" TagPrefix="dx" %>
<%@ Register Assembly="DevExpress.Web.v14.1, Version=14.1.6.0, Culture=neutral, PublicKeyToken=b88d1754d700e49a"
    Namespace="DevExpress.Web.ASPxPanel" TagPrefix="dx" %>
<%@ Register Assembly="DevExpress.Web.v14.1, Version=14.1.6.0, Culture=neutral, PublicKeyToken=b88d1754d700e49a"
    Namespace="DevExpress.Web.ASPxRoundPanel" TagPrefix="dx" %>
<%@ Register Assembly="DevExpress.Web.v14.1, Version=14.1.6.0, Culture=neutral, PublicKeyToken=b88d1754d700e49a"
    Namespace="DevExpress.Web.ASPxCallback" TagPrefix="dx" %>
<%@ Register Assembly="DevExpress.Web.v14.1, Version=14.1.6.0, Culture=neutral, PublicKeyToken=b88d1754d700e49a"
    Namespace="DevExpress.Web.ASPxEditors" TagPrefix="dx1" %>
<asp:Content ID="Content1" ContentPlaceHolderID="head" runat="server">
    <style type="text/css">
        body {
            font-family: Arial, sans-serif;
        }

        .header {
            display: flex;
            justify-content: space-between;
            align-items: center;
        }

            .header h1 {
                margin-left: 5px;
            }

        .logo {
            width: 100px;
            margin-right: 20px;
        }

        .summary-boxes {
            display: flex;
            gap: 10px;
        }

        .box {
            background-color: #f8f8f8;
            border: 1px solid #ccc;
            border-radius: 8px;
            /*padding: 10px;*/
        }

            .box h3 {
                font-size: 20px;
                margin: 0px;
                padding: 5px;
                font-weight:100;
            }

            .box p {
                font-size: 50px;
                font-weight: bold;
                margin: 0;
                padding: 5px;
            }

        .tables-container {
            display: flex;
            justify-content: space-between;
            margin-top: 15px;
            gap: 20px;
        }

        table {
            border-collapse: collapse;
            width: 100%;
            font-size: 12px;
        }

        .table-section {
            width: 50%;
        }

        .table-title {
            font-family: monospace;
            font-size: 2em;
            font-weight: bold;
            background-color: silver;
            padding: 5px 0px;
        }


        .div-container {
            display: flex;
            justify-content: space-between;
            margin-top: 10px;
            gap: 20px;
        }

        .div-section {
            width: 50%;
        }

        .lblLegend {
            font-size: 50px;
            font-weight: bold;
            margin: 0;
            padding: 5px;
        }
    </style>
    <script language="javascript" type="text/javascript">
        window.onload = function () {
            // Untuk Jam
            updateClock();
            setInterval(updateClock, 1000); // update setiap detik
            function updateClock() {
                var now = new Date();

                var tanggal = now.getDate().toString().padStart(2, '0');
                var bulan = now.toLocaleString('id-ID', { month: 'short' }); // Jul, Agu, dst
                var tahun = now.getFullYear();

                var jam = now.getHours().toString().padStart(2, '0');
                var menit = now.getMinutes().toString().padStart(2, '0');
                var detik = now.getSeconds().toString().padStart(2, '0');

                var hasil = `${tanggal} ${bulan} ${tahun} ${jam}:${menit}:${detik}`;

                lblJam.SetText('Date : ' + hasil);
            }

            // Untuk Legend
            refresh_Legend();
            function refresh_Legend() {
                LegendTick.PerformCallback();
            }
            LegendTick.EndCallback.AddHandler(function () {
                setTimeout(refresh_Legend, 10000); // setiap 10 detik 
            });

            // Untuk Grid Delay Pasi
            refresh_grid_DelayPasi();
            function refresh_grid_DelayPasi() {
                grid_DelayPasi.PerformCallback('load');
            }
            grid_DelayPasi.EndCallback.AddHandler(function () {
                setTimeout(refresh_grid_DelayPasi, 10000); // setiap 10 detik 
            });

            // Untuk Grid Delay Supplier
            refresh_grid_DelaySupplier();
            function refresh_grid_DelaySupplier() {
                grid_DelaySupplier.PerformCallback('load');
            }
            grid_DelaySupplier.EndCallback.AddHandler(function () {
                setTimeout(refresh_grid_DelaySupplier, 10000); // setiap 10 detik 
            });

            // Untuk Grid Delay Receive Pasi
            refresh_grid_ReceivePasi();
            function refresh_grid_ReceivePasi() {
                grid_ReceivePasi.PerformCallback('load');
            }
            grid_ReceivePasi.EndCallback.AddHandler(function () {
                setTimeout(refresh_grid_ReceivePasi, 10000); // setiap 10 detik 
            });

            // Untuk Grid Delay Delivery Pasi
            refresh_grid_DeliveryPasi();
            function refresh_grid_DeliveryPasi() {
                grid_DeliveryPasi.PerformCallback('load');
            }
            grid_DeliveryPasi.EndCallback.AddHandler(function () {
                setTimeout(refresh_grid_DeliveryPasi, 10000); // setiap 10 detik 
            });
        };
    </script>
</asp:Content>
<asp:Content ID="Content2" ContentPlaceHolderID="Content" runat="server">
      <dx:ASPxCallback ID="JamTick" runat="server" ClientInstanceName="JamTick" OnCallback="JamTick_Callback">
          <ClientSideEvents EndCallback="function(s, e) {
            var pMsg = s.cpMessage;

            lblJam.SetText(pMsg);
            delete s.cpMessage;
           }" />       
      </dx:ASPxCallback>

    <dx:ASPxCallback ID="LegendTick" runat="server" ClientInstanceName="LegendTick" OnCallback="LegendTick_Callback">
          <ClientSideEvents EndCallback="function(s, e) {
              var pMsg1 = s.cpMessage_CustOrder;
              var pMsg2 = s.cpMessage_DNPasi;
              var pMsg3 = s.cpMessage_InvoicePasi;
              var pMsg4 = s.cpMessage_SuppOrder;
              var pMsg5 = s.cpMessage_DNSupp;
              var pMsg6 = s.cpMessage_InvSupp;
              
              lbl_CustOrderBulanini.SetText(pMsg1);
              lbl_DNPASIBulan.SetText(pMsg2);
              lbl_InvoicePASIBulan.SetText(pMsg3);
              lbl_SuppOrderBulan.SetText(pMsg4);
              lbl_DNSuppBulan.SetText(pMsg5);
              lbl_InvoiceSuppBulan.SetText(pMsg6);
              
              delete s.cpMessage_CustOrder;
              delete s.cpMessage_DNPasi;
              delete s.cpMessage_InvoicePasi;
              delete s.cpMessage_SuppOrder;
              delete s.cpMessage_DNSupp;
              delete s.cpMessage_InvSupp;
           }" />       
      </dx:ASPxCallback>
        

    <div class="header">
        <div>
            <h1>
                <dx:ASPxLabel ID="lblJam" runat="server" Text="Date : " Font-Names="Tahoma" 
                    ClientInstanceName="lblJam" Font-Bold="True" Font-Italic="True" Font-Size="X-Large">
                </dx:ASPxLabel>
            </h1>
        </div>
        <img src="../Images/Logofix.JPG" alt="YAZAKI" class="logo" />
    </div>

    <%--<div class="summary-boxes">
        <div class="box">
            <h3>Cust Order Bulan ini</h3>
            <p>416</p>
        </div>
        <div class="box">
            <h3>DN PASI Bulan ini</h3>
            <p>670</p>
        </div>
        <div class="box">
            <h3>Invoice PASI Bulan ini</h3>
            <p>200</p>
        </div>
        <div class="box">
            <h3>Supp Order Bulan ini</h3>
            <p>2304</p>
        </div>
        <div class="box">
            <h3>DN Supp Bulan ini</h3>
            <p>2068</p>
        </div>
        <div class="box">
            <h3>Invoice Supp Bulan ini</h3>
            <p>1368</p>
        </div>
    </div>--%>

    <div class="div-container" style="margin:0px;">
        <div class="div-section">
            <div class="box">
                <h3>Cust Order Bulan ini</h3>
                <dx:ASPxLabel ID="lbl_CustOrderBulanini" runat="server" Text="0" Font-Names="Tahoma" 
                    ClientInstanceName="lbl_CustOrderBulanini" Font-Bold="True" Font-Italic="True" CssClass="lblLegend">
                </dx:ASPxLabel>
            </div>
            
        </div>

        <div class="div-section">
            <div class="box">
                <h3>DN PASI Bulan ini</h3>
                <dx:ASPxLabel ID="lbl_DNPASIBulan" runat="server" Text="0" Font-Names="Tahoma" 
                    ClientInstanceName="lbl_DNPASIBulan" Font-Bold="True" Font-Italic="True" CssClass="lblLegend">
                </dx:ASPxLabel>
            </div>
            
        </div>

        <div class="div-section">
            <div class="box">
                <h3>Invoice PASI Bulan ini</h3>
                <dx:ASPxLabel ID="lbl_InvoicePASIBulan" runat="server" Text="0" Font-Names="Tahoma" 
                    ClientInstanceName="lbl_InvoicePASIBulan" Font-Bold="True" Font-Italic="True" CssClass="lblLegend">
                </dx:ASPxLabel>
            </div>
            
        </div>
    </div>

    <div class="div-container">
        <div class="div-section">
            <div class="box">
                <h3>Supp Order Bulan ini</h3>
                <dx:ASPxLabel ID="lbl_SuppOrderBulan" runat="server" Text="0" Font-Names="Tahoma" 
                    ClientInstanceName="lbl_SuppOrderBulan" Font-Bold="True" Font-Italic="True" CssClass="lblLegend">
                </dx:ASPxLabel>
            </div>
            
        </div>

        <div class="div-section">
            <div class="box">
                <h3>DN Supp Bulan ini</h3>
                <dx:ASPxLabel ID="lbl_DNSuppBulan" runat="server" Text="0" Font-Names="Tahoma" 
                    ClientInstanceName="lbl_DNSuppBulan" Font-Bold="True" Font-Italic="True" CssClass="lblLegend">
                </dx:ASPxLabel>
            </div>
            
        </div>

        <div class="div-section">
            <div class="box">
                <h3>Invoice Supp Bulan ini</h3>
                <dx:ASPxLabel ID="lbl_InvoiceSuppBulan" runat="server" Text="0" Font-Names="Tahoma" 
                    ClientInstanceName="lbl_InvoiceSuppBulan" Font-Bold="True" Font-Italic="True" CssClass="lblLegend">
                </dx:ASPxLabel>
            </div>
            
        </div>
    </div>

    <div class="tables-container" style="margin-top:10px;">
        <div class="table-section">
            <div class="table-title">Delay PASI Delivery</div>

            <dx:ASPxGridView ID="grid_DelayPasi" runat="server" Width="100%" Font-Names="Tahoma" KeyFieldName="No" AutoGenerateColumns="False"
                ClientInstanceName="grid_DelayPasi">
                <SettingsLoadingPanel Mode="Disabled" />
                <Columns>
                    <dx:GridViewDataTextColumn VisibleIndex="0" Caption="No" FieldName="No" Width="50px" HeaderStyle-HorizontalAlign="Center">
                        <HeaderStyle HorizontalAlign="Center" Wrap="True"></HeaderStyle>
                        <CellStyle Font-Names="Tahoma" HorizontalAlign="Center"></CellStyle>
                    </dx:GridViewDataTextColumn>

                    <dx:GridViewDataTextColumn VisibleIndex="1" Caption="Kanban Date" FieldName="KanbanDate" HeaderStyle-HorizontalAlign="Center">
                        <HeaderStyle HorizontalAlign="Center" Wrap="True"></HeaderStyle>
                        <CellStyle Font-Names="Tahoma" HorizontalAlign="Center"></CellStyle>
                    </dx:GridViewDataTextColumn>

                    <dx:GridViewDataTextColumn VisibleIndex="2" Caption="Kanban No" FieldName="KanbanNo" HeaderStyle-HorizontalAlign="Center">
                        <HeaderStyle HorizontalAlign="Center" Wrap="True"></HeaderStyle>
                        <CellStyle Font-Names="Tahoma" HorizontalAlign="Center"></CellStyle>
                    </dx:GridViewDataTextColumn>

                    <dx:GridViewDataTextColumn VisibleIndex="3" Caption="Delivery Date" FieldName="DeliveryDate" HeaderStyle-HorizontalAlign="Center">
                        <HeaderStyle HorizontalAlign="Center" Wrap="True"></HeaderStyle>
                        <CellStyle Font-Names="Tahoma" HorizontalAlign="Center"></CellStyle>
                    </dx:GridViewDataTextColumn>

                    <dx:GridViewDataTextColumn VisibleIndex="4" Caption="Affiliate" FieldName="Affiliate" HeaderStyle-HorizontalAlign="Center">
                        <HeaderStyle HorizontalAlign="Center" Wrap="True"></HeaderStyle>
                        <CellStyle Font-Names="Tahoma" HorizontalAlign="Center"></CellStyle>
                    </dx:GridViewDataTextColumn>

                    <dx:GridViewDataTextColumn VisibleIndex="5" Caption="Delay" FieldName="Delay" HeaderStyle-HorizontalAlign="Center">
                        <HeaderStyle HorizontalAlign="Center" Wrap="True"></HeaderStyle>
                        <CellStyle Font-Names="Tahoma" HorizontalAlign="Center"></CellStyle>
                    </dx:GridViewDataTextColumn>
                </Columns>
                <SettingsPager Mode="ShowAllRecords" />
                 <Settings ShowHorizontalScrollBar="True" ShowVerticalScrollBar="True" VerticalScrollableHeight="190" ShowGroupButtons="False" ShowStatusBar="Hidden" />
                <Styles>
                    <SelectedRow ForeColor="Black"></SelectedRow>
                </Styles>
                <StylesEditors ButtonEditCellSpacing="0">
                    <ProgressBar Height="21px"></ProgressBar>
                </StylesEditors>
            </dx:ASPxGridView>
        </div>
        <div class="table-section">
            <div class="table-title">Delay Supplier Delivery</div>

            <dx:ASPxGridView ID="grid_DelaySupplier" runat="server" Width="100%" Font-Names="Tahoma" KeyFieldName="No" AutoGenerateColumns="False"
                ClientInstanceName="grid_DelaySupplier" Font-Size="8pt">
                <SettingsLoadingPanel Mode="Disabled" />
                <Columns>
                    <dx:GridViewDataTextColumn VisibleIndex="0" Caption="No" FieldName="No" Width="50px" HeaderStyle-HorizontalAlign="Center">
                        <HeaderStyle HorizontalAlign="Center" Wrap="True"></HeaderStyle>
                        <CellStyle Font-Names="Tahoma" HorizontalAlign="Center"></CellStyle>
                    </dx:GridViewDataTextColumn>

                    <dx:GridViewDataTextColumn VisibleIndex="1" Caption="Kanban Date" FieldName="KanbanDate" HeaderStyle-HorizontalAlign="Center">
                        <HeaderStyle HorizontalAlign="Center" Wrap="True"></HeaderStyle>
                        <CellStyle Font-Names="Tahoma" HorizontalAlign="Center"></CellStyle>
                    </dx:GridViewDataTextColumn>

                    <dx:GridViewDataTextColumn VisibleIndex="2" Caption="Kanban No" FieldName="KanbanNo" HeaderStyle-HorizontalAlign="Center">
                        <HeaderStyle HorizontalAlign="Center" Wrap="True"></HeaderStyle>
                        <CellStyle Font-Names="Tahoma" HorizontalAlign="Center"></CellStyle>
                    </dx:GridViewDataTextColumn>

                    <dx:GridViewDataTextColumn VisibleIndex="3" Caption="Delivery Date" FieldName="DeliveryDate" HeaderStyle-HorizontalAlign="Center">
                        <HeaderStyle HorizontalAlign="Center" Wrap="True"></HeaderStyle>
                        <CellStyle Font-Names="Tahoma" HorizontalAlign="Center"></CellStyle>
                    </dx:GridViewDataTextColumn>

                    <dx:GridViewDataTextColumn VisibleIndex="4" Caption="Affiliate" FieldName="Affiliate" HeaderStyle-HorizontalAlign="Center">
                        <HeaderStyle HorizontalAlign="Center" Wrap="True"></HeaderStyle>
                        <CellStyle Font-Names="Tahoma" HorizontalAlign="Center"></CellStyle>
                    </dx:GridViewDataTextColumn>

                    <dx:GridViewDataTextColumn VisibleIndex="5" Caption="Delay" FieldName="Delay" HeaderStyle-HorizontalAlign="Center">
                        <HeaderStyle HorizontalAlign="Center" Wrap="True"></HeaderStyle>
                        <CellStyle Font-Names="Tahoma" HorizontalAlign="Center"></CellStyle>
                    </dx:GridViewDataTextColumn>
                </Columns>
                <SettingsPager PageSize="100" Position="Top">
                    <Summary Text="Page {0} of {1} [{2} record(s)]" AllPagesText="Page {0} of {1} " />
                    <Summary AllPagesText="Page {0} of {1} " Text="Page {0} of {1} [{2} record(s)]"></Summary>
                </SettingsPager>
                <Settings ShowHorizontalScrollBar="True" ShowVerticalScrollBar="True" ShowGroupButtons="False"
                    ShowStatusBar="Hidden" VerticalScrollableHeight="190" />
                <Styles>
                    <SelectedRow ForeColor="Black"></SelectedRow>
                </Styles>
                <StylesEditors ButtonEditCellSpacing="0">
                    <ProgressBar Height="21px"></ProgressBar>
                </StylesEditors>
            </dx:ASPxGridView>
        </div>
    </div>

    <div class="tables-container">
        <div class="table-section">
            <div class="table-title">Rencana Penerimaan PASI</div>
            
            <dx:ASPxGridView ID="grid_ReceivePasi" runat="server" Width="100%" Font-Names="Tahoma" KeyFieldName="No" AutoGenerateColumns="False"
                ClientInstanceName="grid_ReceivePasi" Font-Size="8pt">
                <SettingsLoadingPanel Mode="Disabled" />
                <Columns>
                    <dx:GridViewDataTextColumn VisibleIndex="0" Caption="No" FieldName="No" Width="50px" HeaderStyle-HorizontalAlign="Center">
                        <HeaderStyle HorizontalAlign="Center" Wrap="True"></HeaderStyle>
                        <CellStyle Font-Names="Tahoma" HorizontalAlign="Center"></CellStyle>
                    </dx:GridViewDataTextColumn>

                    <dx:GridViewDataTextColumn VisibleIndex="1" Caption="Delivery Date" FieldName="DeliveryDate" HeaderStyle-HorizontalAlign="Center">
                        <HeaderStyle HorizontalAlign="Center" Wrap="True"></HeaderStyle>
                        <CellStyle Font-Names="Tahoma" HorizontalAlign="Center"></CellStyle>
                    </dx:GridViewDataTextColumn>

                    <dx:GridViewDataTextColumn VisibleIndex="2" Caption="Supplier" FieldName="Supplier" HeaderStyle-HorizontalAlign="Center">
                        <HeaderStyle HorizontalAlign="Center" Wrap="True"></HeaderStyle>
                        <CellStyle Font-Names="Tahoma" HorizontalAlign="Center"></CellStyle>
                    </dx:GridViewDataTextColumn>

                    <dx:GridViewDataTextColumn VisibleIndex="3" Caption="Surat Jalan" FieldName="SuratJalanNo" HeaderStyle-HorizontalAlign="Center">
                        <HeaderStyle HorizontalAlign="Center" Wrap="True"></HeaderStyle>
                        <CellStyle Font-Names="Tahoma" HorizontalAlign="Center"></CellStyle>
                    </dx:GridViewDataTextColumn>

                    <dx:GridViewDataTextColumn VisibleIndex="4" Caption="Kanban No" FieldName="KanbanNo" HeaderStyle-HorizontalAlign="Center">
                        <HeaderStyle HorizontalAlign="Center" Wrap="True"></HeaderStyle>
                        <CellStyle Font-Names="Tahoma" HorizontalAlign="Center"></CellStyle>
                    </dx:GridViewDataTextColumn>

                    <dx:GridViewDataTextColumn VisibleIndex="5" Caption="Driver Name" FieldName="DriverName" HeaderStyle-HorizontalAlign="Center">
                        <HeaderStyle HorizontalAlign="Center" Wrap="True"></HeaderStyle>
                        <CellStyle Font-Names="Tahoma" HorizontalAlign="Center"></CellStyle>
                    </dx:GridViewDataTextColumn>

                    <dx:GridViewDataTextColumn VisibleIndex="6" Caption="Driver Contact" FieldName="DriverContact" HeaderStyle-HorizontalAlign="Center">
                        <HeaderStyle HorizontalAlign="Center" Wrap="True"></HeaderStyle>
                        <CellStyle Font-Names="Tahoma" HorizontalAlign="Center"></CellStyle>
                    </dx:GridViewDataTextColumn>
                </Columns>
                <SettingsPager PageSize="100" Position="Top">
                    <Summary Text="Page {0} of {1} [{2} record(s)]" AllPagesText="Page {0} of {1} " />
                    <Summary AllPagesText="Page {0} of {1} " Text="Page {0} of {1} [{2} record(s)]"></Summary>
                </SettingsPager>
                <Settings ShowHorizontalScrollBar="True" ShowVerticalScrollBar="True" ShowGroupButtons="False"
                    ShowStatusBar="Hidden" VerticalScrollableHeight="190" />
                <Styles>
                    <SelectedRow ForeColor="Black"></SelectedRow>
                </Styles>
                <StylesEditors ButtonEditCellSpacing="0">
                    <ProgressBar Height="21px"></ProgressBar>
                </StylesEditors>
            </dx:ASPxGridView>
        </div>
        <div class="table-section">
            <div class="table-title">Rencana Pengiriman PASI</div>
            
            <dx:ASPxGridView ID="grid_DeliveryPasi" runat="server" Width="100%" Font-Names="Tahoma" KeyFieldName="No" AutoGenerateColumns="False"
                ClientInstanceName="grid_DeliveryPasi" Font-Size="8pt">
                <SettingsLoadingPanel Mode="Disabled" />
                <Columns>
                    <dx:GridViewDataTextColumn VisibleIndex="0" Caption="No" FieldName="No" Width="50px" HeaderStyle-HorizontalAlign="Center">
                        <HeaderStyle HorizontalAlign="Center" Wrap="True"></HeaderStyle>
                        <CellStyle Font-Names="Tahoma" HorizontalAlign="Center"></CellStyle>
                    </dx:GridViewDataTextColumn>

                    <dx:GridViewDataTextColumn VisibleIndex="1" Caption="Delivery Date" FieldName="DeliveryDate" HeaderStyle-HorizontalAlign="Center">
                        <HeaderStyle HorizontalAlign="Center" Wrap="True"></HeaderStyle>
                        <CellStyle Font-Names="Tahoma" HorizontalAlign="Center"></CellStyle>
                    </dx:GridViewDataTextColumn>

                    <dx:GridViewDataTextColumn VisibleIndex="2" Caption="Supplier" FieldName="Supplier" HeaderStyle-HorizontalAlign="Center">
                        <HeaderStyle HorizontalAlign="Center" Wrap="True"></HeaderStyle>
                        <CellStyle Font-Names="Tahoma" HorizontalAlign="Center"></CellStyle>
                    </dx:GridViewDataTextColumn>

                    <dx:GridViewDataTextColumn VisibleIndex="3" Caption="Surat Jalan" FieldName="SuratJalanNo" HeaderStyle-HorizontalAlign="Center">
                        <HeaderStyle HorizontalAlign="Center" Wrap="True"></HeaderStyle>
                        <CellStyle Font-Names="Tahoma" HorizontalAlign="Center"></CellStyle>
                    </dx:GridViewDataTextColumn>

                    <dx:GridViewDataTextColumn VisibleIndex="4" Caption="Kanban No" FieldName="KanbanNo" HeaderStyle-HorizontalAlign="Center">
                        <HeaderStyle HorizontalAlign="Center" Wrap="True"></HeaderStyle>
                        <CellStyle Font-Names="Tahoma" HorizontalAlign="Center"></CellStyle>
                    </dx:GridViewDataTextColumn>

                    <dx:GridViewDataTextColumn VisibleIndex="5" Caption="Driver Name" FieldName="DriverName" HeaderStyle-HorizontalAlign="Center">
                        <HeaderStyle HorizontalAlign="Center" Wrap="True"></HeaderStyle>
                        <CellStyle Font-Names="Tahoma" HorizontalAlign="Center"></CellStyle>
                    </dx:GridViewDataTextColumn>

                    <dx:GridViewDataTextColumn VisibleIndex="6" Caption="Driver Contact" FieldName="DriverContact" HeaderStyle-HorizontalAlign="Center">
                        <HeaderStyle HorizontalAlign="Center" Wrap="True"></HeaderStyle>
                        <CellStyle Font-Names="Tahoma" HorizontalAlign="Center"></CellStyle>
                    </dx:GridViewDataTextColumn>
                </Columns>
                <SettingsPager PageSize="100" Position="Top">
                    <Summary Text="Page {0} of {1} [{2} record(s)]" AllPagesText="Page {0} of {1} " />
                    <Summary AllPagesText="Page {0} of {1} " Text="Page {0} of {1} [{2} record(s)]"></Summary>
                </SettingsPager>
                <Settings ShowHorizontalScrollBar="True" ShowVerticalScrollBar="True" ShowGroupButtons="False"
                    ShowStatusBar="Hidden" VerticalScrollableHeight="190" />
                <Styles>
                    <SelectedRow ForeColor="Black"></SelectedRow>
                </Styles>
                <StylesEditors ButtonEditCellSpacing="0">
                    <ProgressBar Height="21px"></ProgressBar>
                </StylesEditors>
            </dx:ASPxGridView>
        </div>
    </div>
</asp:Content>
