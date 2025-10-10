Imports System.Drawing
Imports DevExpress
Imports DevExpress.Web


Public Class clsAppearance
    Inherits System.Web.UI.Page

#Region "DECLARE"
    Public Enum ShowHorizontalScrollMode
        Auto = 2
        Hidden = 0
        Visible = 1
    End Enum

    Public Enum PagerMode
        ShowAllRecord = 0
        ShowPager = 1
        EndlessPaging = 2
    End Enum
#End Region
    
#Region "PROCEDURE"
    Public Sub setASPxGridView(ByRef pGridControl As ASPxGridView.ASPxGridView, ByVal pShowHorizontalScroll As ShowHorizontalScrollMode, ByVal pAllowFocusedRow As Boolean, ByVal pAllowSort As Boolean)
        With pGridControl
            'Accessibility
            .KeyboardSupport = True

            'Settings
            '-Settings
            .Settings.ShowGroupButtons = False
            .Settings.HorizontalScrollBarMode = ASPxClasses.ScrollBarMode.Visible
            .Settings.VerticalScrollBarMode = ASPxClasses.ScrollBarMode.Visible
            .Settings.VerticalScrollableHeight = 240
            '-SettingsBehaviour
            .SettingsBehavior.AllowFocusedRow = pAllowFocusedRow
            .SettingsBehavior.AllowSort = pAllowSort
            '.SettingsBehavior.ColumnResizeMode = ASPxGridView.ColumnResizeMode.Control
            .SettingsBehavior.EnableRowHotTrack = True
            '-SettingsPager
            .SettingsPager.Mode = ASPxGridView.GridViewPagerMode.ShowAllRecords
            .SettingsPager.PageSize = 1000
            .SettingsPager.Visible = False

            'Styles
            '-Header
            .Styles.Header.BackColor = Color.FromName("#FFD2A6")
            .Styles.Header.Font.Name = "Verdana"
            .Styles.Header.Font.Size = "8"
            .Styles.Header.ForeColor = Color.Black
            .Styles.Header.Wrap = Utils.DefaultBoolean.True
            '-Row
            .Styles.Row.BackColor = Color.FromName("#FFFFE1")
            .Styles.Row.Font.Name = "Verdana"
            .Styles.Row.Font.Size = "8"
            .Styles.Row.Wrap = Utils.DefaultBoolean.False
            '-RowHotTrack
            .Styles.RowHotTrack.BackColor = Color.FromName("#E8EFFD")
            .Styles.RowHotTrack.Font.Name = "Verdana"
            .Styles.RowHotTrack.Font.Size = "8"
            .Styles.RowHotTrack.ForeColor = Color.Black
            .Styles.RowHotTrack.Wrap = Utils.DefaultBoolean.False
            '-FocusedRow
            .Styles.FocusedRow.BackColor = Color.FromName("#DCE7FC")
            .Styles.FocusedRow.Font.Name = "Verdana"
            .Styles.FocusedRow.Font.Size = "8"
            .Styles.FocusedRow.ForeColor = Color.Black
            .Styles.FocusedRow.Wrap = Utils.DefaultBoolean.False
            '-SelectedRow
            .Styles.SelectedRow.Wrap = Utils.DefaultBoolean.False
        End With
    End Sub

    Public Sub setAppearanceControlsDevEx11(ByRef pRefPage As Page, _
                                            Optional ByVal pShowHorizontalScroll As Boolean = False, _
                                            Optional ByVal pAllowFocusedRow As Boolean = False, _
                                            Optional ByVal pAllowSort As Boolean = False)
        For Each ctlMaster As Control In pRefPage.Controls
            If TypeOf ctlMaster Is MasterPage Then

                For Each ctlForm As Control In ctlMaster.Controls
                    If TypeOf ctlForm Is HtmlForm Then

                        For Each ctlSplitter As Control In ctlForm.Controls
                            If TypeOf ctlSplitter Is ASPxSplitter.ASPxSplitter Then

                                For Each ctlSplitterControl As Control In ctlSplitter.Controls
                                    If TypeOf ctlSplitterControl Is ASPxSplitter.Internal.SplitterControl Then

                                        For Each ctlSplitterContent As Control In ctlSplitterControl.Controls
                                            If TypeOf ctlSplitterContent Is ASPxClasses.Internal.InternalTableRow Then

                                                For Each ctlTblRow As Control In ctlSplitterContent.Controls
                                                    If TypeOf ctlTblRow Is ASPxSplitter.Internal.SplitterPaneCell Then

                                                        For Each ctlSplitterPane As Control In ctlTblRow.Controls
                                                            If TypeOf ctlSplitterPane Is ASPxSplitter.SplitterContentControl Then

                                                                For Each ctlSplitterContentControl As Control In ctlSplitterPane.Controls
                                                                    If TypeOf ctlSplitterContentControl Is ASPxRoundPanel.ASPxRoundPanel Then

                                                                        For Each ctlRoundPanel As Control In ctlSplitterContentControl.Controls
                                                                            If TypeOf ctlRoundPanel Is ASPxRoundPanel.Internal.RPRoundPanelControl Then

                                                                                For Each ctlRoundPanelControl As Control In ctlRoundPanel.Controls
                                                                                    If TypeOf ctlRoundPanelControl Is ASPxClasses.Internal.InternalTable Then

                                                                                        For Each ctlTbl As Control In ctlRoundPanelControl.Controls
                                                                                            If TypeOf ctlTbl Is ASPxClasses.Internal.InternalTableRow Then

                                                                                                For Each ctlTblRow2 As Control In ctlTbl.Controls
                                                                                                    If TypeOf ctlTblRow2 Is ASPxClasses.Internal.InternalTableCell Then

                                                                                                        For Each ctlTblCell As Control In ctlTblRow2.Controls
                                                                                                            If TypeOf ctlTblCell Is ASPxClasses.Internal.InternalTable Then

                                                                                                                For Each ctlTbl2 As Control In ctlTblCell.Controls
                                                                                                                    If TypeOf ctlTbl2 Is ASPxClasses.Internal.InternalTableRow Then

                                                                                                                        For Each ctlTblRow3 As Control In ctlTbl2.Controls
                                                                                                                            If TypeOf ctlTblRow3 Is ASPxClasses.Internal.InternalTableCell Then

                                                                                                                                For Each ctlTblCell2 As Control In ctlTblRow3.Controls
                                                                                                                                    If TypeOf ctlTblCell2 Is ASPxPanel.PanelContent Then

                                                                                                                                        For Each ctlPanelContent As Control In ctlTblCell2.Controls
                                                                                                                                            If TypeOf ctlPanelContent Is ContentPlaceHolder Then

                                                                                                                                                For Each ctlChild As Control In ctlPanelContent.Controls
                                                                                                                                                    'ASPxTextBox
                                                                                                                                                    If TypeOf (ctlChild) Is ASPxEditors.ASPxTextBox Then
                                                                                                                                                        CType(ctlChild, ASPxEditors.ASPxTextBox).Font.Name = "Verdana"
                                                                                                                                                        CType(ctlChild, ASPxEditors.ASPxTextBox).Font.Size = "8"
                                                                                                                                                        CType(ctlChild, ASPxEditors.ASPxTextBox).ForeColor = Color.Black
                                                                                                                                                        CType(ctlChild, ASPxEditors.ASPxTextBox).Height = 20
                                                                                                                                                    End If

                                                                                                                                                    'ASPxComboBox
                                                                                                                                                    If TypeOf (ctlChild) Is ASPxEditors.ASPxComboBox Then
                                                                                                                                                        CType(ctlChild, ASPxEditors.ASPxComboBox).Font.Name = "Verdana"
                                                                                                                                                        CType(ctlChild, ASPxEditors.ASPxComboBox).Font.Size = "8"
                                                                                                                                                        CType(ctlChild, ASPxEditors.ASPxComboBox).ForeColor = Color.Black
                                                                                                                                                        CType(ctlChild, ASPxEditors.ASPxComboBox).Height = 20
                                                                                                                                                        CType(ctlChild, ASPxEditors.ASPxComboBox).IncrementalFilteringMode = ASPxEditors.IncrementalFilteringMode.StartsWith
                                                                                                                                                    End If

                                                                                                                                                    'ASPxGridView
                                                                                                                                                    If TypeOf (ctlChild) Is ASPxGridView.ASPxGridView Then
                                                                                                                                                        With CType(ctlChild, ASPxGridView.ASPxGridView)
                                                                                                                                                            'Accessibility
                                                                                                                                                            .KeyboardSupport = True

                                                                                                                                                            'Settings
                                                                                                                                                            '-Settings
                                                                                                                                                            .Settings.ShowGroupButtons = False
                                                                                                                                                            .Settings.HorizontalScrollBarMode = pShowHorizontalScroll
                                                                                                                                                            .Settings.VerticalScrollBarMode = ASPxClasses.ScrollBarMode.Visible
                                                                                                                                                            .Settings.VerticalScrollableHeight = 240
                                                                                                                                                            '-SettingsBehaviour
                                                                                                                                                            .SettingsBehavior.AllowFocusedRow = pAllowFocusedRow
                                                                                                                                                            .SettingsBehavior.AllowSort = pAllowSort
                                                                                                                                                            .SettingsBehavior.ColumnResizeMode = ASPxClasses.ColumnResizeMode.Control
                                                                                                                                                            .SettingsBehavior.EnableRowHotTrack = True
                                                                                                                                                            '-SettingsPager
                                                                                                                                                            .SettingsPager.Mode = ASPxGridView.GridViewPagerMode.ShowAllRecords
                                                                                                                                                            .SettingsPager.PageSize = 1000
                                                                                                                                                            .SettingsPager.Visible = False

                                                                                                                                                            'Styles
                                                                                                                                                            '-Header
                                                                                                                                                            .Styles.Header.BackColor = Color.FromName("#FFD2A6")
                                                                                                                                                            .Styles.Header.Font.Name = "Verdana"
                                                                                                                                                            .Styles.Header.Font.Size = "8"
                                                                                                                                                            .Styles.Header.ForeColor = Color.Black
                                                                                                                                                            .Styles.Header.Wrap = Utils.DefaultBoolean.True
                                                                                                                                                            '-Row
                                                                                                                                                            .Styles.Row.BackColor = Color.FromName("#FFFFE1")
                                                                                                                                                            .Styles.Row.Font.Name = "Verdana"
                                                                                                                                                            .Styles.Row.Font.Size = "8"
                                                                                                                                                            .Styles.Row.Wrap = Utils.DefaultBoolean.False
                                                                                                                                                            '-RowHotTrack
                                                                                                                                                            .Styles.RowHotTrack.BackColor = Color.FromName("#E8EFFD")
                                                                                                                                                            .Styles.RowHotTrack.Font.Name = "Verdana"
                                                                                                                                                            .Styles.RowHotTrack.Font.Size = "8"
                                                                                                                                                            .Styles.RowHotTrack.ForeColor = Color.Black
                                                                                                                                                            .Styles.RowHotTrack.Wrap = Utils.DefaultBoolean.False
                                                                                                                                                            '-FocusedRow
                                                                                                                                                            .Styles.FocusedRow.BackColor = Color.FromName("#DCE7FC")
                                                                                                                                                            .Styles.FocusedRow.Font.Name = "Verdana"
                                                                                                                                                            .Styles.FocusedRow.Font.Size = "8"
                                                                                                                                                            .Styles.FocusedRow.ForeColor = Color.Black
                                                                                                                                                            .Styles.FocusedRow.Wrap = Utils.DefaultBoolean.False
                                                                                                                                                            '-SelectedRow
                                                                                                                                                            .Styles.SelectedRow.Wrap = Utils.DefaultBoolean.False
                                                                                                                                                        End With
                                                                                                                                                    End If
                                                                                                                                                Next ctlChild

                                                                                                                                            End If
                                                                                                                                        Next ctlPanelContent

                                                                                                                                    End If
                                                                                                                                Next ctlTblCell2

                                                                                                                            End If
                                                                                                                        Next ctlTblRow3

                                                                                                                    End If
                                                                                                                Next ctlTbl2

                                                                                                            End If
                                                                                                        Next ctlTblCell

                                                                                                    End If
                                                                                                Next ctlTblRow2

                                                                                            End If
                                                                                        Next ctlTbl

                                                                                    End If
                                                                                Next ctlRoundPanelControl

                                                                            End If
                                                                        Next ctlRoundPanel

                                                                    End If
                                                                Next ctlSplitterContentControl

                                                            End If
                                                        Next ctlSplitterPane

                                                    End If
                                                Next ctlTblRow

                                            End If
                                        Next ctlSplitterContent

                                    End If
                                Next ctlSplitterControl

                            End If
                        Next ctlSplitter

                    End If
                Next ctlForm

            End If
        Next ctlMaster

    End Sub

    Public Sub setAppearanceControlsDevEx13(ByRef pRefPage As Page, _
                                            Optional ByVal pShowHorizontalScroll As Boolean = False, _
                                            Optional ByVal pEnableRowHotTrack As Boolean = False, _
                                            Optional ByVal pAllowFocusedRow As Boolean = False, _
                                            Optional ByVal pAllowSort As Boolean = False, _
                                            Optional ByVal pFixedCol As Byte = 0, _
                                            Optional ByVal pAllowEditingNonFixedCol As Boolean = False, _
                                            Optional ByVal pPagerMode As PagerMode = PagerMode.ShowAllRecord, _
                                            Optional ByVal pSetFixedBgColor As Boolean = True, _
                                            Optional ByVal pSetNonFixedBgColor As Boolean = False, _
                                            Optional ByVal pSetLabel As Boolean = True, _
                                            Optional ByVal pSetTextBox As Boolean = True)
        For Each ctlMaster As Control In pRefPage.Controls
            If TypeOf ctlMaster Is MasterPage Then

                For Each ctlForm As Control In ctlMaster.Controls
                    If TypeOf ctlForm Is HtmlForm Then

                        For Each ctlSplitter As Control In ctlForm.Controls
                            If TypeOf ctlSplitter Is ASPxSplitter.ASPxSplitter Then

                                For Each ctlSplitterControl As Control In ctlSplitter.Controls
                                    If TypeOf ctlSplitterControl Is ASPxSplitter.Internal.SplitterControl Then

                                        For Each ctlSplitterContent As Control In ctlSplitterControl.Controls
                                            If TypeOf ctlSplitterContent Is ASPxClasses.Internal.InternalTableRow Then

                                                For Each ctlTblRow As Control In ctlSplitterContent.Controls
                                                    If TypeOf ctlTblRow Is ASPxSplitter.Internal.SplitterPaneCell Then

                                                        For Each ctlSplitterPane As Control In ctlTblRow.Controls
                                                            If TypeOf ctlSplitterPane Is ASPxSplitter.SplitterContentControl Then

                                                                For Each ctlSplitterContentControl As Control In ctlSplitterPane.Controls

                                                                    If TypeOf ctlSplitterContentControl Is ContentPlaceHolder Then

                                                                        For Each ctlChild As Control In ctlSplitterContentControl.Controls
                                                                            'ASPxLabel
                                                                            If pSetLabel = True Then
                                                                                If TypeOf (ctlChild) Is ASPxEditors.ASPxLabel Then
                                                                                    CType(ctlChild, ASPxEditors.ASPxLabel).Font.Name = "Verdana"
                                                                                    CType(ctlChild, ASPxEditors.ASPxLabel).Font.Size = "8"
                                                                                End If
                                                                            End If

                                                                            'ASPxTextBox
                                                                            If pSetTextBox = True Then
                                                                                If TypeOf (ctlChild) Is ASPxEditors.ASPxTextBox Then
                                                                                    CType(ctlChild, ASPxEditors.ASPxTextBox).Font.Name = "Verdana"
                                                                                    CType(ctlChild, ASPxEditors.ASPxTextBox).Font.Size = "8"
                                                                                    CType(ctlChild, ASPxEditors.ASPxTextBox).ForeColor = Color.Black
                                                                                    CType(ctlChild, ASPxEditors.ASPxTextBox).Height = 20
                                                                                End If
                                                                            End If

                                                                            'ASPxComboBox
                                                                            If TypeOf (ctlChild) Is ASPxEditors.ASPxComboBox Then
                                                                                CType(ctlChild, ASPxEditors.ASPxComboBox).Font.Name = "Verdana"
                                                                                CType(ctlChild, ASPxEditors.ASPxComboBox).Font.Size = "8"
                                                                                CType(ctlChild, ASPxEditors.ASPxComboBox).ForeColor = Color.Black
                                                                                CType(ctlChild, ASPxEditors.ASPxComboBox).Height = 20
                                                                                CType(ctlChild, ASPxEditors.ASPxComboBox).IncrementalFilteringMode = ASPxEditors.IncrementalFilteringMode.StartsWith
                                                                            End If

                                                                            'ASPxGridView
                                                                            If TypeOf (ctlChild) Is ASPxGridView.ASPxGridView Then
                                                                                With CType(ctlChild, ASPxGridView.ASPxGridView)
                                                                                    'Accessibility
                                                                                    .KeyboardSupport = True

                                                                                    'Settings
                                                                                    '-Settings
                                                                                    .Settings.ShowGroupButtons = False
                                                                                    .Settings.HorizontalScrollBarMode = pShowHorizontalScroll
                                                                                    .Settings.VerticalScrollBarMode = ASPxClasses.ScrollBarMode.Hidden
                                                                                    .Settings.ShowStatusBar = ASPxGridView.GridViewStatusBarMode.Hidden
                                                                                    '-SettingsBehaviour
                                                                                    .SettingsBehavior.AllowFocusedRow = pAllowFocusedRow
                                                                                    .SettingsBehavior.AllowSort = pAllowSort
                                                                                    .SettingsBehavior.ColumnResizeMode = ASPxClasses.ColumnResizeMode.Control
                                                                                    .SettingsBehavior.EnableRowHotTrack = pEnableRowHotTrack
                                                                                    '-SettingsPager
                                                                                    .SettingsPager.Mode = pPagerMode
                                                                                    If pPagerMode = PagerMode.ShowAllRecord Then .SettingsPager.Visible = False Else .SettingsPager.Visible = True : .SettingsPager.AlwaysShowPager = True

                                                                                    'SettingsText
                                                                                    .SettingsText.EmptyDataRow = " "

                                                                                    'SettingsLoadingPanel
                                                                                    .SettingsLoadingPanel.Text = "Please wait..."

                                                                                    'SettingsEditing
                                                                                    .SettingsEditing.NewItemRowPosition = ASPxGridView.GridViewNewItemRowPosition.Bottom
                                                                                    .SettingsEditing.BatchEditSettings.ShowConfirmOnLosingChanges = False

                                                                                    .Font.Name = "Verdana"
                                                                                    .Font.Size = "8"

                                                                                    'Styles
                                                                                    '-Header
                                                                                    .Styles.Header.BackColor = Color.FromName("#FFD2A6")
                                                                                    .Styles.Header.Font.Name = "Verdana"
                                                                                    .Styles.Header.Font.Size = "8"
                                                                                    .Styles.Header.ForeColor = Color.Black
                                                                                    .Styles.Header.Wrap = Utils.DefaultBoolean.True
                                                                                    '-Row
                                                                                    .Styles.Row.BackColor = Color.FromName("#FFFFE1")
                                                                                    .Styles.Row.Font.Name = "Verdana"
                                                                                    .Styles.Row.Font.Size = "8"
                                                                                    .Styles.Row.Wrap = Utils.DefaultBoolean.False
                                                                                    '-RowHotTrack
                                                                                    .Styles.RowHotTrack.BackColor = Color.FromName("#E8EFFD")
                                                                                    .Styles.RowHotTrack.Font.Name = "Verdana"
                                                                                    .Styles.RowHotTrack.Font.Size = "8"
                                                                                    .Styles.RowHotTrack.ForeColor = Color.Black
                                                                                    .Styles.RowHotTrack.Wrap = Utils.DefaultBoolean.False
                                                                                    '-FocusedRow
                                                                                    .Styles.FocusedRow.BackColor = Color.FromName("#DCE7FC")
                                                                                    .Styles.FocusedRow.Font.Name = "Verdana"
                                                                                    .Styles.FocusedRow.Font.Size = "8"
                                                                                    .Styles.FocusedRow.ForeColor = Color.Black
                                                                                    .Styles.FocusedRow.Wrap = Utils.DefaultBoolean.False
                                                                                    '-SelectedRow
                                                                                    .Styles.SelectedRow.Wrap = Utils.DefaultBoolean.False

                                                                                    If pFixedCol > 0 Then
                                                                                        If pAllowEditingNonFixedCol = True Then .Styles.Row.BackColor = Color.White
                                                                                        If pSetNonFixedBgColor = True Then .Styles.Row.BackColor = Color.FromName("#FFFFE1")

                                                                                        Dim iCol As Byte = 0
                                                                                        For iCol = 0 To pFixedCol - 1
                                                                                            .Columns(iCol).FixedStyle = ASPxGridView.GridViewColumnFixedStyle.Left
                                                                                            If pSetFixedBgColor = True Then .Columns(iCol).CellStyle.BackColor = Color.FromName("#FFFFE1")
                                                                                        Next iCol
                                                                                    End If
                                                                                End With

                                                                            End If
                                                                        Next ctlChild

                                                                    End If

                                                                Next ctlSplitterContentControl

                                                            End If
                                                        Next ctlSplitterPane

                                                    End If
                                                Next ctlTblRow

                                            End If
                                        Next ctlSplitterContent

                                    End If
                                Next ctlSplitterControl

                            End If
                        Next ctlSplitter

                    End If
                Next ctlForm

            End If
        Next ctlMaster

    End Sub
#End Region
End Class
