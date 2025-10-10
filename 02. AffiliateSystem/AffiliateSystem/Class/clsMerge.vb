Imports Microsoft.VisualBasic
Imports System
Imports System.Web
Imports System.Web.UI
Imports System.Web.UI.WebControls
Imports DevExpress.Web.ASPxGridView
Imports System.Collections.Generic

Public Class ASPxGridViewCellMerger
    Dim col1 As String
    Private grid_Renamed As ASPxGridView
    Private mergedCells As New Dictionary(Of GridViewDataColumn, TableCell)()
    Private cellRowSpans As New Dictionary(Of TableCell, Integer)()

    Public Sub New(ByVal grid As ASPxGridView, ByVal pColumnName As String)
        Me.grid_Renamed = grid
        AddHandler Me.Grid.HtmlRowCreated, AddressOf grid_HtmlRowCreated
        AddHandler Me.Grid.HtmlDataCellPrepared, AddressOf grid_HtmlDataCellPrepared
        Me.col1 = pColumnName
    End Sub

    Public ReadOnly Property Grid() As ASPxGridView
        Get
            Return grid_Renamed
        End Get
    End Property

    Private Sub grid_HtmlDataCellPrepared(ByVal sender As Object, ByVal e As ASPxGridViewTableDataCellEventArgs)
        'add the attribute that will be used to find which column the cell belongs to
        e.Cell.Attributes.Add("ci", e.DataColumn.VisibleIndex.ToString())

        If cellRowSpans.ContainsKey(e.Cell) Then
            e.Cell.RowSpan = cellRowSpans(e.Cell)
        End If
    End Sub

    Private Sub grid_HtmlRowCreated(ByVal sender As Object, ByVal e As ASPxGridViewTableRowEventArgs)
        If Grid.GetRowLevel(e.VisibleIndex) <> Grid.GroupCount Then
            Return
        End If
        For i As Integer = e.Row.Cells.Count - 1 To 0 Step -1
            Dim dataCell As DevExpress.Web.ASPxGridView.Rendering.GridViewTableDataCell = TryCast(e.Row.Cells(i), DevExpress.Web.ASPxGridView.Rendering.GridViewTableDataCell)
            If dataCell IsNot Nothing Then
                MergeCells(dataCell.DataColumn, e.VisibleIndex, dataCell, col1)
            End If
        Next i
    End Sub

    Private Sub MergeCells(ByVal column As GridViewDataColumn, ByVal visibleIndex As Integer, ByVal cell As TableCell, ByVal pColumn As String)
        Dim isNextTheSame As Boolean = IsNextRowHasSameData(column, visibleIndex)
        Dim iCols As Integer
        Dim BolMerge As Boolean = False

        'MsgBox(column.FieldName)

        For icols = 0 To UBound(Split(pColumn, ","))
            If column.FieldName = Split(pColumn, ",")(iCols) Then
                BolMerge = True
                Exit For
            End If
        Next
        If BolMerge Then
            If isNextTheSame Then
                If (Not mergedCells.ContainsKey(column)) Then
                    mergedCells(column) = cell
                End If
            End If
            If IsPrevRowHasSameData(column, visibleIndex) Then
                CType(cell.Parent, TableRow).Cells.Remove(cell)
                If mergedCells.ContainsKey(column) Then
                    Dim mergedCell As TableCell = mergedCells(column)
                    If (Not cellRowSpans.ContainsKey(mergedCell)) Then
                        cellRowSpans(mergedCell) = 1
                    End If
                    cellRowSpans(mergedCell) = cellRowSpans(mergedCell) + 1
                End If
            End If
            If (Not isNextTheSame) Then
                mergedCells.Remove(column)
            End If
        End If

    End Sub

    Private Function IsNextRowHasSameData(ByVal column As GridViewDataColumn, ByVal visibleIndex As Integer) As Boolean
        'is it the last visible row
        If visibleIndex >= Grid.VisibleRowCount - 1 Then
            Return False
        End If

        Return IsSameData(column.FieldName, visibleIndex, visibleIndex + 1)
    End Function

    Private Function IsPrevRowHasSameData(ByVal column As GridViewDataColumn, ByVal visibleIndex As Integer) As Boolean
        Dim grid As ASPxGridView = column.Grid
        'is it the first visible row
        If visibleIndex <= Me.Grid.VisibleStartIndex Then
            Return False
        End If

        Return IsSameData(column.FieldName, visibleIndex, visibleIndex - 1)
    End Function

    Private Function IsSameData(ByVal fieldName As String, ByVal visibleIndex1 As Integer, ByVal visibleIndex2 As Integer) As Boolean
        ' is it a group row?
        If Grid.GetRowLevel(visibleIndex2) <> Grid.GroupCount Then
            Return False
        End If

        Return Object.Equals(Grid.GetRowValues(visibleIndex1, fieldName), Grid.GetRowValues(visibleIndex2, fieldName))
    End Function
End Class

