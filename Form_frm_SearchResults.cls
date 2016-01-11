VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "Form_frm_SearchResults"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Compare Database

Private Sub btn_First_Click()

DoCmd.GoToRecord , , acFirst

End Sub

Private Sub btn_Last_Click()

DoCmd.GoToRecord , , acLast

End Sub

Private Sub btn_Next_Click()

DoCmd.GoToRecord , , acNext


End Sub

Private Sub btn_Previous_Click()

DoCmd.GoToRecord , , acPrevious

End Sub



Private Sub btn_Export_Click()

'CurrentDb.Execute "SELECT * FROM tbl_SearchResults"
DoCmd.RunCommand acCmdSelectAllRecords
DoCmd.RunCommand acCmdCopy

Dim xl As Excel.Application
Dim xlApp As Excel.Application
Dim xlwkbk As Excel.Workbook
Dim xlsheet As Excel.Worksheet


'We need to set xl to become an excel application
Set xl = New Excel.Application
Set xlwkbk = xl.Workbooks.Add

'Make the instance of Excel visible

xl.Visible = True



    ActiveSheet.PasteSpecial Format:="HTML", Link:=False, DisplayAsIcon:= _
        False, NoHTMLFormatting:=True
    Rows("1:1").Select
    Selection.Delete Shift:=xlUp
    Columns("F:F").Select
    Selection.Replace What:="==> ", Replacement:="~", LookAt:=xlPart, _
        SearchOrder:=xlByRows, MatchCase:=False, SearchFormat:=False, _
        ReplaceFormat:=False
        Columns("F:F").Select
    Selection.Replace What:="==>", Replacement:="", LookAt:=xlPart, _
        SearchOrder:=xlByRows, MatchCase:=False, SearchFormat:=False, _
        ReplaceFormat:=False
    Sheets("Sheet2").Select
    ActiveWindow.SelectedSheets.Delete
    Sheets("Sheet3").Select
    ActiveWindow.SelectedSheets.Delete
    Sheets("Sheet1").Select
    Sheets("Sheet1").Name = "Report"
    Range("G1").Select
    Columns("F:F").Select
    Selection.TextToColumns Destination:=Range("F1"), DataType:=xlDelimited, _
        TextQualifier:=xlDoubleQuote, ConsecutiveDelimiter:=False, Tab:=False, _
        Semicolon:=False, Comma:=False, Space:=False, Other:=True, OtherChar _
        :="~", FieldInfo:=Array(Array(1, 1), Array(2, 1)), TrailingMinusNumbers:=True
    Range("G1").Select
    ActiveCell.FormulaR1C1 = "Desc_value"
    Columns("H:H").Select
    Selection.Delete Shift:=xlToLeft
    Columns("A:E").Select
    Selection.SpecialCells(xlCellTypeConstants, 23).Select
    With Selection.Interior
        .Pattern = xlSolid
        .PatternColorIndex = xlAutomatic
        .ThemeColor = xlThemeColorLight1
        .TintAndShade = 0
        .PatternTintAndShade = 0
    End With
    With Selection.Font
        .ThemeColor = xlThemeColorDark1
        .TintAndShade = 0
    End With
    Selection.Borders(xlEdgeBottom).LineStyle = xlNone
    Selection.Borders(xlEdgeRight).LineStyle = xlNone
    Selection.Borders(xlInsideVertical).LineStyle = xlNone
    Selection.Borders(xlInsideHorizontal).LineStyle = xlNone
    ActiveWindow.SmallScroll Down:=-9
    Columns("A:E").Select
    Selection.SpecialCells(xlCellTypeBlanks).Select
    Selection.FormulaR1C1 = "=R[-1]C"
    Columns("A:G").Select
    Selection.AutoFilter
    Range("A1:G1").Select
    Selection.Font.Bold = True
    Columns("F:F").Select
    Selection.Font.Bold = False
    Selection.Font.Bold = True
    Selection.ColumnWidth = 25.14
    Columns("G:G").ColumnWidth = 39.86
    Columns("G:G").Select
    With Selection
        .HorizontalAlignment = xlLeft
        .VerticalAlignment = xlBottom
        .WrapText = False
        .Orientation = 0
        .AddIndent = False
        .IndentLevel = 0
        .ShrinkToFit = False
        .ReadingOrder = xlContext
        .MergeCells = False
    End With
    Range("A1:G1").Select
    Range(Selection, Selection.End(xlDown)).Select
    Selection.Copy

    Range("A1").Select
    Selection.PasteSpecial Paste:=xlPasteValues, Operation:=xlNone, SkipBlanks _
        :=False, Transpose:=False
        Range("A1").Select
    
    Range("F1").Select
    ActiveCell.FormulaR1C1 = "mem_Desc"
    Range("A1:G1").Select
    With Selection.Font
        .ThemeColor = xlThemeColorDark1
        .TintAndShade = 0
    End With
    With Selection.Interior
        .Pattern = xlSolid
        .PatternColorIndex = xlAutomatic
        .ThemeColor = xlThemeColorLight1
        .TintAndShade = 0
        .PatternTintAndShade = 0
    End With
    
    Columns("A:A").Select
    Selection.EntireColumn.Hidden = True
    Columns("D:E").Select
    Selection.EntireColumn.Hidden = True
    Dim answer As Integer
    Dim p As Integer
    Dim MyRange As Range
    Dim TotalRecords As Integer
    TotalRecords = WorksheetFunction.Max(Range("A:A")) - WorksheetFunction.Max(Range("A:A")) + 1
    
    answer = MsgBox("Do you want to remove Duplicate Entries?", vbYesNo + vbQuestion, "Remove Duplicates")
   
                If answer = vbYes Then
                Columns("A:G").Select
                ActiveSheet.Range("A:G").RemoveDuplicates Columns:=Array(2, 3, 6, 7), _
                Header:=xlYes
                Else
                'Do nothing
                
                End If
                
                MsgBox ("Data Exported Successfully!")

End Sub

 Function CountUnique(ByVal MyRange As Range) As Integer
                    Dim Cell As Range
                    Dim j As Integer
                    Dim iNumCells As Integer
                    Dim iUVals As Integer
                    Dim sUCells() As String
                
                    iNumCells = MyRange.Count
                    ReDim sUCells(iNumCells) As String
                
                    iUVals = 0
                    For Each Cell In MyRange
                        If Cell.Text > "" Then
                            For j = 1 To iUVals
                                If sUCells(j) = Cell.Text Then
                                    Exit For
                                End If
                            Next j
                            If j > iUVals Then
                                iUVals = iUVals + 1
                                sUCells(iUVals) = Cell.Text
                            End If
                        End If
                    Next Cell
                    CountUnique = iUVals
                End Function

