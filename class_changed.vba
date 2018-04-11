Open Explicit

Sub class_changed()

	Dim thisFile As String
	Dim newFile As String
	Dim createdFile As String
	Dim cWB As Workbook
	Dim cWS As Worksheet
    Dim oWS AS Worksheet
    Dim nWS As Worksheet
 	Dim mWB As Workbook

	Set thisFile = "macro_test.xlsm"
 	Set mWb = Workbooks(thisfile)
 	Set newFile = Cells(10, 4).Value

	Set createdFile = "【差分】" & newFile
    Set cWB = Workbooks(createdFile)
    Set cWS = cWB.Worksheets("差分")
    Set oWS = cWB.Worksheets("旧ファイル")
    Set nWS = cWB.Worksheets("新ファイル")

    nWS.Activate

    Cells.Select
    Selection.Find(What:="分類", After:=ActiveCell, LookIn:=xlValues, lookat_
      :=xlWhole, SearchOrder:=xlByRows, SearchDirection:=xlNext, MatchCase:=_
      False, MatchByte:=False).Activate

    ActiveCell.EntireColumn.Select
    Selection.Insert Shift:=xlToRight

    ActiveCell.Offset(0,1).Select


End Sub

Sub Index_Match(WS1 AS Worksheet, WS2 As Worksheet, subject As String)
    'WS1は比較するシート WS2は比較されるシート　subjectは基準
    WS1.Activate
    Dim mat, ind
    Dim I As Integer, rng As Integer
    Dim sch_value As Range
    
    Call search_cell(subject) '基準になるセル
    
    Set sch_value = ActiveCell.Offset(1,0) '基準セル下の住所を保存する
    sch_value.Select
    Range(sch_value, sch_value.End(xlToLeft)).Select

    For I = sch_value.Row To WS1.UsedRange.Rows.Count Step 1
        mat = Application.match(Cells(I, sch_value.Column + 1), WS2.Range(sch_value.Offset(0,1).EntireColumn.Address), 0)
        ind = Application.Index(WS2.Range("a1", WS2.Cells.SpecialCells(xlCellTypeLastCell)), mat, Range(sch_value, sch_value.End(xlToLeft)).Columns.Count +1)
        Cells(I, sch_value.Column).FormulaR1C1 = ind
    Next

End Sub
