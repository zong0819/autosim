Attribute VB_Name = "Module1"
Sub autosim()
Attribute autosim.VB_ProcData.VB_Invoke_Func = "r\n14"
'
' autosim 巨集
'
' 快速鍵: Ctrl+r
'
Dim fId As Integer '若是巨量資料請設Long
Dim bCnt As Integer
bCnt = CInt(InputBox("請輸入目前內科廠區工作簿數量"))
For fId = 1 To bCnt

Workbooks.Open Filename:=ThisWorkbook.Path & "\內科" & fId & "廠.xlsx"

ActiveWorkbook.Sheets(1).Activate '第一張表啟動
'MsgBox ("此廠區資料共" & ActiveSheet.UsedRange.Rows.Count & "筆")
  
  '請將錄製好的巨集貼上在本行下方
    SolverReset
    SolverOk SetCell:="$C$11", MaxMinVal:=1, ValueOf:=0, ByChange:="$F$4:$F$6", _
        Engine:=2, EngineDesc:="Simplex LP"
    SolverAdd CellRef:="$C$9", Relation:=1, FormulaText:="$C$7"
    SolverAdd CellRef:="$C$10", Relation:=1, FormulaText:="$C$8"
    SolverAdd CellRef:="$F$4", Relation:=3, FormulaText:="0"
    SolverAdd CellRef:="$F$5", Relation:=3, FormulaText:="0"
    SolverAdd CellRef:="$F$6", Relation:=3, FormulaText:="0"
    SolverAdd CellRef:="$F$4", Relation:=4, FormulaText:="整數"
    SolverAdd CellRef:="$F$5", Relation:=4, FormulaText:="整數"
    SolverAdd CellRef:="$F$6", Relation:=4, FormulaText:="整數"
    SolverOk SetCell:="$C$11", MaxMinVal:=1, ValueOf:=0, ByChange:="$F$4:$F$6", _
        Engine:=2, EngineDesc:="Simplex LP"
    SolverOk SetCell:="$C$11", MaxMinVal:=1, ValueOf:=0, ByChange:="$F$4:$F$6", _
        Engine:=2, EngineDesc:="Simplex LP"
    SolverSolve

ActiveWorkbook.Save
ActiveWorkbook.Close

Next
End Sub

