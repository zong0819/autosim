Attribute VB_Name = "Module1"
Sub autosim()
Attribute autosim.VB_ProcData.VB_Invoke_Func = "r\n14"
'
' autosim ����
'
' �ֳt��: Ctrl+r
'
Dim fId As Integer '�Y�O���q��ƽг]Long
Dim bCnt As Integer
bCnt = CInt(InputBox("�п�J�ثe����t�Ϥu�@ï�ƶq"))
For fId = 1 To bCnt

Workbooks.Open Filename:=ThisWorkbook.Path & "\����" & fId & "�t.xlsx"

ActiveWorkbook.Sheets(1).Activate '�Ĥ@�i��Ұ�
'MsgBox ("���t�ϸ�Ʀ@" & ActiveSheet.UsedRange.Rows.Count & "��")
  
  '�бN���s�n�������K�W�b����U��
    SolverReset
    SolverOk SetCell:="$C$11", MaxMinVal:=1, ValueOf:=0, ByChange:="$F$4:$F$6", _
        Engine:=2, EngineDesc:="Simplex LP"
    SolverAdd CellRef:="$C$9", Relation:=1, FormulaText:="$C$7"
    SolverAdd CellRef:="$C$10", Relation:=1, FormulaText:="$C$8"
    SolverAdd CellRef:="$F$4", Relation:=3, FormulaText:="0"
    SolverAdd CellRef:="$F$5", Relation:=3, FormulaText:="0"
    SolverAdd CellRef:="$F$6", Relation:=3, FormulaText:="0"
    SolverAdd CellRef:="$F$4", Relation:=4, FormulaText:="���"
    SolverAdd CellRef:="$F$5", Relation:=4, FormulaText:="���"
    SolverAdd CellRef:="$F$6", Relation:=4, FormulaText:="���"
    SolverOk SetCell:="$C$11", MaxMinVal:=1, ValueOf:=0, ByChange:="$F$4:$F$6", _
        Engine:=2, EngineDesc:="Simplex LP"
    SolverOk SetCell:="$C$11", MaxMinVal:=1, ValueOf:=0, ByChange:="$F$4:$F$6", _
        Engine:=2, EngineDesc:="Simplex LP"
    SolverSolve

ActiveWorkbook.Save
ActiveWorkbook.Close

Next
End Sub

