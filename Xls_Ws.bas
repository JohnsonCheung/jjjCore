Attribute VB_Name = "Xls_Ws"
Option Explicit
Type TR1R2
    R1 As Long
    R2 As Long
End Type

Sub R1R2_Dmp(P As TR1R2)
Debug.Print "R1R2=(" & P.R1 & "," & P.R2 & ")"
End Sub

Function Ws_AssertSingleListObj(Ws As Worksheet) As Boolean
If Ws.ListObjects.Count = 1 Then Exit Function
Const C = "Worksheet[{0}] no ListObject.  The workbook is [{1}] which is in folder[{2}]"
Dim Wb As Workbook
Set Wb = Ws.Parent
MsgBox Fmt(C, Ws.Name, Wb.Name, Wb_Pth(Wb)), vbCritical
Ws_AssertSingleListObj = True
End Function

Function Ws_New(Optional WsNm$ = "Sheet1") As Worksheet
Dim O As Worksheet
Set O = Wb_New(WsNm).Sheets(1)
If O.CodeName = "" Then Stop
Set Ws_New = O
End Function

Function Ws_ListObj_R1R2(Ws As Worksheet) As TR1R2
Dim R As Range
Set R = Ws.ListObjects(1).DataBodyRange
'Stop
Ws_ListObj_R1R2 = Rge_R1R2(R)
End Function

Function Ws_Wb(Ws As Worksheet) As Workbook
Set Ws_Wb = Ws.Parent
End Function

Sub Ws_CrtListObj(Ws As Worksheet)
Dim Cell1 As Range, Cell2 As Range, Rge As Range
Set Cell1 = Ws.Cells(1, 1)
Set Cell2 = Ws.Cells.SpecialCells(xlCellTypeLastCell)
Set Rge = Ws.Range(Cell1, Cell2)
Ws.ListObjects.Add xlSrcRange, Rge, , xlYes
End Sub

Function Ws_RCRC(Ws As Worksheet, R1, C1, R2, C2) As Range
Dim Cell1 As Range, Cell2 As Range
Set Cell1 = Ws.Cells(R1, C1)
Set Cell2 = Ws.Cells(R2, C2)
Set Ws_RCRC = Ws.Range(Cell1, Cell2)
End Function

Sub Ws_CrtTbl_ByHdSqv_DtaSqv(Ws As Worksheet, HdSqv, DtaSqv)
Cell_PutSqv Ws.Range("A1"), HdSqv
Cell_PutSqv Ws.Range("A2"), DtaSqv
Ws_CrtListObj Ws
Ws.Columns.AutoFit
End Sub

Sub Ws_FmtCol_AsTxt(Ws As Worksheet, C, NRow&)
Ws_CRR(Ws, C, 2, NRow + 1).NumberFormat = "@"
End Sub

Sub Ws_Zoom(Ws As Worksheet, Zoom%)
Ws.Activate
ActiveWindow.Zoom = Zoom
End Sub

Function Ws_C(Ws As Worksheet, C) As Range
Set Ws_C = Ws_RC(Ws, 1, C).EntireColumn
End Function

Function Ws_Sqv_ByA1DownRight_NoR1(Ws As Worksheet)
Dim R&, C&
R = Ws.Range("A1").End(xlDown).Row
C = Ws.Range("A1").End(xlToRight).Column
Ws_Sqv_ByA1DownRight_NoR1 = Ws_RCRC(Ws, 2, 1, R, C).Value
End Function

Function Ws_A1_To_LastCell(Ws As Worksheet) As Range
Set Ws_A1_To_LastCell = Ws.Range(Ws.Cells(1, 1), Ws_LastCell(Ws))
End Function
Sub Ws_SetSummaryCol__Tst()
'Ws_SetSummaryCol ActiveSheet
Ws_SetSummaryCol ActiveSheet, xlSummaryOnRight
End Sub

Sub Ws_SetSummaryCol(Ws As Worksheet, Optional Where As XlSummaryColumn = XlSummaryColumn.xlSummaryOnLeft)
'To set SummaryCol left or right.
'Note: In order to do so, ActiveCell must outside a "ListObject".
'      So the pgm is required to move the ActiveCell to MaxCell
If Ws.ListObjects.Count = 0 Then
    Ws.OutLine.SummaryColumn = Where
    Exit Sub
End If

Dim App As Application
Dim ScnUpd As Boolean
Dim ActWb As Workbook
Dim ActWs As Worksheet
Dim ActCell As Range
Dim R As Range
    Set App = Ws.Application
    Set ActCell = App.ActiveCell
    Set ActWs = ActCell.Worksheet
    Set ActWb = ActWs.Parent
    ScnUpd = App.ScreenUpdating
    Set R = Ws_MaxCell(Ws)

App.ScreenUpdating = False
Ws.Activate
Dim ActCell1 As Range
    Set ActCell1 = ActiveCell
R.Activate
Ws.OutLine.SummaryColumn = Where    '<===
ActCell1.Activate
ActWb.Activate
ActWs.Activate
ActCell.Activate
App.ScreenUpdating = ScnUpd
End Sub

Function Ws_MaxCell(Ws As Worksheet) As Range
Set Ws_MaxCell = Ws.Cells(Ws_MaxRno(Ws), Ws_MaxCno(Ws))
End Function

Function Ws_CRR(Ws As Worksheet, C, R1, R2) As Range
Set Ws_CRR = Ws_RCRC(Ws, R1, C, R2, C)
End Function

Function Ws_RCC(Ws As Worksheet, R, C1, C2) As Range
Set Ws_RCC = Ws_RCRC(Ws, R, C1, R, C2)
End Function

Function Ws_R(Ws As Worksheet, R) As Range
Set Ws_R = Ws.Rows(R)
End Function

Function Ws_RC(Ws As Worksheet, R, C) As Range
Set Ws_RC = Ws.Cells(R, C)
End Function

Function Ws_Adr$(Ws As Worksheet, R, C, Optional IsAbsolute As Boolean)
Ws_Adr = Ws_RC(Ws, R, C).Address(IsAbsolute, IsAbsolute)
End Function

Function Ws_RR(Ws As Worksheet, R1, R2) As Range
Set Ws_RR = Ws_CRR(Ws, 1, R1, R2).EntireRow
End Function

Function Ws_CC(Ws As Worksheet, C1, C2) As Range
Set Ws_CC = Ws_RCC(Ws, 1, C1, C2).EntireColumn
End Function

Sub Ws_ClrOleObjs(Ws As Worksheet)
Dim J%
For J% = Ws.OLEObjects.Count To 1 Step -1
    Dim iOleObj As OLEObject: Set iOleObj = Ws.OLEObjects(J)
    iOleObj.Delete
Next
End Sub

Sub Ws_SetOutLine(Ws As Worksheet, Rno1%, Rno2%, Optional pLvl As Byte = 2)
Dim Rge As Range: Set Rge = Ws.Range(Ws.Cells(Rno1, 1), Ws.Cells(Rno2, 1))
Rge.EntireRow.OutlineLevel = pLvl
End Sub

Function Ws_Sqv_ByA1ToLastCell_withR1(Ws As Worksheet)
'Find Sqv of From Cell-A1 to Last-Cell
Dim C1 As Range
Dim C2 As Range
Set C1 = Ws.Cells(1, 1)
Set C2 = Ws_LastCell(Ws)
Ws_Sqv_ByA1ToLastCell_withR1 = Ws.Range(C1, C2).Value
End Function

Function Ws_LastRow&(Ws As Worksheet)
Ws_LastRow = Ws_LastCell(Ws).Row
End Function

Function Ws_LastCell(Ws As Worksheet) As Range
Set Ws_LastCell = Ws.Cells.SpecialCells(xlCellTypeLastCell)
End Function

Function Ws_LastCol&(Ws As Worksheet)
Ws_LastCol& = Ws_LastCell(Ws).Column
End Function

Function Ws_LastCol_ByFirstListObj&(Ws As Worksheet)
With Ws.ListObjects(1).DataBodyRange
    Ws_LastCol_ByFirstListObj = .Column + .Columns.Count - 1
End With
End Function

Sub Ws_ClrNames(Ws As Worksheet, pPfx$)
Dim J%
Dim L%: L = Len(pPfx)
For J = Ws.Names.Count To 1 Step -1
    Dim iNm As Name: Set iNm = Ws.Names(J)
    Dim mNm$: mNm = iNm.Name
    If Left(mNm, L) = pPfx Then iNm.Delete
Next
End Sub

Sub Ws_ShwAllDta(Ws As Worksheet)
Dim A As AutoFilter: Set A = Ws.AutoFilter
If TypeName(A) <> "Nothing" Then A.ShowAllData
End Sub

Sub Ws_Dlt(Ws As Worksheet)
Dim mXls As Application: Set mXls = Ws.Application
Dim mSave As Boolean: mSave = mXls.DisplayAlerts
mXls.DisplayAlerts = False
Ws.Delete
mXls.DisplayAlerts = mSave
End Sub

Function Ws_Sqv(Ws As Worksheet)
If IsNothing(Ws) Then Exit Function
Ws_Sqv = Ws.ListObjects(1).DataBodyRange.Value
End Function

Sub Ws_Sort(Ws As Worksheet, pLvcCol$, Optional Rno As Byte = 1)
'Col in pLvcCol can have minus sign as prefix means descending
Ws_ShwAllDta Ws
Dim mA$(): mA = Split(pLvcCol, ",")
Dim mRnoEnd&: mRnoEnd = Ws.Range("A" & Rno).End(xlDown).Row
Dim mColEnd$: mColEnd = Chr(64 + Ws.Range("A" & Rno).End(xlToRight).Column)
Dim J%
With Ws.Sort
    With .SortFields
        .Clear
        For J = 0 To UBound(mA)
            Dim mAA$
            Dim mOrd As XlSortOrder
            If Right(mA(J), 1) = "-" Then
                mOrd = xlDescending
                mAA = Left(mA(J), Len(mA(J)) - 1)
            Else
                mOrd = xlAscending
                mAA = mA(J)
            End If
            Dim mAdr$: mAdr = mAA & Rno & ":" & mAA & mRnoEnd
            .Add Key:=Ws.Range(mAdr), Order:=mOrd
        Next
    End With
    mAdr = "A" & Rno & ":" & mColEnd & mRnoEnd
    .SetRange Ws.Range(mAdr)
    .Header = xlYes
    .MatchCase = False
    .Orientation = xlTopToBottom
    .Apply                              '<== Sort the worksheet
End With
End Sub

Function Ws_MaxRno&(Ws As Worksheet)
Ws_MaxRno = Ws.Rows.Count
End Function
Function Ws_MaxCno&(Ws As Worksheet)
Ws_MaxCno = Ws.Columns.Count
End Function


Sub Tst()
End Sub
