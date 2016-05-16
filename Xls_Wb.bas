Attribute VB_Name = "Xls_Wb"
Option Explicit

Function Wb_Ws(Wb As Workbook, WsNm$) As Worksheet
Set Wb_Ws = Wb.Sheets(WsNm)
End Function

Sub Wb_Sav(Wb As Workbook)
If Wb.Saved Then Exit Sub
Dim App As Application
    Set App = Wb.Application
Dim A As Boolean
    A = App.DisplayAlerts
App.DisplayAlerts = False
Wb.Save
App.DisplayAlerts = A
End Sub

Sub Wb_AssertWsExist(Wb As Workbook, WsNm$)
If Wb_IsWs(Wb, WsNm) Then Exit Sub
Const C = "Worksheet[{0}] not found in" & vbLf & "Workbook[{1}] which is in" & vbLf & "folder[{2}]"
MsgBox Fmt(C, WsNm, Wb.Name, Wb_Pth(Wb)), vbCritical
End Sub

Function Wb_Lik(LikStr$, Optional OCnt%) As Workbook
OCnt = 0
Dim Wb As Workbook
For Each Wb In Workbooks
    If Wb.Name Like LikStr Then
        OCnt = OCnt + 1
        Set Wb_Lik = Wb
    End If
Next
End Function

Function Wb_New(Optional WsNm$ = "Sheet1") As Workbook
Dim O As Workbook, Ws As Worksheet
Set O = Workbooks.Add
If Not Wb_IsWs(O, WsNm) Then Set Ws = O.Sheets(1): Ws.Name = WsNm
Dim IWsNm
For Each IWsNm In Wb_WsNmAy(O)
    If IWsNm <> WsNm Then Wb_DltWs O, CStr(IWsNm)
Next
Set Wb_New = O
End Function

Private Sub Wb_WsNmAy__Tst()
Dim Act$()
    Dim Wb As Workbook
    Set Wb = Wb_New("AA")
    Act = Wb_WsNmAy(Wb)
    Wb.Close False
Debug.Assert Sz(Act) = 1
Debug.Assert Act(0) = "AA"
Pass "Wb_WsNmAy"
End Sub

Function Wb_WsNmAy(Wb As Workbook) As String()
Dim O$()
Dim iWs As Worksheet
For Each iWs In Wb.Sheets
    Push O, iWs.Name
Next
Wb_WsNmAy = O
End Function

Sub Wb_DltWs(Wb As Workbook, WsNm$)
Dim J%
For J% = 1 To Wb.Sheets.Count
    Dim iWs As Worksheet: Set iWs = Wb.Sheets(J)
    If iWs.Name = WsNm Then Ws_Dlt iWs: Exit Sub
Next
End Sub
Function Wb_IsWs(Wb As Workbook, WsNm$) As Boolean
Dim Ws As Worksheet
For Each Ws In Wb.Sheets
    If Ws.Name = WsNm Then Wb_IsWs = True: Exit Function
Next
End Function
Sub Wb_Hid_WsNmAy(Wb As Workbook, WsNmAy$())
Dim Ws As Worksheet, J%
For J = 0 To UB(WsNmAy)
    If Wb_IsWs(Wb, WsNmAy(J)) Then
        Set Ws = Wb.Sheets(WsNmAy(J))
        Ws.Visible = xlSheetHidden
    End If
Next
End Sub

Function Wb_AddName(Wb As Workbook, Nm$, Rge As Range) As Name
Dim WsNm$
    WsNm = Rge_Ws(Rge).Name
Dim Formula$
    Formula = Fmt_QQ("='?'!?", WsNm, Rge.Address)
Set Wb_AddName = Wb.Names.Add(Nm, Formula)
End Function

Function Wb_AddWs_AtEnd(Wb As Workbook, WsNm$, Optional DltBefAdd As Boolean) As Worksheet
If DltBefAdd Then Wb_DltWs Wb, WsNm
Dim O As Worksheet
Set O = Wb.Worksheets.Add(, Wb.Sheets(Wb.Sheets.Count))
O.Name = WsNm
Debug.Print Wb.Application.VBE.ActiveVBProject.VBComponents.Count
If O.CodeName = "" Then
    Dim A%
    A = Wb.Application.VBE.ActiveVBProject.VBComponents.Count
    If O.CodeName = "" Then Stop
End If
Set Wb_AddWs_AtEnd = O
End Function

Function Wb_Pth$(Wb As Workbook)
Wb_Pth = Ffn_Pth(Wb.FullName)
End Function

Sub Wb_ClrNames(Wb As Workbook, pPfx$)
Dim J%
Dim L%: L = Len(pPfx)
For J = Wb.Names.Count To 1 Step -1
    Dim iNm As Name: Set iNm = Wb.Names(J)
    Dim mNm$: mNm = iNm.Name
    Dim mNmX$
    Dim mP%: mP = InStr(mNm, "!")
    If mP > 0 Then
        mNmX = Mid(mNm, mP + 1)
    Else
        mNmX = mNm
    End If
    If pPfx = "" Then
        iNm.Delete
    Else
        If Left(mNmX, L) = pPfx Then iNm.Delete
    End If
Next
End Sub


Sub Tst()
End Sub
