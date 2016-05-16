Attribute VB_Name = "Xls_Sqv"
Option Explicit

Function Sqv_GetDr_Base1(Sqv, R&) As Variant()
Dim U&, J%
U = UBound(Sqv, 2)
ReDim O(1 To U)
For J = 1 To U
    O(J) = Sqv(R, J)
Next
Sqv_GetDr_Base1 = O
End Function

Sub Sqv_PutDr(Sqv, R&, Dr())
Dim J%
For J = 1 To UBound(Dr) + 1
    Sqv(R, J) = Dr(J - 1)
Next
End Sub

Sub Sqv_Brw_NoHd(DtaSqv, Optional WsNm$ = "Data")
Dim Ws As Worksheet
Set Ws = Ws_New(WsNm)
Cell_PutSqv Ws.Range("A1"), DtaSqv
Ws_Wb(Ws).Activate
Ws.Activate
End Sub

Sub Sqv_Brw(HdSqv, DtaSqv, Optional WsNm$ = "Data")
Dim Ws As Worksheet
Set Ws = Ws_New(WsNm)
Cell_PutSqv Ws.Range("A1"), HdSqv
Cell_PutSqv Ws.Range("A2"), DtaSqv
Ws_Wb(Ws).Activate
Ws.Activate
End Sub

Function Sqv_TrimStr(Sqv)
Dim R&, C&
For R = 1 To UBound(Sqv, 1)
    For C = 1 To UBound(Sqv, 2)
        If VarType(Sqv(R, C)) = vbString Then Sqv(R, C) = Trim(Sqv(R, C))
    Next
Next
Sqv_TrimStr = Sqv
End Function

Sub Tst()
End Sub
