Attribute VB_Name = "Vb_Fmt"
Option Explicit
Const CurMdNm$ = "FctFmt"

Function Fmt(FmtStr$, ParamArray Ap())
Dim Av(), I, O$, J%, A$
Av = Ap
O = FmtStr
For Each I In Av
    A = Quote(J, "{}"): J = J + 1
    O = Replace(O, A, I)
Next
Fmt = O
End Function

Function Fmt_QQ(QQ$, ParamArray Ap())
Dim Av(), I, O$
Av = Ap
O = QQ
For Each I In Av
    O = Replace(O, "?", CStr(I), Count:=1)
Next
Fmt_QQ = O
End Function

Sub Tst()
End Sub
