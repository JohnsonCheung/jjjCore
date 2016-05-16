Attribute VB_Name = "Xls_Fn"
Option Explicit
Private Sub Fn_Cno_Col__Tst()
Debug.Assert Fn_Cno_Col(1) = "A"
Debug.Assert Fn_Cno_Col(26) = "Z"
Debug.Assert Fn_Cno_Col(27) = "AA"
Debug.Assert Fn_Cno_Col(123) = "DS"
Debug.Assert Fn_Cno_Col(702) = "ZZ"
Debug.Assert Fn_Cno_Col(703) = "AAA"
Debug.Assert Fn_Cno_Col(16384) = "XFD"
Pass "Fn_Cno_Col"
End Sub
Sub Tst()
Debug.Print "Xls_Fn"
Fn_Cno_Col__Tst
End Sub
Property Get Fn_Cno_Col$(Cno%)
If 1 > Cno Or Cno > 16384 Then Err.Raise 1, , "Fn_Cno_Col: Cno[" & Cno & "] must between 1 and 16384"
Dim N1%, N2%, N3%, O$, C%
C = Cno - 1
Select Case Cno
Case Is <= 26:
    N1 = C + 1
    O = Chr(N1 + 64)
Case Is <= 702      ' 702=26*26+26
    N1 = C \ 26
    N2 = C Mod 26 + 1
    O = Chr(N1 + 64) & Chr(N2 + 64)
Case Else
    N1 = C \ 676        ' 676 = 26*26
    C = C - N1 * 676
    N2 = C \ 26
    N3 = C Mod 26 + 1
    O = Chr(N1 + 64) & Chr(N2 + 64) & Chr(N3 + 64)
End Select
Fn_Cno_Col = O
End Property

Property Get Fn_Col_Cno%(Col$)
Dim N2%, N3%, C$, N1%, O%
C = UCase(Col)
N1 = Asc(C) - 64
Select Case Len(Col)
Case 1
    O = N1
Case 2
    N2 = Asc(Mid(C, 2, 1)) - 64
    O = N1 * 26 + N2
Case 3:
    N2 = Asc(Mid(C, 2, 1)) - 64
    N3 = Asc(Mid(C, 3, 1)) - 64
    O = N1 * 26 * 26 + N2 * 26 + N3
Case Else: Err.Raise 1, , "Fn_Col_Cno: Col[" & Col & "] has len=[" & Len(Col) & "].  Expected=1..3"
End Select
If _
    0 > N1 Or N1 > 26 Or _
    0 > N2 Or N2 > 26 Or _
    0 > N3 Or N3 > 26 Then
    Err.Raise 1, , "Fn_Col_Cno: Col[" & Col & "] must be all letters."
End If
If O > 16384 Then
    Err.Raise 1, , "Fn_Col_Cno: Col[" & Col & "] must be less then [XFD]."
End If
Fn_Col_Cno = O
End Property

Function Fn_RC_Adr$(R&, C%)
Fn_RC_Adr = Fn_Cno_Col(C) & R
End Function
