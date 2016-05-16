Attribute VB_Name = "Vb_Ay"
Option Explicit
Option Compare Text
Function Ay_Sqv(Ay)
Dim N&
N = Sz(Ay)
ReDim O(1 To 1, 1 To N)
Dim J&
For J = 0 To N - 1
    O(1, J + 1) = Ay(J)
Next
Ay_Sqv = O
End Function

Function Ay_SubSet_ByPfx(Ay, Pfx$) As String()
Dim O$(), J&
For J = 0 To UB(Ay)
    If IsPfx(Ay(J), Pfx) Then Push O, Ay(J)
Next
Ay_SubSet_ByPfx = O
End Function

Function Ay_Quote(Ay, Q$) As String()
Dim O$(), U&
U = UB(Ay)
If U = -1 Then Exit Function
ReDim O(U)
Dim J&
For J = 0 To U
    O(J) = Str_Quote(Ay(J), Q)
Next
Ay_Quote = O
End Function

Function Ay_RmvEleAt(Ay, At&)
Dim O, J&, U&, Rmv As Boolean
O = Ay
U = UB(O)
For J = At + 1 To U
    O(J - 1) = O(J)
    Rmv = True
Next
If Rmv Then ReDim Preserve O(U - 1)
Ay_RmvEleAt = O
End Function

Private Sub Ay_AskOne__Tst()
Dim Act$
Act = Ay_AskOne(Split("a"))
Debug.Assert Act = "a"

Act = Ay_AskOne(Split("a b c d"))
Debug.Assert Act = "a"
End Sub

Function Ay_AskOne(Ay, Optional Tit = "Select", Optional Msg$, Optional Dft$ = "")
If Not IsStrAy(Ay) Then Err.Raise 1, , "Ay must be string array"
If Sz(Ay) = 1 Then Ay_AskOne = Ay(0): Exit Function
Dim N%
Dim M$
    Dim A$()
    Dim J%
    N = Sz(Ay)
    ReDim A(N - 1)
    For J = 1 To N
        A(J - 1) = J & ". " & Ay(J - 1)
    Next
    M = Join(A, vbCrLf)
    If Msg <> "" Then M = Msg & vbCrLf & vbCrLf & M
Dim I%
Dim D$
    D = Ay_AskOne__DftNbr(Ay, Dft)
    I = Val(InputBox(M, Tit, D))
If 1 > I Or I > N Then Exit Function
Ay_AskOne = Ay(I - 1)
End Function

Private Function Ay_AskOne__DftNbr$(Ay, Dft$)
If Dft = "" Then Exit Function
Dim J%
For J = 0 To UB(Ay)
    If Ay(J) = Dft Then Ay_AskOne__DftNbr = J + 1: Exit Function
Next
End Function

Sub Ay_Dmp(Ay)
Dim J&
For J = 0 To UB(Ay)
    Debug.Print Ay(J)
Next
End Sub

Private Sub Ay_Brw__Tst()
Dim A$()
A = Split("A b C")
Ay_Brw A
End Sub

Sub Ay_Brw(Ay)
Dim Ft$
    Ft = Tmp_Ft("Ay_Brw")
Ft_WrtStr Ft, Join(Ay, vbCrLf)
Ft_Brw Ft
Kill Ft
End Sub

Function Ay_Idx&(Ay, Itm)
Dim J&
For J = 0 To UB(Ay)
    If Ay(J) = Itm Then Ay_Idx = J: Exit Function
Next
Ay_Idx = -1
End Function

Function Ay_HSqv(Ay)
Dim O, J&
ReDim O(1 To 1, 1 To Sz(Ay))
For J = 0 To UBound(Ay)
    O(1, J + 1) = Ay(J)
Next
Ay_HSqv = O
End Function

Function Ay_AddPfx(Ay, Pfx) As String()
Dim N&, J&
N = Sz(Ay)
If N = 0 Then Exit Function
ReDim O$(N - 1)
For J = 0 To N - 1
    O(J) = Pfx & Ay(J)
Next
Ay_AddPfx = O
End Function

Function Ay_AddSfx(Ay, Sfx) As String()
Dim N&, J&
N = Sz(Ay)
If N = 0 Then Exit Function
ReDim O$(N - 1)
For J = 0 To N - 1
    O(J) = Ay(J) & Sfx
Next
Ay_AddSfx = O
End Function

Sub PushAy_NoDup(O, Ay)
Dim J&
For J = 0 To UB(Ay)
    Push_NoDup O, Ay(J)
Next
End Sub

Sub Push_NoDup(Ay, Itm)
If Not Ay_Has(Ay, Itm) Then
    Push Ay, Itm: Exit Sub
End If
End Sub

Function Ay_Has(Ay, Itm) As Boolean
Dim J&
For J = 0 To UB(Ay)
    If Ay(J) = Itm Then Ay_Has = True: Exit Function
Next
End Function

Function Ay_IsEmpty(Ay) As Boolean
Ay_IsEmpty = Sz(Ay) = 0
End Function

Sub Push(Ay, Itm)
Dim N&
N = Sz(Ay)
ReDim Preserve Ay(N)
If IsObject(Itm) Then
    Set Ay(N) = Itm
Else
    Ay(N) = Itm
End If
End Sub

Sub PushAy(Ay, Ay1)
Dim J&
For J = 0 To UB(Ay1)
    Push Ay, Ay1(J)
Next
End Sub

Function Sz&(Ay)
If Not IsArray(Ay) Then Err.Raise 1
On Error Resume Next
Sz = UBound(Ay) + 1
End Function

Function UB&(Ay)
UB = Sz(Ay) - 1
End Function

Sub Tst()
End Sub
