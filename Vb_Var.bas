Attribute VB_Name = "Vb_Var"
Option Explicit

Function Var_IsNothing(V) As Boolean
Var_IsNothing = TypeName(V) = "Nothing"
End Function
Private Sub IsStrAy__Tst()
Dim A$()
Dim B
Dim C%()
Debug.Assert IsStrAy(A)
Debug.Assert Not IsStrAy(B)
Debug.Assert Not IsStrAy(C)
End Sub
Function IsStrAy(Ay) As Boolean
Dim T&
T = VarType(Ay)
IsStrAy = T = vbArray + vbString
End Function

Sub Tst()
End Sub
