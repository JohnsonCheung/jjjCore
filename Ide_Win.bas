Attribute VB_Name = "Ide_Win"
Option Explicit

Sub Win_CloseAll_ShwLcl_n_ShwImm()
Dim I, W As VBIDE.Window
For Each I In Win_Ay
    Set W = I
    Select Case W.Caption
    Case "Locals", "Immediate"
        If Not W.Visible Then W.Visible = True
    Case Else
        W.Close
    End Select
Next
End Sub
Sub Win_Shw_Imm()
Win_Imm.Visible = True
End Sub
Sub Win_CloseAll()
Dim I
For Each I In Win_Ay
    I.Close
Next
End Sub

Property Get Win_Lcl() As VBIDE.Window
Dim A() As VBIDE.Window, J%
A = Win_Ay
For J = 0 To UBound(A)
    If A(J).Caption = "Locals" Then Set Win_Lcl = A(J): Exit Property
Next
End Property

Property Get Win_Imm() As VBIDE.Window
Dim A() As VBIDE.Window, J%
A = Win_Ay
For J = 0 To UBound(A)
    If A(J).Caption = "Immediate" Then Set Win_Imm = A(J): Exit Property
Next
End Property

Sub Win_Shw_Lcl()
Win_Lcl.Visible = True
End Sub
Property Get Win_Ay() As VBIDE.Window()
Dim N%
N = Application.VBE.Windows.Count
Dim O() As VBIDE.Window
ReDim O(N - 1)
Dim J%
For J = 0 To N - 1
    Set O(J) = Application.VBE.Windows(J + 1)
Next
Win_Ay = O
End Property

Sub Tst()
End Sub
