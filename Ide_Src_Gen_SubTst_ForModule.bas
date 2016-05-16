Attribute VB_Name = "Ide_Src_Gen_SubTst_ForModule"
Option Explicit

Sub Src_Gen_SubTst_ForModule()
Dim Ay() As CodeModule
    Ay = ZMdAy_NoSubTst
If ZSz(Ay) = 0 Then Exit Sub

Dim I, Md As CodeModule
For Each I In Ay
    Set Md = I
    ZAddSubTst Md
Next
End Sub

Private Sub ZAddSubTst(Md As CodeModule)
Const Lines = vbCrLf & "Sub Tst()" & vbCrLf & "End Sub"
Md.InsertLines Md.CountOfLines + 1, Lines
Debug.Print "Inserted: "; Md.Parent.Name
End Sub

Private Sub ZMdAy_NoSubTst__Tst()
Dim A() As CodeModule
    A = ZMdAy_NoSubTst
If ZSz(A) > 0 Then
    Dim O$()
    Dim J&
    ReDim O(UBound(A))
    For J = 0 To UBound(A)
        O(J) = A(J).Parent.Name
    Next
    Ay_Brw O
End If
End Sub

Private Function ZSz&(Ay)
On Error Resume Next
ZSz = UBound(Ay) + 1
End Function

Private Sub ZPush(Ay, I)
Dim N&
N = ZSz(Ay)
ReDim Preserve Ay(N)
If IsObject(I) Then
    Set Ay(N) = I
Else
    Ay(N) = I
End If
End Sub

Private Property Get ZMdAy_NoSubTst() As CodeModule()
Dim O() As CodeModule
    Dim I, Cmp As VBComponent, Md As CodeModule
    For Each I In ZCurPj.VBComponents
        Set Cmp = I
        If Cmp.Type = vbext_ct_StdModule Then
            If Cmp.Name <> "Ide_Src_Gen_SubTst_ForModule" Then
                Set Md = Cmp.CodeModule
                If ZIsNoSubTst(Md) Then ZPush O, Md
            End If
        End If
    Next
ZMdAy_NoSubTst = O
End Property

Private Property Get ZIsNoSubTst(Md As CodeModule) As Boolean
Dim L$
    L = Md.Lines(1, Md.CountOfLines)
Dim P&
    P = InStr(L, "Sub Tst()" & vbCrLf)
ZIsNoSubTst = P = 0
End Property

Private Property Get ZCurPj() As VBProject
Set ZCurPj = Application.VBE.ActiveVBProject
End Property

Sub Tst()
End Sub
