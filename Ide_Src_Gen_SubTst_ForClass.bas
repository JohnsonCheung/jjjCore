Attribute VB_Name = "Ide_Src_Gen_SubTst_ForClass"
Option Explicit

Sub Src_Gen_SubTst_ForClass()
Dim Ay() As CodeModule
    Ay = ZMdAy_NoFriendSubTst
If ZSz(Ay) = 0 Then Exit Sub

Dim I, Md As CodeModule
For Each I In Ay
    Set Md = I
    ZAddFriendSubTst Md
Next
End Sub
Private Sub ZAddFriendSubTst(Md As CodeModule)
Const Lines = vbCrLf & vbCrLf & "Friend Sub Tst()" & vbCrLf & "End Sub"
Md.InsertLines Md.CountOfLines + 1, Lines
Debug.Print "Inserted: "; Md.Parent.Name
End Sub
Private Sub ZMdAy_NoFriendSubTst__Tst()
Dim A() As CodeModule
    A = ZMdAy_NoFriendSubTst
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

Private Property Get ZMdAy_NoFriendSubTst() As CodeModule()
Dim O() As CodeModule
    Dim I, Cmp As VBComponent, Md As CodeModule
    For Each I In ZCurPj.VBComponents
        Set Cmp = I
        If Cmp.Type = vbext_ct_ClassModule Then
            If Cmp.Name <> "Tst" Then
                Set Md = Cmp.CodeModule
                If ZIsNoFriendSubTst(Md) Then ZPush O, Md
            End If
        End If
    Next
ZMdAy_NoFriendSubTst = O
End Property
Private Property Get ZIsNoFriendSubTst(Md As CodeModule) As Boolean
Dim L$
    L = Md.Lines(1, Md.CountOfLines)
Dim P&
    P = InStr(L, "Friend Sub Tst()" & vbCrLf)
ZIsNoFriendSubTst = P = 0
End Property
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
Private Function ZSz&(Ay)
On Error Resume Next
ZSz = UBound(Ay) + 1
End Function
Private Property Get ZCurPj() As VBProject
Set ZCurPj = Application.VBE.ActiveVBProject
End Property

Private Sub ZPass(MethNm$)
Debug.Print "Pass: " & MethNm
End Sub



Sub Tst()
End Sub
