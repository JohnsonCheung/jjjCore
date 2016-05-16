Attribute VB_Name = "Ide_Src_Commit_Pj"
Option Explicit
Private Sub Src_Commit_Pj__Tst()
Src_Commit_Pj
End Sub
Sub Src_Commit_Pj(Optional VBar_Msg$ = "Commit")
Dim A$
    A = Replace(VBar_Msg, "|", vbCrLf)
Src_Export_Pj
ZCpyPj
Dim F$
    If ZIs_DotGitDir_Missing Then
        F = ZBld_BatFil_Init_and_Commit(A)
    Else
        F = ZBld_BatFil_Commit(A)
    End If
Shell F, vbMaximizedFocus
Kill F
End Sub

Private Sub ZCpyPj__Tst()
ZCpyPj
End Sub

Private Sub ZCpyPj()
Dim Fm$, FfnTo$
    Dim P As VBProject
    Set P = Cur_Pj
    Fm = P.Filename
    FfnTo = Pj_SrcPth(P) & Pj_Fn(P)
Fso.CopyFile Fm, FfnTo, True
End Sub
Private Sub ZIs_DotGItDir_Missing__Tst()
Debug.Print ZIs_DotGitDir_Missing
End Sub

Private Property Get ZIs_DotGitDir_Missing() As Boolean
Dim A$
    A = Dir(ZDotGitDir, vbDirectory + vbHidden)
ZIs_DotGitDir_Missing = A = ""
End Property


Private Sub ZDotGitDir__Tst()
Debug.Print ZDotGitDir$
End Sub

Private Property Get ZDotGitDir$()
ZDotGitDir$ = Pj_SrcPth(Cur_Pj) & ".git\"
End Property

Private Function ZBld_BatFil_Commit$(Optional Msg$ = "Commit")
Const F$ = "C:\Temp\Git_Commit.bat"
Dim T%
T = FreeFile(1)
Open F For Output As T
Print #T, ZBatFilBody_Commit(Msg)
Close #T
ZBld_BatFil_Commit = F
End Function
Private Sub ZBld_BatFil_Commit__Tst()
Debug.Assert ZBld_BatFil_Commit() = "C:\Temp\Git_Commit.bat"
End Sub

Private Sub ZBld_BatFil_Init_and_Commit__Tst()
Debug.Assert ZBld_BatFil_Init_and_Commit$ = "C:\Temp\Git_Init_and_Commit.bat"
End Sub
Private Function ZBld_BatFil_Init_and_Commit$(Optional Msg$ = "Commit")
Const F$ = "C:\Temp\Git_Init_and_Commit.bat"
Dim T%
T = FreeFile(1)
Open F For Output As T
Print #T, ZBatFilBody_Init_and_Commit(Msg)
Close #T
ZBld_BatFil_Init_and_Commit = F
End Function

Private Property Get ZBatFilBody$(Optional Msg$ = "Commit", Optional NeedInit As Boolean = False)
Dim P As VBProject
Dim PjFn$
Dim A$
    Set P = Cur_Pj
    A = Pj_SrcPth(P)
    PjFn = Pj_Fn(P)
Dim O$(5)
Dim Drv$
Dim Pth$
    Drv = Left(A, 2)
    Dim B$()
    B = Split(A, "\")
    B(UBound(B) - 1) = PjFn
    Pth = Join(B, "\")
O(0) = Drv
O(1) = "CD " & Pth
O(2) = "Git init": If Not NeedInit Then O(2) = ""
O(3) = "git add *.*"
O(4) = "git commit -a -m """ & Msg & """"
O(5) = "rem Pause"
ZBatFilBody = Join(O, vbCrLf)
End Property

Private Sub ZBatFilBody_Init_and_Commit__Tst()
Debug.Print ZBatFilBody_Init_and_Commit
End Sub

Private Sub ZBatFilBody_Commit__Tst()
Debug.Print ZBatFilBody_Commit
End Sub

Private Property Get ZBatFilBody_Init_and_Commit$(Optional Msg$ = "Commit")
ZBatFilBody_Init_and_Commit = ZBatFilBody(Msg, NeedInit:=True)
End Property

Private Property Get ZBatFilBody_Commit$(Optional Msg$ = "Commit")
ZBatFilBody_Commit = ZBatFilBody(Msg)
End Property


Sub Tst()
End Sub
