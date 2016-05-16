Attribute VB_Name = "Ide_Pj"
Option Explicit

Sub Pj_RenMd(Pj As VBProject, Pfx$, ToPfx$)
Dim Ay$(), J%, Nm$, NewNm$
Ay = Ay_SubSet_ByPfx(Pj_MdNmAy(Pj), Pfx)
For J = 0 To UB(Ay)
    Nm = Ay(J)
    NewNm = Str_ReplPfx(Nm, Pfx, ToPfx)
    If Pj_IsCmp(Pj, NewNm) Then
        Debug.Print NewNm; "<== Exist"
    Else
        Pj.VBComponents(Ay(J)).Name = NewNm
    End If
Next
End Sub

Function Pj_MdAy(Pj As VBProject) As CodeModule()
Dim O() As CodeModule
Dim Cmp As VBComponent
For Each Cmp In Pj.VBComponents
    If Cmp.Type = vbext_ct_ClassModule Or Cmp.Type = vbext_ct_StdModule Then
        Push O, Cmp.CodeModule
    End If
Next
Pj_MdAy = O
End Function

Function Pj_Fn$(Pj As VBProject)
Pj_Fn = Ffn_Fn(Pj.Filename)
End Function

Function Pj_SrcPth$(Pj As VBProject)
Dim P$
Dim F$
Dim O$
    F = Pj_Fn(Pj)
    P = Ffn_Pth(Pj.Filename)
    O = P & "Src\" & F & "\"
Pj_SrcPth = O
End Function

Sub Pj_Export(Pj As VBProject)
Pj_Sav Pj

Dim Ay() As CodeModule
Dim P$
    Ay = Pj_MdAy(Pj)
    P = Pj_SrcPth(Pj)
If Sz(Ay) = 0 Then Exit Sub
'----
Dim Md As CodeModule, I
Pth_CrtEachSeg P
Pth_DltFil P
For Each I In Ay
    Set Md = I
    Md_Export Md
Next
End Sub
Function Pj_Pth$(Pj As VBProject)
Pj_Pth = Ffn_Pth(Pj.Filename)
End Function

Sub Pj_Sav(Pj As VBProject)
If Pj.Saved Then Exit Sub
Stop
Wb_Sav Pj_Wb(Pj)
End Sub

Function Pj_Wb(Pj As VBProject) As Workbook
Dim F$
    F = Pj.Filename
Dim Wb As Workbook
For Each Wb In Application.Workbooks
    If Wb.FullName = F Then Set Pj_Wb = Wb: Exit Function
Next
End Function

Function Pj_FilNm$(Pj As VBProject)
On Error Resume Next
Pj_FilNm = Pj.Filename
End Function

Function Pj_IsCmp(Pj As VBProject, CmpNm$) As Boolean
Dim C As VBComponent
For Each C In Pj.VBComponents
    If C.Name = CmpNm Then Pj_IsCmp = True: Exit Function
Next
End Function

Function Pj_ByNm(Nm$) As VBProject
Dim Pj As VBProject
For Each Pj In Application.VBE.VBProjects
    If Pj.Name = Nm Then Set Pj_ByNm = Pj
Next
End Function

Function Pj_MdNmAy(Pj As VBProject) As String()
Dim O$()
Dim C As VBComponent
For Each C In Pj.VBComponents
    If C.Type = vbext_ct_StdModule Then
        Push O, C.Name
    End If
Next
Pj_MdNmAy = O
End Function


Sub Tst()
End Sub
