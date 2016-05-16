Attribute VB_Name = "Ide_Md"
Option Explicit

Function Md_Nm$(Md As CodeModule)
Md_Nm = Md.Parent.Name
End Function

Function Md_Pj(Md As CodeModule) As VBProject
Set Md_Pj = Md.Parent.Collection.Parent
End Function

Sub Md_CpyToPj(Md As CodeModule, ToPj As VBProject)

End Sub
Sub Md_Export(Md As CodeModule)
Md.Parent.Export Md_SrcFfn(Md)
End Sub

Function Md_SrcPth$(Md As CodeModule)
Md_SrcPth = Pj_SrcPth(Md_Pj(Md))
End Function

Function Md_SrcExt$(Md As CodeModule)
Dim Cmp As VBComponent
Set Cmp = Md.Parent
Select Case Cmp.Type
Case vbext_ct_StdModule: Md_SrcExt = ".bas"
Case vbext_ct_ClassModule: Md_SrcExt = ".cls"
Case Else: Stop
End Select
End Function

Function Md_SrcFn$(Md As CodeModule)
Md_SrcFn = Md.Parent.Name
End Function

Function Md_SrcFfn$(Md As CodeModule)
Md_SrcFfn = Md_SrcPth(Md) & Md_SrcFn(Md) & Md_SrcExt(Md)
End Function

Sub Tst()

End Sub

