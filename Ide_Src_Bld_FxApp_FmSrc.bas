Attribute VB_Name = "Ide_Src_Bld_FxApp_FmSrc"
Option Explicit
Private Sub ZCrt_Pj__Tst()
ZCrt_Pj "Mass Update (Program) v1.3.xlsm"
End Sub
Sub Src_Bld_FxAppFmSrc()
ZCrt_Pj Ay_AskOne(ZSrcPjNmAy, "Build Project")
End Sub
Private Sub ZSrcPjNmAy__Tst()
Ay_Dmp ZSrcPjNmAy
End Sub
Private Property Get ZSrcPjNmAy() As String()
ZSrcPjNmAy = Pth_DirAy(Cur_PjPth & "Src\")
End Property
Private Sub ZCrt_Pj(PjFn$)
Dim Ay$()
Dim Cmps As VBComponents
Dim Pj As VBProject
    Dim PjSrcPth$
    Dim PjFfn$
    Fx_Assert_Opn PjFfn
    PjFfn = Cur_PjPth & PjFn
    If Dir(PjFfn) <> "" Then Err.Raise 1, , "PjFn exist." & vbCrLf & vbCrLf & "PjFn=[" & PjFn & "]" & vbCrLf & vbCrLf & "Fdr=[" & Ffn_Pth(PjFfn) & "]"
    Set Pj = ZCrt_EmptyPj(PjFfn)
    PjSrcPth = Pj_SrcPth(Pj)
    Ay = Pth_FfnAy(PjSrcPth, "*.bas")
    PushAy Ay, Pth_FfnAy(PjSrcPth, "*.cls")
    Set Cmps = Pj.VBComponents
If Sz(Ay) > 0 Then
    Dim I
    For Each I In Ay
        Cmps.Import CStr(I)
    Next
End If
End Sub

Private Function ZCrt_EmptyPj(PjFfn$) As VBProject
'If PjFfn has .xlam, after Wb.SaveAs PjFfn, the Wb.Name does not changed according to PjFfn!!
'So, it is required to close the Wb and re-open it by Workbooks.Open

Dim Wb As Workbook
Dim FilFmt As XlFileFormat
    FilFmt = Fx_FilFmt(Ffn_Ext(PjFfn))
    Set Wb = Application.Workbooks.Add
    Wb.SaveAs PjFfn, FilFmt
    Wb.Close False
    Set Wb = Application.Workbooks.Open(PjFfn)
    Stop
Dim O As VBProject
    Set O = Wb_Pj(Wb)
    O.Name = Ffn_Fnn(PjFfn)
    
Set ZCrt_EmptyPj = O
End Function


Sub Tst()
End Sub
