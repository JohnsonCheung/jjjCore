Attribute VB_Name = "Ide_Cur"
Option Explicit

Property Get Cur_PjNm$()
Cur_PjNm = Cur_Pj.Name
End Property
Property Get Cur_Pj() As VBProject
Set Cur_Pj = Application.VBE.ActiveVBProject
End Property

Property Get Cur_PjSrcPth$()
Cur_PjSrcPth = Pj_SrcPth(Cur_Pj)
End Property

Property Get Cur_PjPth$()
Cur_PjPth = Pj_Pth(Cur_Pj)
End Property

Sub Tst()

End Sub
