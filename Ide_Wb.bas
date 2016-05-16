Attribute VB_Name = "Ide_Wb"
Option Explicit

Sub Wb_Clr_Modules(Wb As VBProject)
Stop
End Sub

Sub Wb_Pj__Tst()
Dim Pj As VBProject
    Set Pj = Wb_Pj(ThisWorkbook)
Debug.Assert Pj.Filename = ThisWorkbook.FullName
End Sub

Function Wb_Pj(Wb As Workbook) As VBProject
Dim Pj As VBProject
For Each Pj In Wb.Application.VBE.VBProjects
    If Pj_FilNm(Pj) = Wb.FullName Then Set Wb_Pj = Pj: Exit Function
Next
End Function


Sub Tst()
End Sub
