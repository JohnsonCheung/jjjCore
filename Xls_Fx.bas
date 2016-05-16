Attribute VB_Name = "Xls_Fx"
Option Explicit

Sub Fx_Assert_Opn(Fx$)
Dim Wb As Workbook
For Each Wb In Application.Workbooks
    If Wb.FullName = Fx Then Err.Raise 1, "Bld_FxApp", "Bld_FxApp: Fx is openned, please close." & vbLf & "Fx=[" & Fx & "]"
Next
End Sub

Function Fx_FilFmt(Fx$) As XlFileFormat
Dim Ext$
Ext = Ffn_Ext(Fx)
Select Case LCase(Ext)
Case ".xlsm": Fx_FilFmt = xlOpenXMLWorkbookMacroEnabled
Case ".xlam": Fx_FilFmt = xlOpenXMLAddIn
Case Else: Stop
End Select
End Function


Sub Tst()
End Sub
