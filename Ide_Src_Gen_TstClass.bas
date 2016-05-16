Attribute VB_Name = "Ide_Src_Gen_TstClass"
Option Explicit
Sub Tst()
Src_Gen_TstClass
End Sub
Private Sub ZBodyModule__Tst()
Str_Brw ZBodyModule
End Sub
Private Sub ZBodyClass__Tst()
Str_Brw ZBodyClass
End Sub
Sub Src_Gen_TstClass()
ZCrt ZBodyClass, "TstC"
ZCrt ZBodyModule, "TstM"
End Sub
Private Property Get ZBodyModule$()
Dim A$()
    A = ZNmAy_AllModule
ZBodyModule = _
    ZBodyModule_Sub_All(A) & _
    ZBodyModule_Each(A)
End Property
Private Property Get ZBodyClass$()
Dim A$()
    A = ZNmAy_AllClass
ZBodyClass = _
    ZBodyClass_Sub_All(A) & _
    ZBodyClass_Each(A)
End Property
Private Property Get ZBodyClass_Sub_All$(ClassNmAy$())
Dim N%, O$()
    N = Sz(ClassNmAy)
ReDim O$(N + 1)

Dim J%
O(0) = "Sub All"
For J = 0 To N - 1
    O(J + 1) = ClassNmAy(J)
Next
O(N + 1) = "End Sub"
ZBodyClass_Sub_All = Join(O, vbCrLf) & vbCrLf & vbCrLf
End Property

Private Property Get ZBodyClass_Each$(ClassNmAy$())
Dim A$()
    A = ClassNmAy
Dim J%
Dim B$()
If Sz(A) = 0 Then Exit Property
ReDim B(UBound(A))
For J = 0 To UBound(A)
    B(J) = ZOneMethBody_Class(A(J))
Next
ZBodyClass_Each = Join(B, vbCrLf)
End Property
Private Property Get ZBodyModule_Sub_All$(ModuleNmAy$())
Dim N%, O$()
    N = Sz(ModuleNmAy)
ReDim O$(N + 1)

Dim J%
O(0) = "Sub All"
For J = 0 To N - 1
    O(J + 1) = ModuleNmAy(J)
Next
O(N + 1) = "End Sub"
ZBodyModule_Sub_All = Join(O, vbCrLf) & vbCrLf & vbCrLf
End Property

Private Property Get ZBodyModule_Each$(ModuleNmAy$())
Dim A$()
    A = ModuleNmAy
Dim PjNm$
    PjNm = Application.VBE.ActiveVBProject.Name
Dim J%
Dim B$()
ReDim B(UBound(A))
For J = 0 To UBound(A)
    B(J) = ZOneMethBody_Module(A(J), PjNm$)
Next
ZBodyModule_Each = Join(B, vbCrLf)
End Property
Private Sub ZNmAy_AllClass__Tst()
Ay_Brw ZNmAy_AllClass
End Sub

Private Property Get ZNmAy_AllClass() As String()
Dim O$()
    Dim I, Cmp As VBComponent
    For Each I In ZCurPj.VBComponents
        Set Cmp = I
        If Cmp.Type = vbext_ct_ClassModule Then
            If Cmp.Name <> "TstC" And Cmp.Name <> "TstM" Then
                Push O, Cmp.Name
            End If
        End If
    Next
ZNmAy_AllClass = O
End Property

Private Sub ZNmAy_AllModule__Tst()
Ay_Brw ZNmAy_AllModule
End Sub

Private Property Get ZNmAy_AllModule() As String()
Dim O$()
    Dim I, Cmp As VBComponent
    For Each I In ZCurPj.VBComponents
        Set Cmp = I
        If Cmp.Type = vbext_ct_StdModule Then
            If Cmp.Name <> "TstM" And Cmp.Name <> "TstC" Then
                Push O, Cmp.Name
            End If
        End If
    Next
ZNmAy_AllModule = O
End Property

Private Property Get ZOneMethBody_Class$(ClassNm$)
Dim O$(4)
O(0) = "Sub " & ClassNm
O(1) = "Dim M As New " & ClassNm
O(2) = "M.Tst"
O(3) = "End Sub"
ZOneMethBody_Class = Join(O, vbCrLf)
End Property

Private Property Get ZOneMethBody_Module$(ModuleNm$, PjNm$)
Dim O$(3)
O(0) = "Sub " & ModuleNm$
O(1) = PjNm & "." & ModuleNm & ".Tst"
O(2) = "End Sub"
ZOneMethBody_Module = Join(O, vbCrLf)
End Property

Private Sub ZCrt(Body$, ClassNm$)
ZDltClass ClassNm
Dim F$
    F = ZCrtTmpTstCls(Body, ClassNm)

Dim A As VBComponent
    Set A = ZCurPj.VBComponents.Add(vbext_ct_ClassModule)
    A.Name = ClassNm
    A.CodeModule.DeleteLines 1, 2
    A.CodeModule.AddFromFile F

Kill F
End Sub

Private Function ZCrtTmpTstCls$(ClassBody$, ClassNm$)
Const O$ = "C:\Temp\Tst.cls"
If Dir(O) <> "" Then Kill O
Dim T%
    T = FreeFile(1)
Open O For Output As T
'Print #T, "VERSION 1.0 CLASS"
'Print #T, "BEGIN"
'Print #T, "  MultiUse = -1  'True"
'Print #T, "End"
Print #T, "Attribute VB_Name = """ & ClassNm & """"
Print #T, "Attribute VB_GlobalNameSpace = False"
Print #T, "Attribute VB_Creatable = False"
Print #T, "Attribute VB_PredeclaredId = False"
Print #T, "Attribute VB_Exposed = False"
Print #T, "Option Explicit"
Print #T, ClassBody
Close #T
ZCrtTmpTstCls = O
End Function
Private Property Get ZCurPj() As VBProject
Set ZCurPj = Application.VBE.ActiveVBProject
End Property

Private Sub ZDltClass(ClassNm$)
Dim A As VBComponent
    Set A = ZClass(ClassNm)
    If TypeName(A) = "Nothing" Then Exit Sub
ZCurPj.VBComponents.Remove A
End Sub
Private Sub ZClass__Tst()
Dim A As VBComponent
    Set A = ZClass("TstM")
Debug.Assert TypeName(A) <> "Nothing"
Debug.Assert A.Name = "TstM"
Debug.Assert A.Type = vbext_ct_ClassModule
ZPass "ZTheClass"
End Sub
Private Sub ZPass(MethNm$)
Debug.Print "Pass: " & MethNm
End Sub
Private Property Get ZClass(ClassNm$) As VBComponent
Dim I, Cmp As VBComponent
For Each I In ZCurPj.VBComponents
    Set Cmp = I
    If Cmp.Name = ClassNm Then
        If Cmp.Type <> vbext_ct_ClassModule Then Stop
        Set ZClass = Cmp: Exit Property
    End If
Next
End Property
