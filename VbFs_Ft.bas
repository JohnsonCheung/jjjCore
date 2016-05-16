Attribute VB_Name = "VbFs_Ft"
Option Explicit

Sub Ft_WrtStr(Ft$, S)
Dim F%
F = FreeFile(1)
Open Ft For Output As #F
Print #F, S
Close #F
End Sub

Sub Ft_BrwMinNoFocus(Ft$, Optional AppWinStyle As VbAppWinStyle = VbAppWinStyle.vbMinimizedNoFocus)
Ft_BrwStyle Ft, vbMinimizedNoFocus
End Sub
Sub Ft_Brw(Ft$)
Ft_BrwMaxFocus Ft
End Sub
Sub Ft_BrwMaxFocus(Ft$, Optional AppWinStyle As VbAppWinStyle = VbAppWinStyle.vbMinimizedNoFocus)
Ft_BrwStyle Ft, vbMaximizedFocus
End Sub

Sub Ft_BrwStyle(Ft$, AppWinStyle As VbAppWinStyle)
Shell "NotePad """ & Ft & """", AppWinStyle
End Sub


Sub Tst()
End Sub
