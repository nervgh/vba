Attribute VB_Name = "mod_Test"
Private Type Rect
    Left As Double
    Top As Double
    Right As Double
    Bottom As Double
End Type


' ----------------------------------------
' How to use
' ----------------------------------------
Sub Example()
    Dim Rect As Rect
    
    Rect = GetRectForExcel2007Plus(ActiveCell)
    
    With UserForm1
        .StartUpPosition = 0
        .Left = Rect.Right
        .Top = Rect.Top
        .Show False
    End With
End Sub
' ----------------------------------------
