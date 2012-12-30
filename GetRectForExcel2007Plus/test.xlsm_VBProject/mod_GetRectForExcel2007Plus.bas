Attribute VB_Name = "mod_GetRectForExcel2007Plus"
Private Type Rect
    Left As Double
    Top As Double
    Right As Double
    Bottom As Double
End Type


' ----------------------------------------
' Returns the cell coordinates in points relative to the screen
'
' @param {Object} Target the cell
' @return {Rect} the cell coordinates
' ----------------------------------------
Function GetRectForExcel2007Plus(ByVal Target As Range) As Rect
    Dim Index As Integer
    Dim Rect As Rect
    
    With ActiveWindow
    
        Set Target = Target.MergeArea
    
        For Index = 1 To .Panes.Count
            If Not Intersect(Target, .Panes(Index).VisibleRange) Is Nothing Then
                
                With .Panes(Index)
                    Rect.Left = PixelsToPoints(.PointsToScreenPixelsX(Target.Left))
                    Rect.Top = PixelsToPoints(.PointsToScreenPixelsY(Target.Top))
                End With
                
                Rect.Right = Target.Width * .Zoom / 100 + Rect.Left
                Rect.Bottom = Target.Height * .Zoom / 100 + Rect.Top

                GetRectForExcel2007Plus = Rect
                Exit Function
                
            End If
        Next
    End With
End Function


' ----------------------------------------
' Converts pixels to points
' More info http://office.microsoft.com/en-us/excel-help/measurement-units-and-rulers-in-excel-HP001151724.aspx
' Important! 96 is DPI of system and may be different
'
' @param {Double} Pixels
' @return {Double} Points
' ----------------------------------------
Private Function PixelsToPoints(ByVal Pixels As Double) As Double
    PixelsToPoints = Pixels / 96 * 72
End Function


' ----------------------------------------
' Converts points to pixels
' More info http://office.microsoft.com/en-us/excel-help/measurement-units-and-rulers-in-excel-HP001151724.aspx
' Important! 96 is DPI of system and may be different
'
' @param {Double} Points
' @return {Double} Pixels
' ----------------------------------------
Private Function PointsToPixels(ByVal Points As Double) As Double
    PointsToPixels = Points / 72 * 96
End Function

