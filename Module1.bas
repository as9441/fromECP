Attribute VB_Name = "Module1"
Option Explicit
Sub FindFirstEmptyCell()
    Dim c
    For Each c In Range("A10:NO10").Cells
        If c = "" Then
            c.Select
            Exit For
        End If
    Next
        
End Sub
