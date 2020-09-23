Attribute VB_Name = "modForms"
Option Explicit

Function IsLoaded(Form As Form) As Boolean
    'On Error Resume Next
    
    Dim Index As Long
    
    For Index = 0 To Forms.Count - 1
        If Forms(Index) Is Form Then
            IsLoaded = True
            Exit For
        End If
    Next Index
End Function

Function IsVisible(Form As Form) As Boolean
    'On Error Resume Next
    
    If IsLoaded(Form) Then
        IsVisible = Form.Visible
    Else
        IsVisible = False
    End If
End Function
