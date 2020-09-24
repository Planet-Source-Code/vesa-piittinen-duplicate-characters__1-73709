Attribute VB_Name = "Module1"
Option Explicit

Public Function HasDuplicates(Search As String) As Long
    Dim Length As Long: Length = LenB(Search)
    If Length >= 4 Then
        Dim X As Long
        For X = 1 To Length - 2 Step 2
            If InStrB(X + 1, Search, MidB$(Search, X, 2)) <> 0 Then
                HasDuplicates = (X - 1) \ 2 + 1
                Exit Function
            End If
        Next X
    End If
End Function
