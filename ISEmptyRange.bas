Attribute VB_Name = "Module1"
Sub IsEmptyRange()
    Dim cell As Range
    Dim bIsEmpty As Boolean
    Dim Sum As Integer
    Dim Max As Integer
    bIsEmpty = False
    Max = 1
    Sum = 0
    For Each cell In Range("A1:AYS1")
        If IsEmpty(cell) = False Then
            If Worksheets("Sheet0").Cells(1, Max).Value Like "Functions" Then
                MsgBox "Found" & Max
                If IsEmpty(Worksheets("Sheet0").Cells(10, Max)) = False Then
                    Sum = Sum + 1
                End If
            End If
            Max = Max + 1
        End If
    Next cell
    Worksheets("Sheet0").Cells(10, 1).Value = Sum
End Sub

