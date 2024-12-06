Attribute VB_Name = "M04_Day4"
Option Explicit

Sub D4_Part1()

    Dim vTabIn As Variant
    
    Dim count As Long
    Dim lastRow As Long
    
    Dim bFind As Boolean
    Const toFind As String = "XMAS"
    
    Dim i As Long
    Dim j As Long
    Dim k As Long
    Dim l As Long
    Dim m As Long
    
    With F_D4
        lastRow = .Cells(.Rows.count, 1).End(xlUp).Row
        vTabIn = .Range(.Cells(1, 1), .Cells(lastRow, 1)).value
    End With
    
    count = 0
    
    For i = LBound(vTabIn, 1) To UBound(vTabIn, 1)
        For j = 1 To Len(vTabIn(i, 1))
        
            For k = -1 To 1
            
                If i + k * (Len(toFind) - 1) <= UBound(vTabIn, 1) And i + k * (Len(toFind) - 1) >= LBound(vTabIn, 1) Then
                For l = -1 To 1
                    
                    If (k <> 0 Or l <> 0) _
                    And (j + l * (Len(toFind) - 1) <= Len(vTabIn(i + k * (Len(toFind) - 1), 1)) And j + l * (Len(toFind) - 1) >= 1) Then
                        bFind = True
                        
                        For m = 1 To Len(toFind)
                            
                            If Mid(toFind, m, 1) <> Mid(vTabIn(i + k * (m - 1), 1), j + l * (m - 1), 1) Then
                                bFind = False
                                Exit For
                            End If
                            
                        Next m
        
                        If bFind Then
                            count = count + 1
                        End If
                    End If
                Next l
                End If
            Next k
        
        Next j
    Next i
    
    Debug.Print count
End Sub

Sub D4_Part2()

    Dim vTabIn As Variant
    
    Dim count As Long
    Dim lastRow As Long
    
    Dim i As Long
    Dim j As Long
    Dim k As Long
    Dim l As Long
    
    With F_D4
        lastRow = .Cells(.Rows.count, 1).End(xlUp).Row
        vTabIn = .Range(.Cells(1, 1), .Cells(lastRow, 1)).value
    End With
    
    count = 0
    
    For i = LBound(vTabIn, 1) + 1 To UBound(vTabIn, 1) - 1
        For j = 2 To Len(vTabIn(i, 1)) - 1
        
            If (Mid(vTabIn(i, 1), j, 1) = "A") _
            And (Mid(vTabIn(i - 1, 1), j - 1, 1) = "M" And Mid(vTabIn(i + 1, 1), j + 1, 1) = "S" _
            Or Mid(vTabIn(i - 1, 1), j - 1, 1) = "S" And Mid(vTabIn(i + 1, 1), j + 1, 1) = "M") _
            And (Mid(vTabIn(i - 1, 1), j + 1, 1) = "M" And Mid(vTabIn(i + 1, 1), j - 1, 1) = "S" _
            Or Mid(vTabIn(i - 1, 1), j + 1, 1) = "S" And Mid(vTabIn(i + 1, 1), j - 1, 1) = "M") Then
                count = count + 1
            End If
            
        Next j
    Next i
    
    Debug.Print count

End Sub
