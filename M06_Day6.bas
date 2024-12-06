Attribute VB_Name = "M06_Day6"
Option Explicit

Function IsInBound(ByRef vTabIn As Variant, i, j) As Boolean
    If i < LBound(vTabIn, 1) Or i > UBound(vTabIn, 1) Then
        IsInBound = False
        Exit Function
    End If
    
    If j < 1 And j > Len(vTabIn(i, 1)) Then
        IsInBound = False
        Exit Function
    End If
    
    IsInBound = True
End Function

Sub D6_Part1()

    Dim vTabIn As Variant
    
    Dim bFound As Boolean
    
    Dim count As Long
    Dim lastRow As Long
    
    Dim i As Long
    Dim j As Long
    Dim k As Long
    Dim l As Long
    
    With F_D6
        lastRow = .Cells(.Rows.count, 1).End(xlUp).Row
        vTabIn = .Range(.Cells(1, 1), .Cells(lastRow, 1)).value
    End With
    
    For i = LBound(vTabIn, 1) To UBound(vTabIn, 1)
        For j = 1 To Len(vTabIn(i, 1))
            If Mid(vTabIn(i, 1), j, 1) = "^" Then
                bFound = True
                Exit For
            End If
        Next j
        
        If bFound Then Exit For
    Next i
    
    k = -1
    l = 0
    count = 0
    Do While IsInBound(vTabIn, i + k, j + l)
        If Mid(vTabIn(i + k, 1), j + l, 1) = "#" Then
            If k <> 0 Then
                l = IIf(k = -1, 1, -1)
                k = 0
            Else
                k = IIf(l = -1, -1, 1)
                l = 0
            End If
        Else
            If Mid(vTabIn(i + k, 1), j + l, 1) <> "X" Then
                count = count + 1
                Mid(vTabIn(i + k, 1), j + l, 1) = "X"
            End If
            
            i = i + k
            j = j + l
        End If
    Loop

    Debug.Print count
    
End Sub
