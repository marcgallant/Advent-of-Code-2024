Attribute VB_Name = "M02_Day2"
Option Explicit

Function IsLineSafe(ByVal coll As Collection) As Boolean
    
    Dim bAscending As Boolean
    
    Dim j As Long
    
    If coll.count > 1 Then
        bAscending = coll(1) < coll(2)
        
        For j = 1 To coll.count - 1
            
            If bAscending Then
                If coll(j) >= coll(j + 1) Or (coll(j + 1) - coll(j)) > 3 Then
                    IsLineSafe = False
                    Exit Function
                End If
            Else
                If coll(j) <= coll(j + 1) Or (coll(j) - coll(j + 1)) > 3 Then
                    IsLineSafe = False
                    Exit Function
                End If
            End If
    
        Next j
    End If
    
    IsLineSafe = True
    
End Function

Sub D2_Part1()
    
    Dim vTabIn As Variant
    Dim vTemp As Variant
    Dim vTabReport As Variant
    
    Dim coll As Collection
    
    Dim count As Long
    Dim lastRow As Long
    Dim lastCol As Long
    
    Dim i As Long
    Dim j As Long
    
    With F_D2
        lastRow = .Cells(.Rows.count, 1).End(xlUp).Row
        vTabIn = .Range(.Cells(1, 1), .Cells(lastRow, 1)).value
    End With
    
    ReDim vTabReport(LBound(vTabIn, 1) To UBound(vTabIn, 1))
    For i = LBound(vTabIn, 1) To UBound(vTabIn, 1)
    
        vTemp = Split(vTabIn(i, 1), " ")
        Set coll = New Collection
        
        For j = LBound(vTemp, 1) To UBound(vTemp, 1)
            coll.Add CInt(vTemp(j))
        Next j
        Set vTabReport(i) = coll
            
    Next i
    
    count = 0
    
    For i = LBound(vTabReport, 1) To UBound(vTabReport, 1)
        
        If IsLineSafe(vTabReport(i)) Then count = count + 1
        
    Next i
    
    Debug.Print count
    
End Sub

Sub D2_Part2()
    
    Dim vTabIn As Variant
    Dim vTemp As Variant
    Dim vTabReport As Variant
    
    Dim coll As Collection
    Dim collTemp As Collection
    
    Dim item As Variant
    
    Dim count As Long
    Dim lastRow As Long
    Dim lastCol As Long
    
    Dim i As Long
    Dim j As Long
    
    With F_D2
        lastRow = .Cells(.Rows.count, 1).End(xlUp).Row
        vTabIn = .Range(.Cells(1, 1), .Cells(lastRow, 1)).value
    End With
    
    ReDim vTabReport(LBound(vTabIn, 1) To UBound(vTabIn, 1))
    For i = LBound(vTabIn, 1) To UBound(vTabIn, 1)
    
        vTemp = Split(vTabIn(i, 1), " ")
        Set coll = New Collection
        For j = LBound(vTemp, 1) To UBound(vTemp, 1)
            
            coll.Add CInt(vTemp(j))
        
        Next j
        Set vTabReport(i) = coll
            
    Next i
    
    count = 0
    
    For i = LBound(vTabReport, 1) To UBound(vTabReport, 1)
        
        For j = 1 To vTabReport(i).count
        
            Set collTemp = New Collection
            For Each item In vTabReport(i)
                collTemp.Add item
            Next item
            
            Call collTemp.Remove(j)
            
            If IsLineSafe(collTemp) Then
                count = count + 1
                Exit For
            End If
            
        Next j
        
    Next i
    
    Debug.Print count
    
End Sub
