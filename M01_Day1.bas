Attribute VB_Name = "M01_Day1"
Option Explicit

Sub D1_Part1()
    Dim vTabIn As Variant
    Dim vTemp As Variant
    
    Dim left As Collection
    Dim right As Collection
    
    Dim distance As Long
    Dim lastRow As Long
    
    Dim i As Long
    
    With F_D1
        lastRow = .Cells(.Rows.count, 1).End(xlUp).Row
        vTabIn = .Range(.Cells(1, 1), .Cells(lastRow, 1))
    End With
    
    Set left = New Collection
    Set right = New Collection
    
    For i = LBound(vTabIn, 1) To UBound(vTabIn, 1)
    
        vTemp = Split(vTabIn(i, 1), "   ")
        left.Add vTemp(0)
        right.Add vTemp(1)
        
    Next i
    
    Call QuickSort(left, 1, left.count)
    Call QuickSort(right, 1, right.count)
    
    distance = 0
    
    For i = 1 To left.count
        
        distance = distance + Abs(left(i) - right(i))
        
    Next i
    
    Debug.Print distance
    
End Sub

Sub D1_Part2()
    Dim vTabIn As Variant
    Dim vTemp As Variant

    Dim Dico As Object
    Dim left As Variant
    Dim right As Variant
    
    Dim similarity As Long
    Dim lastRow As Long
    
    Dim i As Long
    
    With F_D1
        lastRow = .Cells(.Rows.count, 1).End(xlUp).Row
        vTabIn = .Range(.Cells(1, 1), .Cells(lastRow, 1))
    End With
    
    Set left = New Collection
    Set right = New Collection
    
    For i = LBound(vTabIn, 1) To UBound(vTabIn, 1)
    
        vTemp = Split(vTabIn(i, 1), "   ")
        left.Add vTemp(0)
        right.Add vTemp(1)
        
    Next i
    
    Set Dico = CreateObject("Scripting.Dictionary")
    similarity = 0
    
    For i = 1 To right.count
        
        If Not Dico.Exists(right(i)) Then
            Call Dico.Add(right(i), 1)
        Else
            Dico.item(right(i)) = Dico.item(right(i)) + 1
        End If
    
    Next i
    
    For i = 1 To left.count
        
        If Dico.Exists(left(i)) Then
            similarity = similarity + (left(i) * Dico.item(left(i)))
        End If
        
    Next i
    
    Debug.Print similarity
    
End Sub
