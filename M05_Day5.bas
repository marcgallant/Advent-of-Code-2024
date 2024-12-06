Attribute VB_Name = "M05_Day5"
Option Explicit

Const noField = -1

Function Exists(vTabIn As Variant, elt As Variant) As Long
    
    Dim i As Long
    
    If Not IsArray(vTabIn) Then
        Exists = noField
        Exit Function
    End If
    
    If Not IsEmpty(vTabIn) Then
        For i = LBound(vTabIn, 1) To UBound(vTabIn, 1)
            If vTabIn(i) = elt Then
                Exists = i
                Exit Function
            End If
        Next i
    End If
    
    Exists = noField
    
End Function

Sub D5_Part1()
    
    Dim vTabRules As Variant
    Dim vTabUpdates As Variant
    Dim vTabTemp As Variant
    
    Dim Dico As Object
    Dim count As Long
    Dim lastRow As Long
    
    Dim bValid As Boolean
    Dim tmp As Collection
    
    Dim left As String
    Dim right As String
    
    Dim i As Long
    Dim j As Long
    Dim k As Long
    
    Dim pos As Long
    
    With F_D5
        lastRow = .Cells(.Rows.count, 1).End(xlUp).Row
        vTabRules = .Range(.Cells(1, 1), .Cells(lastRow, 1)).value
        
        lastRow = .Cells(.Rows.count, 2).End(xlUp).Row
        vTabUpdates = .Range(.Cells(1, 2), .Cells(lastRow, 2)).value
    End With
    
    Set Dico = CreateObject("Scripting.Dictionary")
    
    For i = LBound(vTabRules, 1) To UBound(vTabRules, 1)
        vTabTemp = Split(vTabRules(i, 1), "|")
        left = vTabTemp(0)
        right = vTabTemp(1)
        
        If Not Dico.Exists(left) Then
            Set tmp = New Collection
            Dico.Add left, tmp
        End If
        
        Call Dico.item(left).Add(right)
    Next i
    
    count = 0
    
    For i = LBound(vTabUpdates, 1) To UBound(vTabUpdates, 1)
        
        bValid = True
        vTabTemp = Split(vTabUpdates(i, 1), ",")
        
        For j = LBound(vTabTemp, 1) To UBound(vTabTemp, 1)
            
            If Dico.Exists(vTabTemp(j)) Then
                Set tmp = Dico.item(vTabTemp(j))
                For k = 1 To tmp.count
                    pos = Exists(vTabTemp, tmp(k))
                    If pos <> -1 And pos < j Then
                        bValid = False
                        Exit For
                    End If
                Next k
            
            End If
            
            If Not bValid Then Exit For
        
        Next j
        
        If bValid Then count = count + vTabTemp(UBound(vTabTemp, 1) / 2)
    
    Next i
    
    Debug.Print count
    
End Sub

Sub D5_Part2()

End Sub
