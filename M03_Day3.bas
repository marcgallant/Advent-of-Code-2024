Attribute VB_Name = "M03_Day3"
Option Explicit

Sub D3_Part1()
    
    Dim matches As Object
    Dim match As Object
    
    Dim result As Long
    
    Dim regex As Object
    
    Set regex = CreateObject("VBScript.RegExp")
    regex.pattern = "mul\((\d+),(\d+)\)"
    regex.Global = True
    
    Set matches = regex.Execute(F_D3.Cells(1, 1).value)
    
    For Each match In matches
        result = result + match.submatches(0) * match.submatches(1)
    Next match
    
    Debug.Print (result)
    
End Sub

Sub D3_Part2()
    
    Dim matches As Object
    Dim match As Object
    
    Dim bEnabled As Boolean
    
    Dim result As Long
    
    Dim regex As Object
    
    Set regex = CreateObject("VBScript.RegExp")
    regex.Global = True
    
    regex.pattern = "mul\((\d+),(\d+)\)|do\(\)|don't\(\)"
    Set matches = regex.Execute(F_D3.Cells(1, 1).value)
    
    bEnabled = True
    For Each match In matches
        
        If match.value = "do()" Then
            bEnabled = True
        ElseIf match.value = "don't()" Then
            bEnabled = False
        ElseIf bEnabled Then
            result = result + match.submatches(0) * match.submatches(1)
        End If

    Next match
    
    Debug.Print (result)

End Sub
