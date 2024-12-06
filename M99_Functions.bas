Attribute VB_Name = "M99_Functions"
Option Explicit

Sub QuickSort(coll As Collection, first As Long, last As Long)
  
  Dim vCentreVal As Variant, vTemp As Variant
  
  Dim lTempLow As Long
  Dim lTempHi As Long
  lTempLow = first
  lTempHi = last
  
  vCentreVal = coll((first + last) \ 2)
  Do While lTempLow <= lTempHi
  
    Do While coll(lTempLow) < vCentreVal And lTempLow < last
      lTempLow = lTempLow + 1
    Loop
    
    Do While vCentreVal < coll(lTempHi) And lTempHi > first
      lTempHi = lTempHi - 1
    Loop
    
    If lTempLow <= lTempHi Then
    
      ' Swap values
      vTemp = coll(lTempLow)
      
      coll.Add coll(lTempHi), After:=lTempLow
      coll.Remove lTempLow
      
      coll.Add vTemp, Before:=lTempHi
      coll.Remove lTempHi + 1
      
      ' Move to next positions
      lTempLow = lTempLow + 1
      lTempHi = lTempHi - 1
      
    End If
    
  Loop
  
  If first < lTempHi Then QuickSort coll, first, lTempHi
  If lTempLow < last Then QuickSort coll, lTempLow, last
  
End Sub
