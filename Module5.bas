Attribute VB_Name = "Module5"
Option Explicit

Function KeyListOnly(x As String) As String

Dim cellTot() As String
Dim rng As Range, rng2 As Range
Dim var1 As String, var2 As String
Dim varEInput As Integer

var1 = "A3"

Worksheets(5).Select

Range("A3").CurrentRegion.Select

Set rng = Selection.Find("R" & x, After:=Range(var1), LookIn:=xlValues, LookAt:=xlWhole, _
                        SearchOrder:=xlRows)
                        
Set rng2 = Selection.Find("R" & x & " Total", After:=Range(var1), LookIn:=xlValues, LookAt:=xlWhole, _
SearchOrder:=xlRows)

var1 = rng.Offset(0, 3).Address

Debug.Print var1

If rng2 Is Nothing Then
                        
    Set rng2 = rng
                                             
End If

var2 = rng2.Address

rng.Select
Selection.Offset(0, 3).Select

Debug.Print var1, var2

If rng = rng2 Then
    
    varEInput = "1"
    
Else

    Debug.Print var2
    cellTot = Split(var2, "$")
    Range(var1 & ":" & "E" & cellTot(2) - 1).Select
    varEInput = Selection.SpecialCells(xlCellTypeConstants).Count
    Range(var1).Select
    
    
End If

KeyListOnly = CStr(varEInput)

End Function



