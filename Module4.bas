Attribute VB_Name = "Module4"
Option Explicit

Private Sub KeyList()

Dim routeSplit() As String, varE() As String, cellTot() As String
Dim minRoute As Integer, i As Integer, l As Integer
Dim rng As Range, rng2 As Range
Dim var1 As String, var2 As String, var3 As String
Dim varEInput As Integer

Call Module3.RouteNos

var3 = "A1"
var1 = "A3"

For i = 0 To UBound(noRoutes) - 1

    routeSplit = Split(noRoutes(i), "-")
    
    minRoute = Min(routeSplit)
    
    Debug.Print (minRoute)
    
    Worksheets(5).Select
    
    Range("A3").CurrentRegion.Select
    
    Set rng = Selection.Find("R" & minRoute, After:=Range(var1), LookIn:=xlValues, LookAt:=xlWhole, _
                            SearchOrder:=xlRows)
                            
    Set rng2 = Selection.Find("R" & minRoute & " Total", After:=Range(var1), LookIn:=xlValues, LookAt:=xlWhole, _
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
        
        varE = Module3.Trial_Func("1")
        
    Else
    
        Debug.Print var2
        cellTot = Split(var2, "$")
        Range(var1 & ":" & "E" & cellTot(2) - 1).Select
        varEInput = Selection.SpecialCells(xlCellTypeConstants).Count
        
        Range(var1).Select

        
        varE = Module3.Trial_Func(CStr(varEInput))
        
    End If
    
    'Debug.Print UBound(varE), varE(0)

    Worksheets(4).Select
    Range(var3).Select
    
    Selection.CurrentRegion.Select
    
    'Debug.Print "Route " & "0" & noRoutes(i)
        
    If minRoute < 10 Then
        Set rng2 = Selection.Find("0" & minRoute, After:=Range(var3), LookIn:=xlValues, LookAt:=xlPart, _
                        SearchOrder:=xlRows)
                        
    Else
            
        Set rng2 = Selection.Find(minRoute, After:=Range(var3), LookIn:=xlValues, LookAt:=xlPart, _
                        SearchOrder:=xlRows)
                        
    End If
                        
                        
                    
    Debug.Print rng2.Address
    
    rng2.Select
    
    var3 = Selection.End(xlDown).Address
    
    Selection.Offset(1, 0).Select
    
    For l = 0 To UBound(varE) - 1
        
        Selection.Offset(l, 0).Value = varE(l)
        
        Next l
        
    Selection.Offset(l, 0).Select
    
    Worksheets(5).Select
    
Next i
    

End Sub

Function Min(arr_string() As String)

Dim i As Integer, minVal As String

minVal = arr_string(0)

If UBound(arr_string) > 0 Then

    For i = 0 To UBound(arr_string)
    
        If CInt(arr_string(i)) < CInt(minVal) Then
        
            minVal = arr_string(i)
            
        End If
        
    Next i
        
End If

Min = minVal

End Function

