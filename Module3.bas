Attribute VB_Name = "Module3"
Option Explicit
Dim tab_rearList As ListObject
Public noRoutes() As String
Dim stopNo() As String
'It is very important to understand that the settings LookIn, LookAt and SearchOrder are saved each time the Find Method is used. For this reason one
'should always specify these settings explicitly each and every time you use the Find Method. If you don't, you run the risk of using the Find Method with settings you were not aware of.

Public Sub StopNames()

Call RouteNos
Call StopsSplit

Dim i As Integer, j As Integer, k As Integer, l As Integer, vCount As Integer
Dim rng As Range, rng2 As Range
Dim var1 As String, var2 As String, varEInput As String
Dim routesSplit() As String, stopSplit() As String, varE() As String


Worksheets(4).Select
Range("A1").Offset(0, 1).Select


Worksheets(5).Select
Range("A3").CurrentRegion.Select

'Debug.Print noRoutes(0)
'Debug.Print stopNo(0)

var1 = Range("A3").Address
var2 = Range("A1").Address

For i = 0 To UBound(noRoutes) - 1

    Range("A3").CurrentRegion.Select
    
    If noRoutes(i) = "" Then
    
        GoTo NextIteration
    
    ElseIf InStr(noRoutes(i), "-") Then 'Stops with "-"
    
        routesSplit = Split(noRoutes(i), "-")
        stopSplit = Split(stopNo(i), "-")
        
        
        For j = 0 To UBound(stopSplit)
                  
            Worksheets(5).Select
             
            Range("A3").CurrentRegion.Select
            
            Set rng = Selection.Find("R" & routesSplit(j), After:=Range(var1), LookIn:=xlValues, LookAt:=xlWhole, _
                        SearchOrder:=xlRows)
                        
            Debug.Print routesSplit(j)
            
            Worksheets(4).Select
            
            If i = 0 And j = 0 Then
                    
                    Selection.End(xlDown).Offset(1, 0).Select
                    
                ElseIf j = 0 Then
                
                     Selection.End(xlDown).End(xlDown).Offset(1, 0).Select
                     
            End If
            
            If rng Is Nothing Then
                        
                GoTo NextIt
            
            Else
    
                varEInput = Module5.KeyListOnly(routesSplit(j))
                
                'Function
                varE = Trial_Func(varEInput)
                
                'New code to integrate into ticket
                
                Worksheets(4).Select
                Debug.Print i, j
 
                For l = 0 To UBound(varE) - 1
                    
                    Selection.Offset(l, 0).Value = "R" & routesSplit(j) & "- " & varE(l)
                    
                    Next l
                    
                Selection.Offset(l, 0).Select
                
                Worksheets(5).Select
                
            End If
            
NextIt:
            
        
        Next j
        
       'New Code
        
        Worksheets(4).Select
        
        var2 = Selection.End(xlDown).End(xlDown).Address
        
        Worksheets(5).Select
        
    
    Else 'Stops with no "-"
        
        Set rng = Selection.Find("R" & noRoutes(i), After:=Range(var1), LookIn:=xlValues, LookAt:=xlWhole, _
        SearchOrder:=xlRows)

        Debug.Print rng
        
        var1 = rng.Address
        
        Worksheets(5).Select
        rng.Select
        Selection.Offset(0, 3).Select
        
        Debug.Print stopNo(i)
        
        varE = Trial_Func(stopNo(i))

    
        'New Code
        Worksheets(4).Select
        Range(var2).Select
        
        Selection.CurrentRegion.Select
        
        'Debug.Print "Route " & "0" & noRoutes(i)
        
        If CInt(noRoutes(i)) < 10 Then
            Set rng2 = Selection.Find("Route " & "0" & noRoutes(i), After:=Range(var2), LookIn:=xlValues, LookAt:=xlWhole, _
                        SearchOrder:=xlRows)
                        
        Else
        
            Set rng2 = Selection.Find("Route " & noRoutes(i), After:=Range(var2), LookIn:=xlValues, LookAt:=xlWhole, _
                        SearchOrder:=xlRows)
                        
        End If
                    
        Debug.Print rng2.Address
        
        Range(rng2.Address).Select
        
        var2 = Selection.End(xlDown).Address
        
        Selection.Offset(1, 0).Select
        
        For l = 0 To UBound(varE) - 1
            
            Selection.Offset(l, 0).Value = "R" & noRoutes(i) & "- " & varE(l)
            
            Next l
            
        Selection.Offset(l, 0).Select
    
        Worksheets(5).Select
        
    End If

NextIteration:
    
Next i
    
Application.CutCopyMode = False

Worksheets(4).Select

End Sub

Function Trial_Func(str_arr As String) As String()

Dim i As Integer, k As Integer, vCount As Integer
Dim varE() As String

vCount = 0

If CInt(str_arr) > 1 Then 'More than 1 stop

    ReDim varE(CInt(str_arr)) As String
    
    For k = 0 To CInt(str_arr) - 1
         
         If Not IsEmpty(Selection.Offset(k + vCount, 0)) Then
             
             varE(k) = Selection.Offset(k + vCount, 0).Value
             
         Else
         
             
             Do While IsEmpty(Selection.Offset(k + vCount, 0))
                 
                 vCount = vCount + 1
                 
             Loop
             
    
             varE(k) = Selection.Offset(k + vCount, 0).Value
             
                   
         End If
         
     Next k
     
     Trial_Func = varE
        
Else
    
    ReDim varE(1)
    varE(0) = Selection.Value
    Trial_Func = varE
    
End If


End Function

Function SplitForSlash(col_or_row As ListObject, col_num As Integer) As String()
    
    Dim n As Integer, x As Integer, i As Integer, sumStops As Integer
    Dim noRoutes() As String, routeNo() As String

    n = Int(col_or_row.ListRows.Count)
    
    ReDim routeNo(n) As String

    For x = 1 To col_or_row.ListRows.Count
        
        If InStr(col_or_row.DataBodyRange(x, col_num), "/") Then
            
            noRoutes() = Split(col_or_row.DataBodyRange(x, col_num), "/")
            
            ' Debug.Print noRoutes(0), noRoutes(1)
             
             routeNo(x - 1) = noRoutes(0)
            
        Else
        
            routeNo(x - 1) = col_or_row.DataBodyRange(x, col_num)
        
        End If
        
    Next x

    SplitForSlash = routeNo
    
End Function

Sub RouteNos()

Dim i As Integer, j As Integer
Dim routeNo_nozero() As String

Set tab_rearList = Worksheets(3).ListObjects("RearLoaderList")

noRoutes = SplitForSlash(tab_rearList, 5)

'Debug.Print UBound(noRoutes), noRoutes(20)

For i = 0 To UBound(noRoutes) - 1
    
    If InStr(noRoutes(i), "-") Then
        
        routeNo_nozero = Split(noRoutes(i), "-")
        
        'Debug.Print UBound(routeNo_nozero)
        
        For j = 0 To UBound(routeNo_nozero)
            
            If Left(routeNo_nozero(j), 1) = "0" Then
            
                routeNo_nozero(j) = Right(routeNo_nozero(j), 1)
                
            End If
            
        Next j
           
        
        noRoutes(i) = Join(routeNo_nozero, "-")
                

    ElseIf Not InStr(noRoutes(i), "-") And Left(noRoutes(i), 1) = "0" Then
        
        noRoutes(i) = Right(noRoutes(i), 1)
        
        
    End If
    
Next i

'Debug.Print noRoutes(2) ', noRoutes(8)

End Sub

Sub StopsSplit()

Set tab_rearList = Worksheets(3).ListObjects("RearLoaderList")

stopNo = SplitForSlash(tab_rearList, 7)

Debug.Print stopNo(2)

End Sub





