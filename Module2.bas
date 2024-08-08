Attribute VB_Name = "Module2"
Option Explicit
Dim ticketInfo() As String

Public Sub TicketGen()

Dim tab_rearList As ListObject
Dim x As Integer, y As Integer, z As Integer, i As Integer, sumStops As Integer
Dim k As Integer
Dim driver() As String
Dim noStops() As String
Dim routeNo() As String
Dim phRng As Range

Worksheets(3).Select

k = 0

Call StoreVarArray

Worksheets(4).Select
Range("A2").Select

For x = 0 To UBound(ticketInfo) - 1 'We are using indice and the first index is '0' hence why theres a '-1' to account for it
    
    driver() = Split(ticketInfo(x), ",", 5)
    'Debug.Print driver(4)

    For y = 0 To UBound(driver) 'driver array has 5 elements
        If driver(y) = "Null" Or driver(4) = "-" Then
            GoTo NextIteration
            
        Else
            Select Case y
                Case 0 ' Time
                    Selection.Value = driver(y)
                
                Case 1 ' Driver Name
                    ActiveCell.Offset(-1, 0).Select
                    Call BoldResize
                    Selection.Value = driver(y)
                
                Case 2 ' Route Num
                    ActiveCell.Offset(1, 1).Select
                    
                    If InStr(driver(y), "/") Then
                        routeNo() = Split(driver(y), "/")
                        Selection.Value = "Route " & routeNo(0)
                    
                    Else
                    Selection.Value = "Route " & driver(y)
                    
                    End If
                
                Case 3 ' Truck Num
                    ActiveCell.Offset(-1, 0).Select
                    Selection.Value = "Truck " & CStr(driver(y))
                
                Case 4 ' Number of Stops
                    
                    k = k + 1
                    If InStr(driver(y), "/") Then
                        noStops() = Split(driver(y), "/")
                        
                        'Debug.Print noStops(0), noStops(1)
                        
                        sumStops = 0
                        
                        If InStr(noStops(0), "-") Then
                            noStops() = Split(noStops(0), "-")
                            
                            
                            For i = 0 To UBound(noStops) - 1
    
                                sumStops = sumStops + CInt(noStops(i))
                                Next i
                             
                             
                            Set phRng = ActiveCell
                            
                            'Debug.Print sumStops
                            
                            Call InnerBorders(Range(ActiveCell.Offset(2, -1), ActiveCell.Offset(sumStops + 2, 0)))
                            
                            phRng.Select
                            ActiveCell.Offset(1 + sumStops + 6, -1).Select
                            
                        Else:
                            
                            Set phRng = ActiveCell
                            Call InnerBorders(Range(ActiveCell.Offset(2, -1), ActiveCell.Offset(CInt(noStops(0)) + 2, 0)))
                            phRng.Select
                            ActiveCell.Offset(1 + CInt(noStops(0)) + 6, -1).Select
                        
                        End If
                    
                    ElseIf InStr(driver(y), "-") Then
                        noStops() = Split(driver(y), "-")
                        
                        'Debug.Print noStops(0), noStops(1), noStops(2)
                        
                        'noStops() = Split(noStops(), "-")
                        
                        sumStops = 0
                            
                            For i = 0 To UBound(noStops)
                                
                                sumStops = sumStops + CInt(noStops(i))
                                
                                Next i
                                
                      '  Debug.Print sumStops
                        
                        Set phRng = ActiveCell
                        
                        Call InnerBorders(Range(ActiveCell.Offset(2, -1), ActiveCell.Offset(sumStops + 1, 0)))
                        
                        phRng.Select
                        
                        ActiveCell.Offset(1 + sumStops + 6, -1).Select
                    
                    Else:
                    
                        Set phRng = ActiveCell
                            
                        Call InnerBorders(Range(ActiveCell.Offset(2, -1), ActiveCell.Offset(CInt(driver(y)) + 1, 0)))
                        
                        phRng.Select
                        
                        ActiveCell.Offset(1 + CInt(driver(y)) + 6, -1).Select
                    
                    End If
                         
            End Select
            
        End If
        
    Next y
    
NextIteration:
    
Next x

ActiveCell.Offset(-1, 0).Value = "End"

Call BordersAll(k)

Worksheets(2).Select
      
End Sub

Public Sub StoreVarArray()
Dim tab_rearList As ListObject
Dim x, n As Integer

'Columns: 2 3 5 6 7

Set tab_rearList = Worksheets(3).ListObjects("RearLoaderList")

n = Int(tab_rearList.ListRows.Count)

ReDim ticketInfo(n) As String

For x = 1 To tab_rearList.ListRows.Count
    
    If tab_rearList.DataBodyRange(x, 5) = "" Then
    
        tab_rearList.DataBodyRange(x, 2).Value = "Null"
        
    End If
        
    ticketInfo(x - 1) = tab_rearList.DataBodyRange(x, 2).Value & "," & tab_rearList.DataBodyRange(x, 3).Value & "," _
                                & tab_rearList.DataBodyRange(x, 5).Value & "," & tab_rearList.DataBodyRange(x, 6).Value & "," _
                                & tab_rearList.DataBodyRange(x, 7).Value
    

Next x

'Debug.Print ticketInfo(UBound(ticketInfo) - 1)

' Debug.Print ticketInfo(2)

End Sub

Public Sub BorderTicket()
Attribute BorderTicket.VB_ProcData.VB_Invoke_Func = "T\n14"

Dim var1 As String, var2 As String

var1 = Selection.Address

Selection.End(xlDown).Select
Selection.End(xlDown).Select

Selection.Offset(-1, 1).Select
var2 = Selection.Address

Range(var1, var2).Select

With Selection.Borders(xlEdgeLeft)
        .LineStyle = xlContinuous
       
End With
    
 With Selection.Borders(xlEdgeTop)
     .LineStyle = xlContinuous
 
 End With
 
 With Selection.Borders(xlEdgeBottom)
     .LineStyle = xlContinuous

 End With
 
 With Selection.Borders(xlEdgeRight)
     .LineStyle = xlContinuous
     .Weight = xlThin
 End With
 
 With Selection.Font
     .ColorIndex = xlAutomatic
 End With

Range(var2).Select

End Sub



Sub BoldResize()
Attribute BoldResize.VB_ProcData.VB_Invoke_Func = "B\n14"

    ActiveCell.Range("A1:B2").Select
    Selection.Font.Bold = True
    ActiveCell.Offset(1, 0).Range("A1").Select
    With Selection
        .HorizontalAlignment = xlLeft
        .VerticalAlignment = xlBottom
        .ReadingOrder = xlContext

    End With
    ActiveCell.Offset(-1, 0).Range("A1:B2").Select
    Selection.Font.Size = 16
    ActiveCell.Columns("A:A").EntireColumn.ColumnWidth = 15.67
    ActiveCell.Offset(0, 1).Columns("A:A").EntireColumn.ColumnWidth = 28
    ActiveCell.Select
    
End Sub

Sub BordersAll(j As Integer)

Dim tab_rearList As ListObject
Dim i As Integer

Worksheets(4).Range("A1").Select

Set tab_rearList = Worksheets(3).ListObjects("RearLoaderList")

For i = 1 To j

    Call BorderTicket
    ActiveCell.Offset(1, -1).Select
    
Next i

End Sub
Sub InnerBorders(rng As Range)
Attribute InnerBorders.VB_ProcData.VB_Invoke_Func = " \n14"
'
' InnerBorderss Macro
'

'
    rng.Select
    
    With Selection.Borders(xlEdgeLeft)
        .LineStyle = xlContinuous
        .Weight = xlThin
    End With
    
    With Selection.Borders(xlEdgeTop)
        .LineStyle = xlContinuous
        .Weight = xlThin
    End With
    With Selection.Borders(xlEdgeBottom)
        .LineStyle = xlContinuous
        .Weight = xlThin
    End With
    With Selection.Borders(xlEdgeRight)
        .LineStyle = xlContinuous
        .Weight = xlThin
    End With
    With Selection.Borders(xlInsideVertical)
        .LineStyle = xlContinuous
        .Weight = xlThin
    End With
    With Selection.Borders(xlInsideHorizontal)
        .LineStyle = xlContinuous
        .Weight = xlThin
    End With


End Sub
