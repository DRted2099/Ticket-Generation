Attribute VB_Name = "Module1"
Option Explicit 'With this enabled, you have to declare the data type of each variable. Ensuring optimal memeory usage

Public Sub filterRear()

' Function to Filter Rears

Call ConvertColumnToText

Dim mystr As String     'string variable to store value from the 'TRUCK NO.' column from the scedule
Dim strSplit() As String 'string array to store thetruck numbers after splitting
Dim tab_rearload As ListObject, tab_schedule As ListObject, rear_filter As ListObject, tab_rearList As ListObject 'Variables to store and call columns of the tables
Dim x As Integer, y As Integer, r As Integer ' x, y are used for looping and r is to keep

Set tab_rearload = ActiveSheet.ListObjects("Rear Loaders")
Set tab_schedule = ActiveSheet.ListObjects("Schedule")

tab_schedule.HeaderRowRange.Copy

Worksheets(3).Select

Range("A1").PasteSpecial xlPasteValues

Worksheets(2).Select

r = 2

For y = 1 To tab_schedule.ListColumns("TRUCK NO.").DataBodyRange.Count

    For x = 1 To tab_rearload.ListColumns("Rear Loaders").DataBodyRange.Count
    
        mystr = tab_schedule.ListColumns("TRUCK NO.").DataBodyRange(y).Value
    
    ' Finds the / and splits the string
    
        If InStr(mystr, "/") > 0 Then
            strSplit() = Split(mystr, "/")
            mystr = strSplit(0)
        
        
        End If
        
    'Copies the rows containing rear loader truck numbers and pastes it in sheet 3
        
        If mystr = tab_rearload.ListColumns("Rear Loaders").DataBodyRange(x) And (tab_schedule.ListColumns("LOAD NO.").DataBodyRange(y).Value <> "-" Or tab_schedule.ListColumns("STOPS").DataBodyRange(y).Value <> "-") Then
        
                tab_schedule.ListRows(y).Range.Copy
                
                Worksheets(3).Select
                Range("A" & r).PasteSpecial xlPasteValues
                r = r + 1
                Range("A1").Select
                
                Worksheets(2).Select
                
                Exit For
                   
        End If

    Next x

Next y

Worksheets(3).Select

Range("A1").Select

If ActiveCell.ListObject Is Nothing Then
        Call ConvertToTable("RearLoaderList")

End If

Set tab_rearList = Worksheets(3).ListObjects("RearLoaderList")
tab_rearList.ShowAutoFilterDropDown = True

Call TableFormat

Worksheets(2).Select
        
End Sub

Public Sub ConvertColumnToText()

    Dim x As Integer
    
    Dim tab_rearload As ListObject, tab_schedule As ListObject
    Dim header_name As String
    Dim schedule As String
    
    Worksheets(2).Select
    
    header_name = "Rear Loaders"
    
    Range("B1").Select
    Selection.End(xlDown).Select
    
    If ActiveCell.ListObject Is Nothing Then
        Call ConvertToTable(header_name)
    
    End If
    
    Selection.ListObject.name = header_name
    ActiveCell.Value = header_name
    
    Set tab_rearload = Worksheets(2).ListObjects(header_name)
    
    tab_rearload.ShowAutoFilterDropDown = False
    
    schedule = "Schedule"
    
    Range("F1").Select
    Selection.End(xlDown).Select
    
    If ActiveCell.ListObject Is Nothing Then
        Call ConvertToTable(schedule)
        
    End If
    
    Selection.ListObject.name = schedule
    
    Set tab_schedule = ActiveSheet.ListObjects(schedule)
    
    tab_schedule.ShowAutoFilterDropDown = False
     
    Call ConvertToText


End Sub

Public Sub ConvertToTable(tableName As String)

Dim tbl As Range
Dim ws As Worksheet

Set tbl = Selection.CurrentRegion
Set ws = ActiveSheet

ws.ListObjects.Add(SourceType:=xlSrcRange, Source:=tbl).name = tableName

End Sub

Sub ConvertToText()

    Range("Rear_Loaders[Rear Loaders]").Select
    Selection.NumberFormat = "@"
    Range("J4").Select
    Range("Schedule[TRUCK NO.]").Select
    Selection.NumberFormat = "@"
    Range("A1").Select

End Sub

Public Sub TableFormat()
'
' TableFormat Macro
'

'
    Range("RearLoaderList[#All]").Select
    Selection.Borders(xlDiagonalDown).LineStyle = xlNone
    Selection.Borders(xlDiagonalUp).LineStyle = xlNone
    
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
    
    With Selection.Borders(xlInsideVertical)
        .LineStyle = xlContinuous
        .Weight = xlThin
    End With
    
    With Selection.Borders(xlInsideHorizontal)
        .LineStyle = xlContinuous
        .Weight = xlThin
    End With
    
    With Selection.Font
        .ColorIndex = xlAutomatic
    End With
    
    Selection.Font.Bold = True
    Selection.Font.Bold = False
    
    Range("A1").Select
    

End Sub


