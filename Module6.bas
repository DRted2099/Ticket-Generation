Attribute VB_Name = "Module6"
Public Sub EraseSheet3and4()
Attribute EraseSheet3and4.VB_ProcData.VB_Invoke_Func = " \n14"
'
' EraseSheet3and4 Macro
'

    Sheets("Rear Loader List - Sheet 3").Select
    Cells.Select
    Selection.Delete Shift:=xlUp
    Range("A1").Select
    
    Sheets("Tickets - Sheet 4").Select
    Cells.Select
    Selection.Delete Shift:=xlUp

    Range("A1").Select
    Sheets("Schedule Copy - Sheet 2").Select
    
End Sub


Public Sub EraseSheet4()

Worksheets(4).Select
Cells.Select
Selection.Delete Shift:=xlUp
Range("A1").Select

Sheets("Schedule Copy - Sheet 2").Select


End Sub
