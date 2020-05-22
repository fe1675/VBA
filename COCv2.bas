Attribute VB_Name = "COC"
Public Sub COC()
Attribute COC.VB_ProcData.VB_Invoke_Func = " \n14"
'
' Macro: Chain of Command
' Author: Fernando Esqueda (FE1675)
' Created: 03/24/2020
' Modified: 03/25/2020
'
' Purpose: Wrap data containing commas with quoatation marks
'

'
    Dim baseCell As Range, totalCols As Long, delHeaders As Variant, header As String   ' Define variable names and types
    
    On Error GoTo errorHandler:                                                         ' Error handling
    
    Application.ScreenUpdating = False                                                  ' Freeze the display to reduce run time
    
    delHeaders = Array("Work State Name", "Bargaining Unit", "CoC Level 1 ATTUID", _
                 "CoC Level 1 Name", "CoC Level 2 Name", "CoC Level 3 Name", _
                 "CoC Level 4 Name", "CoC Level 5 Name", "CoC Level 6 Name", _
                 "CoC Level 7 Name", "CoC Level 8 Name", "CoC Level 9 Name", _
                 "CoC Level 10 Name")                                                   ' Specify the header names of columns that are to be deleted
    
    totalCols = Range("A1").SpecialCells(xlLastCell).Column                             ' Count the number of columns and assign value to the variable, 'totalCols'
    
    Set baseCell = Range("A1")                                                          ' Offsets are based from cell "A1"
    baseCell.Select                                                                     ' Select the cell designated in variable 'baseCell'
    
    For i% = 1 To totalCols                                                             ' Cycle through all columns that contain data
        header = baseCell.Offset(0, i% - 1).Value                                       ' Determine the header name
        If header = "" Then Exit For                                                    ' Exit the cycle if the column is empty
        For j% = LBound(delHeaders) To UBound(delHeaders)                               ' Cycle through the array that contains the header names of the coulmns to be deleted
            If header = delHeaders(j%) Then                                             ' Exeute conditional logic if the header matches a header in the array
                baseCell.Offset(0, i% - 1).EntireColumn.Delete Shift:=xlShiftToLeft     ' Delete the entire column
                i% = i% - 1                                                             ' Reduce i% by one as the column was deleted
            End If
        Next j%
    Next i
    
    ActiveSheet.Cells.Replace what:=",", Replacement:=":", _
                LookAt:=xlPart, SearchOrder:=xlByColumns, MatchCase:=False, _
                SearchFormat:=False, ReplaceFormat:=False                               ' Check for any cells with commas. Replace the commas with a colon
    
    Application.ScreenUpdating = True                                                   ' Refresh the display
    
    Exit Sub                                                        ' Terminate the routine
    '
    '
errorHandler:
    '
    msg = MsgBox("Address: " & ActiveCell.Address & Chr(10) & Err.Description)
End Sub
