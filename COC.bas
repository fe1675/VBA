Attribute VB_Name = "COC"
Public Sub COC()
Attribute COC.VB_ProcData.VB_Invoke_Func = " \n14"
'
' Macro: Chain of Command
' Author: Fernando Esqueda (FE1675)
' Created: 03/24/2020
' Modified: 03/24/2020
'
' Purpose: Wrap data containing commas with quoatation marks
'

'
    ' Define variable names and types
    Dim totalCols As Long, foundComma As Boolean, attuid As String, cellValue As String, rw As Long
    
    On Error GoTo errorHandler:                                     ' Error handling
    
    Application.ScreenUpdating = False                              ' Freeze the display to reduce run time
    
    totalCols = Range("A1").SpecialCells(xlLastCell).Column         ' Count the number of columns and assign value to the variable, 'totalCols'
    
    Set baseCell = Range("A1")                                      ' Offsets are based from cell "A1"
    
    baseCell.Select                                                 ' Select the cell designated in variable 'baseCell'
    totalColsWcommas = 0
    
    For i% = 1 To totalCols                                         ' Cycle through all columns that contain data
        With baseCell.Offset(0, i% - 1).EntireColumn                ' Designate the column to search for commas
            Set c = .Find(",", LookIn:=xlValues)                    ' Search for commas
            If Not c Is Nothing Then                                ' Perform conditional logic if commas are present
                rw = 1                                              ' Set row value to 1
                attuid = baseCell.Offset(rw, 0).Value               ' Assign attuid to variable
                Do While attuid <> ""                               ' Continue sub-routine within each column until reaching the last attuid
                    cellValue = baseCell.Offset(rw, i% - 1).Value   ' Determine contents of cell
                    If cellValue <> "" Then baseCell.Offset(rw, i% - 1).Value = """" & cellValue & """"     ' Wrap cell contents in commas if the cell is not blank
                    rw = rw + 1                                     ' Increase row value by 1
                    attuid = baseCell.Offset(rw, 0).Value           ' Determine the next attuid
                Loop
            End If
        End With
    Next i
    
    Application.ScreenUpdating = True                               ' Refresh the display
    
    Exit Sub                                                        ' Terminate the routine
    '
    '
errorHandler:
    '
    msg = MsgBox("Address: " & ActiveCell.Address & Chr(10) & Err.Description)
End Sub
