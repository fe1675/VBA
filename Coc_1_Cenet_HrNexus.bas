Attribute VB_Name = "Coc_1_Cenet_HrNexus"

Public Sub CoC()
    '
    ' Purpose: To determine the chain of command from a Cenet data file using the data from HR Nexus
    '
    ' Author:  Fernando Esqueda (fe1675)
    ' Created: 5/19/2020
    ' Modified: 5/21/2020
    '
    '
    Dim hrNexusWB As String, hrNexus_sheet As String, hrNexus_rows As Long, hrNexus_cols As Integer, _
        cenetWB As String, cenet_sheet As String, cenet_rows As Long, _
        base_cell As Range, find_mgt_lvl_col As Range, max_supv_lvl As Integer, _
        find_supv_col As Range, supv_col As Integer, _
        find_last_col As Range, last_col As Integer, vl_form As String, _
        find_CoC_col As Range, CoC_attuid As String, CoC_level() As String, _
        keep_headers As Variant, save_header As Boolean, find_col As Range
        
    '
    '
    On Error GoTo errorHandler
    
    'UserForm1.Show
    'If UserForm1.CommandButton2 = True Then Exit Sub
    
    
    'hrNexus = UserForm1.ListBox1
    'cenet = UserForm1.ListBox2
    hrNexusWB = "yk.xlsx"
    cenetWB = "cenet-fe1675_v3.xlsx"
    'cenetWB = "cenet_condensed.xlsx"
    
    Windows(hrNexusWB).Activate
    hrNexus_sheet = ActiveSheet.Name
    
    Set base_cell = Range("A1")
    
    base_cell.Select
    
    'Set find_mgt_lvl_col = base_cell.EntireRow.Find(What:="Management Level Indicator")             ' Create an object as a range that finds the cell containing the Management Level Indicator
    'max_supv_lvl = WorksheetFunction.Max(Range(find_mgt_lvl_col.EntireColumn.Address))              ' Use the Management Level Indicator object to find the highest management level in the data set
    
    hrNexus_rows = base_cell.End(xlDown).Row                                                        ' Determine the total rows in the HR Nexus file
    hrNexus_cols = base_cell.End(xlToRight).Column                                                  ' Determine the total columns in the HR Nexus file
    
    ReDim CoC_level(0)                                                                              ' ReDimension the CoC_level array to erase any values in memory
    
    With base_cell.Range(base_cell.Address, base_cell.End(xlToRight).Address)                       ' Establish range to find the number of columns that contain CoC ATTUIDs
        Set find_CoC_col = .Find("CoC Level*", LookIn:=xlValues)                                    ' Establish an object to use the 'Find' method to locate a string that identifies the CoC headers
        If Not find_CoC_col Is Nothing Then                                                         ' String found
            firstAddress = find_CoC_col.Address                                                     ' Note the cell address of the first string found
            Do                                                                                      ' Initiate a loop
                If Right(find_CoC_col, 6) = "ATTUID" Then                                           ' Conditional logic, proceed if the cell contains the string 'ATTUID'
                    CoC_level(UBound(CoC_level)) = find_CoC_col                                     ' Add the string to the CoC_level array
                    
                    ReDim Preserve CoC_level(UBound(CoC_level) + 1)                                 ' ReDimension the CoC_level array to increase space for another element
                End If
                
                Set find_CoC_col = .FindNext(find_CoC_col)                                          ' Find the next cell that contains a CoC header string
                
            Loop While Not find_CoC_col Is Nothing And find_CoC_col.Address <> firstAddress         ' Loop until a CoC header string is not found or you are returned to the first cell
        End If
    End With
    
    Windows(cenetWB).Activate
    cenet_sheet = ActiveSheet.Name
    
    Set base_cell = Range("A1")
    base_cell.Select
    
    attuid_col = base_cell.EntireColumn.Address                                                     ' Determine the address of the column that contains the base cell
    supv_col = 0
    cenet_rows = base_cell.End(xlDown).Row
    last_col = base_cell.End(xlToRight).Column
    save_header = vbEmpty
    
    keep_headers = Array("ATTUID", "MGT_LEVEL_INDICATOR", "SUPERVISOR_ATTUID", "WORK_STATE", _
                        "EMP_STATUS_CODE", "CENET_ID", "CONSULTANT")

    For i% = 1 To last_col                                                                          ' Cycle through all columns that contain data
        header = base_cell.Offset(0, i% - 1).Value                                                  ' Determine the header name
        If header = "" Then Exit For                                                                ' Exit the loop if the column is empty
        For j% = LBound(keep_headers) To UBound(keep_headers)                                       ' Loop through the array that contains the header names of the coulmns to be deleted
            Select Case header
                Case keep_headers(j%)
                    save_header = True
            End Select
        Next j%
        
        If Not save_header = True Then
            base_cell.Offset(0, i% - 1).EntireColumn.Delete shift:=xlShiftToLeft                    ' Delete the entire column for any header that does not match the array
            i% = i% - 1                                                                             ' Reduce i% by one as the column was deleted
        End If
        save_header = False
    Next i
    
    
    Set find_supv_col = base_cell.EntireRow.Find(What:="SUPERVISOR_ATTUID")                         ' Create an object as a range that finds the cell containing the supervisor data
    supv_col = find_supv_col.Column                                                                 ' Determine the column address for the cell that contains the Supervisor ATTUID and assign to variable 'supv_col'
    
    last_col = base_cell.End(xlToRight).Column
    
    For i% = LBound(CoC_level) To UBound(CoC_level)
        If Not CoC_level(i%) = "" Then
            base_cell.Offset(0, last_col + i%) = CoC_level(i%)
            
            vl_form = "=iferror(VLOOKUP(RC[-" & (last_col + i%) & "],'[" & hrNexusWB & "]" & _
                       hrNexus_sheet & "'!R1C1:R" & hrNexus_rows & "C" & hrNexus_cols & "," & (9 + (i% * 2)) & ",FALSE)," & _
                       "iferror(VLOOKUP(RC[-" & (last_col - supv_col + 1 + i%) & "],'[" & hrNexusWB & "]" & _
                       hrNexus_sheet & "'!R1C1:R" & hrNexus_rows & "C" & hrNexus_cols & "," & (9 + (i% * 2)) & ",FALSE)," & """" & "" & """" & "))"
                       
            base_cell.Offset(1, last_col + i%).FormulaR1C1 = vl_form
        End If
    Next i%
    
    Set find_col = base_cell.EntireRow.Find(What:="CENET_ID")
    col% = find_col.Column
    
    base_cell.Offset(0, col% - 1).EntireColumn.Cut
    base_cell.End(xlToRight).Offset(0, 1).Insert shift:=xlToRight
    
    base_cell.Select
    last_col = base_cell.End(xlToRight).Column
    
    For i% = 1 To last_col
        header = base_cell.Offset(0, i%)
        
        Select Case Left(UCase(header), 5)
            Case "MGT_L"
                base_cell.Offset(0, i%).EntireColumn.ColumnWidth = 10
            Case "SUPER"
                base_cell.Offset(0, i%).EntireColumn.ColumnWidth = 18
            Case "WORK_"
                base_cell.Offset(0, i%).EntireColumn.ColumnWidth = 8.3
            Case "EMP_S"
                base_cell.Offset(0, i%).EntireColumn.ColumnWidth = 11
            Case "CONSU"
                base_cell.Offset(0, i%).EntireColumn.ColumnWidth = 12
            Case "COC L"
                base_cell.Offset(0, i%).EntireColumn.ColumnWidth = 17.75
            Case "CENET"
                base_cell.Offset(0, i%).EntireColumn.ColumnWidth = 10
        End Select
    Next i%
    
    Set find_col = base_cell.EntireRow.Find(What:="CoC Level 1 ATTUID")
    col% = find_col.Column
    
    base_cell.Offset(cenet_rows - 1, col% - 1) = "CoC Placeholder"                                  ' Locate the last row and insert a string as a placeholder
    
    base_cell.Range(base_cell.Offset(1, col% - 1), base_cell.Offset(1, col% - 1).End(xlToRight).Offset(0, -1)).Select ' Select formulas and fill down to the last row
    Range(Selection, Selection.End(xlDown)).Select                                                      '
    Selection.FillDown                                                                                  '
    
    base_cell.Range(base_cell.Offset(0, col% - 1), base_cell.Offset(0, col% - 1).End(xlToRight).Offset(0, -1)).Select ' Convert formulas to values
    Range(Selection, Selection.End(xlDown)).Select                                                      '
    Selection.Copy                                                                                      '
    Selection.PasteSpecial Paste:=xlPasteValues, Operation:=xlNone, SkipBlanks _
        :=False, Transpose:=False                                                                       '
    Application.CutCopyMode = False                                                                     '
    
    base_cell.Select
    
    ActiveSheet.Cells.Replace What:=",", Replacement:=":", _
                LookAt:=xlPart, SearchOrder:=xlByColumns, MatchCase:=False, _
                SearchFormat:=False, ReplaceFormat:=False                                           ' Check for any cells with commas. Replace the commas with a colon
    
    Exit Sub
    '
    '
errorHandler:
    '
    msg = MsgBox("File: " & Chr(9) & ActiveWorkbook.Name & Chr(10) & _
                 "Sheet:    " & Chr(9) & ActiveSheet.Name & Chr(10) & _
                 "Cell:     " & Chr(9) & ActiveCell.Address & Chr(10) & _
                 "Error:    " & Chr(9) & Err.Description)
End Sub




