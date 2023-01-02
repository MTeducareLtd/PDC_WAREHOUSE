Attribute VB_Name = "ModGridToExcel"
'The definition of the arrays are shown.
Public NumberRows As Long          'Number of Rows
Public NumberColumns As Long       'Number of Columns
Public FieldNames() As String      'Row Heading
Public FieldData() As String       'Data
Public Head1, Head2, Head3, Head4, Head5 As String

Public Function SaveExcelWorksheet_Profund(argProjectName As String, argWorkbookFile As String, argErrorMessage As String) As Integer
'  Calling Arguments
'  argProjectName       Project Name or Identifier
'  argWorkbookFile      Workbook File
'  argErrorMessage      Returned Error Message

'Local Definitions
Dim XL As Object
Dim wrkRow As Long
Dim wrkColumn As Integer
Dim wrkFlag As Long
Dim wrkStringLength As Integer
Dim ExcelWork() As Variant
Dim ExcelRightJustify() As Integer
Dim ExcelColumnWidth() As Single
Dim wrkString As String
Dim wrkRowOut As Integer

'  Set Return and Error Trap
SaveExcelWorksheet_Profund = False
On Error GoTo SaveExcelWorksheet_Error


'  Set the Size of Justify, Width and Work Arrays
ReDim ExcelRightJustify(1 To NumberColumns)
ReDim ExcelColumnWidth(1 To NumberColumns)
ReDim ExcelWork(1 To NumberColumns)

'  Open the Excel Worksheet
Set XL = CreateObject("Excel.Sheet")

'  Write Project Details
With XL.Application
    If (argProjectName <> "") Then
'            .Range(.Cells(1, 1), .Cells(6, 1)).HorizontalAlignment = xlLeft
'        .Cells(1, 1).Value = argProjectName
'        .Cells(1, 1).Font.Bold = True
'        .Cells(1, 1).Font.Size = 14
'        .Cells(3, 1).Value = Head1
'        .Cells(4, 1).Value = Head2
'        .Cells(5, 1).Value = Head3
'        .Cells(6, 1).Value = Head4
'        .Cells(7, 1).Value = Head5
       '.Range(.Cells(1, 1), .Cells(8, 1)).HorizontalAlignment = xlLeft
    End If
'
'   Write Field Names in Row 9 of Worksheet
    For wrkColumn = 1 To NumberColumns
        ExcelColumnWidth(wrkColumn) = 8.11
        wrkStringLength = Len(FieldNames(wrkColumn))
        If (wrkStringLength > 8) Then
            ExcelColumnWidth(wrkColumn) = CSng(wrkStringLength) + 0.11
        End If
        
        ExcelRightJustify(wrkColumn) = True
        '.Columns(wrkColumn).HorizontalAlignment = xlRight
        .Columns(wrkColumn).NumberFormat = "General"
        .Columns(wrkColumn).ColumnWidth = ExcelColumnWidth(wrkColumn)
        ExcelWork(wrkColumn) = FieldNames(wrkColumn)
        
    Next wrkColumn
'   .Range(.Cells(9, 1), .Cells(9, NumberColumns)).HorizontalAlignment = xlCenter
'   .Range(.Cells(1, 1), .Cells(8, 1)).HorizontalAlignment = xlLeft
    .Range(.Cells(1, 1), .Cells(1, NumberColumns)).Value = ExcelWork
    .Range(.Cells(1, 1), .Cells(1, NumberColumns)).Font.Bold = False
'   .Range(.Cells(1, 1), .Cells(8, 1)).HorizontalAlignment = xlLeft
End With

'  Make Excel Visible to a Maximum Size
XL.Application.Visible = True

'  Set the Size of the Work Array
ReDim ExcelWork(1 To 100, 1 To NumberColumns)

'  Set the Number of Rows Output
wrkRowOut = 0

'  Write Data to Excel Worksheet
'Go Through Each Row of Data
For wrkRow = 1 To NumberRows
'   Increase the Number of Rows Exported
    wrkRowOut = wrkRowOut + 1
'   Go Through Each Column of Data
    For wrkColumn = 1 To NumberColumns
'       Obtain and Store Data in Working String and Array
        wrkString = FieldData(wrkRow, wrkColumn)
        ExcelWork(wrkRowOut, wrkColumn) = wrkString

'       Test whether the Data is Numeric or Time and can be left Right Justified
        If (ExcelRightJustify(wrkColumn) = True) Then
            If (wrkString <> "") Then
                wrkFlag = IsNumeric(wrkString)
                'If (wrkFlag = False) Then wrkFlag = IsTime(wrkString)
                ExcelRightJustify(wrkColumn) = wrkFlag

'               If Data is now Left Justified, set the Alignment of the Column
                If (wrkFlag = False) Then
                    With XL.Application
                        '.Columns(wrkColumn).HorizontalAlignment = xlLeft
                        '.Range(.Cells(9, 1), .Cells(9, NumberColumns)).HorizontalAlignment = xlCenter
                        If (wrkColumn = 1) Then
                            '.Range(.Cells(1, 1), .Cells(8, 1)).HorizontalAlignment = xlLeft
                        End If
                    End With
                End If
            End If
        End If
'       Test whether the Column is Wide Enough
        wrkStringLength = Len(wrkString)
        If (wrkStringLength > ExcelColumnWidth(wrkColumn)) Then
            ExcelColumnWidth(wrkColumn) = CSng(wrkStringLength) + 0.11
            XL.Application.Columns(wrkColumn).ColumnWidth = ExcelColumnWidth(wrkColumn)
        End If
    Next wrkColumn

'   Test Whether to Store the Data in the Worksheet
    If (wrkRowOut = 100 Or (wrkRow <= 25 And wrkRowOut = 10)) Then
'
'   Save the Data in the Worksheet
        With XL.Application
            .Range(.Cells(wrkRow - wrkRowOut + 2, 1), _
                .Cells(wrkRow + 1, NumberColumns)).Value = ExcelWork
        End With
        wrkRowOut = 0
    End If

'  Put an Interrupt Test Here
    DoEvents

'Get Next Row of Data
Next wrkRow

'Put any unstored Rows into the Worksheet
With XL.Application
    If (wrkRowOut > 0) Then
        .Range(.Cells(NumberRows - wrkRowOut + 2, 1), _
           .Cells(NumberRows + 1, NumberColumns)).Value = ExcelWork
    End If

'  Reset the Alignment of the Heading Cells
    '.Range(.Cells(9, 1), .Cells(9, NumberColumns)).HorizontalAlignment = xlCenter
    '.Range(.Cells(1, 1), .Cells(8, 1)).HorizontalAlignment = xlLeft
'
'   Set the Title Rows
    .ActiveSheet.PageSetup.PrintTitleRows = .ActiveSheet.Rows("1:10").Address
'
'   Set the Footer
    .ActiveSheet.PageSetup.CenterFooter = "Page &P of &N"
'
'   Set the Page Print Order
'    .ActiveSheet.PageSetup.Order = xlOverThenDown
'
'   Arrange the Windows
'    .Windows.Arrange arrangeStyle:=xlCascade
End With
'
'  Save the Worksheet
    XL.SaveAs argWorkbookFile
''
''  Close Excel Application
    XL.Application.Quit
    Set XL = Nothing
    SaveExcelWorksheet_Profund = True
    Exit Function
'
'  Build a Return Error
SaveExcelWorksheet_Error:
    'wrkErrorMessage "Error Creating Spreadsheet." & Chr$(13) & Chr$(10) & Chr$(10) & Error$
    Exit Function
End Function

Public Function SaveExcelWorksheetOld(argProjectName As String, argWorkbookFile As String, argErrorMessage As String) As Integer
'  Calling Arguments
'  argProjectName       Project Name or Identifier
'  argWorkbookFile      Workbook File
'  argErrorMessage      Returned Error Message

'Local Definitions
Dim XL As Object
Dim wrkRow As Long
Dim wrkColumn As Integer
Dim wrkFlag As Long
Dim wrkStringLength As Integer
Dim ExcelWork() As Variant
Dim ExcelRightJustify() As Integer
Dim ExcelColumnWidth() As Single
Dim wrkString As String
Dim wrkRowOut As Integer

'  Set Return and Error Trap
SaveExcelWorksheetOld = False
On Error GoTo SaveExcelWorksheet_Error


'  Set the Size of Justify, Width and Work Arrays
ReDim ExcelRightJustify(1 To NumberColumns)
ReDim ExcelColumnWidth(1 To NumberColumns)
ReDim ExcelWork(1 To NumberColumns)

'  Open the Excel Worksheet
Set XL = CreateObject("Excel.Sheet")

'  Write Project Details
With XL.Application
    If (argProjectName <> "") Then
'            .Range(.Cells(1, 1), .Cells(6, 1)).HorizontalAlignment = xlLeft
        .Cells(1, 1).Value = argProjectName
        .Cells(1, 1).Font.Bold = True
        .Cells(1, 1).Font.Size = 14
        .Cells(3, 1).Value = Head1
        .Cells(4, 1).Value = Head2
        .Cells(5, 1).Value = Head3
        .Cells(6, 1).Value = Head4
        .Cells(7, 1).Value = Head5
       '.Range(.Cells(1, 1), .Cells(8, 1)).HorizontalAlignment = xlLeft
    End If
'
'   Write Field Names in Row 9 of Worksheet
    For wrkColumn = 1 To NumberColumns
        ExcelColumnWidth(wrkColumn) = 8.11
        wrkStringLength = Len(FieldNames(wrkColumn))
        If (wrkStringLength > 8) Then
            ExcelColumnWidth(wrkColumn) = CSng(wrkStringLength) + 0.11
        End If
        
        ExcelRightJustify(wrkColumn) = True
        '.Columns(wrkColumn).HorizontalAlignment = xlRight
        .Columns(wrkColumn).NumberFormat = "General"
        .Columns(wrkColumn).ColumnWidth = ExcelColumnWidth(wrkColumn)
        ExcelWork(wrkColumn) = FieldNames(wrkColumn)
        
    Next wrkColumn
'   .Range(.Cells(9, 1), .Cells(9, NumberColumns)).HorizontalAlignment = xlCenter
'   .Range(.Cells(1, 1), .Cells(8, 1)).HorizontalAlignment = xlLeft
    .Range(.Cells(9, 1), .Cells(9, NumberColumns)).Value = ExcelWork
    .Range(.Cells(9, 1), .Cells(9, NumberColumns)).Font.Bold = True
'   .Range(.Cells(1, 1), .Cells(8, 1)).HorizontalAlignment = xlLeft
End With

'  Make Excel Visible to a Maximum Size
XL.Application.Visible = True

'  Set the Size of the Work Array
ReDim ExcelWork(1 To 100, 1 To NumberColumns)

'  Set the Number of Rows Output
wrkRowOut = 0

'  Write Data to Excel Worksheet
'Go Through Each Row of Data
For wrkRow = 1 To NumberRows
'   Increase the Number of Rows Exported
    wrkRowOut = wrkRowOut + 1
'   Go Through Each Column of Data
    For wrkColumn = 1 To NumberColumns
'       Obtain and Store Data in Working String and Array
        wrkString = FieldData(wrkRow, wrkColumn)
        ExcelWork(wrkRowOut, wrkColumn) = wrkString

'       Test whether the Data is Numeric or Time and can be left Right Justified
        If (ExcelRightJustify(wrkColumn) = True) Then
            If (wrkString <> "") Then
                wrkFlag = IsNumeric(wrkString)
                'If (wrkFlag = False) Then wrkFlag = IsTime(wrkString)
                ExcelRightJustify(wrkColumn) = wrkFlag

'               If Data is now Left Justified, set the Alignment of the Column
                If (wrkFlag = False) Then
                    With XL.Application
                        '.Columns(wrkColumn).HorizontalAlignment = xlLeft
                        '.Range(.Cells(9, 1), .Cells(9, NumberColumns)).HorizontalAlignment = xlCenter
                        If (wrkColumn = 1) Then
                            '.Range(.Cells(1, 1), .Cells(8, 1)).HorizontalAlignment = xlLeft
                        End If
                    End With
                End If
            End If
        End If
'       Test whether the Column is Wide Enough
        wrkStringLength = Len(wrkString)
        If (wrkStringLength > ExcelColumnWidth(wrkColumn)) Then
            ExcelColumnWidth(wrkColumn) = CSng(wrkStringLength) + 0.11
            XL.Application.Columns(wrkColumn).ColumnWidth = ExcelColumnWidth(wrkColumn)
        End If
    Next wrkColumn

'   Test Whether to Store the Data in the Worksheet
    If (wrkRowOut = 100 Or (wrkRow <= 25 And wrkRowOut = 10)) Then
'
'   Save the Data in the Worksheet
        With XL.Application
            .Range(.Cells(wrkRow - wrkRowOut + 11, 1), _
                .Cells(wrkRow + 10, NumberColumns)).Value = ExcelWork
        End With
        wrkRowOut = 0
    End If

'  Put an Interrupt Test Here
    DoEvents

'Get Next Row of Data
Next wrkRow

'Put any unstored Rows into the Worksheet
With XL.Application
    If (wrkRowOut > 0) Then
        .Range(.Cells(NumberRows - wrkRowOut + 11, 1), _
           .Cells(NumberRows + 10, NumberColumns)).Value = ExcelWork
    End If

'  Reset the Alignment of the Heading Cells
    '.Range(.Cells(9, 1), .Cells(9, NumberColumns)).HorizontalAlignment = xlCenter
    '.Range(.Cells(1, 1), .Cells(8, 1)).HorizontalAlignment = xlLeft
'
'   Set the Title Rows
    .ActiveSheet.PageSetup.PrintTitleRows = .ActiveSheet.Rows("1:10").Address
'
'   Set the Footer
    .ActiveSheet.PageSetup.CenterFooter = "Page &P of &N"
'
'   Set the Page Print Order
'    .ActiveSheet.PageSetup.Order = xlOverThenDown
'
'   Arrange the Windows
'    .Windows.Arrange arrangeStyle:=xlCascade
End With
'
'  Save the Worksheet
    XL.SaveAs argWorkbookFile
''
''  Close Excel Application
    XL.Application.Quit
    Set XL = Nothing
    SaveExcelWorksheetOld = True
    Exit Function
'
'  Build a Return Error
SaveExcelWorksheet_Error:
    'wrkErrorMessage "Error Creating Spreadsheet." & Chr$(13) & Chr$(10) & Chr$(10) & Error$
    Exit Function
End Function


Public Function SaveExcelWorksheet(argProjectName As String, argWorkbookFile As String, argErrorMessage As String) As Integer
Dim objXL As New Excel.Application
Dim wbXL As New Excel.Workbook
Dim wsXL As New Excel.Worksheet
Dim intRow As Integer ' counter
Dim intCol As Integer ' counter
If Not IsObject(objXL) Then
    MsgBox "You need Microsoft Excel to use this function", _
       vbExclamation, "Print to Excel"
    Exit Function
End If

'On Error Resume Next is necessary because
'someone may pass more rows
'or columns than the flexgrid has'you can instead check for this,
'or rewrite the function so that
'it exports all non-fixed cells
'to Excel

On Error Resume Next ' open Excel
objXL.Visible = True
Set wbXL = objXL.Workbooks.Add
Set wsXL = objXL.ActiveSheet ' name the worksheet
With wsXL
    If Not WorkSheetName = "" Then
        .Name = WorkSheetName
    End If
End With

With wsXL
    .Cells(1, 1).Value = argProjectName
    .Cells(1, 1).Font.Bold = True
    .Cells(1, 1).Font.Size = 14
    
    .Cells(3, 1).Value = Head1 'argWorkbookFile
    .Cells(4, 1).Value = Head2
    .Cells(5, 1).Value = Head3
    .Cells(6, 1).Value = Head4
    .Cells(7, 1).Value = Head5

'   Write Field Names in Row 9 of Worksheet
    For intCol = 1 To NumberColumns
       wsXL.Cells(9, intCol).Value = FieldNames(intCol)
        
    Next
End With



' fill worksheet
For intRow = 1 To NumberRows
    For intCol = 1 To NumberColumns
        wsXL.Cells(intRow + 9, intCol).Value = FieldData(intRow, intCol) & " "
    Next
Next

' format the look
For intCol = 1 To NumberColumns
    wsXL.Columns(intCol).AutoFit
    'wsXL.Columns(intCol).AutoFormat (1)
    wsXL.Range("a1", Right(wsXL.Columns(NumberColumns).AddressLocal, 1) & NumberRows).AutoFormat GridStyle
Next

SaveExcelWorksheet = True
Head2 = ""
Head3 = ""
Head4 = ""
Head5 = ""

Exit Function
End Function

