Attribute VB_Name = "VBALibrary"
Option Explicit

' ==========================================================================================================================================================================
'                                   Colors
' ==========================================================================================================================================================================

Public Const cCLEAR As Long = -4142 'use colorindex instead of color for this one
Public Const cBLACK As Long = 0
Public Const cBEIGE As Long = 12900829
Public Const cBLUE1 As Long = 15849925 'lightest
Public Const cBLUE2 As Long = 14136213
Public Const cBLUE3 As Long = 13929812
Public Const cBLUE4 As Long = 15773696 'darkest
Public Const cGREEN1 As Long = 12379352 'lightest
Public Const cGREEN2 As Long = 5296274
Public Const cGREEN3 As Long = 5287936 'darkest
Public Const cPURPLE As Long = 13082801
Public Const cGRAY1 As Long = 14277081 'lightest
Public Const cGRAY2 As Long = 12566463
Public Const cGRAY3 As Long = 8421504 'darkest
Public Const cORANGE As Long = 49407
Public Const cYELLOW As Long = 65535
Public Const cRED As Long = 255
Public Const cMAROON As Long = 5066944
Public Const cWHITE As Long = 16777215

' ==========================================================================================================================================================================
'                                   Types used in functions below
' ==========================================================================================================================================================================

' This is used in GetArrayOfFilePaths()
Public Type FileData
    fPath As String
    fParentFolder As String
    fName As String
    fSize As Long
    fDateCreated As Date
    fDateModified As Date
    fType As String
End Type


' ==========================================================================================================================================================================
'                                               Getting last row or column
' ==========================================================================================================================================================================

'Get the last used row for the specified worksheet and column number
Public Function GetLastRowByWorksheetNumber(columnNumber As Integer, worksheetNumber As Integer) As Integer
    GetLastRowByWorksheetNumber = ThisWorkbook.Worksheets(worksheetNumber).Cells(ThisWorkbook.Worksheets(worksheetNumber).Rows.count - 1, columnNumber).End(xlUp).row
End Function

Public Function GetLastRowByWorksheetName(columnNumber As Integer, worksheetName As String) As Integer
    GetLastRowByWorksheetName = ThisWorkbook.Worksheets(worksheetName).Cells(ThisWorkbook.Worksheets(worksheetName).Rows.count - 1, columnNumber).End(xlUp).row
End Function

Public Function GetLastRowByWorksheet(columnNumber As Integer, actualWorksheet As Worksheet) As Integer
    GetLastRowByWorksheet = actualWorksheet.Cells(actualWorksheet.Rows.count - 1, columnNumber).End(xlUp).row
End Function


'Get the last used column for the specified worksheet and row number
Public Function GetLastColumnByWorksheetNumber(rowNumber As Integer, worksheetNumber As Integer) As Integer
    GetLastColumnByWorksheetNumber = ThisWorkbook.Worksheets(worksheetNumber).Cells(rowNumber, ThisWorkbook.Worksheets(worksheetNumber).Columns.count - 1).End(xlToLeft).Column
End Function

Public Function GetLastColumnByWorksheetName(rowNumber As Integer, worksheetName As String) As Integer
    GetLastColumnByWorksheetName = ThisWorkbook.Worksheets(worksheetName).Cells(rowNumber, ThisWorkbook.Worksheets(worksheetName).Columns.count - 1).End(xlToLeft).Column
End Function

Public Function GetLastColumnByWorksheet(rowNumber As Integer, actualWorksheet As Worksheet) As Integer
    GetLastColumnByWorksheet = actualWorksheet.Cells(rowNumber, actualWorksheet.Columns.count - 1).End(xlToLeft).Column
End Function


'Counts the number of rows until a blank cell is encountered.  Returns last non-empty row
Public Function GetRowCountUntilBlank(rowNum As Integer, colNum As Integer, ws As Worksheet) As Integer
    Dim temp As String
    Dim count As Integer
        
    count = 0
    Do
        temp = ws.Cells(rowNum + count, colNum).Value
        If temp = "" Then Exit Do
        count = count + 1
    Loop Until temp = ""

    GetRowCountUntilBlank = count
End Function







' ==========================================================================================================================================================================
'                                   General Worksheet Value related functions
' ==========================================================================================================================================================================

'Returns the column letter of a column number (i.e. Column 5 returns E, 26 returns Z, 27 returns AA, etc.)
Public Function GetColumnLetterFromColumnNumber(sheet As Worksheet, row As Integer, col As Integer) As String
    Dim temp() As String, colLetter As String
    temp = Split(sheet.Cells(row, col).Address, "$")
    colLetter = temp(1)
    GetColumnLetterFromColumnNumber = colLetter
End Function


'Returns the column letter of a column number
Public Function GetColumnLetterFromColumnNumberOnly(col As Integer) As String
    Dim temp() As String, colLetter As String
    temp = Split(ThisWorkbook.Worksheets(1).Cells(1, col).Address, "$")
    colLetter = temp(1)
    GetColumnLetterFromColumnNumberOnly = colLetter
End Function


'Copy and paste a cell from one sheet to another
Public Sub CopyPasteCell(sourceSheet As Worksheet, destSheet As Worksheet, sourceRow As Integer, sourceCol As Integer, destRow As Integer, destCol As Integer)
    sourceSheet.Cells(sourceRow, sourceCol).Copy
    destSheet.Cells(destRow, destCol).PasteSpecial xlPasteValues
End Sub


'Selects the min or max value in a range of values
Public Sub SelectMinOrMaxOfRange(Optional findMin As Boolean = True)
    
    Dim i As Long, index As Long
    Dim r As Range
    Dim v As Variant
    
    Set r = Selection
    
    If r.Columns.count = 1 Then
        
        ' Make the initial not a blank value
        For i = 1 To r.Rows.count
            If IsEmpty(r(i, 1)) = False Then
                v = r(i, 1)
                Exit For
            End If
        Next
        
        index = i
        
        'Find the min or max in the range
        For i = 1 To r.Rows.count
            If findMin = True Then
                If r(i, 1).Value < v And IsEmpty(r(i, 1)) = False Then
                    v = r(i, 1)
                    index = i
                End If
            Else
                If r(i, 1).Value > v Then
                    v = r(i, 1)
                    index = i
                End If
            End If
        Next
        r(index, 1).Select
    Else
        MsgBox "This only works on a single column selection"
    End If
    
End Sub


' Copy a range to a new destination removing all the blank rows
Sub CopyRangeWithoutBlanks()
    Dim i As Integer, j As Integer, count As Integer
    Dim sourceRange As Range, destRange As Range
    
    Set sourceRange = Selection
    Set destRange = Application.InputBox("Select Destination Cell to Paste values into", Type:=8)
    
    count = 0
    For i = 1 To sourceRange.Rows.count
        If sourceRange(i, 1) <> "" Then
            For j = 1 To sourceRange.Columns.count
                destRange.Offset(count, j - 1) = sourceRange(i, j)
            Next
            count = count + 1
        End If
    Next
End Sub


'Inserts a number of blank rows
Sub InsertBlankRows(numRows As Integer)
    Dim i As Integer
    For i = 1 To numRows
        Selection.EntireRow.Insert Shift:=xlShiftDown
    Next
End Sub

'Delete rows in a selection where the cells are blank
Public Sub DeleteBlankRows()
    
    Dim i As Integer
    Dim count As Integer
    Dim r As Range
    
    Set r = Selection
    
    If r.Columns.count > 1 Then
        MsgBox "You can only select 1 column"
        Exit Sub
    End If
    
    count = 0
    For i = r.Rows.count To 1 Step -1
        If r(i, 1) = "" Then r(i, 1).EntireRow.Delete
    Next
End Sub










' ==========================================================================================================================================================================
'                                   Worksheet Appearance related functions
' ==========================================================================================================================================================================

' Adds the standard black borders around a range of cells
Public Sub AddBlackBordersAroundCells(borderRange As Range)
    borderRange.Borders.LineStyle = xlContinuous
    borderRange.BorderAround xlContinuous, xlThick, xlColorIndexAutomatic, 0
End Sub

'Add a thick border to the cells within a range
Public Sub AddThickBorderToRange(r As Range)
    r.BorderAround xlContinuous, xlThick, xlColorIndexAutomatic, RGB(0, 0, 0)
End Sub













' ==========================================================================================================================================================================
'                                   Append to Array functions
' ==========================================================================================================================================================================

'Append a string onto the end of an array
Public Sub AppendToStringArray(ByRef myArray() As String, text As String, ByRef count As Integer)
    count = count + 1
    ReDim Preserve myArray(1 To count)
    myArray(count) = text
End Sub

'Append a integer onto the end of an array
Public Sub AppendToIntegerArray(ByRef myArray() As Integer, number As Integer, ByRef count As Integer)
    count = count + 1
    ReDim Preserve myArray(1 To count)
    myArray(count) = number
End Sub

'Append a double onto the end of an array
Public Sub AppendToDoubleArray(ByRef myArray() As Double, number As Double, ByRef count As Integer)
    count = count + 1
    ReDim Preserve myArray(1 To count)
    myArray(count) = number
End Sub



' ==========================================================================================================================================================================
'                                   Directory and File Functions
' ==========================================================================================================================================================================

' Returns false if dirctory doesn't exist
Public Function CheckFolderExists(filePath As String) As Boolean
    If Dir(filePath, vbDirectory) = "" Then CheckFolderExists = False Else CheckFolderExists = True
End Function


' Saves a copy of the workbook to the specified filepath
Public Sub SaveCopyOfWorkbook(filePath As String, wb As Workbook, Optional displayAlerts As Boolean = False)
    Application.displayAlerts = displayAlerts
    wb.SaveCopyAs filePath
    Application.displayAlerts = True
End Sub


'Write a string to a text file
Public Sub WriteStringToTextFile(filePath As String, mode As IOMode, text As String)
    Dim fso As New FileSystemObject
    Dim stream As TextStream
    
    If Dir(filePath) = "" Then Call fso.CreateTextFile(filePath, False) 'create file if it doesn't exist
        
    Set stream = fso.OpenTextFile(filePath, mode, False)
    stream.WriteLine (text)
    stream.Close
End Sub


'Write a string array to a text file
Public Sub WriteStringArrayToTextFile(filePath As String, mode As IOMode, text() As String)
    Dim i As Integer
    Dim fso As New FileSystemObject
    Dim stream As TextStream
    
    If Dir(filePath) = "" Then Call fso.CreateTextFile(filePath, False) 'create file if it doesn't exist
    
    Set stream = fso.OpenTextFile(filePath, mode, False)
    
    For i = LBound(text) To UBound(text)
        stream.WriteLine (text(i))
    Next
    
    stream.Close
End Sub


' Returns an array of all the filePaths from the folderPath and any sub folders within that folder
Public Function GetArrayOfFilePaths(basePath As String) As FileData()
    
    Dim fso As New FileSystemObject
    Dim filePaths() As FileData
    Dim count As Integer
    Dim baseFolder As Folder
    
    count = 0

    Set baseFolder = fso.GetFolder(basePath)
    Call GetAllFilesFromFolder(baseFolder, count, filePaths)
    
    GetArrayOfFilePaths = filePaths
End Function


' Returns all files in all folders for a particular base folder - called from function above
Private Sub GetAllFilesFromFolder(baseFolder As Folder, ByRef count As Integer, ByRef filePaths() As FileData)
    
    Dim subFolder As Folder
    Dim aFile As File
    Dim arrayForExtension() As String
    
    If baseFolder.SubFolders.count > 0 Then
        For Each subFolder In baseFolder.SubFolders
            Call GetAllFilesFromFolder(subFolder, count, filePaths) 'Call GetFiles recursively
        Next
    End If
    
    If baseFolder.Files.count > 0 Then
        For Each aFile In baseFolder.Files
            arrayForExtension = Split(aFile.Name, ".")
            count = count + 1
            ReDim Preserve filePaths(1 To count)
            filePaths(count).fPath = aFile.Path ' add filepath to array
            filePaths(count).fParentFolder = aFile.ParentFolder
            filePaths(count).fName = aFile.Name
            filePaths(count).fSize = aFile.Size / 1024
            filePaths(count).fDateCreated = Format(aFile.DateCreated, "dd/mm/yy")
            filePaths(count).fDateModified = Format(aFile.DateLastModified, "dd/mm/yy")
            filePaths(count).fType = arrayForExtension(UBound(arrayForExtension))
        Next
    End If
End Sub


'Starts an exe file
Public Function StartExecutable(filePath As String)
    Dim fso As New FileSystemObject
    
    If fso.FileExists(filePath) = False Then
        MsgBox "Cannot find exe file : " & filePath
    Else
        Shell filePath, vbNormalFocus
    End If
End Function

'Opens a txt file in notepad
Public Function StartNotePad(filePath As String)
    Dim fso As New FileSystemObject
    
    If fso.FileExists(filePath) = False Then
        MsgBox "Cannot find text file : " & filePath
    Else
        Shell "notepad.exe " & filePath, vbNormalFocus
    End If
End Function


'Opens a folder in explorer
Public Function OpenFolder(folderPath As String)
    Dim fso As New FileSystemObject
    
    If CheckFolderExists(folderPath) = False Then
        MsgBox "Cannot find folder : " & folderPath
    Else
        Debug.Print folderPath
        Shell "explorer.exe " & Chr(34) & folderPath & Chr(34), vbNormalFocus
    End If
End Function









' ==========================================================================================================================================================================
'                                   Date related functions
' ==========================================================================================================================================================================

'Returns the days of the year as a string array, eg. (Mon, 11 Feb)
Public Function GetDaysOfYear() As String()
    
    Dim i As Integer, count As Integer
    Dim currentDay As String, numDays As Integer
    Dim startDate As Date
    Dim data() As String
       
    startDate = "1/1/" & Year(DateTime.Now) '1st of January this year
    
    'account for leap years to 2024
    If Year(DateTime.Now) = "2016" Or Year(DateTime.Now) = "2020" Or Year(DateTime.Now) = "2024" Then
        numDays = 366
    Else
        numDays = 365
    End If
    
    
    count = 1
    'Loop 365 or 366 times
    For i = 0 To numDays - 1
        currentDay = WeekdayName(Weekday(startDate + i), False, vbSunday)
        If currentDay <> "Saturday" And currentDay <> "Sunday" Then
            ReDim Preserve data(1 To 2, 1 To count)
            data(1, count) = Left(currentDay, 3)
            data(2, count) = Format(startDate + i, "d mmm")
            count = count + 1
        End If
    Next
        
    GetDaysOfYear = data
End Function


'Returns the integer index of todays date in the list of ALL dates for this year
Public Function GetTodayAsIndexOfDates() As Integer
    
    Dim i As Integer
    Dim dates() As String
    Dim today As String, temp As String
    
    dates = GetDaysOfYear
    today = Format(DateTime.Now, "ddd d mmm")
    
    For i = 1 To UBound(dates, 2)
        temp = dates(1, i) & " " & dates(2, i)
        If InStr(temp, "Mar") Then
            Dim x As Integer
            x = 5
        End If
        If temp = today Then
            Exit For
        End If
    Next
    
    If i = UBound(dates, 2) + 1 Then i = 0
        
    GetTodayAsIndexOfDates = i
End Function









' ==========================================================================================================================================================================
'                                   Form Functions
' ==========================================================================================================================================================================

' Display a yes/no message box and return a boolean with the result of the selection, true is Yes
Public Function ShowYesNoMessageBox(message As String, Optional title As String = "") As Boolean
    Dim msgBoxAnswer As String
    msgBoxAnswer = MsgBox(message, vbYesNo, title)
    If msgBoxAnswer = vbYes Then ShowYesNoMessageBox = True Else ShowYesNoMessageBox = False
End Function











' ==========================================================================================================================================================================
'                                   Colour Picker Dialog
' ==========================================================================================================================================================================

'Picks new color - http://vba-corner.livejournal.com/1691.html
Sub ShowColourPickerDialog()
    Const BGColor As Long = 13160660  'background color of dialogue
    Const ColorIndexLast As Long = 32 'index of last custom color in palette
    
    Dim i_OldColor As Double
    Dim myOrgColor As Double, myNewColor As Double
    Dim myRGB_R As Integer, myRGB_G As Integer, myRGB_B As Integer
        
    i_OldColor = Selection.Interior.Color
    myOrgColor = ActiveWorkbook.Colors(ColorIndexLast)
    
    If i_OldColor = xlNone Then
        Call Color2RGB(BGColor, myRGB_R, myRGB_G, myRGB_B)
    Else
        Call Color2RGB(i_OldColor, myRGB_R, myRGB_G, myRGB_B)
    End If
    
    'call the color picker dialogue
    If Application.Dialogs(xlDialogEditColor).Show(ColorIndexLast, myRGB_R, myRGB_G, myRGB_B) = True Then
        Selection.Interior.Color = ActiveWorkbook.Colors(ColorIndexLast)
        ActiveWorkbook.Colors(ColorIndexLast) = myOrgColor
    Else
        If i_OldColor = 16777215 Then
            Selection.Interior.Color = -4142
        Else
            Selection.Interior.Color = i_OldColor
        End If
    End If
End Sub

'Converts a color to RGB values - Part of function above
Private Sub Color2RGB(ByVal i_Color As Long, o_R As Integer, o_G As Integer, o_B As Integer)
    o_R = i_Color Mod 256
    i_Color = i_Color \ 256
    o_G = i_Color Mod 256
    i_Color = i_Color \ 256
    o_B = i_Color Mod 256
End Sub









' ==========================================================================================================================================================================
'                                   VBA IDE functions
' ==========================================================================================================================================================================

' Returns a list of the names of the modules in the modules folder
Public Function GetListOfModules() As String()

    Dim count As Integer
    Dim project As VBIDE.VBProject
    Dim component As VBIDE.VBComponent
    Dim moduleNames() As String
    
    Set project = ThisWorkbook.VBProject
    
    count = 0
    For Each component In project.VBComponents
        If component.Type = vbext_ct_StdModule Then
            count = count + 1
            ReDim Preserve moduleNames(1 To count)
            moduleNames(count) = component.Name
            
        End If
    Next
    
    GetListOfModules = moduleNames
End Function

' Create a sub template
Public Sub CreateSub(moduleName As String, worksheetName As String)

    Dim VBAEditor As VBIDE.VBE
    Dim project As VBIDE.VBProject
    Dim component As VBIDE.VBComponent
    Dim module As VBIDE.CodeModule
    Dim subLineNumber As String
 
    Set VBAEditor = Application.VBE
    Set project = ThisWorkbook.VBProject
    Set component = project.VBComponents.Item(moduleName)
    Set module = component.CodeModule
 
    'Create sub
    With module
        .InsertLines .CountOfLines + 1, ""
        .InsertLines .CountOfLines + 1, "'TempSub"
        .InsertLines .CountOfLines + 1, "Public Sub TempSubAuto" & .CountOfLines & "()"
        .InsertLines .CountOfLines + 1, "    "
        .InsertLines .CountOfLines + 1, "    Dim i as Integer, j as integer"
        .InsertLines .CountOfLines + 1, "    Dim count as Integer"
        .InsertLines .CountOfLines + 1, "    Dim tempString as String"
        .InsertLines .CountOfLines + 1, "    Dim r as Range"
        .InsertLines .CountOfLines + 1, "    Dim ws as Worksheet"
        .InsertLines .CountOfLines + 1, "    "
        .InsertLines .CountOfLines + 1, "    Set ws = ThisWorkbook.Worksheets(" & Chr(34) & worksheetName & Chr(34) & ")"
        .InsertLines .CountOfLines + 1, "    "
        .InsertLines .CountOfLines + 1, "    count = 0"
        .InsertLines .CountOfLines + 1, "    For i = 1 to 1"
        .InsertLines .CountOfLines + 1, "        "
        .InsertLines .CountOfLines + 1, "    Next"
        .InsertLines .CountOfLines + 1, "    "
        .InsertLines .CountOfLines + 1, "    "
        .InsertLines .CountOfLines + 1, "    "
        .InsertLines .CountOfLines + 1, "End Sub"
    End With
    
    'Open VBA editor at the sub
    VBAEditor.MainWindow.Visible = True
    module.CodePane.Show
    subLineNumber = module.CountOfLines
    module.CodePane.SetSelection subLineNumber, 1, subLineNumber, 1
End Sub
 

