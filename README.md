# library
library of Macros
Sub DeleteAllTextboxes()

Dim i As Long
With ActiveSheet.Shapes
For i = .Count To 1 Step -1
If .Item(i).Type = msoTextBox Then
.Item(i).Delete
End If
Next i
End With

End Sub

Sub Position()
'
' Milestone Positioning
'

'
xStart = 88.6
xFinish = 574.7
yStart = 40.8
yFinish = 342.4
Dim positionCell As String
Dim labelCell As String

For counter = 0 To 30

positionCell = "X" & CStr(counter + 15)

If Range(positionCell).Value = 0 Then
    Exit Sub
    End If

labelCell = "=$U$" & CStr(counter + 15)


milestonePosition = Range(positionCell)


    ActiveSheet.Shapes.AddConnector(msoConnectorStraight, milestonePosition, yStart, milestonePosition, yFinish).Select
    Selection.Name = "MilestoneLine " & CStr(counter + 1)
    With Selection.ShapeRange.Line
        .ForeColor.RGB = RGB(0, 0, 0)
    End With
    ActiveSheet.Shapes.AddTextbox(msoTextOrientationUpward, milestonePosition + -20, yStart + 20, 20, 280).Select
    Selection.Name = "MilestoneText " & CStr(counter + 1)
    'Selection.ShapeRange.IncrementLeft -0.75
    'ActiveSheet.Shapes.Range(Array("Milestone1")).Select
    'ActiveSheet.Shapes.Range(Array("Text Box 1")).Select
    Selection.Formula = labelCell
    Selection.ShapeRange.Fill.Visible = msoFalse
    Selection.ShapeRange.TextFrame2.TextRange.Font.Bold = msoTrue
    Selection.ShapeRange.Line.Visible = msoFalse
    
Next counter
   
End Sub

Sub KillShapes()
Dim Sh As Shape
With Worksheets("Chart")
   For Each Sh In .Shapes
       If Not Application.Intersect(Sh.TopLeftCell, .Range("C1:M30")) Is Nothing Then
         If Sh.Name <> "TextBox 13" And Sh.Type = msoTextBox Then Sh.Delete
         If Sh.Name <> "TextBox 13" And Sh.Type = msoConnectorStraight Then Sh.Delete
       End If
    Next Sh
End With
End Sub

Sub PrintArea()
'
'


 Dim Lastrow As Long

 ActiveCell.SpecialCells(xlLastCell).Select
 Lastrow = ActiveCell.SpecialCells(xlLastCell).Row
 ActiveSheet.PageSetup.PrintArea = "C24:O" & ActiveCell.SpecialCells(xlLastCell).Row
 End Sub

Sub SetDate()
'
'
Dim dte As Variant
Dim ws As Worksheet

dte = InputBox("Enter Bill through Date")
Sheets("Setup").Range("D24").Value = dte


End Sub
Sub WhiteOut()

If Sheets("Cover Page").Range("B197") = "Enter Comment Here" Then
    Sheets("Cover Page").Range("B197").Font.Color = vbWhite
    Else
    Sheets("Cover Page").Range("B197").Font.Color = vbBlack
    End If
    
End Sub

Sub ResetFont()

Sheets("Cover Page").Range("B197").Font.Color = vbBlack
Sheets("Cover Page").Range("B197") = "Enter Comment Here"
End Sub

Sub HideRowNotes()
'
Application.ScreenUpdating = False
If Range("C15").Value = "Notes:  " Then
     Rows("15:16").Select
    Selection.EntireRow.Hidden = True
    Else
End If

End Sub
Sub Unhide()
Rows("15:16").Select
Selection.EntireRow.Hidden = True
End Sub

Sub DateMatch()
Dim answer As Integer
If Sheets("Cover Page").Range("L13").Value = Sheets("Setup").Range("C14").Value Then
GoTo er
Else
answer = MsgBox("Would you like to update invoice to today's date?", vbYesNo + vbQuestion)
    If answer = vbYes Then
     Sheets("Cover Page").Range("L13").Value = Sheets("Setup").Range("C14").Value
Else
    'do nothing
End If
End If
er:
End Sub

Sub Pathfinder()
  FilePath = Application.GetOpenFilename("Text Files(*.txt), *.txt")
  If FilePath <> False Then
    Range("D26").Value = FilePath
End If
     
End Sub

Sub SelectFolder()
    Dim fd As FileDialog
    Dim sPath As String
    
    Set fd = Application.FileDialog(msoFileDialogFolderPicker)
   
    If fd.Show = -1 Then
        sPath = fd.SelectedItems(1)
    End If
    
    'sPath now holds the path to the folder or nothing if the user clicked the cancel button
    Range("D26").Value = sPath
End Sub
Sub FolderPath()
Range("G26").Value = ThisWorkbook.Path
End Sub


Sub HideSomeRows()
'This macro hides rows dependent on the values in a specific column or columns

Application.ScreenUpdating = False
    
Dim column1 As String
'1: Replace 'A1:A100' in the following line of code with a reference
    'to the column of data used to determine if a row should be hidden
column1 = "I5:I85"
'OPTIONAL: If you have an additional column to test for your hide condition define it here
    'and remove the comment mark
column2 = "X5:X85"

Dim cell As Range, RangeToHide As Range

For Each cell In Range(column1)
    '2: Replace '= 0' in the next line of code with the condition used to determine if a row should be hidden.
        'As it is written this will hide each row that contains a zero in the column
        If cell.Value = 0 And cell.Offset(0, 15).Value = 0 Then 'OPTIONAL: If you need an additional test on another column add the following text before 'Then':
        'AND cell.Offset(0, 1).Value = 0
        'where 1 changes to the number of columns to the right of the first and
        ' '= 0' is replaced with the condition on the 2nd column
            If RangeToHide Is Nothing Then
                Set RangeToHide = cell.EntireRow
            Else
                Set RangeToHide = Union(RangeToHide, cell.EntireRow)
            End If
        End If
    Next

RangeToHide.EntireRow.Hidden = True

column1 = "e89:e154"
'OPTIONAL: If you have an additional column to test for your hide condition define it here
    'and remove the comment mark
'column2 = "X5:X85"

Dim cell2 As Range, RangeToHide2 As Range

For Each cell2 In Range(column1)
    '2: Replace '= 0' in the next line of code with the condition used to determine if a row should be hidden.
        'As it is written this will hide each row that contains a zero in the column
        If cell2.Value = 0 Then 'OPTIONAL: If you need an additional test on another column add the following text before 'Then':
        'AND cell.Offset(0, 1).Value = 0
        'where 1 changes to the number of columns to the right of the first and
        ' '= 0' is replaced with the condition on the 2nd column
            If RangeToHide2 Is Nothing Then
                Set RangeToHide2 = cell2.EntireRow
            Else
                Set RangeToHide2 = Union(RangeToHide2, cell2.EntireRow)
            End If
        End If
    Next

RangeToHide2.EntireRow.Hidden = True


End Sub

Sub SaveToPDF1()
'
'This will save file as .pdf
'

'Choosing the file name
'This section contains several options for choosing the file name

Dim filename As String 'declare variable to hold
Dim sameName As Boolean
sameName = False

'Option 1: Use the value in a cell as the file name.
    'Replace the 'A1' in the following line of code with the correct cell reference
'fileName = Range("A1").Value

'Option 2: Get the filename from a message box.
    'If using this option, remove the comment box befor the next line.
filename = InputBox("Enter the file name.")

'Option 3: Get the filename from the Excel file name. If using this option, remove the comment box befor the next line
'sameName = True


'Choosing Sheets
'This section contains options for which sheets will be included int the pdf.

'Use the following lines of code to include all visible sheets.
Dim ws As Worksheet

For Each ws In Sheets

    If ws.Visible Then ws.Select (False)

Next

'Or duplicate the following line of code for each sheet that should be included
    'For each sheet use the following line of code with "Sheet Name" replaced with the actual sheet name.
    
    'Sheets("Sheet Name").Select (False)
    'Sheets("Sheet Name").Select (False)
    
    
'Export as PDF
'This saves and exports .xlsm file into .pdf formatting.

On Error GoTo ErrHandler:


If sameName Then
    ActiveSheet.ExportAsFixedFormat Type:=xlTypePDF, OpenAfterPublish:=True
    
    Else
              
    ActiveSheet.ExportAsFixedFormat Type:=xlTypePDF, filename:= _
        ThisWorkbook.Path & Application.PathSeparator & filename, _
        Quality:=xlQualityStandard, IncludeDocProperties:=True, IgnorePrintAreas _
        :=False, OpenAfterPublish:=True

End If

ActiveSheet.Select
Exit Sub
ErrHandler:
MsgBox ("Save Failed. PDF with the same name is open in another window. Close to save over it.")

Resume Next


End Sub


Sub PrintPlanLookup()
Application.ScreenUpdating = False
Dim lastColumn
Dim i
Dim ws As Worksheet
Dim strPath As String
Dim myFile As Variant
Dim strFile As String

Worksheets("Lists").Activate
lastColumn = Range("AH2").Value + 4
For i = 5 To lastColumn
Worksheets("Plan Grid").Activate
activeBL = Cells(3, i).Value
activePG = Cells(4, i).Value
Worksheets("Plan Lookup").Activate
Range("B3").Value = activeBL
Range("H3").Value = activePG


Set ws = ActiveSheet


Rows("56:85").EntireRow.Hidden = False
'Selection.EntireRow.Hidden = False
Dim empcount
empcount = 86 - Range("T11").Value
If Range("T11").Value > 0 Then
Rows(empcount & ":86").EntireRow.Hidden = True
'Selection.EntireRow.Hidden = True
End If


strFile = activePG _
           & ".pdf"


For Each v In Array("/", "\", "|", ":", "*", "?", "<", ">", """")
            strFile = Replace(strFile, v, "_")
        Next
strFile = ThisWorkbook.Path & "\Plans\" & strFile
If strFile <> "False" Then
    ws.ExportAsFixedFormat _
        Type:=xlTypePDF, _
        filename:=strFile, _
        Quality:=xlQualityStandard, _
        IncludeDocProperties:=True, _
        IgnorePrintAreas:=False, _
        OpenAfterPublish:=False
End If
Next i

Application.ScreenUpdating = True
End Sub



Sub ErrorHandler()

On Error GoTo ErrHandler:


If sameName Then
    ActiveSheet.ExportAsFixedFormat Type:=xlTypePDF, OpenAfterPublish:=True
    
    Else
              
    ActiveSheet.ExportAsFixedFormat Type:=xlTypePDF, filename:= _
        ThisWorkbook.Path & Application.PathSeparator & filename, _
        Quality:=xlQualityStandard, IncludeDocProperties:=True, IgnorePrintAreas _
        :=False, OpenAfterPublish:=True

End If

ActiveSheet.Select
Exit Sub
ErrHandler:
MsgBox ("Save Failed. PDF with the same name is open in another window. Close to save over it.")

Resume Next


End Sub


Sub Unhide_All_Sheets()
'Unhide all sheets

Application.ScreenUpdating = False

Dim ws As Worksheet
For Each ws In ActiveWorkbook.Worksheets
    ws.Visible = xlSheetVisible
Next ws

Worksheets("Setup").Activate

End Sub

Sub Save_as_Project_InvoiceAHTDNEW()
'
' Save_as_Project_Invoice Macro
Application.DisplayAlerts = False  ' This eliminates the error that Excel gives when saving over a file that already exists

Dim WarningText As String

If Dir(ThisWorkbook.Path & "\" & Sheets("SETUP").Cells(15, "C").Value & "-" & Sheets("SETUP").Cells(4, "C").Value & ".xlsb") = "" Then
'Check to see if file with invoice number already exists
    GoTo Saved
Else
    WarningText = "WARNING!  File Already Exists!  Do you still want to "
End If
If MsgBox(WarningText & "Save as    " & Sheets("Setup").Cells(15, "C").Value & "-" & Sheets("Setup").Cells(4, "C").Value & "    to    " & ThisWorkbook.Path & "    ?", vbYesNo + vbQuestion) = vbNo Then
    Exit Sub  'if the user decides not to save the file
Saved:
    End If

    ActiveWorkbook.SaveAs filename:=ThisWorkbook.Path & "\" & Sheets("SETUP").Cells(15, "C").Value & "-" & Sheets("SETUP").Cells(4, "C").Value, _
        FileFormat:=xlExcel12, Password:="", WriteResPassword:="", _
        ReadOnlyRecommended:=False, CreateBackup:=False
    Application.DisplayAlerts = True
    


End Sub
Sub Save_as_Project_InvoiceStandardNEW()

'Saves this file in the folder that it currently is in but with the invoice number as the name.

Application.DisplayAlerts = False  ' This eliminates the error that Excel gives when saving over a file that already exists

Call IfNotes    '10/31/14 JNP   Looks to see if there are notes

Dim WarningText As String

If Dir(ThisWorkbook.Path & "\" & Sheets("Setup").Cells(21, 2).Value & "-" & Sheets("Setup").Cells(24, 3).Value & ".xlsb") = "" Then  'Check to see if file with invoice number already exists
    GoTo Save
Else
    WarningText = "WARNING!  File Already Exists!  Do you still want to "
End If

'Display a Message telling the user if they are saving over a file, the directory they are saving in, and the name of the file.  Gives the option to save or not
If MsgBox(WarningText & "Save as    " & Sheets("Setup").Cells(21, 2).Value & "-" & Sheets("Setup").Cells(24, 3).Value & "    to    " & ThisWorkbook.Path & "    ?", vbYesNo + vbQuestion) = vbNo Then
    Exit Sub  'if the user decides not to save the file
Save:
End If   'Save the file if the user selects "Yes"
    ActiveWorkbook.SaveAs filename:=ThisWorkbook.Path & "\" & Sheets("Setup").Cells(21, 2).Value & "-" & Sheets("Setup").Cells(24, 3).Value, _
        FileFormat:=50, Password:="", WriteResPassword:="", _
        ReadOnlyRecommended:=False, CreateBackup:=False
    Application.DisplayAlerts = True
    


End Sub
Sub PrinttoPDFAHTD_withextras3()

   
Call Exhibit_1_Hide_Zero__Lines
Call Exhibit3_Hide_Unused_Lines
Call Exhibit_1A_Hide_Zero__Lines
Call Exhibit_2A_Hide_Unused_Lines
Call Unhide_All_Sheets

    Dim rateVisible As Boolean
    Dim empVisible As Boolean
    Dim setupVisible As Boolean
    
    rateVisible = Sheets("Rate Schedule").Visible
    empVisible = Sheets("Employees").Visible
    setupVisible = Sheets("Setup").Visible
    
    Sheets("Rate Schedule").Visible = False
    Sheets("Employees").Visible = False
    Sheets("Setup").Visible = False
    
Call HideUnusedSheets


  
Dim SheetIndexer As Integer
Dim SheetsToSelect() As String
Dim size As Integer
size = 0

For SheetIndexer = 1 To Sheets.Count
    If Sheets(SheetIndexer).Visible Then
    size = size + 1
    ReDim Preserve SheetsToSelect(size - 1)
    SheetsToSelect(size - 1) = ThisWorkbook.Sheets(SheetIndexer).Name
    End If
Next SheetIndexer
ThisWorkbook.Sheets(SheetsToSelect).Select

Dim WarningText As String

If Dir(ThisWorkbook.Path & "\" & Sheets("Setup").Cells(15, "C").Value & "-" & Sheets("Setup").Cells(4, "C").Value & " Unsigned" & ".pdf") = "" Then  'Check to see if file with invoice number already exists
    GoTo Save
Else
    WarningText = "WARNING!  File Already Exists!  Do you still want to "
End If

'Display a Message telling the user if they are saving over a file, the directory they are saving in, and the name of the file.  Gives the option to save or not
If MsgBox(WarningText & "Save as    " & Sheets("Setup").Cells(15, "C").Value & "-" & Sheets("Setup").Cells(4, "C").Value & " Unsigned" & "    to    " & ThisWorkbook.Path & "    ?", vbYesNo + vbQuestion) = vbNo Then
    Exit Sub  'if the user decides not to save the file
Save:
End If   'Save the file if the user selects "Yes"
ActiveSheet.ExportAsFixedFormat Type:=xlTypePDF, filename:= _
        ThisWorkbook.Path & Application.PathSeparator & Sheets("SETUP").Cells(15, "C").Value & "-" & Sheets("Setup").Cells(4, "C").Value & " Unsigned", _
        Quality:=xlQualityStandard, IncludeDocProperties:=True, IgnorePrintAreas _
        :=False, OpenAfterPublish:=True
    
    
    Sheets("Rate Schedule").Visible = rateVisible
    Sheets("Employees").Visible = rateVisible
    Sheets("Setup").Visible = rateVisible

Call HideUnusedSheets

Sheets("Cover").Activate

End Sub


Sub Unhide_()
'Unhide all sheets

Application.ScreenUpdating = False

Dim ws As Worksheet
For Each ws In ActiveWorkbook.Worksheets
    ws.Visible = xlSheetVisible
Next ws

Worksheets("Cover").Activate

End Sub

Sub HideUnusedSheets()

 Dim rateVisible As Boolean
    Dim empVisible As Boolean
    Dim setupVisible As Boolean
    
    rateVisible = Sheets("Rate Schedule").Visible
    empVisible = Sheets("Employees").Visible
    setupVisible = Sheets("Setup").Visible
    
    Sheets("Rate Schedule").Visible = False
    Sheets("Employees").Visible = False
    Sheets("Setup").Visible = False
    
If Sheets("Exhibit1").Range("F114") = 0 Then
    Sheets("Exhibit1").Visible = False
    End If
If Sheets("Exhibit2").Range("E53") = 0 Then
    Sheets("Exhibit2").Visible = False
    End If
If Sheets("Invoice").Range("E28") = 0 Then
    Sheets("Exhibit3").Visible = False
    End If
If Sheets("Exhibit 1A - Title II").Range("F118") = 0 Then
   Sheets("Exhibit 1A - Title II").Visible = False
   End If
If Sheets("Exhibit 2A").Range("E44") = 0 Then
   Sheets("Exhibit 2A").Visible = False
   End If

End Sub


Sub HideD()

Sheets("Cover Page").Visible = True

End Sub

Sub FindPrintArea()
'
'
Dim target1 As String
Dim targert As String
'Dim tracker As Integer
Dim invert As Integer

For Index = 1 To 500
    invert = 500 - Index
    target = "O" & CStr(invert)
If Range(target).Value <> "" Then
    Index = 501
    End If
Next Index

target1 = "C24:" & target
'Range(target1).Select
ActiveSheet.PageSetup.PrintArea = target1

End Sub

Sub CopyPasteChart()
'
' CopyPasteChart Macro
'

'
    ActiveSheet.ChartObjects("Chart 4").Activate
    ActiveChart.Axes(xlValue).MajorGridlines.Select
    ActiveChart.ChartArea.Copy
    Sheets("Sheet2").Select
    ActiveSheet.Paste
End Sub

Sub Hide_Row_Notes()
'
Application.ScreenUpdating = False
    If Range("C15").Value = "Notes:  " Then
        Rows("15:16").Select
        Selection.EntireRow.Hidden = True
    Else
'do nothing
End If


End Sub

Sub ShowHideNotes()

    ActiveSheet.Unprotect Password:="grotjohn2"
    
Dim LinesHidden As Boolean

LinesHidden = False

If Range("C15").Value = "Notes:" & "" Then
If (Range("C15").EntireRow.Hidden = False) Then
   Rows("15:16").Select
   Selection.EntireRow.Hidden = True
LinesHidden = True
End If
End If

If LinesHidden = False Then
Rows("15:16").Select
Selection.EntireRow.Hidden = False
End If


End Sub

Sub ToggleStuff()
If Rows("15:16").EntireRow.Hidden = False Then
    Rows("15:16").Select
    Selection.EntireRow.Hidden = True
Else
    Rows("15:16").Select
    Selection.EntireRow.Hidden = False
End If

End Sub
Sub FindFormulaCells()
For Each cl In ActiveSheet.UsedRange
If cl.HasFormula() = True Then
cl.Interior.ColorIndex = 24
End If
Next cl
End Sub
Sub CopyFormulasDown()

Application.ScreenUpdating = False

'Replace the values for the next two variables with the row and column of the first cell that contains the data
checkColumn = "A"
checkRow = "3"

'Use the following two variables to define the first and last cell of the first row of formulas
    'These two cells should be on the same row
firstFormulaCell = "D3"
lastFormulaCell = "E3"

'Use this line of code when the data in the check column sometimes gets shorter
Range(Cells(Range(firstFormulaCell).Row + 1, Range(firstFormulaCell).Column), Cells(Range(lastFormulaCell).End(xlDown).Row, Range(lastFormulaCell).Column)).ClearContents

'To find first row without data (starting from the check row and column), uncomment formula below
Lastrow = ActiveSheet.Range(checkColumn & checkRow).End(xlDown).Row

'To use last nonblank row (starting from bottom), uncomment formula below
    'Generally, use this if your check column has empty cells mixed in with the data
'Lastrow = Range(checkColumn & Rows.Count).End(xlUp).Row

endColumn = Range(lastFormulaCell).Column
Range(firstFormulaCell, lastFormulaCell).Select
Selection.AutoFill Destination:=Range(firstFormulaCell, Cells(Lastrow, endColumn))
Range(checkColumn & checkRow).Select

End Sub
Sub LastSaveTime()

Function LastModified() As Date
Application.Volatile
LastModified = ActiveWorkbook.BuiltinDocumentProperties("Last Save Time")
End Function

End Sub

