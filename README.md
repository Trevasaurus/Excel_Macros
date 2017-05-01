'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
'''''''''''''''    email     ''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''

Sub Email()
Application.ScreenUpdating = False
Application.Calculation = xlCalculationManual

    Dim OutApp As Object
    Dim OutMail As Object
    Dim rng As Range
    Dim SigString As String
    Dim Signature As String
    Dim StrBody As String
    Dim data_file As String
    Dim file_date As String
    Dim folder_date As String
    Dim week_num As String
    Dim subject_date As String
    Dim secondary_contact As String
    
    secondary_contact = "xxxxxxxxxxxxx"
    
    
    email_date = Range("email_date")
    file_date = Range("CD_DateSave")
    
    total_pnl = Format(Range("total_pnl"), "Currency")
    clean_pnl = Format(Range("clean_pnl"), "Currency")
    activity_pnl = Format(Range("activity_pnl"), "Currency")
    flash_activity = Format(Range("flash_activity"), "Currency")
    flash_legacy = Format(Range("flash_legacy"), "Currency")
    flash_diff = Format(Range("flash_diff"), "Currency")
    flash_diff_adj = Format(Range("flash_diff_adj"), "Currency")
    residual_pnl = Format(Range("residual_pnl"), "Currency")
    
    Set rng = Range("email_table")

''''''Attachment File name & location
    data_file = "file_name" & file_date & ".xlsx"
    Dim sPath As String: sPath = "C:\file pah"
    Dim sFilename1 As String: sFilename1 = sPath & data_file

    Set OutApp = CreateObject("Outlook.Application")
    Set OutMail = OutApp.CreateItem(0)
       
    'create email body
    StrBody = "All,<br><br>" & _
            "Attached are final numbers for COB " & "<b>" & email_date & "</b>. " & _
            "Total Secondary PnL is " & "<b>" & total_pnl & "mm " & "</b>" & _
            "with " & clean_pnl & "mm of Clean PnL and " & _
            residual_pnl & "mm in residual. Today's PnL is primarily driven by <br><br>" & _
            "Flash to Actual PnL difference of " & flash_diff & "mm is driven by " & _
            "Activity PnL of " & flash_activity & "mm and Legacy Book PnL of " & flash_legacy & "mm. " & _
            "Adjusting for these items that are not included in the Flash PnL " & _
            "the remaining difference is " & flash_diff_adj & "mm. <br><br>" & _
            RangetoHTML(rng) & "<br><br>" & _
            "<br><br>" & _
            "Note. <br><br>" & _
            "Please contact xxxxxxxxxxxxxxxxxxxxxx " & _
            "or " & xxxxxxxxxxxx & " email group with any questions or comments. <br><br>" & _
            "PLEASE DO NOT 'Reply All' TO THIS EMAIL <br><br>"
            
    If Dir(SigString) <> "" Then
        Signature = ""
    Else
        Signature = ""
    End If

        On Error Resume Next
        With OutMail
            .SentOnBehalfOfName = "sender email addres"
            .To = "xxxxxxxxxxxxxxxx.com"
            .CC = ""
            .BCC = "xxxxxxxxxxxxxxxxxx.com"
            .Subject = "xxxxxxxxxxxxxxxx " & email_date
            .HTMLBody = StrBody
            .Attachments.Add sFilename1
            .Display
            .VotingOptions = "Approve P&L; Reject P&L"
        End With
        
    On Error GoTo 0

    Set OutMail = Nothing
    Set OutApp = Nothing
    
Application.ScreenUpdating = True
Application.Calculation = xlCalculationAutomatic
End Sub

Function RangetoHTML(rng As Range)

    Dim fso As Object
    Dim ts As Object
    Dim TempFile As String
    Dim TempWB As Workbook

    TempFile = Environ$("temp") & "\" & Format(Now, "dd-mm-yy h-mm-ss") & ".htm"

    'Copy the range and create a new workbook to past the data in
    Range("email_table").Copy
    Set TempWB = Workbooks.Add(1)
    With TempWB.Sheets(1)
        .Cells(1).PasteSpecial Paste:=8
        .Cells(1).PasteSpecial xlPasteValues, , False, False
        .Cells(1).PasteSpecial xlPasteFormats, , False, False
        .Cells(1).Select
        Application.CutCopyMode = False
        On Error Resume Next
        .DrawingObjects.Visible = True
        .DrawingObjects.Delete
        On Error GoTo 0
    End With

    'Publish the sheet to a htm file
    With TempWB.PublishObjects.Add( _
         SourceType:=xlSourceRange, _
         Filename:=TempFile, _
         Sheet:=TempWB.Sheets(1).Name, _
         Source:=TempWB.Sheets(1).UsedRange.Address, _
         HtmlType:=xlHtmlStatic)
        .Publish (True)
    End With

    'Read all data from the htm file into RangetoHTML
    Set fso = CreateObject("Scripting.FileSystemObject")
    Set ts = fso.GetFile(TempFile).OpenAsTextStream(1, -2)
    RangetoHTML = ts.ReadAll
    ts.Close
    RangetoHTML = Replace(RangetoHTML, "align=center x:publishsource=", _
                          "align=left x:publishsource=")

    'Close TempWB
    TempWB.Close savechanges:=False

    'Delete the htm file we used in this function
    Kill TempFile

    Set ts = Nothing
    Set fso = Nothing
    Set TempWB = Nothing
End Function

''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''

Sub Delete_EmptyRows()
'
'This macro will delete all rows that are missing data in a cell
'underneath and including the selected cell.
'Importnant: To avoid run time error, get an accurate row count for your sheet!
'
Dim Counter
Dim i As Integer

Counter = InputBox("Enter the total number of rows to process")

ActiveCell.Select

For i = 1 To Counter

    If ActiveCell = "" Then

        Selection.EntireRow.Delete
        Counter = Counter - 1

    Else

ActiveCell.Offset(1, 0).Select
    End If

Next i

End Sub

''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''

'This will delete all rows that are completely blank.

Sub DeleteALL_blankRows()
Dim R As Long
Dim C As Range
Dim N As Long
Dim Rng As Range

On Error GoTo EndMacro
Application.ScreenUpdating = False
Application.Calculation = xlCalculationManual

If Selection.Rows.Count > 1 Then

Set Rng = Selection

Else
Set Rng = ActiveSheet.UsedRange.Rows

End If

N = 0
For R = Rng.Rows.Count To 1 Step -1
If Application.WorksheetFunction.CountA(Rng.Rows(R).EntireRow) = 0 Then
Rng.Rows(R).EntireRow.Delete
N = N + 1
End If
Next R

EndMacro:

Application.ScreenUpdating = True
Application.Calculation = xlCalculationAutomatic

End Sub




''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''

Sub WrapFormula()

Dim rng As Range
Dim cell As Range
Dim x As String

'Determine if a single cell or range is selected
  If Selection.Cells.Count = 1 Then
    Set rng = Selection
    If Not rng.HasFormula Then GoTo NoFormulas
  Else
    'Get Range of Cells that Only Contain Formulas
      On Error GoTo NoFormulas
        Set rng = Selection.SpecialCells(xlCellTypeFormulas)
      On Error GoTo 0
  End If

'Loop Through Each Cell in Range and add *run_rate
  For Each cell In rng.Cells
    x = cell.Formula
    'cell = "=IFERROR(" & Right(x, Len(x) - 1) & "," & Chr(34) & Chr(34) & ")"
    cell = x & "* run_rate"
  Next cell

Exit Sub

'Error Handler
NoFormulas:
  MsgBox "There were no formulas found in your selection!"

End Sub

''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''

Sub MakeAbsoluteorRelative()

Dim RdoRange As Range
Dim i As Integer
Dim Reply As String
 
'Ask whether Relative or Absolute
Reply = InputBox("Change formulas to?" & Chr(13) & Chr(13) _
 & "Relative row/Absolute column = 1" & Chr(13) _
 & "Absolute row/Relative column = 2" & Chr(13) _
 & "Absolute all = 3" & Chr(13) _
 & "Relative all = 4", "OzGrid Business Applications")
 
   'They cancelled
   If Reply = "" Then Exit Sub
   
    On Error Resume Next
    'Set Range variable to formula cells only
    Set RdoRange = Selection.SpecialCells(Type:=xlFormulas)
 
        'determine the change type
    Select Case Reply
     Case 1 'Relative row/Absolute column
        
        For i = 1 To RdoRange.Areas.Count
            RdoRange.Areas(i).Formula = _
            Application.ConvertFormula _
            (Formula:=RdoRange.Areas(i).Formula, _
            FromReferenceStyle:=xlA1, _
            ToReferenceStyle:=xlA1, ToAbsolute:=xlRelRowAbsColumn)
        Next i
       
     Case 2 'Absolute row/Relative column
        
        For i = 1 To RdoRange.Areas.Count
            RdoRange.Areas(i).Formula = _
            Application.ConvertFormula _
            (Formula:=RdoRange.Areas(i).Formula, _
            FromReferenceStyle:=xlA1, _
            ToReferenceStyle:=xlA1, ToAbsolute:=xlAbsRowRelColumn)
        Next i
       
     Case 3 'Absolute all
        
        For i = 1 To RdoRange.Areas.Count
            RdoRange.Areas(i).Formula = _
            Application.ConvertFormula _
            (Formula:=RdoRange.Areas(i).Formula, _
            FromReferenceStyle:=xlA1, _
            ToReferenceStyle:=xlA1, ToAbsolute:=xlAbsolute)
        Next i
       
      Case 4 'Relative all
        
        For i = 1 To RdoRange.Areas.Count
            RdoRange.Areas(i).Formula = _
            Application.ConvertFormula _
            (Formula:=RdoRange.Areas(i).Formula, _
            FromReferenceStyle:=xlA1, _
            ToReferenceStyle:=xlA1, ToAbsolute:=xlRelative)
        Next i
         
       
     Case Else 'Typo
        MsgBox "Change type not recognised!", vbCritical, _
        "OzGrid Business Applications"
 End Select
 
    'Clear memory
    Set RdoRange = Nothing
End Sub


''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''

Sub Duplicate_Highlight()
   '
   ' NOTE: You must select the first cell in the column AND
   ' sort the column before running the macro
   '
   ScreenUpdating = False
   FirstItem = ActiveCell.Value
   SecondItem = ActiveCell.Offset(1, 0).Value
   Offsetcount = 1
   Do While ActiveCell <> ""
      If FirstItem = SecondItem Then
        ActiveCell.Offset(Offsetcount, 0).Interior.Color = RGB(255, 0, 0)
        Offsetcount = Offsetcount + 1
        SecondItem = ActiveCell.Offset(Offsetcount, 0).Value
      Else
        ActiveCell.Offset(Offsetcount, 0).Select
        FirstItem = ActiveCell.Value
        SecondItem = ActiveCell.Offset(1, 0).Value
        Offsetcount = 1
      End If
   Loop
   ScreenUpdating = True
End Sub

	
	

''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''

Sub error_finder()
' Check for possible missing or erroneous links in
' formulas and list possible errors in a summary sheet

  Dim iSh As Integer
  Dim sShName As String
  Dim sht As Worksheet
  Dim c, sChar As String
  Dim rng As Range
  Dim i As Integer, j As Integer
  Dim wks As Worksheet
  Dim sChr As String, addr As String
  Dim sFormula As String, scVal As String
  Dim lNewRow As Long
  Dim vHeaders

  vHeaders = Array("Sheet Name", "Cell", "Cell Value", "Formula")
  'check if 'Summary' worksheet is in workbook
  'and if so, delete it
  With Application
    .ScreenUpdating = False
    .DisplayAlerts = False
    .Calculation = xlCalculationManual
  End With

  For i = 1 To Worksheets.Count
    If Worksheets(i).Name = "Summary" Then
      Worksheets(i).Delete
    End If
  Next i

  iSh = Worksheets.Count

  'create a new summary sheet
    Sheets.Add After:=Sheets(iSh)
    Sheets(Sheets.Count).Name = "Summary"
  With Sheets("Summary")
    Range("A1:D1") = vHeaders
  End With
  lNewRow = 2

  ' this will not work if the sheet is protected,
  ' assume that sheet should not be changed; so ignore it
  On Error Resume Next

  For i = 1 To iSh
    sShName = Worksheets(i).Name
    Application.Goto Sheets(sShName).Cells(1, 1)
    Set rng = Cells.SpecialCells(xlCellTypeFormulas, 23)

    For Each c In rng
      addr = c.Address
      sFormula = c.Formula
      scVal = c.Text

      For j = 1 To Len(c.Formula)
        sChr = Mid(c.Formula, j, 1)

        If sChr = "[" Or sChr = "!" Or _
          IsError(c) Then
          'write values to summary sheet
          With Sheets("Summary")
            .Cells(lNewRow, 1) = sShName
            .Cells(lNewRow, 2) = addr
            .Cells(lNewRow, 3) = scVal
            .Cells(lNewRow, 4) = "'" & sFormula
          End With
          lNewRow = lNewRow + 1
          Exit For
        End If
      Next j
    Next c
  Next i

' housekeeping
  With Application
    .ScreenUpdating = True
    .DisplayAlerts = True
    .Calculation = xlCalculationAutomatic
  End With

' tidy up
  Sheets("Summary").Select
  Columns("A:D").EntireColumn.AutoFit
  Range("A1:D1").Font.Bold = True
  Range("A2").Select
End Sub

''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''

Sub Import_data()

Application.ScreenUpdating = False
Application.Calculation = xlCalculationManual

Worksheets("worksheet name").Select
Range("AA16").Value = Range("A10").Value
    
''''''Directory Components'''''''''''''
Dim filepath, home, data_file As String
Dim rng As Range
Dim sht As Worksheet
Dim LastColumn As Long
Dim year_folder as string

year_folder = Range("year_folder")
home = ThisWorkbook.Name
data_file = "name of data file.xlsx"

filepath = "C:\file path"

''''''''''''Import PnL'''''''''''''''''''''''''
Application.DisplayAlerts = False

Workbooks.Open Filename:=filepath & data_file, UpdateLinks:=False, ReadOnly:=True
    Worksheets("Email").Select
    Range("O1:O35").Select
    Selection.Copy
    
Windows(home).Activate
    Worksheets("data tab").Select
        LastColumn = Range("A1").CurrentRegion.Columns.Count + 1
        Set rng = Cells(1, LastColumn)
        rng.Select
        Selection.PasteSpecial Paste:=xlPasteValues
        Range(Cells(37, LastColumn - 1), Cells(44, LastColumn)).Select
        Selection.FillRight
        
    'Application.DisplayAlerts = False
    Windows(data_file).Close savechanges:=False
      
Worksheets("control").Select
Range("A3").Select

Calculate
    
Application.ScreenUpdating = True
Application.Calculation = xlCalculationAutomatic

End Sub



''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''

Public Sub SetDataFieldsToSum()

'change highlighted range in pivot table to from count to sum

Dim xPF As PivotField
Dim WorkRng As Range
Set WorkRng = Application.Selection
With WorkRng.PivotTable
   .ManualUpdate = True
   For Each xPF In .DataFields
      With xPF
         .Function = xlSum
         .NumberFormat = "#,##0"
      End With
   Next
   .ManualUpdate = False
End With
End Sub	


''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''

Sub Refresh_Pivot()

Application.ScreenUpdating = False
Application.Calculation = xlCalculationManual

 Dim pvt As PivotTable
 Dim sh As Worksheet

 For Each sh In Worksheets
   For Each pvt In sh.PivotTables
     pvt.RefreshTable
     pvt.PreserveFormatting = True
   Next pvt
 Next sh

Application.ScreenUpdating = True
Application.Calculation = xlCalculationAutomatic

End Sub

''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Sub Refresh_Query_Table()

Application.ScreenUpdating = False
Application.Calculation = xlCalculationManual

 Dim qry As QueryTable
 Dim sh As Worksheet

 For Each sh In Worksheets
   For Each qry In sh.QueryTables
     qry.Refresh
   Next qry
 Next sh

Application.ScreenUpdating = True
Application.Calculation = xlCalculationAutomatic

End Sub



''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''

'this will delete a row based on a constraint in one column (set to delete rows where date < today)

Sub Remove_Old_Dates()
Dim ws1 As Worksheet:   Set ws1 = "name of sheet"
Dim lastrow As Long, icell As Long
Dim restraint as vairant

lastrow = ws1.Range("A" & Rows.Count).End(xlUp).Row 
restraint = date

For icell = lastrow To 1 Step -1
    If ws1.Range("A" & icell).Value < restraint Then 
        ws1.Range("A" & icell).EntireRow.Delete Shift:=xlUp
    End If
Next icell

End Sub


''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''

Sub Reverse_Rows_or_Columns()

'This will transpose a selection of rows or columns.
'Note: you cannot select an etire row or column, but one
'cell less than that will work fine.

Dim Arr() As Variant
Dim Rng As Range
Dim C As Range
Dim Rw As Long
Dim Cl As Long

On Error GoTo EndMacro

Application.ScreenUpdating = False
Application.Calculation = xlCalculationManual
Application.EnableEvents = False

Set Rng = Selection
Rw = Selection.Rows.Count
Cl = Selection.Columns.Count
If Rw > 1 And Cl > 1 Then
MsgBox "Must select either a range of rows or columns, but not simultaneaously columns and rows.", _
vbExclamation, "Reverse Rows or Columns"
Exit Sub
End If

If Rng.Cells.Count = ActiveCell.EntireRow.Cells.Count Then
MsgBox "Can't select an entire row, only up to one cell less than an entire row.", vbExclamation, _
"Reverse Rows or Columns"
Exit Sub
End If

If Rng.Cells.Count = ActiveCell.EntireColumn.Cells.Count Then
MsgBox "Can't select an entire column, only up to one cell less than an entire column.", vbExclamation, _
"Reverse Rows or Columns"
Exit Sub
End If

If Rw > 1 Then
ReDim Arr(Rw)
Else
ReDim Arr(Cl)
End If

Rw = 0
For Each C In Rng
Arr(Rw) = C.Formula
Rw = Rw + 1
Next C

Rw = Rw - 1
For Each C In Rng
C.Formula = Arr(Rw)
Rw = Rw - 1
Next C

EndMacro:
Application.ScreenUpdating = True
Application.Calculation = xlCalculationAutomatic
Application.EnableEvents = True

End Sub


''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''

Public Sub Roll_Data()

    Application.ScreenUpdating = False
        
    Worksheets("sheet1").Select
    
    Range("Pre_range").Value = Range("Cur_range").Value
 
    Worksheets("Exposures").Calculate
    Application.ScreenUpdating = True

End Sub

''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Sub WrapIfError()

Dim rng As Range
Dim cell As Range
Dim x As String

'Determine if a single cell or range is selected
  If Selection.Cells.Count = 1 Then
    Set rng = Selection
    If Not rng.HasFormula Then GoTo NoFormulas
  Else
    'Get Range of Cells that Only Contain Formulas
      On Error GoTo NoFormulas
        Set rng = Selection.SpecialCells(xlCellTypeFormulas)
      On Error GoTo 0
  End If

'Loop Through Each Cell in Range and add =IFERROR([formula],"")
  For Each cell In rng.Cells
    x = cell.Formula
    cell = "=IFERROR(" & Right(x, Len(x) - 1) & "," & Chr(34) & Chr(34) & ")"
  Next cell

Exit Sub

'Error Handler
NoFormulas:
  MsgBox "There were no formulas found in your selection!"

End Sub

''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''


