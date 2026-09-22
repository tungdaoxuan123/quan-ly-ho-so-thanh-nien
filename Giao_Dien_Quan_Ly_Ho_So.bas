Attribute VB_Name = "GiaoDienQuanLyHoSo"
Option Explicit

Private Const UiSheetName As String = "Quan ly ho so"

Public Sub CreateProfileManagementUI()
    Dim sourceSheet As Worksheet
    Dim uiSheet As Worksheet
    Dim lastRow As Long
    Dim sourceName As String
    Dim formulaPrefix As String
    Dim generateButton As Shape

    Set sourceSheet = ActiveSheet
    If sourceSheet.Name = UiSheetName Then
        MsgBox "Open your data sheet first, then run CreateProfileManagementUI again.", vbExclamation
        Exit Sub
    End If

    sourceName = sourceSheet.Name
    lastRow = sourceSheet.Cells(sourceSheet.Rows.Count, 2).End(xlUp).Row
    If lastRow < 3 Then
        MsgBox "The active sheet does not contain person records starting on row 3.", vbCritical
        Exit Sub
    End If

    Application.DisplayAlerts = False
    On Error Resume Next
    Worksheets(UiSheetName).Delete
    On Error GoTo 0
    Application.DisplayAlerts = True

    Set uiSheet = Worksheets.Add(After:=Worksheets(Worksheets.Count))
    uiSheet.Name = UiSheetName

    On Error Resume Next
    ThisWorkbook.Names("ProfileRecordNumbers").Delete
    On Error GoTo 0
    ThisWorkbook.Names.Add Name:="ProfileRecordNumbers", RefersTo:="='" & Replace(sourceName, "'", "''") & "'!$B$3:$B$" & CStr(lastRow)

    With uiSheet
        .Columns("A").ColumnWidth = 24
        .Columns("B").ColumnWidth = 52
        .Range("A1:B1").Merge
        .Range("A1").Value = "Quan ly ho so thanh nien"
        .Range("A1").Font.Bold = True
        .Range("A1").Font.Size = 18
        .Range("A1").HorizontalAlignment = xlCenter
        .Range("A1:B1").Interior.Color = RGB(31, 78, 121)
        .Range("A1").Font.Color = RGB(255, 255, 255)

        .Range("A2").Value = "Source sheet"
        .Range("B2").Value = sourceName
        .Range("A4").Value = "Choose person (STT)"
        .Range("B4").Validation.Delete
        .Range("B4").Validation.Add Type:=xlValidateList, AlertStyle:=xlValidAlertStop, Formula1:="=ProfileRecordNumbers"
        .Range("B4").Value = sourceSheet.Cells(3, 2).Value
        .Range("A6").Value = "Name"
        .Range("A7").Value = "Date of birth"
        .Range("A8").Value = "CCCD"
        .Range("A9").Value = "Current address"
        .Range("A10").Value = "Occupation"
        .Range("A11").Value = "Parent names"

        formulaPrefix = "=IFERROR(INDEX('" & Replace(sourceName, "'", "''") & "'!"
        .Range("B6").Formula = formulaPrefix & "$C$3:$C$" & lastRow & ",MATCH($B$4,'" & Replace(sourceName, "'", "''") & "'!$B$3:$B$" & lastRow & ",0)),\"\")"
        .Range("B7").Formula = formulaPrefix & "$E$3:$E$" & lastRow & ",MATCH($B$4,'" & Replace(sourceName, "'", "''") & "'!$B$3:$B$" & lastRow & ",0)),\"\")&\"/\"&IFERROR(INDEX('" & Replace(sourceName, "'", "''") & "'!$F$3:$F$" & lastRow & ",MATCH($B$4,'" & Replace(sourceName, "'", "''") & "'!$B$3:$B$" & lastRow & ",0)),\"\")&\"/\"&IFERROR(INDEX('" & Replace(sourceName, "'", "''") & "'!$G$3:$G$" & lastRow & ",MATCH($B$4,'" & Replace(sourceName, "'", "''") & "'!$B$3:$B$" & lastRow & ",0)),\"\")"
        .Range("B8").Formula = formulaPrefix & "$H$3:$H$" & lastRow & ",MATCH($B$4,'" & Replace(sourceName, "'", "''") & "'!$B$3:$B$" & lastRow & ",0)),\"\")"
        .Range("B9").Formula = formulaPrefix & "$O$3:$O$" & lastRow & ",MATCH($B$4,'" & Replace(sourceName, "'", "''") & "'!$B$3:$B$" & lastRow & ",0)),\"\")"
        .Range("B10").Formula = formulaPrefix & "$M$3:$M$" & lastRow & ",MATCH($B$4,'" & Replace(sourceName, "'", "''") & "'!$B$3:$B$" & lastRow & ",0)),\"\")"
        .Range("B11").Formula = formulaPrefix & "$W$3:$W$" & lastRow & ",MATCH($B$4,'" & Replace(sourceName, "'", "''") & "'!$B$3:$B$" & lastRow & ",0)),\"\")&\" / \"&IFERROR(INDEX('" & Replace(sourceName, "'", "''") & "'!$AF$3:$AF$" & lastRow & ",MATCH($B$4,'" & Replace(sourceName, "'", "''") & "'!$B$3:$B$" & lastRow & ",0)),\"\")"
        .Range("A4:A11").Font.Bold = True
        .Range("A4:B11").Borders.LineStyle = xlContinuous
        .Range("A4:B11").Borders.Color = RGB(217, 217, 217)
        .Range("B4").Interior.Color = RGB(255, 242, 204)
        .Range("B6:B11").Interior.Color = RGB(242, 242, 242)
        .Range("A1:B14").VerticalAlignment = xlCenter
        .Rows("1:14").RowHeight = 24
        .Rows(1).RowHeight = 34
    End With

    Set generateButton = uiSheet.Shapes.AddShape(msoShapeRoundedRectangle, 135, 320, 260, 42)
    generateButton.Name = "GenerateWordDocumentButton"
    generateButton.TextFrame2.TextRange.Text = "Generate Word Document"
    generateButton.TextFrame2.TextRange.Font.Size = 16
    generateButton.TextFrame2.TextRange.Font.Fill.ForeColor.RGB = RGB(255, 255, 255)
    generateButton.Fill.ForeColor.RGB = RGB(31, 78, 121)
    generateButton.Line.ForeColor.RGB = RGB(31, 78, 121)
    generateButton.OnAction = "GenerateProfileFromUI"

    uiSheet.Activate
    MsgBox "The profile management screen is ready. Choose a person, then click Generate Word Document.", vbInformation
End Sub

Public Sub GenerateProfileFromUI()
    Dim uiSheet As Worksheet
    Dim sourceSheet As Worksheet
    Dim recordNumber As Variant
    Dim foundCell As Range

    Set uiSheet = Worksheets(UiSheetName)
    Set sourceSheet = Worksheets(CStr(uiSheet.Range("B2").Value))
    recordNumber = uiSheet.Range("B4").Value
    Set foundCell = sourceSheet.Columns(2).Find(What:=recordNumber, LookIn:=xlValues, LookAt:=xlWhole)
    If foundCell Is Nothing Then
        MsgBox "Choose a valid STT number from the dropdown.", vbExclamation
        Exit Sub
    End If
    GenerateProfileForRow foundCell.Row, sourceSheet.Name
End Sub

Private Sub GenerateProfileForRow(ByVal sourceRow As Long, ByVal sourceSheetName As String)
    Dim baseFolder As String
    Dim scriptPath As String
    Dim templatePath As String
    Dim pythonPath As String
    Dim command As String

    baseFolder = ThisWorkbook.Path
    scriptPath = baseFolder & "/tao_ho_so_word.py"
    templatePath = baseFolder & "/Mau_Ho_So_Thanh_Nien_Phan_I.docx"
    pythonPath = baseFolder & "/.venv/bin/python3"
    If Len(baseFolder) = 0 Or Dir(scriptPath) = "" Or Dir(templatePath) = "" Or Dir(pythonPath) = "" Then
        MsgBox "Keep this workbook, generator files, template, and .venv folder together.", vbCritical
        Exit Sub
    End If

    command = QuoteForShell(pythonPath) & " " & QuoteForShell(scriptPath) & _
        " --workbook " & QuoteForShell(ThisWorkbook.FullName) & _
        " --sheet " & QuoteForShell(sourceSheetName) & _
        " --row " & CStr(sourceRow) & _
        " --template " & QuoteForShell(templatePath) & _
        " --output-dir " & QuoteForShell(baseFolder & "/Ho so thanh nien da tao") & _
        " > " & QuoteForShell(baseFolder & "/Tao ho so Word.log") & " 2>&1"
    Shell command, vbNormalFocus
    MsgBox "Generating Word document. Check the Ho so thanh nien da tao folder in a moment.", vbInformation
End Sub

Private Function QuoteForShell(ByVal value As String) As String
    QuoteForShell = Chr$(34) & Replace(value, Chr$(34), "\" & Chr$(34)) & Chr$(34)
End Function
