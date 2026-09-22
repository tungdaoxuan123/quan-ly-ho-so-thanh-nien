Attribute VB_Name = "GenerateTSQA3"
Option Explicit

Private Function QuoteForShell(ByVal value As String) As String
    QuoteForShell = Chr$(34) & Replace(value, Chr$(34), "\" & Chr$(34)) & Chr$(34)
End Function

Public Sub GenerateTSQA3ForSelectedPerson()
    Dim selectedRow As Long
    Dim baseFolder As String
    Dim command As String
    Dim scriptPath As String
    Dim templatePath As String
    Dim pythonPath As String

    selectedRow = ActiveCell.Row
    If selectedRow < 3 Then
        MsgBox "Select any cell in a person's row (row 3 or later), then try again.", vbExclamation
        Exit Sub
    End If

    baseFolder = ThisWorkbook.Path
    If Len(baseFolder) = 0 Then
        MsgBox "Save this workbook before generating a Word document.", vbExclamation
        Exit Sub
    End If

    scriptPath = baseFolder & "/generate_tsq_a3.py"
    templatePath = baseFolder & "/TSQ_A3_Section_I_Template.docx"
    pythonPath = baseFolder & "/.venv/bin/python3"
    If Dir(scriptPath) = "" Or Dir(templatePath) = "" Then
        MsgBox "Put generate_tsq_a3.py and TSQ_A3_Section_I_Template.docx in the same folder as this workbook.", vbCritical
        Exit Sub
    End If
    If Dir(pythonPath) = "" Then
        MsgBox "Run the one-time virtual-environment setup in README.md before using this button.", vbCritical
        Exit Sub
    End If

    command = QuoteForShell(pythonPath) & " " & QuoteForShell(scriptPath) & _
        " --workbook " & QuoteForShell(ThisWorkbook.FullName) & _
        " --row " & CStr(selectedRow) & _
        " --template " & QuoteForShell(templatePath) & _
        " --output-dir " & QuoteForShell(baseFolder & "/Generated TSQ A3") & _
        " > " & QuoteForShell(baseFolder & "/Generated TSQ A3.log") & " 2>&1"

    Shell command, vbNormalFocus
    MsgBox "Generating Word document. It will be saved in the Generated TSQ A3 folder.", vbInformation
End Sub
