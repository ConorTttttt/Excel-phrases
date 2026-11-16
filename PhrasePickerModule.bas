Attribute VB_Name = "PhrasePickerModule"
Option Explicit

' This module belongs in PERSONAL.XLSB, not the downloaded workbook.
Public PickerTarget As Excel.Range

Public Sub OpenPhrasePicker()
    Dim libraryPath As String

    If TypeName(Selection) <> "Range" Then
        MsgBox "Select a cell in your downloaded workbook first.", vbInformation
        Exit Sub
    End If

    If ActiveWorkbook Is ThisWorkbook Then
        MsgBox "Please select a cell in a downloaded workbook first.", vbInformation
        Exit Sub
    End If

    Set PickerTarget = ActiveCell
    libraryPath = GetPhraseLibraryPath()
    If Len(libraryPath) = 0 Then Exit Sub

    If frmPhrasePicker.ReadPhrases(libraryPath) Then
        frmPhrasePicker.Show
    Else
        Unload frmPhrasePicker
    End If
End Sub

Public Function GetPhraseLibraryPath() As String
    Dim savedPath As String
    Dim chosenPath As Variant

    savedPath = GetSetting("ExcelPhrasePicker", "Settings", "LibraryPath", "")
    If Len(savedPath) > 0 Then
        On Error Resume Next
        If Len(Dir$(savedPath)) > 0 Then
            GetPhraseLibraryPath = savedPath
            Exit Function
        End If
        On Error GoTo 0
    End If

    chosenPath = Application.GetOpenFilename( _
        FileFilter:="Excel workbooks (*.xlsx),*.xlsx", _
        Title:="Choose the separate PhraseLibrary.xlsx workbook")

    If VarType(chosenPath) = vbBoolean Then Exit Function ' user cancelled

    GetPhraseLibraryPath = CStr(chosenPath)
    SaveSetting "ExcelPhrasePicker", "Settings", "LibraryPath", GetPhraseLibraryPath
End Function

Public Sub ChooseDifferentPhraseLibrary()
    Dim chosenPath As Variant
    chosenPath = Application.GetOpenFilename( _
        FileFilter:="Excel workbooks (*.xlsx),*.xlsx", _
        Title:="Choose a different phrase library workbook")

    If VarType(chosenPath) = vbBoolean Then Exit Sub
    SaveSetting "ExcelPhrasePicker", "Settings", "LibraryPath", CStr(chosenPath)
    MsgBox "Phrase library updated.", vbInformation
End Sub
