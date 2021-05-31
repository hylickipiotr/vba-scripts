Declare PtrSafe Function ShellExecute Lib "shell32.dll" _
Alias "ShellExecuteA" (ByVal hwnd As LongPtr, ByVal lpOperation As String, _
ByVal lpFile As String, ByVal lpParameters As String, ByVal lpDirectory As _
String, ByVal nShowCmd As Long) As LongPtr


Sub PrintFile(ByVal strPathAndFilename As String)
  Call ShellExecute(Application.hwnd, "print", strPathAndFilename, vbNullString, vbNullString, 0)
End Sub

Sub BatchPrintPdfDocuments()
  Dim strFile As String
  Dim strFolder As String

  On Error GoTo Err
  strFolder = InputBox("Podaj ścieżkę do folderu", "Ścieżka folderu")

  If Not Right(strFolder, 1) = "\" Then
    strFolder = strFolder & "\"
  End If


  strCopiesCount = InputBox("Ilość kopii", "Ilość kopii", "1")
  strFile = Dir(strFolder & "*.pdf", vbNormal)
  copiesCount = Int(strCopiesCount)

  While strFile <> ""
    For i = 1 To copiesCount
      PrintFile(strFolder & strFile)
    Next i
  Wend

  MsgBox "Wszystkie pliki zostały przekazane do druku. Możesz już bezpiecznie zamknąć Word'a."

  Exit Sub
Err:
  MsgBox "Coś poszło nie tak. Spróbuj ponownie"
End Sub