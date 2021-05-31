Declare Function apiShellExecute Lib "shell32.dll" Alias "ShellExecuteA" ( _
ByVal hwnd As Long, _
ByVal lpOperation As String, _
ByVal lpFile As String, _
ByVal lpParameters As String, _
ByVal lpDirectory As String, _
ByVal nShowCmd As Long) _
As Long

Sub PrintFile(ByVal strPathAndFilename As String)
  Call apiShellExecute(Application.hwnd, "print", strPathAndFilename, vbNullString, vbNullString, 0)
End Sub

Sub BatchPrintPdfDocuments(PartNum As String)
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
    For i As Integer = 1 To copiesCount
      PrintFile(strFolder & strFile)
    Next i
  End While

  MsgBox "Wszystkie pliki zostały przekazane do druku. Możesz już bezpiecznie zamknąć Word'a."

  Exit Sub
Err:
  MsgBox "Coś poszło nie tak. Spróbuj ponownie"
End Sub