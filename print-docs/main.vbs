Sub BatchPrintWordDocuments()
  Dim objWordApplication As New Word.Application
  Dim strFile As String
  Dim strFolder As String

  On Error GoTo Err
  strFolder = InputBox("Podaj ścieżkę do folderu", "Ścieżka folderu")
  
  If Not Right(strFolder, 1) = "\" Then
      strFolder = strFolder & "\"
  End If
  
  strCopiesCount = InputBox("Ilość kopii", "Ilość kopii", "1")
  strFile = Dir(strFolder & "*.doc*", vbNormal)
  copiesCount = Int(strCopiesCount)
 
  While strFile <> ""
      With objWordApplication
        .Documents.Open (strFolder & strFile)
        .ActiveDocument.PrintOut Copies:=copiesCount, ManualDuplexPrint:=False
        .ActiveDocument.Close
      End With
    strFile = Dir()
  Wend
 
  Set objWordApplication = Nothing
 
  MsgBox "Wszystkie pliki zostały przekazane do druku. Możesz już bezpiecznie zamknąć Word'a."
  
  Exit Sub
Err:
  MsgBox "Coś poszło nie tak. Spróbuj ponownie"
End Sub
