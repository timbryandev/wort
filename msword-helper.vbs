DIM wdFormat

If Wscript.Arguments.Count > 1 Then
    If Wscript.Arguments(1) = "" Then
        wdFormat = 16 ' .docx by default
    Else
        wdFormat = CInt(WScript.Arguments(1))
    End If
End If


DIM fileIn
    fileIn = WScript.Arguments(0)
DIM fileOut
    fileOut = fileIn
	fileOut = Left(fileOut, InstrRev(fileOut, ".")-1)

SET objWord = CreateObject("Word.Application")
    objWord.Visible = False
    objWord.Documents.Open(fileIn)

SET objDoc = objWord.ActiveDocument
    objDoc.SaveAs fileOut, wdFormat

    objDoc.Close
SET objDoc = Nothing

    objWord.Quit
SET objWord = Nothing
