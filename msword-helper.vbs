DIM wdFormat, fileIn, fileOut

If Wscript.Arguments.Count > 1 Then
    If Wscript.Arguments(1) = "" Then
        wdFormat = 16 ' .docx by default
    Else
        wdFormat = CInt(WScript.Arguments(1))
    End If
End If


fileIn = WScript.Arguments(0)
fileOut = Left(fileIn, InstrRev(fileIn, ".")-1)

SET objWord = CreateObject("Word.Application")
    objWord.Visible = False
    objWord.Documents.Open(fileIn)

SET objDoc = objWord.ActiveDocument
    objDoc.SaveAs fileOut, wdFormat
    objDoc.Close

    objWord.Quit

SET objDoc = Nothing
SET objWord = Nothing
