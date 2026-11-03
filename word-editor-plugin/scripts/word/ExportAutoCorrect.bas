Attribute VB_Name = "ExportAutoCorrect"
Option Explicit

' Exports Word's AutoCorrect list to a UTF-8 TSV file in your Downloads folder.
' One entry per line: <replace> TAB <with>. Windows only (uses ADODB.Stream).
' Word keeps a separate list per language; run it with the cursor in text
' of the language whose list you want.
Public Sub ExportAutoCorrectToTsv()
    Dim stm As Object
    Dim entry As AutoCorrectEntry
    Dim path As String
    Dim count As Long, skipped As Long

    path = Environ("USERPROFILE") & "\Downloads\autocorrect-word.tsv"

    Set stm = CreateObject("ADODB.Stream")
    stm.Type = 2              ' text
    stm.Charset = "utf-8"
    stm.Open

    For Each entry In Application.AutoCorrect.Entries
        ' Formatted (rich text) entries cannot be represented as plain text.
        If entry.RichText Or InStr(entry.Value, vbTab) > 0 Or InStr(entry.Value, vbCr) > 0 Or InStr(entry.Value, vbLf) > 0 Then
            skipped = skipped + 1
        Else
            stm.WriteText entry.Name & vbTab & entry.Value & vbLf
            count = count + 1
        End If
    Next entry

    stm.SaveToFile path, 2    ' overwrite if it exists
    stm.Close

    MsgBox "Exported " & count & " entries (" & skipped & " skipped) to:" & vbCrLf & path, vbInformation
End Sub
