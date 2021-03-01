Attribute VB_Name = "main"
Option Explicit
Sub findHeading4()

' Only print headings
Dim para As Paragraph
For Each para In ActiveDocument.Paragraphs
  If para.Range.Style = "Heading 4" And Left(para.Range.Text, 7) = "SUBPART" Then
    Debug.Print para.Range.Text
  End If
Next

End Sub
Private Sub dev01()

' Works! - appends results to text file
' Only print headings
Dim file_num As Long
Dim append_text As String
Dim p As Paragraph

file_num = FreeFile
Open "C:\Users\smats\Documents\office\word\far\dfars_dev.txt" For Append As #file_num

For Each p In ThisDocument.Paragraphs
  If p.Range.Style = "Heading 4" And Left(p.Range.Text, 7) = "SUBPART" Then
    Print #file_num, p.Range.Text
  End If
Next

Close #file_num

End Sub
Private Sub dev02()

' Works! - finds headings and checks to see if they even exist
Dim txt As String
Dim pos As Long
txt = "SUBPART 201.1 —PURPOSE, AUTHORITY, ISSUANCE"
pos = InStr(9, txt, " ", vbTextCompare)
txt = Left(txt, pos)

With ThisDocument.Range
  With .Find
    .ClearFormatting
    .Format = True
    .Forward = True
    .Wrap = wdFindContinue
    .MatchWholeWord = False
    .Style = "Heading 4"
    .Text = txt
    .Execute
  End With
  If .Find.Found = True Then
    .Select
  End If
End With

End Sub
