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
Sub ufOpen()

With uf_dev
  .StartUpPosition = 0
  .Left = Application.Left + (0.5 * Application.Width) - (0.5 * .Width)
  .Top = Application.Top + (0.5 * Application.Height) - (0.5 * .Height)
  .Show
End With

End Sub
Private Sub dev01()

Dim txt As String
Dim p As Paragraph
Dim sect As Long

With ThisDocument
  For Each p In .Paragraphs
    ' Start adding in sections: first find each by heading, then insert section
    If p.Style = "Heading 4" And Left(p.Range.Text, 7) = "SUBPART" Then
      p.Range.Select
      Selection.HomeKey
      Selection.InsertBreak Type:=wdSectionBreakNextPage
      
      ' Add the Subpart to each header
      ' Must start off at section 2, since first section is previous one after adding new one
      If sect = 0 Then
        sect = sect + 2
      Else
        sect = sect + 1
      End If
      
      ' Add the text
      txt = p.Range.Text
      With ThisDocument.Sections(sect).Headers(wdHeaderFooterPrimary)
        .LinkToPrevious = False
        .Range.Text = txt & vbCrLf
        
        ' Add in hyperlinks to the header
        Dim txt_ref, txt_far, txt_dfars, txt_nmcars, wh_far As String
        Dim pos, pos2, diff As Long
        
        ' Find position of second space to find the reference
        pos = InStr(9, txt, " ", vbTextCompare)
        txt_ref = Mid(txt, 9, pos - 9)
        
        ' Find the position of the . to see what FAR you are currently in
        pos2 = InStr(txt, ".")
        diff = pos2 - 8
        If diff = 2 Then
          wh_far = "FAR"
        ElseIf diff = 3 Then
          wh_far = "DFARS"
        ElseIf diff = 4 Then
          wh_far = "NMCARS"
        End If
        
        ' Rename references for hyperlinks
        ' Start with FAR
        If diff = 2 Then
          txt_far = txt_ref
          txt_dfars = "20" & txt_ref
          txt_nmcars = "520" & txt_ref
        
''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
' Stopped here - may need to convert DFARS and may need to see how the handbooks
'   handle citations
''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
        ' DFARS
        ElseIf diff = 3 Then
          
          txt_dfars = "2" & txt_ref
          txt_nmcars = "52" & txt_ref
          
        ' NCMARS
        ElseIf diff = 4 Then
          
          txt_dfars = "2" & txt_ref
          txt_nmcars = "52" & txt_ref
        End If
        txt_far = "SUBPART " & txt_far
        txt_dfars = "SUBPART " & txt_dfars
        txt_nmcars = "SUBPART " & txt_nmcars
        
        
        
        
        ' Peform hyperlinks for FAR
        If wh_far = "FAR" Then
          ' Rename references

        
          
        End If
        
      End With
      
    End If
  Next p
End With

End Sub
Private Sub dev02()

' WORKS!

ThisDocument.Sections(1).Headers(wdHeaderFooterPrimary).LinkToPrevious = False
ThisDocument.Sections(1).Headers(wdHeaderFooterPrimary).Range.Text = "Section 1"
ThisDocument.Sections(2).Headers(wdHeaderFooterPrimary).LinkToPrevious = False
ThisDocument.Sections(2).Headers(wdHeaderFooterPrimary).Range.Text = "Section 2"
ThisDocument.Sections(3).Headers(wdHeaderFooterPrimary).LinkToPrevious = False
ThisDocument.Sections(3).Headers(wdHeaderFooterPrimary).Range.Text = "Section 3"

End Sub
Private Sub dev03()

' Works!

Dim txt As String
Dim p As Paragraph
Dim cnt As Long

With ThisDocument
  For Each p In .Paragraphs
    If p.Style = "Heading 4" And Left(p.Range.Text, 7) = "SUBPART" Then
      cnt = cnt + 1
      'p.Range.Select
      'Selection.HomeKey
      'Selection.InsertBreak Type:=wdSectionBreakNextPage
    End If
  Next p
End With

Debug.Print cnt

End Sub
Private Sub dev04()

' Works! - text parsing
Dim txt, txt_ref, txt_dfars, txt_nmcars As String
Dim pos, pos2, diff As Long

txt = "SUBPART 1.000 - Ugh Derp"
pos = InStr(9, txt, " ", vbTextCompare)
pos2 = InStr(txt, ".")
txt_ref = Mid(txt, 9, pos - 9)
diff = pos2 - 8
If diff = 2 Then
  txt_dfars = "20" & txt_ref
  txt_nmcars = "520" & txt_ref
Else
  txt_dfars = "2" & txt_ref
  txt_nmcars = "52" & txt_ref
End If
txt_dfars = "SUBPART " & txt_dfars
txt_nmcars = "SUBPART " & txt_nmcars

Debug.Print pos2; " - "; txt_ref; " - "; diff; " - "; txt_dfars; " - "; txt_nmcars

'ActiveDocument.Hyperlinks.Add _
'  Anchor:=Selection.Range, _
'  Address:="", _
'  SubAddress:="p3", _
'  ScreenTip:="Ugh Derp!", _
'  TextToDisplay:="Page 2"

End Sub
