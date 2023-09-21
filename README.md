# MergeContentMacro
/* A Macro in MS Word that reads all Heading 1 in "ActiveDocument," then copies and pastes corresponding content from "MotherDocument."
I'm trying to develop this macro right now, but it keeps making MS Word crash. I'm not sure what I'm doing wrong, yet.
Before running the macro, you need two documents (MotherDocument, and ActiveDocument). MotherDocument is the repository for all your content. Content is organized under headings (system default Heading 1). In ActiveDocument, write out the headings you want to include, then run the macro. The macro should scan each of the headings, find the corresponding content in MotherDocument, then copy and paste that content into ActiveDocument. I'm pretty new to programming and got help from ChatGPT to get this far.
Here's the code I have so far: */

Sub MergeHeadings()
    Dim currentDoc As Document
    Dim motherDoc As Document
    Dim currentH1 As Paragraph
    Dim motherH1 As Paragraph
    
    ' Set the current document
    Set currentDoc = ActiveDocument
    
    ' Open the MotherDocument (update the path)
    Set motherDoc = Documents.Open("C:\Path\To\MotherDocument.docx")
    
    ' Loop through H1 headings in CurrentDocument
    For Each currentH1 In currentDoc.Paragraphs
        If currentH1.Style = "Heading 1" Then
            ' Search for a matching heading in MotherDocument
            For Each motherH1 In motherDoc.Paragraphs
                If motherH1.Style = "Heading 1" And motherH1.Range.Text = currentH1.Range.Text Then
                    ' Copy the matching heading and its contents from MotherDocument
                    motherH1.Range.Copy
                    ' Paste it to CurrentDocument under the current heading
                    currentH1.Range.Collapse wdCollapseEnd
                    currentH1.Range.Paste
                End If
            Next motherH1
        End If
    Next currentH1
    
    ' Close the MotherDocument without saving changes
    motherDoc.Close SaveChanges:=wdDoNotSaveChanges
End Sub
