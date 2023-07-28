Sub ExportComments()
    ' Exports comments from a MS Word document to Excel and associates them with the heading paragraphs
    ' they are included in. Useful for outline numbered section, i.e. 3.2.1.5....

    Dim xlApp As Excel.Application
    Dim xlWB As Excel.Workbook
    Dim i As Integer, HeadingRow As Integer
    Dim objPara As Paragraph
    Dim objComment As comment
    Dim strSection As String
    Dim strTemp
    Dim myRange As Range

    On Error Resume Next ' Enable error handling

    ' Create Excel Application and Workbook
    Set xlApp = CreateObject("Excel.Application")
    xlApp.Visible = True
    Set xlWB = xlApp.Workbooks.Add 'create a new workbook

    ' Set the heading row in Excel
    HeadingRow = 1
    With xlWB.Worksheets(1)
        .Cells(HeadingRow, 1).Formula = "Comment ID"
        .Cells(HeadingRow, 2).Formula = "Page"
        .Cells(HeadingRow, 3).Formula = "Paragraph"
        .Cells(HeadingRow, 5).Formula = "Thread"
        .Cells(HeadingRow, 4).Formula = "Comment"
        .Cells(HeadingRow, 6).Formula = "Reviewer"
        .Cells(HeadingRow, 7).Formula = "Ministry"
        .Cells(HeadingRow, 8).Formula = "Date"
        .Cells(HeadingRow, 9).Formula = "Acceptance"
    End With

    ' Find the starting page number of the document
    Dim docStartPage As Long
    docStartPage = GetDocumentStartPage(ActiveDocument)

    ' Loop through each comment in the document
    For i = 1 To ActiveDocument.Comments.Count
        Set myRange = ActiveDocument.Comments(i).Scope
        strSection = ParentLevel(myRange.Paragraphs(1)) ' find the section heading for this comment

        With xlWB.Worksheets(1)
            .Cells(i + HeadingRow, 1).Formula = ActiveDocument.Comments(i).Index
            .Cells(i + HeadingRow, 2).Value = GetCommentPageNumber(ActiveDocument.Comments(i), docStartPage) ' Get the page number
            .Cells(i + HeadingRow, 3).Value = strSection
            .Cells(i + HeadingRow, 5).Value = GetCommentReplyTo(ActiveDocument.Comments(i)) ' Get the comment it was replying to
            .Cells(i + HeadingRow, 4).Formula = ActiveDocument.Comments(i).Range

            Dim authorName As String
            Dim ministryCode As String ' New variable to hold the Ministry code

            ' Get the author name and extract the Ministry code
            ExtractAuthorAndMinistryCode ActiveDocument.Comments(i).author, authorName, ministryCode

            .Cells(i + HeadingRow, 6).Value = authorName
            
            ' Remove ":EX" or ":IN" from the Ministry code
            ministryCode = RemoveEXINFromMinistryCode(ministryCode)

            .Cells(i + HeadingRow, 7).Value = ministryCode ' Populate the Ministry column

            .Cells(i + HeadingRow, 8).Value = Format(ActiveDocument.Comments(i).Date, "DD-MM-YYYY") ' Format the date
            .Cells(i + HeadingRow, 9).Formula = ActiveDocument.Comments(i).Done

            ' Check if the paragraph is "Not a reply" and apply conditional formatting
            If .Cells(i + HeadingRow, 5).Value = "New Thread" Then
                .Rows(i + HeadingRow).Interior.Color = RGB(191, 225, 255) ' Light blue background
            End If
        End With
    Next i

    ' Save the Excel file with the desired name after user confirmation
    Dim fileName As String
    fileName = GetSourceFileName(ActiveDocument) & "_comment_output_" & Format(Now(), "YYYY-MM-DD") & ".xlsx"
    
    ' Ask the user if they want to save the Excel file
    Dim saveChoice As String
    saveChoice = UCase(InputBox("Do you want to save the exported comments to Excel?" & vbCrLf & "Type 'YES' to save or 'NO' to cancel.", , "YES"))


    If saveChoice = "YES" Then
        xlWB.SaveAs fileName
        MsgBox "Comments exported and saved to: " & fileName, vbInformation, "Export Successful"
    Else
        MsgBox "Export canceled. Comments were not saved.", vbExclamation, "Export Canceled"
    End If
    
    
    Set xlWB = Nothing
    Set xlApp = Nothing
End Sub

Function ParentLevel(ByVal para As Word.Paragraph) As String
    ' Finds the first outlined numbered paragraph above the given paragraph object

    On Error Resume Next ' Enable error handling

    Dim ParaAbove As Word.Paragraph
    Set ParaAbove = para
    sStyle = para.Range.ParagraphStyle
    sStyle = Left(sStyle, 4)
    If sStyle = "Head" Then
        GoTo Skip
    End If
    Do While ParaAbove.OutlineLevel = para.OutlineLevel
        If ParaAbove.Previous Is Nothing Then
            Exit Do
        End If
        Set ParaAbove = ParaAbove.Previous
    Loop

Skip:
    strTitle = ParaAbove.Range.Text
    strTitle = Left(strTitle, Len(strTitle) - 1)
    ParentLevel = ParaAbove.Range.ListFormat.ListString & " " & strTitle
End Function

Function GetCommentReplyTo(comment As comment) As String
    ' Get the text of the comment that the input comment is replying to

    On Error Resume Next ' Enable error handling

    Dim replyComment As comment
    Dim para As Paragraph
    Dim commentRange As Range
    Dim commentText As String

    Set para = comment.Scope.Paragraphs(1)
    Set commentRange = para.Range
    commentText = commentRange.Text

    For Each replyComment In ActiveDocument.Comments
        If replyComment.Index < comment.Index Then
            Set para = replyComment.Scope.Paragraphs(1)
            Set commentRange = para.Range
            If InStr(commentRange.Text, commentText) > 0 Then
                GetCommentReplyTo = replyComment.Range.Text
                Exit Function
            End If
        End If
    Next replyComment

    GetCommentReplyTo = "New Thread"
End Function

Function GetCommentPageNumber(comment As comment, docStartPage As Long) As Long
    ' Get the adjusted page number for the comment

    On Error Resume Next ' Enable error handling

    Dim commentRange As Range
    Set commentRange = comment.Scope

    ' Move the range to the start of the comment
    commentRange.Collapse wdCollapseStart

    ' Find the adjusted page number for the comment
    Dim commentPage As Long
    commentPage = commentRange.Information(wdActiveEndAdjustedPageNumber)

    ' Calculate the adjusted page number
    GetCommentPageNumber = docStartPage + commentPage - 1
End Function

Function GetSourceFileName(doc As Document) As String
    ' Get the source file name without extension

    Dim fileName As String
    fileName = doc.FullName
    GetSourceFileName = Left(fileName, InStrRev(fileName, ".") - 1)
End Function

Function GetDocumentStartPage(doc As Document) As Long
    ' Get the starting page number of the document

    On Error Resume Next ' Enable error handling

    Dim docRange As Range
    Set docRange = doc.Range

    ' Find the starting page number of the document
    Dim startPage As Long
    startPage = docRange.Information(wdActiveEndAdjustedPageNumber)

    GetDocumentStartPage = startPage
End Function

Function ExtractAuthorAndMinistryCode(ByVal author As String, ByRef cleanedAuthor As String, ByRef ministryCode As String)
    ' Extract the author's name and Ministry code from the given author string
    cleanedAuthor = author
    ministryCode = ""

    ' Look for a space-separated code at the end of the author's name
    Dim spacePosition As Integer
    spacePosition = InStrRev(author, " ")

    If spacePosition > 0 Then
        cleanedAuthor = Trim(Left(author, spacePosition - 1))
        ministryCode = Trim(Mid(author, spacePosition + 1))
    End If
End Function

Function RemoveEXINFromMinistryCode(ByVal code As String) As String
    ' Remove ":EX" or ":IN" from the Ministry code
    RemoveEXINFromMinistryCode = Replace(code, ":EX", "")
    RemoveEXINFromMinistryCode = Replace(RemoveEXINFromMinistryCode, ":IN", "")
End Function