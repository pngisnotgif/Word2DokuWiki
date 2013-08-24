'Please do not delete this section
'
'Developed by Tania Hew 07/2008
'
'


'Cancel Macro'
Private Sub CancelButton_Click()
YesLocation.Value = False
NoLocation.Value = False
Word2DokuWiki.Hide
End Sub

'Convert'
Private Sub ConvertButton_Click()
  If YesLocation.Value = False And NoLocation.Value = False Then
    MsgBox ("Please select whether or not to replace images by a specific image location")
  Else
    Dim FileName As String
    FileName = GetFilename(ActiveDocument.Name)
    
    Application.ScreenUpdating = False
    HideRevisions
    ReplaceQuotes
    DokuWikiEscapeChars
    
'    // 2011-06-20 by Taggic
    DokuWikiConvertFootnotes
    
    DokuWikiConvertHyperlinks
    DokuWikiConvertH1
    DokuWikiConvertH2
    DokuWikiConvertH3
    DokuWikiConvertH4
    DokuWikiConvertH5
    DokuWikiConvertItalic
    DokuWikiConvertBold
    DokuWikiConvertUnderline
    DokuWikiConvertStrikeThrough
    DokuWikiConvertSuperscript
    DokuWikiConvertSubscript
    DokuWikiConvertLists
    DokuWikiConvertTable
    UndoDokuWikiEscapeChars
    DokuWikiSaveAsHTMLAndConvertImages
    MoveJPGFilesToNewFolder
    MovePNGFilesToNewFolder
    MoveGIFFilesToNewFolder
    removeImages
    ActiveDocument.Content.Copy 'Copy to clipboard
    Application.ScreenUpdating = True
    AutoCopyToFile
    'ManualCopyToFile
    
    'CLEAN UP
    'DeleteHTMFile 'Remove HTM File'
    'DeleteHTMFolder 'Remove HTM Folder and contents'
    
    'Workaround to have original Word document open at end of conversion
'    ActiveDocument.Close
'    Application.Documents.Open (FileName)
    
    'Close Word to DokuWiki Converter dialog
    Word2DokuWiki.Hide
    
    MsgBox ("Word to DokuWiki Conversion complete!")
  End If
End Sub


Private Sub NoLocation_Click()
ImageLocation.Locked = True
ImageLocation.BackColor = &H8000000F
NoLabel.Visible = True
YesLabel.Visible = False
End Sub

Private Sub YesLocation_Click()
ImageLocation.Locked = False
ImageLocation.BackColor = &H80000005
YesLabel.Visible = True
NoLabel.Visible = False
End Sub

Private Sub HideRevisions()
    ActiveDocument.ShowRevisions = False
End Sub

Private Sub DokuWikiConvertH1()
    ReplaceHeading wdStyleHeading1, "======"
End Sub

Private Sub DokuWikiConvertH2()
    ReplaceHeading wdStyleHeading2, "====="
End Sub

Private Sub DokuWikiConvertH3()
    ReplaceHeading wdStyleHeading3, "===="
End Sub

Private Sub DokuWikiConvertH4()
        ReplaceHeading wdStyleHeading4, "==="
End Sub

Private Sub DokuWikiConvertH5()
    ReplaceHeading wdStyleHeading5, "=="
End Sub

Private Sub DokuWikiConvertH6()
    ReplaceHeading wdStyleHeading5, "="
End Sub

Private Sub DokuWikiConvertBold()
    ActiveDocument.Select
    With Selection.Find
        .ClearFormatting
        .Font.Bold = True
        .Text = ""
        .Format = True
        .MatchCase = False
        .MatchWholeWord = False
        .MatchWildcards = False
        .MatchSoundsLike = False
        .MatchAllWordForms = False
        .Forward = True
        .Wrap = wdFindContinue
       
        Do While .Execute
            With Selection
                If Len(.Text) > 1 And InStr(1, .Text, vbCr) Then
                    ' Just process the chunk before any newline characters
                    ' We'll pick-up the rest with the next search
                    .Collapse
                    .MoveEndUntil vbCr
                End If
                                      
                ' Don't bother to markup newline characters (prevents a loop, as well)
                
                If Not .Text = vbCr Then
                    If Not Left(.Text, 2) = "**" Then
                    .InsertBefore "**"
                    End If
                    If Not Right(.Text, 2) = "**" Then
                    .InsertAfter "**"
                    End If
                End If
               
                .Style = ActiveDocument.Styles("Standard")
                .Font.Bold = False
            End With
        Loop
    End With
End Sub
 
Private Sub DokuWikiConvertItalic()
    ActiveDocument.Select
   
    With Selection.Find
   
        .ClearFormatting
        .Font.Italic = True
        .Text = ""
       
        .Format = True
        .MatchCase = False
        .MatchWholeWord = False
        .MatchWildcards = False
        .MatchSoundsLike = False
        .MatchAllWordForms = False
       
        .Forward = True
        .Wrap = wdFindContinue
       
        Do While .Execute
            With Selection
                If Len(.Text) > 1 And InStr(1, .Text, vbCr) Then
                    ' Just process the chunk before any newline characters
                    ' We'll pick-up the rest with the next search
                    .Collapse
                    .MoveEndUntil vbCr
                End If
                                      
                ' Don't bother to markup newline characters (prevents a loop, as well)
                If Not .Text = vbCr Then
                    If Not Left(.Text, 2) = "//" Then
                    .InsertBefore "//"
                    End If
                    If Not Right(.Text, 2) = "//" Then
                    .InsertAfter "//"
                    End If
                End If
               
                .Style = ActiveDocument.Styles("Standard")
                .Font.Italic = False
            End With
        Loop
    End With
End Sub
 
Private Sub DokuWikiConvertUnderline()
    ActiveDocument.Select
   
    With Selection.Find
   
        .ClearFormatting
        .Font.Underline = True
        .Text = ""
       
        .Format = True
        .MatchCase = False
        .MatchWholeWord = False
        .MatchWildcards = False
        .MatchSoundsLike = False
        .MatchAllWordForms = False
       
        .Forward = True
        .Wrap = wdFindContinue
       
        Do While .Execute
            With Selection
                If Len(.Text) > 1 And InStr(1, .Text, vbCr) Then
                    ' Just process the chunk before any newline characters
                    ' We'll pick-up the rest with the next search
                    .Collapse
                    .MoveEndUntil vbCr
                End If
                                       
                ' Don't bother to markup newline characters (prevents a loop, as well)
                If Not .Text = vbCr Then
                    If Not Left(.Text, 2) = "__" Then
                    .InsertBefore "__"
                    End If
                    If Not Right(.Text, 2) = "__" Then
                    .InsertAfter "__"
                    End If
                End If
                
                .Style = ActiveDocument.Styles("Standard")
                .Font.Underline = False
            End With
        Loop
    End With
End Sub
 
Private Sub DokuWikiConvertStrikeThrough()
    ActiveDocument.Select
   
    With Selection.Find
   
        .ClearFormatting
        .Font.StrikeThrough = True
        .Text = ""
       
        .Format = True
        .MatchCase = False
        .MatchWholeWord = False
        .MatchWildcards = False
        .MatchSoundsLike = False
        .MatchAllWordForms = False
       
        .Forward = True
        .Wrap = wdFindContinue
       
        Do While .Execute
            With Selection
                If Len(.Text) > 1 And InStr(1, .Text, vbCr) Then
                    ' Just process the chunk before any newline characters
                    ' We'll pick-up the rest with the next search
                    .Collapse
                    .MoveEndUntil vbCr
                End If
                                      
                ' Don't bother to markup newline characters (prevents a loop, as well)
                If Not .Text = vbCr Then
                    If Not Left(.Text, 2) = "<del>" Then
                    .InsertBefore "<del>"
                    End If
                    If Not Right(.Text, 2) = "</del>" Then
                    .InsertAfter "</del>"
                    End If
                End If
               
                .Style = ActiveDocument.Styles("Standard")
                .Font.StrikeThrough = False
            End With
        Loop
    End With
End Sub
 
Private Sub DokuWikiConvertSuperscript()
    ActiveDocument.Select
   
    With Selection.Find
   
        .ClearFormatting
        .Font.Superscript = True
        .Text = ""
       
        .Format = True
        .MatchCase = False
        .MatchWholeWord = False
        .MatchWildcards = False
        .MatchSoundsLike = False
        .MatchAllWordForms = False
       
        .Forward = True
        .Wrap = wdFindContinue
       
        Do While .Execute
            With Selection
                .Text = Trim(.Text)
                If Len(.Text) > 1 And InStr(1, .Text, vbCr) Then
                    ' Just process the chunk before any newline characters
                    ' We'll pick-up the rest with the next search
                    .Collapse
                    .MoveEndUntil vbCr
                End If
                                       
                ' Don't bother to markup newline characters (prevents a loop, as well)
                If Not .Text = vbCr Then
                    If Not Left(.Text, 2) = "<sup>" Then
                    .InsertBefore "<sup>"
                    End If
                    If Not Right(.Text, 2) = "</sup>" Then
                    .InsertAfter "</sup>"
                    End If
                End If
                
                .Style = ActiveDocument.Styles("Standard")
                .Font.Superscript = False
            End With
        Loop
    End With
End Sub
 
Private Sub DokuWikiConvertSubscript()
    ActiveDocument.Select
   
    With Selection.Find
   
        .ClearFormatting
        .Font.Subscript = True
        .Text = ""
       
        .Format = True
        .MatchCase = False
        .MatchWholeWord = False
        .MatchWildcards = False
        .MatchSoundsLike = False
        .MatchAllWordForms = False
       
        .Forward = True
        .Wrap = wdFindContinue
       
        Do While .Execute
            With Selection
                .Text = Trim(.Text)
                If Len(.Text) > 1 And InStr(1, .Text, vbCr) Then
                    ' Just process the chunk before any newline characters
                    ' We'll pick-up the rest with the next search
                    .Collapse
                    .MoveEndUntil vbCr
                End If
                                       
                ' Don't bother to markup newline characters (prevents a loop, as well)
                If Not .Text = vbCr Then
                    If Not Left(.Text, 2) = "<sub>" Then
                    .InsertBefore "<sub>"
                    End If
                    If Not Right(.Text, 2) = "</sub>" Then
                    .InsertAfter "</sub>"
                    End If
                End If
               
                .Style = ActiveDocument.Styles("Standard")
                .Font.Subscript = False
            End With
        Loop
    End With
End Sub
 
Private Sub DokuWikiConvertLists()
    Dim para As Paragraph
    For Each para In ActiveDocument.ListParagraphs
        With para.Range
            .InsertBefore "  "
             If .ListFormat.ListType = wdListBullet Then
                 .InsertBefore "*"
             Else
                  .InsertBefore "-"
              End If
            For i = 1 To .ListFormat.ListLevelNumber
                   .InsertBefore "  "
           Next i
            .ListFormat.RemoveNumbers
        End With
    Next para
End Sub
 
'   // 2011-06-20 add by Taggic
Private Sub DokuWikiConvertFootnotes()
    Dim footnoteCount As Integer
    footnoteCount = ActiveDocument.Footnotes.Count
    For i = 1 To footnoteCount
        With ActiveDocument.Footnotes(1)
            Dim addr As String
            
            addr = .Range.Text
            
            .Reference.InsertAfter "((" & addr & "))"
            .Delete
        End With
    Next i
End Sub

 

Private Sub DokuWikiConvertHyperlinks()
    Dim hyperCount As Integer
   
    hyperCount = ActiveDocument.Hyperlinks.Count
   
    For i = 1 To hyperCount
        With ActiveDocument.Hyperlinks(1)
            Dim addr As String
            addr = .Address
            .Delete
            .Range.InsertBefore "["
            .Range.InsertAfter "-" & addr & "]"
        End With
    Next i
End Sub
 
' Replace all smart quotes with their dumb equivalents
Private Sub ReplaceQuotes()
    Dim quotes As Boolean
    quotes = Options.AutoFormatAsYouTypeReplaceQuotes
    Options.AutoFormatAsYouTypeReplaceQuotes = False
    ReplaceString ChrW(8220), """"
    ReplaceString ChrW(8221), """"
    ReplaceString "?, " '"
    ReplaceString "?, " '"
    Options.AutoFormatAsYouTypeReplaceQuotes = quotes
End Sub
 
Private Sub DokuWikiEscapeChars()
    EscapeCharacter "*"
    EscapeCharacter "#"
    EscapeCharacter "_"
    EscapeCharacter "-"
    EscapeCharacter "+"
    EscapeCharacter "{"
    EscapeCharacter "}"
    EscapeCharacter "["
    EscapeCharacter "]"
    EscapeCharacter "~"
    EscapeCharacter "^^"
    EscapeCharacter "|"
    EscapeCharacter "'"
End Sub
 
Private Function ReplaceHeading(styleHeading As String, headerPrefix As String)
    Dim normalStyle As Style
    Set normalStyle = ActiveDocument.Styles(wdStyleNormal)
   
    ActiveDocument.Select
   
    With Selection.Find
   
        .ClearFormatting
        .Style = ActiveDocument.Styles(styleHeading)
        .Text = ""

      
        .Format = True
        .MatchCase = False
        .MatchWholeWord = False
        .MatchWildcards = False
        .MatchSoundsLike = False
        .MatchAllWordForms = False
       
        .Forward = True
        .Wrap = wdFindContinue
       
        Do While .Execute
            With Selection
                If InStr(1, .Text, vbCr) Then
                    ' Just process the chunk before any newline characters
                    ' We'll pick-up the rest with the next search
                    .Collapse
                    .MoveEndUntil vbCr
                End If
                                       
                ' Don't bother to markup newline characters (prevents a loop, as well)
               If Not .Text = vbCr Then
                   .InsertBefore headerPrefix
                   .InsertBefore vbCr
                   .InsertAfter headerPrefix
               End If
               .Style = normalStyle
           End With
       Loop
   End With
End Function

Private Sub DokuWikiConvertTable()
  Dim TotTables As Long
  Do While ActiveDocument.Tables.Count() > 0
    ActiveDocument.Tables(1).Range.Select
    Selection.Find.ClearFormatting
    Selection.Find.Replacement.ClearFormatting
    With Selection.Find
      .Text = " $s$|$s$ "
      .Replacement.Text = "I"
      .Forward = True
      .Wrap = wdFindContinue
      .Format = False
      .MatchCase = False
      .MatchWholeWord = False
      .MatchWildcards = False
      .MatchSoundsLike = False
      .MatchAllWordForms = False
    End With
    Selection.Find.Execute Replace:=wdReplaceAll
    Selection.Find.ClearFormatting
    Selection.Find.Replacement.ClearFormatting
    With Selection.Find
      .Text = " $s$^^$s$ "
      .Replacement.Text = "/\"
      .Forward = True
      .Wrap = wdFindContinue
      .Format = False
      .MatchCase = False
      .MatchWholeWord = False
      .MatchWildcards = False
      .MatchSoundsLike = False
      .MatchAllWordForms = False
    End With
    Selection.Find.Execute Replace:=wdReplaceAll
    Selection.Find.ClearFormatting
    Application.DefaultTableSeparator = "|"
    Selection.Rows.ConvertToText Separator:=wdSeparateByDefaultListSeparator, NestedTables:=True
    Selection.Find.ClearFormatting
    Selection.Find.Replacement.ClearFormatting
    With Selection.Find
      .Text = "^p"
      .Replacement.Text = "|^p|"
      .Forward = True
      .Wrap = wdFindStop
      .Format = False
      .MatchCase = False
      .MatchWholeWord = False
      .MatchWildcards = False
      .MatchSoundsLike = False
      .MatchAllWordForms = False
    End With
    Selection.Find.Execute Replace:=wdReplaceAll
    Selection.InsertBefore ("|")
    Selection.InsertParagraphAfter
    Selection.Find.ClearFormatting
    Selection.Find.Replacement.ClearFormatting
    With Selection.Find
      .Text = "^p|^p"
      .Replacement.Text = "^p"
      .Forward = True
      .Wrap = wdFindStop
      .Format = False
      .MatchCase = False
      .MatchWholeWord = False
      .MatchWildcards = False
      .MatchSoundsLike = False
      .MatchAllWordForms = False
    End With
    Selection.Find.Execute Replace:=wdReplaceAll
    Selection.Find.ClearFormatting
    Selection.Find.Replacement.ClearFormatting
    With Selection.Find
      .Text = "$s$blank$s$"
      .Replacement.Text = ""
      .Forward = True
      .Wrap = wdFindContinue
      .Format = False
      .MatchCase = False
      .MatchWholeWord = False
      .MatchWildcards = False
      .MatchSoundsLike = False
      .MatchAllWordForms = False
    End With
    Selection.Find.Execute Replace:=wdReplaceAll
    Selection.Find.ClearFormatting
    Selection.Find.Replacement.ClearFormatting
    With Selection.Find
      .Text = "||"
      .Replacement.Text = "|  |"
      .Forward = True
      .Wrap = wdFindStop
      .Format = False
      .MatchCase = False
      .MatchWholeWord = False
      .MatchWildcards = False
      .MatchSoundsLike = False
      .MatchAllWordForms = False
    End With
    Selection.Find.Execute Replace:=wdReplaceAll
    With Selection.Find
      .Text = "||"
      .Replacement.Text = "|  |"
      .Forward = True
      .Wrap = wdFindStop
      .Format = False
      .MatchCase = False
      .MatchWholeWord = False
      .MatchWildcards = False
      .MatchSoundsLike = False
      .MatchAllWordForms = False
    End With
    Selection.Find.Execute Replace:=wdReplaceAll
    Selection.Find.ClearFormatting
    Selection.Find.Replacement.ClearFormatting
    With Selection.Find
      .Text = "| |"
      .Replacement.Text = "|  |"
      .Forward = True
      .Wrap = wdFindStop
      .Format = False
      .MatchCase = False
      .MatchWholeWord = False
      .MatchWildcards = False
      .MatchSoundsLike = False
      .MatchAllWordForms = False
    End With
    Selection.Find.Execute Replace:=wdReplaceAll
    With Selection.Find
      .Text = "| |"
      .Replacement.Text = "|  |"
      .Forward = True
      .Wrap = wdFindStop
      .Format = False
      .MatchCase = False
      .MatchWholeWord = False
      .MatchWildcards = False
      .MatchSoundsLike = False
      .MatchAllWordForms = False
    End With
    Selection.Find.Execute Replace:=wdReplaceAll
    Selection.Paragraphs(1).Range.Select
    Selection.Find.ClearFormatting
    Selection.Find.Replacement.ClearFormatting
    With Selection.Find
      .Text = "|"
      .Replacement.Text = "^^"
      .Forward = True
      .Wrap = wdFindStop
      .Format = False
      .MatchCase = False
      .MatchWholeWord = False
      .MatchWildcards = False
      .MatchSoundsLike = False
      .MatchAllWordForms = False
    End With
    Selection.Find.Execute Replace:=wdReplaceAll
  Loop
End Sub

Private Sub UndoDokuWikiEscapeChars()
    UndoEscapeCharacter "*"
    UndoEscapeCharacter "#"
    UndoEscapeCharacter "_"
    UndoEscapeCharacter "-"
    UndoEscapeCharacter "+"
    UndoEscapeCharacter "{"
    UndoEscapeCharacter "}"
    UndoEscapeCharacter "["
    UndoEscapeCharacter "]"
    UndoEscapeCharacter "~"
    UndoEscapeCharacter "^^"
    UndoEscapeCharacter "|"
    UndoEscapeCharacter "'"
End Sub

Private Function EscapeCharacter(char As String)
    ReplaceString char, " $s$" & char & "$s$ "
End Function

Private Function UndoEscapeCharacter(char As String)
    ReplaceString " $s$" & char & "$s$ ", char
End Function

Private Function ReplaceString(findStr As String, replacementStr As String)
    Selection.Find.ClearFormatting
    Selection.Find.Replacement.ClearFormatting
    With Selection.Find
        .Text = findStr
        .Replacement.Text = replacementStr
        .Forward = True
        .Wrap = wdFindContinue
        .Format = False
        .MatchCase = False
        .MatchWholeWord = False
        .MatchWildcards = False
        .MatchSoundsLike = False
        .MatchAllWordForms = False
    End With
    Selection.Find.Execute Replace:=wdReplaceAll
End Function





'begin my functions'

'function to get file name path of document - full path plus file name minus file extension
'Example if file is C:\Documents & Settings\taniah\My Documents\docname.doc, this would
'return C:\Documents & Settings\taniah\My Documents\docname'
Private Function GetFilename(ByVal strPath As String) As String
    GetFilename = ActiveDocument.Path & "\" & ActiveDocument.Name
  
    'Strip the .doc from the end
    GetFilename = Left(GetFilename, Len(GetFilename) - 4)
End Function

'function to get file name only of text document minus extension
'Example if file is C:\Documents & Settings\taniah\My Documents\docname.doc, this would
'return docname'
Private Function GetFilenameOnly(ByVal strPath As String) As String
    Dim lngPos As Long
    Dim fName As String
 
    If (Left$(strPath, 4) <> "*.txt") And (Len(strPath) > 0) Then
      On Error GoTo LocalHandler
        'Get all characters up to .txt string
        lngPos = InStr(strPath, ".txt")
        GetFilenameOnly = Left$(strPath, lngPos - 1)
      Else
LocalHandler:
        'Return error
        MsgBox ("There was an error retrieving file name. Please ensure that current file is a text document")
        'Application.Quit
    End If
End Function

Private Sub DokuWikiSaveAsHTMLAndConvertImages()
    Dim s As Shape
    Dim FileLocation As String

    For Each s In ActiveDocument.Shapes
        s.ConvertToInlineShape
    Next

    FileLocation = ActiveDocument.Path + "\" + ActiveDocument.Name
    FileName = GetFilename(ActiveDocument.Name)
    FolderName = FileName + "_files"

    ActiveDocument.SaveAs FileName:=FileName + ".htm", _
                  FileFormat:=wdFormatFilteredHTML, LockComments:=False, Password:="", _
                  AddToRecentFiles:=True, WritePassword:="", ReadOnlyRecommended:=False, _
                  EmbedTrueTypeFonts:=False, SaveNativePictureFormat:=False, SaveFormsData _
                  :=False, SaveAsAOCELetter:=False

    'Rename all the files with a Unique name
    'strDir = Dir(FileName & "_files\*.jpg")


    'Ask for image location on wiki'
    Dim iShape As InlineShape

    If YesLocation.Value = True Then
    sLocation = ImageLocation.Text
             
                'Put image location link in DokuWiki format for all images produced in text file'
                Set fs = CreateObject("Scripting.FileSystemObject")
                If fs.FolderExists(FolderName) Then
                    Set f = fs.GetFolder(FolderName)
            
                    Set fc = f.Files
                    i = 1
                    For Each f In fc
                        If i <= ActiveDocument.InlineShapes.Count Then
                            Set iShape = ActiveDocument.InlineShapes.Item(i)
                            iShape.Range.InsertBefore "{{" + sLocation + ":" + f.Name & "|}}"
                            i = i + 1
                        End If
                    Next
               End If

      ElseIf NoLocation.Value = True Then
            'Go through every image that has been produced and substitute the link in the DokuWiki page
            Set fs = CreateObject("Scripting.FileSystemObject")
            If fs.FolderExists(FolderName) Then
                Set f = fs.GetFolder(FolderName)
        
                Set fc = f.Files
                i = 1
                For Each f In fc
                    If i <= ActiveDocument.InlineShapes.Count Then
                        Set iShape = ActiveDocument.InlineShapes.Item(i)
                        iShape.Range.InsertBefore "!IMAGE: " + f.Name & " :IMAGE!"
                        i = i + 1
                    End If
                Next
              End If
      Else
          'If Cancel was chosen, do nothing.
          'Shell "explorer.exe " + FileName + "_files", vbNormalFocus

    End If
    'MsgBox ("HTML creation done")
End Sub

'function to move jpg files from one folder to a newly created folder
Private Sub MoveJPGFilesToNewFolder()
    Dim FSO As Object
    Dim FromPath As String
    Dim ToPath As String
    Dim FileExt As String
    Dim FNames As String
    
    FileName = GetFilename(ActiveDocument.Name)
    FolderName = FileName + "_files"
    
    FromPath = FolderName
    ToPath = FileName + " IMAGES"

    FileExt = "*.jpg*"
    

    If Right(FromPath, 1) <> "\" Then
        FromPath = FromPath & "\"
    End If

    JPGFNames = Dir(FromPath & FileExt)

    If (Len(JPGFNames) = 0) Then
        'MsgBox "No files in " & FromPath
        Exit Sub
    End If

    Set FSO = CreateObject("scripting.filesystemobject")

    If FSO.FolderExists(ToPath) = False Then
        FSO.CreateFolder (ToPath)
    End If

    FSO.MoveFile Source:=FromPath & FileExt, Destination:=ToPath
    MsgBox "You can find the image files associated with the created wiki page here: " & ToPath
End Sub

'function to move png files from one folder to a newly created folder
Private Sub MovePNGFilesToNewFolder()
    Dim FSO As Object
    Dim FromPath As String
    Dim ToPath As String
    Dim FileExt As String
    Dim FNames As String
    
    FileName = GetFilename(ActiveDocument.Name)
    FolderName = FileName + "_files"
    
    FromPath = FolderName
    ToPath = FileName + " IMAGES"

    FileExt = "*.png*"
    

    If Right(FromPath, 1) <> "\" Then
        FromPath = FromPath & "\"
    End If

    PNGFNames = Dir(FromPath & FileExt)

    If (Len(PNGFNames) = 0) Then
        'MsgBox "No files in " & FromPath
        Exit Sub
    End If

    Set FSO = CreateObject("scripting.filesystemobject")

    If FSO.FolderExists(ToPath) = False Then
        FSO.CreateFolder (ToPath)
    End If

    FSO.MoveFile Source:=FromPath & FileExt, Destination:=ToPath
    MsgBox "You can find the image files associated with the created wiki page here: " & ToPath
End Sub

'function to move gif files from one folder to a newly created folder
Private Sub MoveGIFFilesToNewFolder()
    Dim FSO As Object
    Dim FromPath As String
    Dim ToPath As String
    Dim FileExt As String
    Dim FNames As String
    
    FileName = GetFilename(ActiveDocument.Name)
    FolderName = FileName + "_files"
    
    FromPath = FolderName
    ToPath = FileName + " IMAGES"

    FileExt = "*.gif*"
    

    If Right(FromPath, 1) <> "\" Then
        FromPath = FromPath & "\"
    End If

    GIFFNames = Dir(FromPath & FileExt)

    If (Len(GIFFNames) = 0) Then
        'MsgBox "No files in " & FromPath
        Exit Sub
    End If

    Set FSO = CreateObject("scripting.filesystemobject")

    If FSO.FolderExists(ToPath) = False Then
        FSO.CreateFolder (ToPath)
    End If

    FSO.MoveFile Source:=FromPath & FileExt, Destination:=ToPath
    MsgBox "You can find the image files associated with the created wiki page here: " & ToPath
End Sub

'Function to delete the HTM file created
Private Function DeleteHTMFile()
  Dim HTMFile As String
  
  'HTM File should have same name as current document minus extension
  HTMFile = GetFilenameOnly(ActiveDocument.Name) + ".htm"
  
  'Look for specified file
  For Each HTMLDoc In Application.Documents
  If HTMLDoc.Name = HTMFile Then
    wdDoc.Close
  End If
  Next HTMLDoc
  
  'Delete File
  Set objFSO = CreateObject("Scripting.FileSystemObject")
  objFSO.deletefile (ActiveDocument.Path + "\" + HTMFile), True
End Function

'function to Delete HTML folder that is automatically created when html file is created
Private Sub DeleteHTMFolder()
  Dim FSO As Object
   Dim FolderName As String
   Dim FileName As String
    
    Set FSO = CreateObject("scripting.filesystemobject")

    FileName = GetFilename(ActiveDocument.Name)
    FolderName = FileName + "_files"

    If Right(FolderName, 1) = "\" Then
        FolderName = Left(FolderName, Len(FolderName) - 1)
    End If

    If FSO.FolderExists(FolderName) = False Then
        'MsgBox FolderName & " doesn't exist"
        'there were no images found in source document
        MsgBox "No images were found in source document."
        Exit Sub
    End If

    On Error GoTo 0
    'Delete files
    FSO.deletefile FolderName & "\*.*", True
    'Delete subfolders
    FSO.deletefolder FolderName & "\*.*", True
    
    Dir "C:\" 'This line added so that folder can be deleted without error'
    'Delete folder
    FSO.deletefolder FolderName, True
    On Error GoTo 0
    
    
End Sub

'Function to remove images from a document'
Private Sub removeImages()
' enregistrée le 23/10/2006 par OLIVIER

  Selection.Find.ClearFormatting
  Selection.Find.Replacement.ClearFormatting
  With Selection.Find
    .Text = "^g"
    .Replacement.Text = ""
    .Forward = True
    .Wrap = wdFindContinue
    .Format = False
    .MatchCase = False
    .MatchWholeWord = False
    .MatchWildcards = False
    .MatchSoundsLike = False
    .MatchAllWordForms = False
  End With
  Selection.Find.Execute Replace:=wdReplaceAll
End Sub

'Create Dokuwiki version of page using user-specified file name to store contents'
'Function not used
Private Sub ManualCopyToFile()
   Dim sTemp As String

   'retrieve clipboard text content
   sTemp = ActiveDocument.Content

  'Open File Save As Dialog'
   With Application.Dialogs(wdDialogFileSaveAs)
    .Name = "*.txt"
    .Show
    If Err Then
       'This code runs if the dialog was cancelled
      MsgBox "Dialog Cancelled"
      Exit Sub
    End If
  End With
End Sub


'automatically create text file with DokuWiki syntax of document content
Private Sub AutoCopyToFile()

   Dim sTemp As String
   Dim fullDocName As String
   Dim docName As String
      
   'retrieve clipboard text content
   sTemp = ActiveDocument.Content
   
   'get full document name
   fullDocName = ActiveDocument.Name
   

   'get filename excluding file extension
   docName = GetFilename(fullDocName) + ".txt"
   
   'THIS SECTION CAN REPLACE SECTION BELOW IF WANT TO BE GIVEN OPTION
   'TO NOT OVERWRITE FILES
   'Dim strMsg As String
   'strMsg = "A file called " & docName & " already exists. Do you want to replace the existing " & strSaveAsName & "?"
   
   ' Check if the file already exists
   'If Dir(docName & "*") = "" Then
    'If file does not exist, save without prompting.
    'save clipboard content to a text file having same name as Word document
    'ActiveDocument.SaveAs FileName:=docName, FileFormat:=wdFormatText, FileFormat:=wdFormatText, _
    'LockComments:=False, Password:="", AddToRecentFiles:=True, WritePassword _
    ':="", ReadOnlyRecommended:=False, EmbedTrueTypeFonts:=False, _
    'SaveNativePictureFormat:=False, SaveFormsData:=False, SaveAsAOCELetter:= _
    'False
   'Else
      ' If file does exist, prompt with warning message.
      ' Check value of button clicked in message box.
      'Select Case MsgBox(strMsg, vbYesNoCancel + vbExclamation)
         'Case vbYes
         ' If Yes was chosen, save and overwrite existing file.
            'On Error GoTo LocalHandler
    'save clipboard content to a text file having same name as Word document
    'ActiveDocument.SaveAs FileName:=docName, FileFormat:=wdFormatText, FileFormat:=wdFormatText, _
    'LockComments:=False, Password:="", AddToRecentFiles:=True, WritePassword _
    ':="", ReadOnlyRecommended:=False, EmbedTrueTypeFonts:=False, _
    'SaveNativePictureFormat:=False, SaveFormsData:=False, SaveAsAOCELetter:= _
    'False
         'Case vbNo
         ' If No was chosen, prompt for file name
         ' using the File SaveAs dialog box.
            'With Dialogs(wdDialogFileSaveAs)
               '.Name = "*.txt"
               '.Show
            'End With
         'Case Else
         ' If Cancel was chosen, do nothing.
      'End Select
   'End If
   
   'Need to add code below to rename images folder so name is based on the user-provided text file name
   'END THIS SECTION
   
    On Error GoTo LocalHandler
    MsgBox ("Any existing text files will be overwritten.")
    'save clipboard content to a text file having same name as Word document
    ActiveDocument.SaveAs FileName:=docName, FileFormat:=wdFormatText, _
    LockComments:=False, Password:="", AddToRecentFiles:=True, WritePassword _
    :="", ReadOnlyRecommended:=False, EmbedTrueTypeFonts:=False, _
    SaveNativePictureFormat:=False, SaveFormsData:=False, SaveAsAOCELetter:= _
    False
    
LocalHandler:
        'MsgBox ("There was an error saving the text file. ")
        'Application.Quit

End Sub

