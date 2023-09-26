'App.config in Windows Forms app trips up with some exception on first call to ConfigurationManager.AppSettings
'with exception details, If I recall correctly, saying that <system.diagnostics> Is not recognized and/or initialized.
'Removing <system.diagnostics> element And child elements solves the problem! Don't know if that's the right thing to do.
Imports System.Configuration

Imports System.IO
Imports System.Xml
Module MainBlogExportFileToBook
    ' The (free reuse) license for this code created by me, Ravi S. Iyer, is provided in my blog post:
    ' All my blog data and books publicly accessible on Google Drive; Permission for free reuse,
    ' https://ravisiyer.blogspot.com/p/all-my-blogbooks-publicly-accessible-on.html .
    '
    ' I (Ravi S. Iyer) am now an obsolete software developer as till end June 2023, I had stayed away from coding
    ' for around, if not over, a decade now, barring very tiny JavaScript tweaking of others' code to suit my needs.
    ' The initial version of this software was done in Visual Basic for Applications (VBA). Now I am trying to port it
    ' to VB.Net. I am out of touch with .Net programming which I had done many years ago mainly in C#.
    ' Prior to this exploration (VBA project) I had very limited exposure to VBA - so the VBA development platform
    ' as well as the
    ' development environment is new to me. I want to limit the time I spend on this work. So I am looking at
    ' fast code development without focusing much on niceties which may be bad coding style.
    ' I also may not be calling the right VBA functions or calling them the right way - I just don't have the time
    ' to read up on the development platform / VBA barring just quick viewing of reference pages and help I get
    ' from Google Search result links, to figure out what I should try.
    '
    ' I am focusing on code for my specific needs and not for any general purpose needs. So my code may
    ' have bugs and problems which do not come into play when I am using it for my specific needs but come into
    ' play when others use it for different needs. I am not in a position now to help fix such issues. Other
    ' developers are absolutely welcome to do such fixes or other changes and re-publish their version of this
    ' software.

    Public LastXMLInputFilePath As String

    Public Const MsgBoxTitle As String = "Blog Backup XML file to HTML Blogbook"

    Public Const SelectFeedEntryXMLTerm = "atom:entry"
    Public Const CategoryXMLName = "atom:category"
    Public Const CategoryTermAttributeXMLName = "term"
    Public Const CategoryTermKindPostXMLName = "post"
    Public Const CategoryTermKindPageXMLName = "page"
    Public Const CategoryTermKindCommentXMLName = "comment"
    Public Const CategoryTermKindSettingsXMLName = "settings"
    Public Const CategoryTermKindTemplateXMLName = "template"

    Public LogFileStream As System.IO.StreamWriter

    Public RunBEFToBookProgram As Boolean

    Structure PPCEntryType ' PPC stands for Page or Post or Comment
        Public Index As Integer
        Public CategoryTermKind As String
        Public Title As String
        Public Content As String
        Public href As String
        Public AuthorName As String
        Public PublishedDate As String
        Public UpdatedDate As String
    End Structure

    Structure BlogDetailsType
        Public UpdatedDate As String
        Public Title As String
        Public href As String  ' Alternate href among multiple link elements
        Public AuthorName As String
    End Structure

    Sub DriveDefaultBlogExportFileSearchToBook()

        Dim DefaultInputDirectory = ConfigurationManager.AppSettings("DefaultInputDirectory")
        If DefaultInputDirectory <> "" And Not Directory.Exists(DefaultInputDirectory) Then
            MsgBox("Default Input Directory: " & DefaultInputDirectory & " does not exist. " &
                    "DriveDefaultBlogExportFileSearchToBook() is aborting!", , MsgBoxTitle)
            Exit Sub
        End If

        Dim DefaultXMLInputFilePathAndName As String = GetDefaultXMLInputFilePathAndName()
        If Not File.Exists(DefaultXMLInputFilePathAndName) Then
            MsgBox("Default XML Input File: " & DefaultXMLInputFilePathAndName & " does not exist. " &
                    "DriveDefaultBlogExportFileSearchToBook() is aborting!", , MsgBoxTitle)
            Exit Sub
        End If

        Dim SearchStrArr(4) As String ' Defines array of 5 elements with index from 0 to 4
        'Ref: https://learn.microsoft.com/en-us/office/vba/language/concepts/getting-started/declaring-arrays
        Dim ix
        For ix = 0 To 4
            SearchStrArr(ix) = ""
        Next ix

        Dim CreateBlogBook As Boolean
        CreateBlogBook = True

        Dim BlogBookFilePathAndName As String, LogFilePathAndName As String
        Dim CommentsDictFilePathAndName As String, CommentsLogFilePathAndName As String
        BlogBookFilePathAndName = Left(DefaultXMLInputFilePathAndName,
                                       Len(DefaultXMLInputFilePathAndName) - 4) & "-BlogBook.html"
        LogFilePathAndName = Left(DefaultXMLInputFilePathAndName,
                                  Len(DefaultXMLInputFilePathAndName) - 4) & "-BlogBookLog.txt"
        CommentsDictFilePathAndName = Left(DefaultXMLInputFilePathAndName,
                                           Len(DefaultXMLInputFilePathAndName) - 4) _
        & "-CommentsDict.html"
        CommentsLogFilePathAndName = Left(DefaultXMLInputFilePathAndName,
                                          Len(DefaultXMLInputFilePathAndName) - 4) _
        & "-CommentsLog.txt"

        Dim CommentsDict As Dictionary(Of String, String)
        CommentsDict = New Dictionary(Of String, String)
        CreateCommentsDictionary(CommentsDict, DefaultXMLInputFilePathAndName, SearchStrArr,
            False, CreateBlogBook, CommentsLogFilePathAndName)

        WriteCommentsDictionaryToFile(CommentsDict, CommentsDictFilePathAndName)

        Dim NumPostsAndPagesProcessed
        NumPostsAndPagesProcessed = BlogExportFileSearchToBook(DefaultXMLInputFilePathAndName,
        SearchStrArr, False, BlogBookFilePathAndName, LogFilePathAndName, CommentsDict)

        Dim msg
        msg = "Functions: CreateCommentsDictionary, WriteCommentsDictionaryToFile and " &
            "BlogExportFileSearchToBook finished execution." & vbCrLf & vbCrLf &
            "Number of pages and posts processed: " & NumPostsAndPagesProcessed & vbCrLf &
            "For details, see log files: '" & Path.GetFileName(LogFilePathAndName) &
            "' and '" & Path.GetFileName(CommentsLogFilePathAndName) & "'."
        MsgBox(msg, , MsgBoxTitle)

    End Sub

    Sub DrivePromptInputBlogExportFileSearchToBook()
        Dim XMLInputFilePathAndName As String
        Dim fd As New OpenFileDialog()

        If LastXMLInputFilePath = "" Then
            Dim DefaultInputDirectory As String
            'DefaultXMLInputFileDirectory = Directory.GetParent(DefaultXMLInputFilePathAndName).ToString()
            DefaultInputDirectory = ConfigurationManager.AppSettings("DefaultInputDirectory")
            If DefaultInputDirectory <> "" And System.IO.Directory.Exists(DefaultInputDirectory) Then
                fd.InitialDirectory = DefaultInputDirectory
            End If
        Else
            fd.InitialDirectory = LastXMLInputFilePath
        End If
        fd.Title = "Generate HTML Blogbook from Blog Backup XML file" & ": Select input XML export file"
        fd.Filter = "XML Files|*.xml"
        If fd.ShowDialog() = DialogResult.OK Then
            XMLInputFilePathAndName = fd.FileName
            LastXMLInputFilePath = Directory.GetParent(XMLInputFilePathAndName).ToString()
        Else
            MsgBox("Input XML file Not chosen. Aborting Sub DrivePromptInputBlogExportFileSearchToBook",
                , MsgBoxTitle)
            Exit Sub
        End If
        Dim msg, Style, response

        Dim BlogBookFilePathAndName As String, LogFilePathAndName As String
        Dim CommentsDictFilePathAndName As String, CommentsLogFilePathAndName As String

        Dim SearchStrArr(4) As String ' Defines array of 5 elements with index from 0 to 4
        Dim ix
        For ix = 0 To 4
            SearchStrArr(ix) = ""
        Next ix

        Dim CreateBlogBook As Boolean
        CreateBlogBook = True

        Dim CombineSearchStrAsAND As Boolean
        CombineSearchStrAsAND = False

        Dim ContinueRun As Boolean
        ContinueRun = True

        Dim FirstRun As Boolean
        FirstRun = True

        Do While ContinueRun

            RunBEFToBookProgram = False
            BEFToBookParamForm.ShowDialog()

            If RunBEFToBookProgram Then
                With BEFToBookParamForm
                    SearchStrArr(0) = .TextBoxSS1.Text
                    SearchStrArr(1) = .TextBoxSS2.Text
                    SearchStrArr(2) = .TextBoxSS3.Text
                    SearchStrArr(3) = .TextBoxSS4.Text
                    SearchStrArr(4) = .TextBoxSS5.Text
                    CombineSearchStrAsAND = .CheckBoxCombineUsingAND.Checked
                    CreateBlogBook = .CheckBoxCreateBlogbook.Checked
                End With
            Else
                MsgBox("User clicked Cancel command button or closed form window. Aborting!",
                    , MsgBoxTitle)
                Exit Sub
            End If

            If CreateBlogBook Then
                BlogBookFilePathAndName = Left(XMLInputFilePathAndName, Len(XMLInputFilePathAndName) - 4) & "-BlogBook.html"
            Else
                BlogBookFilePathAndName = ""
            End If
            LogFilePathAndName = Left(XMLInputFilePathAndName, Len(XMLInputFilePathAndName) - 4) & "-BlogBookLog.txt"
            If CreateBlogBook Then
                CommentsDictFilePathAndName = Left(XMLInputFilePathAndName, Len(XMLInputFilePathAndName) - 4) _
                & "-CommentsDict.html"
            Else
                CommentsDictFilePathAndName = ""
            End If

            CommentsLogFilePathAndName = Left(XMLInputFilePathAndName, Len(XMLInputFilePathAndName) - 4) _
            & "-CommentsLog.txt"

            If CreateBlogBook Or FirstRun Then
                FirstRun = False
                msg = "Input file: " & XMLInputFilePathAndName & vbCrLf
                For ix = 0 To 4
                    msg = msg & "SearchStrArr(" & ix & "): " & SearchStrArr(ix) & vbCrLf
                Next ix
                msg = msg & "CombineSearchStrAsAND: " & CombineSearchStrAsAND & vbCrLf

                msg = msg & "Output files will be created in same directory as input file and will be overwritten if they exist." &
            " So only their filenames are given below." & vbCrLf & vbCrLf
                If BlogBookFilePathAndName <> "" Then
                    msg = msg & "Blog book: " & Path.GetFileName(BlogBookFilePathAndName) & vbCrLf
                End If
                msg = msg & "Main Log: " & Path.GetFileName(LogFilePathAndName) & vbCrLf

                If CommentsDictFilePathAndName <> "" Then
                    msg = msg & "Comments Dictionary book: " & Path.GetFileName(CommentsDictFilePathAndName) & vbCrLf
                End If
                msg = msg & "Comments Log: " & Path.GetFileName(CommentsLogFilePathAndName) & vbCrLf & vbCrLf

                If CreateBlogBook Then
                    msg = msg & "Do you want to continue and run Functions: CreateCommentsDictionary, " &
                    "WriteCommentsDictionaryToFile and BlogExportFileSearchToBook to create blog book, comments dictionary book" &
                    " as well as log files?"
                Else
                    msg = msg & "Do you want to continue and run Functions: CreateCommentsDictionary, " &
                    "and BlogExportFileSearchToBook to only create log files?"
                End If

                Style = vbYesNo Or vbDefaultButton1    ' Define buttons.
                response = MsgBox(msg, Style, MsgBoxTitle)
                If response = vbNo Then    ' User chose No.
                    ContinueRun = False
                    Exit Do
                End If
            End If

            Dim CommentsDict As Dictionary(Of String, String)
            CommentsDict = New Dictionary(Of String, String)
            CreateCommentsDictionary(CommentsDict, XMLInputFilePathAndName, SearchStrArr, CombineSearchStrAsAND,
                    CreateBlogBook, CommentsLogFilePathAndName)

            If CreateBlogBook Then
                WriteCommentsDictionaryToFile(CommentsDict, CommentsDictFilePathAndName)
            End If

            Dim NumPostsAndPagesProcessed
            NumPostsAndPagesProcessed = BlogExportFileSearchToBook(XMLInputFilePathAndName, SearchStrArr,
            CombineSearchStrAsAND, BlogBookFilePathAndName, LogFilePathAndName, CommentsDict)

            If CreateBlogBook Then
                msg = "Functions: CreateCommentsDictionary, WriteCommentsDictionaryToFile and " &
                    "BlogExportFileSearchToBook finished execution." & vbCrLf & vbCrLf &
                    "Number of pages and posts processed: " & NumPostsAndPagesProcessed & vbCrLf &
                    "For details, see log files: '" & Path.GetFileName(LogFilePathAndName) &
                    "' and '" & Path.GetFileName(CommentsLogFilePathAndName) & "'."
            Else
                msg = "Functions: CreateCommentsDictionary and " & "BlogExportFileSearchToBook finished execution " &
                    "with settings to only create log files." & vbCrLf & vbCrLf &
                    "Number of pages and posts processed: " & NumPostsAndPagesProcessed & vbCrLf &
                    "Log files created are: '" & Path.GetFileName(LogFilePathAndName) &
                    "' and '" & Path.GetFileName(CommentsLogFilePathAndName) & "'."
            End If
            msg = msg & vbCrLf & vbCrLf & "Do you want to run main functions BlogExportFileSearchtoBook etc." _
                & " again on same input file?"
            Style = vbYesNo Or vbDefaultButton1    ' Define buttons.
            response = MsgBox(msg, Style, MsgBoxTitle)
            If response = vbNo Then    ' User chose No.
                ContinueRun = False
            End If
        Loop

        BEFToBookParamForm.Close()
    End Sub

    Function BlogExportFileSearchToBook(XMLInputFilePathAndName As String, SearchStrArr() As String,
    CombineSearchStrAsAND As Boolean, BlogBookFilePathAndName As String, LogFilePathAndName As String,
                                        CommentsDict As Dictionary(Of String, String)) As Integer

        Dim StartTime, EndTime As String
        StartTime = Now

        LogFileStream = My.Computer.FileSystem.OpenTextFileWriter(LogFilePathAndName, False) 'Overwrite not append

        Dim msg As String
        msg = "Function BlogExportFileSearchToBook start date & time is: " & StartTime
        Debug.Print(msg)
        LogFileStream.WriteLine(msg)

        msg = "Arguments to Function BlogExportFileSearchToBook are: XMLInputFilePathAndName = " &
        XMLInputFilePathAndName & vbCrLf
        Dim ix
        For ix = 0 To 4
            msg = msg & "SearchStrArr(" & ix & "): " & SearchStrArr(ix) & vbCrLf
        Next ix
        msg = msg & "CombineSearchStrAsAND = " & CombineSearchStrAsAND & vbCrLf
        msg = msg & "BlogBookFilePathAndName = " & BlogBookFilePathAndName & vbCrLf &
            "LogFilePathAndName = " & LogFilePathAndName & vbCrLf &
            "CommentsDict object (data not logged)" & vbCrLf
        Debug.Print(msg)
        LogFileStream.WriteLine(msg)

        Dim CreateBlogBook As Boolean

        If BlogBookFilePathAndName = "" Then
            CreateBlogBook = False
        Else
            CreateBlogBook = True
        End If

        Dim XDoc As XmlDocument
        XDoc = New XmlDocument()

        ' https://stackoverflow.com/questions/24734/selectnodes-not-working-on-stackoverflow-feed
        Dim nsmgr As XmlNamespaceManager
        nsmgr = New XmlNamespaceManager(XDoc.NameTable)
        nsmgr.AddNamespace("atom", "http://www.w3.org/2005/Atom")

        XDoc.Load(XMLInputFilePathAndName)
        msg = "XmlDocument (var XDoc) created and file: '" & XMLInputFilePathAndName &
            "' loaded into it."
        Debug.Print(msg)
        LogFileStream.WriteLine(msg)

        Dim BlogDetails As BlogDetailsType
        BlogDetails = GetBlogDetailsFromXMLDoc(XDoc, nsmgr)

        Dim Entries As XmlNodeList
        Const XMLSelectNodesParameter As String = SelectFeedEntryXMLTerm
        Entries = XDoc.DocumentElement.SelectNodes(XMLSelectNodesParameter, nsmgr)

        Dim NumPostsAndPagesProcessed
        NumPostsAndPagesProcessed = IterateThroughAllEntries(XMLInputFilePathAndName, BlogDetails,
            Entries, nsmgr, SearchStrArr, CombineSearchStrAsAND, CommentsDict, CreateBlogBook,
            BlogBookFilePathAndName)

        EndTime = Now
        msg = "Date & time now is: " & Now & vbCrLf &
            "Time taken for this function execution (in seconds): " _
            & DateDiff("s", StartTime, EndTime) & " (excludes dialog interaction by driver (invoker) sub)"
        Debug.Print(msg)
        LogFileStream.WriteLine(msg)

        LogFileStream.Close()

        BlogExportFileSearchToBook = NumPostsAndPagesProcessed
    End Function


    Function IterateThroughAllEntries(XMLInputFilePathAndName As String, BlogDetails As BlogDetailsType,
        Entries As XmlNodeList, nsmgr As XmlNamespaceManager, SearchStrArr() As String,
        CombineSearchStrAsAND As Boolean, CommentsDict As Dictionary(Of String, String),
        CreateBlogBook As Boolean, BlogBookFilePathAndName As String) As Integer

        Dim BlogBookFileStream As System.IO.StreamWriter


        If CreateBlogBook Then
            BlogBookFileStream = My.Computer.FileSystem.OpenTextFileWriter(BlogBookFilePathAndName, False)
        End If

        Dim msg As String

        msg = vbCrLf & "BlogDetails:" & vbCrLf & "Updated = " & BlogDetails.UpdatedDate & vbCrLf &
        "Title = " & BlogDetails.Title & vbCrLf &
        "href = " & BlogDetails.href & vbCrLf &
        "AuthorName = " & BlogDetails.AuthorName
        Debug.Print(msg)
        LogFileStream.WriteLine(msg)

        WriteEntriesHeaderToLog(Entries.Count)

        Dim Entry As XmlNode

        Dim EntryIndex As Long
        EntryIndex = 1

        Dim TotalPostsAndPages
        TotalPostsAndPages = 0
        Dim TotalContentLength
        TotalContentLength = 0

        Dim ContentHTML As String
        ContentHTML = ""

        If CreateBlogBook Then
            ContentHTML = GetHeaderHTMLForBlogBook(XMLInputFilePathAndName, BlogDetails, SearchStrArr,
            CombineSearchStrAsAND)
            BlogBookFileStream.Write(ContentHTML)
            ContentHTML = ""
        End If

        Dim MatchedSearchString As String
        MatchedSearchString = ""
        For Each Entry In Entries
            Dim PPCEntry As PPCEntryType
            ' Code below is to avoid getting a compiler warning about unitialized PPCEntry
            ' In quick browse, I could not get a way to doing it better
            With PPCEntry
                .Index = 0
                .CategoryTermKind = ""
                .Title = ""
                .Content = ""
                .href = ""
                .AuthorName = ""
                .PublishedDate = ""
                .UpdatedDate = ""
            End With
            ReadEntryIntoPPCEntry(Entry, nsmgr, PPCEntry)
            If ((PPCEntry.CategoryTermKind = CategoryTermKindPostXMLName) Or
                    (PPCEntry.CategoryTermKind = CategoryTermKindPageXMLName)) Then ' continue processing the entry

                Dim CategoryTermKindProperCase
                CategoryTermKindProperCase = UCase(Left(PPCEntry.CategoryTermKind, 1)) &
                Mid(PPCEntry.CategoryTermKind, 2)
                MatchedSearchString = DoesEntryMatchSearchCriteria(SearchStrArr, CombineSearchStrAsAND, PPCEntry)
                If MatchedSearchString <> "" Then 'Entry matched or SearchStrArr is empty and Entry has to be
                    'processed.
                    LogPPCEntry(EntryIndex, PPCEntry, MatchedSearchString)

                    If CreateBlogBook Then
                        ContentHTML &= CreateHTMLForPPCEntry(PPCEntry, EntryIndex)
                    End If

                    'Check if post/page has comments and if so, insert comments
                    Dim PlainHTTPhref
                    PlainHTTPhref = GetPlainHTTPhref(PPCEntry)
                    If CommentsDict.ContainsKey(PlainHTTPhref) Then
                        msg = CategoryTermKindProperCase & " has comment(s)."
                        Debug.Print(msg)
                        LogFileStream.WriteLine(msg)

                        If CreateBlogBook Then
                            ContentHTML &= PageOrPostComments(msg, CommentsDict.Item(PlainHTTPhref))
                        End If
                    End If

                    If CreateBlogBook Then
                        ContentHTML = ContentHTML &
                    "============================End of " & CategoryTermKindProperCase &
                    "==========================<br/><br/>" & vbCrLf

                        BlogBookFileStream.Write(ContentHTML)
                        ContentHTML = ""
                    End If

                    TotalPostsAndPages += 1
                    TotalContentLength += Len(PPCEntry.Content)
                End If ' DoesEntryMatchSearchCriteria() <> ""
            End If ' CategoryTermKind is post or page
            EntryIndex += 1
        Next

        If CreateBlogBook Then
            ContentHTML = GetEndInfoHTMLForBlogBook()
            BlogBookFileStream.Write(ContentHTML)
            BlogBookFileStream.Close()
        End If

        msg = "Function BlogExportFileSearchToBook execution finishing." & vbCrLf
        If CreateBlogBook Then
            msg = msg + "Blog book created (and comments dictionary book should have been created). " &
                "Total posts and pages in blog book: " &
                TotalPostsAndPages & " and its (total) content length: " & TotalContentLength & vbCrLf
        Else
            msg = msg + "Only log files created. Total posts and pages mentioned in main log file: " &
                TotalPostsAndPages & " and total content length of posts & pages mentioned in main log file: " &
                TotalContentLength & vbCrLf
        End If
        Debug.Print(msg)
        LogFileStream.WriteLine(msg)

        IterateThroughAllEntries = TotalPostsAndPages ' Return number of posts & pages that match search criteria

    End Function

    Sub WriteEntriesHeaderToLog(NumEntries)
        Dim msg
        msg = vbCrLf & "Number of entries in input file:" & NumEntries & vbCrLf
        Debug.Print(msg)
        LogFileStream.WriteLine(msg)

        msg = "Entries that are written to output are listed below. Format of data about these entries " &
            vbCrLf & "(tab is the separator between data items on each line making it easy to import" &
            " into/view in programs like Excel):"
        Debug.Print(msg)
        LogFileStream.WriteLine(msg)

        Dim TabChar As String
        TabChar = Chr(9)

        ' Print long form of format for each post/page
        msg = "Index of entry\tCategoryTermKind\tPublished-Date-YYYY-MM-DD\t" &
        "Updated-Date-YYYY-MM-DD\t" &
        "Title (full)\tLink\tContent-length\tAuthor-Name\t(Optional)Matched Search String"
        Debug.Print(msg)
        LogFileStream.WriteLine(msg)

        ' Then print short form name for each field/column. This will be useful when viewing in Excel
        msg = "Ind" & TabChar & "Catg" & TabChar & "Pub-Date" & TabChar & "Upd-Date" & TabChar &
        "Title (full)" & TabChar & "Link" & TabChar & "ContLen" & TabChar & "Author" &
        TabChar & "Opt. Match-Str"
        Debug.Print(msg)
        LogFileStream.WriteLine(msg)

    End Sub

    Function GetHeaderHTMLForBlogBook(XMLInputFilePathAndName As String, BlogDetails As BlogDetailsType,
    SearchStrArr() As String, CombineSearchStrAsAND As Boolean) As String
        Dim ContentHTML As String
        Dim BlogUpdatedDateAsDate As Date
        Dim dstr As String
        ' From "2023-07-12T01:09:28.850-07:00" to "2023-07-12 01:09:28"
        ' Note that "2023-07-12T01:09:28.850-07:00" does not seem to work with CDate
        dstr = Left(BlogDetails.UpdatedDate, 10) & " " & Mid(BlogDetails.UpdatedDate, 12, 8)
        BlogUpdatedDateAsDate = CDate(dstr)

        ContentHTML = "<html><head><meta charset=""utf-8""><title>" & BlogDetails.Title &
        "</title></head><body>" &
        "<style>body{overflow-wrap:break-word; word-break:break-word; word-wrap:break-word}</style>"
        ContentHTML = ContentHTML & "<h1>" & BlogDetails.Title & " Blogbook</h1><br/>" &
        "Blog XML Backup file (input file) path and name: " & XMLInputFilePathAndName & "<br/>" &
        "Blog last updated date & time as per XML Backup file: " & BlogUpdatedDateAsDate &
        " [Raw data with TZ: " & BlogDetails.UpdatedDate & "]<br/>" &
        "Blog address (URL): " & "<a href=""" & BlogDetails.href & """>" & BlogDetails.href & "</a><br/>" &
        "Blog author name: " & BlogDetails.AuthorName & "<br/>" &
        "Date & time of creation of this HTML blogbook (excluding contents links (TOC)): " & Now & "<br/>"
        If IsEmptySearchStrArr(SearchStrArr) Then
            ContentHTML &= "<br/>This blogbook has contents of all pages and posts in input file.<br/>"
        Else
            ContentHTML = ContentHTML & "<br/>This blogbook has contents of pages and posts in input file " &
            "matching the following search strings:<br/> "
            Dim ix
            For ix = 0 To 4
                If SearchStrArr(ix) <> "" Then
                    ContentHTML = ContentHTML & SearchStrArr(ix) & "<br/>"
                End If
            Next ix
            If CombineSearchStrAsAND Then
                ContentHTML &= "---<br/>Above search strings are combined as AND.<br/>"
            Else
                ContentHTML &= "---<br/>Above search strings are combined as OR.<br/>"
            End If
        End If
        ContentHTML &= "=================================================================<br/>"
        GetHeaderHTMLForBlogBook = ContentHTML
    End Function

    Function GetEndInfoHTMLForBlogBook() As String
        Dim ContentHTML As String
        Dim FooterHTML
        ContentHTML = "<br/>============================<b>End of Book</b>=========================" & "<br/><br/>"
        FooterHTML = ReadBlogBookFooterHTMLFromFile()
        ContentHTML &= FooterHTML

        GetEndInfoHTMLForBlogBook = ContentHTML

    End Function

    Function ReadBlogBookFooterHTMLFromFile() As String
        Dim InputFileData As String
        Dim InputTSO As System.IO.StreamReader
        Dim BlogBookFooterFileName As String = ConfigurationManager.AppSettings("BlogBookFooterFileName")
        If File.Exists(BlogBookFooterFileName) Then
            InputTSO = My.Computer.FileSystem.OpenTextFileReader(BlogBookFooterFileName) ' In same directory
            '                               as running program and so only filename is enough
            InputFileData = InputTSO.ReadToEnd()
            InputTSO.Close()
            ReadBlogBookFooterHTMLFromFile = InputFileData
        Else
            Dim msg = "Blogbook Footer File: " & BlogBookFooterFileName & " not found. " &
            "So ReadBlogBookFooterHTMLFromFile() is returning empty string"
            Debug.Print(msg)
            LogFileStream.WriteLine(msg)
            MsgBox(msg, , MsgBoxTitle)
            ReadBlogBookFooterHTMLFromFile = ""
        End If
    End Function

    Sub ReadEntryIntoPPCEntry(Entry As XmlNode, nsmgr As XmlNamespaceManager, ByRef PPCEntry As PPCEntryType)

        PPCEntry.CategoryTermKind = GetCategoryTermKindFromEntry(Entry, nsmgr)
        PPCEntry.Title = GetTitleFromEntry(Entry, nsmgr)
        PPCEntry.Content = GetContentFromEntry(Entry, nsmgr)
        PPCEntry.PublishedDate = GetPublishedDateFromEntry(Entry, nsmgr)
        PPCEntry.UpdatedDate = GetUpdatedDateFromEntry(Entry, nsmgr)
        PPCEntry.href = GetAlternateHrefFromEntry(Entry, nsmgr)
        PPCEntry.AuthorName = GetAuthorNameFromEntry(Entry, nsmgr)

    End Sub

    Sub LogPPCEntry(EntryIndex As Long, PPCEntry As PPCEntryType, Optional MatchedSearchStr As String = "")

        Dim TabChar As String
        TabChar = Chr(9)

        Dim msg, Title
        Dim ix

        ix = InStr(PPCEntry.Title, vbLf)
        If ix > 0 Then
            Title = Left(PPCEntry.Title, ix - 1)
        Else
            Title = PPCEntry.Title
        End If

        msg = EntryIndex & TabChar & PPCEntry.CategoryTermKind & TabChar & Left(PPCEntry.PublishedDate, 10) & TabChar _
        & Left(PPCEntry.UpdatedDate, 10) & TabChar & Title & TabChar & PPCEntry.href _
        & TabChar & Len(PPCEntry.Content) & TabChar & PPCEntry.AuthorName

        If MatchedSearchStr <> "" Then
            msg = msg & TabChar & MatchedSearchStr
        End If

        Debug.Print(msg)
        LogFileStream.WriteLine(msg)

    End Sub

    Function IsEmptySearchStrArr(SearchStrArr() As String) As Boolean

        Dim EmptySearchStrArr As Boolean
        EmptySearchStrArr = True
        Dim ix
        For ix = 0 To 4
            If SearchStrArr(ix) <> "" Then
                EmptySearchStrArr = False
                Exit For
            End If
        Next ix

        IsEmptySearchStrArr = EmptySearchStrArr

    End Function

    'Return values:
    ' "-": Empty Search String Array and so entry should be processed
    ' "": Entry did not match any Search String in (non empty) Search String Array
    ' For OR op: Search String: Entry matched one of the (non empty) Search Strings in (non empty) Search String Array and
    '  that (non empty) Matched String is returned with some additional characters before and after
    ' For AND op: Search String: Entry matched all of the (non empty) Search Strings in (non empty) Search String Array
    ' and those (non empty) Matched Strings are returned with some additional characters before and after
    Function DoesEntryMatchSearchCriteria(SearchStrArr() As String, CombineSearchStrAsAND As Boolean,
    PPCEntry As PPCEntryType) As String

        If IsEmptySearchStrArr(SearchStrArr) Then
            DoesEntryMatchSearchCriteria = "-"
            Exit Function
        End If

        Dim MatchedStringPlus
        If CombineSearchStrAsAND Then
            MatchedStringPlus = DoesEntryMatchSearchStingsUsingAND(SearchStrArr, PPCEntry)
            ' Matched strings plus in case of AND
        Else
            MatchedStringPlus = DoesEntryMatchSearchStingsUsingOR(SearchStrArr, PPCEntry)
        End If

        DoesEntryMatchSearchCriteria = MatchedStringPlus
    End Function

    Function DoesEntryMatchSearchStingsUsingOR(SearchStrArr() As String, PPCEntry As PPCEntryType) As String
        Dim MatchedStringPlus
        Dim ix, InStrix

        For ix = 0 To 4
            If SearchStrArr(ix) <> "" Then
                If PPCEntry.Title <> "" Then
                    InStrix = InStr(1, PPCEntry.Title, SearchStrArr(ix), vbTextCompare) 'case insensitive search
                    If (InStrix > 0) Then
                        MatchedStringPlus = GetMatchedStringPlus(PPCEntry.Title, SearchStrArr(ix), InStrix)
                        DoesEntryMatchSearchStingsUsingOR = MatchedStringPlus
                        Exit Function
                    End If
                End If

                If PPCEntry.Content <> "" Then
                    InStrix = InStr(1, PPCEntry.Content, SearchStrArr(ix), vbTextCompare) 'case insensitive search
                    If (InStrix > 0) Then
                        MatchedStringPlus = GetMatchedStringPlus(PPCEntry.Content, SearchStrArr(ix), InStrix)
                        DoesEntryMatchSearchStingsUsingOR = MatchedStringPlus
                        Exit Function
                    End If
                End If
            End If
        Next ix
        DoesEntryMatchSearchStingsUsingOR = ""
    End Function

    Function DoesEntryMatchSearchStingsUsingAND(SearchStrArr() As String, PPCEntry As PPCEntryType) As String
        Dim MatchedStringsPlus
        Dim CurrentMatchedString
        Dim ix, InStrix

        MatchedStringsPlus = ""

        For ix = 0 To 4
            If SearchStrArr(ix) <> "" Then
                CurrentMatchedString = ""
                If PPCEntry.Title <> "" Then
                    InStrix = InStr(1, PPCEntry.Title, SearchStrArr(ix), vbTextCompare) 'case insensitive search
                    If (InStrix > 0) Then
                        CurrentMatchedString = GetMatchedStringPlus(PPCEntry.Title, SearchStrArr(ix), InStrix)
                        MatchedStringsPlus = MatchedStringsPlus & ix & ") " & CurrentMatchedString & " | "
                    End If
                End If

                If CurrentMatchedString = "" And PPCEntry.Content <> "" Then
                    InStrix = InStr(1, PPCEntry.Content, SearchStrArr(ix), vbTextCompare) 'case insensitive search
                    If (InStrix > 0) Then
                        CurrentMatchedString = GetMatchedStringPlus(PPCEntry.Content, SearchStrArr(ix), InStrix)
                        MatchedStringsPlus = MatchedStringsPlus & ix & ") " & CurrentMatchedString & " | "
                    Else
                        DoesEntryMatchSearchStingsUsingAND = ""
                        Exit Function
                    End If
                End If
            End If
        Next ix
        DoesEntryMatchSearchStingsUsingAND = MatchedStringsPlus
    End Function

    Function GetMatchedStringPlus(Content, SearchStr, InStrix) As String
        Dim MSPStartPos, MSPEndPos, x
        Dim SSLen, ContentLen

        If InStrix <= 10 Then
            MSPStartPos = 1
        Else
            MSPStartPos = InStrix - 10
        End If

        SSLen = Len(SearchStr)
        x = InStrix + SSLen + 10
        ContentLen = Len(Content)
        If x < ContentLen Then
            MSPEndPos = x
        Else
            MSPEndPos = ContentLen
        End If
        GetMatchedStringPlus = Mid(Content, MSPStartPos, MSPEndPos - MSPStartPos + 1)
    End Function

    ' Should this function be renamed to CreateHTMLForPageOrPostEntry as it is meant only for them and not
    ' comments, I think?
    Function CreateHTMLForPPCEntry(PPCEntry As PPCEntryType, EntryIndex As Long) As String


        Dim CategoryTermKindProperCase
        CategoryTermKindProperCase = UCase(Left(PPCEntry.CategoryTermKind, 1)) _
        & Mid(PPCEntry.CategoryTermKind, 2)

        Dim ContentLength
        ContentLength = Len(PPCEntry.Content)

        Dim ContentHTML As String
        ' Example of id in header tag: <h2 id="toc1">dum1 h2</h2>
        ContentHTML = "<h1 id=""entry-" & EntryIndex & """>" & PPCEntry.Title & "; Published: " _
        & Left(PPCEntry.PublishedDate, 10) & "</h1><br/>"
        ContentHTML = ContentHTML & CategoryTermKindProperCase & " link (URL) on blog: <a href=""" &
        PPCEntry.href & """>" & PPCEntry.href & "</a><br/><br/>"
        ContentHTML = ContentHTML & PPCEntry.Content & "<br/><br/>"

        CreateHTMLForPPCEntry = ContentHTML
    End Function

    Function GetPlainHTTPhref(PPCEntry As PPCEntryType) As String
        ' Dictionary has hrefs from XML input file comments entries which use http://
        ' href in page/posts entries seems to be https://
        ' So if needed convert href of page/posts to http:// before checking in dictionary
        Dim PlainHTTPhref, ix
        ix = InStr(1, PPCEntry.href, "https://")
        If (ix = 1) Then
            PlainHTTPhref = "http://" & Right(PPCEntry.href, Len(PPCEntry.href) - 8)
        Else
            PlainHTTPhref = PPCEntry.href
        End If
        GetPlainHTTPhref = PlainHTTPhref
    End Function

    Function PageOrPostComments(msg As String, CommentsHTML As String) As String

        Dim ContentHTML
        ContentHTML = "=======================End of main body=======================<br/><br/>"
        ContentHTML = ContentHTML & "<b>" & msg & "</b><br/><br/>" & CommentsHTML

        PageOrPostComments = ContentHTML
    End Function

    Function GetBlogDetailsFromXMLDoc(XDoc, nsmgr) As BlogDetailsType
        Dim BlogDetails As BlogDetailsType

        Dim Node As XmlNode
        Node = XDoc.SelectSingleNode("/atom:feed/atom:updated", nsmgr)
        BlogDetails.UpdatedDate = Node.InnerText

        Node = XDoc.SelectSingleNode("/atom:feed/atom:title", nsmgr)
        BlogDetails.Title = Node.InnerText

        Node = XDoc.SelectSingleNode("/atom:feed/atom:author/atom:name", nsmgr)
        BlogDetails.AuthorName = Node.InnerText

        Dim relAttr As XmlNode
        Dim hrefAttr As XmlNode
        Dim hrefText, LinkNode
        hrefText = ""

        Dim LinkNodes As XmlNodeList
        LinkNodes = XDoc.SelectNodes("/atom:feed/atom:link", nsmgr)
        'Dim Attr
        For Each LinkNode In LinkNodes
            relAttr = LinkNode.Attributes.getNamedItem("rel")
            If (relAttr.Value = "alternate") Then
                hrefAttr = LinkNode.Attributes.getNamedItem("href")
                hrefText = hrefAttr.Value
            End If
        Next
        BlogDetails.href = hrefText

        GetBlogDetailsFromXMLDoc = BlogDetails

    End Function

    Function GetCategoryTermKindFromEntry(Entry As XmlNode, nsmgr As XmlNamespaceManager) As String
        Dim termAttr As XmlNode
        Dim termKindText

        Dim CategoryNode As XmlNode
        CategoryNode = Entry.SelectSingleNode(CategoryXMLName, nsmgr)

        termAttr = CategoryNode.Attributes.GetNamedItem(CategoryTermAttributeXMLName)
        If (Right(termAttr.Value, 4) = CategoryTermKindPostXMLName) Then
            termKindText = CategoryTermKindPostXMLName
        ElseIf (Right(termAttr.Value, 4) = CategoryTermKindPageXMLName) Then
            termKindText = CategoryTermKindPageXMLName
        ElseIf (Right(termAttr.Value, 7) = CategoryTermKindCommentXMLName) Then
            termKindText = CategoryTermKindCommentXMLName
        ElseIf (Right(termAttr.Value, 8) = CategoryTermKindSettingsXMLName) Then
            termKindText = CategoryTermKindSettingsXMLName
        ElseIf (Right(termAttr.Value, 8) = CategoryTermKindTemplateXMLName) Then
            termKindText = CategoryTermKindTemplateXMLName
        Else
            termKindText = "" ' Unexpected category term kind
        End If

        GetCategoryTermKindFromEntry = termKindText

    End Function

    Function GetAlternateHrefFromEntry(Entry As XmlNode, nsmgr As XmlNamespaceManager) As String
        Dim relAttr As XmlNode
        Dim hrefAttr As XmlNode
        Dim hrefText, LinkNode
        hrefText = ""

        Dim LinkNodes As XmlNodeList
        LinkNodes = Entry.SelectNodes("atom:link", nsmgr)
        For Each LinkNode In LinkNodes
            relAttr = LinkNode.Attributes.getNamedItem("rel")
            If (relAttr.Value = "alternate") Then
                hrefAttr = LinkNode.Attributes.getNamedItem("href")
                hrefText = hrefAttr.Value
            End If
        Next

        GetAlternateHrefFromEntry = hrefText

    End Function

    Function GetTitleFromEntry(Entry As XmlNode, nsmgr As XmlNamespaceManager) As String
        Dim TitleText

        Dim TitleNode As XmlNode
        TitleNode = Entry.SelectSingleNode("atom:title", nsmgr)

        TitleText = TitleNode.InnerText

        GetTitleFromEntry = TitleText

    End Function

    Function GetContentFromEntry(Entry As XmlNode, nsmgr As XmlNamespaceManager) As String
        Dim ContentText

        Dim ContentNode As XmlNode
        ContentNode = Entry.SelectSingleNode("atom:content", nsmgr)

        ContentText = ContentNode.InnerText

        GetContentFromEntry = ContentText

    End Function

    Function GetPublishedDateFromEntry(Entry As XmlNode, nsmgr As XmlNamespaceManager) As String
        Dim PublishedDateText

        Dim PublishedNode As XmlNode
        PublishedNode = Entry.SelectSingleNode("atom:published", nsmgr)

        PublishedDateText = PublishedNode.InnerText

        GetPublishedDateFromEntry = PublishedDateText
        '   I think the above function inner code can be just one line given below (the part of the line after =
        '   can be used in caller to avoid having a function invocation itself).
        '   But I was new to these XML parsing functions at the time I wrote functions like these and so
        '   needed variables capturing the data for easy debugging.
        '   I think time taken for the above code is insignificant and besides the debugging advantage, the code
        '   is quite readable too. So I am choosing to retain the above code in this function and similar other functions.
        '    GetPublishedDateFromEntry = Entry.SelectSingleNode("published").Text
    End Function

    Function GetUpdatedDateFromEntry(Entry As XmlNode, nsmgr As XmlNamespaceManager) As String
        Dim UpdatedDateText

        Dim UpdatedNode As XmlNode
        UpdatedNode = Entry.SelectSingleNode("atom:updated", nsmgr)

        UpdatedDateText = UpdatedNode.InnerText

        GetUpdatedDateFromEntry = UpdatedDateText

    End Function

    Function GetAuthorNameFromEntry(Entry As XmlNode, nsmgr As XmlNamespaceManager) As String
        Dim AuthorNameText

        Dim AuthorNameNode As XmlNode
        AuthorNameNode = Entry.SelectSingleNode("atom:author/atom:name", nsmgr)
        AuthorNameText = AuthorNameNode.InnerText ' Inner Text has the name text
        GetAuthorNameFromEntry = AuthorNameText

    End Function

    Function GetDefaultXMLInputFilePathAndName() As String
        Dim DefaultXMLInputFilePathAndName As String
        Dim DefaultInputDirectory As String
        DefaultInputDirectory = ConfigurationManager.AppSettings("DefaultInputDirectory")
        If (DefaultInputDirectory = "") Then
            DefaultXMLInputFilePathAndName = ConfigurationManager.AppSettings("DefaultXMLInputFileName")
        Else
            DefaultXMLInputFilePathAndName = DefaultInputDirectory & "\" &
                ConfigurationManager.AppSettings("DefaultXMLInputFileName")
        End If
        GetDefaultXMLInputFilePathAndName = DefaultXMLInputFilePathAndName
    End Function


End Module
