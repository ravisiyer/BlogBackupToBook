Imports System.Xml

Module ModCreateCommentsDictionary
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

    Function CreateCommentsDictionary(dict As Dictionary(Of String, String), XMLInputFilePathAndName As String,
                                      SearchStrArr() As String, CombineSearchStrAsAND As Boolean,
                                      CreateBlogBook As Boolean, LogFilePathAndName As String) As Integer

        Dim StartTime As String, EndTime As String
        StartTime = Now

        LogFileStream = My.Computer.FileSystem.OpenTextFileWriter(LogFilePathAndName, False) 'Overwrite not append

        Dim msg As String
        msg = "Function CreateCommentsDictionary start date & time is: " & StartTime
        Debug.Print(msg)
        LogFileStream.WriteLine(msg)

        msg = "Arguments to Fn. CreateCommentsDictionary are: dict object (data not logged)" & vbCrLf &
        "XMLInputFilePathAndName = " & XMLInputFilePathAndName & vbCrLf
        Dim ix
        For ix = 0 To 4
            msg = msg & "SearchStrArr(" & ix & ") = " & SearchStrArr(ix) & vbCrLf
        Next ix
        msg = msg & "CombineSearchStrAsAND = " & CombineSearchStrAsAND & vbCrLf
        msg = msg & "CreateBlogBook = " & CreateBlogBook & vbCrLf &
            "LogFilePathAndName = " & LogFilePathAndName & vbCrLf

        Debug.Print(msg)
        LogFileStream.WriteLine(msg)

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

        Dim Entries As XmlNodeList
        Const XMLSelectNodesParameter As String = SelectFeedEntryXMLTerm
        Entries = XDoc.DocumentElement.SelectNodes(XMLSelectNodesParameter, nsmgr)

        If (Entries.Count <= 0) Then
            'Something is wrong. Abort.
            msg = "SelectNodes returned 0 entries! Aborting Function CreateCommentsDictionary!"
            Debug.Print(msg)
            LogFileStream.WriteLine(msg)
            MsgBox(msg, , MsgBoxTitle)
            CreateCommentsDictionary = 0
            Exit Function
        End If
        Dim NumCommentsProcessed
        Dim TotalContentLength As Long
        TotalContentLength = 0
        NumCommentsProcessed = IterateThroughAllCommentEntries(dict,
        Entries, nsmgr, SearchStrArr, CombineSearchStrAsAND, TotalContentLength)

        EndTime = Now
        msg = "Function CreateCommentsDictionary execution finishing." & vbCrLf
        msg = msg + "Total comments in input file: " &
            NumCommentsProcessed & " and its (total) content length: " & TotalContentLength & vbCrLf
        msg = msg + "Date & time now is: " & Now & vbCrLf &
            "Time taken for this function execution (in seconds): " & DateDiff("s", StartTime, EndTime)
        Debug.Print(msg)
        LogFileStream.WriteLine(msg)

        LogFileStream.Close()

        CreateCommentsDictionary = NumCommentsProcessed
    End Function

    Function IterateThroughAllCommentEntries(dict As Dictionary(Of String, String), Entries As XmlNodeList,
                                             nsmgr As XmlNamespaceManager, SearchStrArr() As String,
                                             CombineSearchStrAsAND As Boolean,
                                             ByRef TotalContentLength As Long) As Integer

        WriteEntriesHeaderToLog(Entries.Count)

        Dim Entry As XmlNode

        Dim EntryIndex As Long
        EntryIndex = 1

        Dim TotalComments
        TotalComments = 0

        Dim ContentHTML
        ContentHTML = ""
        Dim msg

        For Each Entry In Entries

            ' For this processing I think Page or Post or Comment entries can be handled similarly
            ' If my thinking is right then PageOrPostEntry may be renamed to something that indicates
            ' it is common to page, post and comment ... Renamed it ... code has to be tested
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

            If (PPCEntry.CategoryTermKind = CategoryTermKindCommentXMLName) Then
                ' process the entry

                Dim CategoryTermKindProperCase
                CategoryTermKindProperCase = UCase(Left(PPCEntry.CategoryTermKind, 1)) &
                Mid(PPCEntry.CategoryTermKind, 2)

                Dim MatchedSearchString As String
                MatchedSearchString = ""
                If Not IsEmptySearchStrArr(SearchStrArr) Then
                    MatchedSearchString = DoesEntryMatchSearchCriteria(SearchStrArr, CombineSearchStrAsAND,
                                                                       PPCEntry)
                End If

                LogPPCEntry(EntryIndex, PPCEntry, MatchedSearchString)

                If MatchedSearchString <> "" Then 'Entry matched
                    'Later consider creating separate dictionary of such entries
                    'As of now, besides above addition of MatchedSearchString to comment log line, add a line to log file
                    'immediately after comment log line flagging it as matching search criteria. As this will be
                    'in initial part of line, it will be easy to spot when viewing the file.
                    'Later I may create a collection object of all such comments and print them as
                    'matching comments HTML output file.
                    msg = "*****Above logged comment matches search string. Matched string: " & MatchedSearchString
                    Debug.Print(msg)
                    LogFileStream.WriteLine(msg)
                End If

                Dim instrindex
                instrindex = InStr(PPCEntry.Title, "<")
                If (instrindex > 0) Then ' "<" found at 1-based instrindex pos
                    PPCEntry.Title = Left(PPCEntry.Title, instrindex - 1)
                End If

                If (PPCEntry.href <> "") Then
                    ContentHTML = ContentHTML & CategoryTermKindProperCase & " link (URL) on blog: <a href=""" _
                & PPCEntry.href & """>" & PPCEntry.href & "</a><br/><br/>"
                End If

                ContentHTML = ContentHTML & "<b>Comment by " & PPCEntry.AuthorName & "</b> on " _
                & Left(PPCEntry.PublishedDate, 10) & "<br/><br/>"

                ContentHTML = ContentHTML & PPCEntry.Content & "<br/><br/>"
                ContentHTML = ContentHTML &
            "============================End of " & CategoryTermKindProperCase & "==========================<br/><br/>"

                If (PPCEntry.href <> "") Then
                    Dim hrefWithoutQ

                    instrindex = InStr(PPCEntry.href, "?")
                    If (instrindex > 0) Then ' "?" found at 1-based instrindex pos
                        hrefWithoutQ = Left(PPCEntry.href, instrindex - 1)
                    Else
                        hrefWithoutQ = PPCEntry.href
                    End If

                    If dict.ContainsKey(hrefWithoutQ) Then
                        Dim newVal As String
                        newVal = dict.Item(hrefWithoutQ) & ContentHTML
                        dict.Item(hrefWithoutQ) = newVal
                    Else
                        dict.Add(hrefWithoutQ, ContentHTML)
                    End If
                End If

                ContentHTML = ""
                TotalComments += 1
                TotalContentLength += Len(PPCEntry.Content)
            End If ' PPCEntry.CategoryTermKind = CategoryTermKindCommentXMLName
            EntryIndex += 1
        Next
        IterateThroughAllCommentEntries = TotalComments
    End Function

    Sub WriteCommentsDictionaryToFile(dict As Dictionary(Of String, String),
                                      CommentsDictFilePathAndName As String)
        Dim CommentsDictFileStream As System.IO.StreamWriter
        CommentsDictFileStream = My.Computer.FileSystem.OpenTextFileWriter(CommentsDictFilePathAndName,
                                                                           False)

        CommentsDictFileStream.Write("<html><body><meta charset=""utf-8"">" &
            "<h1>Printing Comments Dictionary</h1><h2>Dictionary entries count: " _
            & dict.Count & "</h2><br/>" &
            "==================================================================<br/><br/>")

        Dim key, i
        Dim ContentHTML
        i = 1
        For Each key In dict.Keys
            ContentHTML = "<h2>href key no. " & i & " = " & key & "</h2><br/>" &
        "<b>========================Start of Comments for href=====================</b><br/><br/>" &
        dict.Item(key) & "<br/>"
            CommentsDictFileStream.Write(ContentHTML)
            i += 1
        Next key

        CommentsDictFileStream.Write("<b>==========================End of " _
        & "Dictionary==========================</b><br/></body></html>")
        CommentsDictFileStream.Close

    End Sub


End Module
