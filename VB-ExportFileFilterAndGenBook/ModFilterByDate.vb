Imports System.Configuration
Imports System.IO
Imports System.Xml

Module ModFilterByDate
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

    Public RunFilterProgram As Boolean
    Public Const InitializedDateStr = "0001/01/01"
    '
    Structure MatchCriteriaType
        '    Title As String
        '    Content As String
        '    href As String
        '    AuthorName As String
        Public StartPublishedDateStr As String
        Public StartPublishedDate As Date
        Public EndPublishedDateStr As String
        Public EndPublishedDate As Date
        Public StartUpdatedDateStr As String
        Public StartUpdatedDate As Date
        Public EndUpdatedDateStr As String
        Public EndUpdatedDate As Date
    End Structure

    Enum DateMatchType
        BothStartAndEnd = 1
        StartOnly = 2
        EndOnly = 3
        None = 4
    End Enum

    Sub DriveDefaultFilterBlogExportFile()
        Dim DefaultInputDirectory = ConfigurationManager.AppSettings("DefaultInputDirectory")
        If DefaultInputDirectory <> "" And Not Directory.Exists(DefaultInputDirectory) Then
            MsgBox("Default Input Directory: " & DefaultInputDirectory & " does not exist. " &
                    "DriveDefaultFilterBlogExportFile() is aborting!", , MsgBoxTitle)
            Exit Sub
        End If

        Dim DefaultXMLInputFilePathAndName As String = GetDefaultXMLInputFilePathAndName()
        If Not File.Exists(DefaultXMLInputFilePathAndName) Then
            MsgBox("Default XML Input File: " & DefaultXMLInputFilePathAndName & " does not exist. " &
                    "DriveDefaultBlogExportFileSearchToBook() is aborting!", , MsgBoxTitle)
            Exit Sub
        End If

        Dim MatchCriteria As MatchCriteriaType
        MatchCriteria = GetInitializedMatchCriteria()
        Dim msg
        msg = FilterBlogExportFile(DefaultXMLInputFilePathAndName, MatchCriteria)
        MsgBox(msg, , MsgBoxTitle)
    End Sub

    Sub DrivePromptFilterBlogExportFile()

        Dim XMLInputFileName As String
        XMLInputFileName = ""

        Dim fd As New OpenFileDialog()
        If LastXMLInputFilePath = "" Then
            Dim DefaultInputDirectory As String
            DefaultInputDirectory = ConfigurationManager.AppSettings("DefaultInputDirectory")
            If DefaultInputDirectory <> "" And System.IO.Directory.Exists(DefaultInputDirectory) Then
                fd.InitialDirectory = DefaultInputDirectory
            End If
        Else
            fd.InitialDirectory = LastXMLInputFilePath
        End If
        fd.Title = "Filter Blog Backup XML file by Date" & ": Select input XML export file"
        fd.Filter = "XML Files|*.xml"
        If fd.ShowDialog() = DialogResult.OK Then
            XMLInputFileName = fd.FileName
            LastXMLInputFilePath = Directory.GetParent(XMLInputFileName).ToString()
        Else
            MsgBox("Input XML file Not chosen. Aborting Sub DrivePromptFilterBlogExportFile",
                , MsgBoxTitle)
            Exit Sub
        End If

        Dim ContinueRun As Boolean
        ContinueRun = True

        Dim msg, Style, response
        Dim MatchCriteria As MatchCriteriaType
        MatchCriteria = GetInitializedMatchCriteria()

        Do While ContinueRun
            RunFilterProgram = False
            FilterDateRangesForm.ShowDialog()

            If RunFilterProgram Then
                With FilterDateRangesForm
                    MatchCriteria.StartPublishedDateStr = .StartPublishedDateTB.Text
                    MatchCriteria.EndPublishedDateStr = .EndPublishedDateTB.Text
                    MatchCriteria.StartUpdatedDateStr = .StartUpdatedDateTB.Text
                    MatchCriteria.EndUpdatedDateStr = .EndUpdatedDateTB.Text
                End With
                If MatchCriteria.StartPublishedDateStr <> "" Then
                    MatchCriteria.StartPublishedDate = CDate(MatchCriteria.StartPublishedDateStr)
                End If
                If MatchCriteria.EndPublishedDateStr <> "" Then
                    MatchCriteria.EndPublishedDate = CDate(MatchCriteria.EndPublishedDateStr)
                End If
                If MatchCriteria.StartUpdatedDateStr <> "" Then
                    MatchCriteria.StartUpdatedDate = CDate(MatchCriteria.StartUpdatedDateStr)
                End If
                If MatchCriteria.EndUpdatedDateStr <> "" Then
                    MatchCriteria.EndUpdatedDate = CDate(MatchCriteria.EndUpdatedDateStr)
                End If
            Else
                MsgBox("User clicked Cancel command button or closed form window. " &
                    "Aborting Sub DrivePromptFilterBlogExportFile!", , MsgBoxTitle)
                Exit Sub
            End If
            msg = FilterBlogExportFile(XMLInputFileName, MatchCriteria)
            msg = msg & vbCrLf & vbCrLf & "Do you want to run Sub FilterBlogExportFile" _
                & " again on same input file?"
            Style = vbYesNo Or vbDefaultButton2    ' Define buttons.
            response = MsgBox(msg, Style, MsgBoxTitle)
            If response = vbNo Then    ' User chose No.
                ContinueRun = False
            End If
        Loop

        FilterDateRangesForm.Close()
    End Sub

    Function FilterBlogExportFile(XMLInputFileName As String, MatchCriteria As MatchCriteriaType) As String
        Dim OutFileName
        Dim LogFileName
        OutFileName = Left(XMLInputFileName, Len(XMLInputFileName) - 4) & "filt.xml"
        LogFileName = Left(XMLInputFileName, Len(XMLInputFileName) - 4) & "filtlog.txt"
        Dim StartTime, EndTime As String
        StartTime = Now

        LogFileStream = My.Computer.FileSystem.OpenTextFileWriter(LogFileName, False) 'Overwrite not append

        Dim msg As String
        msg = "Function FilterBlogExportFile start date & time is: " & StartTime
        Debug.Print(msg)
        LogFileStream.WriteLine(msg)


        msg = "Arguments to Fn. FilterBlogExportFile are: XMLInputFilePathAndName = " &
            vbCrLf & XMLInputFileName & vbCrLf
        Debug.Print(msg)
        LogFileStream.WriteLine(msg)

        LogMatchCriteria(MatchCriteria)

        Dim PublishedDateMatchT As DateMatchType
        PublishedDateMatchT = GetPublishedDateMatchT(MatchCriteria) ' GetPublishedDateMatchT() writes to Debug & logfile

        Dim UpdatedDateMatchT As DateMatchType
        UpdatedDateMatchT = GetUpdatedDateMatchT(MatchCriteria) ' GetUpdatedDateMatchT() writes to Debug & logfile

        Dim OutFileStream As System.IO.StreamWriter

        OutFileStream = My.Computer.FileSystem.OpenTextFileWriter(OutFileName, False)

        Dim XDoc As XmlDocument
        XDoc = New XmlDocument()

        ' https://stackoverflow.com/questions/24734/selectnodes-not-working-on-stackoverflow-feed
        Dim nsmgr As XmlNamespaceManager
        nsmgr = New XmlNamespaceManager(XDoc.NameTable)
        nsmgr.AddNamespace("atom", "http://www.w3.org/2005/Atom")

        XDoc.Load(XMLInputFileName)
        'msg = "XmlDocument (var XDoc) created and file: '" & XMLInputFileName &
        '    "' loaded into it."
        'Debug.Print(msg)
        'LogFileStream.WriteLine(msg)


        Dim Node As XmlNode
        Node = XDoc.FirstChild 'xml
        msg = Node.OuterXml
        OutFileStream.WriteLine(msg)
        Node = Node.NextSibling ' xml-stylesheet
        msg = Node.OuterXml
        OutFileStream.WriteLine(msg)

        Node = Node.NextSibling ' feed

        Dim ix
        ix = InStr(Node.OuterXml, "<entry>")
        If (ix <= 0) Then
            'Something is wrong. Abort.
            msg = "Missing '<entry>' in feed XML! Aborting Function FilterBlogExportFile!"
            Debug.Print(msg)
            LogFileStream.WriteLine(msg)
            MsgBox(msg, , MsgBoxTitle)
            FilterBlogExportFile = msg
            Exit Function
        End If
        msg = Left(Node.OuterXml, ix - 1)
        OutFileStream.WriteLine(msg)
        msg = GetRunHeaderXML(XMLInputFileName, MatchCriteria, StartTime)
        ' If this XML is inserted  before <feed> it breaks BlogExportFileSearchToBook as this inserted xml seems to
        ' make feed element no longer the root element. So /feed/xxx does not work.
        ' Manually moved this XML to end of file just before </feed>. That worked.
        ' Don't know if this approach of writing the RunHeader XML withing <feed> data just before the
        ' the first <entry> element is appropriate. But I think it will work and not interfere with other code that
        ' uses "/feed/entry" to retrieve the entry elements (entries)
        OutFileStream.WriteLine(msg)

        Dim Entries As XmlNodeList

        Entries = XDoc.DocumentElement.SelectNodes("atom:entry", nsmgr)

        If (Entries.Count <= 0) Then
            'Something is wrong. Abort.
            msg = "SelectNodes returned 0 entries! Aborting Function FilterBlogExportFile!"
            Debug.Print(msg)
            LogFileStream.WriteLine(msg)
            MsgBox(msg, , MsgBoxTitle)
            FilterBlogExportFile = msg
            Exit Function
        End If

        Dim Entry As XmlNode

        Dim PPCMatchesCriteria As Boolean ' PPC stands for Page, Post and/or Comment

        Dim NumEntriesWritten
        NumEntriesWritten = 0

        Dim EntryIndex
        EntryIndex = 1

        WriteEntriesHeaderToLog(Entries.Count)

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
            If PPCEntry.CategoryTermKind = "post" Or PPCEntry.CategoryTermKind = "page" Then
                PPCMatchesCriteria = DoesPPCMatchCriteria(Entry, nsmgr, PublishedDateMatchT, UpdatedDateMatchT,
                MatchCriteria)
                If PPCMatchesCriteria Then
                    OutFileStream.WriteLine(Entry.OuterXml)
                    LogPPCEntry(EntryIndex, PPCEntry)
                    NumEntriesWritten += 1
                End If
            ElseIf PPCEntry.CategoryTermKind = "comment" Then 'copy all comment entries irrespective of PublishedDate
                OutFileStream.WriteLine(Entry.OuterXml)
                LogPPCEntry(EntryIndex, PPCEntry)
                NumEntriesWritten += 1
            End If 'Skip all entries other than comments and matching post and page entries
            EntryIndex += 1
        Next

        '    This Hard-coded solution seems OK as it is only the end tag
        msg = "</feed>"
        OutFileStream.WriteLine(msg)

        '    Set XDoc = Nothing
        OutFileStream.Close()

        msg = "Fn. FilterBlogExportFile finished execution." & vbCrLf & vbCrLf &
            "Output file: " & vbCrLf & OutFileName & vbCrLf &
            "Number of entries written to output file:" & NumEntriesWritten & vbCrLf & vbCrLf

        EndTime = Now
        msg = msg & "Date & time now is: " & Now & vbCrLf &
            "Time taken for this function execution (in seconds): " _
            & DateDiff("s", StartTime, EndTime) &
            " (excludes dialog interaction by driver (invoker) sub)"
        Debug.Print(msg)
        LogFileStream.WriteLine(msg)

        LogFileStream.Close()

        FilterBlogExportFile = msg

    End Function

    Function GetPublishedDateMatchT(MatchCriteria As MatchCriteriaType) As DateMatchType
        Dim PublishedDateMatchT As DateMatchType

        If MatchCriteria.StartPublishedDateStr = "" Then
            If MatchCriteria.EndPublishedDateStr = "" Then
                PublishedDateMatchT = DateMatchType.None
            Else
                PublishedDateMatchT = DateMatchType.EndOnly
            End If
        Else
            If MatchCriteria.EndPublishedDateStr = "" Then
                PublishedDateMatchT = DateMatchType.StartOnly
            Else
                PublishedDateMatchT = DateMatchType.BothStartAndEnd
            End If
        End If

        Dim msg
        msg = "PublishedDateMatchT = " & PublishedDateMatchT &
        " [BothStartAndEnd = 1, StartOnly = 2, EndOnly = 3, None = 4]"
        Debug.Print(msg)
        LogFileStream.WriteLine(msg)

        GetPublishedDateMatchT = PublishedDateMatchT
    End Function

    Function GetUpdatedDateMatchT(MatchCriteria As MatchCriteriaType) As DateMatchType
        Dim UpdatedDateMatchT As DateMatchType

        If MatchCriteria.StartUpdatedDateStr = "" Then
            If MatchCriteria.EndUpdatedDateStr = "" Then
                UpdatedDateMatchT = DateMatchType.None
            Else
                UpdatedDateMatchT = DateMatchType.EndOnly
            End If
        Else
            If MatchCriteria.EndUpdatedDateStr = "" Then
                UpdatedDateMatchT = DateMatchType.StartOnly
            Else
                UpdatedDateMatchT = DateMatchType.BothStartAndEnd
            End If
        End If

        Dim msg
        msg = "UpdatedDateMatchT = " & UpdatedDateMatchT &
        " [BothStartAndEnd = 1, StartOnly = 2, EndOnly = 3, None = 4]"
        Debug.Print(msg)
        LogFileStream.WriteLine(msg)

        GetUpdatedDateMatchT = UpdatedDateMatchT
    End Function

    Function DoesPPCMatchCriteria(Entry As XmlNode, nsmgr As XmlNamespaceManager, PublishedDateMatchT As DateMatchType,
    UpdatedDateMatchT As DateMatchType, MatchCriteria As MatchCriteriaType) As Boolean

        Dim PPCMatchesCriteria As Boolean
        Dim PublishedDateText
        Dim PublishedDate As Date
        Dim UpdatedDateText
        Dim UpdatedDate As Date

        If PublishedDateMatchT = DateMatchType.None And UpdatedDateMatchT = DateMatchType.None Then
            ' No match criteria to check
            DoesPPCMatchCriteria = True
            Exit Function
        End If

        PPCMatchesCriteria = False

        PublishedDateText = GetPublishedDateFromEntry(Entry, nsmgr)
        PublishedDate = CDate(Left(PublishedDateText, 10))

        Select Case PublishedDateMatchT
            Case DateMatchType.BothStartAndEnd
                If PublishedDate >= MatchCriteria.StartPublishedDate And
            PublishedDate <= MatchCriteria.EndPublishedDate Then
                    PPCMatchesCriteria = True
                End If
            Case DateMatchType.StartOnly
                If PublishedDate >= MatchCriteria.StartPublishedDate Then
                    PPCMatchesCriteria = True
                End If
            Case DateMatchType.EndOnly
                If PublishedDate <= MatchCriteria.EndPublishedDate Then
                    PPCMatchesCriteria = True
                End If
                '    Case DateMatchType.None ' Commented out after adding Updated Date range match
                '        PPCMatchesCriteria = True
        End Select

        ' Published Date and Updated Date ranges when specified together work as OR not AND
        ' So if Published date matches Published Date range we can set return value to true
        ' and exit the function now itself.
        ' We need to check Updated Date range match only if Published Date range does not match
        If PPCMatchesCriteria Then
            DoesPPCMatchCriteria = True
            Exit Function
        End If

        UpdatedDateText = GetUpdatedDateFromEntry(Entry, nsmgr)
        UpdatedDate = CDate(Left(UpdatedDateText, 10))

        Select Case UpdatedDateMatchT
            Case DateMatchType.BothStartAndEnd
                If UpdatedDate >= MatchCriteria.StartUpdatedDate And
            UpdatedDate <= MatchCriteria.EndUpdatedDate Then
                    PPCMatchesCriteria = True
                End If
            Case DateMatchType.StartOnly
                If UpdatedDate >= MatchCriteria.StartUpdatedDate Then
                    PPCMatchesCriteria = True
                End If
            Case DateMatchType.EndOnly
                If UpdatedDate <= MatchCriteria.EndUpdatedDate Then
                    PPCMatchesCriteria = True
                End If
                '    Case DateMatchType.None
                '        PPCMatchesCriteria = True
        End Select

        DoesPPCMatchCriteria = PPCMatchesCriteria
    End Function

    Function GetRunHeaderXML(XMLInputFileName As String, MatchCriteria As MatchCriteriaType,
                             StartTime As String) As String
        Dim XMLOutput As String

        XMLOutput = "<VBDotNet-FilterXMLExportFile>"

        XMLOutput = XMLOutput & "<RunDateTime>" & StartTime & "</RunDateTime>"

        XMLOutput = XMLOutput & "<InputFileName>" & XMLInputFileName & "</InputFileName>"

        XMLOutput &= "<MatchCriteria>"
        If MatchCriteria.StartPublishedDateStr <> "" Then
            XMLOutput = XMLOutput & "<StartPublishedDate>" & MatchCriteria.StartPublishedDateStr &
        "</StartPublishedDate>"
        End If
        If MatchCriteria.EndPublishedDateStr <> "" Then
            XMLOutput = XMLOutput & "<EndPublishedDate>" & MatchCriteria.EndPublishedDateStr &
        "</EndPublishedDate>"
        End If
        If MatchCriteria.StartUpdatedDateStr <> "" Then
            XMLOutput = XMLOutput & "<StartUpdatedDate>" & MatchCriteria.StartUpdatedDateStr &
        "</StartUpdatedDate>"
        End If
        If MatchCriteria.EndUpdatedDateStr <> "" Then
            XMLOutput = XMLOutput & "<EndUpdatedDate>" & MatchCriteria.EndUpdatedDateStr &
        "</EndUpdatedDate>"
        End If
        XMLOutput &= "</MatchCriteria>"

        XMLOutput &= "</VBDotNet-FilterXMLExportFile>"

        GetRunHeaderXML = XMLOutput
    End Function


    Function GetInitializedMatchCriteria() As MatchCriteriaType
        Dim MatchCriteria As MatchCriteriaType
        MatchCriteria.StartPublishedDateStr = ""
        MatchCriteria.StartPublishedDate = InitializedDateStr
        MatchCriteria.EndPublishedDateStr = ""
        MatchCriteria.EndPublishedDate = InitializedDateStr

        MatchCriteria.StartUpdatedDateStr = ""
        MatchCriteria.StartUpdatedDate = InitializedDateStr
        MatchCriteria.EndUpdatedDateStr = ""
        MatchCriteria.EndUpdatedDate = InitializedDateStr

        GetInitializedMatchCriteria = MatchCriteria
    End Function

    Sub LogMatchCriteria(MatchCriteria As MatchCriteriaType)
        Dim msg
        msg = "MatchCriteria.StartPublishedDateStr = " & MatchCriteria.StartPublishedDateStr & vbCrLf
        msg = msg & "MatchCriteria.StartPublishedDate = " & MatchCriteria.StartPublishedDate & vbCrLf
        msg = msg & "MatchCriteria.EndPublishedDateStr = " & MatchCriteria.EndPublishedDateStr & vbCrLf
        msg = msg & "MatchCriteria.EndPublishedDate = " & MatchCriteria.EndPublishedDate & vbCrLf
        Debug.Print(msg)
        LogFileStream.WriteLine(msg)

        msg = "MatchCriteria.StartUpdatedDateStr = " & MatchCriteria.StartUpdatedDateStr & vbCrLf
        msg = msg & "MatchCriteria.StartUpdatedDate = " & MatchCriteria.StartUpdatedDate & vbCrLf
        msg = msg & "MatchCriteria.EndUpdatedDateStr = " & MatchCriteria.EndUpdatedDateStr & vbCrLf
        msg = msg & "MatchCriteria.EndUpdatedDate = " & MatchCriteria.EndUpdatedDate & vbCrLf
        Debug.Print(msg)
        LogFileStream.WriteLine(msg)

    End Sub

End Module
