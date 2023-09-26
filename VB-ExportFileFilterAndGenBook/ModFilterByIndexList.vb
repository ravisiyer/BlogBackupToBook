Imports System.Configuration
Imports System.IO
Imports System.Xml

Module ModFilterByIndexList
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

    '    Public Const DefaultXMLInputFileBaseName As String = "20230801-ravisiyermisc-blog.xml"
    '    Public Const DefaultIndexListFileBaseName As String = "20230801-ravisiyermisc-blog-IndexList.txt"
    '    Public Const DefaultInputFileDirectory As String _
    '= "D:\Users\Ravi-user\source\repos\VS-ExportFileFilterByIndexList\Data"

    'Public RunFilterProgram As Boolean

    'Sub ShowTopLevelRunOptions()
    '    Dim RunOptionsForm As MainForm
    '    RunOptionsForm = New MainForm
    '    RunOptionsForm.Show()
    '    RunOptionsForm.Close()
    '    'Unload RunOptionsForm
    'End Sub

    Sub DriveDefaultExportFileFilterByIndexList()
        Dim DefaultInputDirectory = ConfigurationManager.AppSettings("DefaultInputDirectory")
        If DefaultInputDirectory <> "" And Not Directory.Exists(DefaultInputDirectory) Then
            MsgBox("Default Input Directory: " & DefaultInputDirectory & " does not exist. " &
                    "DriveDefaultExportFileFilterByIndexList() is aborting!", , MsgBoxTitle)
            Exit Sub
        End If

        'Dim DefaultXMLInputFilePathAndName
        'DefaultXMLInputFilePathAndName = DefaultInputFileDirectory & "\" & DefaultXMLInputFileBaseName
        'Dim DefaultIndexListFilePathAndName
        'DefaultIndexListFilePathAndName = DefaultInputFileDirectory & "\" & DefaultIndexListFileBaseName

        'If Not System.IO.File.Exists(DefaultXMLInputFilePathAndName) Then
        '    MsgBox("Default XML Input File: " & DefaultXMLInputFilePathAndName & " does not exist!" &
        '        vbCrLf & "Aborting Sub DriveDefaultExportFileFilterByIndexList!")
        '    Exit Sub
        'End If
        Dim DefaultXMLInputFilePathAndName As String = GetDefaultXMLInputFilePathAndName()
        If Not File.Exists(DefaultXMLInputFilePathAndName) Then
            MsgBox("Default XML Input File: " & DefaultXMLInputFilePathAndName & " does not exist. " &
                    "DriveDefaultExportFileFilterByIndexList() is aborting!", , MsgBoxTitle)
            Exit Sub
        End If

        'If Not System.IO.File.Exists(DefaultIndexListFilePathAndName) Then
        '    MsgBox("Default Index List Input File: " & DefaultIndexListFilePathAndName & " does not exist!" &
        '        vbCrLf & "Aborting Sub DriveDefaultExportFileFilterByIndexList!")
        '    Exit Sub
        'End If
        Dim DefaultIndexListFilePathAndName As String = GetDefaultIndexListFilePathAndName()
        If Not File.Exists(DefaultIndexListFilePathAndName) Then
            MsgBox("Default Index List File: " & DefaultIndexListFilePathAndName & " does not exist. " &
                    "DriveDefaultExportFileFilterByIndexList() is aborting!", , MsgBoxTitle)
            Exit Sub
        End If

        Dim msg
        msg = ExportFileFilterByIndexList(DefaultXMLInputFilePathAndName, DefaultIndexListFilePathAndName)
        MsgBox(msg)
    End Sub

    Sub DrivePromptExportFileFilterByIndexList()

        Dim XMLInputFileName As String

        Dim fd As New OpenFileDialog With {
            .Title = "Specify Input Blogger Blog Backup/Export XML file"
        }

        Dim DefaultInputDirectory = ConfigurationManager.AppSettings("DefaultInputDirectory")
        If DefaultInputDirectory <> "" And System.IO.Directory.Exists(DefaultInputDirectory) Then
            fd.InitialDirectory = DefaultInputDirectory
        End If
        fd.Filter = "XML Files|*.xml"

        If fd.ShowDialog() = DialogResult.OK Then
            XMLInputFileName = fd.FileName
        Else
            MsgBox("Input XML file Not chosen. Aborting Sub DrivePromptExportFileFilterByIndexList",
                , MsgBoxTitle)
            Exit Sub
        End If

        Dim ParentFolderName
        ParentFolderName = Directory.GetParent(XMLInputFileName).ToString()
        Dim IndexListFileName As String

        fd.Title = "Filter Blog Backup XML file by Index List" & ": Specify Input Index List text file"
        fd.InitialDirectory = ParentFolderName
        fd.Filter = "Txt Files|*.txt"

        If fd.ShowDialog() = DialogResult.OK Then
            IndexListFileName = fd.FileName
        Else
            MsgBox("Index List file not chosen. Aborting Sub DrivePromptExportFileFilterByIndexList",
                , MsgBoxTitle)
            Exit Sub
        End If

        Dim msg
        msg = ExportFileFilterByIndexList(XMLInputFileName, IndexListFileName)
        MsgBox(msg)

    End Sub

    Function ExportFileFilterByIndexList(XMLInputFileName, IndexListFileName) As String

        Dim OutFileName
        Dim LogFileName
        OutFileName = Left(XMLInputFileName, Len(XMLInputFileName) - 4) & "ixfilt.xml"
        LogFileName = Left(XMLInputFileName, Len(XMLInputFileName) - 4) & "ixfiltlog.txt"

        Dim StartTime, EndTime As String
        StartTime = Now

        Dim LogFileStream As System.IO.StreamWriter
        LogFileStream = My.Computer.FileSystem.OpenTextFileWriter(LogFileName, False) 'Overwrite not append

        Dim msg As String
        msg = "Function ExportFileFilterByIndexList start date & time is: " & StartTime
        Debug.Print(msg)
        LogFileStream.WriteLine(msg)

        msg = "Arguments to Sub. ExportFileFilterByIndexList are: XMLInputFileName: " &
            vbCrLf & XMLInputFileName & vbCrLf & "IndexListFileName: " &
            IndexListFileName
        Debug.Print(msg)
        LogFileStream.WriteLine(msg)

        Dim IndexListFileStream As System.IO.StreamReader
        IndexListFileStream = My.Computer.FileSystem.OpenTextFileReader(IndexListFileName)

        Dim OutFileStream As System.IO.StreamWriter
        OutFileStream = My.Computer.FileSystem.OpenTextFileWriter(OutFileName, False)

        Dim XDoc As XmlDocument
        XDoc = New XmlDocument()

        ' https://stackoverflow.com/questions/24734/selectnodes-not-working-on-stackoverflow-feed
        Dim nsmgr As XmlNamespaceManager
        nsmgr = New XmlNamespaceManager(XDoc.NameTable)
        nsmgr.AddNamespace("atom", "http://www.w3.org/2005/Atom")

        XDoc.Load(XMLInputFileName)

        Dim Node As XmlNode
        Node = XDoc.FirstChild 'xml
        msg = Node.OuterXml
        OutFileStream.WriteLine(msg)
        Node = Node.NextSibling ' xml-stylesheet
        msg = Node.OuterXml
        OutFileStream.WriteLine(msg)

        Node = Node.NextSibling ' feed
        Dim ix

        ix = InStr(Node.OuterXml, "<entry")
        If (ix <= 0) Then
            'Something is wrong. Abort.
            msg = "Missing '<entry>' in feed XML! Aborting Function ExportFileFilterByIndexList!"
            Debug.Print(msg)
            LogFileStream.WriteLine(msg)
            MsgBox(msg, , MsgBoxTitle)
            ExportFileFilterByIndexList = msg
            Exit Function
        End If
        msg = Left(Node.OuterXml, ix - 1)
        OutFileStream.WriteLine(msg)

        Dim Entries As XmlNodeList

        Entries = XDoc.DocumentElement.SelectNodes("atom:entry", nsmgr)

        If (Entries.Count <= 0) Then
            'Something is wrong. Abort.
            msg = "SelectNodes returned 0 entries! Aborting Function ExportFileFilterByIndexList!"
            Debug.Print(msg)
            LogFileStream.WriteLine(msg)
            MsgBox(msg, , MsgBoxTitle)
            ExportFileFilterByIndexList = msg
            Exit Function
        End If

        Dim Entry As XmlNode

        Dim NumEntriesWritten
        NumEntriesWritten = 0

        Dim EntryIndex
        EntryIndex = 1

        Dim TextLine As String
        Dim PPCIndexNum As Integer
        PPCIndexNum = 0
        Dim FirstField As String
        Dim TabStr As String
        TabStr = Chr(9)

        With IndexListFileStream
            Do Until .EndOfStream
                TextLine = .ReadLine
                ix = InStr(TextLine, TabStr)
                If (ix > 0) Then
                    FirstField = Left(TextLine, ix - 1)
                    If IsNumeric(FirstField) Then
                        PPCIndexNum = FirstField
                        If PPCIndexNum > Entries.Count Then
                            'Abort
                            msg = "Index in Index List file: " & PPCIndexNum & " > Count of Entries: " &
                                Entries.Count & "!" & vbCrLf & "Aborting Function ExportFileFilterByIndexList!"
                            Debug.Print(msg)
                            LogFileStream.WriteLine(msg)
                            MsgBox(msg, , MsgBoxTitle)
                            ExportFileFilterByIndexList = msg
                            Exit Function
                        End If
                        Entry = Entries(PPCIndexNum - 1)
                        If Not IsEntryComment(Entry, nsmgr) Then 'All comments are picked up later on. Don't duplicate comments
                            OutFileStream.WriteLine(Entry.OuterXml)
                            NumEntriesWritten += 1
                        End If
                    End If
                End If
            Loop
            .Close()
        End With

        ' Now pull in all comments from entries and write to output file
        For Each Entry In Entries
            If IsEntryComment(Entry, nsmgr) Then
                OutFileStream.WriteLine(Entry.OuterXml)
                NumEntriesWritten += 1
            End If 'Skip all entries other than comments and matching post and page entries
            EntryIndex += 1
        Next

        '    This Hard-coded solution seems OK as it is only the end tag
        msg = "</feed>"
        OutFileStream.WriteLine(msg)

        OutFileStream.Close()

        msg = "Sub. ExportFileFilterByIndexList finished execution." & vbCrLf & vbCrLf &
        "Output file: " & vbCrLf & OutFileName & vbCrLf &
        "Number of entries written to output file:" & NumEntriesWritten & vbCrLf & vbCrLf

        EndTime = Now
        msg = msg & "Date & time now is: " & Now & vbCrLf &
        "Time taken for this macro sub. execution (in seconds): " _
        & DateDiff("s", StartTime, EndTime) & " (excludes dialog interaction by driver (invoker) sub)"
        Debug.Print(msg)
        LogFileStream.WriteLine(msg)

        LogFileStream.Close()

        ExportFileFilterByIndexList = msg

    End Function

    Function IsEntryComment(Entry As XmlNode, nsmgr As XmlNamespaceManager) As Boolean
        Dim CategoryNode As XmlNode
        Dim termAttr As XmlNode

        CategoryNode = Entry.SelectSingleNode("atom:category", nsmgr)
        termAttr = CategoryNode.Attributes.GetNamedItem("term")
        If (Right(termAttr.Value, 7) = "comment") Then
            IsEntryComment = True
        Else
            IsEntryComment = False
        End If
    End Function

    Function GetDefaultIndexListFilePathAndName() As String
        Dim DefaultIndexListFilePathAndName As String
        Dim DefaultInputDirectory As String
        DefaultInputDirectory = ConfigurationManager.AppSettings("DefaultInputDirectory")
        If (DefaultInputDirectory = "") Then
            DefaultIndexListFilePathAndName = ConfigurationManager.AppSettings("DefaultIndexListFileName")
        Else
            DefaultIndexListFilePathAndName = DefaultInputDirectory & "\" &
                ConfigurationManager.AppSettings("DefaultIndexListFileName")
        End If
        GetDefaultIndexListFilePathAndName = DefaultIndexListFilePathAndName
    End Function


End Module
