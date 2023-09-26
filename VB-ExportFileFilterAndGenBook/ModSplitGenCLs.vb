Imports System.Configuration
Imports System.IO
Imports HtmlAgilityPack

Module ModSplitGenCLs
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

    ' 5 MB Default SplitSize. CLng seems to force long computation else get overflow error
    Public Const DefaultSplitSize As Long = 5242880 'CLng(5) * 1024 * 1024 gives compile error
    'So am giving the resultant number for 5 MB
    ' 50 KB (roughly 1 % of SplitSize) default overflow read chunk size till end of post/page
    Public Const DefaultSplitOverflowReadSize As Long = 51200 ' CLng(5) * 10 * 1024

    Public Const DefaultEndOfPostMarker = "End of Post==========================<br/><br/>"
    Public Const DefaultEndOfPageMarker = "End of Page==========================<br/><br/>"

    Public Const SplitFileNameSuffix = "-Split"

    Public Const DefaultContentsLinksInsertMarker = "<h1 id="
    Public Const ContentsLinksInsertedOutputFileBaseSuffix As String = "wCLs"

    Public RunProgram As Boolean


    Sub DriveDefaultSplitBBAndGenCLs()

        Dim DefaultInputDirectory = ConfigurationManager.AppSettings("DefaultInputDirectory")
        If DefaultInputDirectory <> "" And Not Directory.Exists(DefaultInputDirectory) Then
            MsgBox("Default Input Directory: " & DefaultInputDirectory & " does not exist. " &
                    "DriveDefaultSplitBBAndGenCLs() is aborting!", , MsgBoxTitle)
            Exit Sub
        End If

        Dim DefaultHTMLInputFilePathAndName = GetDefaultHTMLInputFilePathAndName()
        If Not File.Exists(DefaultHTMLInputFilePathAndName) Then
            MsgBox("Default HTML Input File: " & DefaultHTMLInputFilePathAndName & " does not exist. " &
                    "DriveDefaultSplitBBAndGenCLs() is aborting!", , MsgBoxTitle)
            Exit Sub
        End If

        Dim LogFilePathAndName = Left(DefaultHTMLInputFilePathAndName,
                                  Len(DefaultHTMLInputFilePathAndName) - 5) & "-SBBGCLsLog.txt"

        LogFileStream = My.Computer.FileSystem.OpenTextFileWriter(LogFilePathAndName, False) 'Overwrite not append

        SplitBBAndGenCLs(DefaultHTMLInputFilePathAndName, DefaultSplitSize, DefaultEndOfPostMarker, DefaultEndOfPageMarker,
        True, DefaultContentsLinksInsertMarker) 'True for GenerateContentsLinks

        LogFileStream.Close()

    End Sub

    Sub DriveTestSplitBBAndGenCLs()

        Dim TestInputDirectory

        TestInputDirectory = ConfigurationManager.AppSettings("TestInputDirectory")

        If Not Directory.Exists(TestInputDirectory) Then
            MsgBox("TestInputDirectory: " & TestInputDirectory & " does not exist. " &
                   "DriveTestSplitBBAndGenCLs is aborting!", , MsgBoxTitle)
            Exit Sub
        End If
        Dim LogFilePathAndName
        LogFilePathAndName = TestInputDirectory & "\" & "test-SBBGCLsLog.txt"

        LogFileStream = My.Computer.FileSystem.OpenTextFileWriter(LogFilePathAndName, False) 'Overwrite not append

        Dim TestInputFilePathAndName As String
        TestInputFilePathAndName = TestInputDirectory & "\" & "test.html"
        SplitBBAndGenCLs(TestInputFilePathAndName, 50, DefaultEndOfPostMarker, DefaultEndOfPageMarker,
            True, DefaultContentsLinksInsertMarker) 'True for GenerateContentsLinks

        TestInputFilePathAndName = TestInputDirectory & "\" & "test.xml"
        SplitBBAndGenCLs(TestInputFilePathAndName, 50, "</entry>", DefaultEndOfPageMarker,
        False, DefaultContentsLinksInsertMarker) 'False for GenerateContentsLinks

        TestInputFilePathAndName = TestInputDirectory & "\" & "test.txt"
        SplitBBAndGenCLs(TestInputFilePathAndName, 50, DefaultEndOfPostMarker, DefaultEndOfPageMarker,
        False, DefaultContentsLinksInsertMarker) 'False for GenerateContentsLinks

        LogFileStream.Close()

    End Sub

    Sub DrivePromptInputSplitBBAndGenCLs()

        Dim InputFilePathAndName As String
        InputFilePathAndName = ""
        Dim fd As New OpenFileDialog With {
            .Filter = "HTML Files|*.html|XML Files|*.xml|Text Files|*.txt"
        }

        If LastXMLInputFilePath = "" Then
            Dim DefaultInputDirectory = ConfigurationManager.AppSettings("DefaultInputDirectory")
            If System.IO.Directory.Exists(DefaultInputDirectory) Then
                fd.InitialDirectory = DefaultInputDirectory
            End If
        Else
            fd.InitialDirectory = LastXMLInputFilePath
        End If
        fd.Title = "Generate Contents Links And Split HTML Blogbook" & ": Select input file"
        If fd.ShowDialog() = DialogResult.OK Then
            InputFilePathAndName = fd.FileName
        Else
            MsgBox("Input file not chosen. Aborting Sub DrivePromptInputSplitBBAndGenCLs", , MsgBoxTitle)
            Exit Sub
        End If

        Dim Message
        Dim SplitSize As Long
        Dim EndOfPostMarker As String
        Dim GenerateContentsLinks As Boolean
        Dim ContentsLinksInsertMarker As String

        RunProgram = False

        Dim ix

        ix = InStr(InputFilePathAndName, ".html")
        If (ix > 0) Then ' we have an html file; does user want contents links to be generated and inserted?
            ParamForm.CBGenerateContentsLinks.Enabled = True
            ParamForm.CBGenerateContentsLinks.Checked = True
            ParamForm.TBContentsLinksInsertMarker.Enabled = True
            ParamForm.TBContentsLinksInsertMarker.Text = DefaultContentsLinksInsertMarker
            GenerateContentsLinks = True
        Else
            ParamForm.CBGenerateContentsLinks.Enabled = False
            ParamForm.CBGenerateContentsLinks.Checked = False
            ParamForm.TBContentsLinksInsertMarker.Enabled = False
            GenerateContentsLinks = False
        End If

        ContentsLinksInsertMarker = ""
        ParamForm.TBEndOfPostMarker.Text = DefaultEndOfPostMarker
        ParamForm.ShowDialog()
        If RunProgram Then
            With ParamForm
                If .TBSplitSize.Text = "" Then
                    SplitSize = 0
                ElseIf IsNumeric(.TBSplitSize.Text) Then
                    SplitSize = .TBSplitSize.Text
                Else
                    SplitSize = 0
                End If
                EndOfPostMarker = .TBEndOfPostMarker.Text
                If .CBGenerateContentsLinks.Enabled = True Then
                    GenerateContentsLinks = .CBGenerateContentsLinks.Checked
                    ContentsLinksInsertMarker = .TBContentsLinksInsertMarker.Text
                End If
            End With
            If SplitSize < 0 Then
                SplitSize = 0
            End If
            If EndOfPostMarker = "" Then
                EndOfPostMarker = DefaultEndOfPostMarker
            End If
            If GenerateContentsLinks Then
                If ContentsLinksInsertMarker = "" Then
                    ContentsLinksInsertMarker = DefaultContentsLinksInsertMarker
                End If
            End If
        Else
            MsgBox("User clicked Cancel command button or closed form window." &
                   " Aborting DrivePromptInputSplitBBAndGenCLs!", , MsgBoxTitle)
            Exit Sub
        End If

        Dim LogFilePath, LogFileBaseName, LogFilePathAndName
        LogFilePath = Directory.GetParent(InputFilePathAndName).ToString()
        LogFileBaseName = Path.GetFileNameWithoutExtension(InputFilePathAndName)
        LogFilePathAndName = LogFilePath & "\" & LogFileBaseName & "-SBBGCLsLog.txt"


        LogFileStream = My.Computer.FileSystem.OpenTextFileWriter(LogFilePathAndName, False) 'Overwrite not append

        If (SplitSize > 0) Then
            SplitBBAndGenCLs(InputFilePathAndName, SplitSize, EndOfPostMarker, DefaultEndOfPageMarker,
            GenerateContentsLinks, ContentsLinksInsertMarker)
        ElseIf GenerateContentsLinks Then
            Dim CLOutputFilePathAndName
            Dim StartTime, EndTime As String
            StartTime = Now
            Message = "Date & time just before calling function GenerateAndInsertContentsLinksIntoFile Is: " _
            & StartTime
            Debug.Print(Message)
            LogFileStream.WriteLine(Message)

            CLOutputFilePathAndName = GenerateAndInsertContentsLinksIntoFile(InputFilePathAndName,
            ContentsLinksInsertMarker)

            Message = "GenerateAndInsertContentsLinksIntoFile function was invoked for input file: " &
            InputFilePathAndName & " And ContentsLinksInsertMarker specified was: " & ContentsLinksInsertMarker &
            "." & vbCrLf & vbCrLf
            If CLOutputFilePathAndName = "" Then
                Message &= "GenerateAndInsertContentsLinksIntoFile failed."
            Else
                Message = Message & "GenerateAndInsertContentsLinksIntoFile succeeded. " & vbCrLf &
            "Output file with contents links created with path And name: " &
            CLOutputFilePathAndName & " ." & vbCrLf & vbCrLf
            End If

            EndTime = Now
            Message = Message + "Date & time now Is: " & Now & vbCrLf &
            "Time taken for GenerateAndInsertContentsLinksIntoFile function execution (in seconds): " _
            & DateDiff("s", StartTime, EndTime)
            Debug.Print(Message)
            LogFileStream.WriteLine(Message)
            MsgBox(Message, , MsgBoxTitle)
        End If

        LogFileStream.Close()

    End Sub

    Sub SplitBBAndGenCLs(InputFilePathAndName As String, SplitSize As Long, EndOfPostMarker As String,
        EndOfPageMarker As String, GenerateContentsLinks As Boolean, ContentsLinksInsertMarker As String)

        Dim InputFileData As String
        Dim InputTSO As System.IO.StreamReader
        Dim OutputTSO As System.IO.StreamWriter
        Dim msg

        Dim CurrentSplitFilePathAndName

        Dim StartTime, EndTime As String
        StartTime = Now

        msg = "Function SplitBBAndGenCLs start date & time Is: " & StartTime
        Debug.Print(msg)
        LogFileStream.WriteLine(msg)

        InputTSO = My.Computer.FileSystem.OpenTextFileReader(InputFilePathAndName)

        Dim SplitInProgress As Boolean
        SplitInProgress = True
        Dim SplitNum, OutputSplitFilesWritten
        SplitNum = 1
        OutputSplitFilesWritten = 0

        Dim CurrentSplitCharsWritten, TotalSplitsCharsWritten As Long
        TotalSplitsCharsWritten = 0
        CurrentSplitCharsWritten = 0

        Dim SplitFilePathAndNameBase As String
        SplitFilePathAndNameBase = Directory.GetParent(InputFilePathAndName).ToString() & "\" &
        Path.GetFileNameWithoutExtension(InputFilePathAndName) & SplitFileNameSuffix

        Dim InputFileExt As String
        InputFileExt = Path.GetExtension(InputFilePathAndName)


        InputFileData = InputTSO.ReadToEnd() ' Is able to read 16 M chars UNICODE file 
        '        'But don't know at what bigger size it will trip up
        InputTSO.Close()

        Dim InputFileDataLen As Long
        InputFileDataLen = Len(InputFileData)
        msg = "Number of UNICODE characters in input file (InputFileDataLen) = " & Len(InputFileData)
        Debug.Print(msg)
        LogFileStream.WriteLine(msg)

        Dim tmp As String = ConfigurationManager.AppSettings("PromptAfterNumFilesWritten")
        Dim PromptAfterNumFilesWritten As Integer
        If IsNumeric(tmp) Then
            PromptAfterNumFilesWritten = Integer.Parse(tmp)
        Else
            msg = "PromptAfterNumFilesWritten App Setting is not numeric! Using 10 as the value"
            Debug.Print(msg)
            LogFileStream.WriteLine(msg)
            PromptAfterNumFilesWritten = 10
        End If

        Do While InputFileDataLen > 0

            CurrentSplitFilePathAndName = SplitFilePathAndNameBase & SplitNum & InputFileExt

            OutputTSO = My.Computer.FileSystem.OpenTextFileWriter(CurrentSplitFilePathAndName,
                                                                  False)

            If InputFileExt = ".html" And SplitNum > 1 Then 'First html split file has HTML headers copied from input file
                ' For other html split files we need to add HTML headers. It seems OK to have meta charset UTF-8 even
                ' in a UTF-16 LE file. It does not trip up Chrome or Word in Devanagari characters display.
                msg = "<html><head><meta charset=""UTF-8""></head><body>"
                OutputTSO.Write(msg)
                CurrentSplitCharsWritten += Len(msg)
                TotalSplitsCharsWritten += Len(msg)
            End If

            Dim InputDataCurrentChunk As String
            Dim InputDataCurrentChunkLen As Long
            InputDataCurrentChunk = Left(InputFileData, SplitSize)
            InputDataCurrentChunkLen = Len(InputDataCurrentChunk)
            OutputTSO.Write(InputDataCurrentChunk)
            Dim TempDataLen As Long
            TempDataLen = InputFileDataLen - InputDataCurrentChunkLen
            InputFileData = Right(InputFileData, TempDataLen)
            InputFileDataLen = Len(InputFileData) 'Should be TempDataLen

            CurrentSplitCharsWritten += InputDataCurrentChunkLen
            TotalSplitsCharsWritten += InputDataCurrentChunkLen

            Dim ix, LenTillEndOfPostOrPage As Long
            Dim CurrentSplitInProgress As Boolean
            CurrentSplitInProgress = True
            Do While CurrentSplitInProgress
                Dim EndMarkerLen
                EndMarkerLen = 0
                ix = InStr(InputFileData, EndOfPostMarker) ' Posts come first followed by pages in backup/export XML file
                ' So blog book html will also be in that order
                If (ix > 0) Then
                    EndMarkerLen = Len(EndOfPostMarker)
                Else
                    ix = InStr(InputFileData, EndOfPageMarker)
                    If (ix > 0) Then
                        EndMarkerLen = Len(EndOfPageMarker)
                    End If
                End If
                If (ix > 0) Then
                    'found end of post or page marker
                    LenTillEndOfPostOrPage = ix + EndMarkerLen - 1
                    OutputTSO.Write(Left(InputFileData, LenTillEndOfPostOrPage))
                    TempDataLen = InputFileDataLen - LenTillEndOfPostOrPage
                    InputFileData = Right(InputFileData, TempDataLen)
                    InputFileDataLen = Len(InputFileData) 'Should be TempDataLen

                    CurrentSplitCharsWritten += LenTillEndOfPostOrPage
                    TotalSplitsCharsWritten += LenTillEndOfPostOrPage
                    CurrentSplitInProgress = False
                Else
                    ' May happen after last end of post or page is processed and text after that till
                    ' end of file is being processed or we may have finished reading all data in which case
                    ' we need not write anything
                    If InputFileData <> "" Then
                        OutputTSO.Write(InputFileData)
                        InputFileData = ""
                        InputFileDataLen = 0
                        CurrentSplitCharsWritten += Len(InputFileData)
                        TotalSplitsCharsWritten += Len(InputFileData)
                    End If
                    CurrentSplitInProgress = False
                End If
            Loop

            If InputFileExt = ".html" Then
                If InputFileDataLen > 0 Then 'not last split file; so write end tags
                    msg = "</body></html>"
                    OutputTSO.Write(msg)
                    CurrentSplitCharsWritten += Len(msg)
                    TotalSplitsCharsWritten += Len(msg)
                End If
            End If

            OutputTSO.Close()
            OutputSplitFilesWritten += 1

            msg = Path.GetFileName(CurrentSplitFilePathAndName) & " having " &
            CurrentSplitCharsWritten & " UNICODE characters written."
            Debug.Print(msg)
            LogFileStream.WriteLine(msg)

            'Now seems to be a good time to read the split file just created and closed, to generate
            'Contents links for it and insert it into the data at the right place and create new file
            ' with contents links.
            Dim CLOutputFilePathAndName
            CLOutputFilePathAndName = ""
            If InputFileExt = ".html" And GenerateContentsLinks Then
                msg = "GenerateAndInsertContentsLinksIntoFile function Is now being invoked for file: " &
                    Path.GetFileName(CurrentSplitFilePathAndName) &
                    " And ContentsLinksInsertMarker specified Is: " & ContentsLinksInsertMarker &
                    "."
                Debug.Print(msg)
                LogFileStream.WriteLine(msg)
                CLOutputFilePathAndName = GenerateAndInsertContentsLinksIntoFile(CurrentSplitFilePathAndName,
                    ContentsLinksInsertMarker)
                If CLOutputFilePathAndName = "" Then
                    msg = "GenerateAndInsertContentsLinksIntoFile did Not create Contents Links output file."
                Else
                    msg = "GenerateAndInsertContentsLinksIntoFile succeeded. " &
                    "Output file with contents links created, filename: " &
                    Path.GetFileName(CLOutputFilePathAndName) & " ."
                End If
                Debug.Print(msg)
                LogFileStream.WriteLine(msg)
            End If

            If (OutputSplitFilesWritten > PromptAfterNumFilesWritten) Then
                ' Safety check for bad arguments causing inadvertent flooding of file system with split files
                Dim Style, response

                msg = "No. of split files written Is: " & OutputSplitFilesWritten &
                " which Is greater than PromptAfterNumFilesWritten which Is: " & PromptAfterNumFilesWritten &
                vbCrLf & vbCrLf & "Are you sure you want to continue?"

                Style = vbYesNo Or vbDefaultButton1    ' Define buttons.
                response = MsgBox(msg, Style, MsgBoxTitle)
                If response = vbYes Then    ' User chose Yes.
                    msg = "User chose to continue after warning:" & vbCrLf & msg
                    Debug.Print(msg)
                    LogFileStream.WriteLine(msg)
                Else
                    msg = "User chose to Not continue (abort) after warning:" & vbCrLf & msg &
                    vbCrLf & "Aborting by exiting Sub!"
                    Debug.Print(msg)
                    LogFileStream.WriteLine(msg)
                    Exit Sub
                End If
            End If

            SplitNum += 1
            CurrentSplitCharsWritten = 0

        Loop

        EndTime = Now

        msg = "Input file: " & InputFilePathAndName & " has been split into " &
        OutputSplitFilesWritten & " files." & vbCrLf &
        "SplitSize specified Is: " & SplitSize & " characters" & vbCrLf &
        "End markers (2) specified are:" & vbCrLf &
        EndOfPostMarker & vbCrLf &
        EndOfPageMarker & vbCrLf & vbCrLf
        msg = msg & "The " & OutputSplitFilesWritten & " output split files have been created in folder: " &
        Directory.GetParent(InputFilePathAndName).ToString() &
        ". They are named as " & Path.GetFileNameWithoutExtension(InputFilePathAndName) & SplitFileNameSuffix &
        "x" & InputFileExt & " where x Is a running number starting with 1." & vbCrLf & vbCrLf
        If InputFileExt = ".html" And GenerateContentsLinks Then
            msg = msg & "GenerateAndInsertContentsLinksIntoFile function was invoked for each output split file." &
            " ContentsLinksInsertMarker specified was: " & ContentsLinksInsertMarker &
            ". Additional output file(s) with contents links may have been created with basename suffix: " &
            ContentsLinksInsertedOutputFileBaseSuffix &
            ". For details, please check log file with suffix: -SBBGCLsLog.txt ."
        End If

        MsgBox(msg, , MsgBoxTitle)
        msg = vbCrLf & "Final SplitBBAndGenCLs Sub. MsgBox messages repeated below." & vbCrLf & msg
        Debug.Print(msg)
        LogFileStream.WriteLine(msg)

        msg = "----- End Final SplitBBAndGenCLs Sub. MsgBox messages -----" & vbCrLf
        Debug.Print(msg)
        LogFileStream.WriteLine(msg)

        msg = "Function SplitBBAndGenCLs execution finishing." & vbCrLf
        msg = msg + "Date & time now (after showing Message Box with program execution info.) is: " & Now & vbCrLf &
            "Time taken for this function execution excluding dialog interaction (in seconds): " _
            & DateDiff("s", StartTime, EndTime)
        Debug.Print(msg)
        LogFileStream.WriteLine(msg)

    End Sub

    ' If output file with contents links has been created, its path and name are returned, else empty string is returned
    Function GenerateAndInsertContentsLinksIntoFile(InputFilePathAndName As String,
                ContentsLinksInsertMarker As String) As String
        Dim InputTSO As System.IO.StreamReader
        Dim InputFileData As String
        Dim InputFileDataLen As Long
        Dim OutputFileData As String
        Dim msg

        InputTSO = My.Computer.FileSystem.OpenTextFileReader(InputFilePathAndName)

        InputFileData = InputTSO.ReadToEnd()
        InputTSO.Close()

        InputFileDataLen = Len(InputFileData)

        Dim ContentsLinksHTML As String
        ContentsLinksHTML = GenerateContentsLinks(InputFileData)

        If ContentsLinksHTML = "" Then
            '        Got nothing to write
            msg = "GenerateContentsLinks(InputFileData) returned empty ContentsLinksHTML for file: " &
            InputFilePathAndName & ". So not creating ContentsLinks output file."
            Debug.Print(msg)
            LogFileStream.WriteLine(msg)
            'MsgBox(msg, , MsgBoxTitle)
            GenerateAndInsertContentsLinksIntoFile = ""
            Exit Function
        End If

        ' Insert generated contents links in input file data at right place
        Dim ix As Long
        ix = InStr(InputFileData, ContentsLinksInsertMarker)
        If ix = 0 Then
            msg = "Did not get insert marker for contents! " &
                "Aborting function GenerateAndInsertContentsLinksIntoFile!"
            Debug.Print(msg)
            LogFileStream.WriteLine(msg)
            '        MsgBox(msg, , MsgBoxTitle)
            GenerateAndInsertContentsLinksIntoFile = ""
            Exit Function
        End If

        msg = "Contents Internal Links created on " & Now & " by Microsoft Visual Basic (.Net) (VB.Net) " _
            & "Function GenerateAndInsertContentsLinksIntoFile in ExportFileFilterAndGenBook project." &
            "<br/><br/>"
        Dim InsertPos As Long
        InsertPos = ix
        ' Insert Contents into InputFileData string at iy and then write out InputFileData to new file
        Dim DataLenAfterInsertPos
        DataLenAfterInsertPos = (InputFileDataLen - InsertPos) + 1
        OutputFileData = Left(InputFileData, InsertPos - 1) & "<br/><br/>" &
                    msg &
                    ContentsLinksHTML &
                    Right(InputFileData, DataLenAfterInsertPos)
        Dim OutputFilePathAndName
        ' Generate output filename with a suffix appended to file base name
        OutputFilePathAndName = Directory.GetParent(InputFilePathAndName).ToString() & "\" &
            Path.GetFileNameWithoutExtension(InputFilePathAndName) &
            ContentsLinksInsertedOutputFileBaseSuffix & ".html"

        Dim OutputTSO As System.IO.StreamWriter

        OutputTSO = My.Computer.FileSystem.OpenTextFileWriter(OutputFilePathAndName, False) 'Overwrite not append

        OutputTSO.Write(OutputFileData)

        OutputTSO.Close()
        'Debug.Print "Inserted Contents Links into input file data and created output file: " & OutputFilePathAndName

        GenerateAndInsertContentsLinksIntoFile = OutputFilePathAndName

    End Function

    Sub TestDriveGenerateContentsLinks()
        Dim InputData As String, TOC As String
        InputData = "<html><body><h1 id=""entry-52"">Thuravoor Narayana Sastrigal related passages ...</h1>" &
        "<br/> data ...<br/><br/>" &
        "<h1 id=""entry-53"">Sahitya Kutuhala: Almost century old book ...</h1><br/> data ...<br/><br/></body></html>"

        TOC = GenerateContentsLinks(InputData)

        Dim msg
        msg = "For InputData:" & vbCrLf & InputData & vbCrLf & vbCrLf &
            "GenerateContentsLinks(InputData) returned:" & vbCrLf & TOC
        Debug.Print(msg)
        MsgBox(msg, , MsgBoxTitle)

        Dim DefaultInputDirectory = ConfigurationManager.AppSettings("DefaultInputDirectory")
        If DefaultInputDirectory <> "" And Not Directory.Exists(DefaultInputDirectory) Then
            MsgBox("Default Input Directory: " & DefaultInputDirectory & " does not exist. " &
                    "TestDriveGenerateContentsLinks() is aborting!", , MsgBoxTitle)
            Exit Sub
        End If

        Dim DefaultHTMLInputFilePathAndName = GetDefaultHTMLInputFilePathAndName()
        If Not File.Exists(DefaultHTMLInputFilePathAndName) Then
            MsgBox("Default HTML Input File: " & DefaultHTMLInputFilePathAndName & " does not exist. " &
                    "TestDriveGenerateContentsLinks() is aborting!", , MsgBoxTitle)
            Exit Sub
        End If

        Dim InputFileData As String
        Dim InputTSO As System.IO.StreamReader
        InputTSO = My.Computer.FileSystem.OpenTextFileReader(DefaultHTMLInputFilePathAndName)

        InputFileData = InputTSO.ReadToEnd()
        InputTSO.Close

        TOC = GenerateContentsLinks(InputFileData)

        msg = "For input file: " & DefaultHTMLInputFilePathAndName & vbCrLf & vbCrLf &
            "GenerateContentsLinks(InputData) returned:" & vbCrLf & TOC
        Debug.Print(msg)
        MsgBox(msg, , MsgBoxTitle)

    End Sub

    'Uses HTMLAgilityPack classes, https://html-agility-pack.net/
    ' Code below seems to work!
    ' Code below needs "Imports HtmlAgilityPack" statement at top of file
    Function GenerateContentsLinks(InputData As String) As String
        Dim HTMLDoc As HtmlDocument 'of HTML Agility Pack
        Dim htmlH1s As HtmlNodeCollection
        Dim htmlH1 As HtmlNode

        Dim ContentsLinksHTML As String
        Dim OneContentLinkHTML As String

        Dim InputDataLen As Long

        Dim ContentID As String
        Dim h1Text As String

        Dim NumPostsAndPages
        Dim FirstPostTitle
        Dim LastPostTitle

        h1Text = ""
        NumPostsAndPages = 0
        FirstPostTitle = ""
        LastPostTitle = ""

        InputDataLen = Len(InputData)
        If InputDataLen = 0 Then
            GenerateContentsLinks = ""
            Exit Function
        End If

        ContentsLinksHTML = ""
        HTMLDoc = New HtmlDocument()
        HTMLDoc.LoadHtml(InputData)
        htmlH1s = HTMLDoc.DocumentNode.SelectNodes("/html/body/h1")

        If htmlH1s IsNot Nothing Then
            'Loop through all h1-elements
            For Each htmlH1 In htmlH1s
                ContentID = htmlH1.Id
                If ContentID <> "" Then
                    NumPostsAndPages += 1
                    h1Text = htmlH1.InnerText
                    If FirstPostTitle = "" Then
                        FirstPostTitle = h1Text
                    End If
                    ' <a href="#entry-52">post-title</a><br/><br/>
                    OneContentLinkHTML = "<a href=""#" & ContentID & """>" & h1Text &
                "</a><br/><br/>" & vbCrLf
                    'Debug.Print "OneContentLinkHTML = " & OneContentLinkHTML
                    ContentsLinksHTML &= OneContentLinkHTML
                    OneContentLinkHTML = ""
                End If
            Next
        End If


        If h1Text <> "" Then
            LastPostTitle = h1Text
        End If

        If ContentsLinksHTML <> "" Then
            Dim temp As String
            temp = "<h1>Contents Internal Links</h1>" &
            "Number of posts and pages: " & NumPostsAndPages & "<br/>" & vbCrLf &
            "First post/page title & date: " & FirstPostTitle & "<br/>" & vbCrLf &
            "Last post/page title & date: " & LastPostTitle & "<br/><br/>" & vbCrLf
            temp = temp & ContentsLinksHTML &
            "<br/>=========================================================<br/>" & vbCrLf
            ContentsLinksHTML = temp
        End If

        GenerateContentsLinks = ContentsLinksHTML

    End Function

    ''WebBrowser version which I think seems To slow program down quite a bit, at times (Not always).
    ''I read somewhere that it Is resource heavy but am Not able To Get that link Or similar link now.
    ''HtmlDocument Class is there both In HTML Agility pack And one Of the standard .Net related 
    '''packs?' (I think that is System.Windows.Forms).
    ''So if below code Is to be used, the "Imports HtmlAgilityPack" statement at top of file needs
    ''to be commented out Or deleted.
    'Function GenerateContentsLinks(InputData As String) As String
    '    Dim HTMLDoc As HtmlDocument
    '    Dim htmlH1s As HtmlElementCollection
    '    Dim htmlH1 As HtmlElement

    '    'https://www.vbforums.com/showthread.php?646095-Convert-html-string-to-HTMLDocument-for-parsing
    '    Dim WB As New WebBrowser
    '    'I suspect that WB WebBrowser object is slowing down program at times.
    '    'Need to check if I can do this without using WB.
    '    'HTMLDoc = WB.Document.OpenNew(True) ' does not work
    '    Dim ContentsLinksHTML As String
    '    Dim OneContentLinkHTML As String

    '    Dim InputDataLen As Long

    '    Dim ContentID As String
    '    Dim h1Text As String

    '    Dim NumPostsAndPages
    '    Dim FirstPostTitle
    '    Dim LastPostTitle

    '    h1Text = ""
    '    NumPostsAndPages = 0
    '    FirstPostTitle = ""
    '    LastPostTitle = ""

    '    InputDataLen = Len(InputData)
    '    If InputDataLen = 0 Then
    '        GenerateContentsLinks = ""
    '        Exit Function
    '    End If

    '    ContentsLinksHTML = ""
    '    'https://learn.microsoft.com/en-us/dotnet/api/system.windows.forms.webbrowser.document?view=windowsdesktop-7.0
    '    WB.ScriptErrorsSuppressed = True
    '    WB.DocumentText = "<HTML></HTML>" ' Hack as I do not how to create a starter HTML document
    '    '                       in the web browser control. Note WB.Document is a read only propery. If a starter
    '    '                       HTML document is not created then WB.Document is Nothing which results in 
    '    '                       if I recall correctly, WB.Document.OpenNew(True) to throw an exception

    '    'WB.DocumentText = InputData ' creates empty <HTML></HTML> document. so does not work
    '    'Below code works; That should help solve our problem 
    '    ' https://learn.microsoft.com/en-us/dotnet/api/system.windows.forms.htmldocument.write?
    '    ' view=windowsdesktop-7.0#system-windows-forms-htmldocument-write(system-string)
    '    'If (WB.Document IsNot Nothing) Then
    '    '    Dim doc As HtmlDocument = WB.Document.OpenNew(True)
    '    '    doc.Write("<HTML><BODY>This is a new HTML document.</BODY></HTML>")
    '    'End If
    '    'Get the source (code) of the webpage
    '    'HTMLDoc.body.innerHTML = InputData
    '    'WB.Document.Body.InnerHtml = InputData ' did not work; threw an exception
    '    HTMLDoc = WB.Document.OpenNew(True)
    '    HTMLDoc.Write(InputData)
    '    'HTMLDoc = WB.Document
    '    'HTMLDoc.Body.InnerHtml = InputData ' throws a null reference exception
    '    'Set reference to all h1-elements in htmlDoc.body.innerHTML
    '    htmlH1s = HTMLDoc.GetElementsByTagName("h1")

    '    'Loop through all h1-elements
    '    For Each htmlH1 In htmlH1s
    '        ContentID = htmlH1.GetAttribute("id")
    '        If ContentID <> "" Then
    '            NumPostsAndPages = NumPostsAndPages + 1
    '            h1Text = htmlH1.InnerText
    '            If FirstPostTitle = "" Then
    '                FirstPostTitle = h1Text
    '            End If
    '            ' <a href="#entry-52">post-title</a><br/><br/>
    '            OneContentLinkHTML = "<a href=""#" & ContentID & """>" & h1Text &
    '            "</a><br/><br/>" & vbCrLf
    '            'Debug.Print "OneContentLinkHTML = " & OneContentLinkHTML
    '            ContentsLinksHTML = ContentsLinksHTML & OneContentLinkHTML
    '            OneContentLinkHTML = ""
    '        End If
    '    Next

    '    If h1Text <> "" Then
    '        LastPostTitle = h1Text
    '    End If

    '    If ContentsLinksHTML <> "" Then
    '        Dim temp As String
    '        temp = "<h1>Contents Internal Links</h1>" &
    '        "Number of posts and pages: " & NumPostsAndPages & "<br/>" & vbCrLf &
    '        "First post/page title & date: " & FirstPostTitle & "<br/>" & vbCrLf &
    '        "Last post/page title & date: " & LastPostTitle & "<br/><br/>" & vbCrLf
    '        temp = temp & ContentsLinksHTML &
    '        "<br/>=========================================================<br/>" & vbCrLf
    '        ContentsLinksHTML = temp
    '    End If

    '    GenerateContentsLinks = ContentsLinksHTML

    'End Function
    Function GetDefaultHTMLInputFilePathAndName() As String
        Dim DefaultHTMLInputFilePathAndName As String
        Dim DefaultInputDirectory As String
        DefaultInputDirectory = ConfigurationManager.AppSettings("DefaultInputDirectory")
        If (DefaultInputDirectory = "") Then
            DefaultHTMLInputFilePathAndName = ConfigurationManager.AppSettings("DefaultHTMLInputFileName")
        Else
            DefaultHTMLInputFilePathAndName = DefaultInputDirectory & "\" &
                ConfigurationManager.AppSettings("DefaultHTMLInputFileName")
        End If
        GetDefaultHTMLInputFilePathAndName = DefaultHTMLInputFilePathAndName
    End Function

End Module
