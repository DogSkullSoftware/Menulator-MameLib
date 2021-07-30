Imports System.Net
Imports System.IO



Public Class App
    Implements IDisposable

    'Public Event StdOut(ByVal sender As Object, ByVal e As DataReceivedEventArgs)
    'Public Event StdErr(ByVal sender As Object, ByVal e As DataReceivedEventArgs)
    'Public Event GameFinished(ByVal sender As Object, ByVal e As EventArgs)
#Region "Shared"
    Public Shared ReadOnly Property Version(ByVal strMamePath As String) As String
        Get
            'Dim s As String = "" '= HelpString(strMamePath)  redirect2.Start(MamePath, "-?")
            If Len(strMamePath) = 0 Then Return ""
            If IO.File.Exists(strMamePath) = False Then Return ""
            Dim s As String = MameProcess.Redirect(strMamePath, "-?")
            'Dim proc As New Process
            'With proc
            '    With .StartInfo
            '        .Arguments = "-?"
            '        .CreateNoWindow = True
            '        .FileName = strMamePath
            '        .RedirectStandardError = True
            '        .RedirectStandardOutput = True
            '        .UseShellExecute = False
            '        .WorkingDirectory = IO.Path.GetDirectoryName(strMamePath)
            '    End With

            '    'AddHandler .ErrorDataReceived, _
            '    '          AddressOf NMakeErrorDataHandler
            '    If .Start() Then
            '        .BeginErrorReadLine()
            '        Dim sr As New StreamReader(.StandardOutput.BaseStream)
            '        'Do
            '        If sr.Peek() >= 0 Then
            '            s = sr.ReadLine '& vbCrLf
            '        End If
            '        'Application.DoEvents()
            '        'Loop Until .HasExited AndAlso sr.Peek < 0
            '        sr.Close()
            '        'logMutex.WaitOne()
            '        'RaiseEvent ProcessTimes(.StartTime, .ExitTime)
            '        'Return True
            '    End If
            '    .Close()
            '    .Dispose()
            '    'Return False
            'End With
            Try
                's = Left$(s, InStr(s, vbCr) - 1)
                Dim i As Integer = InStr(s, "v")
                s = Trim(Mid(s, i + 1, InStr(s, ")") - i))
                Return s
            Catch ex As Exception
                s = Trim(Mid$(s, 11, 20 - 1))
                Return s
            End Try
        End Get
    End Property
    'Public Shared Function VerifyRomsStructured(ByVal strMamePath As String, Optional ByVal strRom As String = "") As VerifiedRoms
    '    myVerified = New VerifiedRoms
    '    SyncRedirect(strMamePath, "-verifyroms " & strRom, AddressOf VerifyRomsOut, Nothing, AddressOf VerifyRomsExit)
    '    Return myVerified
    'End Function
    'Shared proc As Process
    'Private Shared Sub SyncRedirect(ByVal strPath As String, ByVal strParams As String, ByVal STDOutHandler As DataReceivedEventHandler, ByVal STDErrHandler As DataReceivedEventHandler, ByVal ExitHandler As System.EventHandler)
    '    proc = New Process
    '    With proc
    '        .StartInfo.UseShellExecute = False
    '        .StartInfo.RedirectStandardOutput = True
    '        .StartInfo.FileName = strPath
    '        .StartInfo.Arguments = strParams
    '        .StartInfo.WorkingDirectory = IO.Path.GetDirectoryName(strPath)
    '        .StartInfo.CreateNoWindow = True


    '        .StartInfo.RedirectStandardOutput = True
    '        .StartInfo.RedirectStandardError = True
    '        .EnableRaisingEvents = True
    '        AddHandler .OutputDataReceived, STDOutHandler
    '        AddHandler .ErrorDataReceived, STDErrHandler
    '        AddHandler .Exited, ExitHandler


    '        .Start()

    '        .BeginErrorReadLine()
    '        .BeginOutputReadLine()

    '        .WaitForExit()

    '        .Close()
    '    End With
    'End Sub
    'Public Shared Function DownloadLatestMame() As String
    '    Return DownloadLatestMame(IO.Path.GetTempPath)
    'End Function
    '''' <summary>
    '''' Downloads the installer not the direct exe!
    '''' </summary>
    '''' <param name="strTargetDirectory"></param>
    '''' <remarks></remarks>
    'Public Shared Function DownloadLatestMame(ByVal strTargetDirectory As String) As String
    '    Dim s As String = LatestMameBuild(True)
    '    Dim objwebClient As New WebClient
    '    objwebClient.Headers.Add("user-agent", "Mozilla/4.0 (compatible; MSIE 6.0;Windows NT 5.1")
    '    Dim strTarget As String = Path.Combine(strTargetDirectory, Path.GetFileName(s))
    '    objwebClient.DownloadFile(s, strTarget)
    '    objwebClient.Dispose()
    '    Return strTarget
    'End Function

    '''' <summary>
    '''' Returns the latest known MAME build from the internet. Note: Do not rely on this data it is often inacurate
    '''' </summary>
    '''' <param name="bolGiveUrl"></param>
    '''' <returns></returns>
    '''' <remarks></remarks>
    '<Obsolete>
    'Public Shared Function LatestMameBuild(Optional ByVal bolGiveUrl As Boolean = False) As String
    '    Dim objwebClient As New WebClient
    '    Dim s As String = IO.Path.GetTempFileName
    '    Try
    '        objwebClient.Headers.Add("user-agent", "Mozilla/4.0 (compatible; MSIE 6.0;Windows NT 5.1")
    '        objwebClient.DownloadFile("http://various.ru/mame/latest.php", s)
    '    Catch
    '        MsgBox("Internet connection failed. Make sure you are connected.")
    '        Return "Unknown"

    '    End Try
    '    objwebClient.Dispose()
    '    Dim f As New IO.FileStream(s, FileMode.Open)
    '    Dim r As New IO.StreamReader(f)
    '    Dim strMameFile As String = ""
    '    While r.EndOfStream = False
    '        Dim s2 As String = r.ReadLine
    '        If s2.Contains("Download") Then
    '            Dim i As Integer = s2.IndexOf(Chr(34))
    '            Dim i2 As Integer = s2.IndexOf(Chr(34), i + 1)
    '            If bolGiveUrl Then
    '                strMameFile = s2.Substring(i + 1, i2 - i - 1)
    '                'strMameFile = "V" & Path.GetFileNameWithoutExtension(strMameFile).Replace("mame", "")
    '                strMameFile = "http://various.ru/mame/" & Path.GetFileName(strMameFile)
    '            Else
    '                strMameFile = s2.Substring(i + 1, i2 - i - 1)
    '                strMameFile = "V" & Path.GetFileNameWithoutExtension(strMameFile).Replace("mame", "")
    '            End If
    '            Exit While
    '        End If

    '    End While
    '    r.Close()
    '    r.Dispose()
    '    f.Close()
    '    f.Dispose()
    '    Return strMameFile

    'End Function


#End Region


    Private ReadOnly strMamePath As String

    Public Sub New(ByVal sMameExe As String)
        'MyBase.New(sMameExe)
        strMamePath = sMameExe
        'myMame = New MameProcess(sMameExe)
    End Sub
    Public ReadOnly Property MamePath() As String
        Get
            Return strMamePath
        End Get
    End Property
    ''' <summary>
    ''' Displays current MAME version and copyright notice.
    ''' </summary>
    ''' <value></value>
    ''' <returns></returns>
    ''' <remarks></remarks>
    Public ReadOnly Property HelpString() As String
        Get

            Return MameProcess.Redirect(strMamePath, "-?")
        End Get
    End Property
    Public ReadOnly Property Version() As String
        Get
            Dim s As String = HelpString
            Try
                's = Left$(s, InStr(s, vbCr) - 1)
                Dim i As Integer = InStr(s, "v")
                s = Trim(Mid(s, i + 1, InStr(s, ")") - i + 1))
                Return s
            Catch ex As Exception
                s = Trim(Mid$(s, 11, 20))
                Return s
            End Try
        End Get
    End Property
    ''' <summary>
    ''' Displays a summary of all the command line options.
    ''' </summary>
    ''' <returns></returns>
    ''' <remarks></remarks>
    Public Function ShowUsage() As String
        Return MameProcess.Redirect(strMamePath, "-su")
    End Function

    ''' <summary>
    ''' Performs internal validation on every driver in the system. Run this
    ''' before submitting changes to ensure that you haven't violated any of
    ''' the core system rules.
    ''' </summary>
    ''' <value></value>
    ''' <returns></returns>
    ''' <remarks></remarks>
    Public ReadOnly Property Validate() As String
        Get
            Return MameProcess.Redirect(strMamePath, "-validate")
        End Get
    End Property
    ''' <summary>
    ''' Creates the default mame.ini file. All the configuration options
    ''' (not commands) described below can be permanently changed by editing
    ''' this configuration file.
    ''' </summary>
    ''' <remarks></remarks>
    Public Sub CreateConfig()
        MameProcess.Redirect(strMamePath, "-cc")
    End Sub
    Public Shared Sub CreateConfig(ByVal strMamePath As String)
        MameProcess.Redirect(strMamePath, "-cc")
    End Sub
    ''' <summary>
    ''' Displays the current configuration settings.
    ''' </summary>
    ''' <returns></returns>
    ''' <remarks></remarks>
    Public Function ShowConfig() As String
        Return MameProcess.Redirect(strMamePath, "-sc")
    End Function
    ''' <summary>
    ''' Checks for invalid or missing ROM images. By default all drivers that
    ''' have valid ZIP files or directories in the rompath are verified;
    ''' however, you can limit this list by specifying a driver name or
    ''' wildcard.
    ''' </summary>
    ''' <param name="strRom"></param>
    ''' <returns></returns>
    ''' <remarks></remarks>
    Public Function VerifyRoms(Optional ByVal strRom As String = "") As String
        Return MameProcess.Redirect(strMamePath, strRom)
    End Function

    'Public Class GoodRom
    '    Private ReadOnly sName As String
    '    Private sStatus As String
    '    Private ReadOnly sParent As String
    '    Protected Friend Sub New(ByVal strName As String, ByVal strStatus As String, Optional ByVal strParent As String = "")
    '        sName = strName
    '        sStatus = strStatus
    '        sParent = strParent
    '    End Sub
    '    Public ReadOnly Property Name() As String
    '        Get
    '            Return sName
    '        End Get
    '    End Property
    '    Public Property Status() As String
    '        Get
    '            Return sStatus
    '        End Get
    '        Protected Friend Set(ByVal value As String)
    '            sStatus = value
    '        End Set
    '    End Property
    'End Class

    'Public Class BadRom
    '    Inherits GoodRom

    '    Public Class ProblemFile
    '        <Flags()> _
    '        Public Enum Problems
    '            Unknown
    '            NotFound
    '            NeedsRedump
    '            NoGoodDumpKnown
    '            IncorrectChecksum
    '            IncorrectLength
    '            SharedWithParent
    '        End Enum
    '        Private sFileName As String
    '        Private sFileSize As String
    '        'Private strexpectedCRC As String
    '        'Private strexpectedSha1 As String
    '        'Private strreturnedCRC As String
    '        Private strreturnedSize As String
    '        Private eProbs As Problems
    '        Public Property Problem() As Problems
    '            Get
    '                Return eProbs
    '            End Get
    '            Set(ByVal value As Problems)
    '                eProbs = value
    '            End Set

    '        End Property
    '        Public Sub New()

    '        End Sub
    '        Public Sub New(ByVal strFile As String, ByVal strSize As String, ByVal p As Problems)
    '            sFileName = strFile
    '            sFileSize = strSize
    '            eProbs = p
    '        End Sub
    '        Public Property FileName() As String
    '            Get
    '                Return sFileName
    '            End Get
    '            Set(ByVal value As String)
    '                sFileName = value
    '            End Set
    '        End Property
    '        Public Property FileSize() As String
    '            Get
    '                Return sFileSize
    '            End Get
    '            Set(ByVal value As String)
    '                sFileSize = value
    '            End Set
    '        End Property
    '        'Public Property ExpectedCRC() As String
    '        '    Get
    '        '        Return strexpectedCRC
    '        '    End Get
    '        '    Set(ByVal value As String)
    '        '        strexpectedCRC = value
    '        '    End Set
    '        'End Property
    '        'Public Property ExpectedSHA1() As String
    '        '    Get
    '        '        Return strexpectedSha1
    '        '    End Get
    '        '    Set(ByVal value As String)
    '        '        strexpectedSha1 = value
    '        '    End Set
    '        'End Property
    '        'Public Property ReturnedCRC() As String
    '        '    Get
    '        '        Return strreturnedCRC
    '        '    End Get
    '        '    Set(ByVal value As String)
    '        '        strreturnedCRC = value
    '        '    End Set
    '        'End Property
    '        Public Property ReturnedSize() As String
    '            Get
    '                Return strreturnedSize
    '            End Get
    '            Set(ByVal value As String)
    '                strreturnedSize = value
    '            End Set
    '        End Property
    '    End Class

    '    Private lErrors As List(Of ProblemFile)
    '    Public Property Errors() As List(Of ProblemFile)
    '        Get
    '            Return lErrors
    '        End Get
    '        Protected Friend Set(ByVal value As List(Of ProblemFile))
    '            lErrors = value
    '        End Set
    '    End Property

    '    Protected Friend Sub New(ByVal strName As String, ByVal strStatus As String, ByVal strParent As String)
    '        MyBase.New(strName, strStatus, strParent)
    '        lErrors = New List(Of ProblemFile)
    '    End Sub

    'End Class

    'Public Class VerifiedRoms
    '    Private lGood As Dictionary(Of String, GoodRom), lBad As Dictionary(Of String, BadRom)
    '    Private good As Integer, total As Integer
    '    Public ReadOnly Property GoodRoms() As Dictionary(Of String, GoodRom)
    '        Get
    '            Return lGood
    '        End Get
    '    End Property
    '    Public ReadOnly Property BadRoms() As Dictionary(Of String, BadRom)
    '        Get
    '            Return lBad
    '        End Get
    '    End Property
    '    Public Sub New()
    '        lGood = New Dictionary(Of String, GoodRom)
    '        lBad = New Dictionary(Of String, BadRom)
    '    End Sub
    '    Public Property GoodCount() As Integer
    '        Get
    '            Return good
    '        End Get
    '        Set(ByVal value As Integer)
    '            good = value
    '        End Set
    '    End Property
    '    Public Property TotalCount() As Integer
    '        Get
    '            Return total
    '        End Get
    '        Set(ByVal value As Integer)
    '            total = value
    '        End Set
    '    End Property
    'End Class

    'Shared myVerified As VerifiedRoms
    'Dim WithEvents myMame As MameProcess
    'Private Shared Sub VerifyRomsOut(ByVal sender As Object, ByVal e As DataReceivedEventArgs)
    '    If e.Data Is Nothing Then Return
    '    Dim s() As String = Split(e.Data, " ")
    '    If e.Data.Contains("romsets found, ") Then
    '        myVerified.GoodCount = s(3)
    '        myVerified.TotalCount = s(0)
    '        Return
    '    End If
    '    Dim strRom As String, strParent As String = "", strResult As String

    '    Select Case Left(e.Data, 7)
    '        Case "romset "
    '            strRom = s(1)
    '            If s(2).StartsWith("[") Then
    '                strParent = Mid(s(2), 2, s(2).Length - 2)
    '            End If
    '            strResult = Mid(e.Data, InStr(e.Data, " is "))
    '            Select Case strResult
    '                Case "good"
    '                    If Not myVerified.GoodRoms.ContainsKey(strRom) Then
    '                        myVerified.GoodRoms.Add(strRom, New GoodRom(strRom, strResult, strParent))
    '                    End If
    '                Case "bad"
    '                    If Not myVerified.BadRoms.ContainsKey(strRom) Then
    '                        myVerified.BadRoms.Add(strRom, New BadRom(strRom, strResult, strParent))
    '                    Else
    '                        myVerified.BadRoms(strRom).Status = strResult
    '                    End If
    '            End Select
    '        Case "EXPECTE", "   FOUN"
    '            'myVerified.BadRoms.Last.Value.Errors.Add(e.Data)
    '        Case Else
    '            strRom = Mid(e.Data, 1, InStr(e.Data, ":") - 1).Trim ' s(0).Replace(":", "")
    '            If Not myVerified.BadRoms.ContainsKey(strRom) Then
    '                myVerified.BadRoms.Add(strRom, New BadRom(strRom, "", ""))
    '            End If
    '            Dim f As New BadRom.ProblemFile
    '            Dim i As Integer = InStr(e.Data, ":") + 1
    '            f.FileName = Mid(e.Data, i, InStr(i + 1, e.Data, " ") - i).Trim

    '            If e.Data.Contains("NEEDS REDUMP") Then
    '                f.Problem = BadRom.ProblemFile.Problems.NeedsRedump
    '            End If
    '            If e.Data.Contains("NO GOOD DUMP KNOWN") Then
    '                f.Problem = BadRom.ProblemFile.Problems.NoGoodDumpKnown
    '            End If
    '            If e.Data.Contains("NOT FOUND") Then
    '                f.Problem = f.Problem Or BadRom.ProblemFile.Problems.NotFound
    '            End If
    '            If e.Data.Contains("INCORRECT CHECKSUM") Then
    '                f.Problem = f.Problem Or BadRom.ProblemFile.Problems.IncorrectChecksum
    '            End If
    '            If e.Data.Contains("INCORRECT LENGTH") Then
    '                f.Problem = f.Problem Or BadRom.ProblemFile.Problems.IncorrectLength
    '                i += f.FileName.Length + 3
    '                f.FileSize = Mid(e.Data, i, InStr(i + 1, e.Data, ")") - i)
    '                f.ReturnedSize = Mid(e.Data, InStrRev(e.Data, ":") + 2)
    '            End If

    '            If e.Data.Contains("(shared with parent)") Then
    '                f.Problem = f.Problem Or BadRom.ProblemFile.Problems.SharedWithParent
    '            End If

    '            myVerified.BadRoms(strRom).Errors.Add(f)
    '    End Select
    'End Sub
    'Private Shared Sub VerifyRomsExit(ByVal sender As Object, ByVal e As EventArgs)

    'End Sub
    Public Shared Sub DeleteNvram(ByVal strMamePath As String, ByVal strRomName As String, ByVal NVDirectory As String)
        'Dim i As New MAME.INI(strMamePath)
        'Dim s() As String = MAME.INI.ExpandPath(strMamePath, i.RootValues("nvram_directory").ValueOrDefault)
        Dim s As String = MAME.INI.ExpandPath(strMamePath, NVDirectory)
        'For Each st In s
        Dim strBuild As String = IO.Path.Combine(s, strRomName & ".nv")
        If IO.File.Exists(strBuild) Then
            IO.File.Delete(strBuild)
        End If
        'Next
    End Sub
    Public Shared Sub DeleteNvram(ByVal strMamePath As String, ByVal strRomName As String)
        Dim i As New MAME.INI(strMamePath)
        Dim s() As String = MAME.INI.ExpandPath(strMamePath, i.RootValues("nvram_directory").ValueOrDefault)
        For Each st In s
            Dim strBuild As String = IO.Path.Combine(st, strRomName & ".nv")
            If IO.File.Exists(strBuild) Then
                IO.File.Delete(strBuild)
            End If
        Next
    End Sub

    Public Shared Sub DeleteCfg(ByVal strMamePath As String, ByVal strRomName As String)
        Dim i As New MAME.INI(strMamePath)
        Dim s() As String = MAME.INI.ExpandPath(strMamePath, i.RootValues("cfg_directory").ValueOrDefault)
        For Each st In s
            Dim strBuild As String = IO.Path.Combine(st, strRomName & ".cfg")
            If IO.File.Exists(strBuild) Then
                IO.File.Delete(strBuild)
            End If
        Next
    End Sub
    'Public Sub RunGame(ByVal strRomName As String)
    '    'SyncRedirect(strMamePath, strRomName, AddressOf GameOut, AddressOf GameErr, AddressOf GameExit)
    '    myMame.PlayRom(strRomName)
    'End Sub
    'Private Sub GameOut(ByVal sender As Object, ByVal e As DataReceivedEventArgs) Handles myMame.StdOut
    '    RaiseEvent StdOut(Me, e)
    'End Sub
    'Private Sub GameErr(ByVal sender As Object, ByVal e As DataReceivedEventArgs) Handles myMame.ErrOut
    '    RaiseEvent StdErr(Me, e)
    'End Sub
    'Private Sub GameExit(ByVal sender As Object, ByVal e As System.EventArgs) Handles myMame.Exited
    '    RaiseEvent GameFinished(Me, e)
    'End Sub

    'Public Function CloseMAME() As Boolean
    '    Try
    '        If Not proc Is Nothing Then If proc.HasExited = False Then Return proc.CloseMainWindow() Else Return True
    '    Catch ex As Exception
    '        Return False
    '    End Try
    'End Function


#Region " IDisposable Support "
    Private disposedValue As Boolean = False        ' To detect redundant calls

    ' IDisposable
    Protected Overridable Sub Dispose(ByVal disposing As Boolean)
        If Not Me.disposedValue Then
            If disposing Then

                'If proc IsNot Nothing Then proc.Dispose()
            End If


        End If
        Me.disposedValue = True
    End Sub


    ' This code added by Visual Basic to correctly implement the disposable pattern.
    Public Sub Dispose() Implements IDisposable.Dispose
        ' Do not change this code.  Put cleanup code in Dispose(ByVal disposing As Boolean) above.
        Dispose(True)
        GC.SuppressFinalize(Me)
    End Sub
#End Region

End Class
