Imports System.Drawing
Imports System.Runtime.InteropServices
Imports System.Net
Imports System.IO

Public Module clsMameLib
    Public Class MameProcess
        Implements IDisposable

        Public ReadOnly Property PrivilegedProcessorTime() As TimeSpan
            Get
                Return p.PrivilegedProcessorTime
            End Get
        End Property
        Public ReadOnly Property TotalProcessorTime() As TimeSpan
            Get
                Return p.TotalProcessorTime
            End Get
        End Property
        Public ReadOnly Property UserProcessorTime() As TimeSpan
            Get
                Return p.UserProcessorTime
            End Get
        End Property
        Public ReadOnly Property ExitTime() As Date
            Get
                Return p.ExitTime
            End Get
        End Property
        Public ReadOnly Property StartTime() As Date
            Get
                Return p.StartTime
            End Get
        End Property

        Public Property MinWorkingSet() As IntPtr
            Get
                Return p.MinWorkingSet
            End Get
            Set(ByVal value As IntPtr)
                p.MinWorkingSet = value
            End Set
        End Property
        Public Property MaxWorkingSet() As IntPtr
            Get
                Return p.MaxWorkingSet
            End Get
            Set(ByVal value As IntPtr)
                p.MaxWorkingSet = value
            End Set
        End Property
        Public ReadOnly Property NonpagedSystemMemorySize() As Long
            Get
                Return p.NonpagedSystemMemorySize64
            End Get
        End Property
        Public ReadOnly Property PagedMemorySize() As Long
            Get
                Return p.PagedMemorySize64
            End Get
        End Property
        Public ReadOnly Property PagedSystemMemorySize() As Long
            Get
                Return p.PagedSystemMemorySize64
            End Get
        End Property
        Public ReadOnly Property PeakPagedMemorySize() As Long
            Get
                Return p.PeakPagedMemorySize64
            End Get
        End Property
        Public ReadOnly Property PeakVirtualMemorySize() As Long
            Get
                Return p.PeakVirtualMemorySize64
            End Get
        End Property
        Public ReadOnly Property PeakWorkingSet() As Long
            Get
                Return p.PeakWorkingSet64
            End Get
        End Property
        Public ReadOnly Property PrivateMemorySize() As Long
            Get
                Return p.PrivateMemorySize64
            End Get
        End Property

        Public ReadOnly Property VirtualMemorySize() As Long
            Get
                Return p.VirtualMemorySize64
            End Get
        End Property
        Public ReadOnly Property WorkingSet() As Long
            Get
                Return p.WorkingSet64
            End Get
        End Property

        Dim WithEvents p As Process
        Dim strMamePath As String
        Public Event StdOut As DataReceivedEventHandler
        Public Event ErrOut As DataReceivedEventHandler
        Public Event Started As EventHandler
        Public Event Exited As EventHandler

        Public Sub New(ByVal sMamePath As String)
            p = New Process()
            strMamePath = sMamePath
        End Sub

        Public Function PlayRom(ByVal strRomName As String, Optional ByVal strParameters As String = "") As Boolean
            'SyncRedirect(strMamePath, strRomName, AddressOf GameOut, AddressOf GameErr, AddressOf GameExit)
            If Load(strMamePath, strRomName & " " & strParameters) Then
                p.WaitForExit()
                'p.Close()
                Return True
            Else
                Return False
            End If
        End Function
        Public Function PlayRomAsync(ByVal strRomName As String, Optional ByVal strParameters As String = "") As Boolean
            Return Load(strMamePath, strRomName & " " & strParameters)
        End Function

        Private Function Load(ByVal strPath As String, ByVal strParams As String) As Boolean

            With p
                .StartInfo.UseShellExecute = False
                .StartInfo.RedirectStandardOutput = True
                .StartInfo.FileName = strPath
                .StartInfo.Arguments = strParams
                .StartInfo.WorkingDirectory = IO.Path.GetDirectoryName(strPath)
                .StartInfo.CreateNoWindow = True


                .StartInfo.RedirectStandardOutput = True
                .StartInfo.RedirectStandardError = True
                .EnableRaisingEvents = True
                'AddHandler .OutputDataReceived, STDOutHandler
                'AddHandler .ErrorDataReceived, STDErrHandler
                'AddHandler .Exited, ExitHandler


                If .Start() Then

                    .BeginErrorReadLine()
                    .BeginOutputReadLine()
                    Return True
                Else
                    Return False
                End If
                '.WaitForExit()

                '.Close()
            End With
        End Function
        Public Function CloseMame() As Boolean
            Return p.CloseMainWindow()
        End Function
        Public Property Priority() As ProcessPriorityClass
            Get
                Return p.PriorityClass
            End Get
            Set(ByVal value As ProcessPriorityClass)
                p.PriorityClass = value
            End Set
        End Property
        Public Property PriorityBoost() As Boolean
            Get
                Return p.PriorityBoostEnabled
            End Get
            Set(ByVal value As Boolean)
                p.PriorityBoostEnabled = value

            End Set
        End Property
        Public ReadOnly Property HasExited() As Boolean
            Get
                Return p.HasExited
            End Get
        End Property
        Public ReadOnly Property IsResponding() As Boolean
            Get
                Return p.Responding
            End Get
        End Property
        Public Sub Close()
            p.Close()
        End Sub
        Public Shared Function RedirectStream(ByVal strFileName As String, ByVal strParams As String) As StreamReader
            Dim myProcess As New Process
            'Using XMLProc As New Process
            With myProcess
                With .StartInfo
                    .FileName = strFileName
                    .Arguments = strParams
                    .WorkingDirectory = IO.Path.GetDirectoryName(strFileName)
                    .CreateNoWindow = True
                    'set to async            
                    .RedirectStandardOutput = True
                    .UseShellExecute = False
                End With
                'init XML
                'refrence the stream
                'RaiseEvent Begin()
                .Start()
                'timeStart = .StartTime
                'timeEnd = Nothing
                RedirectStream = .StandardOutput
            End With
            'End Using
        End Function
        Public Shared Function Redirect(ByVal strPath As String, ByVal strParams As String) As String
            'Dim consoleApp As New Process
            Dim s As String = ""
            Dim Proc As New Process
            With Proc
                .StartInfo.UseShellExecute = False
                .StartInfo.RedirectStandardOutput = True
                .StartInfo.FileName = strPath
                .StartInfo.Arguments = strParams
                .StartInfo.WorkingDirectory = IO.Path.GetDirectoryName(strPath)
                .StartInfo.CreateNoWindow = True
                'RaiseEvent Begin()
                .Start()
                'timeStart = .StartTime
                'timeEnd = Nothing
                Dim sr As New StreamReader(.StandardOutput.BaseStream)

                Do

                    s &= sr.ReadLine & vbCrLf
                    'System.windows.forms.Application.DoEvents()
                    Threading.Thread.Sleep(0)
                Loop Until sr.EndOfStream
                '.WaitForExit()
                'timeEnd = .ExitTime
                .Close()

            End With
            Redirect = s
            Proc.Dispose()
            'Return consoleApp.StandardOutput.ReadToEnd()
        End Function

#Region " IDisposable Support "
        Private disposedValue As Boolean = False        ' To detect redundant calls

        ' IDisposable
        Protected Overridable Sub Dispose(ByVal disposing As Boolean)
            If Not Me.disposedValue Then
                If disposing Then

                    p.Dispose()
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

        Private Sub p_Disposed(ByVal sender As Object, ByVal e As System.EventArgs) Handles p.Disposed

        End Sub
        Dim strError As String
        Public ReadOnly Property ErrorData() As String
            Get
                Return strError
            End Get
        End Property
        Private Sub p_ErrorDataReceived(ByVal sender As Object, ByVal e As System.Diagnostics.DataReceivedEventArgs) Handles p.ErrorDataReceived
            strError &= e.Data & vbCrLf
            RaiseEvent ErrOut(Me, e)
        End Sub

        Private Sub p_Exited(ByVal sender As Object, ByVal e As System.EventArgs) Handles p.Exited
            RaiseEvent Exited(Me, e)
        End Sub

        Private Sub p_OutputDataReceived(ByVal sender As Object, ByVal e As System.Diagnostics.DataReceivedEventArgs) Handles p.OutputDataReceived
            RaiseEvent ErrOut(Me, e)
        End Sub
    End Class

    <DllImport("WININET", CharSet:=CharSet.Auto)> _
  Private Function InternetGetConnectedState( _
        ByRef lpdwFlags As InternetConnectionState, _
        ByVal dwReserved As Integer) As Boolean

    End Function
    Public Function IsInternetConnected() As Boolean
        Dim flag As InternetConnectionState
        Return InternetGetConnectedState(flag, 0)

    End Function
    <Flags()> _
Enum InternetConnectionState As Integer

        INTERNET_CONNECTION_MODEM = &H1
        INTERNET_CONNECTION_LAN = &H2
        INTERNET_CONNECTION_PROXY = &H4
        INTERNET_RAS_INSTALLED = &H10
        INTERNET_CONNECTION_OFFLINE = &H20
        INTERNET_CONNECTION_CONFIGURED = &H40
    End Enum
    Public Function LoadImage(ByVal ImageURL As String) As Image
        Dim sURL As String = Trim(ImageURL)
        Dim pb As Image = Nothing
        Dim objwebClient As New WebClient
        objwebClient.Headers.Add("user-agent", "Mozilla/4.0 (compatible; MSIE 6.0;Windows NT 5.1")
        Try
            If Not sURL.ToLower().StartsWith("http://") Then
                pb = Image.FromFile(ImageURL)
            Else
                If Not IsInternetConnected() Then Return Nothing
                'sURL = "http://" & sURL
                'Dim objImage As MemoryStream
                AddHandler objwebClient.DownloadDataCompleted, AddressOf DownloadDataCompleted
                bolDone = False
                bolError = False

                'objImage = New MemoryStream(objwebClient.DownloadData(sURL))
                objwebClient.DownloadDataAsync(New Uri(sURL))
                Do Until bolDone

                Loop
                If bolError = False Then
                    'pb = Image.FromStream(objImage)
                    pb = Image.FromStream(New MemoryStream(data, False))
                End If
                'bAns = True
            End If


        Catch ex As Exception
            Debug.Print(ex.Message)
            'bAns = False
            objwebClient.CancelAsync()
            Return Nothing
        Finally
            RemoveHandler objwebClient.DownloadDataCompleted, AddressOf DownloadDataCompleted
            objwebClient.Dispose()
        End Try

        Return pb
    End Function
    Private data() As Byte, bolDone As Boolean, bolError As Boolean
    Private Sub DownloadDataCompleted(ByVal sender As Object, ByVal e As DownloadDataCompletedEventArgs)
        bolDone = True
        If e.Cancelled Or e.Error IsNot Nothing Then
            bolError = True
            Return
        End If
        data = e.Result
    End Sub



    Public Delegate Function LoadImageDelegate(ByVal ImageURL As String) As Image
    Public Sub LoadImage(ByVal ImageURL As String, ByVal Async As AsyncCallback, Optional ByVal StateObject As Object = Nothing)
        Dim l As New LoadImageDelegate(AddressOf LoadImage)
        l.BeginInvoke(ImageURL, Async, StateObject)
    End Sub
End Module
