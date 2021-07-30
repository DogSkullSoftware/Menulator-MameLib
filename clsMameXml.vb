Imports System.Drawing
Imports System.Xml
Imports MAME


Public Interface IRom
    Property Name As String
    Property Path As String
    Property ImagePath As String
    Property Description As String
    Property Favorite As Boolean
End Interface

Public Class MameGame
    Implements IRom

    '    <?xml version="1.0"?>
    '<!DOCTYPE mame [
    '<!ELEMENT mame (machine+)>
    '	<!ATTLIST mame build CDATA #IMPLIED>
    '	<!ATTLIST mame debug (yes|no) "no">
    '	<!ATTLIST mame mameconfig CDATA #REQUIRED>
    '	<!ELEMENT machine (description, year?, manufacturer?, biosset*, rom*, disk*, device_ref*, sample*, chip*, display*, sound?, input?, dipswitch*, configuration*, port*, adjuster*, driver?, device*, slot*, softwarelist*, ramoption*)>
    '		<!ATTLIST machine name CDATA #REQUIRED>
    '		<!ATTLIST machine sourcefile CDATA #IMPLIED>
    '		<!ATTLIST machine isbios (yes|no) "no">
    '		<!ATTLIST machine isdevice (yes|no) "no">
    '		<!ATTLIST machine ismechanical (yes|no) "no">
    '		<!ATTLIST machine runnable (yes|no) "yes">
    '		<!ATTLIST machine cloneof CDATA #IMPLIED>
    '		<!ATTLIST machine romof CDATA #IMPLIED>
    '		<!ATTLIST machine sampleof CDATA #IMPLIED>
    '		<!ELEMENT description (#PCDATA)>
    '		<!ELEMENT year (#PCDATA)>
    '		<!ELEMENT manufacturer (#PCDATA)>
    '		<!ELEMENT biosset EMPTY>
    '			<!ATTLIST biosset name CDATA #REQUIRED>
    '			<!ATTLIST biosset description CDATA #REQUIRED>
    '			<!ATTLIST biosset default (yes|no) "no">
    '		<!ELEMENT rom EMPTY>
    '			<!ATTLIST rom name CDATA #REQUIRED>
    '			<!ATTLIST rom bios CDATA #IMPLIED>
    '			<!ATTLIST rom size CDATA #REQUIRED>
    '			<!ATTLIST rom crc CDATA #IMPLIED>
    '			<!ATTLIST rom sha1 CDATA #IMPLIED>
    '			<!ATTLIST rom merge CDATA #IMPLIED>
    '			<!ATTLIST rom region CDATA #IMPLIED>
    '			<!ATTLIST rom offset CDATA #IMPLIED>
    '			<!ATTLIST rom status (baddump|nodump|good) "good">
    '			<!ATTLIST rom optional (yes|no) "no">
    '		<!ELEMENT disk EMPTY>
    '			<!ATTLIST disk name CDATA #REQUIRED>
    '			<!ATTLIST disk sha1 CDATA #IMPLIED>
    '			<!ATTLIST disk merge CDATA #IMPLIED>
    '			<!ATTLIST disk region CDATA #IMPLIED>
    '			<!ATTLIST disk index CDATA #IMPLIED>
    '			<!ATTLIST disk writable (yes|no) "no">
    '			<!ATTLIST disk status (baddump|nodump|good) "good">
    '			<!ATTLIST disk optional (yes|no) "no">
    '		<!ELEMENT device_ref EMPTY>
    '			<!ATTLIST device_ref name CDATA #REQUIRED>
    '		<!ELEMENT sample EMPTY>
    '			<!ATTLIST sample name CDATA #REQUIRED>
    '		<!ELEMENT chip EMPTY>
    '			<!ATTLIST chip name CDATA #REQUIRED>
    '			<!ATTLIST chip tag CDATA #IMPLIED>
    '			<!ATTLIST chip type (cpu|audio) #REQUIRED>
    '			<!ATTLIST chip clock CDATA #IMPLIED>
    '		<!ELEMENT display EMPTY>
    '			<!ATTLIST display tag CDATA #IMPLIED>
    '			<!ATTLIST display type (raster|vector|lcd|unknown) #REQUIRED>
    '			<!ATTLIST display rotate (0|90|180|270) #REQUIRED>
    '			<!ATTLIST display flipx (yes|no) "no">
    '			<!ATTLIST display width CDATA #IMPLIED>
    '			<!ATTLIST display height CDATA #IMPLIED>
    '			<!ATTLIST display refresh CDATA #REQUIRED>
    '			<!ATTLIST display pixclock CDATA #IMPLIED>
    '			<!ATTLIST display htotal CDATA #IMPLIED>
    '			<!ATTLIST display hbend CDATA #IMPLIED>
    '			<!ATTLIST display hbstart CDATA #IMPLIED>
    '			<!ATTLIST display vtotal CDATA #IMPLIED>
    '			<!ATTLIST display vbend CDATA #IMPLIED>
    '			<!ATTLIST display vbstart CDATA #IMPLIED>
    '		<!ELEMENT sound EMPTY>
    '			<!ATTLIST sound channels CDATA #REQUIRED>
    '		<!ELEMENT input (control*)>
    '			<!ATTLIST input service (yes|no) "no">
    '			<!ATTLIST input tilt (yes|no) "no">
    '			<!ATTLIST input players CDATA #REQUIRED>
    '			<!ATTLIST input coins CDATA #IMPLIED>
    '			<!ELEMENT control EMPTY>
    '				<!ATTLIST control type CDATA #REQUIRED>
    '				<!ATTLIST control player CDATA #IMPLIED>
    '				<!ATTLIST control buttons CDATA #IMPLIED>
    '				<!ATTLIST control reqbuttons CDATA #IMPLIED>
    '				<!ATTLIST control minimum CDATA #IMPLIED>
    '				<!ATTLIST control maximum CDATA #IMPLIED>
    '				<!ATTLIST control sensitivity CDATA #IMPLIED>
    '				<!ATTLIST control keydelta CDATA #IMPLIED>
    '				<!ATTLIST control reverse (yes|no) "no">
    '				<!ATTLIST control ways CDATA #IMPLIED>
    '				<!ATTLIST control ways2 CDATA #IMPLIED>
    '				<!ATTLIST control ways3 CDATA #IMPLIED>
    '		<!ELEMENT dipswitch (dipvalue*)>
    '			<!ATTLIST dipswitch name CDATA #REQUIRED>
    '			<!ATTLIST dipswitch tag CDATA #REQUIRED>
    '			<!ATTLIST dipswitch mask CDATA #REQUIRED>
    '			<!ELEMENT dipvalue EMPTY>
    '				<!ATTLIST dipvalue name CDATA #REQUIRED>
    '				<!ATTLIST dipvalue value CDATA #REQUIRED>
    '				<!ATTLIST dipvalue default (yes|no) "no">
    '		<!ELEMENT configuration (confsetting*)>
    '			<!ATTLIST configuration name CDATA #REQUIRED>
    '			<!ATTLIST configuration tag CDATA #REQUIRED>
    '			<!ATTLIST configuration mask CDATA #REQUIRED>
    '			<!ELEMENT confsetting EMPTY>
    '				<!ATTLIST confsetting name CDATA #REQUIRED>
    '				<!ATTLIST confsetting value CDATA #REQUIRED>
    '				<!ATTLIST confsetting default (yes|no) "no">
    '		<!ELEMENT port (analog*)>
    '			<!ATTLIST port tag CDATA #REQUIRED>
    '			<!ELEMENT analog EMPTY>
    '				<!ATTLIST analog mask CDATA #REQUIRED>
    '		<!ELEMENT adjuster EMPTY>
    '			<!ATTLIST adjuster name CDATA #REQUIRED>
    '			<!ATTLIST adjuster default CDATA #REQUIRED>
    '		<!ELEMENT driver EMPTY>
    '			<!ATTLIST driver status (good|imperfect|preliminary) #REQUIRED>
    '			<!ATTLIST driver emulation (good|imperfect|preliminary) #REQUIRED>
    '			<!ATTLIST driver color (good|imperfect|preliminary) #REQUIRED>
    '			<!ATTLIST driver sound (good|imperfect|preliminary) #REQUIRED>
    '			<!ATTLIST driver graphic (good|imperfect|preliminary) #REQUIRED>
    '			<!ATTLIST driver cocktail (good|imperfect|preliminary) #IMPLIED>
    '			<!ATTLIST driver protection (good|imperfect|preliminary) #IMPLIED>
    '			<!ATTLIST driver savestate (supported|unsupported) #REQUIRED>
    '		<!ELEMENT device (instance*, extension*)>
    '			<!ATTLIST device type CDATA #REQUIRED>
    '			<!ATTLIST device tag CDATA #IMPLIED>
    '			<!ATTLIST device fixed_image CDATA #IMPLIED>
    '			<!ATTLIST device mandatory CDATA #IMPLIED>
    '			<!ATTLIST device interface CDATA #IMPLIED>
    '			<!ELEMENT instance EMPTY>
    '				<!ATTLIST instance name CDATA #REQUIRED>
    '				<!ATTLIST instance briefname CDATA #REQUIRED>
    '			<!ELEMENT extension EMPTY>
    '				<!ATTLIST extension name CDATA #REQUIRED>
    '		<!ELEMENT slot (slotoption*)>
    '			<!ATTLIST slot name CDATA #REQUIRED>
    '			<!ELEMENT slotoption EMPTY>
    '				<!ATTLIST slotoption name CDATA #REQUIRED>
    '				<!ATTLIST slotoption devname CDATA #REQUIRED>
    '				<!ATTLIST slotoption default (yes|no) "no">
    '		<!ELEMENT softwarelist EMPTY>
    '			<!ATTLIST softwarelist name CDATA #REQUIRED>
    '			<!ATTLIST softwarelist status (original|compatible) #REQUIRED>
    '			<!ATTLIST softwarelist filter CDATA #IMPLIED>
    '		<!ELEMENT ramoption (#PCDATA)>
    '			<!ATTLIST ramoption default CDATA #IMPLIED>
    ']>




    'isbios = false isrunnable = true


    Public Property Name As String Implements IRom.Name
    Public Property Path As String Implements IRom.Path
    Public Property Description As String Implements IRom.Description
    Public Property Year As String
    Public Property Category As String
    Public Property Manufacturer As String
    Public Property InputPlayers As String
    Public Property DriverStatus As String
    Public Property ImagePath As String Implements IRom.ImagePath
    Public Property Rating As String
    Public Property Verified As Boolean
    Public Property CloneOf As String
    Public Property IsBios As Boolean

    Dim _xml As XElement
    Public Sub New(data As XElement)
        _xml = data
    End Sub
    Public ReadOnly Property XML As XElement
        Get
            Return _xml
        End Get
    End Property

    Public Overridable ReadOnly Property DisplayProperties As Dictionary(Of String, Object)
        Get
            Dim d As New Dictionary(Of String, Object)
            d.Add("Name", Name)
            d.Add("Description", Description)
            d.Add("Year", Year)
            d.Add("Category", Category)
            d.Add("Manufacturer", Manufacturer)
            d.Add("InputPlayers", InputPlayers)
            d.Add("DriverStatus", DriverStatus)
            d.Add("ImagePath", ImagePath)
            d.Add("Rating", Rating)
            d.Add("Verified", Verified)
            d.Add("CloneOf", CloneOf)
            d.Add("IsBios", IsBios)

            Return d
        End Get
    End Property

    Public Property Favorite As Boolean Implements IRom.Favorite

End Class



Public Class MameXml
    Implements IDisposable

    'Public Root As XStreamingElement
    Dim Reader As XmlReader
    'Dim sMame As StreamMame
    Public Sub New()
        Init()
    End Sub
    Public Sub Init(whereClause As Func(Of XElement, Boolean))
        Me.whereClause = whereClause
        'sMame = New StreamMame(IO.Path.Combine(My.Application.Info.DirectoryPath, "Systems\MAME.xml"))
        'Count = sMame.Count
        'sMame.Reset()
        ''Root = New XStreamingElement("mame",
        '                                 From el In sMame Where ((el.@<runnable> Is Nothing OrElse el.@<runnable> = "yes") AndAlso (el.@<isbios> Is Nothing OrElse el.@<isbios> = "no") AndAlso (el.@<isdevice> Is Nothing OrElse el.@<isdevice> = "no")))
        Dim x As New XmlReaderSettings With {
            .ProhibitDtd = False,
            .CloseInput = True
        }

        Reader = XmlReader.Create(New IO.StreamReader(IO.Path.Combine(My.Application.Info.DirectoryPath, "Systems\MAME.xml")), x)
        Reader.MoveToContent()

        Count = (Aggregate a As XElement In CType(XDocument.ReadFrom(Reader), XElement).Elements("game").Where(whereClause) Where String.IsNullOrEmpty(a.<cloneof>.Value) Into Count)
        Reader.Close()

        Reader = XmlReader.Create(New IO.StreamReader(IO.Path.Combine(My.Application.Info.DirectoryPath, "Systems\MAME.xml")), x)
        Reader.MoveToContent()


        'Root = New XStreamingElement("system", Reader)

    End Sub
    Private Sub Init()
        Init(Function(a)
                 Return True
             End Function)
    End Sub
    Dim whereClause As Func(Of XElement, Boolean)


    ''' <summary>
    ''' Resets the XML Index to 0
    ''' </summary>
    Public Sub Reset()
        Init()
    End Sub
    Dim _count As Integer
    Public Property Count As Integer
        Get
            Return _count
        End Get
        Private Set(value As Integer)
            _count = value
        End Set
    End Property

    ''' <summary>
    ''' 
    ''' </summary>
    ''' <param name="x">Number of items to fetch after current cursor position</param>
    ''' <returns></returns>
    Public Function GetNextX(x As Integer) As List(Of MameGame)
        Dim bucket As New List(Of MameGame)

        'For Each i In (From el In sMame Where (el.@<runnable> Is Nothing OrElse el.@<runnable> = "yes") AndAlso (el.@<isbios> Is Nothing OrElse el.@<isbios> = "no") AndAlso (el.@<isdevice> Is Nothing OrElse el.@<isdevice> = "no") Take x)
        'For Each i In (From el In sMame Take x)
        '    bucket.Add(New MameGame() With {.Name = i.@<name>, .Description = i.<description>.Value, .Year = i.<year>.Value, .Manufacturer = i.<manufacturer>.Value, .Category = i.<genre>.Value, .Rating = i.<rating>.Value, .Verified = IIf(i.<enabled>.Value = "Yes", True, False)})
        'Next
        Dim count As Integer = 0
        While Reader.Read
            If Reader.NodeType = XmlNodeType.Element Then
                If Reader.Name = "game" Then
                    Dim current As XElement = XElement.ReadFrom(Reader)
                    If Not String.IsNullOrEmpty(current.<cloneof>.Value) Then
                        Continue While
                    End If


                    If whereClause IsNot Nothing Then
                        If Not whereClause(current) Then
                            Continue While
                        End If
                    End If

                    Dim m = New MameGame(current) With {.Name = current.@<name>,
                               .Description = current.<description>.Value,
                               .Year = current.<year>.Value,
                               .Manufacturer = current.<manufacturer>.Value,
                               .Category = current.<genre>.Value,
                               .Rating = current.<rating>.Value,
                               .CloneOf = current.<cloneof>.Value,
                               .DriverStatus = current.<status>.Value,
                               .InputPlayers = current.<players>.Value,
                               .Verified = IIf(current.<enabled>.Value = "Yes", True, False),
                                .IsBios = (Not IsNothing(current.<isbios>.Value)) OrElse current.<isbios>.Value = "yes"
                    }
                    bucket.Add(m)
                    count += 1
                End If
            End If

            If count = x Then Exit While
        End While
        Return bucket
    End Function

    Public Shared Function VerifyRoms(strMamePath As String) As Integer
        Dim newElements As New List(Of XElement)
        Dim newDocument As XDocument '= <?xml version="1.0"?>
        '                               <system name="MAME" version=<%= MAME.App.Version(strMamePath) %> lastupdate=<%= Now.Date %>>
        '                               </system>
        'Dim mameprocess As New MAME.MameProcess(strMamePath)

        'TODO: stream this from -verifyroms directly
        'Dim f = MAME.MameProcess.RedirectStream(strMamePath, "-verifyroms") 
        Dim f = IO.File.OpenText(IO.Path.Combine(IO.Path.GetDirectoryName(strMamePath), "verify.txt"))

        Dim x As New XmlReaderSettings With {
            .ProhibitDtd = False
        }
        If Not IO.File.Exists(IO.Path.Combine(IO.Path.GetDirectoryName(strMamePath), "list.xml")) Then
            Dim xx = DirectCast(MameProcess.RedirectStream(strMamePath, "-listxml").BaseStream, IO.FileStream)
            Dim yy As New IO.FileStream(IO.Path.Combine(IO.Path.GetDirectoryName(strMamePath), "list.xml"), IO.FileMode.CreateNew)
            Dim buffer(8 * 1024) As Byte
            Dim Len As Integer
            Do
                Len = xx.Read(buffer, 0, buffer.Length)
                yy.Write(buffer, 0, Len)
            Loop Until Len = 0
            yy.Close()
            xx.Close()
        End If
        Dim xReader = XmlReader.Create(IO.Path.Combine(IO.Path.GetDirectoryName(strMamePath), "list.xml"), x) 'MameProcess.RedirectStream(strMamePath, "-listxml"), x)
        xReader.MoveToContent()
        Dim master = XDocument.Load(xReader)

        While Not f.EndOfStream
            Dim strRom As String = ""
            Dim strParent As String = ""
            Select Case MAME.MameXml.ParseVerify(f.ReadLine, strRom, strParent)
                Case 0
'do nothing
                Case 1
                    Dim listXml = master.Element("mame").<machine>.Where(Function(xa) xa.@<name> = strRom).Single

                    'Debug.Print("")
                    'Dim listxml = <machine>
                    '                  <description><%= strRom %></description>
                    '                  <year></year>
                    '                  <manufacturer></manufacturer>
                    '              </machine>

                    If listXml.@<isbios> = "yes" Then
                        Exit Select
                    End If
                    newElements.Add(<game name=<%= strRom %>>
                                        <%= listXml.<description> %>
                                        <cloneof><%= listXml.@<cloneof> %></cloneof>
                                        <crc></crc>
                                        <%= listXml.<manufacturer> %>
                                        <%= listXml.<year> %>
                                        <players><%= listXml.<input>.@<players> %></players>
                                        <status><%= listXml.<driver>.@<status> %></status>
                                        <genre></genre>
                                        <rating></rating>
                                        <enabled>Yes</enabled>
                                    </game>)
                    'Try
                    'Catch
                    '    Exit While
                    'End Try
            End Select
        End While
        f.Close()
        'mameprocess.Close()

        newElements.Sort(New DescriptionSorter)
        If IO.File.Exists(IO.Path.Combine(IO.Path.GetDirectoryName(strMamePath), "Cat32\Category.ini")) Then
            Dim ff = IO.File.OpenText(IO.Path.Combine(IO.Path.GetDirectoryName(strMamePath), "Cat32\Category.ini"))
            Dim header As String = ""
            While Not ff.EndOfStream
                Dim s = ff.ReadLine()
                If String.IsNullOrEmpty(s) Then Continue While
                If s(0) = "["c Then
                    header = s.Trim("[", "]")
                    Continue While
                ElseIf header = "FOLDER_SETTINGS" Then
                    Continue While
                ElseIf header = "Adult" Then
                    Dim p = newElements.Where(Function(a) a.@<name> = s).SingleOrDefault
                    If p IsNot Nothing Then
                        If String.IsNullOrEmpty(p.<genre>.Value) Then
                            p.<genre>.Value = p.<genre>.Value & " * Mature *"
                        Else
                            p.<genre>.Value = " * Mature *"
                        End If
                    End If
                Else
                    Dim p = newElements.Where(Function(a) a.@<name> = s).SingleOrDefault
                    If p IsNot Nothing Then
                        If String.IsNullOrEmpty(p.<genre>.Value) Then
                            p.<genre>.Value = header
                        ElseIf p.<genre>.Value = " * Mature *" Then
                            p.<genre>.Value = header & p.<genre>.Value
                        Else
                            Dim s1 = Split(p.<genre>.Value, " / ")
                            Dim s2 = Split(header, " / ")
                            Dim i As Integer, count = Math.Min(s1.Count - 1, s2.Count - 1)
                            For i = 0 To count
                                If s1(i) = s2(i) Then
                                    'skip
                                Else
                                    Exit For
                                End If
                            Next
                            If i = count Then Continue While
                            p.<genre>.Value &= " / " & Join(s2.Skip(i).ToArray, " / ")
                        End If
                    End If
                End If
            End While
            ff.Dispose()
        End If
        newDocument = <?xml version="1.0"?>
                      <system name="MAME" version=<%= MAME.App.Version(strMamePath) %> lastupdate=<%= FormatDate(Now) %>>
                          <%= newElements %>
                      </system>

        ' (From a In newElements Order By a.<game>.<description>.Value)
        newDocument.Save(IO.Path.Combine(My.Application.Info.DirectoryPath, "Systems\MAME.xml"))
        newDocument = Nothing
        'MasterFile = Nothing

        GC.Collect()
        Return newElements.Count
    End Function


    Private Class DescriptionSorter
        Implements IComparer(Of XElement)

        Public Function Compare(x As XElement, y As XElement) As Integer Implements IComparer(Of XElement).Compare
            Return String.Compare(x.<description>.Value, y.<description>.Value)
        End Function
    End Class
    Public Shared Function ParseVerify(ByVal strPrint As String, Optional ByRef strRom As String = "", Optional ByRef strParent As String = "") As Integer
        Dim strs() As String
        'Dim s() As String

        's = Split(strPrint, vbCrLf)
        'For t As Integer = 0 To UBound(s)
        If Strings.Left(strPrint, 6) = "romset" Then
            strs = Split(strPrint, " ")
            strRom = strs(1)
            If strs(2).StartsWith("[") Then strParent = strs(2).Trim("[", "]")
            Select Case strs(UBound(strs))
                Case "good", "available"
                    Return 1
                Case Else
                    Return 0
            End Select

            'DoSQL("UPDATE tblGAME SET verified='" & i & "' WHERE name='" & strs(1) & "'")
        End If
        'Next
        Return -1
    End Function

#Region "IDisposable Support"
    Private disposedValue As Boolean ' To detect redundant calls

    ' IDisposable
    Protected Overridable Sub Dispose(disposing As Boolean)
        If Not disposedValue Then
            If disposing Then
                ' TODO: dispose managed state (managed objects).
                'sMame = Nothing
                'Root = Nothing
                Reader.Close()
            End If

            ' TODO: free unmanaged resources (unmanaged objects) and override Finalize() below.
            ' TODO: set large fields to null.
        End If
        disposedValue = True
    End Sub

    ' TODO: override Finalize() only if Dispose(disposing As Boolean) above has code to free unmanaged resources.
    'Protected Overrides Sub Finalize()
    '    ' Do not change this code.  Put cleanup code in Dispose(disposing As Boolean) above.
    '    Dispose(False)
    '    MyBase.Finalize()
    'End Sub

    ' This code added by Visual Basic to correctly implement the disposable pattern.
    Public Sub Dispose() Implements IDisposable.Dispose
        ' Do not change this code.  Put cleanup code in Dispose(disposing As Boolean) above.
        Dispose(True)
        ' TODO: uncomment the following line if Finalize() is overridden above.
        ' GC.SuppressFinalize(Me)
    End Sub
#End Region


    Public Class StreamMame
        Implements IEnumerable(Of XElement), IDisposable

        Public Shared _reader As XmlReader
        Private _uri As String
        Public _Enum As StreamMameEnumerator
        Public Sub New(uri As String)
            _uri = uri
            Reset()
        End Sub
        Public Sub Reset()
            Dim x As New XmlReaderSettings With {
                .ProhibitDtd = False,
                .CloseInput = True
            }
            _reader = XmlReader.Create(_uri, x)
            _reader.MoveToContent()

        End Sub
        Public Function GetEnumerator() As IEnumerator(Of XElement) Implements IEnumerable(Of XElement).GetEnumerator
            _Enum = New StreamMameEnumerator(_uri, _reader)
            Return _Enum
        End Function

        Private Function IEnumerable_GetEnumerator() As IEnumerator Implements IEnumerable.GetEnumerator
            Return Me.GetEnumerator
        End Function

#Region "IDisposable Support"
        Private disposedValue As Boolean ' To detect redundant calls

        ' IDisposable
        Protected Overridable Sub Dispose(disposing As Boolean)
            If Not disposedValue Then
                If disposing Then
                    ' TODO: dispose managed state (managed objects).
                    _Enum.Dispose()
                    _Enum = Nothing
                    If _reader IsNot Nothing Then
                        _reader.Close()
                        _reader = Nothing
                    End If
                End If

                ' TODO: free unmanaged resources (unmanaged objects) and override Finalize() below.
                ' TODO: set large fields to null.
            End If
            disposedValue = True
        End Sub

        ' TODO: override Finalize() only if Dispose(disposing As Boolean) above has code to free unmanaged resources.
        'Protected Overrides Sub Finalize()
        '    ' Do not change this code.  Put cleanup code in Dispose(disposing As Boolean) above.
        '    Dispose(False)
        '    MyBase.Finalize()
        'End Sub

        ' This code added by Visual Basic to correctly implement the disposable pattern.
        Public Sub Dispose() Implements IDisposable.Dispose
            ' Do not change this code.  Put cleanup code in Dispose(disposing As Boolean) above.
            Dispose(True)
            ' TODO: uncomment the following line if Finalize() is overridden above.
            ' GC.SuppressFinalize(Me)
        End Sub
#End Region
    End Class


    Public Class StreamMameEnumerator
        Implements IEnumerator(Of XElement)

        Private _current As XElement
        Private _uri As String
        Friend _reader As XmlReader
        Public Sub New(uri As String, reader As XmlReader)
            _uri = uri
            _reader = reader
            Reset()
        End Sub

        Public ReadOnly Property Current As XElement Implements IEnumerator(Of XElement).Current
            Get
                Return _current
            End Get
        End Property

        Private ReadOnly Property IEnumerator_Current As Object Implements IEnumerator.Current
            Get
                Return Me.Current
            End Get
        End Property

        Public Sub Reset() Implements IEnumerator.Reset

        End Sub


        Public Function MoveNext() As Boolean Implements IEnumerator.MoveNext

            While _reader.Read
                If _reader.NodeType = XmlNodeType.Element Then
                    Select Case _reader.Name
                        'Case "machine"
                        '    _current = XElement.ReadFrom(_reader)
                        '    Return True
                        Case "game"
                            _current = XElement.ReadFrom(_reader)
                            Return True
                    End Select
                End If
            End While
            Return False
        End Function

#Region "IDisposable Support"
        Private disposedValue As Boolean ' To detect redundant calls

        ' IDisposable
        Protected Overridable Sub Dispose(disposing As Boolean)
            If Not disposedValue Then
                If disposing Then
                    ' TODO: dispose managed state (managed objects).
                    '_reader.Close()
                End If

                ' TODO: free unmanaged resources (unmanaged objects) and override Finalize() below.
                ' TODO: set large fields to null.
            End If
            disposedValue = True
        End Sub

        ' TODO: override Finalize() only if Dispose(disposing As Boolean) above has code to free unmanaged resources.
        'Protected Overrides Sub Finalize()
        '    ' Do not change this code.  Put cleanup code in Dispose(disposing As Boolean) above.
        '    Dispose(False)
        '    MyBase.Finalize()
        'End Sub

        ' This code added by Visual Basic to correctly implement the disposable pattern.
        Public Sub Dispose() Implements IDisposable.Dispose
            ' Do not change this code.  Put cleanup code in Dispose(disposing As Boolean) above.
            Dispose(True)
            ' TODO: uncomment the following line if Finalize() is overridden above.
            ' GC.SuppressFinalize(Me)
        End Sub
#End Region


    End Class


End Class
