
Public Class INI
    Public Shared NullString As String = "<NULL> (not set)"
    Private Shared ii As Information
    Private Shared _Headers As New Dictionary(Of String, Header)
    Private Shared _FieldProperties As New Dictionary(Of String, FieldProperty)

    Private strMamePath As String, strMameIni As String
    Private RootMame As ValueDictionary

    Public Class ValueDictionary
        Inherits Dictionary(Of String, FieldValue)
        Public Structure KeyValue
            Dim d As KeyValuePair(Of String, FieldValue)

            Public ReadOnly Property Key() As String
                Get
                    Return d.Key
                End Get
            End Property
            Public ReadOnly Property Value() As FieldValue
                Get
                    Return d.Value
                End Get
            End Property
            Public Sub New(ByVal Key As String, ByVal Value As FieldValue)
                d = New KeyValuePair(Of String, FieldValue)(Key, Value)
            End Sub
            Public Shared Widening Operator CType(ByVal o As KeyValuePair(Of String, FieldValue)) As KeyValue
                Return New KeyValue(o.Key, o.Value)
            End Operator


        End Structure
        Dim _headers As New Dictionary(Of String, Header)
        Public Shadows Sub Add(ByVal key As String, ByVal value As FieldValue)
            MyBase.Add(key, value)
            If Not _headers.ContainsKey(value.Header.Name) Then
                _headers.Add(value.Header.Name, value.Header)
            End If

        End Sub


        Public ReadOnly Property Headers() As Dictionary(Of String, Header)
            Get
                Return _headers
            End Get
        End Property
    End Class

    Public Structure Information
        Private sVersion As String
        Private cRemChar As Char
        Private sHeaderFormat As String
        Private iColumnSpace As Integer
        Private lDBCount As Long
        Public Property Version() As String
            Get
                Return sVersion
            End Get
            Friend Set(ByVal value As String)
                sVersion = value
            End Set
        End Property
        Public Property RemChar() As Char
            Get
                Return cRemChar
            End Get
            Friend Set(ByVal value As Char)
                cRemChar = value
            End Set
        End Property
        Public Property HeaderFormat() As String
            Get
                Return sHeaderFormat
            End Get
            Friend Set(ByVal value As String)
                sHeaderFormat = value
            End Set
        End Property
        Public Property ColumnSpace() As Integer
            Get
                Return iColumnSpace
            End Get
            Friend Set(ByVal value As Integer)
                iColumnSpace = value
            End Set
        End Property
        Public Property DBCount() As Long
            Get
                Return lDBCount
            End Get
            Friend Set(ByVal value As Long)
                lDBCount = value
            End Set
        End Property
    End Structure
    Public Class Header
        Private ReadOnly sName As String
        Private fFields As New List(Of FieldProperty)
        ''' <summary>
        ''' Returns a list of all fields contained within this header
        ''' </summary>
        ''' <value></value>
        ''' <returns></returns>
        ''' <remarks></remarks>
        Public ReadOnly Property Fields() As List(Of FieldProperty)
            Get
                Return fFields
            End Get
        End Property
        ''' <summary>
        ''' Returns the unique name of this header
        ''' </summary>
        ''' <value></value>
        ''' <returns></returns>
        ''' <remarks></remarks>
        Public ReadOnly Property Name() As String
            Get
                Return sName
            End Get
        End Property
        Friend Sub New(ByVal strName As String)
            If strName Is Nothing Then strName = ""
            sName = strName
        End Sub
    End Class
    Public Enum FieldType
        Unknown
        Path
        MultiPath
        File
        [Boolean]
        Int
        Float
        List
        [String]
        ScreenList
        AspectList
        ResolutionList

    End Enum
    ''' <summary>
    ''' Contains information of known fields supplied by config.xml
    ''' </summary>
    ''' <remarks></remarks>
    Public Class FieldProperty
        Private ReadOnly sName As String
        Private tType As FieldType
        Private oDefaultValue As Object
        Private bIsRoot As Boolean
        Private sDescription As String
        Private fMin As Single
        Private fMax As Single
        Private sList As String()
        Private sHeader As String
        Private sAlts() As String
        ''' <summary>
        ''' Returns the associated Header object
        ''' </summary>
        ''' <returns></returns>
        ''' <remarks></remarks>
        Public Function GetHeader() As Header
            Return _Headers(sHeader)
        End Function
        ''' <summary>
        ''' Unique name of the field
        ''' </summary>
        ''' <value></value>
        ''' <returns></returns>
        ''' <remarks></remarks>
        Public ReadOnly Property Name() As String
            Get
                Return sName
            End Get
        End Property
        Friend Sub New(ByVal strName As String)
            sName = strName
        End Sub
        ''' <summary>
        ''' Returns the expression type for this field
        ''' </summary>
        ''' <value></value>
        ''' <returns></returns>
        ''' <remarks></remarks>
        Public Property Type() As FieldType
            Get
                Return tType
            End Get
            Friend Set(ByVal value As FieldType)
                tType = value
            End Set
        End Property
        ''' <summary>
        ''' Returns true is field is marked as a root only field
        ''' </summary>
        ''' <value></value>
        ''' <returns></returns>
        ''' <remarks></remarks>
        Public Property IsRoot() As Boolean
            Get
                Return bIsRoot
            End Get
            Friend Set(ByVal value As Boolean)
                bIsRoot = value
            End Set
        End Property
        ''' <summary>
        ''' Returns descriptive instructions on the use of this field
        ''' </summary>
        ''' <value></value>
        ''' <returns></returns>
        ''' <remarks></remarks>
        Public Property Description() As String
            Get
                Return sDescription
            End Get
            Friend Set(ByVal value As String)
                sDescription = value
            End Set
        End Property
        ''' <summary>
        ''' Returns the lowest possible value for this field
        ''' </summary>
        ''' <value></value>
        ''' <returns></returns>
        ''' <remarks></remarks>
        Public Property Min() As Single
            Get
                Return fMin
            End Get
            Friend Set(ByVal value As Single)
                fMin = value
            End Set
        End Property
        ''' <summary>
        ''' Returns the highest possible value for this field
        ''' </summary>
        ''' <value></value>
        ''' <returns></returns>
        ''' <remarks></remarks>
        Public Property Max() As Single
            Get
                Return fMax
            End Get
            Friend Set(ByVal value As Single)
                fMax = value
            End Set
        End Property
        ''' <summary>
        ''' Returns a list of valid values for this field
        ''' </summary>
        ''' <value></value>
        ''' <returns></returns>
        ''' <remarks></remarks>
        Public Property List() As String()
            Get
                Return sList
            End Get
            Friend Set(ByVal value As String())
                sList = value
            End Set
        End Property
        ''' <summary>
        ''' Returns the name of the header associated with this field
        ''' </summary>
        ''' <value></value>
        ''' <returns></returns>
        ''' <remarks></remarks>
        Public Property Header() As String
            Get
                Return sHeader
            End Get
            Friend Set(ByVal value As String)
                sHeader = value

            End Set
        End Property
        ''' <summary>
        ''' Returns a list of alternate names for the field if any
        ''' </summary>
        ''' <value></value>
        ''' <returns></returns>
        ''' <remarks></remarks>
        Public Property Alts() As String()
            Get
                Return sAlts
            End Get
            Friend Set(ByVal value As String())
                sAlts = value
            End Set
        End Property
        ''' <summary>
        ''' Returns the system default value for this field
        ''' </summary>
        ''' <value></value>
        ''' <returns></returns>
        ''' <remarks></remarks>
        Public Property DefaultValue() As Object
            Get
                Return oDefaultValue
            End Get
            Friend Set(ByVal value As Object)
                oDefaultValue = value
            End Set
        End Property
    End Class
    ''' <summary>
    ''' Contains values for fields from an INI file
    ''' </summary>
    ''' <remarks></remarks>
    Public Class FieldValue
        Private vValue As Object
        Private ReadOnly sName As String
        Private bEnabled As Boolean

        Public Sub New(ByVal strFieldName As String)
            sName = strFieldName

        End Sub
        ''' <summary>
        ''' Gets/sets the current value for this field
        ''' </summary>
        ''' <value></value>
        ''' <returns></returns>
        ''' <remarks></remarks>
        Public Property Value() As Object
            Get
                Return vValue
            End Get
            Set(ByVal value As Object)
                vValue = value
            End Set
        End Property
        ''' <summary>
        ''' Returns the current value if it is not null, otherwise returns the system default value for this field (can be null)
        ''' </summary>
        ''' <value></value>
        ''' <returns></returns>
        ''' <remarks></remarks>
        Public ReadOnly Property ValueOrDefault() As Object
            Get
                If vValue Is Nothing OrElse vValue = NullString Then
                    Return Properties.[DefaultValue]
                Else
                    If TypeOf vValue Is String AndAlso vValue = "" Then
                        Return Nothing
                    Else
                        Return vValue
                    End If
                End If
            End Get
        End Property
        Public ReadOnly Property ValueOrNullString() As Object
            Get
                If vValue Is Nothing OrElse vValue = NullString Then
                    Return NullString
                Else
                    If TypeOf vValue Is String AndAlso vValue = "" Then
                        Return NullString
                    Else
                        Return vValue
                    End If
                End If
            End Get
        End Property
        ''' <summary>
        ''' Returns the unique name for this field
        ''' </summary>
        ''' <value></value>
        ''' <returns></returns>
        ''' <remarks></remarks>
        Public ReadOnly Property Name() As String
            Get
                Return sName
            End Get
        End Property
        ''' <summary>
        ''' Gets/sets whether this field will be commented out
        ''' </summary>
        ''' <value></value>
        ''' <returns></returns>
        ''' <remarks></remarks>
        Public Property Enabled() As Boolean
            Get
                Return bEnabled
            End Get
            Set(ByVal value As Boolean)
                bEnabled = value
            End Set
        End Property
        ''' <summary>
        ''' Returns the associated Header object
        ''' </summary>
        ''' <value></value>
        ''' <returns></returns>
        ''' <remarks></remarks>
        Public ReadOnly Property Header() As Header
            Get
                Try
                    Return _FieldProperties(sName).GetHeader
                Catch
                    Return Nothing
                End Try
            End Get
        End Property
        ''' <summary>
        ''' Returns a FieldProperty object found from the config.xml file containing more information about this field
        ''' </summary>
        ''' <value></value>
        ''' <returns></returns>
        ''' <remarks></remarks>
        Public ReadOnly Property Properties() As FieldProperty
            Get
                Try
                    Return _FieldProperties(sName)
                Catch
                    Return Nothing
                End Try
            End Get
        End Property
    End Class

    Shared Sub New()
        If Not IO.File.Exists("config.xml") Then
            Return
        End If
        Dim x = XDocument.Load("config.xml")
        ii = New Information
        With x.<INI>
            ii.Version = .@<version>
            ii.RemChar = .@<remchar>
            ii.HeaderFormat = .@<headerformat>.Replace("|", vbCrLf)
            ii.DBCount = .@<dbcount>
            ii.ColumnSpace = .@<columnspace>
        End With


        For Each y In x.Elements.Elements
            'Debug.Print(y.@<name>)
            Dim h As New Header(y.@<name>)

            For Each z In y.Elements
                'Debug.Print(z.@<name>)
                Dim f As New FieldProperty(z.@<name>) With {.[DefaultValue] = z.@<default>, .IsRoot = CBool(Val(z.@<root>)), .Description = z.<description>.Value}
                For Each j In z.<alt>
                    If f.Alts Is Nothing Then ReDim f.Alts(0) Else ReDim Preserve f.Alts(UBound(f.Alts) + 1)
                    f.Alts(UBound(f.Alts)) = j.Value
                Next
                Dim g As Object = z.@<type>
                If g IsNot Nothing Then
                    Try
                        f.Type = [Enum].Parse(GetType(FieldType), g, True)
                    Catch
                        f.Type = FieldType.Unknown
                    End Try
                Else
                    f.Type = FieldType.Unknown
                End If

                Select Case f.Type
                    Case FieldType.Int, FieldType.Float
                        f.Min = z.@<min>
                        f.Max = z.@<max>
                    Case FieldType.List
                        f.List = Split(z.@<list>, "|")
                    Case FieldType.File
                        f.List = Split(z.@<ext>, "|")
                End Select
                f.Header = h.Name
                _FieldProperties.Add(f.Name, f)

                h.Fields.Add(f)
            Next
            _Headers.Add(h.Name, h)
        Next
        x = Nothing
    End Sub
    ''' <summary>
    ''' Dictonary of known header strings
    ''' </summary>
    ''' <remarks></remarks>  
    Public Shared ReadOnly Property Headers() As Dictionary(Of String, Header)
        Get
            Return _Headers
        End Get
    End Property
    ''' <summary>
    ''' Dictionary of all known ini fields (relies on config.xml)
    ''' </summary>
    ''' <remarks></remarks> 
    Public Shared ReadOnly Property FieldProperties() As Dictionary(Of String, FieldProperty)
        Get
            Return _FieldProperties
        End Get
    End Property
    ''' <summary>
    ''' Attempts to find a valid mame*.ini file
    ''' </summary>
    ''' <param name="strMameExe"></param>
    ''' <param name="strMameIniPath"></param>
    ''' <returns></returns>
    ''' <remarks></remarks>
    Public Shared Function MameIniExists(ByVal strMameExe As String, Optional ByRef strMameIniPath As String = "") As Boolean
        strMameIniPath = IO.Path.Combine(IO.Path.GetDirectoryName(strMameExe), "mame.ini")
        If IO.File.Exists(strMameIniPath) Then
            Return True
        End If
        strMameIniPath = IO.Path.Combine(IO.Path.GetDirectoryName(strMameExe), IO.Path.GetFileNameWithoutExtension(strMameExe) & ".ini")
        If IO.File.Exists(strMameIniPath) Then
            Return True
        End If

        strMameIniPath = IO.Path.Combine(IO.Path.GetDirectoryName(strMameExe), "mamed.ini")
        If IO.File.Exists(strMameIniPath) Then
            Return True
        End If
        '        Throw New IO.FileLoadException("No Mame.Ini found. Use App.CreateConfig first.")
        Return False
    End Function
    ''' <summary>
    ''' Returns information gathered from config.xml
    ''' </summary>
    ''' <value></value>
    ''' <returns></returns>
    ''' <remarks></remarks>
    Public Shared ReadOnly Property INIInformation() As Information
        Get
            Return ii
        End Get
    End Property
    Public Shared Function ExpandPath(ByVal strMamePath As String, ByVal strRelativePath As String) As String
        Try
            Dim s() As String = Split(strRelativePath, ";")
            If s Is Nothing Then ReDim s(0) : s(0) = strRelativePath
            If IO.Path.IsPathRooted(s(0)) = False Then

                Return IO.Path.GetFullPath(IO.Path.Combine(IO.Path.GetDirectoryName(strMamePath), s(0)))
            Else
                Return s(0)
            End If
        Catch ex As ArgumentException
            Return Nothing
        End Try
    End Function
    Public Shared Function ExpandPath(ByVal strMamePath As String, ByVal strRelativePaths() As String) As String()
        Dim d(UBound(strRelativePaths)) As String
        For t As Integer = 0 To UBound(strRelativePaths)
            d(t) = ExpandPath(strMamePath, strRelativePaths(t))
        Next
        Return d
    End Function
    Public Shared Function CollapsePath(ByVal strMamePath As String, ByVal strPath As String) As String
        Dim i As String = IO.Path.GetDirectoryName(IO.Path.GetFullPath(strPath)) & "\"
        Dim i2 As String = IO.Path.GetDirectoryName(IO.Path.GetFullPath(strMamePath)) & "\"

        Return i.Replace(i2, "") & IO.Path.GetFileName(strPath)
    End Function
    Public Shared Function CollapsePath(ByVal strMamePath As String, ByVal strPaths As String()) As String()
        'Dim i As String = IO.Path.GetDirectoryName(IO.Path.GetFullPath(strPath)) & "\"
        'Dim i2 As String = IO.Path.GetDirectoryName(IO.Path.GetFullPath(strMamePath)) & "\"

        'Return i.Replace(i2, "") & IO.Path.GetFileName(strPath)
        Dim d(UBound(strPaths)) As String
        For t As Integer = 0 To UBound(strPaths)
            d(t) = CollapsePath(strMamePath, strPaths(t))
        Next
        Return d
    End Function
    Public Shared Function GetNameFromAlt(ByVal strAltName As String) As String
        If _FieldProperties.ContainsKey(strAltName) Then
            Return strAltName
        End If
        Dim test2 As String = ""
        If strAltName.StartsWith("no") Then
            'could be a boolean
            test2 = Mid(strAltName, 3)
            If _FieldProperties.ContainsKey(test2) Then
                Return test2
            End If
        End If
        For Each f In _FieldProperties
            If f.Value.Alts Is Nothing Then Continue For
            If f.Value.Alts.Contains(strAltName) Then
                Return f.Key
            End If
            If test2 <> "" AndAlso f.Value.Alts.Contains(test2) Then
                Return f.Key
            End If
        Next
        Return Nothing
    End Function

    Public Sub New(ByVal strMameExe As String)
        strMamePath = strMameExe
        strMameIni = IO.Path.Combine(IO.Path.GetDirectoryName(strMameExe), "mame.ini")
        If Not IO.File.Exists(strMameIni) Then
            strMameIni = IO.Path.Combine(IO.Path.GetDirectoryName(strMameExe), IO.Path.GetFileNameWithoutExtension(strMameExe) & ".ini")
            If Not IO.File.Exists(strMameIni) Then
                strMameIni = IO.Path.Combine(IO.Path.GetDirectoryName(strMameExe), "mamed.ini")
                If Not IO.File.Exists(strMameIni) Then
                    Throw New IO.FileLoadException("No Mame.Ini found. Use App.CreateConfig first.")
                End If
            End If
        End If
        RootMame = _Load(strMameIni, _FieldProperties, Nothing)
    End Sub
    ''' <summary>
    ''' Attempts to find an ini file from a rom name
    ''' </summary>
    ''' <param name="strRomName">Rom to find ini for</param>
    ''' <param name="strRomIniPath">Pointer to a string that will recieve the found ini path</param>
    ''' <returns></returns>
    ''' <remarks></remarks>
    Public Function RomIniExists(ByVal strRomName As String, ByRef strRomIniPath As String) As Boolean
        Dim strIniPaths() As String = ExpandPath(Split(RootValues("inipath").Value, ";"))
        For Each s In strIniPaths
            strRomIniPath = IO.Path.Combine(s, strRomName & ".ini")
            If IO.File.Exists(strRomIniPath) Then
                Return True
            End If
        Next
        Return False
    End Function

    Public ReadOnly Property IsINIRoot() As Boolean
        Get
            Dim strTest As String = IO.Path.Combine(IO.Path.GetDirectoryName(strMamePath), "mame.ini")
            If IO.File.Exists(strTest) Then
                Return True
            End If
            strTest = IO.Path.Combine(IO.Path.GetDirectoryName(strMamePath), IO.Path.GetFileNameWithoutExtension(strMamePath) & ".ini")
            If IO.File.Exists(strTest) Then
                Return True
            End If

            strTest = IO.Path.Combine(IO.Path.GetDirectoryName(strMamePath), "mamed.ini")
            If IO.File.Exists(strTest) Then
                Return True
            End If
        End Get
    End Property
    ''' <summary>
    ''' Returns a list of existing ini's in the order in which mame will parse them
    ''' </summary>
    ''' <param name="strRomName"></param>
    ''' <param name="bolIsVector"></param>
    ''' <param name="strDriverName"></param>
    ''' <param name="strCloneOf"></param>
    ''' <returns></returns>
    ''' <remarks></remarks>
    Public Function IniHierarchy(ByVal strRomName As String, ByVal bolIsVector As Boolean, ByVal strDriverName As String, Optional ByVal strCloneOf As String = "") As String()
        Dim strTest As String
        Dim theReturn As String() = Nothing
        Dim IniPaths As String() = ExpandPath(strMamePath, Split(RootMame("inipath").ValueOrDefault, ";"))
        'ReDim Preserve IniPaths(UBound(IniPaths) + 1)
        'For z As Integer = UBound(IniPaths) To 1 Step -1
        '    IniPaths(z) = IniPaths(z - 1)
        'Next
        'IniPaths(0) = IO.Path.GetDirectoryName(strMamePath)
        Dim strTests() As String = {"mame.ini", IO.Path.Combine(IO.Path.GetDirectoryName(strMamePath), IO.Path.GetFileNameWithoutExtension(strMamePath) & ".ini"), "mamed.ini", _
                                    "vector.ini", strDriverName & ".ini", strCloneOf & ".ini", strRomName & ".ini"}
        For z As Integer = 0 To UBound(strTests)
            If z = 3 And bolIsVector = False Then Continue For
            If z = 4 And strDriverName = "" Then Continue For
            If z = 5 And strCloneOf = "" Then Continue For
            If z = 6 And strRomName = "" Then Continue For
            For Each s As String In IniPaths
                strTest = IO.Path.Combine(s, strTests(z))
                If IO.File.Exists(strTest) Then
                    If theReturn Is Nothing Then
                        ReDim theReturn(0)
                    Else
                        ReDim Preserve theReturn(UBound(theReturn) + 1)
                    End If
                    theReturn(UBound(theReturn)) = strTest
                End If
            Next
        Next
        Return theReturn
    End Function
    ''' <summary>
    ''' Returns an expanded absolute path from mame's root directory
    ''' </summary>
    ''' <param name="strRelativePath"></param>
    ''' <returns></returns>
    ''' <remarks></remarks>
    Public Function ExpandPath(ByVal strRelativePath As String) As String
        Return ExpandPath(strMamePath, strRelativePath)
    End Function
    Public Function ExpandPath(ByVal strRelativePaths() As String) As String()
        Return ExpandPath(strMamePath, strRelativePaths)
    End Function
    'Public Function CollapsePath(ByVal strPath As String)
    '    Return CollapsePath(Me.strMamePath, strPath)
    'End Function
    'Public Function CollapseDirectory(ByVal strDir As String)
    '    Dim i As String = IO.Path.GetDirectoryName(IO.Path.GetFullPath(strMamePath)) & "\"
    '    Dim i2 As String = IO.Path.GetFullPath(strDir) & "\"

    '    i = i2.Replace(i, "")
    '    If i = "" Then i = "." Else If i.EndsWith("\") Then i = i.Remove(i.Count - 1, 1)
    '    Return i
    'End Function

    ''' <summary>
    ''' Returns the path to the mame directory for this instance
    ''' </summary>
    ''' <value></value>
    ''' <returns></returns>
    ''' <remarks></remarks>
    Public ReadOnly Property MamePath() As String
        Get
            Return strMamePath
        End Get
    End Property
    Public ReadOnly Property MameINI() As String
        Get
            Return strMameIni
        End Get
    End Property
    ''' <summary>
    ''' Returns a complete dictionary of fields filled with data gathered from mame*.ini
    ''' </summary>
    ''' <value></value>
    ''' <returns></returns>
    ''' <remarks></remarks>
    Public ReadOnly Property RootValues() As ValueDictionary
        Get
            Return RootMame
        End Get
    End Property
    ''' <summary>
    ''' Loads an INI file based from the rom name. Loops through all inipaths until a file is found, if bMergeWithRoot is True then the mame.ini is loaded into the Return first then any values found in the file overwrite. If no file is found then if bMergeWithRoot is True the Return is DefaultMame otherwise Return is Nothing
    ''' </summary>
    ''' <param name="strRomName"></param>
    ''' <param name="bMergeWithRoot"></param>
    ''' <returns></returns>
    ''' <remarks></remarks>
    Public Function LoadRom(ByVal strRomName As String, Optional ByVal bMergeWithRoot As Boolean = False) As ValueDictionary
        Dim s() As String = ExpandPath(strMamePath, Split(RootMame("inipath").Value, ";"))
        For Each st As String In s
            Dim test As String = IO.Path.Combine(st, strRomName & ".ini")
            If IO.File.Exists(test) Then
                Return LoadIni(test, bMergeWithRoot)
            End If
        Next
        If bMergeWithRoot Then
            Return RootMame
        Else
            Return Nothing
        End If
    End Function
    Public Function LoadIni(ByVal strIni As String, Optional ByVal bMergeWithRoot As Boolean = False, Optional ByVal bLoadDefaults As Boolean = True) As ValueDictionary
        If IO.File.Exists(strIni) Then
            If bMergeWithRoot Then
                If bLoadDefaults Then
                    Return _Load(strIni, INI._FieldProperties, RootMame)
                Else
                    Return _Load(strIni, INI._FieldProperties, Nothing, False)
                End If
            Else
                Return _Load(strIni)
            End If
        End If
        If bMergeWithRoot Then
            Return RootMame
        Else
            Return Nothing
        End If
    End Function
    ''' <summary>
    ''' Returns a dictionary with only the fields found in the ini
    ''' </summary>
    ''' <param name="strFile"></param>
    ''' <returns></returns>
    ''' <remarks></remarks>
    Private Function _Load(ByVal strFile As String) As ValueDictionary
        Dim reader As New FileIO.TextFieldParser(strFile)
        reader.TrimWhiteSpace = True
        reader.CommentTokens = New String() {ii.RemChar}

        'reader.Delimiters = New String() {vbCrLf}
        Dim l As New ValueDictionary
        While Not reader.EndOfData
            Dim s As String = reader.ReadLine
            s = Trim(s)
            If Len(s) = 0 Then Continue While
            If Left(s, 1) = ii.RemChar Then
                Continue While
                'f.Enabled = False
            Else
                'f.Enabled = True
            End If
            Dim s2(1) As String
            If InStr(s, " ") <> 0 Then
                s2(0) = Trim(Left$(s, InStr(s, " ")))
                s2(1) = Trim(Mid(s, Len(s2(0)) + 1))
            Else
                s2(0) = s
                s2(1) = ""
            End If

            If UBound(s2) Then
                Dim f As FieldValue
                f = New FieldValue(s2(0))
                f.Value = s2(1)
                f.Enabled = True
                l.Add(f.Name, f)
            End If



        End While
        reader.Close()
        reader.Dispose()
        Return l
    End Function
    ''' <summary>
    ''' Returns a dictionary with all known fields and their default values whether present in the ini or not 
    ''' then fills in values from the ini where needed
    ''' </summary>
    ''' <param name="strIni"></param>
    ''' <param name="Properties"></param>
    ''' <param name="Defaults"></param>
    ''' <returns></returns>
    ''' <remarks></remarks>
    Private Function _Load(ByVal strIni As String, ByVal Properties As Dictionary(Of String, FieldProperty), ByVal Defaults As ValueDictionary, Optional ByVal bPopulateInternalDefaults As Boolean = True) As ValueDictionary
        Dim a As New ValueDictionary
        For Each j As KeyValuePair(Of String, FieldProperty) In Properties
            Dim f As New FieldValue(j.Key)
            Dim testf As FieldValue = Nothing
            If Defaults IsNot Nothing AndAlso Defaults.TryGetValue(j.Key, testf) Then
                f.Enabled = testf.Enabled
                f.Value = testf.Value
            Else
                f.Enabled = False
                If bPopulateInternalDefaults Then f.Value = j.Value.[DefaultValue]
            End If
            a.Add(j.Key, f)
        Next

        _Load(strIni, a)
        Return a
    End Function
    ''' <summary>
    ''' Fills found values into the supplied dictionary
    ''' </summary>
    ''' <param name="strIni"></param>
    ''' <param name="dic"></param>
    ''' <remarks></remarks>
    Private Sub _Load(ByVal strIni As String, ByRef dic As ValueDictionary)
        Dim reader As New FileIO.TextFieldParser(strIni)
        reader.TrimWhiteSpace = True
        reader.CommentTokens = New String() {ii.RemChar}

        'reader.Delimiters = New String() {vbCrLf}
        While Not reader.EndOfData
            Dim s As String = reader.ReadLine
            s = Trim(s)
            If Len(s) = 0 Then Continue While
            If Left(s, 1) = ii.RemChar Then
                Continue While
                'f.Enabled = False
            Else
                'f.Enabled = True
            End If
            Dim s2(1) As String
            s2(0) = Trim(Left$(s, InStr(s, " ")))
            s2(1) = Trim(Mid(s, Len(s2(0)) + 1))
            If s2(0) = "" Then s2(0) = s2(1) : s2(1) = ""
            If UBound(s2) Then
                Dim f As FieldValue
                f = New FieldValue(s2(0))
                f.Value = s2(1)
                f.Enabled = True
                If _FieldProperties.ContainsKey(s2(0)) = False Then
                    Dim f2 As New FieldProperty(s2(0)) With {.Header = ""}
                    _FieldProperties.Add(s2(0), f2)
                    _Headers("").Fields.Add(f2)
                End If
                dic(f.Name) = f
            End If

        End While
        reader.Close()
        reader.Dispose()
    End Sub
    ''' <summary>
    ''' Returns an empty, but complete, dictionary 
    ''' </summary>
    ''' <remarks></remarks>
    Private Function _Load(ByVal Properties As Dictionary(Of String, FieldProperty), ByVal bolRootValues As Boolean) As ValueDictionary
        Dim a As New ValueDictionary
        For Each j As KeyValuePair(Of String, FieldProperty) In Properties
            If bolRootValues = False And j.Value.IsRoot Then Continue For
            Dim f As New FieldValue(j.Key)
            'Dim testf As FieldValue = Nothing
            'If Defaults IsNot Nothing AndAlso Defaults.TryGetValue(j.Key, testf) Then
            ' f.Enabled = testf.Enabled
            ' f.Value = testf.Value
            'Else

            f.Enabled = False
            'f.Value = j.Value.Default
            'End If
            a.Add(j.Key, f)
        Next
        Return a
    End Function
    ''' <summary>
    ''' Returns a dictionary with all known fields set to null
    ''' </summary>
    ''' <param name="bolIncludeRootValues"></param>
    ''' <returns></returns>
    ''' <remarks></remarks>
    Public Function LoadEmpty(ByVal bolIncludeRootValues As Boolean) As ValueDictionary
        Return _Load(_FieldProperties, bolIncludeRootValues)
    End Function
    ''' <summary>
    ''' Returns a dictionary with all known fields set to their system default value.
    ''' </summary>
    ''' <returns></returns>
    ''' <param name="bolIncludeRootValues">Set to True if the dictionary is to include root values (values only used by the root mame*.ini file)</param>
    ''' <remarks></remarks>
    Public Function LoadDefaults(ByVal bolIncludeRootValues As Boolean) As ValueDictionary
        Dim a As New ValueDictionary
        For Each j As KeyValuePair(Of String, FieldProperty) In _FieldProperties
            If bolIncludeRootValues = False And j.Value.IsRoot Then Continue For
            Dim f As New FieldValue(j.Key)
            Dim testf As FieldValue = Nothing
            'If RootMame.TryGetValue(j.Key, testf) Then
            '    f.Enabled = testf.Enabled
            '    f.Value = testf.Value
            'Else
            f.Enabled = False
            f.Value = j.Value.[DefaultValue]
            'End If
            a.Add(j.Key, f)
        Next
        Return a
    End Function
    ''' <summary>
    ''' Returns a dictionary with all known fields set to the current mame*.ini values or its default if not found in mame*.ini
    ''' </summary>
    ''' <param name="bolIncludeRootValues"></param>
    ''' <returns></returns>
    ''' <remarks></remarks>
    Public Function LoadFromRoot(ByVal bolIncludeRootValues As Boolean) As ValueDictionary
        Dim a As New ValueDictionary
        For Each j As KeyValuePair(Of String, FieldProperty) In _FieldProperties
            If bolIncludeRootValues = False And j.Value.IsRoot Then Continue For
            Dim f As New FieldValue(j.Key)
            Dim testf As FieldValue = Nothing
            If RootMame.TryGetValue(j.Key, testf) Then
                f.Enabled = testf.Enabled
                f.Value = testf.Value
            Else
                f.Enabled = False
                f.Value = j.Value.[DefaultValue]
            End If
            a.Add(j.Key, f)
        Next
        Return a
    End Function
    Public Function LoadFromINI(ByVal strFile As String, ByVal bolIncludeRootValues As Boolean) As ValueDictionary
        Dim reader As New FileIO.TextFieldParser(strFile)
        reader.TrimWhiteSpace = True
        reader.CommentTokens = New String() {ii.RemChar}

        'reader.Delimiters = New String() {vbCrLf}
        Dim l As ValueDictionary = LoadEmpty(bolIncludeRootValues)
        While Not reader.EndOfData
            Dim s As String = reader.ReadLine
            'Dim bolEnabled As Boolean


            s = Trim(s)
            If Len(s) = 0 Then Continue While
            If Left(s, 1) = ii.RemChar Then
                Continue While
                'f.Enabled = False
                'bolEnabled = False
                's = s.Replace(ii.RemChar, "")
            Else
                'f.Enabled = True
                'bolEnabled = True
            End If
            Dim s2(1) As String
            If InStr(s, " ") <> 0 Then
                s2(0) = Trim(Left$(s, InStr(s, " ")))
                s2(1) = Trim(Mid(s, Len(s2(0)) + 1))
            Else
                s2(0) = s
                s2(1) = ""
            End If

            If UBound(s2) Then
                'Dim f As FieldValue
                'f = New FieldValue(s2(0))
                'f.Value = s2(1)
                'f.Enabled = True
                'l.Add(f.Name, f)
                l(s2(0)).Value = s2(1)
                l(s2(0)).Enabled = True
            End If



        End While
        reader.Close()
        reader.Dispose()
        Return l
    End Function
    Public Function LoadFromRomName(ByVal strRomName As String, ByVal bolIncludeRootValues As Boolean) As ValueDictionary
        Dim strRomIni As String = ""
        If RomIniExists(strRomName, strRomIni) Then
            Return LoadFromINI(strRomIni, bolIncludeRootValues)
        End If
        Return Nothing
    End Function
    'Public Sub Save(Optional ByVal strSaveAsPath As String = "")
    '    If strSaveAsPath = "" Then strSaveAsPath = strMameIni
    '    Dim fil As New IO.StreamWriter(strSaveAsPath, False)

    '    For t As Integer = 0 To Headers.Values.Count - 1
    '        Dim bolWroteHeader As Boolean = False
    '        If bolIsRoot Then
    '            If Len(Headers(t).Name) > 0 Then
    '                fil.WriteLine(Replace(ii.HeaderFormat, "$", Headers(t).Name))
    '            End If
    '            bolWroteHeader = True
    '        End If

    '        If Headers(t).Children IsNot Nothing Then
    '            For z As Integer = 0 To UBound(Headers(t).Children)
    '                With fFields(Headers(t).Children(z))
    '                    If bolIsRoot = False And .RootField = True Then Continue For
    '                    If bolIsRoot = False And .Enabled = False Then Continue For
    '                    Dim a As Integer = 0
    '                    If .Value.ToString = NullString Then .Enabled = False
    '                    If Not .Enabled Then fil.Write(ii.RemChar & " ") : a = 2
    '                    If bolWroteHeader = False Then
    '                        If Len(Headers(t).Name) > 0 Then
    '                            fil.WriteLine(Replace(ii.HeaderFormat, "$", Headers(t).Name))
    '                        End If
    '                        bolWroteHeader = True
    '                    End If
    '                    fil.WriteLine(.Name & Space(ii.ColumnSpace - Len(.Name) - a) & .Value(True))
    '                End With
    '            Next
    '            fil.WriteLine()
    '        End If

    '    Next
    '    fil.Close()

    'End Sub
    Public Shared Sub Save(ByVal strSaveAsFile As String, ByVal d As ValueDictionary, ByVal bolWriteRootFields As Boolean, ByVal bolWriteDisabled As Boolean, ByVal bolWriteNulls As Boolean)
        Dim fil As New IO.StreamWriter(strSaveAsFile, False)
        Dim WriteHeader As New Dictionary(Of String, Boolean)
        WriteHeader.Add("", False)
        'Dim WriteValue As New List(Of String)
        For Each h In d.Values
            With h
                If .Properties.IsRoot = True AndAlso bolWriteRootFields = False Then Continue For
                'If (.Enabled = True AndAlso .Value IsNot Nothing AndAlso .Value <> NullString) = False Then Continue For
                If (.Enabled = False And bolWriteDisabled = False) Then Continue For
                If ((.Value Is Nothing OrElse .Value = NullString) And bolWriteNulls = False) Then Continue For

                'If (.Enabled = True AndAlso .Value IsNot Nothing AndAlso .Value <> NullString) Or (bolWriteDisabled = True And h.Enabled = False) Or (bolWriteNulls And (h.Value Is Nothing OrElse h.Value = NullString)) Then
                If Not WriteHeader.ContainsKey(h.Header.Name) Then WriteHeader.Add(h.Header.Name, False)
                'WriteValue.Add(h.Name)
            End With
        Next


        For Each f In _Headers
            If WriteHeader.ContainsKey(f.Value.Name) Then
                If f.Value.Name <> "" Then fil.WriteLine(Replace(ii.HeaderFormat, "$", f.Value.Name.ToUpper))
                Dim a As Integer = 0
                For Each j In f.Value.Fields
                    If d.ContainsKey(j.Name) Then
                        With d(j.Name)
                            a = 0
                            If .Enabled = False OrElse .Value Is Nothing OrElse .Value = NullString Then
                                fil.Write(ii.RemChar & " ") : a = 2
                            End If
                            fil.WriteLine(.Name & Space(ii.ColumnSpace - Len(.Name) - a) & .Value)
                        End With
                    End If
                Next
                fil.WriteLine()
            End If
        Next
        fil.Close()
    End Sub
    Public Sub Save(ByVal d As ValueDictionary, Optional ByVal bolWriteRootFields As Boolean = True, Optional ByVal bolWriteDisabled As Boolean = True, Optional ByVal bolWriteNulls As Boolean = True)
        Save(Me.strMameIni, d, bolWriteRootFields, bolWriteDisabled, bolWriteNulls)
    End Sub


End Class
