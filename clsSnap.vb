Imports System.Drawing

Public NotInheritable Class Snap
    Public Enum WebSnaps
        ''''' <summary>
        ''''' Local file
        ''''' </summary>
        ''''' <remarks></remarks>
        ''Local 'search Snap directory
        '''' <summary>
        '''' Icon from Maws
        '''' </summary>
        '''' <remarks></remarks>
        '[Icon] 'http://www.mameworld.net/maws/img/ico16_gif/.gif
        ''''' <summary>
        ''''' In game snap from Maws
        ''''' </summary>
        ''''' <remarks></remarks>
        ''MWSnap ' http://www.mameworld.net/maws/img/shots/mwsnap/
        ''''' <summary>
        ''''' Title screen snap from Maws
        ''''' </summary>
        ''''' <remarks></remarks>
        ''Title 'http://www.mameworld.net/maws/img/shots/titles
        ''''' <summary>
        ''''' In game snap from CrashTest
        ''''' </summary>
        ''''' <remarks></remarks>
        ''Snap 'http://www.mameworld.net/maws/img/shots/snap/
        ''''' <summary>
        ''''' Select screen from CrashTest
        ''''' </summary>
        ''''' <remarks></remarks>
        ''[Select] 'http://www.mameworld.net/maws/img/shots/select
        ''''' <summary>
        ''''' In game snap from Enaitz Jar
        ''''' </summary>
        ''''' <remarks></remarks>
        ''EJ 'http://www.mameworld.net/maws/img/shots/ej

        'PSSnap 'http://maws.mameworld.info/img/ps/snap
        'PSTitle 'http://maws.mameworld.info/img/ps/titles
        'PSSelect 'http://maws.mameworld.info/img/ps/select
        'PSScores 'http://maws.mameworld.info/img/ps/scores
        'PSGameOver 'http://maws.mameworld.info/img/ps/gameover/
        'PSBosses 'http://maws.mameworld.info/img/ps/bosses
        'PSVersus 'http://maws.mameworld.info/img/ps/versus/
        'CTSnap 'http://maws.mameworld.info/maws/img/shots/snap/
        'CTTitle 'http://maws.mameworld.info/maws/img/shots/titles
        'CTSelect 'http://maws.mameworld.info/maws/img/shots/select/
        'EJInGame 'http://maws.mameworld.info/maws/img/shots/ej/
        'MrDoArtwork 'http://mrdo.mameworld.info/mame_artwork


        Paradise_Snap '<img class="lazy" src="http://s.emuparadise.org/MAME/snap/umk3.png" data-original="http://s.emuparadise.org/MAME/snap/umk3.png" alt=" Screenshot" style="display: inline;">
        Paradise_Title '<img class="lazy" src="http://s.emuparadise.org/MAME/titles/umk3.png" data-original="http://s.emuparadise.org/MAME/titles/umk3.png" alt=" Title Screen" height="240" width="320" style="display: inline;">
    End Enum
    'Private strRom As String
    'Private sSnapList() As String
    'Private snapIndex As Integer
    'Private strMamePath As String
    'Private strIniPath As String

    'Public Sub New(ByVal strMame As String, ByVal IniPaths As String)
    '    strMamePath = strMame
    '    strIniPath = IniPaths
    'End Sub
    'Public Sub New(ByVal strMame As String)
    '    strMamePath = strMame
    'End Sub
    'Public ReadOnly Property Count() As Integer
    '    Get
    '        If sSnapList IsNot Nothing Then
    '            Return UBound(sSnapList)
    '        Else
    '            Return -1
    '        End If
    '    End Get
    'End Property

    'Public Property RomName() As String
    '    Get
    '        Return strRom
    '    End Get
    '    Set(ByVal value As String)
    '        strRom = value
    '        sSnapList = FoundSnaps(strMamePath, value, strIniPath)
    '        If sSnapList IsNot Nothing Then
    '            snapIndex = 0
    '        Else
    '            snapIndex = -1
    '        End If
    '    End Set
    'End Property
    'Public Property Index() As Integer
    '    Get
    '        Return snapIndex
    '    End Get
    '    Set(ByVal value As Integer)
    '        If sSnapList IsNot Nothing Then
    '            snapIndex = value
    '            If snapIndex > UBound(sSnapList) Then
    '                snapIndex = 0
    '            End If
    '        End If
    '    End Set
    'End Property
    'Public ReadOnly Property List() As String()
    '    Get
    '        Return sSnapList
    '    End Get
    'End Property

    'Public Function RenderLocalSnap(ByVal X As Integer, ByVal Y As Integer, ByVal Width As Integer, ByVal Height As Integer) As Bitmap
    '    If sSnapList IsNot Nothing Then
    '        Return RenderWithAspect(New Bitmap(sSnapList(snapIndex)), X, Y, Width, Height)
    '    Else
    '        Return RenderWithAspect(MAME.App.MameLogo, X, Y, Width, Height)
    '    End If
    'End Function
    'Public Function RenderLocalSnap() As Bitmap
    '    Return New Bitmap(sSnapList(snapIndex))
    'End Function

    'Public Shared Function FoundSnaps(ByVal strMamePath As String, ByVal strRomName As String, ByVal strSnapPath As String) As String()

    '    '    'Dim strSnapPaths() As String = Split(strSnapPath, ";") 'MAME.INI.(strMamePath, INI.IniFields.path_snapshot_directory, True)
    '    '    'Dim strTestFile As String
    '    '    'Dim strFoundSnaps() As String = Nothing

    '    '    'If strSnapPaths IsNot Nothing Then
    '    '    '    'ReDim strFoundSnaps(0)
    '    '    '    For t As Integer = 0 To UBound(strSnapPaths)
    '    '    '        strTestFile = IO.Path.Combine(INI.ExpandPath(strMamePath, strSnapPaths(t)), strRomName & ".png")
    '    '    '        If Len(strTestFile) > 0 And IO.File.Exists(strTestFile) Then
    '    '    '            If strFoundSnaps Is Nothing Then
    '    '    '                ReDim strFoundSnaps(0)
    '    '    '            Else
    '    '    '                ReDim Preserve strFoundSnaps(UBound(strFoundSnaps) + 1)
    '    '    '            End If
    '    '    '            strFoundSnaps(UBound(strFoundSnaps)) = strTestFile
    '    '    '        End If
    '    '    '    Next
    '    '    'End If
    '    '    ''ReDim Preserve strFoundSnaps(UBound(strFoundSnaps) - 1)
    '    '    'Return strFoundSnaps
    '    '    FoundSnaps()
    '    FoundSnaps(MAME.INI.ValueDictionary)
    'End Function
    Public Shared Function FoundSnaps(ByVal MamePath As String, ByVal strExpandedIniPaths As String(), ByVal strRomName As String) As String()
        'Dim strSnapPaths() As String = Split(strExpandedIniPaths, ";") 'MAME.INI.(strMamePath, INI.IniFields.path_snapshot_directory, True)
        Dim l As New List(Of String)
        If strExpandedIniPaths IsNot Nothing Then
            MamePath = IO.Path.GetDirectoryName(MamePath)
            For t As Integer = 0 To UBound(strExpandedIniPaths)
                If Len(strExpandedIniPaths(t)) Then
                    Dim testpath As String = IO.Path.Combine(MamePath, strExpandedIniPaths(t))
                    Dim strTestFile As String = IO.Path.Combine(testpath, strRomName)
                    If IO.Directory.Exists(strTestFile) Then
                        l.AddRange(IO.Directory.GetFiles(strTestFile, "*.png"))
                    End If
                    If IO.Directory.Exists(testpath) Then
                        l.AddRange(IO.Directory.GetFiles(testpath, "_" & strRomName & ".png"))
                        l.AddRange(IO.Directory.GetFiles(testpath, strRomName & ".png"))
                    End If

                End If
            Next
        End If
        Return l.ToArray
    End Function
    Public Shared Function FoundSnaps(ByVal strExpandedIniPaths As String(), ByVal strRomName As String) As String()
        'Dim strSnapPaths() As String = Split(strExpandedIniPaths, ";") 'MAME.INI.(strMamePath, INI.IniFields.path_snapshot_directory, True)
        Dim l As New List(Of String)
        If strExpandedIniPaths IsNot Nothing Then
            'MamePath = IO.Path.GetDirectoryName(MamePath)
            For t As Integer = 0 To UBound(strExpandedIniPaths)
                If Len(strExpandedIniPaths(t)) Then
                    'Dim testpath As String = IO.Path.Combine(strExpandedIniPaths(t), strExpandedIniPaths(t))
                    Dim strTestFile As String = IO.Path.Combine(strExpandedIniPaths(t), strRomName)
                    If IO.Directory.Exists(strTestFile) Then
                        l.AddRange(IO.Directory.GetFiles(strTestFile, "*.png"))
                    End If

                    If IO.Directory.Exists(strExpandedIniPaths(t)) Then
                        l.AddRange(IO.Directory.GetFiles(strExpandedIniPaths(t), "_" & strRomName & ".png"))
                        l.AddRange(IO.Directory.GetFiles(strExpandedIniPaths(t), strRomName & ".png"))
                    End If
                End If
            Next
        End If
        Return l.ToArray
    End Function
    Public Shared Function FoundWebSnaps(strRomName As String) As String()
        Dim s(1)
        's(i + 1) = GetAddress(WebSnaps.MWSnap, strRomname)
        s(0) = GetAddress(WebSnaps.Paradise_Snap, strRomName)
        s(1) = GetAddress(WebSnaps.Paradise_Title, strRomName)
        's(i + 3) = GetAddress(WebSnaps.PSBosses, strRomname)
        's(i + 4) = GetAddress(WebSnaps.PSGameOver, strRomname)
        's(i + 5) = GetAddress(WebSnaps.PSScores, strRomname)
        's(i + 6) = GetAddress(WebSnaps.PSSnap, strRomname)
        's(i + 7) = GetAddress(WebSnaps.PSVersus, strRomname)
        's(i + 8) = GetAddress(WebSnaps.CTSnap, strRomname)
        's(i + 9) = GetAddress(WebSnaps.CTSelect, strRomname)
        's(i + 10) = GetAddress(WebSnaps.CTTitle, strRomname)
        's(i + 11) = GetAddress(WebSnaps.EJInGame, strRomname)
        's(i + 12) = GetAddress(WebSnaps.MrDoArtwork, strRomname)

        Return s
    End Function
    Public Shared Function FoundSnapsWithWebImages(ByVal MamePath As String, ByVal strExpandedIniPaths As String(), ByVal strRomname As String) As String()
        Dim s() As String
        s = FoundSnaps(MamePath, strExpandedIniPaths, strRomname)
        Dim i As Integer = UBound(s)
        ReDim Preserve s(i + 2)
        's(i + 1) = GetAddress(WebSnaps.MWSnap, strRomname)
        s(i + 1) = GetAddress(WebSnaps.Paradise_Snap, strRomname)
        s(i + 2) = GetAddress(WebSnaps.Paradise_Title, strRomname)
        's(i + 3) = GetAddress(WebSnaps.PSBosses, strRomname)
        's(i + 4) = GetAddress(WebSnaps.PSGameOver, strRomname)
        's(i + 5) = GetAddress(WebSnaps.PSScores, strRomname)
        's(i + 6) = GetAddress(WebSnaps.PSSnap, strRomname)
        's(i + 7) = GetAddress(WebSnaps.PSVersus, strRomname)
        's(i + 8) = GetAddress(WebSnaps.CTSnap, strRomname)
        's(i + 9) = GetAddress(WebSnaps.CTSelect, strRomname)
        's(i + 10) = GetAddress(WebSnaps.CTTitle, strRomname)
        's(i + 11) = GetAddress(WebSnaps.EJInGame, strRomname)
        's(i + 12) = GetAddress(WebSnaps.MrDoArtwork, strRomname)

        Return s
    End Function
    Public Shared Function FoundSnapsWithWebImages(ByVal strExpandedIniPaths As String(), ByVal strRomname As String) As String()
        Dim s() As String
        s = FoundSnaps(strExpandedIniPaths, strRomname)
        Dim i As Integer = UBound(s)
        ReDim Preserve s(i + 2)
        's(i + 1) = GetAddress(WebSnaps.MWSnap, strRomname)
        s(i + 1) = GetAddress(WebSnaps.Paradise_Snap, strRomname)
        s(i + 2) = GetAddress(WebSnaps.Paradise_Title, strRomname)
        's(i + 3) = GetAddress(WebSnaps.PSBosses, strRomname)
        's(i + 4) = GetAddress(WebSnaps.PSGameOver, strRomname)
        's(i + 5) = GetAddress(WebSnaps.PSScores, strRomname)

        's(i + 6) = GetAddress(WebSnaps.PSSnap, strRomname)
        's(i + 7) = GetAddress(WebSnaps.PSVersus, strRomname)
        's(i + 8) = GetAddress(WebSnaps.CTSnap, strRomname)
        's(i + 9) = GetAddress(WebSnaps.CTSelect, strRomname)
        's(i + 10) = GetAddress(WebSnaps.CTTitle, strRomname)
        's(i + 11) = GetAddress(WebSnaps.EJInGame, strRomname)
        's(i + 12) = GetAddress(WebSnaps.MrDoArtwork, strRomname)
        Return s
    End Function

    Public Shared Function GetAddress(ByVal index As WebSnaps, ByVal strRomName As String) As String
        Select Case index
            Case WebSnaps.Paradise_Snap
                Return "http://s.emuparadise.org/MAME/snap/" & strRomName & ".png"
            Case WebSnaps.Paradise_Title
                Return "http://s.emuparadise.org/MAME/titles/" & strRomName & ".png"


                ''Case WebSnaps.MWSnap
                ''    'Return "http://www.mameworld.net/maws/img/shots/mwsnap/" & strRomName & ".png"
                ''    Return "http://maws.mameworld.info/maws/img/shots/progettosnaps/ingame/" & strRomName & ".png"
                ''Case WebSnaps.Title
                ''    'Return "http://www.mameworld.net/maws/img/shots/titles/" & strRomName & ".png"
                ''    Return "http://maws.mameworld.info/maws/img/shots/progettosnaps/titles/" & strRomName & ".png"
                ''Case WebSnaps.Snap
                ''    'Return "http://www.mameworld.net/mas/img/shots/snap/" & strRomName & ".png"
                ''    Return "http://maws.mameworld.info/maws/img/shots/progettosnaps/select/" & strRomName & ".png"
                ''Case WebSnaps.EJ
                ''    Return "http://www.mameworld.net/maws/img/shots/ej/" & strRomName & ".png"
                'Case WebSnaps.Icon
                '    Return "http://maws.mameworld.info/maws/img/ico16_gif/" & strRomName & ".gif"
                'Case WebSnaps.PSSnap
                '    Return "http://maws.mameworld.info/img/ps/snap/" & strRomName & ".png"

                'Case WebSnaps.PSTitle
                '    Return "http://maws.mameworld.info/img/ps/titles/" & strRomName & ".png"
                'Case WebSnaps.PSSelect
                '    Return "http://maws.mameworld.info/img/ps/select/" & strRomName & ".png"
                'Case WebSnaps.PSScores
                '    Return "http://maws.mameworld.info/img/ps/scores/" & strRomName & ".png"
                'Case WebSnaps.PSGameOver
                '    Return "http://maws.mameworld.info/img/ps/gameover/" & strRomName & ".png"
                'Case WebSnaps.PSBosses
                '    Return "http://maws.mameworld.info/img/ps/bosses/" & strRomName & ".png"
                'Case WebSnaps.PSVersus
                '    Return "http://maws.mameworld.info/img/ps/versus/" & strRomName & ".png"
                'Case WebSnaps.CTSnap
                '    Return "http://maws.mameworld.info/maws/img/shots/snap/" & strRomName & ".png"
                'Case WebSnaps.CTTitle
                '    Return "http://maws.mameworld.info/maws/img/shots/titles/" & strRomName & ".png"
                'Case WebSnaps.CTSelect
                '    Return "http://maws.mameworld.info/maws/img/shots/select/" & strRomName & ".png"
                'Case WebSnaps.EJInGame
                '    Return "http://maws.mameworld.info/maws/img/shots/ej/" & strRomName & ".png"
                'Case WebSnaps.MrDoArtwork
                '    Return "http://mrdo.mameworld.info/mame_artwork/" & strRomName & ".png"
                '    'Case WebSnaps.Select
                '    '    Return "http://www.mameworld.net/maws/img/shots/select/" & strRomName & ".png"
        End Select
        Return Nothing
    End Function
    Public Shared Function GetAddress(ByVal i As WebSnaps) As String
        Select Case i
            Case WebSnaps.Paradise_Snap
                Return "http://s.emuparadise.org/MAME/snap/"
            Case WebSnaps.Paradise_Title
                Return "http://s.emuparadise.org/MAME/titles/"
                ''Case WebSnaps.MWSnap
                ''    'Return "http://www.mameworld.net/maws/img/shots/mwsnap/" & strRomName & ".png"
                ''    Return "http://maws.mameworld.info/maws/img/shots/progettosnaps/ingame/"
                ''    'Case WebSnaps.Title
                ''    'Return "http://www.mameworld.net/maws/img/shots/titles/" & strRomName & ".png"
                ''    Return "http://maws.mameworld.info/maws/img/shots/progettosnaps/titles/" & strRomName & ".png"
                ''Case WebSnaps.Snap
                ''    'Return "http://www.mameworld.net/mas/img/shots/snap/" & strRomName & ".png"
                ''    Return "http://maws.mameworld.info/maws/img/shots/progettosnaps/select/" & strRomName & ".png"
                ''Case WebSnaps.EJ
                ''    Return "http://www.mameworld.net/maws/img/shots/ej/" & strRomName & ".png"
                'Case WebSnaps.Icon
                '    Return "http://maws.mameworld.info/maws/img/ico16_gif/"
                'Case WebSnaps.PSSnap
                '    Return "http://maws.mameworld.info/img/ps/snap/"

                'Case WebSnaps.PSTitle
                '    Return "http://maws.mameworld.info/img/ps/titles/"
                'Case WebSnaps.PSSelect
                '    Return "http://maws.mameworld.info/img/ps/select/"
                'Case WebSnaps.PSScores
                '    Return "http://maws.mameworld.info/img/ps/scores/"
                'Case WebSnaps.PSGameOver
                '    Return "http://maws.mameworld.info/img/ps/gameover/"
                'Case WebSnaps.PSBosses
                '    Return "http://maws.mameworld.info/img/ps/bosses/"
                'Case WebSnaps.PSVersus
                '    Return "http://maws.mameworld.info/img/ps/versus/"
                'Case WebSnaps.CTSnap
                '    Return "http://maws.mameworld.info/maws/img/shots/snap/"
                'Case WebSnaps.CTTitle
                '    Return "http://maws.mameworld.info/maws/img/shots/titles/"
                'Case WebSnaps.CTSelect
                '    Return "http://maws.mameworld.info/maws/img/shots/select/"
                'Case WebSnaps.EJInGame
                '    Return "http://maws.mameworld.info/maws/img/shots/ej/"
                'Case WebSnaps.MrDoArtwork
                '    Return "http://mrdo.mameworld.info/mame_artwork/"
                '    'Case WebSnaps.Select
                '    '    Return "http://www.mameworld.net/maws/img/shots/select/" & strRomName & ".png"
        End Select
        Return Nothing
    End Function
    Public Shared Function LoadWebSnap(ByVal strRomName As String, ByVal intIndex As WebSnaps) As Image
        Dim s As String = GetAddress(intIndex, strRomName)
        Return LoadImage(s)
    End Function
    Public Shared Sub LoadWebSnap(ByVal strRomName As String, ByVal intIndex As WebSnaps, ByVal Async As AsyncCallback)
        Dim s As String = GetAddress(intIndex, strRomName)
        LoadImage(s, Async)
    End Sub
    'Public Shared Function KeepAspect(ByVal source As Size, ByVal dest As SizeF, Optional ByVal bolStretch As Boolean = True) As RectangleF
    '    Dim ScalePercent As Single, addX As Integer, addY As Integer, NewWidth As Integer, NewHeight As Integer
    '    'With cImages(Index)
    '    If source.Width > dest.Width Or source.Height > dest.Height Then
    '        'If bolStretch Then
    '        If (source.Width > source.Height) Then
    '            ScalePercent = source.Width / dest.Width
    '        Else
    '            ScalePercent = source.Height / dest.Height
    '        End If
    '        NewWidth = Math.Floor(source.Width / ScalePercent)
    '        NewHeight = Math.Floor(source.Height / ScalePercent)
    '        addX = Math.Floor((dest.Width / 2) - (NewWidth / 2))
    '        addY = Math.Floor((dest.Height / 2) - (NewHeight / 2))
    '        'Else
    '        '    NewWidth = source.Width
    '        '    NewHeight = source.Height
    '        '    addX = Math.Floor((source.Width / 2) - (dest.Width / 2))
    '        '    addY = Math.Floor((source.Height / 2) - (dest.Height / 2))
    '        'End If

    '    Else
    '        If bolStretch Then
    '            If (source.Width > source.Height) Then
    '                ScalePercent = source.Width / dest.Width
    '            Else
    '                ScalePercent = source.Height / dest.Height
    '            End If
    '            NewWidth = Math.Floor(source.Width / ScalePercent)
    '            NewHeight = Math.Floor(source.Height / ScalePercent)
    '            addX = Math.Floor((dest.Width / 2) - (NewWidth / 2))
    '            addY = Math.Floor((dest.Height / 2) - (NewHeight / 2))
    '        Else
    '            NewWidth = source.Width
    '            NewHeight = source.Height
    '            addX = Math.Floor((dest.Width / 2) - (source.Width / 2))
    '            addY = Math.Floor((dest.Height / 2) - (source.Height / 2))
    '        End If
    '    End If
    '    KeepAspect = New RectangleF(addX, addY, NewWidth, NewHeight)
    '    If KeepAspect.Width = 0 Then KeepAspect.Width = 1
    '    If KeepAspect.Height = 0 Then KeepAspect.Height = 1
    'End Function
    'Public Shared Function KeepAspect(ByVal source As SizeF, ByVal dest As SizeF, Optional ByVal bolStretch As Boolean = True) As RectangleF

    '    Dim ScalePercent As Single, addX As Single, addY As Single, NewWidth As Single, NewHeight As Single
    '    'With cImages(Index)
    '    If source.Width > dest.Width Or source.Height > dest.Height Then
    '        'If bolStretch Then
    '        If (source.Width > source.Height) Then
    '            ScalePercent = source.Width / dest.Width
    '        Else
    '            ScalePercent = source.Height / dest.Height
    '        End If
    '        NewWidth = source.Width / ScalePercent
    '        NewHeight = source.Height / ScalePercent
    '        addX = (dest.Width / 2) - (NewWidth / 2)
    '        addY = (dest.Height / 2) - (NewHeight / 2)
    '        'Else
    '        '    NewWidth = source.Width
    '        '    NewHeight = source.Height
    '        '    addX = Math.Floor((source.Width / 2) - (dest.Width / 2))
    '        '    addY = Math.Floor((source.Height / 2) - (dest.Height / 2))
    '        'End If

    '    Else
    '        If bolStretch Then
    '            If (source.Width > source.Height) Then
    '                ScalePercent = source.Width / dest.Width
    '            Else
    '                ScalePercent = source.Height / dest.Height
    '            End If
    '            NewWidth = source.Width / ScalePercent
    '            NewHeight = source.Height / ScalePercent
    '            addX = (dest.Width / 2) - (NewWidth / 2)
    '            addY = (dest.Height / 2) - (NewHeight / 2)
    '        Else
    '            NewWidth = source.Width
    '            NewHeight = source.Height
    '            addX = (dest.Width / 2) - (source.Width / 2)
    '            addY = (dest.Height / 2) - (source.Height / 2)
    '        End If
    '    End If
    '    KeepAspect = New RectangleF(addX, addY, NewWidth, NewHeight)
    '    'If KeepAspect.Width = 0.0F Then KeepAspect.Width = 1.0F
    '    'If KeepAspect.Height = 0.0F Then KeepAspect.Height = 1.0F
    'End Function


End Class