Imports System.Data.SqlClient
Imports System.IO


Public Class Database
    Implements IDisposable

    Public NotInheritable Class SearchHelper
        Implements ICloneable

        Public Shared Function StripSql(ByVal strIn As String) As String
            If strIn.Contains("'") Then
                Dim first = InStr(strIn, "'", CompareMethod.Binary) + 1
                Dim last = InStr(first, strIn, "'", CompareMethod.Binary) - 1
                'the contents of this string should be preserved lest %
                Return Mid(strIn, first, last - first).Replace("%", "").Trim
            Else
                Return strIn.Trim("=", ">", "<", "LIKE", "like", "not", "NOT", "!")
            End If
        End Function

        Public NotInheritable Class TableCollection
            Inherits System.Collections.ObjectModel.KeyedCollection(Of String, Table)
            Implements ICloneable

            Protected Overrides Function GetKeyForItem(ByVal item As Table) As String
                Return item.Name
            End Function
            Default Public Overloads ReadOnly Property Item(ByVal Table As String, ByVal Field As String) As Field
                Get
                    Return MyBase.Item(Table)(Field)
                End Get
            End Property
            Public Overloads Sub Add(ByVal Table As String)
                MyBase.Add(New Table(Table))
            End Sub

            Public Function Clone() As Object Implements System.ICloneable.Clone
                Dim i As New TableCollection
                For Each e In Me
                    i.Add(e)
                Next
                Return i
            End Function
        End Class
        Public NotInheritable Class FieldCollection
            Inherits System.Collections.ObjectModel.KeyedCollection(Of String, Field)
            Implements ICloneable

            Protected Overrides Function GetKeyForItem(ByVal item As Field) As String
                Return item.Name
            End Function
            ReadOnly sParent As String
            Public Sub New(ByVal Parent As String)
                sParent = Parent
            End Sub
            Public ReadOnly Property Parent() As String
                Get
                    Return sParent
                End Get
            End Property
            Public Overloads Sub Add(ByVal Field As String, Optional ByVal sAlias As String = "", Optional ByVal Sort As SortOrder = SortOrder.Unspecified, Optional ByVal bShow As Boolean = True)
                MyBase.Add(New Field(Field, sParent) With {.Alias = sAlias, .Sort = Sort, .Show = bShow})
            End Sub

            Public Function Clone() As Object Implements System.ICloneable.Clone
                Dim i As New FieldCollection(sParent)
                For Each e In Me
                    i.Add(e)
                Next
                Return i
            End Function
        End Class
        Public NotInheritable Class JoinTableCollection
            Inherits System.Collections.ObjectModel.KeyedCollection(Of String, JoinTable)
            Implements ICloneable

            Protected Overrides Function GetKeyForItem(ByVal item As JoinTable) As String
                Return item.Name
            End Function
            Private ReadOnly sParent As String
            Public ReadOnly Property Parent() As String
                Get
                    Return sParent
                End Get
            End Property
            Public Sub New(ByVal Parent As String)
                sParent = Parent
            End Sub
            Public Overloads Sub Add(ByVal name As String, ByVal PrimaryKey As String)
                MyBase.Add(New JoinTable(name, sParent, PrimaryKey))
            End Sub

            Public Function Clone() As Object Implements System.ICloneable.Clone
                Dim i As New JoinTableCollection(sParent)
                For Each e In Me
                    i.Add(e)
                Next
                Return i
            End Function
        End Class

        Public Class Table
            Implements ICloneable

            Private fFields As FieldCollection
            Private fJoinTables As JoinTableCollection
            Private sName As String
            Public Property Fields() As FieldCollection
                Get
                    Return fFields
                End Get
                Set(ByVal value As FieldCollection)
                    fFields = value
                End Set
            End Property
            Default Public ReadOnly Property Item(ByVal Key As String) As Field
                Get
                    Return fFields(Key)
                End Get
            End Property
            Default Public ReadOnly Property Item(ByVal Index As Integer) As Field
                Get
                    Return fFields(Index)
                End Get
            End Property

            Public Property JoinTables() As JoinTableCollection
                Get
                    Return fJoinTables
                End Get
                Set(ByVal value As JoinTableCollection)
                    fJoinTables = value
                End Set
            End Property
            Public Property Name() As String
                Get
                    Return sName
                End Get
                Set(ByVal value As String)
                    sName = value
                End Set
            End Property
            Friend Sub New(ByVal sName As String)
                fFields = New FieldCollection(sName)
                fJoinTables = New JoinTableCollection(sName)
                Me.sName = sName
            End Sub

            Public Function Clone() As Object Implements System.ICloneable.Clone
                Dim i As New Table(sName)
                With Me
                    i.fFields = .fFields.Clone
                    i.fJoinTables = .fJoinTables.Clone
                End With
                Return i
            End Function
        End Class
        Public Class Field
            Implements ICloneable

            Private cCondition As Condition
            Private sAlias As String
            'Private pParentTable As Table
            Private sSort As SortOrder = SortOrder.Unspecified
            Private sName As String
            Private bShow As Boolean = True
            Public Property [Alias]() As String
                Get
                    Return sAlias
                End Get
                Set(ByVal value As String)
                    If value = "" Then value = Nothing
                    sAlias = value
                End Set
            End Property
            'Public ReadOnly Property ParentTable() As Table
            '    Get
            '        Return pParentTable
            '    End Get
            'End Property
            Public Property Name() As String
                Get
                    Return sName
                End Get
                Set(ByVal value As String)
                    sName = value
                End Set
            End Property
            Public Property Conditions() As Condition
                Get
                    Return cCondition
                End Get
                Set(ByVal value As Condition)
                    cCondition = value
                End Set
            End Property
            Default Public ReadOnly Property Condition(ByVal strKey As String) As String()
                Get
                    Return cCondition(strKey)
                End Get
            End Property
            Public Property Sort() As SortOrder
                Get
                    Return sSort
                End Get
                Set(ByVal value As SortOrder)
                    sSort = value
                End Set
            End Property
            Public Property Show() As Boolean
                Get
                    Return bShow
                End Get
                Set(ByVal value As Boolean)
                    bShow = value
                End Set
            End Property
            Friend Sub New(ByVal name As String, ByVal parent As String)
                cCondition = New Condition(parent, name)
                sName = name
            End Sub

            Public Function Clone() As Object Implements System.ICloneable.Clone
                Dim i As New Field(Me.Conditions.ParentTable, Me.Conditions.ParentField)
                With Me
                    i.bShow = i.bShow
                    i.cCondition = Me.cCondition.Clone
                    i.sAlias = i.sAlias
                    i.sSort = i.sSort
                End With
                Return i
            End Function
        End Class
        Public Class JoinTable
            Implements ICloneable

            Private fFields As FieldCollection
            Private sName As String

            Private sParentLink As String
            Private sAlias As String
            Default Public ReadOnly Property Item(ByVal key As String) As Field
                Get
                    Return fFields(key)
                End Get
            End Property
            Default Public ReadOnly Property Item(ByVal Index As Integer) As Field
                Get
                    Return fFields(Index)
                End Get
            End Property
            Public Property Fields() As FieldCollection
                Get
                    Return fFields
                End Get
                Set(ByVal value As FieldCollection)
                    fFields = value
                End Set
            End Property
            Public ReadOnly Property Name() As String
                Get
                    Return sName
                End Get

            End Property
            Public Property ParentColumn() As String
                Get
                    Return sParentLink
                End Get
                Set(ByVal value As String)
                    sParentLink = value
                End Set
            End Property
            Public Property [Alias]() As String
                Get
                    Return sAlias
                End Get
                Set(ByVal value As String)
                    sAlias = value
                End Set
            End Property
            Friend Sub New(ByVal Name As String, ByVal parent As String, ByVal sIDKey As String)
                sParentLink = sIDKey
                sName = Name
                sAlias = Name
                fFields = New FieldCollection(Name)
                fFields.Add(New Field(sIDKey, Name) With {.Show = False})
            End Sub

            Public Function Clone() As Object Implements System.ICloneable.Clone
                Dim i As New JoinTable(sName, sParentLink, Me.fFields.Parent)
                Return i
            End Function
        End Class
        Public Class Condition
            Implements ICloneable

            Public Class MultiKeyedCollection
                'Dim internal As New Dictionary(Of String, String())
                Inherits Dictionary(Of String, String())
                Implements ICloneable
                Friend Sub New()

                End Sub
                Friend Sub New(ByVal dic As Dictionary(Of String, String()))
                    MyBase.New(dic)
                End Sub
                Public Overloads Sub Add(ByVal key As String, ByVal value As String)
                    If ContainsKey(key) Then
                        Dim i = MyBase.Item(key)
                        ReDim Preserve i(UBound(i) + 1)
                        i(UBound(i)) = value
                        MyBase.Item(key) = i
                    Else
                        MyBase.Add(key, New String() {value})
                    End If
                End Sub


                Public Function Clone() As Object Implements System.ICloneable.Clone
                    Return New MultiKeyedCollection(Me)
                End Function
            End Class

            Private ReadOnly pParent As String
            Public ReadOnly Property ParentTable() As String
                Get
                    Return pParent
                End Get

            End Property
            Private ReadOnly fParentField As String
            Public ReadOnly Property ParentField() As String
                Get
                    Return fParentField
                End Get
            End Property

            'Private sORs() As String, sANDs() As String
            Private _sANDs, _sORs As New MultiKeyedCollection
            'Public Sub Add(ByVal strCondition As String, ByVal ParamArray sOr() As String)
            '    'AddOr(strCondition)
            '    'For Each s As String In sOr
            '    '    AddOr(s)
            '    'Next
            '    Add("", strCondition, sOr)
            'End Sub
            Public Sub Add(ByVal strConditionName As String, ByVal strCondition As String, ByVal ParamArray sOr() As String)
                'AddOr (strCondition )
                _sORs.Add(strConditionName, strCondition)
                For Each s As String In sOr
                    _sORs.Add(strConditionName, s)
                Next
            End Sub
            Public Sub AddOr(ByVal strConditionName As String, ByVal sOr As String)
                _sORs.Add(strConditionName, sOr)
            End Sub
            Public Sub AddOr(ByVal sOr As String)
                AddOr("", sOr)
            End Sub
            Public Sub AddAnd(ByVal strConditionName As String, ByVal sAnd As String)
                _sANDs.Add(strConditionName, sAnd)
            End Sub
            Public Sub AddAnd(ByVal sAnd As String)
                AddAnd("", sAnd)
            End Sub
            'Public Sub AddOr(ByVal sOr As String)
            '    If sORs Is Nothing Then ReDim sORs(0) Else ReDim Preserve sORs(UBound(sORs) + 1)
            '    sORs(UBound(sORs)) = sOr
            'End Sub

            'Public Sub AddAnd(ByVal sAnd As String)
            '    If sANDs Is Nothing Then ReDim sANDs(0) Else ReDim Preserve sANDs(UBound(sANDs) + 1)
            '    sANDs(UBound(sANDs)) = sAnd
            'End Sub
            'Default Public ReadOnly Property Item(ByVal index As Integer) As String
            '    Get
            '        If _sORs IsNot Nothing Then
            '            If index > _sORs.Count - 1 Then
            '                'Return sANDs(UBound(sORs) + 1 - index)
            '                Return _sANDs (
            '            Else
            '                Return sORs(index)
            '            End If
            '        ElseIf sANDs IsNot Nothing Then
            '            Return sANDs(index)
            '        End If
            '        Return Nothing
            '    End Get
            'End Property
            Public ReadOnly Property Count() As Integer
                Get
                    Dim ret As Integer
                    'If sORs IsNot Nothing Then ret += UBound(sORs) + 1
                    'If sANDs IsNot Nothing Then ret += UBound(sANDs) + 1
                    If _sORs.Count Then ret += _sORs.Count
                    If _sANDs.Count Then ret += _sANDs.Count
                    Return ret
                End Get
            End Property
            Public Overrides Function ToString() As String
                'If sORs Is Nothing AndAlso sANDs Is Nothing Then Return ""
                If _sORs.Count = 0 AndAlso _sANDs.Count = 0 Then Return ""
                Dim ret As String = ""
                Dim s() As String = Nothing
                'If sORs IsNot Nothing Then
                If _sORs.Count > 0 Then
                    ret = "("
                    For Each x In _sORs
                        For Each xs In x.Value
                            If s Is Nothing Then ReDim s(0) Else ReDim Preserve s(UBound(s) + 1)
                            s(UBound(s)) = "[" & ParentTable & "].[" & ParentField & "] " & xs
                        Next
                    Next
                    ret &= Join(s, " OR ") & ")"
                End If
                Erase s
                'If sANDs IsNot Nothing Then
                If _sANDs.Count > 0 Then
                    If ret = "" Then ret = "(" Else ret &= " AND ("
                    For Each x In _sANDs
                        For Each xs In x.Value
                            If s Is Nothing Then ReDim s(0) Else ReDim Preserve s(UBound(s) + 1)
                            s(UBound(s)) = "[" & ParentTable & "].[" & ParentField & "] " & xs
                        Next
                    Next
                    ret &= Join(s, ") AND (") & ")"
                End If
                Return ret
            End Function
            Public Function SimpleString() As String
                Dim ret As String = ""
                'If sORs IsNot Nothing Then
                If _sORs.Count > 0 Then
                    For Each s In _sORs
                        'ret = Join(sORs)
                        ret &= Join(s.Value)
                    Next
                End If
                'If sANDs IsNot Nothing Then
                If _sANDs.Count > 0 Then
                    'ret &= Join(sANDs)
                    For Each s In _sANDs
                        ret &= Join(s.Value)
                    Next
                End If
                Return ret
            End Function

            Default Public Property Item(ByVal strKey As String) As String()
                Get
                    Dim s As String() = Nothing
                    If _sORs.ContainsKey(strKey) Then
                        For Each st In _sORs
                            For Each sz In st.Value
                                If s Is Nothing Then ReDim s(0) Else ReDim Preserve s(UBound(s) + 1)
                                s(UBound(s)) = sz
                            Next
                        Next

                    End If
                    If _sANDs.ContainsKey(strKey) Then
                        For Each st In _sANDs
                            For Each sz In st.Value
                                If s Is Nothing Then ReDim s(0) Else ReDim Preserve s(UBound(s) + 1)
                                s(UBound(s)) = sz
                            Next
                        Next
                    End If
                    Return s
                End Get
                Set(ByVal value As String())

                End Set
            End Property

            Public ReadOnly Property HasConditions() As Boolean
                Get
                    'If sORs Is Nothing AndAlso sANDs Is Nothing Then Return False
                    If _sORs.Count = 0 AndAlso _sANDs.Count = 0 Then Return False
                    Return True
                End Get
            End Property
            Public ReadOnly Property HasCondition(ByVal strKey As String) As Boolean
                Get
                    If _sORs.ContainsKey(strKey) Then Return True
                    If _sANDs.ContainsKey(strKey) Then Return True
                    Return False
                End Get
            End Property
            Public Sub Remove(ByVal strKey As String)
                If _sORs.ContainsKey(strKey) Then _sORs.Remove(strKey)
                If _sANDs.ContainsKey(strKey) Then _sANDs.Remove(strKey)
            End Sub
            Friend Sub New(ByVal Parent As String, ByVal Field As String)
                pParent = Parent
                fParentField = Field
            End Sub
            Public Sub Clear()
                'Erase sORs
                'Erase sANDs
                _sORs.Clear()
                _sANDs.Clear()
            End Sub

            Public Function Clone() As Object Implements System.ICloneable.Clone
                Dim i As New Condition(pParent, fParentField)
                With Me
                    'i.sANDs = .sANDs.Clone
                    'i.sORs = .sORs.Clone
                    i._sANDs = ._sANDs.Clone
                    i._sORs = ._sORs.Clone
                End With
                Return i
            End Function
        End Class

        Dim rRoot As New TableCollection
        Dim bDistinct As Boolean = True
        Public Property Distinct() As Boolean
            Get
                Return bDistinct
            End Get
            Set(ByVal value As Boolean)
                bDistinct = value
            End Set
        End Property
        Public Property RootTables() As TableCollection
            Get
                Return rRoot
            End Get
            Set(ByVal value As TableCollection)
                rRoot = value
            End Set
        End Property

        Default Public ReadOnly Property Item(ByVal key As String) As Table
            Get
                Return rRoot(key)
            End Get
        End Property
        Default Public ReadOnly Property Item(ByVal index As Integer) As Table
            Get
                Return rRoot(index)
            End Get
        End Property
        Default Public ReadOnly Property Item(ByVal Table As String, ByVal Field As String) As Field
            Get
                Return rRoot(Table)(Field)
            End Get
        End Property
        Default Public ReadOnly Property Item(ByVal Table As String, ByVal JoinTable As String, ByVal Field As String) As Field
            Get
                Return rRoot(Table).JoinTables(JoinTable)(Field)
            End Get
        End Property
        Public Function CreateStatement() As String
            Dim SELECTs() As String = Nothing
            Dim JOINs() As String = Nothing
            Dim WHEREs() As String = Nothing
            Dim ORDERs() As String = Nothing
            Dim ROOTs() As String = Nothing
            For Each r In rRoot


                'Dim temp As String
                For Each h In r.Fields
                    'temp = ""
                    If h.Show Then
                        If SELECTs Is Nothing Then ReDim SELECTs(0) Else ReDim Preserve SELECTs(UBound(SELECTs) + 1)
                        SELECTs(UBound(SELECTs)) = "[" & r.Name & "].[" & h.Name & "]"
                        If h.Alias IsNot Nothing Then
                            SELECTs(UBound(SELECTs)) &= " AS [" & h.Alias & "]"
                        End If
                    End If
                    'If h.Conditions.Count Then
                    'For Each t In h.Conditions
                    If h.Conditions.HasConditions Then
                        If WHEREs Is Nothing Then ReDim WHEREs(0) Else ReDim Preserve WHEREs(UBound(WHEREs) + 1)
                        WHEREs(UBound(WHEREs)) = h.Conditions.ToString
                    End If
                    'Next
                    'End If
                    If h.Sort <> SortOrder.Unspecified Then
                        If ORDERs Is Nothing Then ReDim ORDERs(0) Else ReDim Preserve ORDERs(UBound(ORDERs) + 1)
                        ORDERs(UBound(ORDERs)) = "[" & r.Name & "].[" & h.Name & "]"
                        If h.Sort = SortOrder.Descending Then ORDERs(UBound(ORDERs)) &= " DESC"
                    End If
                Next

                For Each h In r.JoinTables
                    Dim j() As String = Nothing

                    For Each i In h.Fields
                        If j Is Nothing Then ReDim j(0) Else ReDim Preserve j(UBound(j) + 1)
                        If i.Show Then
                            If SELECTs Is Nothing Then ReDim SELECTs(0) Else ReDim Preserve SELECTs(UBound(SELECTs) + 1)
                            SELECTs(UBound(SELECTs)) = "[" & h.Name & "].[" & i.Name & "]"
                            If i.Alias IsNot Nothing Then
                                SELECTs(UBound(SELECTs)) &= " AS [" & i.Alias & "]"
                            End If
                        End If

                        j(UBound(j)) = i.Name
                        If i.Alias IsNot Nothing Then
                            j(UBound(j)) &= " AS " & i.Alias
                        End If
                        If i.Conditions.HasConditions Then
                            'For Each t In i.Conditions
                            If WHEREs Is Nothing Then ReDim WHEREs(0) Else ReDim Preserve WHEREs(UBound(WHEREs) + 1)
                            WHEREs(UBound(WHEREs)) = i.Conditions.ToString
                            'Next
                        End If
                        If i.Sort <> SortOrder.Unspecified Then
                            If ORDERs Is Nothing Then ReDim ORDERs(0) Else ReDim Preserve ORDERs(UBound(ORDERs) + 1)
                            ORDERs(UBound(ORDERs)) = "[" & h.Name & "].[" & i.Name & "]"
                            If i.Sort = SortOrder.Descending Then ORDERs(UBound(ORDERs)) &= " DESC"
                        End If
                    Next
                    If JOINs Is Nothing Then ReDim JOINs(0) Else ReDim Preserve JOINs(UBound(JOINs) + 1)
                    JOINs(UBound(JOINs)) = " INNER JOIN (SELECT " & Join(j, ", ") & " FROM " & h.Name & " AS " & h.Name & "_1) AS " & h.Alias & " ON [" & r.Name & "].[" & h.ParentColumn & "]=[" & h.Name & "].[" & h.ParentColumn & "]"
                Next

                If ROOTs Is Nothing Then ReDim ROOTs(0) Else ReDim Preserve ROOTs(UBound(ROOTs) + 1)
                ROOTs(UBound(ROOTs)) = "[" & r.Name & "]"
            Next
            Dim s As New Text.StringBuilder()
            s.Append("SELECT " & IIf(bDistinct, "DISTINCT ", ""))

            s.Append(Join(SELECTs, ", ") & " FROM " & Join(ROOTs, ", "))
            If JOINs IsNot Nothing Then s.Append(Join(JOINs, ""))
            If WHEREs IsNot Nothing Then s.Append(" WHERE " & Join(WHEREs, " AND "))
            If ORDERs IsNot Nothing Then s.Append(" ORDER BY " & Join(ORDERs, ", "))

            Return s.ToString
        End Function
        Public Function Execute() As DataTable
            Dim d As New DataTable

            Dim s As String '= "select "
            s = CreateStatement()
            Debug.Print(s)
            Dim c As New SqlCommand(s, OpenDatabase(DatabaseName))

            'Try
            d.Load(c.ExecuteReader)
            'Catch ex As Exception
            '    MsgBox(ex.Message)
            '    Return Nothing
            'Finally
            c.Dispose()
            'End Try
            Return d
        End Function
        Public Function ContainsKey(ByVal Table As String, ByVal Field As String) As Boolean
            If Not rRoot.Contains(Table) Then Return False
            Return rRoot(Table).Fields.Contains(Field)
        End Function
        Public Function Remove(ByVal Table As String, ByVal Field As String) As Boolean
            Return rRoot(Table).Fields.Remove(Field)
        End Function

        Public Function Clone() As Object Implements System.ICloneable.Clone
            Dim i As New SearchHelper
            i.Distinct = Me.Distinct
            i.rRoot = Me.rRoot.Clone
            Return i
        End Function
    End Class

    Shared strMameExe As String
    Shared strBuild As String
    Shared Sub New()
        Try
            _New()
        Catch
        End Try
        'TODO: should close database??
    End Sub
    Private Shared Sub _New()
        If NamespaceExists() AndAlso TablesExists Then
            Dim ca As New SqlCommand("select * from mame", OpenDatabase(DatabaseName))
            Dim reader = ca.ExecuteReader
            reader.Read()
            strMameExe = reader.Item("mamepath").ToString
            strBuild = reader.Item("build").ToString
            reader.Close()
            ca.Dispose()
        End If
    End Sub
    Public Shared ReadOnly Property TablesExists() As Boolean
        Get
            Try
                Dim cmd As New SqlCommand("select name from sysobjects where xtype='u' and name = 'mame'", OpenDatabase("Menulator"))
                TablesExists = Not (cmd.ExecuteScalar) = ""
                cmd.Dispose()
            Catch
                Return False
            End Try
        End Get
    End Property
    Public Shared ReadOnly Property TablesExists(ByVal ParamArray Tables() As String) As Boolean
        Get
            Try

                Dim cmd As New SqlCommand("select name from sysobjects where xtype='u' and name = '" & Join(Tables, "' and name = ") & "'", OpenDatabase("Menulator"))
                TablesExists = Not (cmd.ExecuteScalar) = ""
                cmd.Dispose()
            Catch
                Return False
            End Try
        End Get
    End Property
    Public Shared Property ExtractMameExe() As String
        Get
            If strMameExe Is Nothing Then
                _New()
            End If
            Return strMameExe
        End Get
        Set(ByVal value As String)
            If value = "" Then Throw New InvalidDataException
            If DoSQL("update mame set mamepath='" & value & "' where mamepath='" & strMameExe & "'") Then
                strMameExe = value
                'If App.Version(value) <> m_MameHeader.Build Then
                '    'unequal
                '    'Throw New MameInequalityException("WARNING: New Mame exe is a different build than the database!")
                'End If
            Else
                Throw New Data.DataException
            End If
        End Set
    End Property
    Public Shared ReadOnly Property ExtractMameBuild() As String
        Get
            Return strBuild
        End Get
    End Property
    Public Class RomStats
        Dim sRomName As String, sTime As TimeSpan
        Dim dLast As Nullable(Of DateTime)
        Dim iPLay As Integer
        Public Sub New(ByVal strRomName As String, ByVal playTime As TimeSpan, ByVal sLast As Nullable(Of DateTime), ByVal intCount As Integer)
            sRomName = strRomName
            sTime = playTime
            dLast = sLast
            iPLay = intCount
        End Sub
        Public ReadOnly Property RomName() As String
            Get
                Return sRomName
            End Get
        End Property
        Public ReadOnly Property PlayTime() As TimeSpan
            Get
                Return sTime
            End Get
        End Property
        Public ReadOnly Property LastPlayed() As Nullable(Of DateTime)
            Get
                Return dLast
            End Get
        End Property
        Public ReadOnly Property PlayCount() As Integer
            Get
                Return iPLay
            End Get
        End Property
    End Class

    Public Shared Function PlayRom(ByVal strMamePath As String, ByVal strRomName As String, ByVal strAdditionalParam As String, Optional ByRef ErrorString As String = "") As RomStats

        CloseDatabase()
        GC.Collect()
        GC.WaitForPendingFinalizers()
        'Dim timeNow As DateTime = Now
        'Redirect.SyncRedirect(strMamePath, strRomName, Nothing, Nothing, Nothing)
        Dim mymame As New MameProcess(strMamePath)
        If mymame.PlayRom(strRomName, strAdditionalParam) Then
            ErrorString = mymame.ErrorData
            Dim diff = mymame.ExitTime - mymame.StartTime
            'Dim timeEnd As TimeSpan = Now - timeNow
            Return UpdateStats(strRomName, Now, diff)
        End If
        ErrorString = mymame.ErrorData
        Return Nothing
        '        OpenDatabase(DatabaseName)
        '        DoSQL("UPDATE game SET playcount=coalesce(playcount,0)+1" & _
        '", playtime=coalesce(playtime,0)+" & CInt(timeEnd.TotalSeconds) & _
        '", playlast='" & Now.ToString & "'" & _
        '" WHERE name='" & strRomName & "'")
    End Function
    'Public Shared Function RomNameID(ByVal strRomName As String) As Integer
    '    Dim con = New SqlConnection("Integrated Security=TRUE;Initial Catalog=" & DatabaseName & ";Data Source=.\SQLEXPRESS;Connection Timeout=0;")
    '    con.Open()
    '    Dim cmd As New SqlCommand("select game_Id, name from game where name='" & strRomName & "'", con)
    '    RomNameID = cmd.ExecuteScalar
    '    cmd.Dispose()
    '    con.Dispose()
    'End Function
    ''' <summary>
    ''' Create or update user stats in the database.
    ''' </summary>
    ''' <param name="strRomName">Rom to update.</param>
    ''' <param name="LastPlayed">Date this rom was played last.</param>
    ''' <param name="PlayLength">Total time (in seconds) playing this rom.</param>
    ''' <param name="IncrementCount">Integer to add to the existing count. Use 0 to disable the play count.</param>
    Public Shared Function UpdateStats(ByVal strRomName As String, ByVal LastPlayed As DateTime, ByVal PlayLength As TimeSpan, Optional ByVal IncrementCount As Integer = 1) As RomStats
        OpenDatabase(DatabaseName)
        'Dim id As Integer = RomNameID(strRomName)
        DoSQL("IF EXISTS (select name from userstats where name=@1) " & _
"UPDATE userstats SET playcount=coalesce(playcount,0)+@2" & _
", playtime=coalesce(playtime,0)+@3" & _
", playlast=@4" & _
" from userstats where name=@1 ELSE " & _
"INSERT INTO userstats (name, playcount, playtime, playlast) VALUES (@1,@2,@3,@4)", strRomName, IncrementCount, CInt(PlayLength.TotalSeconds), LastPlayed.ToString)

        Dim t As DataTable = GetTable("select name, playcount, playtime from userstats where name='" & strRomName & "'")
        UpdateStats = New RomStats(strRomName, TimeSpan.FromSeconds(t.Rows(0)("playtime")), LastPlayed, t.Rows(0)("playcount"))
        t.Dispose()
    End Function
    Public Shared Function ClearStats(ByVal strRomName As String) As RomStats
        OpenDatabase(DatabaseName)
        DoSQL("UPDATE userstats SET playcount=0, playtime=0, playlast='' from userstats wHERE name='" & strRomName & "'")
        Return New RomStats(strRomName, TimeSpan.Zero, Nothing, 0)
    End Function
    Public Shared Sub ClearStats()
        OpenDatabase(DatabaseName)
        DoSQL("UPDATE userstats SET playcount=0, playtime=0, playlast=''")

    End Sub

#Region "SQL"
    Protected Friend Const DatabaseName As String = "Menulator"
    Shared sharedConnection As SqlConnection
    ''' <summary>
    ''' Returns if the Menulator namespace exists in the master table.
    ''' </summary>
    ''' <returns></returns>
    ''' <remarks></remarks>
    Public Shared Function NamespaceExists() As Boolean
        'Dim conn As New SqlConnection("Integrated Security=TRUE;USER INSTANCE=TRUE;Initial Catalog=;Data Source=.\SQLEXPRESS;")
        'conn.Open()
        Dim cmd As New SqlCommand("select db_id('" & DatabaseName & "')", OpenDatabase(""))
        NamespaceExists = Not cmd.ExecuteScalar() Is DBNull.Value
        '        DatabaseNamespaceExists = 
        cmd.Dispose()
        'conn.Close()
        'conn.Dispose()

    End Function
    ''' <summary>
    ''' Deletes the SQL namespace from the master table and deletes physical files.
    ''' </summary>
    ''' <returns></returns>
    ''' <remarks></remarks>
    Public Shared Function DeleteNamespace() As Boolean
        Dim cmd As SqlCommand = Nothing
        Try
            SqlConnection.ClearAllPools()
            cmd = New SqlCommand("DROP DATABASE " & DatabaseName, OpenDatabase(""))
            'cmd.CommandText = "DROP DATABASE " & strName  '[C:\GAMES\EMULATION\MAME\MAMEPP.MDF]"
            Dim i = cmd.ExecuteNonQuery
        Catch ex As Exception
            '        Throw ex
            Return False
        Finally
            cmd.Dispose()
        End Try
        Return True
    End Function

    'Public Shared Function AttachDatabase(ByVal strDBPath As String, ByVal bolPathIsToMameEXE As Boolean) As Boolean
    '    CloseDatabase(True)
    '    OpenDatabase("")
    '    If bolPathIsToMameEXE Then
    '        strDBPath = IO.Path.Combine(IO.Path.GetDirectoryName(strDBPath), IO.Path.GetFileNameWithoutExtension(strDBPath) & ".mdf")
    '    End If
    '    Dim strRO As String = strDBPath.Replace(".mdf", ".ndf") ' IO.Path.Combine(IO.Path.GetDirectoryName(strDBPath), IO.Path.GetFileNameWithoutExtension(strDBPath) & ".ndf")
    '    Dim strLog As String = strDBPath.Replace(".mdf", "_log.ldf")
    '    Try
    '        If Not NamespaceExists() Then
    '            'no namespace
    '            If IO.File.Exists(strDBPath) Then
    '                _DoSQL("CREATE DATABASE " & DatabaseName & _
    '                      " ON PRIMARY ( FILENAME = '" & strDBPath & "') " & _
    '                      " FOR ATTACH ")
    '            Else
    '                _DoSQL("CREATE DATABASE " & DatabaseName & " ON PRIMARY (Name='" & IO.Path.GetFileNameWithoutExtension(strDBPath) & "', " & _
    '                                          "filename = '" & strDBPath & "', size=2,maxsize=unlimited, filegrowth=1) ")
    '            End If

    '            _DoSQL("ALTER DATABASE " & DatabaseName & " ADD FileGroup ReadOnlyDB")
    '            If Not IO.File.Exists(strRO) Then
    '                'create a file 
    '                _DoSQL("ALTER DATABASE " & DatabaseName & " ADD FILE (name='ReadOnlyDB', filename='" & strRO & "', " & _
    '                      "size=3,maxsize=unlimited,filegrowth=1) TO FILEGROUP ReadOnlyDB")
    '            Else
    '                'attach a file
    '                _DoSQL("sp_attach_db @dbname='" & DatabaseName & "', @filename1='" & strRO & "'")
    '            End If
    '            If Not IO.File.Exists(strLog) Then
    '                'create a file
    '                _DoSQL("ALTER DATABASE " & DatabaseName & " ADD LOG FILE (filename='" & strLog & "')")
    '            Else
    '                'attach a file
    '                _DoSQL("sp_attach_db @dbname='" & DatabaseName & "', @filename1='" & strLog & "'")
    '            End If
    '        Else
    '            'namespace exists
    '            If Not IO.File.Exists(strDBPath) Then
    '                'no primary
    '                _DoSQL("ALTER DATABASE " & DatabaseName & " ADD FILE (name='" & IO.Path.GetFileNameWithoutExtension(strDBPath) & "', " & _
    '                        "filename='" & strDBPath & "',size=2,maxsize=unlimited,filegrowth=1)")
    '            Else
    '                'file exists assume correct
    '            End If
    '            If Not IO.File.Exists(strRO) Then
    '                'no readonly (this is the bulk of the mame db)
    '                _DoSQL("ALTER DATABASE " & DatabaseName & " ADD FILE (name='ReadOnlyDB', filename='" & strRO & "', " & _
    '      "size=3,maxsize=unlimited,filegrowth=1) TO FILEGROUP ReadOnlyDB")
    '            Else
    '                'file exists
    '            End If
    '            If Not IO.File.Exists(strLog) Then
    '                _DoSQL("ALTER DATABASE " & DatabaseName & " ADD LOG FILE (filename='" & strLog & "')")
    '            Else

    '            End If
    '        End If
    '    Catch ex As Exception
    '        Return False
    '    End Try
    '    Return True
    '    'attaches 
    'End Function
    'Public Shared Function DetachDatabase() As Boolean
    '    OpenDatabase("")
    '    Try
    '        _DoSQL("sp_detach_db @dbname='" & DatabaseName & "'")
    '    Catch
    '        Return False
    '    End Try
    '    Return True
    'End Function

    ''' <summary>
    ''' Creates a Menulator SQL namespace in the master table.
    ''' </summary>
    ''' <param name="strDBPath"></param>
    ''' <param name="bolPathIsToMameEXE">Set to True if strDbPath is pointing to a MAME executable instead of the Menulator database file.</param>
    ''' <remarks>This also creates the physical database if one doesnot already exist. The default save location is the path root of the mame executable.</remarks>
    Public Shared Function CreateNamespace(ByVal strDBPath As String, ByVal bolPathIsToMameEXE As Boolean) As Boolean
        'Dim e As New OperationChangedArgs(OperationChangedArgs.OperationTypes.CreateNamespace, "Create Namespace", False, 0, True, False, False)
        'DoOperationChanged(e)
        If bolPathIsToMameEXE Then
            strDBPath = IO.Path.Combine(IO.Path.GetDirectoryName(strDBPath), IO.Path.GetFileNameWithoutExtension(strDBPath) & ".mdf")
        End If
        Dim cmd As SqlCommand
        Try
            If Not IO.File.Exists(strDBPath) Then
                cmd = New SqlCommand("CREATE DATABASE " & DatabaseName & " ON PRIMARY (Name='" & IO.Path.GetFileNameWithoutExtension(strDBPath) & "', " & _
                                          "filename = '" & strDBPath & "', size=2,maxsize=unlimited, filegrowth=1), " & _
                                          "FILEGROUP ReadOnlyDB (Name='ReadOnlyDB', filename='" & IO.Path.Combine(IO.Path.GetDirectoryName(strDBPath), IO.Path.GetFileNameWithoutExtension(strDBPath) & ".ndf") & "', " & _
                                          "size=3,maxsize=unlimited, filegrowth=1)", OpenDatabase(""))
                Try
                    cmd.ExecuteNonQuery()
                Catch ex As SqlException When ex.Number = 1802
                    'file exists but no namespace
                    'Kill(strDBPath)
                    'Kill(IO.Path.ChangeExtension(strDBPath, "_log.ldf"))

                    cmd.CommandText = "sp_attach_db @dbname = '" & DatabaseName & "', " & _
                        "@filename1 = '" & strDBPath & "', " & _
                        "@filename2 = '" & strDBPath.Replace(".mdf", "_log.ldf") & "', " & _
                        "@filename3 = '" & strDBPath.Replace(".mdf", ".ndf") & "'"
                    Try
                        cmd.ExecuteNonQuery()
                    Catch ex2 As Exception
                        'e.Canceled = True
                        Throw ex2
                    End Try
                Catch ex As Exception
                    'e.Canceled = True
                    Throw ex
                    'Finally
                    '    cmd.CommandText = "ALTER DATABASE " & DatabaseName & " SET AUTO_SHRINK ON"
                    '    cmd.ExecuteNonQuery()
                    '    cmd.Dispose()
                    '    'e.IsComplete = True
                    '    'DoOperationChanged(e)
                End Try
            Else
                cmd = New SqlCommand("", OpenDatabase(""))
                Try

                    cmd.CommandText = "sp_attach_db @dbname = '" & DatabaseName & "', " & _
                                           "@filename1 = '" & strDBPath & "', " & _
                                           "@filename2 = '" & strDBPath.Replace(".mdf", "_log.ldf") & "', " & _
                                           "@filename3 = '" & strDBPath.Replace(".mdf", ".ndf") & "'"
                    Try
                        cmd.ExecuteNonQuery()
                    Catch ex2 As Exception
                        'e.Canceled = True
                        Throw ex2
                    End Try

                Catch ex As Exception
                    Throw ex
                Finally

                End Try
            End If
        Catch ex As Exception
            Throw ex
            Return False
        Finally

            cmd = New SqlCommand("ALTER DATABASE " & DatabaseName & " SET AUTO_SHRINK ON", OpenDatabase(""))
            cmd.ExecuteNonQuery()

            'cmd.CommandText = "DBCC ShrinkDatabase (" & DatabaseName & ")"
            'cmd.ExecuteNonQuery()

            'cmd.CommandText = "ALTER DATABASE " & DatabaseName & " MODIFY Filegroup ReadOnlyDB Read_Only"
            'cmd.ExecuteNonQuery()
            cmd.Dispose()
        End Try
        Return True
        'cmd.CommandText = 'log on" + "(name=Mamepp_log, filename='" & IO.Path.ChangeExtension(strPath, "_log.ldf") & "',size=3," + "maxsize=20,filegrowth=10%)"
    End Function
    Shared lastCatalog As String = "nothing"
    Public Shared Sub ShrinkDatabase()
        CloseDatabase()
        Dim con = New SqlConnection("Integrated Security=TRUE;Initial Catalog=" & "" & ";Data Source=.\SQLEXPRESS;Connection Timeout=0;")
        con.Open()
        Dim cmd As New SqlCommand("DBCC ShrinkDatabase (" & DatabaseName & ")", con)
        cmd.ExecuteNonQuery()
        cmd.Dispose()
        con.Dispose()
    End Sub
    Public Shared Sub MakeReadOnly()
        CloseDatabase(True)
        Try
            DisconnectConnections()
        Catch ex As Exception

        End Try
        Dim con = New SqlConnection("Integrated Security=TRUE;Initial Catalog=" & "" & ";Data Source=.\SQLEXPRESS;Connection Timeout=0;")
        con.Open()
        Dim cmd As New SqlCommand("ALTER DATABASE " & DatabaseName & " MODIFY Filegroup ReadOnlyDB Read_Only", con)
        cmd.ExecuteNonQuery()
        cmd.Dispose()
        con.Dispose()
    End Sub
    Public Shared Sub MakeReadWrite()
        CloseDatabase(True)
        Dim con = New SqlConnection("Integrated Security=TRUE;Initial Catalog=" & "" & ";Data Source=.\SQLEXPRESS;Connection Timeout=0;")
        con.Open()
        Dim cmd As New SqlCommand("ALTER DATABASE " & DatabaseName & " MODIFY Filegroup ReadOnlyDB READ_WRITE", con)

        cmd.ExecuteNonQuery()
        cmd.Dispose()
        con.Dispose()
    End Sub
    Public Shared Function TryMakeReadWrite() As Boolean
        Try
            MakeReadWrite()
        Catch ex As Exception
            Return False
        End Try
        Return True
    End Function
    Public Shared Function OpenNewDatabase(Optional ByVal Catalog As String = DatabaseName) As SqlConnection
        'If sharedConnection IsNot Nothing AndAlso sharedConnection.State = ConnectionState.Open Then
        '    If Catalog = lastCatalog Then
        '        Return sharedConnection
        '    Else
        '        sharedConnection.Close()
        '    End If
        'End If
        Dim con As SqlConnection
        Dim connstring As String = "Integrated Security=TRUE;Initial Catalog=" & Catalog & ";Data Source=.\SQLEXPRESS;Connection Timeout=0;"
        con = New SqlConnection(connstring)

        Try

            con.Open()
        Catch ex As Exception
            con.Dispose()
            'DoError(New ErrorEventArgs(ex))
        End Try
        Return con
    End Function
    Public Shared Function OpenDatabase(Optional ByVal Catalog As String = DatabaseName) As SqlConnection
        Debug.Print("<MAME.DATABASE.OPEN>")
        If sharedConnection IsNot Nothing AndAlso sharedConnection.State = ConnectionState.Open Then
            If Catalog = lastCatalog Then
                Return sharedConnection
            Else
                sharedConnection.Close()
            End If
        End If
        Dim connstring As String = "Integrated Security=TRUE;Initial Catalog=" & Catalog & ";Data Source=.\SQLEXPRESS;Connection Timeout=0;"
        sharedConnection = New SqlConnection(connstring)

        Try

            sharedConnection.Open()
            lastCatalog = Catalog
        Catch ex As Exception
            sharedConnection.Dispose()
            lastCatalog = "nothing"
            'DoError(New ErrorEventArgs(ex))
        End Try
        Return sharedConnection
    End Function
    Public Shared Sub CloseDatabase(Optional ByVal bolClearPools As Boolean = False)
        If sharedConnection IsNot Nothing Then
            'Debug.Print(sharedConnection.State.ToString)
            Try
                If sharedConnection.State <> ConnectionState.Closed Then
                    If bolClearPools Then SqlConnection.ClearAllPools()
                    sharedConnection.Close()
                End If
            Catch
            Finally
                sharedConnection.Dispose()
                Debug.Print("</MAME.DATABASE.CLOSE>")
            End Try
        End If
    End Sub

    ''' <summary>
    ''' Attempts to close all open connections to the Menulator database. Almost always fails.
    ''' </summary>
    ''' <returns></returns>
    ''' <remarks></remarks>
    Public Shared Function DisconnectConnections() As Boolean
        lastCatalog = "nothing"
        Dim cmd As New SqlCommand("ALTER DATABASE " & DatabaseName & " SET SINGLE_USER WITH ROLLBACK IMMEDIATE", OpenDatabase(""))
        Dim i As Object
        Try
            i = cmd.ExecuteNonQuery
            cmd.CommandText = "ALTER DATABASE " & DatabaseName & " SET MULTI_USER"
            cmd.ExecuteNonQuery()
        Catch ex As Exception
            Throw ex
            Return False
        Finally

            cmd.Dispose()
        End Try
        Return True
    End Function
    ''' <summary>
    ''' Executes a SQL statement against the database.
    ''' </summary>
    ''' <param name="strCMD">SQL command statement.</param>
    ''' <returns>Number of rows affected.</returns>
    ''' <remarks></remarks>
    Public Shared Function DoSQL(ByVal strCMD As String) As Integer
        Dim command As New SqlCommand(strCMD, OpenDatabase(DatabaseName))
        DoSQL = command.ExecuteNonQuery()
        command.Dispose()
    End Function
    Public Shared Function TrySQL(ByVal strCmd As String, ByRef Result As Integer) As Boolean
        Try
            Result = DoSQL(strCmd)
        Catch ex As Exception
            Return False
        End Try
        Return True
    End Function
    Public Shared Function DoSQL(ByVal strCMD As String, ByVal ParamArray parameters() As Object) As Integer
        Dim command As New SqlCommand(strCMD, OpenDatabase(DatabaseName))
        For t As Integer = 0 To UBound(parameters)
            command.Parameters.Add(New SqlParameter("@" & t + 1, parameters(t)))
        Next
        DoSQL = command.ExecuteNonQuery()
        command.Dispose()
    End Function
    Private Shared Function _DoSQL(ByVal strCmd As String) As Integer
        Dim command As New SqlCommand(strCmd, sharedConnection)
        _DoSQL = command.ExecuteNonQuery()
        command.Dispose()
    End Function
    ''' <summary>
    ''' Executes a SQL command against the database.
    ''' </summary>
    ''' <param name="strSQLCommand"></param>
    ''' <returns>First column of the first row value (as a string)</returns>
    ''' <remarks></remarks>
    Public Shared Function DoScalar(ByVal strSQLCommand As String) As String
        Dim command As New SqlCommand(strSQLCommand, OpenDatabase(DatabaseName))
        DoScalar = CStr(command.ExecuteScalar)
        command.Dispose()
    End Function
    ''' <summary>
    ''' Executes a SQL command against the database.
    ''' </summary>
    ''' <param name="strCommand"></param>
    ''' <returns>A new SqlReader from the supplied SQL command.</returns>
    ''' <remarks></remarks>
    Public Shared Function GetReader(ByVal strCommand As String) As SqlDataReader
        Dim cmd As New SqlCommand(strCommand, OpenDatabase(DatabaseName))
        GetReader = cmd.ExecuteReader()
        cmd.Dispose()

    End Function
    ''' <summary>
    ''' Executes a SQL command against the database.
    ''' </summary>
    ''' <param name="strCommand"></param>
    ''' <returns>A new DataTable from the supplied SQL command.</returns>
    ''' <remarks></remarks>
    Public Shared Function GetTable(ByVal strCommand As String) As DataTable
        Dim cmd As New SqlCommand(strCommand, OpenDatabase(DatabaseName))
        Dim d As New DataTable
        d.Load(cmd.ExecuteReader)
        GetTable = d
        cmd.Dispose()
    End Function

    Public Shared ReadOnly Property DatabaseConnection() As SqlConnection
        Get
            Return sharedConnection
        End Get
    End Property
#End Region

    Private Structure MameHeader

        Public sMamePath As String
        Public Build As String
        Public IsDebug As Boolean
        Public IsQuick As Boolean
        'Public CatVerFile As String
        Public Property MamePath() As String
            Get
                If sMamePath = "" Then Throw New InvalidDataException
                Return sMamePath
            End Get
            Set(ByVal value As String)
                sMamePath = value
                If value = "" Then Throw New InvalidDataException
            End Set
        End Property
    End Structure
    Public Sub New()
        m_MameHeader = New MameHeader
        m_MameHeader.Build = MAME.Database.ExtractMameBuild
        m_MameHeader.sMamePath = MAME.Database.ExtractMameExe
    End Sub
    Dim m_MameHeader As MameHeader
    ''' <summary>
    ''' Path to the MAME executable
    ''' </summary>
    ''' <value>Path to a new MAME executable</value>
    ''' <returns></returns>
    ''' <remarks>Setting this property also updates the recorded database mame path but does not update any potential changes.</remarks>
    Public Property MameExe() As String
        Get
            Return m_MameHeader.MamePath
        End Get
        Set(ByVal value As String)
            If value = "" Then Throw New InvalidDataException
            tryMakeReadWrite()
            Try
                If DoSQL("update mame set mamepath='" & value & "' where mamepath='" & m_MameHeader.MamePath & "'") Then
                    m_MameHeader.MamePath = value
                    'If App.Version(value) <> m_MameHeader.Build Then
                    '    'unequal
                    '    'Throw New MameInequalityException("WARNING: New Mame exe is a different build than the database!")
                    'End If
                Else
                    MsgBox("Can't")
                End If
            Finally
                MakeReadOnly()
            End Try
        End Set
    End Property
#Region "IDisposable Support"
    Private disposedValue As Boolean ' To detect redundant calls

    ' IDisposable
    Protected Overridable Sub Dispose(ByVal disposing As Boolean)
        If Not Me.disposedValue Then
            If disposing Then
                ' TODO: dispose managed state (managed objects).

            End If

            ' TODO: free unmanaged resources (unmanaged objects) and override Finalize() below.
            ' TODO: set large fields to null.
        End If
        Me.disposedValue = True
    End Sub

    ' TODO: override Finalize() only if Dispose(ByVal disposing As Boolean) above has code to free unmanaged resources.
    'Protected Overrides Sub Finalize()
    '    ' Do not change this code.  Put cleanup code in Dispose(ByVal disposing As Boolean) above.
    '    Dispose(False)
    '    MyBase.Finalize()
    'End Sub

    ' This code added by Visual Basic to correctly implement the disposable pattern.
    Public Sub Dispose() Implements IDisposable.Dispose
        ' Do not change this code.  Put cleanup code in Dispose(ByVal disposing As Boolean) above.
        Dispose(True)
        GC.SuppressFinalize(Me)
    End Sub
#End Region
    Public Shared Sub DownloadLatestCatVerFile(ByVal strMameExe As String)
        'http://www.progettoemma.net/public/catveren.zip
        'Dim s As String = LatestMameBuild(True)
        Dim objwebClient As New Net.WebClient
        objwebClient.Headers.Add("user-agent", "Mozilla/4.0 (compatible; MSIE 6.0;Windows NT 5.1")
        Dim catPath As String = IO.Path.Combine(IO.Path.GetDirectoryName(strMameExe), "cat32")
        If Not IO.Directory.Exists(catPath) Then
            IO.Directory.CreateDirectory(catPath)

        End If
        Dim s As String = IO.Path.Combine(catPath, "catveren.zip")
        'Dim strTarget As String = Path.Combine(strTargetDirectory, Path.GetFileName(s))
        objwebClient.DownloadFile("http://www.progettoemma.net/public/catveren.zip", s)
        objwebClient.Dispose()
        Unzip(s, IO.Path.GetDirectoryName(s), "catveren\Catver.ini", True)

        'Return IO.Path.Combine(IO.Path.GetDirectoryName(s), "Catver.ini")
    End Sub

    Public Shared Sub DownloadLatestCat32File(ByVal strMameExe As String)
        'http://www.progettoemma.net/public/catveren.zip
        'Dim s As String = LatestMameBuild(True)
        Dim objwebClient As New Net.WebClient
        objwebClient.Headers.Add("user-agent", "Mozilla/4.0 (compatible; MSIE 6.0;Windows NT 5.1")
        Dim catPath As String = IO.Path.Combine(IO.Path.GetDirectoryName(strMameExe), "cat32")
        If Not IO.Directory.Exists(catPath) Then
            IO.Directory.CreateDirectory(catPath)

        End If
        Dim s As String = IO.Path.Combine(catPath, "cat32en.zip")
        'Dim strTarget As String = Path.Combine(strTargetDirectory, Path.GetFileName(s))
        objwebClient.DownloadFile("http://www.progettoemma.net/public/cat32en.zip", s)
        objwebClient.Dispose()
        Unzip(s, IO.Path.GetDirectoryName(s), True)
    End Sub




    Shared WithEvents MyThread As New ManagedThread
    Public Class OperationCompleteArgs
        Inherits EventArgs
        ReadOnly _Operation As String
        ReadOnly _Success As Boolean
        Public Sub New(ByVal sOperation As String, ByVal bSuccess As Boolean)
            Me._Operation = sOperation
            Me._Success = bSuccess
        End Sub
        Public ReadOnly Property Operation() As String
            Get
                Return _Operation
            End Get
        End Property
        Public ReadOnly Property Success() As Boolean
            Get
                Return _Success
            End Get
        End Property
    End Class
    Public Class OperationProgressArgs
        Inherits EventArgs
        ReadOnly _Operation As String
        ReadOnly _Status As String
        ReadOnly _Percent As Integer
        Public Sub New(ByVal sOperation As String, ByVal sStatus As String, ByVal iPercent As Integer)
            Me._Operation = sOperation
            Me._Status = sStatus
            Me._Percent = iPercent
        End Sub
        Public ReadOnly Property Operation() As String
            Get
                Return _Operation
            End Get
        End Property
        Public ReadOnly Property Status() As String
            Get
                Return _Status
            End Get
        End Property
        Public ReadOnly Property Percent() As Integer
            Get
                Return _Percent
            End Get
        End Property
    End Class
    'Public Delegate Sub OperationCompleteDelegate(ByVal operation As String, ByVal bSuccess As Boolean)
    'Public Delegate Sub OperationProgressUpdateDelegate(ByVal operation As String, ByVal strStatus As String, ByVal iPercent As Integer)
    Public Delegate Sub OperationCompleteDelegate(ByVal e As OperationCompleteArgs)
    Public Delegate Sub OperationProgressUpdateDelegate(ByVal e As OperationProgressArgs)

    Private MustInherit Class OperationBase
        Public MameExe As String
        Public ProgressDelegate As OperationProgressUpdateDelegate
        Public CompleteDelegate As OperationCompleteDelegate
        Public Ping As Threading.AutoResetEvent
        Public IsAsync As Boolean
        Public MustOverride ReadOnly Property Name() As String
        Public Sub TryInvokeProgress(ByVal e As OperationProgressArgs, Optional ByVal bAsync As Boolean = False)
            If ProgressDelegate Is Nothing Then Return
            If bAsync = False Then
                ProgressDelegate.Invoke(e)
            Else
                ProgressDelegate.BeginInvoke(e, AddressOf genericCallback, Nothing)
            End If
        End Sub
        Public Sub TryInvokeComplete(ByVal e As OperationCompleteArgs, Optional ByVal bAsync As Boolean = False)
            If CompleteDelegate Is Nothing Then Return
            If bAsync = False Then
                CompleteDelegate.Invoke(e)
            Else
                CompleteDelegate.BeginInvoke(e, AddressOf genericCallback, Nothing)
            End If
        End Sub
        'Private Sub ProgressCallBack(ByVal ar As IAsyncResult)
        '    With DirectCast(ar.AsyncState, Runtime.Remoting.Messaging.AsyncResult)
        '        If Not .EndInvokeCalled Then
        '            DirectCast(.AsyncDelegate, OperationProgressUpdateDelegate).EndInvoke(ar)
        '        End If
        '    End With
        'End Sub
        'Private Sub CompleteCallback(ByVal ar As IAsyncResult)

        'End Sub
        Public Sub New(ByVal sMameExe As String, ByVal complete As OperationCompleteDelegate, ByVal Progress As OperationProgressUpdateDelegate, ByVal ping As Threading.AutoResetEvent, ByVal bAsync As Boolean)
            MameExe = sMameExe
            ProgressDelegate = Progress
            CompleteDelegate = complete
            Me.Ping = ping
            IsAsync = bAsync
        End Sub
        Private Sub genericCallback(ByVal ar As IAsyncResult)
            With DirectCast(ar.AsyncState, Runtime.Remoting.Messaging.AsyncResult)
                If Not .EndInvokeCalled Then
                    .AsyncDelegate.EndInvoke(ar)
                End If
            End With
        End Sub
        Public Function Wait() As Boolean
            If Ping IsNot Nothing Then
                Return Ping.WaitOne()
            End If
            Return True
        End Function
    End Class
    Private Class MakeTablesOperation
        Inherits OperationBase

        Public Sub New(ByVal sMameExe As String, ByVal complete As OperationCompleteDelegate, ByVal Progress As OperationProgressUpdateDelegate, ByVal ping As Threading.AutoResetEvent, ByVal bAsync As Boolean)
            MyBase.New(sMameExe, complete, Progress, ping, bAsync)
        End Sub
        Public Overrides ReadOnly Property Name() As String
            Get
                Return "Creating Tables"
            End Get
        End Property
    End Class
    Private Class VerifyRomsOperation
        Inherits OperationBase
        Public ForceProgress As Boolean
        Public Sub New(ByVal sMameExe As String, ByVal complete As OperationCompleteDelegate, ByVal Progress As OperationProgressUpdateDelegate, ByVal ping As Threading.AutoResetEvent, ByVal bAsync As Boolean, Optional ByVal bForceProgress As Boolean = True)
            MyBase.New(sMameExe, complete, Progress, ping, bAsync)
            Me.ForceProgress = bForceProgress
        End Sub

        Public Overrides ReadOnly Property Name() As String
            Get
                Return "Verifying Roms"
            End Get
        End Property
    End Class
    Private Class CategorizeOperation
        Inherits OperationBase
        Public Sub New(ByVal sMameExe As String, ByVal complete As OperationCompleteDelegate, ByVal Progress As OperationProgressUpdateDelegate, ByVal ping As Threading.AutoResetEvent, ByVal bAsync As Boolean)
            MyBase.New(sMameExe, complete, Progress, ping, bAsync)
        End Sub

        Public Overrides ReadOnly Property Name() As String
            Get
                Return "Categorizing Roms"
            End Get
        End Property
    End Class
    Private Class RebuildOperation
        Inherits OperationBase
        Public Sub New(ByVal sMameExe As String, ByVal complete As OperationCompleteDelegate, ByVal Progress As OperationProgressUpdateDelegate, ByVal ping As Threading.AutoResetEvent, ByVal bAsync As Boolean)
            MyBase.New(sMameExe, complete, Progress, ping, bAsync)
        End Sub

        Public Overrides ReadOnly Property Name() As String
            Get
                Return "Rebuild Database"
            End Get
        End Property
    End Class


    Private Shared Function _Rebuild(ByVal o As Object) As Boolean
        'Dim t = Now
        'Try
        Dim m As RebuildOperation = o
        'If UBound(o) >= 3 AndAlso o(3) IsNot Nothing Then
        '    DirectCast(o(3), Threading.AutoResetEvent).WaitOne()
        'End If
        m.Wait()
        Dim bError As Boolean = False
        bError = MyThread.RequestedStop
        If Not bError Then bError = Not _CreateTablesFromMame2(New MakeTablesOperation(m.MameExe, Nothing, m.ProgressDelegate, Nothing, m.IsAsync)) 'bError = Not _CreateTablesFromMame2(New Object() {o(0), Nothing, o(2)})

        If Not bError Then bError = MyThread.RequestedStop
        If Not bError Then bError = Not _VerifyRoms(New VerifyRomsOperation(m.MameExe, Nothing, m.ProgressDelegate, Nothing, m.IsAsync)) 'Not _VerifyRoms(New Object() {o(0), Nothing, o(2)})

        If Not bError Then bError = MyThread.RequestedStop
        If Not bError Then bError = Not _Categorize(New CategorizeOperation(m.MameExe, Nothing, m.ProgressDelegate, Nothing, m.IsAsync)) '_Categorize(New Object() {o(0), Nothing, o(2)})
        'If UBound(o) > 0 Then
        'If o(1) IsNot Nothing Then DirectCast(o(1), OperationCompleteDelegate).Invoke(New OperationCompleteArgs("Rebuild", Not bError))
        'End If
        m.TryInvokeComplete(New OperationCompleteArgs("Rebuild Database", Not bError))
        Return Not bError
        'If o(2) Then
        '    DoOperationCompleted(New OperationCompleteArgs(OperationTypes.FullRebuild, True, Now - t, "", -1))
        'End If
        'Catch ex As Threading.ThreadAbortException
        '    'If o(2) Then
        '    '    DoOperationCompleted(New OperationCompleteArgs(OperationTypes.FullRebuild, False, Now - t, "", -1))
        '    'End If
        '    Return
        'End Try
    End Function
    ''' <summary>
    ''' Creates tables, categorizes, and verifies all in one routine
    ''' </summary>
    ''' <param name="strMamePath"></param>
    ''' <param name="bolAsync"></param>
    ''' <remarks></remarks>
    Public Shared Sub RebuildDatabase(ByVal strMamePath As String, Optional ByVal bolAsync As Boolean = True)
        'CanCancel = True
        If bolAsync Then
            If MyThread.State <> ManagedThread.States.Stopped Then

                MyThread.Stop()
            End If
            'MyThread.Start(AddressOf _Rebuild, New Object() {strMamePath, Nothing, Nothing})
            MyThread.Start(AddressOf _Rebuild, New RebuildOperation(strMamePath, Nothing, Nothing, Nothing, True))
        Else
            '_Rebuild(New Object() {strMamePath, Nothing, Nothing})
            _Rebuild(New RebuildOperation(strMamePath, Nothing, Nothing, Nothing, False))
        End If
    End Sub
    Public Shared Sub RebuildDatabaseWithCallback(ByVal strMamePath As String, ByVal OnFinish As OperationCompleteDelegate, ByVal OnProgress As OperationProgressUpdateDelegate, ByVal ping As Threading.AutoResetEvent)
        'CanCancel = True
        If MyThread.State <> ManagedThread.States.Stopped Then

            MyThread.Stop()
        End If
        'If bolAsync Then
        'MyThread.Start(AddressOf _Rebuild, New Object() {strMamePath, OnFinish, OnProgress, ping})
        MyThread.Start(AddressOf _Rebuild, New RebuildOperation(strMamePath, OnFinish, OnProgress, ping, True))
        'Else
        '    _Rebuild(New Object() {strMamePath})
        'End If
    End Sub
    ''' <summary>
    ''' Begins the Verify Roms operation
    ''' </summary>
    ''' <param name="strMamePath"></param>
    ''' <param name="bolAsync"></param>
    ''' <remarks></remarks>
    Public Shared Sub VerifyRoms(ByVal strMamePath As String, Optional ByVal bolAsync As Boolean = True)
        'CanCancel = True
        If bolAsync Then
            If MyThread.State <> ManagedThread.States.Stopped Then

                MyThread.Stop()
            End If
            'MyThread.Start(AddressOf _VerifyRoms, New Object() {strMamePath})
            MyThread.Start(AddressOf _VerifyRoms, New VerifyRomsOperation(strMamePath, Nothing, Nothing, Nothing, True))
        Else
            '_VerifyRoms(New Object() {strMamePath})
            _VerifyRoms(New VerifyRomsOperation(strMamePath, Nothing, Nothing, Nothing, False))
        End If
    End Sub
    Public Shared Sub VerifyRomsWithCallback(ByVal strMamePath As String, ByVal OnFinish As OperationCompleteDelegate, ByVal OnProgress As OperationProgressUpdateDelegate, ByVal ping As Threading.AutoResetEvent)
        If MyThread.State <> ManagedThread.States.Stopped Then

            MyThread.Stop()
        End If
        'MyThread.Start(AddressOf _VerifyRoms, New Object() {strMamePath, OnFinish, OnProgress, ping})
        MyThread.Start(AddressOf _VerifyRoms, New VerifyRomsOperation(strMamePath, OnFinish, OnProgress, ping, True))
    End Sub
    Public Shared Function Categorize(ByVal strMamePath As String, Optional ByVal bolAsync As Boolean = True) As Boolean
        'DoOperationChanged(New OperationChangedArgs(OperationChangedArgs.OperationTypes.CategorizeDatabase, "Categorize Database", True, 0, True, False, False))
        If bolAsync Then
            If MyThread.State <> ManagedThread.States.Stopped Then

                MyThread.Stop()
            End If
            '            MyThread.Start(AddressOf _Categorize, New Object() {strMamePath})
            MyThread.Start(AddressOf _Categorize, New CategorizeOperation(strMamePath, Nothing, Nothing, Nothing, True))
            Return True
        Else
            'Return _Categorize(New Object() {strMamePath})
            Return _Categorize(New CategorizeOperation(strMamePath, Nothing, Nothing, Nothing, False))
        End If
    End Function
    Public Shared Sub CategorizeWithCallback(ByVal strMamePath As String, ByVal OnFinish As OperationCompleteDelegate, ByVal OnProgress As OperationProgressUpdateDelegate, ByVal ping As Threading.AutoResetEvent)
        If MyThread.State <> ManagedThread.States.Stopped Then

            MyThread.Stop()
        End If
        'MyThread.Start(AddressOf _Categorize, New Object() {strMamePath, OnFinish, OnProgress, ping})
        MyThread.Start(AddressOf _Categorize, New CategorizeOperation(strMamePath, OnFinish, OnProgress, ping, True))
    End Sub
    Public Shared Sub CreateTablesFromMame(ByVal strMamePath As String, Optional ByVal bolAsync As Boolean = True)
        'DoOperationChanged(New OperationChangedArgs(OperationChangedArgs.OperationTypes.CategorizeDatabase, "Categorize Database", True, 0, True, False, False))
        If bolAsync Then
            If MyThread.State <> ManagedThread.States.Stopped Then

                MyThread.Stop()
            End If
            'MyThread.Start(AddressOf _CreateTablesFromMame2, New Object() {strMamePath, Nothing, Nothing})
            MyThread.Start(AddressOf _CreateTablesFromMame2, New MakeTablesOperation(strMamePath, Nothing, Nothing, Nothing, True))
        Else
            '_CreateTablesFromMame2(New Object() {strMamePath, Nothing, Nothing})
            _CreateTablesFromMame2(New MakeTablesOperation(strMamePath, Nothing, Nothing, Nothing, False))
        End If
    End Sub
    Public Shared Sub CreateTablesFromMameCallback(ByVal strMamePath As String, ByVal OnFinish As OperationCompleteDelegate, ByVal OnProgress As OperationProgressUpdateDelegate, ByVal ping As Threading.AutoResetEvent)
        If MyThread.State <> ManagedThread.States.Stopped Then

            MyThread.Stop()
        End If
        'MyThread.Start(AddressOf _CreateTablesFromMame2, New Object() {strMamePath, OnFinish, OnProgress, ping})
        MyThread.Start(AddressOf _CreateTablesFromMame2, New MakeTablesOperation(strMamePath, OnFinish, OnProgress, ping, True))
    End Sub
    'Private Shared Function _CreateTablesFromMame(ByVal oe As Object) 'ByVal strMame As String, Optional ByVal bolQuick As Boolean = False)
    '    Dim strMame As String = oe(0)
    '    'Dim bolQuick As Boolean = oe(1)
    '    'Dim SuppressFinalEvent As Boolean = oe(2)
    '    Dim _FinishDelegate As OperationCompleteDelegate = Nothing
    '    Dim _ProgressDelegate As OperationProgressUpdateDelegate = Nothing
    '    If UBound(oe) > 0 Then
    '        If oe(1) IsNot Nothing Then _FinishDelegate = oe(1)
    '        If oe(2) IsNot Nothing Then _ProgressDelegate = oe(2)
    '    End If

    '    Debug.Print("Start ")
    '    GC.Collect()
    '    GC.WaitForPendingFinalizers()

    '    'startTime = Now
    '    'Dim e As New OperationStartedArgs(OperationTypes.CreateDatabase, "Create Tables", Windows.Forms.ProgressBarStyle.Marquee)
    '    'Dim e2 As New OperationStatsArgs(OperationTypes.CreateDatabase)
    '    If _ProgressDelegate IsNot Nothing Then _ProgressDelegate.Invoke("CreateTables", "Preparing data...", -1)
    '    Dim s1 As String

    '    'DoOperationStarted(e)
    '    'DoCancelOperationValueChanged(False)
    '    'e2.Begin()
    '    'e2.Description = "Building Database Tables"
    '    'DoUpdateOperationStats(e2)
    '    If TablesExists Then
    '        'Dim smame As String = ExtractMameExe
    '        CloseDatabase()

    '        If Not DeleteNamespace() Then
    '            'DoOperationCompleted(New OperationCompleteArgs(e.Operation, False, Now - startTime, "Can't delete namespace", -1))
    '            If _FinishDelegate IsNot Nothing Then _FinishDelegate.Invoke("CreateTables", False)
    '            Return False
    '        End If
    '        'If smame = "" Then
    '        'Else
    '        'Kill(IO.Path.Combine(IO.Path.GetDirectoryName(smame), IO.Path.GetFileNameWithoutExtension(smame) & ".mdf"))
    '        'End If

    '    End If
    '    CloseDatabase()
    '    If IO.File.Exists(IO.Path.Combine(IO.Path.GetDirectoryName(strMame), IO.Path.GetFileNameWithoutExtension(strMame) & ".mdf")) Then Kill(IO.Path.Combine(IO.Path.GetDirectoryName(strMame), IO.Path.GetFileNameWithoutExtension(strMame) & ".mdf"))

    '    If NamespaceExists() = False Then
    '        CreateNamespace(strMame, True)
    '    End If
    '    'Dim lines As New List(Of String)
    '    Dim xreader = RedirectStream(strMame, "-lx")
    '    Dim defaults As New Dictionary(Of String, Dictionary(Of String, String))
    '    Dim notnulls As New Dictionary(Of String, List(Of String))
    '    Dim ranges As New Dictionary(Of String, Dictionary(Of String, String))
    '    Dim fields As New Dictionary(Of String, List(Of String))
    '    'Dim quickList() As String = {"mame", "game"}
    '    With xreader
    '        While .Peek > 0
    '            s1 = .ReadLine.Trim
    '            'lines.Add(s1)
    '            'sb.AppendLine(s)
    '            If s1.StartsWith("<!ATTLIST") Then
    '                Dim u() As String = Split(s1.Replace("<", "").Replace(">", ""), " ")
    '                For t As Integer = UBound(u) To 2 Step -1
    '                    If u(t) = "#REQUIRED" Then
    '                        If Not notnulls.ContainsKey(u(1)) Then notnulls.Add(u(1), New List(Of String))
    '                        notnulls(u(1)).Add(u(2))
    '                    End If
    '                    If u(t).EndsWith(Chr(34)) Then
    '                        If Not defaults.ContainsKey(u(1)) Then defaults.Add(u(1), New Dictionary(Of String, String))
    '                        defaults(u(1)).Add(u(2), u(t).Replace(Chr(34), "'"))
    '                    End If
    '                    If u(t).StartsWith("(") Then
    '                        If Not ranges.ContainsKey(u(1)) Then ranges.Add(u(1), New Dictionary(Of String, String))
    '                        ranges(u(1)).Add(u(2), u(t).Replace("(", "").Replace(")", ""))
    '                    End If
    '                Next
    '            End If
    '            If s1 = "]>" Then Exit While
    '        End While
    '        .Close()

    '    End With
    '    Dim d As New DataSet("Mame")
    '    'Return False
    '    xreader.Dispose()
    '    xreader = RedirectStream(strMame, "-lx")

    '    d.ReadXml(xreader, XmlReadMode.InferTypedSchema)
    '    xreader.Dispose()
    '    'infered gives good valuetypes
    '    'infer sucks
    '    'ignore does nothing
    '    'fragment breaks
    '    'diff breask   

    '    'e.IsComplete = True
    '    'DoOperationChanged(e)
    '    'e.Operation = OperationChangedArgs.OperationTypes.PopulateDatabase
    '    'e.Description = "Populating Database"
    '    'e.HasPercent = True
    '    'e.CanCancel = True
    '    'e.IsComplete = False
    '    'DoOperationChanged(e)

    '    'DoOperationStarted(New OperationStartedArgs(OperationTypes.PopulateDatabase, "Create Tables", Windows.Forms.ProgressBarStyle.Continuous))
    '    'DoCancelOperationValueChanged(True)

    '    'rowCount = 0
    '    'rowPos = 0

    '    Dim s As String = "", s2() As String = Nothing
    '    Dim adapter As SqlDataAdapter = Nothing
    '    Dim cmd As SqlCommand = Nothing
    '    'Dim e2 As New OperationStatsArgs(OperationTypes.CreateDatabase)
    '    'e2.Begin()

    '    Dim bError As Boolean
    '    Try

    '        adapter = New SqlDataAdapter("", OpenDatabase(DatabaseName))
    '        'AddHandler adapter.RowUpdated, AddressOf adapter_RowUpdated
    '        'adapter.UpdateBatchSize = 20
    '        cmd = New SqlCommand("", sharedConnection)
    '        For Each t As DataTable In d.Tables
    '            'If bolQuick Then
    '            '    If Not quickList.Contains(t.TableName) Then Continue For
    '            'End If
    '            'rowCount += t.Rows.Count

    '            s = "CREATE TABLE " & t.TableName & " ("
    '            'For Each g As DataRelation In t.ParentRelations
    '            '    'Debug.Print(g.ChildKeyConstraint.)
    '            'Next
    '            fields.Add(t.TableName, New List(Of String))

    '            Erase s2

    '            For Each c As DataColumn In t.Columns
    '                If c.ColumnName = "mame_Id" Then Continue For
    '                If notnulls.ContainsKey(t.TableName) AndAlso notnulls(t.TableName).Contains(c.ColumnName) Then
    '                    c.AllowDBNull = False
    '                End If
    '                If defaults.ContainsKey(t.TableName) AndAlso defaults(t.TableName).ContainsKey(c.ColumnName) Then
    '                    c.DefaultValue = defaults(t.TableName)(c.ColumnName)
    '                End If

    '                If s2 Is Nothing Then ReDim s2(0) Else ReDim Preserve s2(UBound(s2) + 1)

    '                Dim hint As Object = Nothing
    '                If ranges.ContainsKey(t.TableName) AndAlso ranges(t.TableName).ContainsKey(c.ColumnName) Then
    '                    hint = ranges(t.TableName)(c.ColumnName)
    '                End If

    '                s2(UBound(s2)) = "[" & c.ColumnName & "] " & GetDBType(c.DataType, False, hint) & IIf(c.MaxLength > -1, "(" & c.MaxLength & ")", "")
    '                fields(t.TableName).Add("[" & c.ColumnName & "]")
    '                If c.DefaultValue IsNot System.DBNull.Value Then s2(UBound(s2)) &= " DEFAULT " & c.DefaultValue.ToString
    '                If t.PrimaryKey.Contains(c) Then
    '                    s2(UBound(s2)) &= " PRIMARY KEY"
    '                ElseIf c.Unique Then
    '                    s2(UBound(s2)) &= " UNIQUE"
    '                ElseIf c.AllowDBNull = False Then
    '                    s2(UBound(s2)) &= " NOT NULL"
    '                Else
    '                    If t.ParentRelations.Count Then
    '                        For Each da As DataRelation In t.ParentRelations
    '                            For Each ta As DataColumn In da.ParentColumns
    '                                If ta.ColumnName = c.ColumnName Then
    '                                    'With DirectCast(da.ParentColumns(Array.IndexOf(da.ParentColumns, c)), DataColumn)
    '                                    s2(UBound(s2)) &= " foreign key references " & da.ParentTable.TableName & "("
    '                                    Dim q() As String = Nothing
    '                                    For Each cc As DataColumn In da.ParentColumns
    '                                        If q Is Nothing Then ReDim q(0) Else ReDim Preserve q(UBound(q) + 1)
    '                                        q(UBound(q)) = cc.ColumnName
    '                                    Next cc
    '                                    s2(UBound(s2)) &= Join(q, ", ") & ") on delete cascade"
    '                                    'Debug.Print(da.ParentKeyConstraint.ExtendedProperties.ToString)
    '                                    'End With
    '                                End If
    '                            Next ta
    '                        Next da
    '                    End If
    '                End If
    '            Next c
    '            Select Case t.TableName
    '                Case "game"
    '                    Dim i As Integer = UBound(s2)
    '                    ReDim Preserve s2(UBound(s2) + 7)
    '                    s2(i + 1) = "mature bit default 0"
    '                    s2(i + 2) = "verified bit default 0"
    '                    s2(i + 3) = "category varchar(255)"
    '                    s2(i + 4) = "playtime int default 0"
    '                    s2(i + 5) = "playcount int default 0"
    '                    s2(i + 6) = "playlast varchar(255) default ''"
    '                    s2(i + 7) = "favorite bit default 0"
    '                Case "mame"
    '                    Dim i As Integer = UBound(s2)
    '                    ReDim Preserve s2(UBound(s2) + 1)
    '                    s2(i + 1) = "mamepath varchar(255)"
    '                    's2(i + 2) = "quick bit"
    '                    t.Columns.Add("mamepath", GetType(System.String))
    '                    t.Rows(0)![mamepath] = strMame
    '                    fields("mame").Add("mamepath")
    '                    't.Columns.Add("quick", GetType(Boolean))
    '                    't.Rows(0)![quick] = bolQuick
    '                    'fields("mame").Add("quick")
    '            End Select
    '            s &= Join(s2, ", ") & ")"

    '            Debug.Print(s)
    '            cmd.CommandText = s
    '            Try
    '                cmd.ExecuteNonQuery()
    '            Catch ex As Exception
    '                Debug.Print("Error: " & ex.Message)
    '            End Try
    '        Next t

    '        For Each t As DataTable In d.Tables
    '            'If bolQuick Then
    '            '    If Not quickList.Contains(t.TableName) Then Continue For
    '            'End If
    '            'e2.Description = "Populating Table '" & t.TableName & "'"
    '            'e2.Count += 1
    '            'DoUpdateOperationStats(e2)
    '            adapter.SelectCommand.CommandText = "select " & Join(fields(t.TableName).ToArray, ", ") & " from " & t.TableName
    '            Dim o As New SqlCommandBuilder(adapter)
    '            adapter.InsertCommand = o.GetInsertCommand

    '            For Each p As SqlParameter In adapter.InsertCommand.Parameters
    '                If defaults.ContainsKey(t.TableName) AndAlso defaults(t.TableName).ContainsKey(p.SourceColumn) Then
    '                    p.Value = "DEFAULT" 'defaults(t.TableName)(p.SourceColumn)
    '                End If
    '            Next p

    '            adapter.Update(d, t.TableName)

    '        Next t
    '        'e2.Description = "Create Database Complete!"
    '        'e2.Stop()
    '        'DoUpdateOperationStats(e2)
    '    Catch ex As Exception
    '        'Throw ex
    '        'e.Canceled = True
    '        'DoOperationCompleted(New OperationCompleteArgs(OperationTypes.PopulateDatabase, False, Now - startTime, ex.Message, -1))
    '        bError = True
    '    Finally

    '        If adapter IsNot Nothing Then
    '            'RemoveHandler adapter.RowUpdated, AddressOf adapter_RowUpdated
    '            adapter.Dispose()

    '        End If
    '        If cmd IsNot Nothing Then cmd.Dispose()

    '        d.Dispose()
    '        'CloseDatabase()
    '        'If SuppressFinalEvent = False Then DoOperationCompleted(New OperationCompleteArgs(OperationTypes.PopulateDatabase, True, Now - startTime, "Created Tables", rowCount))
    '        Debug.Print("DONE")
    '        'e.IsComplete = True
    '        'DoOperationChanged(e)
    '        GC.Collect()
    '        GC.WaitForPendingFinalizers()
    '        If _FinishDelegate IsNot Nothing Then
    '            _FinishDelegate.Invoke("CreateTables", Not bError)
    '        End If
    '    End Try
    '    Return Not bError
    'End Function
    'Private Shared Function _UpdateTablesFromMame(ByVal oe As Object) 'ByVal strMame As String, Optional ByVal bolQuick As Boolean = False)
    '    Dim strMame As String = oe(0)
    '    Dim _FinishDelegate As OperationCompleteDelegate = Nothing
    '    Dim _ProgressDelegate As OperationProgressUpdateDelegate = Nothing
    '    If UBound(oe) > 0 Then
    '        If oe(1) IsNot Nothing Then _FinishDelegate = oe(1)
    '        If oe(2) IsNot Nothing Then _ProgressDelegate = oe(2)
    '    End If

    '    Debug.Print("Start ")
    '    GC.Collect()
    '    GC.WaitForPendingFinalizers()

    '    If _ProgressDelegate IsNot Nothing Then _ProgressDelegate.Invoke("CreateTables", "Preparing data...", -1)
    '    Dim s1 As String


    '    'If TablesExists Then
    '    '    CloseDatabase()

    '    '    If Not DeleteNamespace() Then
    '    '        If _FinishDelegate IsNot Nothing Then _FinishDelegate.Invoke("CreateTables", False)
    '    '        Return False
    '    '    End If
    '    'End If
    '    'CloseDatabase()
    '    'If IO.File.Exists(IO.Path.Combine(IO.Path.GetDirectoryName(strMame), IO.Path.GetFileNameWithoutExtension(strMame) & ".mdf")) Then Kill(IO.Path.Combine(IO.Path.GetDirectoryName(strMame), IO.Path.GetFileNameWithoutExtension(strMame) & ".mdf"))

    '    'If NamespaceExists() = False Then
    '    '    CreateNamespace(strMame, True)
    '    'End If

    '    Dim xreader = RedirectStream(strMame, "-lx")
    '    Dim defaults As New Dictionary(Of String, Dictionary(Of String, String))
    '    Dim notnulls As New Dictionary(Of String, List(Of String))
    '    Dim ranges As New Dictionary(Of String, Dictionary(Of String, String))
    '    Dim fields As New Dictionary(Of String, List(Of String))
    '    With xreader
    '        While .Peek > 0
    '            s1 = .ReadLine.Trim
    '            If s1.StartsWith("<!ATTLIST") Then
    '                Dim u() As String = Split(s1.Replace("<", "").Replace(">", ""), " ")
    '                For t As Integer = UBound(u) To 2 Step -1
    '                    If u(t) = "#REQUIRED" Then
    '                        If Not notnulls.ContainsKey(u(1)) Then notnulls.Add(u(1), New List(Of String))
    '                        notnulls(u(1)).Add(u(2))
    '                    End If
    '                    If u(t).EndsWith(Chr(34)) Then
    '                        If Not defaults.ContainsKey(u(1)) Then defaults.Add(u(1), New Dictionary(Of String, String))
    '                        defaults(u(1)).Add(u(2), u(t).Replace(Chr(34), "'"))
    '                    End If
    '                    If u(t).StartsWith("(") Then
    '                        If Not ranges.ContainsKey(u(1)) Then ranges.Add(u(1), New Dictionary(Of String, String))
    '                        ranges(u(1)).Add(u(2), u(t).Replace("(", "").Replace(")", ""))
    '                    End If
    '                Next
    '            End If
    '            If s1 = "]>" Then Exit While
    '        End While
    '        .Close()

    '    End With
    '    Dim d As New DataSet("Mame")
    '    xreader.Dispose()
    '    xreader = RedirectStream(strMame, "-lx")

    '    d.ReadXml(xreader, XmlReadMode.InferTypedSchema)
    '    xreader.Dispose()
    '    'infered gives good valuetypes
    '    'infer sucks
    '    'ignore does nothing
    '    'fragment breaks
    '    'diff breask   

    '    'e.IsComplete = True
    '    'DoOperationChanged(e)
    '    'e.Operation = OperationChangedArgs.OperationTypes.PopulateDatabase
    '    'e.Description = "Populating Database"
    '    'e.HasPercent = True
    '    'e.CanCancel = True
    '    'e.IsComplete = False
    '    'DoOperationChanged(e)

    '    'DoOperationStarted(New OperationStartedArgs(OperationTypes.PopulateDatabase, "Create Tables", Windows.Forms.ProgressBarStyle.Continuous))
    '    'DoCancelOperationValueChanged(True)

    '    'rowCount = 0
    '    'rowPos = 0

    '    Dim s As String = "", s2() As String = Nothing
    '    Dim adapter As SqlDataAdapter = Nothing
    '    Dim cmd As SqlCommand = Nothing
    '    'Dim e2 As New OperationStatsArgs(OperationTypes.CreateDatabase)
    '    'e2.Begin()

    '    Dim bError As Boolean
    '    Try

    '        adapter = New SqlDataAdapter("", OpenDatabase(DatabaseName))
    '        'AddHandler adapter.RowUpdated, AddressOf adapter_RowUpdated
    '        'adapter.UpdateBatchSize = 20
    '        cmd = New SqlCommand("", sharedConnection)
    '        For Each t As DataTable In d.Tables
    '            'If bolQuick Then
    '            '    If Not quickList.Contains(t.TableName) Then Continue For
    '            'End If
    '            'rowCount += t.Rows.Count

    '            s = "CREATE TABLE " & t.TableName & " ("
    '            'For Each g As DataRelation In t.ParentRelations
    '            '    'Debug.Print(g.ChildKeyConstraint.)
    '            'Next
    '            fields.Add(t.TableName, New List(Of String))

    '            Erase s2

    '            For Each c As DataColumn In t.Columns
    '                If c.ColumnName = "mame_Id" Then Continue For
    '                If notnulls.ContainsKey(t.TableName) AndAlso notnulls(t.TableName).Contains(c.ColumnName) Then
    '                    c.AllowDBNull = False
    '                End If
    '                If defaults.ContainsKey(t.TableName) AndAlso defaults(t.TableName).ContainsKey(c.ColumnName) Then
    '                    c.DefaultValue = defaults(t.TableName)(c.ColumnName)
    '                End If

    '                If s2 Is Nothing Then ReDim s2(0) Else ReDim Preserve s2(UBound(s2) + 1)

    '                Dim hint As Object = Nothing
    '                If ranges.ContainsKey(t.TableName) AndAlso ranges(t.TableName).ContainsKey(c.ColumnName) Then
    '                    hint = ranges(t.TableName)(c.ColumnName)
    '                End If

    '                s2(UBound(s2)) = "[" & c.ColumnName & "] " & GetDBType(c.DataType, False, hint) & IIf(c.MaxLength > -1, "(" & c.MaxLength & ")", "")
    '                fields(t.TableName).Add("[" & c.ColumnName & "]")
    '                If c.DefaultValue IsNot System.DBNull.Value Then s2(UBound(s2)) &= " DEFAULT " & c.DefaultValue.ToString
    '                If t.PrimaryKey.Contains(c) Then
    '                    s2(UBound(s2)) &= " PRIMARY KEY"
    '                ElseIf c.Unique Then
    '                    s2(UBound(s2)) &= " UNIQUE"
    '                ElseIf c.AllowDBNull = False Then
    '                    s2(UBound(s2)) &= " NOT NULL"
    '                Else
    '                    If t.ParentRelations.Count Then
    '                        For Each da As DataRelation In t.ParentRelations
    '                            For Each ta As DataColumn In da.ParentColumns
    '                                If ta.ColumnName = c.ColumnName Then
    '                                    'With DirectCast(da.ParentColumns(Array.IndexOf(da.ParentColumns, c)), DataColumn)
    '                                    s2(UBound(s2)) &= " foreign key references " & da.ParentTable.TableName & "("
    '                                    Dim q() As String = Nothing
    '                                    For Each cc As DataColumn In da.ParentColumns
    '                                        If q Is Nothing Then ReDim q(0) Else ReDim Preserve q(UBound(q) + 1)
    '                                        q(UBound(q)) = cc.ColumnName
    '                                    Next cc
    '                                    s2(UBound(s2)) &= Join(q, ", ") & ") on delete cascade"
    '                                    'Debug.Print(da.ParentKeyConstraint.ExtendedProperties.ToString)
    '                                    'End With
    '                                End If
    '                            Next ta
    '                        Next da
    '                    End If
    '                End If
    '            Next c
    '            Select Case t.TableName
    '                Case "game"
    '                    Dim i As Integer = UBound(s2)
    '                    ReDim Preserve s2(UBound(s2) + 7)
    '                    s2(i + 1) = "mature bit default 0"
    '                    s2(i + 2) = "verified bit default 0"
    '                    s2(i + 3) = "category varchar(255)"
    '                    s2(i + 4) = "playtime int default 0"
    '                    s2(i + 5) = "playcount int default 0"
    '                    s2(i + 6) = "playlast varchar(255) default ''"
    '                    s2(i + 7) = "favorite bit default 0"
    '                Case "mame"
    '                    Dim i As Integer = UBound(s2)
    '                    ReDim Preserve s2(UBound(s2) + 1)
    '                    s2(i + 1) = "mamepath varchar(255)"
    '                    's2(i + 2) = "quick bit"
    '                    t.Columns.Add("mamepath", GetType(System.String))
    '                    t.Rows(0)![mamepath] = strMame
    '                    fields("mame").Add("mamepath")
    '                    't.Columns.Add("quick", GetType(Boolean))
    '                    't.Rows(0)![quick] = bolQuick
    '                    'fields("mame").Add("quick")
    '            End Select
    '            s &= Join(s2, ", ") & ")"

    '            Debug.Print(s)
    '            cmd.CommandText = s
    '            Try
    '                cmd.ExecuteNonQuery()
    '            Catch ex As Exception
    '                Debug.Print("Error: " & ex.Message)
    '            End Try
    '        Next t

    '        For Each t As DataTable In d.Tables
    '            'If bolQuick Then
    '            '    If Not quickList.Contains(t.TableName) Then Continue For
    '            'End If
    '            'e2.Description = "Populating Table '" & t.TableName & "'"
    '            'e2.Count += 1
    '            'DoUpdateOperationStats(e2)
    '            adapter.SelectCommand.CommandText = "select " & Join(fields(t.TableName).ToArray, ", ") & " from " & t.TableName
    '            Dim o As New SqlCommandBuilder(adapter)
    '            adapter.InsertCommand = o.GetInsertCommand

    '            For Each p As SqlParameter In adapter.InsertCommand.Parameters
    '                If defaults.ContainsKey(t.TableName) AndAlso defaults(t.TableName).ContainsKey(p.SourceColumn) Then
    '                    p.Value = "DEFAULT" 'defaults(t.TableName)(p.SourceColumn)
    '                End If
    '            Next p

    '            adapter.Update(d, t.TableName)

    '        Next t
    '        'e2.Description = "Create Database Complete!"
    '        'e2.Stop()
    '        'DoUpdateOperationStats(e2)
    '    Catch ex As Exception
    '        'Throw ex
    '        'e.Canceled = True
    '        'DoOperationCompleted(New OperationCompleteArgs(OperationTypes.PopulateDatabase, False, Now - startTime, ex.Message, -1))
    '        bError = True
    '    Finally

    '        If adapter IsNot Nothing Then
    '            'RemoveHandler adapter.RowUpdated, AddressOf adapter_RowUpdated
    '            adapter.Dispose()

    '        End If
    '        If cmd IsNot Nothing Then cmd.Dispose()

    '        d.Dispose()
    '        'CloseDatabase()
    '        'If SuppressFinalEvent = False Then DoOperationCompleted(New OperationCompleteArgs(OperationTypes.PopulateDatabase, True, Now - startTime, "Created Tables", rowCount))
    '        Debug.Print("DONE")
    '        'e.IsComplete = True
    '        'DoOperationChanged(e)
    '        GC.Collect()
    '        GC.WaitForPendingFinalizers()
    '        If _FinishDelegate IsNot Nothing Then
    '            _FinishDelegate.Invoke("CreateTables", Not bError)
    '        End If
    '    End Try
    '    Return Not bError
    'End Function
    Public Shared Sub CancelOperation()
        If MyThread IsNot Nothing Then MyThread.RequestStop()
    End Sub
    Public Shared ReadOnly Property DatabaseFile(ByVal strMame As String) As String
        Get
            Return IO.Path.Combine(IO.Path.GetDirectoryName(strMame), IO.Path.GetFileNameWithoutExtension(strMame) & ".mdf")
        End Get
    End Property

    
    
    Private Shared Function _CreateTablesFromMame2(ByVal oe As Object) As Boolean  'ByVal strMame As String, Optional ByVal bolQuick As Boolean = False)
        Dim startTime = Now

        'Dim strMame As String = oe(0)
        'Dim bolQuick As Boolean = oe(1)
        'Dim SuppressFinalEvent As Boolean = oe(2)
        'Dim _FinishDelegate As OperationCompleteDelegate = Nothing
        'Dim _ProgressDelegate As OperationProgressUpdateDelegate = Nothing
        'If UBound(oe) > 0 Then
        'If oe(1) IsNot Nothing Then _FinishDelegate = oe(1)
        'If oe(2) IsNot Nothing Then _ProgressDelegate = oe(2)
        'End If
        'If UBound(oe) >= 3 AndAlso oe(3) IsNot Nothing Then
        'DirectCast(oe(3), Threading.AutoResetEvent).WaitOne()
        'End If
        Dim m As MakeTablesOperation = oe
        m.Wait()

        'Debug.Print("Start ")
        GC.Collect()
        GC.WaitForPendingFinalizers()

        'startTime = Now
        'Dim e As New OperationStartedArgs(OperationTypes.CreateDatabase, "Create Tables", Windows.Forms.ProgressBarStyle.Marquee)
        'Dim e2 As New OperationStatsArgs(OperationTypes.CreateDatabase)
        'If _ProgressDelegate IsNot Nothing Then _ProgressDelegate.Invoke(New OperationProgressArgs("CreateTables", "Preparing data...", -1))
        m.TryInvokeProgress(New OperationProgressArgs("Create Tables", "Preparing Data...", -1))
        Dim s1 As String

        '<old way
        'Dim UserStats As DataTable = Nothing
        'If TablesExists Then

        '    OpenDatabase()
        '    Try
        '        UserStats = GetTable("select * from UserStats")
        '    Catch
        '        UserStats = Nothing
        '    End Try
        '    CloseDatabase()

        '    If Not DeleteNamespace() Then
        '        'DoOperationCompleted(New OperationCompleteArgs(e.Operation, False, Now - startTime, "Can't delete namespace", -1))
        '        If _FinishDelegate IsNot Nothing Then _FinishDelegate.Invoke("CreateTables", False)
        '        Return False
        '    End If
        'End If

        'If NamespaceExists() = True Then
        '    DeleteNamespace()
        'End If
        'CloseDatabase()
        'If IO.File.Exists(IO.Path.Combine(IO.Path.GetDirectoryName(strMame), IO.Path.GetFileNameWithoutExtension(strMame) & ".ndf")) Then
        '    Kill(IO.Path.Combine(IO.Path.GetDirectoryName(strMame), IO.Path.GetFileNameWithoutExtension(strMame) & ".ndf"))
        'End If
        'CreateNamespace(strMame, True)
        '/>

        '<new way doesn't work :(
        Dim needNewUsersStats As Boolean = True
        'If NamespaceExists() Then
        '    CloseDatabase(True)
        '    If DetachDatabase() Then
        '        'files are now seperated
        '        'kill the readonly table
        '        My.Computer.FileSystem.DeleteFile(IO.Path.Combine(IO.Path.GetDirectoryName(strMame), IO.Path.GetFileNameWithoutExtension(strMame) & ".ndf"), FileIO.UIOption.OnlyErrorDialogs, FileIO.RecycleOption.DeletePermanently, FileIO.UICancelOption.DoNothing)
        '        'reattach; this keeps the userstats
        '        If Not AttachDatabase(strMame, True) Then
        '            'throw error
        '        End If
        '    Else

        '    End If
        'Else
        '    needNewUsersStats = True
        '    If Not CreateNamespace(strMame, True) Then
        '        'throw error

        '    End If
        'End If
        '/>

        '< the new NEW way
        If NamespaceExists() Then
            'hope the cascade works!
            If TablesExists Then
                Try
                    CloseDatabase(True)
                    MakeReadWrite()
                Catch
                End Try
                'Dim t = GetTable("select TABLE_NAME from INFORMATION_SCHEMA.TABLES")
                'Dim tables() As String = Nothing
                'For Each r As DataRow In t.Rows
                '    'Debug.Print(r(0))
                '    If r(0).ToString.ToLower <> "userstats" Then
                '        If tables Is Nothing Then ReDim tables(0) Else ReDim Preserve tables(UBound(tables) + 1)
                '        tables(UBound(tables)) = r(0)
                '        DoSQL("delete from " & r(0))
                '    End If
                'Next

                'Dim t As DataTable = GetTable("SELECT K_Table = FK.TABLE_NAME, FK_Column = CU.COLUMN_NAME, PK_Table = PK.TABLE_NAME, PK_Column = PT.COLUMN_NAME, " & _
                '"Constraint_Name = C.CONSTRAINT_NAME FROM INFORMATION_SCHEMA.REFERENTIAL_CONSTRAINTS C INNER JOIN INFORMATION_SCHEMA.TABLE_CONSTRAINTS FK ON C.CONSTRAINT_NAME = FK.CONSTRAINT_NAME " & _
                '"INNER JOIN INFORMATION_SCHEMA.TABLE_CONSTRAINTS PK ON C.UNIQUE_CONSTRAINT_NAME = PK.CONSTRAINT_NAME INNER JOIN INFORMATION_SCHEMA.KEY_COLUMN_USAGE CU ON C.CONSTRAINT_NAME = CU.CONSTRAINT_NAME " & _
                '"INNER JOIN (SELECT i1.TABLE_NAME, i2.COLUMN_NAME FROM INFORMATION_SCHEMA.TABLE_CONSTRAINTS i1 INNER JOIN INFORMATION_SCHEMA.KEY_COLUMN_USAGE i2 ON i1.CONSTRAINT_NAME = i2.CONSTRAINT_NAME " & _
                '"WHERE i1.CONSTRAINT_TYPE = 'PRIMARY KEY') PT ON PT.TABLE_NAME = PK.TABLE_NAME")

                'For Each r As DataRow In t.Rows
                '    DoSQL("alter table " & r(0) & " DROP CONSTRAINT " & r("Constraint_Name"))
                '    'DoSQL("alter table " & r(0) & " DROP CONSTRAINT " & r("FK_Column"))
                '    'DoSQL("DROP INDEX " & r("PK_Column") & " ON " & r("PK_Table"))
                'Next
                'For Each r As DataRow In t.Rows
                '    Dim i = DoSQL("DROP TABLE " & r(0))
                'Next
                'Dim i As Integer
                'i = DoSQL("delete from game")
                'i = DoSQL("delete from mame")

                Dim t = GetTable("select TABLE_NAME from INFORMATION_SCHEMA.TABLES")
                Dim tables As New List(Of String)
                For Each r As DataRow In t.Rows
                    'Debug.Print(r(0))
                    If r(0).ToString.ToLower <> "userstats" Then
                        'If tables Is Nothing Then ReDim tables(0) Else ReDim Preserve tables(UBound(tables) + 1)
                        tables.Add(r(0))
                        'DoSQL("drop table " & r(0))
                    Else
                        needNewUsersStats = False
                    End If
                Next
                Do Until tables Is Nothing OrElse tables.Count = 0
                    For z As Integer = 0 To tables.Count - 1
                        If z > tables.Count - 1 Then Exit For

                        Dim i As Integer
                        If TrySQL("drop table " & tables(z), i) Then
                            tables.Remove(tables(z))
                        End If

                    Next
                Loop
            End If
        Else
            needNewUsersStats = True
            CreateNamespace(m.MameExe, True)
        End If
        '/>
        'Return False
        'Dim lines As New List(Of String)
        Dim xreader = RedirectStream(m.MameExe, "-lx")
        Dim defaults As New Dictionary(Of String, Dictionary(Of String, String))
        Dim notnulls As New Dictionary(Of String, List(Of String))
        Dim ranges As New Dictionary(Of String, Dictionary(Of String, String))
        Dim fields As New Dictionary(Of String, List(Of String))
        'Dim quickList() As String = {"mame", "game"}
        With xreader
            While .Peek > 0
                s1 = .ReadLine.Trim
                'lines.Add(s1)
                'sb.AppendLine(s)
                If s1.StartsWith("<!ATTLIST") Then
                    Dim u() As String = Split(s1.Replace("<", "").Replace(">", ""), " ")
                    For t As Integer = UBound(u) To 2 Step -1
                        If u(t) = "#REQUIRED" Then
                            If Not notnulls.ContainsKey(u(1)) Then notnulls.Add(u(1), New List(Of String))
                            notnulls(u(1)).Add(u(2))
                        End If
                        If u(t).EndsWith(Chr(34)) Then
                            If Not defaults.ContainsKey(u(1)) Then defaults.Add(u(1), New Dictionary(Of String, String))
                            defaults(u(1)).Add(u(2), u(t).Replace(Chr(34), "'"))
                        End If
                        If u(t).StartsWith("(") Then
                            If Not ranges.ContainsKey(u(1)) Then ranges.Add(u(1), New Dictionary(Of String, String))
                            ranges(u(1)).Add(u(2), u(t).Replace("(", "").Replace(")", ""))
                        End If
                    Next
                End If
                If s1 = "]>" Then Exit While
            End While
            .Close()

        End With
        Dim d As New DataSet("Mame")
        'Return False
        xreader.Dispose()
        xreader = RedirectStream(m.MameExe, "-lx")

        d.ReadXml(xreader, XmlReadMode.InferTypedSchema)
        xreader.Dispose()
        'infered gives good valuetypes
        'infer sucks
        'ignore does nothing
        'fragment breaks
        'diff breask   

        'e.IsComplete = True
        'DoOperationChanged(e)
        'e.Operation = OperationChangedArgs.OperationTypes.PopulateDatabase
        'e.Description = "Populating Database"
        'e.HasPercent = True
        'e.CanCancel = True
        'e.IsComplete = False
        'DoOperationChanged(e)

        'DoOperationStarted(New OperationStartedArgs(OperationTypes.PopulateDatabase, "Create Tables", Windows.Forms.ProgressBarStyle.Continuous))
        'DoCancelOperationValueChanged(True)

        'rowCount = 0
        'rowPos = 0

        Dim s As String = "", s2() As String = Nothing
        Dim adapter As SqlDataAdapter = Nothing
        Dim cmd As SqlCommand = Nothing
        'Dim e2 As New OperationStatsArgs(OperationTypes.CreateDatabase)
        'e2.Begin()

        Dim bError As Boolean
        Try
            'CloseDatabase()
            adapter = New SqlDataAdapter("", OpenDatabase(DatabaseName))
            'AddHandler adapter.RowUpdated, AddressOf adapter_RowUpdated
            'adapter.UpdateBatchSize = 20
            cmd = New SqlCommand("", sharedConnection)
            'GoTo fillUserStats

            For Each t As DataTable In d.Tables
                'If bolQuick Then
                '    If Not quickList.Contains(t.TableName) Then Continue For
                'End If
                'rowCount += t.Rows.Count

                s = "CREATE TABLE " & t.TableName & " ("
                'For Each g As DataRelation In t.ParentRelations
                '    'Debug.Print(g.ChildKeyConstraint.)
                'Next
                fields.Add(t.TableName, New List(Of String))

                Erase s2

                For Each c As DataColumn In t.Columns
                    If c.ColumnName = "mame_Id" Then Continue For
                    If notnulls.ContainsKey(t.TableName) AndAlso notnulls(t.TableName).Contains(c.ColumnName) Then
                        c.AllowDBNull = False
                    End If
                    If defaults.ContainsKey(t.TableName) AndAlso defaults(t.TableName).ContainsKey(c.ColumnName) Then
                        c.DefaultValue = defaults(t.TableName)(c.ColumnName)
                    End If

                    If s2 Is Nothing Then ReDim s2(0) Else ReDim Preserve s2(UBound(s2) + 1)

                    Dim hint As Object = Nothing
                    If ranges.ContainsKey(t.TableName) AndAlso ranges(t.TableName).ContainsKey(c.ColumnName) Then
                        hint = ranges(t.TableName)(c.ColumnName)
                    End If

                    s2(UBound(s2)) = "[" & c.ColumnName & "] " & GetDBType(c.DataType, False, hint) & IIf(c.MaxLength > -1, "(" & c.MaxLength & ")", "")
                    fields(t.TableName).Add("[" & c.ColumnName & "]")
                    If c.DefaultValue IsNot System.DBNull.Value Then s2(UBound(s2)) &= " DEFAULT " & c.DefaultValue.ToString
                    If t.PrimaryKey.Contains(c) Then
                        s2(UBound(s2)) &= " PRIMARY KEY"
                    ElseIf c.Unique Then
                        s2(UBound(s2)) &= " UNIQUE"
                    ElseIf c.AllowDBNull = False Then
                        s2(UBound(s2)) &= " NOT NULL"
                    Else
                        If t.ParentRelations.Count Then
                            For Each da As DataRelation In t.ParentRelations
                                For Each ta As DataColumn In da.ParentColumns
                                    If ta.ColumnName = c.ColumnName Then
                                        'With DirectCast(da.ParentColumns(Array.IndexOf(da.ParentColumns, c)), DataColumn)
                                        s2(UBound(s2)) &= " foreign key references " & da.ParentTable.TableName & "("
                                        Dim q() As String = Nothing
                                        For Each cc As DataColumn In da.ParentColumns
                                            If q Is Nothing Then ReDim q(0) Else ReDim Preserve q(UBound(q) + 1)
                                            q(UBound(q)) = cc.ColumnName
                                        Next cc
                                        s2(UBound(s2)) &= Join(q, ", ") & ") on delete cascade"
                                        'Debug.Print(da.ParentKeyConstraint.ExtendedProperties.ToString)
                                        'End With
                                    End If
                                Next ta
                            Next da
                        End If
                    End If
                Next c
                Select Case t.TableName
                    'Case "game"
                    '    ReDim Preserve s2(UBound(s2) + 1)
                    '    s2(UBound(s2)) = "verified bit default 0"
                    '    Dim i As Integer = UBound(s2)
                    '    ReDim Preserve s2(UBound(s2) + 7)
                    '    s2(i + 1) = "mature bit default 0"
                    '    s2(i + 2) = "verified bit default 0"
                    '    s2(i + 3) = "category varchar(255)"
                    '    s2(i + 4) = "playtime int default 0"
                    '    s2(i + 5) = "playcount int default 0"
                    '    s2(i + 6) = "playlast varchar(255) default ''"
                    '    s2(i + 7) = "favorite bit default 0"
                    Case "mame"
                        Dim i As Integer = UBound(s2)
                        ReDim Preserve s2(UBound(s2) + 1)
                        s2(i + 1) = "mamepath varchar(255)"
                        's2(i + 2) = "quick bit"
                        t.Columns.Add("mamepath", GetType(System.String))
                        t.Rows(0)![mamepath] = m.MameExe
                        fields("mame").Add("mamepath")
                        't.Columns.Add("quick", GetType(Boolean))
                        't.Rows(0)![quick] = bolQuick
                        'fields("mame").Add("quick")
                End Select
                s &= Join(s2, ", ") & ") ON ReadOnlyDB"

                Debug.Print(s)
                cmd.CommandText = s
                m.TryInvokeProgress(New OperationProgressArgs("Creating Tables", "Creating Table '" & t.TableName & "'", (d.Tables.IndexOf(t) / (d.Tables.Count - 1) * 50)))
                Try
                    cmd.ExecuteNonQuery()
                Catch ex As Exception
                    Debug.Print("Error: " & ex.Message)
                End Try
            Next t

            'create UserStats table
            If needNewUsersStats Then
                cmd.CommandText = "CREATE TABLE UserStats ([name] nvarchar(255) NOT NULL, " & _
                    "mature bit default 0, category varchar(255), playtime int default 0, playcount int default 0, " & _
                    "playlast varchar(255) default '', favorite bit default 0, verified bit default 0)"
                cmd.ExecuteNonQuery()
            End If
            'Dim bulk As New SqlBulkCopy(cmd.Connection)

            'bulk.WriteToServer(d.CreateDataReader)
            For Each t As DataTable In d.Tables
                'If bolQuick Then
                '    If Not quickList.Contains(t.TableName) Then Continue For
                'End If
                'e2.Description = "Populating Table '" & t.TableName & "'"
                'e2.Count += 1
                'DoUpdateOperationStats(e2)
                'bulk.DestinationTableName = t.TableName
                'bulk.ColumnMappings.Clear()
                'For Each c As DataColumn In t.Columns
                '    Select Case t.TableName
                '        Case "mame"
                '            Select Case c.ColumnName
                '                Case "mamepath"
                '                Case Else
                '                    bulk.ColumnMappings.Add(New SqlBulkCopyColumnMapping(c.ColumnName, c.ColumnName))
                '            End Select
                '        Case "game"
                '            Select Case c.ColumnName
                '                Case "verified"
                '                Case Else
                '                    bulk.ColumnMappings.Add(New SqlBulkCopyColumnMapping(c.ColumnName, c.ColumnName))
                '            End Select
                '        Case Else
                '            Select Case c.ColumnName
                '                Case "mame_Id"
                '                Case Else
                '                    bulk.ColumnMappings.Add(New SqlBulkCopyColumnMapping(c.ColumnName, c.ColumnName))
                '            End Select

                '    End Select
                'Next


                Dim trans = sharedConnection.BeginTransaction(IsolationLevel.Serializable)
                cmd.Transaction = trans
                cmd.CommandText = "select " & Join(fields(t.TableName).ToArray, ", ") & " from " & t.TableName
                adapter.SelectCommand = cmd '.CommandText = "select " & Join(fields(t.TableName).ToArray, ", ") & " from " & t.TableName

                Dim o As New SqlCommandBuilder(adapter)

                adapter.InsertCommand = o.GetInsertCommand.Clone
                If defaults.ContainsKey(t.TableName) Then
                    For Each p As SqlParameter In adapter.InsertCommand.Parameters
                        If defaults(t.TableName).ContainsKey(p.SourceColumn) Then
                            p.Value = "DEFAULT" 'defaults(t.TableName)(p.SourceColumn)
                        End If
                    Next p
                End If
                o.DataAdapter = Nothing

                m.TryInvokeProgress(New OperationProgressArgs("Creating Tables", "Updating Table '" & t.TableName & "'", 50 + (d.Tables.IndexOf(t) / (d.Tables.Count - 1) * 50)))
                adapter.Update(d, t.TableName)
                trans.Commit()
                o.Dispose()
                trans.Dispose()
                'bulk.WriteToServer(t)

            Next t

            'fillUserStats:
            'If needNewUsersStats Then
            '    'If UserStats IsNot Nothing Then
            '    '    For Each r As DataRow In d.Tables("game").Rows
            '    '        cmd.CommandText = "select * from userstats where game_Id=" & r.Item("game_ID")
            '    '        If cmd.ExecuteScalar Is Nothing Then
            '    '            UserStats.Rows.Add(r.Item("game_Id"), Nothing, Nothing, Nothing, Nothing, Nothing, Nothing, Nothing)
            '    '        End If
            '    '    Next
            '    '    'adapter.Fill(UserStats)
            '    'Else
            '    Dim UserStats As New DataTable("userstats")
            '    For Each r As DataRow In d.Tables("game").Rows
            '        UserStats.Rows.Add(r.Item("name"), Nothing, Nothing, Nothing, Nothing, Nothing, Nothing, Nothing)
            '    Next
            '    'End If
            '    adapter.SelectCommand.CommandText = "select * from userstats"
            '    Dim o2 = New SqlCommandBuilder(adapter)

            '    adapter.InsertCommand = o2.GetInsertCommand
            '    adapter.Update(UserStats)
            'End If

            'e2.Description = "Create Database Complete!"
            'e2.Stop()
            'DoUpdateOperationStats(e2)
            CloseDatabase()

            'cmd.CommandText = "DBCC ShrinkDatabase (" & DatabaseName & ")"
            'cmd.ExecuteNonQuery()
            ShrinkDatabase()

            MakeReadOnly()
        Catch ex As Exception
            'Throw ex
            'e.Canceled = True
            'DoOperationCompleted(New OperationCompleteArgs(OperationTypes.PopulateDatabase, False, Now - startTime, ex.Message, -1))
            bError = True
            Throw ex
        Finally

            If adapter IsNot Nothing Then
                'RemoveHandler adapter.RowUpdated, AddressOf adapter_RowUpdated
                adapter.Dispose()

            End If
            If cmd IsNot Nothing Then cmd.Dispose()

            d.Dispose()
            'CloseDatabase()
            'If SuppressFinalEvent = False Then DoOperationCompleted(New OperationCompleteArgs(OperationTypes.PopulateDatabase, True, Now - startTime, "Created Tables", rowCount))
            Debug.Print("DONE")
            'e.IsComplete = True
            'DoOperationChanged(e)
            GC.Collect()
            GC.WaitForPendingFinalizers()
            'If _FinishDelegate IsNot Nothing Then
            '    _FinishDelegate.Invoke(New OperationCompleteArgs("CreateTables", Not bError))
            'End If
            m.TryInvokeComplete(New OperationCompleteArgs("Create Tables", Not bError))
            Debug.Print((Now - startTime).ToString)
        End Try
        Return Not bError
    End Function

    Private Shared Function _VerifyRoms(ByVal o As Object) As Boolean
        'Dim op As OperationCompleteArgs = Nothing
        'Dim Suppress As Boolean = o(2)
        'Dim e As New OperationChangedArgs(OperationChangedArgs.OperationTypes.VerifingRoms, "Verify Roms", False, 0, True, False, False)
        'startTime = Now
        'DoOperationStarted(New OperationStartedArgs(OperationTypes.VerifingRoms, "Verifing Roms...", Windows.Forms.ProgressBarStyle.Marquee))
        'Dim _FinishDelegate As OperationCompleteDelegate = Nothing
        'Dim _ProgressDelegate As OperationProgressUpdateDelegate = Nothing
        'If UBound(o) > 0 Then
        '    If o(1) IsNot Nothing Then _FinishDelegate = o(1)
        '    If o(2) IsNot Nothing Then _ProgressDelegate = o(2)
        'End If
        'If UBound(o) >= 3 AndAlso o(3) IsNot Nothing Then
        '    DirectCast(o(3), Threading.AutoResetEvent).WaitOne()
        'End If
        'Dim ForceProgress As Boolean = True
        Dim m As VerifyRomsOperation = o
        m.Wait()
        Dim Proc As New Process     'Dim s As String = ""
        Dim sr As StreamReader = Nothing
        'If _ProgressDelegate IsNot Nothing Then _ProgressDelegate.Invoke(New OperationProgressArgs("Verfiying Roms", "Initializing...", -1))
        m.TryInvokeProgress(New OperationProgressArgs("Verifying Roms", "Initializing...", -1))
        OpenDatabase(DatabaseName)
        'Dim con = New SqlConnection("Integrated Security=TRUE;Initial Catalog=" & DatabaseName & ";Data Source=.\SQLEXPRESS;Connection Timeout=0;")
        'con.Open()
        'Dim cmd As New SqlCommand("select game_Id, name from game where name=" & "@romname", con)
        'cmd.Parameters.AddWithValue("@romname", "")

        'RomNameID = cmd.ExecuteScalar
        Dim trans As SqlTransaction = sharedConnection.BeginTransaction
        Dim c As New SqlCommand("", sharedConnection, trans)
        'Dim s As String
        c.CommandText = "UPDATE userstats SET verified=0"
        c.ExecuteNonQuery()

        c.CommandText = "IF EXISTS(select name from userstats where name=@romname) " & _
        "UPDATE userstats SET verified=@verified from userstats where name=@romname " & _
        "ELSE INSERT INTO userstats (name, verified) VALUES (@romname, @verified)"
        'c.Parameters.AddWithValue("@game_Id", 0)
        c.Parameters.AddWithValue("@romname", "")
        c.Parameters.AddWithValue("@verified", False)
        'Dim e2 As New OperationStatsArgs(OperationTypes.VerifingRoms)
        'e2.Begin()
        Dim bError As Boolean
        Try

            With Proc
                .StartInfo.UseShellExecute = False
                .StartInfo.RedirectStandardOutput = True
                .StartInfo.FileName = m.MameExe  'o(0) 'strPath
                .StartInfo.Arguments = "-verifyroms" 'strParams
                .StartInfo.WorkingDirectory = IO.Path.GetDirectoryName(m.MameExe) 'strPath)
                .StartInfo.CreateNoWindow = True
                'RaiseEvent Begin()
                .Start()
                'timeStart = .StartTime
                'timeEnd = Nothing

                Dim t As List(Of String) = Nothing
                If m.ForceProgress Then
                    t = New List(Of String)
                End If
                Dim s As String = ""
                Dim strRom As String = ""
                Dim strParent As String = ""
                Dim i As Integer = 0
                m.TryInvokeProgress(New OperationProgressArgs("Verifying Roms", "Process Started...", -1))

                sr = New StreamReader(.StandardOutput.BaseStream)

                Do Until sr.EndOfStream Or MyThread.RequestedStop
                    s = sr.ReadLine
                    'If InStr(s, vbCrLf) Then
                    Dim ssplit As String() = Split(s, vbCrLf)
                    For Each ss In ssplit
                        i = ParseVerify(ss, strRom, strParent)
                        If i = -1 Then Continue Do
                        'If i = 1 Then
                        '    e2.Count += 1
                        '    e2.Description = "Found rom '" & strRom & "'"
                        'End If
                        If m.ForceProgress Then
                            t.Add(strRom)
                        Else
                            '_ProgressDelegate.Invoke(New OperationProgressArgs("Verifing Roms", "Found Rom '" & strRom & "'", -1))
                            m.TryInvokeProgress(New OperationProgressArgs("Verifying Roms", "Found Rom '" & strRom & "'", -1))
                            c.Parameters("@romname").Value = strRom
                            c.Parameters("@verified").Value = i
                            c.ExecuteNonQuery()
                        End If
                    Next
                    'Else
                    'i = ParseVerify(s, strRom, strParent)
                    'If i = -1 Then Continue Do
                    ''If i = 1 Then
                    ''    e2.Count += 1
                    ''    e2.Description = "Found rom '" & strRom & "'"
                    ''End If
                    ''cmd.Parameters("@romname").Value = strRom
                    ''Dim cc = cmd.ExecuteScalar
                    ''c.Parameters("@game_Id").Value = cc
                    'c.Parameters("@romname").Value = strRom
                    'c.Parameters("@verified").Value = i
                    'c.ExecuteNonQuery()
                    'End If
                    't.Add(s)
                    's &= sr.ReadLine & vbCrLf
                    'System.Windows.Forms.Application.DoEvents()
                    'DoUpdateOperationStats(e2)
                    '_ProgressDelegate.Invoke ("Verifing Roms",
                Loop
                bError = MyThread.RequestedStop
                '.WaitForExit()
                'timeEnd = .ExitTime
                'e2.Stop()
                'e2.Description = "Verify Complete!"
                'OnOperationUpdateStats(e2)
                .Close()
                'e.IsComplete = True
                'DoOperationChanged(e)
                'e = New OperationChangedArgs(OperationChangedArgs.OperationTypes.VerifingRoms, "Verifing Roms", True, 0, True, False, False)
                'DoOperationChanged(e)
                If m.ForceProgress And Not MyThread.RequestedStop Then
                    For q As Integer = 0 To t.Count - 1
                        'ParseVerify(t.Keys(q), t.Values(q))
                        'DoProgressChanged(New ProgressChangedArgs((q / (t.Count - 1)) * 100))

                        '_ProgressDelegate.Invoke(New OperationProgressArgs("Verifing Rom", "Found Rom '" & t(q) & "'", (q / (t.Count - 1)) * 100))
                        m.TryInvokeProgress(New OperationProgressArgs("Verifying Roms", "Found Rom '" & t(q) & "'", (q / (t.Count - 1) * 100)))
                        c.Parameters("@romname").Value = t(q)
                        c.Parameters("@verified").Value = 1
                        c.ExecuteNonQuery()

                        'e.Percent = (q / (t.Count - 1)) * 100
                        'DoOperationChanged(e)
                    Next
                End If
                trans.Commit()
                'Debug.Print(s)

                'op = New OperationCompleteArgs(OperationTypes.VerifingRoms, True, Now - startTime, s, 0)
            End With
        Catch ex As Threading.ThreadAbortException
            If sr IsNot Nothing Then sr.Close()
            trans.Rollback()

            'op = New OperationCompleteArgs(OperationTypes.VerifingRoms, False, Now - startTime, "", 0)
            'e.IsComplete = True
            'e.Canceled = True
            bError = True
        Finally
            Proc.Close()
            If sr IsNot Nothing Then sr.Dispose()
            Proc.Dispose()
            c.Dispose()
            trans.Dispose()

            'cmd.Dispose()
            'con.Dispose()

            'CanCancel = False
            'RaiseEvent VerifyDone(op)
            'If Suppress = False Then DoOperationCompleted(op)
            'e.IsComplete = True
            'DoOperationChanged(e)
            'If _FinishDelegate IsNot Nothing Then
            '    _FinishDelegate.Invoke(New OperationCompleteArgs("VerifyRoms", Not bError))
            'End If
            m.TryInvokeComplete(New OperationCompleteArgs("Verifying Roms", Not bError))
        End Try
        Return Not bError
        'Return s
        'Return consoleApp.StandardOutput.ReadToEnd()
    End Function
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

    Private Shared Function _Categorize(ByVal oe As Object) As Boolean
        'ByVal strMamePath As String, ByVal SuppressFinalComplete As Boolean)
        'Dim strMamePath As String = oe(0)
        'Dim _FinishDelegate As OperationCompleteDelegate = Nothing
        'Dim _ProgressDelegate As OperationProgressUpdateDelegate = Nothing
        'If UBound(oe) > 0 Then
        '    If oe(1) IsNot Nothing Then _FinishDelegate = oe(1)
        '    If oe(2) IsNot Nothing Then _ProgressDelegate = oe(2)
        'End If
        'If UBound(oe) >= 3 AndAlso oe(3) IsNot Nothing Then
        '    DirectCast(oe(3), Threading.AutoResetEvent).WaitOne()
        'End If
        Dim m As CategorizeOperation = oe
        m.Wait()
        'Dim SuppressFinalComplete As Boolean = oe(1)
        'Dim t1 As Date = Now
        'DoOperationStarted(New OperationStartedArgs(OperationTypes.CategorizeDatabase, "Categorize", Windows.Forms.ProgressBarStyle.Continuous))
        'Dim e As New OperationChangedArgs(OperationChangedArgs.OperationTypes.CategorizeDatabase, "Categorize Database", True, 0, True, False, False)
        'DoOperationChanged(e)
        'CanCancel = True
        'Dim o As OperationCompleteArgs = Nothing
        Dim bError As Boolean
        Try
            OpenDatabase(DatabaseName)
            Dim catVer = GenerateCatVerFile(m.MameExe)
            Dim cat32 = GenerateCat32File(m.MameExe)
            If IO.File.Exists(cat32) Then bError = Not ParseCat32(cat32, m.ProgressDelegate)
            If IO.File.Exists(catVer) Then bError = Not ParseCatVer(catVer, , m.ProgressDelegate)
            If IO.File.Exists(cat32) = False AndAlso IO.File.Exists(catVer) = False Then bError = True
            'Dim cmd As New SqlCommand("update mame set catver='" & g & "' where mamepath='" & strmamepath & "'", sharedConnection)
            'cmd.ExecuteNonQuery()
            'cmd.Dispose()
            'e.IsComplete = True
            'o = New OperationCompleteArgs(OperationTypes.CategorizeDatabase, True, Now - t1, "Categorize Complete", 0)
        Catch ex As Threading.ThreadAbortException
            'e.IsComplete = True
            'e.Canceled = True
            'o = New OperationCompleteArgs(OperationTypes.CategorizeDatabase, False, Now - t1, "User Canceled", 0)
            bError = True
        Finally
            'CanCancel = False
            'If SuppressFinalComplete = False Then DoOperationCompleted(o)
            'DoOperationChanged(e)
            'If _FinishDelegate IsNot Nothing Then
            '    _FinishDelegate.Invoke(New OperationCompleteArgs("Categorize", Not bError))
            'End If
            m.TryInvokeComplete(New OperationCompleteArgs("Categorize", Not bError))
        End Try
        Return Not bError
    End Function
    Private Shared Function ParseCat32(ByVal strFile As String, Optional ByVal delegateProgress As OperationProgressUpdateDelegate = Nothing) As Boolean
        Dim strRomName As String

        Dim trans As SqlTransaction = sharedConnection.BeginTransaction
        'Dim ce As New SqlCommand("Update userstats set category=@category, mature=@mature from userstats inner join (select game_id, name from game where name=@name or cloneof=@name)", sharedConnection, trans)
        Dim ce As New SqlCommand("IF NOT EXISTS(Select name from userstats where name=@romname) INSERT INTO userstats (name, category, mature) SELECT name=@romname, category=@category, mature=@mature", _
                         sharedConnection, trans)

        ce.Parameters.AddWithValue("@romname", "")
        ce.Parameters.AddWithValue("@category", "")
        ce.Parameters.AddWithValue("@mature", 0)

        Dim fs As New System.IO.FileStream(strFile, System.IO.FileMode.Open, System.IO.FileAccess.Read)
        Dim d As New System.IO.StreamReader(fs)
        d.BaseStream.Seek(0, System.IO.SeekOrigin.Begin)
        Dim BaseLen As Long = (d.BaseStream.Length)
        Dim current As String = ""
        'Dim e2 As New OperationStatsArgs(OperationTypes.CategorizeDatabase)
        'e2.Begin()
        Dim matures() As String = Nothing
        While d.Peek > -1
            strRomName = d.ReadLine.Trim
            If delegateProgress IsNot Nothing Then delegateProgress.Invoke(New OperationProgressArgs("Categorize", "Updating Roms", ((d.BaseStream.Position / BaseLen) * 100)))
            'ce.Parameters.Clear()
            If strRomName.Length = 0 Then Continue While
            'If Strings.Left(strRomName, 2) = ";;" Then
            '    Dim s() As String = Split(strRomName.Replace(";;", ""), " / ")
            '    's(0) = name [CatVer (rev. 1)]
            '    's(1) = build date [11-Sep-05]
            '    's(2) = build version [MAME .99u2]
            '    's(3) = url [http://www.mameworld.net/catlist]
            '    'ParseCat32 = "v" & s(2).Replace("MAME ", "") & " (" & s(1) & ")"
            '    Continue While
            'End If
            If Left(strRomName, 1) = "[" Then
                'If strRomName <> "[Category]" Then Exit While
                current = strRomName
                'If strRomName = "[FOLDER_SETTINGS]" Or strRomName = "[ROOT_FOLDER]" Then
                current = current.Replace("[", "").Replace("]", "")
                Continue While
            End If
            If current = "FOLDER_SETTINGS" OrElse current = "ROOT_FOLDER" Then Continue While
            'RaiseEvent CategorizeProgressChanged((d.BaseStream.Position / BaseLen) * 100)
            'DoProgressChanged(New ProgressChangedArgs((d.BaseStream.Position / BaseLen) * 100))
            'If SourceArgs IsNot Nothing Then
            '    SourceArgs.Percent = (d.BaseStream.Position / BaseLen) * 100
            '    'DoOperationChanged(SourceArgs)
            'End If
            'If current = "[Category]" Then
            'Dim s() As String = Split(strRomName, "=")
            'strRomName = s(0)
            ce.Parameters("@romname").Value = strRomName
            'If s(1).Contains("*Mature*") Then
            '    ce.Parameters("@mature").Value = 1
            '    s(1) = Replace(s(1), " *Mature*", "")
            'Else
            '    ce.Parameters("@mature").Value = 0
            'End If
            If current = "Adult" Then
                If matures Is Nothing Then ReDim matures(0) Else ReDim Preserve matures(UBound(matures) + 1)
                matures(UBound(matures)) = strRomName
                '                ce.Parameters("@mature").Value = 1
            Else
                If matures IsNot Nothing AndAlso matures.Contains(strRomName) Then ce.Parameters("@mature").Value = 1
                ce.Parameters("@category").Value = current
                ce.ExecuteNonQuery()
                'e2.Count += 1
                'e2.Description = "Category found for '" & strRomName & "'"
                'DoUpdateOperationStats(e2)
            End If

            'End If

        End While
        'e2.Stop()
        'DoUpdateOperationStats(e2)
        trans.Commit()

        d.Close()
        trans.Dispose()
        ce.Dispose()
        d.Dispose()
        fs.Dispose()
    End Function
    Private Shared Function ParseCatVer(ByVal strFile As String, Optional ByRef strVersion As String = "", Optional ByVal delegateProgress As OperationProgressUpdateDelegate = Nothing) As Boolean
        Dim strRomName As String

        Dim trans As SqlTransaction = sharedConnection.BeginTransaction
        'UPDATE UserStats SET category = 'test' FROM UserStats INNER JOIN game ON UserStats.game_Id = game.game_Id WHERE (game.name = 'test') OR (game.cloneof = 'test')
        'Dim ce As New SqlCommand("Update userstats set category=@category, mature=@mature where name=@name or cloneof=@name", sharedConnection, trans)
        'Dim ce As New SqlCommand("UPDATE userstats set category=@category, mature=@mature FROM userstats INNER JOIN game ON userstats.game_Id = game.game_Id WHERE (game.name=@name or game.cloneof=@name)", sharedConnection, trans) ' 1:43
        'Dim ce As New SqlCommand("UPDATE userstats set category=@category, mature=@mature FROM userstats INNER JOIN (select name, cloneof from game where cloneof=@romname) as game ON userstats.name = game.name where userstats.name=@romname", sharedConnection, trans) '1:31
        'Dim ce As New SqlCommand("UPDATE userstats SET category=@category, mature=@mature from userstats WHERE name=@romname IF @@ROWCOUNT=0 INSERT INTO userstats VALUES @category, @mature, @romname")
        Dim ce As New SqlCommand("IF NOT EXISTS(Select name from userstats where name=@romname) INSERT INTO userstats (name, category, mature) SELECT name=@romname, category=@category, mature=@mature", _
                                 sharedConnection, trans)
        ce.Parameters.AddWithValue("@romname", "")
        ce.Parameters.AddWithValue("@category", "")
        ce.Parameters.AddWithValue("@mature", 0)

        Dim fs As New System.IO.FileStream(strFile, System.IO.FileMode.Open, System.IO.FileAccess.Read)
        Dim d As New System.IO.StreamReader(fs)
        d.BaseStream.Seek(0, System.IO.SeekOrigin.Begin)
        Dim BaseLen As Long = (d.BaseStream.Length)
        Dim current As String = ""
        'Dim e2 As New OperationStatsArgs(OperationTypes.CategorizeDatabase)
        'e2.Begin()
        Dim bError As Boolean
        Try
            While d.Peek > -1 AndAlso MyThread.RequestedStop = False
                strRomName = d.ReadLine.Trim
                If delegateProgress IsNot Nothing Then delegateProgress.Invoke(New OperationProgressArgs("Categorize", "Updating Roms", (d.BaseStream.Position / BaseLen) * 100))
                'ce.Parameters.Clear()
                If strRomName.Length = 0 Then Continue While
                If Strings.Left(strRomName, 2) = ";;" Then
                    Dim s() As String = Split(strRomName.Replace(";;", ""), " / ")
                    's(0) = name [CatVer (rev. 1)]
                    's(1) = build date [11-Sep-05]
                    's(2) = build version [MAME .99u2]
                    's(3) = url [http://www.mameworld.net/catlist]
                    strVersion = "v" & s(2).Replace("MAME ", "") & " (" & s(1) & ")"
                    Continue While
                End If
                If Left(strRomName, 1) = "[" Then
                    'If strRomName <> "[Category]" Then Exit While
                    current = strRomName
                    Continue While
                End If
                'RaiseEvent CategorizeProgressChanged((d.BaseStream.Position / BaseLen) * 100)
                'DoProgressChanged(New ProgressChangedArgs((d.BaseStream.Position / BaseLen) * 100))
                'If SourceArgs IsNot Nothing Then
                '    SourceArgs.Percent = (d.BaseStream.Position / BaseLen) * 100
                '    'DoOperationChanged(SourceArgs)
                'End If
                If current = "[Category]" Then
                    Dim s() As String = Split(strRomName, "=")
                    If s(1) = "" Then Continue While
                    strRomName = s(0)
                    ce.Parameters("@romname").Value = strRomName
                    If s(1).Contains("*Mature*") Then
                        ce.Parameters("@mature").Value = 1
                        s(1) = Replace(s(1), "*Mature*", "").Trim
                    ElseIf s(1).Contains("* Mature *") Then
                        ce.Parameters("@mature").Value = 1
                        s(1) = Replace(s(1), "* Mature *", "").Trim

                    Else
                        ce.Parameters("@mature").Value = 0
                    End If
                    ce.Parameters("@category").Value = s(1)

                    ce.ExecuteNonQuery()

                    'e2.Count += 1
                    'e2.Description = "Category found for '" & strRomName & "'"
                    'DoUpdateOperationStats(e2)
                End If

            End While
            bError = MyThread.RequestedStop
            trans.Commit()
        Catch ex As Exception
            trans.Rollback()
            bError = True
            Throw ex
        Finally
            'e2.Stop()
            'DoUpdateOperationStats(e2)

            d.Close()
            trans.Dispose()
            ce.Dispose()
            d.Dispose()
            fs.Dispose()
        End Try
        Return Not bError
    End Function
    Public Shared ReadOnly Property CatVerFileExists(ByVal strMamePath As String) As Boolean
        Get
            Return File.Exists(GenerateCatVerFile(strMamePath))
        End Get
    End Property
    Public Shared ReadOnly Property Cat32FileExists(ByVal strMamePath As String) As Boolean
        Get
            Return IO.File.Exists(GenerateCat32File(strMamePath))
        End Get
    End Property
    Public Shared Function GenerateCatVerFile(ByVal strMamePath As String) As String
        Return Path.Combine(Path.GetDirectoryName(strMamePath), "cat32\Catver.ini")
    End Function
    Private Shared Function GenerateCat32File(ByVal strMamePath As String) As String
        Return Path.Combine(Path.GetDirectoryName(strMamePath), "cat32\Category.ini")
    End Function

    Private Shared Function RedirectStream(ByVal strFileName As String, ByVal strParams As String) As StreamReader
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
    Private Shared Function GetDBType(ByVal theType As System.Type, ByVal bolNoUnicode As Boolean, Optional ByVal texthint As Object = Nothing) As String
        Dim p1 As SqlClient.SqlParameter
        Dim tc As System.ComponentModel.TypeConverter
        p1 = New SqlClient.SqlParameter()
        tc = System.ComponentModel.TypeDescriptor.GetConverter(p1.DbType)
        If tc.CanConvertFrom(theType) Then
            p1.DbType = tc.ConvertFrom(theType.Name)
        Else
            Select Case Type.GetTypeCode(theType)
                Case TypeCode.Boolean
                    p1.SqlDbType = SqlDbType.Bit
                Case TypeCode.Byte, TypeCode.SByte
                    p1.SqlDbType = SqlDbType.TinyInt
                Case TypeCode.Char, TypeCode.Int16
                    p1.SqlDbType = SqlDbType.SmallInt
                Case TypeCode.DateTime
                    p1.SqlDbType = SqlDbType.DateTime
                Case TypeCode.Decimal
                    p1.SqlDbType = SqlDbType.Decimal
                Case TypeCode.Double
                    p1.SqlDbType = SqlDbType.Float
                Case TypeCode.Int32, TypeCode.UInt32, TypeCode.UInt16
                    p1.SqlDbType = SqlDbType.Int
                Case TypeCode.Int64, TypeCode.UInt64
                    p1.SqlDbType = SqlDbType.BigInt
                Case TypeCode.Object
                    p1.SqlDbType = SqlDbType.Variant
                Case TypeCode.Single
                    p1.SqlDbType = SqlDbType.Real
                Case TypeCode.String
                    p1.SqlDbType = SqlDbType.VarChar
            End Select

        End If
        Dim adder As String = ""
        Select Case p1.SqlDbType
            Case SqlDbType.Char, SqlDbType.NChar, SqlDbType.NText, SqlDbType.NVarChar, SqlDbType.Text, SqlDbType.VarChar
                If texthint IsNot Nothing Then
                    Dim s() As String = Split(texthint, "|")
                    Dim longest As Integer = s(0).Length
                    For t As Integer = 1 To UBound(s)
                        If s(t).Length > longest Then longest = s(t).Length
                    Next
                    adder = "(" & longest & ")"
                Else
                    adder = "(255)"
                End If
        End Select
        If bolNoUnicode = True Then
            Select Case p1.SqlDbType
                Case SqlDbType.NChar
                    p1.SqlDbType = SqlDbType.Char
                Case SqlDbType.NText
                    p1.SqlDbType = SqlDbType.Text
                Case SqlDbType.NVarChar
                    p1.SqlDbType = SqlDbType.VarChar

            End Select
        End If
        Return p1.SqlDbType.ToString.ToLower & adder
    End Function


    Private Shared Sub MyThread_OnError(ByVal state As Object) Handles MyThread.OnError

    End Sub
End Class
