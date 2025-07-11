

Public Class clsTabularModel

    Enum enFieldType
        Level = 1
        Measure = 2
    End Enum

    Enum enDataType
        Text = 1
        Number = 2
        DateTime = 3
        Bool = 4
    End Enum

    Public Enum enSelectionMode
        AllSelected = 1
        AllSearch = 2
        DeSelectMember = 3
        SelectMember = 4
    End Enum

    Public Enum enSort
        none = 0
        asc = 1
        desc = 2
    End Enum


    Private Property conn As Object


    Private CubeName As String
    Private ds As DataSet
    Public Cubes As List(Of Cube)

    Private ImageList32 As ImageList
    Private ImageList48 As ImageList

    Public ReadOnly Property AllLevels As List(Of Level)
        Get
            Dim lstRes As New List(Of Level)
            For Each c As Cube In Me.Cubes
                For Each d As Dimension In c.Dimensions
                    For Each h As Hierarchy In d.Hierarchies
                        For Each l As Level In h.Levels
                            If l.nds.Count > 0 Then
                                lstRes.Add(l)
                            End If
                        Next l
                    Next h
                Next d
            Next c
            Return lstRes
        End Get
    End Property

    Public ReadOnly Property AllMeasures As List(Of Measure)
        Get
            Dim lstRes As New List(Of Measure)
            For Each c As Cube In Me.Cubes
                For Each d As Dimension In c.Dimensions
                    For Each m As Measure In d.Measures
                        If m.nds.Count > 0 Then
                            lstRes.Add(m)
                        End If
                    Next m
                Next d
            Next c
            Return lstRes
        End Get
    End Property

    Public ReadOnly Property GetLevel(UniName As String) As Level
        Get
            For Each c As Cube In Me.Cubes
                For Each d As Dimension In c.Dimensions
                    For Each h As Hierarchy In d.Hierarchies
                        For Each l As Level In h.Levels
                            If l.UniName.ToLower = UniName.ToLower Then
                                Return l
                            End If
                        Next l
                    Next h
                Next d
            Next c
            Return Nothing
        End Get
    End Property

    Public ReadOnly Property GetMeasure(UniName As String) As Measure
        Get
            For Each c As Cube In Me.Cubes
                For Each d As Dimension In c.Dimensions
                    For Each m As Measure In d.Measures
                        If m.UniName.ToLower = UniName.ToLower Then
                            Return m
                        End If
                    Next m
                Next d
            Next c
            Return Nothing
        End Get
    End Property


    Public ReadOnly Property IsMD As Boolean
        Get
            Dim blnIsMD As Boolean = False

            If Me.ds Is Nothing Then
                Return blnIsMD
            End If

            Try
                For Each r As DataRow In Me.ds.Tables("CUBES").Rows
                    If r("PREFERRED_QUERY_PATTERNS").ToString = "0" Then
                        blnIsMD = True
                        Exit For
                    End If
                Next r
            Catch ex As Exception
                blnIsMD = False
            End Try


            Return blnIsMD

        End Get
    End Property

    Public ReadOnly Property DimensionCount As Integer
        Get
            Dim intRes As Integer
            For Each cb In Me.Cubes
                intRes += cb.Dimensions.Count
            Next cb
            Return intRes
        End Get
    End Property

    Public ReadOnly Property CubeCount As Integer
        Get
            Return Me.Cubes.Count
        End Get
    End Property

    Friend ctrlDaxQuery As ctrlDaxQuery
    Public Sub New(CubeName As String, ImageList32 As ImageList, ImageList48 As ImageList, ctrlDaxQuery As ctrlDaxQuery)
        Me.Cubes = New List(Of Cube)
        Me.ImageList32 = ImageList32
        Me.ImageList48 = ImageList48
        Me.CubeName = CubeName
        Me.ctrlDaxQuery = ctrlDaxQuery
        Me.conn = Me.ctrlDaxQuery.conn
        Me.Update()
    End Sub

    Public Function GetAllSelectedObjectFields(tabObj As Object) As List(Of Object)

        Dim lstRes As New List(Of Object)

        If tabObj Is Nothing Then
            Return lstRes
        End If


        If Not TryCast(tabObj, clsTabularModel.Level) Is Nothing Then

            lstRes.Add(TryCast(tabObj, clsTabularModel.Level))

        ElseIf Not TryCast(tabObj, clsTabularModel.Measure) Is Nothing Then

            lstRes.Add(TryCast(tabObj, clsTabularModel.Measure))

        ElseIf Not TryCast(tabObj, clsTabularModel.Dimension) Is Nothing Then

            Dim d As clsTabularModel.Dimension = TryCast(tabObj, clsTabularModel.Dimension)
            For Each l In d.LevelsAttrVisible
                lstRes.Add(l)
            Next l
            For Each m In d.MeasuresVisible
                lstRes.Add(m)
            Next m

        ElseIf Not TryCast(tabObj, clsTabularModel.DisplayFolder) Is Nothing Then

            Dim df As clsTabularModel.DisplayFolder = TryCast(tabObj, clsTabularModel.DisplayFolder)
            For Each l As clsTabularModel.Level In df.AllLevelsSortedByName
                lstRes.Add(l)
            Next l
            For Each m As clsTabularModel.Measure In df.AllMeasuresSortedByName
                lstRes.Add(m)
            Next m

        ElseIf Not TryCast(tabObj, clsTabularModel.Hierarchy) Is Nothing Then

            Dim h As clsTabularModel.Hierarchy = TryCast(tabObj, clsTabularModel.Hierarchy)
            For Each l As clsTabularModel.Level In h.LevelsAttrVisible
                lstRes.Add(l)
            Next l


        ElseIf Not TryCast(tabObj, clsTabularModel.Members) Is Nothing Then

            Dim l As clsTabularModel.Level = TryCast(tabObj, clsTabularModel.Members).Level
            lstRes.Add(l)

        ElseIf Not TryCast(tabObj, clsTabularModel.Cube) Is Nothing Then

            Dim c As clsTabularModel.Cube = TryCast(tabObj, clsTabularModel.Cube)
            For Each d As clsTabularModel.Dimension In c.Dimensions
                For Each l As clsTabularModel.Level In d.LevelsAttrVisible
                    lstRes.Add(l)
                Next l
                For Each m As clsTabularModel.Measure In d.MeasuresVisible
                    lstRes.Add(m)
                Next m
            Next d

        End If



        For i As Integer = lstRes.Count - 1 To 0 Step -1
            If lstRes.Item(i).IsSelected = False Then
                lstRes.RemoveAt(i)
            End If
        Next i

        Dim strKeys As New List(Of String)
        For i As Integer = lstRes.Count - 1 To 0 Step -1
            If strKeys.Contains(lstRes.Item(i).UniName.ToString.ToLower) Then
                lstRes.RemoveAt(i)
            Else
                strKeys.Add(lstRes.Item(i).UniName.ToString.ToLower)
            End If
        Next i



        Return lstRes


    End Function



    Public Function GetAllObjectFields(tabObj As Object) As List(Of Object)

        Dim lstRes As New List(Of Object)

        If tabObj Is Nothing Then
            Return lstRes
        End If


        If Not TryCast(tabObj, clsTabularModel.Level) Is Nothing Then

            lstRes.Add(TryCast(tabObj, clsTabularModel.Level))

        ElseIf Not TryCast(tabObj, clsTabularModel.Measure) Is Nothing Then

            lstRes.Add(TryCast(tabObj, clsTabularModel.Measure))

        ElseIf Not TryCast(tabObj, clsTabularModel.Dimension) Is Nothing Then

            Dim d As clsTabularModel.Dimension = TryCast(tabObj, clsTabularModel.Dimension)
            For Each l In d.LevelsAttrVisible
                lstRes.Add(l)
            Next l
            For Each m In d.MeasuresVisible
                lstRes.Add(m)
            Next m

        ElseIf Not TryCast(tabObj, clsTabularModel.DisplayFolder) Is Nothing Then

            Dim df As clsTabularModel.DisplayFolder = TryCast(tabObj, clsTabularModel.DisplayFolder)
            For Each l As clsTabularModel.Level In df.AllLevelsSortedByName
                lstRes.Add(l)
            Next l
            For Each m As clsTabularModel.Measure In df.AllMeasuresSortedByName
                lstRes.Add(m)
            Next m

        ElseIf Not TryCast(tabObj, clsTabularModel.Hierarchy) Is Nothing Then

            Dim h As clsTabularModel.Hierarchy = TryCast(tabObj, clsTabularModel.Hierarchy)
            For Each l As clsTabularModel.Level In h.LevelsAttrVisible
                lstRes.Add(l)
            Next l


        ElseIf Not TryCast(tabObj, clsTabularModel.Members) Is Nothing Then

            Dim l As clsTabularModel.Level = TryCast(tabObj, clsTabularModel.Members).Level
            lstRes.Add(l)

        ElseIf Not TryCast(tabObj, clsTabularModel.Cube) Is Nothing Then

            Dim c As clsTabularModel.Cube = TryCast(tabObj, clsTabularModel.Cube)
            For Each d As clsTabularModel.Dimension In c.Dimensions
                For Each l As clsTabularModel.Level In d.LevelsAttrVisible
                    lstRes.Add(l)
                Next l
                For Each m As clsTabularModel.Measure In d.MeasuresVisible
                    lstRes.Add(m)
                Next m
            Next d

        End If

        Return lstRes

    End Function


    Public Sub Update()

        Me.GetDataSet()



        For Each cr As DataRow In Me.ds.Tables("CUBES").Rows
            If cr("CUBE_NAME").ToString.ToLower = Me.CubeName.ToLower Then
                Dim c As New Cube(cr("CUBE_NAME"), Me.ds, Me)

                For Each dr As DataRow In cr.GetChildRows("Cube_to_Dim")
                    Dim d As New Dimension(c) With {
                        .ds = Me.ds,
                        .CUBE_NAME = dr("CUBE_NAME"),
                        .DIMENSION_NAME = dr("DIMENSION_NAME"),
                        .DIMENSION_UNIQUE_NAME = dr("DIMENSION_UNIQUE_NAME"),
                        .DIMENSION_CAPTION = dr("DIMENSION_CAPTION"),
                        .DIMENSION_ORDINAL = dr("DIMENSION_ORDINAL"),
                        .DIMENSION_TYPE = dr("DIMENSION_TYPE"),
                        .DIMENSION_CARDINALITY = dr("DIMENSION_CARDINALITY"),
                        .DEFAULT_HIERARCHY = dr("DEFAULT_HIERARCHY"),
                        .DESCRIPTION = dr("DESCRIPTION"),
                        .DIMENSION_IS_VISIBLE = dr("DIMENSION_IS_VISIBLE")
                    }

                    For Each mr As DataRow In dr.GetChildRows("Dim_to_Meas")
                        Dim m As New Measure(d) With {
                            .CUBE_NAME = mr("CUBE_NAME"),
                            .MEASURE_NAME = mr("MEASURE_NAME"),
                            .MEASURE_UNIQUE_NAME = mr("MEASURE_UNIQUE_NAME"),
                            .MEASURE_CAPTION = mr("MEASURE_CAPTION"),
                            .MEASURE_AGGREGATOR = mr("MEASURE_AGGREGATOR"),
                            .DATA_TYPE = mr("DATA_TYPE"),
                            .NUMERIC_PRECISION = mr("NUMERIC_PRECISION"),
                            .NUMERIC_SCALE = mr("NUMERIC_SCALE"),
                            .EXPRESSION = mr("EXPRESSION"),
                            .MEASURE_IS_VISIBLE = mr("MEASURE_IS_VISIBLE"),
                            .MEASUREGROUP_NAME = mr("MEASUREGROUP_NAME"),
                            .MEASURE_DISPLAY_FOLDER = mr("MEASURE_DISPLAY_FOLDER"),
                            .DEFAULT_FORMAT_STRING = mr("DEFAULT_FORMAT_STRING")
                        }
                    Next mr

                    For Each hr As DataRow In dr.GetChildRows("Dim_to_Hier")
                        Dim h As New Hierarchy(d) With {
                            .CUBE_NAME = hr("CUBE_NAME"),
                            .DIMENSION_UNIQUE_NAME = hr("DIMENSION_UNIQUE_NAME"),
                            .HIERARCHY_NAME = hr("HIERARCHY_NAME"),
                            .HIERARCHY_UNIQUE_NAME = hr("HIERARCHY_UNIQUE_NAME"),
                            .HIERARCHY_CAPTION = hr("HIERARCHY_CAPTION"),
                            .DIMENSION_TYPE = hr("DIMENSION_TYPE"),
                            .HIERARCHY_CARDINALITY = hr("HIERARCHY_CARDINALITY"),
                            .DEFAULT_MEMBER = hr("DEFAULT_MEMBER"),
                            .ALL_MEMBER = hr("ALL_MEMBER"),
                            .DESCRIPTION = hr("DESCRIPTION"),
                            .DIMENSION_IS_VISIBLE = hr("DIMENSION_IS_VISIBLE"),
                            .HIERARCHY_ORDINAL = hr("HIERARCHY_ORDINAL"),
                            .HIERARCHY_IS_VISIBLE = hr("HIERARCHY_IS_VISIBLE"),
                            .HIERARCHY_ORIGIN = hr("HIERARCHY_ORIGIN"),
                            .HIERARCHY_DISPLAY_FOLDER = hr("HIERARCHY_DISPLAY_FOLDER")
                        }

                        For Each lr As DataRow In hr.GetChildRows("Hier_to_Lev")
                            Dim l As New Level(h) With {
                                .CUBE_NAME = lr("CUBE_NAME"),
                                .DIMENSION_UNIQUE_NAME = lr("DIMENSION_UNIQUE_NAME"),
                                .HIERARCHY_UNIQUE_NAME = lr("HIERARCHY_UNIQUE_NAME"),
                                .LEVEL_NAME = lr("LEVEL_NAME"),
                                .LEVEL_UNIQUE_NAME = lr("LEVEL_UNIQUE_NAME"),
                                .LEVEL_CAPTION = lr("LEVEL_CAPTION"),
                                .LEVEL_NUMBER = lr("LEVEL_NUMBER"),
                                .LEVEL_TYPE = lr("LEVEL_TYPE"),
                                .DESCRIPTION = lr("DESCRIPTION"),
                                .LEVEL_IS_VISIBLE = lr("LEVEL_IS_VISIBLE"),
                                .LEVEL_ORDERING_PROPERTY = lr("LEVEL_ORDERING_PROPERTY"),
                                .LEVEL_DBTYPE = lr("LEVEL_DBTYPE"),
                                .LEVEL_ATTRIBUTE_HIERARCHY_NAME = lr("LEVEL_ATTRIBUTE_HIERARCHY_NAME"),
                                .LEVEL_KEY_CARDINALITY = lr("LEVEL_KEY_CARDINALITY"),
                                .LEVEL_ORIGIN = lr("LEVEL_ORIGIN")
                            }

                        Next lr

                    Next hr

                Next dr

                Me.Cubes.Add(c)
            End If
        Next cr


    End Sub

    Public Class Cube
        Public Property CUBE_NAME As String

        Public Property ds As DataSet

        Public ReadOnly Property Catalog As String
            Get
                Try
                    For Each r As DataRow In Me.ds.Tables("CUBES").Rows
                        Return (r("CATALOG_NAME").ToString)
                    Next r
                Catch ex As Exception
                    Return ""
                End Try
                Return ""
            End Get
        End Property


        Public ReadOnly Property DimensionsSortedByTypeAndName As List(Of Dimension)
            Get
                Me.Dimensions.Sort(Function(x As Dimension, y As Dimension) x.SortByTypeAndName.CompareTo(y.SortByTypeAndName))
                Return Me.Dimensions
            End Get
        End Property



        Private _Dimensions As List(Of Dimension)
        Public ReadOnly Property Dimensions As List(Of Dimension)
            Get
                If Not _Dimensions Is Nothing Then
                    Return _Dimensions
                Else
                    _Dimensions = New List(Of Dimension)
                End If

                For Each r As DataRow In Me.ds.Tables("DIMENSIONS").Rows
                    If r("CUBE_NAME").ToString.ToLower = Me.CUBE_NAME.ToLower Then
                        Dim d As New Dimension(Me) With {
                            .ds = Me.ds,
                            .CUBE_NAME = r("CUBE_NAME").ToString,
                            .DIMENSION_NAME = r("DIMENSION_NAME").ToString,
                            .DIMENSION_UNIQUE_NAME = r("DIMENSION_UNIQUE_NAME").ToString,
                            .DIMENSION_CAPTION = r("DIMENSION_CAPTION").ToString,
                            .DIMENSION_ORDINAL = r("DIMENSION_ORDINAL"),
                            .DIMENSION_TYPE = r("DIMENSION_TYPE"),
                            .DIMENSION_CARDINALITY = r("DIMENSION_CARDINALITY"),
                            .DEFAULT_HIERARCHY = r("DEFAULT_HIERARCHY").ToString,
                            .DESCRIPTION = r("DESCRIPTION").ToString,
                            .DIMENSION_IS_VISIBLE = r("DIMENSION_IS_VISIBLE")
                        }

                        For Each mr As DataRow In r.GetChildRows("Dim_to_Meas")
                            Dim m As New Measure(d) With {
                                .CUBE_NAME = mr("CUBE_NAME").ToString,
                                .MEASURE_NAME = mr("MEASURE_NAME").ToString,
                                .MEASURE_UNIQUE_NAME = mr("MEASURE_UNIQUE_NAME").ToString,
                                .MEASURE_CAPTION = mr("MEASURE_CAPTION").ToString,
                                .MEASURE_AGGREGATOR = mr("MEASURE_AGGREGATOR"),
                                .DATA_TYPE = mr("DATA_TYPE"),
                                .NUMERIC_PRECISION = mr("NUMERIC_PRECISION"),
                                .NUMERIC_SCALE = mr("NUMERIC_SCALE"),
                                .EXPRESSION = mr("EXPRESSION").ToString,
                                .MEASURE_IS_VISIBLE = mr("MEASURE_IS_VISIBLE"),
                                .MEASUREGROUP_NAME = mr("MEASUREGROUP_NAME").ToString,
                                .MEASURE_DISPLAY_FOLDER = mr("MEASURE_DISPLAY_FOLDER").ToString,
                                .DEFAULT_FORMAT_STRING = mr("DEFAULT_FORMAT_STRING").ToString
                            }
                            If m.MEASURE_DISPLAY_FOLDER = "" Then
                                d.Measures.Add(m)
                            Else
                                Dim df_l As List(Of DisplayFolder) = d.GetDisplayFolders(m.MEASURE_DISPLAY_FOLDER)
                                For Each df As DisplayFolder In df_l
                                    m.DisplayFolder = df
                                    df.Measures.Add(m)
                                    d.Measures.Add(m)
                                Next df
                            End If

                        Next mr

                        For Each hr As DataRow In r.GetChildRows("Dim_to_Hier")
                            Dim h As New Hierarchy(d) With {
                                .CUBE_NAME = hr("CUBE_NAME").ToString,
                                .DIMENSION_UNIQUE_NAME = hr("DIMENSION_UNIQUE_NAME").ToString,
                                .HIERARCHY_NAME = hr("HIERARCHY_NAME").ToString,
                                .HIERARCHY_UNIQUE_NAME = hr("HIERARCHY_UNIQUE_NAME").ToString,
                                .HIERARCHY_CAPTION = hr("HIERARCHY_CAPTION").ToString,
                                .DIMENSION_TYPE = hr("DIMENSION_TYPE"),
                                .HIERARCHY_CARDINALITY = hr("HIERARCHY_CARDINALITY"),
                                .DEFAULT_MEMBER = hr("DEFAULT_MEMBER").ToString,
                                .ALL_MEMBER = hr("ALL_MEMBER").ToString,
                                .DESCRIPTION = hr("DESCRIPTION").ToString,
                                .DIMENSION_IS_VISIBLE = hr("DIMENSION_IS_VISIBLE"),
                                .HIERARCHY_ORDINAL = hr("HIERARCHY_ORDINAL"),
                                .HIERARCHY_IS_VISIBLE = hr("HIERARCHY_IS_VISIBLE"),
                                .HIERARCHY_ORIGIN = hr("HIERARCHY_ORIGIN"),
                                .HIERARCHY_DISPLAY_FOLDER = hr("HIERARCHY_DISPLAY_FOLDER").ToString
                            }
                            If h.HIERARCHY_DISPLAY_FOLDER = "" Then
                                d.Hierarchies.Add(h)
                            Else
                                Dim df_l As List(Of DisplayFolder) = d.GetDisplayFolders(h.HIERARCHY_DISPLAY_FOLDER)
                                For Each df As DisplayFolder In df_l
                                    h.DisplayFolder = df
                                    df.Hierarchies.Add(h)
                                    d.Hierarchies.Add(h)
                                Next df
                            End If

                            For Each lr As DataRow In hr.GetChildRows("Hier_to_Lev")
                                Dim l As New Level(h) With {
                                    .CUBE_NAME = lr("CUBE_NAME").ToString,
                                    .DIMENSION_UNIQUE_NAME = lr("DIMENSION_UNIQUE_NAME").ToString,
                                    .HIERARCHY_UNIQUE_NAME = lr("HIERARCHY_UNIQUE_NAME").ToString,
                                    .LEVEL_NAME = lr("LEVEL_NAME").ToString,
                                    .LEVEL_UNIQUE_NAME = lr("LEVEL_UNIQUE_NAME").ToString,
                                    .LEVEL_CAPTION = lr("LEVEL_CAPTION").ToString,
                                    .LEVEL_NUMBER = lr("LEVEL_NUMBER"),
                                    .LEVEL_TYPE = lr("LEVEL_TYPE"),
                                    .DESCRIPTION = lr("DESCRIPTION").ToString,
                                    .LEVEL_IS_VISIBLE = lr("LEVEL_IS_VISIBLE"),
                                    .LEVEL_ORDERING_PROPERTY = lr("LEVEL_ORDERING_PROPERTY").ToString,
                                    .LEVEL_DBTYPE = lr("LEVEL_DBTYPE"),
                                    .LEVEL_ATTRIBUTE_HIERARCHY_NAME = lr("LEVEL_ATTRIBUTE_HIERARCHY_NAME").ToString,
                                    .LEVEL_KEY_CARDINALITY = lr("LEVEL_KEY_CARDINALITY"),
                                    .LEVEL_ORIGIN = lr("LEVEL_ORIGIN")
                                }
                                h.Levels.Add(l)

                            Next lr


                        Next hr


                        _Dimensions.Add(d)
                    End If
                Next r

                Return _Dimensions

            End Get
        End Property

        Friend Property tm As clsTabularModel

        Public Sub New(CUBE_NAME As String, ds As DataSet, tm As clsTabularModel)
            Me.tm = tm
            Me.ds = ds
            Me.CUBE_NAME = CUBE_NAME
        End Sub

        Public Sub SelectLevel(strUniName As String)
            For Each d As clsTabularModel.Dimension In Me.Dimensions
                For Each l As clsTabularModel.Level In d.LevelsAttrVisible
                    If l.UniName.ToLower.Trim = strUniName.ToLower.Trim Then
                        l.IsSelected = True
                        'nd = l.nds.Item(0)
                        If l.LevelSiblings.Count > 0 Then
                            For Each ls As clsTabularModel.Level In l.LevelSiblings
                                ls.IsSelected = True
                            Next ls
                        End If
                        Exit For
                    End If
                Next l
            Next d


        End Sub




        Public Sub DeSelect(strUniName As String)
            Dim nd As C1.Win.C1FlexGrid.Node = Nothing
            For Each d As clsTabularModel.Dimension In Me.Dimensions
                For Each m As clsTabularModel.Measure In d.MeasuresVisible
                    If m.UniName.ToLower.Trim = strUniName.ToLower.Trim Then
                        m.IsSelected = False
                        nd = m.nds.Item(0)
                        Exit For
                    End If
                Next m
                For Each l As clsTabularModel.Level In d.LevelsAttrVisible
                    If l.UniName.ToLower.Trim = strUniName.ToLower.Trim Then
                        l.IsSelected = False
                        nd = l.nds.Item(0)
                        If l.LevelSiblings.Count > 0 Then
                            For Each ls As clsTabularModel.Level In l.LevelSiblings
                                ls.IsSelected = False
                            Next ls
                        End If
                        Exit For
                    End If
                Next l
            Next d

            If Not nd Is Nothing Then
                Dim ndP As C1.Win.C1FlexGrid.Node = nd
                Do While ndP.Level > 1
                    If ndP.Level = 1 Then
                        Exit Do
                    Else
                        ndP = ndP.GetNode(C1.Win.C1FlexGrid.NodeTypeEnum.Parent)
                    End If
                Loop

                Dim ndS As C1.Win.C1FlexGrid.Node = ndP.GetNode(C1.Win.C1FlexGrid.NodeTypeEnum.NextSibling)
                If ndS Is Nothing Then
                    ndS = Me.tm.ctrlDaxQuery.ctrlCube.fg.Rows(Me.tm.ctrlDaxQuery.ctrlCube.fg.Rows.Count - 1).Node
                End If
                For i = ndS.GetCellRange.r1 To ndP.GetCellRange.r1 Step -1
                    If Not TryCast(Me.tm.ctrlDaxQuery.ctrlCube.fg.Rows(i).UserData, clsTabularModel.DisplayFolder) Is Nothing Then
                        TryCast(Me.tm.ctrlDaxQuery.ctrlCube.fg.Rows(i).UserData, clsTabularModel.DisplayFolder).SetDisplaystate()
                    ElseIf Not TryCast(Me.tm.ctrlDaxQuery.ctrlCube.fg.Rows(i).UserData, clsTabularModel.Hierarchy) Is Nothing Then
                        TryCast(Me.tm.ctrlDaxQuery.ctrlCube.fg.Rows(i).UserData, clsTabularModel.Hierarchy).SetDisplayState()
                    ElseIf Not TryCast(Me.tm.ctrlDaxQuery.ctrlCube.fg.Rows(i).UserData, clsTabularModel.Dimension) Is Nothing Then
                        TryCast(Me.tm.ctrlDaxQuery.ctrlCube.fg.Rows(i).UserData, clsTabularModel.Dimension).SetDisplayState()
                    End If
                Next i
            End If



        End Sub




    End Class

    Public Class Dimension

        Public Property Cube As Cube
        Public Property ds As DataSet
        Public Property CUBE_NAME As String
        Public Property DIMENSION_NAME As String
        Public Property DIMENSION_UNIQUE_NAME As String
        Public Property DIMENSION_CAPTION As String


        Public Property DIMENSION_ORDINAL As Integer
        Public Property DIMENSION_TYPE As Integer
        Public Property DIMENSION_CARDINALITY As Integer
        Public Property DEFAULT_HIERARCHY As String
        Public Property DESCRIPTION As String
        Public Property DIMENSION_IS_VISIBLE As Boolean

        Public Property nd As C1.Win.C1FlexGrid.Node

        Public ReadOnly Property HierarchiesSortedbyOrdinal As List(Of Hierarchy)
            Get
                Me.Hierarchies.Sort(Function(x, y) x.HIERARCHY_ORDINAL.CompareTo(y.HIERARCHY_ORDINAL))
                Return Me.Hierarchies
            End Get
        End Property

        Public ReadOnly Property IsSelected As Boolean
            Get
                For Each l In Me.LevelsAttrVisible
                    If l.IsSelected = True Then
                        Return True
                    End If
                Next l
                For Each m In Me.MeasuresVisible
                    If m.IsSelected = True Then
                        Return True
                    End If
                Next m
                Return False
            End Get
        End Property





        Public Enum enDisplayState
            Blank = 0
            Filtered = 1
            Selected = 2
            Tristate = 3
            SelectedAndFiltered = 4
            MouseOverSelected = 5
            MouseOverUnSelected = 6
            MouseOverSelectedAndFiltered = 7
        End Enum

        Public Sub SetDisplayState()
            Dim intCtr As Integer = 0
            Dim intSel As Integer = 0
            Dim intNSel As Integer = 0

            For Each l In Me.LevelsAttrVisible
                intCtr += 1
                If l.IsSelected = True Then
                    intSel += 1
                Else
                    intNSel += 1
                End If
            Next l
            For Each m In Me.MeasuresVisible
                intCtr += 1
                If m.IsSelected = True Then
                    intSel += 1
                Else
                    intNSel += 1
                End If
            Next m

            If intSel = 0 Then
                If Me.DIMENSION_IS_MEASURE_TABLE Then
                    Me.nd.Image = Me.Cube.tm.ImageList32.Images("PBI_MeasureTableUnSelected32.ico")
                    'Me.nd.Image = Me.Cube.tm.ImageList32.Images("PBI_MeasureTableBlank32.ico")
                ElseIf Me.VisibleMeasures = True Then
                    Me.nd.Image = Me.Cube.tm.ImageList32.Images("PBI_TableWithMeasuresUnSelected32.ico")
                    'Me.nd.Image = Me.Cube.tm.ImageList32.Images("PBI_TableWithMeasuresBlank32.ico")
                Else
                    Me.nd.Image = Me.Cube.tm.ImageList32.Images("PBI_TableUnSelected32.ico")
                    'Me.nd.Image = Me.Cube.tm.ImageList32.Images("PBI_TableBlank32.ico")
                End If
            ElseIf intSel <> intCtr Then
                If Me.DIMENSION_IS_MEASURE_TABLE Then
                    Me.nd.Image = Me.Cube.tm.ImageList32.Images("PBI_MeasureTableTristate32.ico")
                ElseIf Me.VisibleMeasures = True Then
                    Me.nd.Image = Me.Cube.tm.ImageList32.Images("PBI_TableWithMeasuresTristate32.ico")
                Else
                    Me.nd.Image = Me.Cube.tm.ImageList32.Images("PBI_TableTristate32.ico")
                End If
            Else
                If Me.DIMENSION_IS_MEASURE_TABLE Then
                    Me.nd.Image = Me.Cube.tm.ImageList32.Images("PBI_MeasureTableSelected32.ico")
                ElseIf Me.VisibleMeasures = True Then
                    Me.nd.Image = Me.Cube.tm.ImageList32.Images("PBI_TableWithMeasuresSelected32.ico")
                Else
                    Me.nd.Image = Me.Cube.tm.ImageList32.Images("PBI_TableSelected32.ico")
                End If
            End If
        End Sub








        Private _BasicDisplayState As enDisplayState
        Private _DisplayState As enDisplayState
        Public Sub ResetDisplayState()
            Me.DisplayState = _BasicDisplayState
        End Sub
        Public Property DisplayState As enDisplayState
            Get
                Return Me._DisplayState
            End Get
            Set(value As enDisplayState)

                Me._DisplayState = value


                For Each m As Measure In Me.MeasuresVisible
                    If Me._DisplayState = enDisplayState.Blank Then
                        m.IsSelected = False
                    ElseIf Me._DisplayState = enDisplayState.Selected Then
                        m.IsSelected = True
                    End If
                Next m
                For Each l As Level In Me.LevelsAttrVisible
                    If Me._DisplayState = enDisplayState.Blank Then
                        l.IsSelected = False
                    ElseIf Me._DisplayState = enDisplayState.Selected Then
                        l.IsSelected = True
                    End If
                Next l






                If Me.DIMENSION_IS_MEASURE_TABLE Then
                    If Me._DisplayState = enDisplayState.Blank Then
                        Me.nd.Image = Me.Cube.tm.ImageList32.Images("PBI_MeasureTableUnSelected32.ico")
                        'Me.nd.Image = Me.Cube.tm.ImageList32.Images("PBI_MeasureTableBlank32.ico")
                    ElseIf Me._DisplayState = enDisplayState.Selected Then
                        Me.nd.Image = Me.Cube.tm.ImageList32.Images("PBI_MeasureTableSelected32.ico")
                    ElseIf Me._DisplayState = enDisplayState.Tristate Then
                        Me.nd.Image = Me.Cube.tm.ImageList32.Images("PBI_MeasureTableTristate32.ico")
                    End If
                ElseIf Me.VisibleMeasures = True Then
                    If Me._DisplayState = enDisplayState.Blank Then
                        Me.nd.Image = Me.Cube.tm.ImageList32.Images("PBI_TableWithMeasuresUnSelected32.ico")
                        'Me.nd.Image = Me.Cube.tm.ImageList32.Images("PBI_TableWithMeasuresBlank32.ico")
                    ElseIf Me._DisplayState = enDisplayState.Selected Then
                        Me.nd.Image = Me.Cube.tm.ImageList32.Images("PBI_TableWithMeasuresSelected32.ico")
                    ElseIf Me._DisplayState = enDisplayState.Tristate Then
                        Me.nd.Image = Me.Cube.tm.ImageList32.Images("PBI_TableWithMeasuresTristate32.ico")
                    End If
                Else
                    If Me._DisplayState = enDisplayState.Blank Then
                        Me.nd.Image = Me.Cube.tm.ImageList32.Images("PBI_TableUnSelected32.ico")
                        'Me.nd.Image = Me.Cube.tm.ImageList32.Images("PBI_TableBlank32.ico")
                    ElseIf Me._DisplayState = enDisplayState.Selected Then
                        Me.nd.Image = Me.Cube.tm.ImageList32.Images("PBI_TableSelected32.ico")
                    ElseIf Me._DisplayState = enDisplayState.Tristate Then
                        Me.nd.Image = Me.Cube.tm.ImageList32.Images("PBI_TableTristate32.ico")
                    End If
                End If

            End Set
        End Property



        Public Property Measures As List(Of Measure)

        Public ReadOnly Property SortByTypeAndName As String
            Get
                Dim strRes As String = ""
                If Me.DIMENSION_IS_MEASURE_TABLE = True Then
                    strRes = "A_" & Me.DIMENSION_CAPTION
                Else
                    strRes = "Z_" & Me.DIMENSION_CAPTION
                End If
                Return strRes
            End Get
        End Property


        Public ReadOnly Property MeasuresSortedbyName As List(Of Measure)
            Get
                Me.Measures.Sort(Function(x As Measure, y As Measure) x.MEASURE_CAPTION.CompareTo(y.MEASURE_CAPTION))
                Return Me.Measures
            End Get
        End Property

        Public ReadOnly Property MeasuresVisible As List(Of Measure)
            Get
                Dim lstRes As New List(Of Measure)
                For Each m In Me.Measures
                    If m.MEASURE_IS_VISIBLE = True Then
                        If lstRes.Contains(m) = False Then
                            lstRes.Add(m)
                        End If
                    End If
                Next m
                Return lstRes
            End Get
        End Property


        Public ReadOnly Property LevelsAttrVisible As List(Of Level)
            Get
                Dim lstRes As New List(Of Level)
                For Each h In Me.Hierarchies
                    'If h.HIERARCHY_ORIGIN <> 1 Then
                    For Each l As Level In h.Levels
                        If l.LEVEL_IS_VISIBLE = True Then
                            If l.LEVEL_TYPE <> 1 Then
                                If lstRes.Contains(l) = False Then
                                    lstRes.Add(l)
                                End If
                            End If
                        End If
                    Next l
                    'End If
                Next h
                lstRes.Sort(Function(x As Level, y As Level) x.Hierarchy.HIERARCHY_ORDINAL.CompareTo(y.Hierarchy.HIERARCHY_ORDINAL))
                Return lstRes

            End Get
        End Property


        Public Property Hierarchies As List(Of Hierarchy)
        Private Property htDisplayFolder As Hashtable

        Public ReadOnly Property DisplayFolders As List(Of DisplayFolder)
            Get
                Dim lstRes As New List(Of DisplayFolder)
                For Each k As DisplayFolder In htDisplayFolder.Values
                    lstRes.Add(k)
                Next k



                lstRes.Sort(Function(x As DisplayFolder, y As DisplayFolder)
                                Dim ComPres As Integer = x.nLevel.CompareTo(y.nLevel)
                                If ComPres = 0 Then
                                    ComPres = x.Name.CompareTo(y.Name)
                                End If
                                Return ComPres
                            End Function)
                Return lstRes
            End Get
        End Property

        Public ReadOnly Property RootDisplayFolders As List(Of DisplayFolder)
            Get
                Dim lstRes As New List(Of DisplayFolder)
                For Each k As DisplayFolder In htDisplayFolder.Values
                    If k.dfParent Is Nothing Then
                        lstRes.Add(k)
                    End If
                Next k
                Return lstRes
            End Get
        End Property

        Public ReadOnly Property VisibleMeasures As Boolean
            Get
                For Each m In Me.Measures
                    If m.MEASURE_IS_VISIBLE = True Then
                        Return True
                    End If
                Next m
                For Each df In Me.DisplayFolders
                    For Each m In df.Measures
                        If m.MEASURE_IS_VISIBLE = True Then
                            Return True
                        End If
                    Next m
                Next df
                Return False
            End Get
        End Property



        Public ReadOnly Property DIMENSION_IS_MEASURE_TABLE As Boolean
            Get
                Dim blnRes As Boolean = False
                Dim blnMea As Boolean = False
                Dim blnDim As Boolean = False

                For Each m As Measure In Me.Measures
                    If m.MEASURE_IS_VISIBLE = True Then
                        blnMea = True
                        Exit For
                    End If
                Next m
                For Each h As Hierarchy In Me.Hierarchies
                    If h.HIERARCHY_IS_VISIBLE = True Then
                        blnDim = True
                        Exit For
                    End If
                Next h

                If blnMea = True And blnDim = False Then
                    blnRes = True
                Else
                    blnRes = False
                End If

                Return blnRes
            End Get
        End Property

        Public ReadOnly Property DIMENSION_IS_MD_MEASURES As Boolean
            Get
                If Me.DIMENSION_UNIQUE_NAME.ToLower.Trim = "[measures]" Then
                    Return True
                Else
                    Return False
                End If
            End Get
        End Property

        Public Sub New(Cube As Cube)
            Me.Cube = Cube
            Me.Measures = New List(Of Measure)
            Me.Hierarchies = New List(Of Hierarchy)
            Me.htDisplayFolder = New Hashtable
        End Sub

        Public Function GetDisplayFolders(Path As String) As List(Of DisplayFolder)
            Dim strPaths() As String = Split(Path, ";")
            Dim lstRes As New List(Of DisplayFolder)
            For Each s As String In strPaths
                If s.Trim <> "" Then
                    lstRes.Add(GetDisplayFolder(s))
                End If
            Next s
            Return lstRes

        End Function


        Private Function GetDisplayFolder(Path As String) As DisplayFolder

            If Me.htDisplayFolder.ContainsKey(Path) = True Then
                Return htDisplayFolder.Item(Path)
            End If

            Dim df As New DisplayFolder(Path, Me)
            htDisplayFolder.Add(Path, df)
            If Path.Contains("\") = False Then
                Return df
            End If

            Dim _df As DisplayFolder = df
            Dim ndf As DisplayFolder
            Dim strPath As String = Path.Substring(0, Path.LastIndexOf("\"))
            Do While strPath <> ""
                If Me.htDisplayFolder.ContainsKey(strPath) = False Then
                    ndf = New DisplayFolder(strPath, Me)
                    _df.dfParent = ndf
                    ndf.DisplayFolders.Add(_df)
                    htDisplayFolder.Add(strPath, ndf)
                    _df = ndf
                Else
                    ndf = Me.htDisplayFolder.Item(strPath)
                    _df.dfParent = ndf
                    ndf.DisplayFolders.Add(_df)
                    _df = ndf
                End If
                If strPath.Contains("\") = True Then
                    strPath = strPath.Substring(0, strPath.LastIndexOf("\"))
                Else
                    strPath = ""
                End If

            Loop

            Return df
        End Function
    End Class

    Public Class DisplayFolder

        Public Dimension As clsTabularModel.Dimension
        Public Path As String
        Public Measures As List(Of Measure)
        Public Hierarchies As List(Of Hierarchy)
        Public dfParent As DisplayFolder
        Public DisplayFolders As List(Of DisplayFolder)
        Public nd As C1.Win.C1FlexGrid.Node

        Public ReadOnly Property HierarchiesSortedbyOrdinal As List(Of Hierarchy)
            Get
                Me.Hierarchies.Sort(Function(x, y) x.HIERARCHY_ORDINAL.CompareTo(y.HIERARCHY_ORDINAL))
                Return Me.Hierarchies
            End Get
        End Property

        Public ReadOnly Property IsSelected As Boolean
            Get
                For Each l As clsTabularModel.Level In Me.AllLevelsSortedByName
                    If l.IsSelected = True Then
                        Return True
                    End If
                Next l
                For Each m As clsTabularModel.Measure In Me.AllMeasuresSortedByName
                    If m.IsSelected = True Then
                        Return True
                    End If
                Next m
                Return False
            End Get
        End Property

        Public Enum enDisplayState
            Blank = 0
            Filtered = 1
            Selected = 2
            Tristate = 3
            SelectedAndFiltered = 4
            MouseOverSelected = 5
            MouseOverUnSelected = 6
            MouseOverSelectedAndFiltered = 7
        End Enum

        Public Sub SetDisplaystate()

            Dim intCtr As Integer = 0
            Dim intSel As Integer = 0
            Dim intNSel As Integer = 0
            For Each l As clsTabularModel.Level In Me.AllLevelsSortedByName
                intCtr += 1
                If l.IsSelected = True Then
                    intSel += 1
                Else
                    intNSel += 1
                End If
            Next l
            For Each m As clsTabularModel.Measure In Me.AllMeasuresSortedByName
                intCtr += 1
                If m.IsSelected = True Then
                    intSel += 1
                Else
                    intNSel += 1
                End If
            Next m

            If intSel = 0 Then
                Me.nd.Image = Me.Dimension.Cube.tm.ImageList32.Images("PBI_DisplayFolderUnSelected32.ico")
                'Me.nd.Image = Me.Dimension.Cube.tm.ImageList32.Images("PBI_DisplayFolderBlank32.ico")
                'Me.DisplayState = enDisplayState.Blank
            ElseIf intSel = intCtr Then
                Me.nd.Image = Me.Dimension.Cube.tm.ImageList32.Images("PBI_DisplayFolderSelected32.ico")
                'Me.DisplayState = enDisplayState.Tristate
            Else
                Me.nd.Image = Me.Dimension.Cube.tm.ImageList32.Images("PBI_DisplayFolderTristate32.ico")
                'Me.DisplayState = enDisplayState.Selected
            End If



        End Sub



        Private _BasicDisplayState As enDisplayState
        Private _DisplayState As enDisplayState
        Public Sub ResetDisplayState()
            Me.DisplayState = _BasicDisplayState
        End Sub
        Public Property DisplayState As enDisplayState
            Get
                Return Me._DisplayState
            End Get
            Set(value As enDisplayState)
                If value <> enDisplayState.MouseOverSelected And value <> enDisplayState.MouseOverUnSelected Then
                    _BasicDisplayState = value
                End If
                Me._DisplayState = value

                If Me._DisplayState = enDisplayState.Blank Then

                    For Each m As Measure In Me.AllMeasuresSortedByName
                        m.IsSelected = False
                    Next m
                    For Each l As Level In Me.AllLevelsSortedByName
                        l.IsSelected = False
                    Next l

                    Me.nd.Image = Me.Dimension.Cube.tm.ImageList32.Images("PBI_DisplayFolderUnSelected32.ico")
                    'Me.nd.Image = Me.Dimension.Cube.tm.ImageList32.Images("PBI_DisplayFolderBlank32.ico")
                ElseIf Me._DisplayState = enDisplayState.Selected Then

                    For Each m As Measure In Me.AllMeasuresSortedByName
                        m.IsSelected = True
                    Next m
                    For Each l As Level In Me.AllLevelsSortedByName
                        l.IsSelected = True
                    Next l

                    Me.nd.Image = Me.Dimension.Cube.tm.ImageList32.Images("PBI_DisplayFolderSelected32.ico")
                ElseIf Me._DisplayState = enDisplayState.Tristate Then
                    Me.nd.Image = Me.Dimension.Cube.tm.ImageList32.Images("PBI_DisplayFolderTristate32.ico")
                End If

            End Set
        End Property


        Public ReadOnly Property MeasuresSortedbyName As List(Of Measure)
            Get
                Me.Measures.Sort(Function(x As Measure, y As Measure) x.MEASURE_CAPTION.CompareTo(y.MEASURE_CAPTION))
                Return Me.Measures
            End Get
        End Property

        Public ReadOnly Property AllMeasuresSortedByName As List(Of Measure)
            Get
                Me._lstMeas = New List(Of Measure)
                Me._GetAllMeasures(Me)
                Me._lstMeas.Sort(Function(x As Measure, y As Measure) x.MEASURE_CAPTION.CompareTo(y.MEASURE_CAPTION))
                Return Me._lstMeas
            End Get
        End Property

        Dim _lstMeas As List(Of Measure)
        Private Sub _GetAllMeasures(df As DisplayFolder)
            For Each m As Measure In df.Measures
                If Me._lstMeas.Contains(m) = False AndAlso m.MEASURE_IS_VISIBLE = True Then
                    Me._lstMeas.Add(m)
                End If
            Next m
            For Each dfc As DisplayFolder In df.DisplayFolders
                For Each m As clsTabularModel.Measure In dfc.Measures
                    If Me._lstMeas.Contains(m) = False AndAlso m.MEASURE_IS_VISIBLE = True Then
                        Me._lstMeas.Add(m)
                    End If
                Next m
                Me._GetAllMeasures(dfc)
            Next dfc
        End Sub

        Public ReadOnly Property AllLevelsSortedByName As List(Of Level)
            Get
                Me._lstLvl = New List(Of Level)
                Me._GetAllLevels(Me)
                Me._lstLvl.Sort(Function(x As Level, y As Level) (x.Hierarchy.Dimension.DIMENSION_CAPTION & ":" & x.LEVEL_CAPTION).CompareTo(y.Hierarchy.Dimension.DIMENSION_CAPTION & ":" & y.LEVEL_CAPTION))
                Return Me._lstLvl
            End Get
        End Property

        Dim _lstLvl As List(Of Level)
        Private Sub _GetAllLevels(df As DisplayFolder)
            For Each h As Hierarchy In df.Hierarchies
                For Each l As Level In h.LevelsAttrVisible
                    If Me._lstLvl.Contains(l) = False Then
                        Me._lstLvl.Add(l)
                    End If
                Next l
            Next h
            For Each dfc As DisplayFolder In df.DisplayFolders
                For Each h As clsTabularModel.Hierarchy In dfc.Hierarchies
                    For Each l As clsTabularModel.Level In h.LevelsAttrVisible
                        If Me._lstLvl.Contains(l) = False Then
                            Me._lstLvl.Add(l)
                        End If
                    Next l
                Next h
                Me._GetAllLevels(dfc)
            Next dfc
        End Sub

        Public ReadOnly Property Name As String
            Get
                Dim strRes As String = ""
                If Me.dfParent Is Nothing Then
                    strRes = Me.Path
                Else
                    strRes = Me.Path.Substring(Me.Path.LastIndexOf("\") + 1)
                End If
                Return strRes
            End Get
        End Property

        Public ReadOnly Property nLevel As Integer
            Get
                If Me.dfParent Is Nothing Then
                    Return 0
                Else
                    Return Split(Me.Path, "\").Length - 1
                End If
            End Get
        End Property


        Public Sub New(Path As String, Dimension As clsTabularModel.Dimension)
            Me.Dimension = Dimension
            Me.DisplayFolders = New List(Of DisplayFolder)
            Me.Measures = New List(Of Measure)
            Me.Hierarchies = New List(Of Hierarchy)
            Me.Path = Path
            If Me.Path = "" Then Me.Path = "XXX"
        End Sub


    End Class




    Public Class Hierarchy

        Public Property Dimension As Dimension
        Public Property CUBE_NAME As String
        Public Property DIMENSION_UNIQUE_NAME As String
        Public Property HIERARCHY_NAME As String
        Public Property HIERARCHY_UNIQUE_NAME As String
        Public Property HIERARCHY_CAPTION As String
        Public Property DIMENSION_TYPE As Integer
        Public Property HIERARCHY_CARDINALITY As Integer
        Public Property DEFAULT_MEMBER As String
        Public Property ALL_MEMBER As String
        Public Property DESCRIPTION As String
        Public Property DIMENSION_IS_VISIBLE As Boolean
        Public Property HIERARCHY_ORDINAL As Integer
        Public Property HIERARCHY_IS_VISIBLE As Boolean
        Public Property HIERARCHY_ORIGIN As Integer
        Public Property HIERARCHY_DISPLAY_FOLDER As String

        Public Property Levels As List(Of Level)

        Public Property DisplayFolder As DisplayFolder
        Public nd As C1.Win.C1FlexGrid.Node

        Public Enum enDisplayState
            Blank = 0
            Filtered = 1
            Selected = 2
            Tristate = 3
            SelectedAndFiltered = 4
            MouseOverSelected = 5
            MouseOverUnSelected = 6
            MouseOverSelectedAndFiltered = 7
        End Enum

        Public Sub SetDisplayState()

            Dim intCtr As Integer = 0
            Dim intSel As Integer = 0
            Dim intNSel As Integer = 0
            For Each l As clsTabularModel.Level In Me.LevelsAttrVisible
                intCtr += 1
                If l.IsSelected = True Or l.SiblingIsSelected = True Then
                    intSel += 1
                Else
                    intNSel += 1
                End If
            Next l

            If intSel = 0 Then
                'Me.nd.Image = Me.Dimension.Cube.tm.ImageList32.Images("PBI_HierarchyBlank32.ico")
                Me.nd.Image = Me.Dimension.Cube.tm.ImageList32.Images("PBI_HierarchyUnSelected32.ico")
            ElseIf intSel <> intCtr Then
                Me.nd.Image = Me.Dimension.Cube.tm.ImageList32.Images("PBI_HierarchyTristate32.ico")
            Else
                Me.nd.Image = Me.Dimension.Cube.tm.ImageList32.Images("PBI_HierarchySelected32.ico")
            End If
        End Sub

        Private _BasicDisplayState As enDisplayState
        Private _DisplayState As enDisplayState


        Public Property DisplayState As enDisplayState
            Get
                Return Me._DisplayState
            End Get
            Set(value As enDisplayState)

                Me._DisplayState = value

                If Me._DisplayState = enDisplayState.Blank Then
                    For Each l As Level In Me.LevelsAttrVisible
                        l.IsSelected = False
                    Next l
                    Me.nd.Image = Me.Dimension.Cube.tm.ImageList32.Images("PBI_HierarchyUnSelected32.ico")
                    'Me.nd.Image = Me.Dimension.Cube.tm.ImageList32.Images("PBI_HierarchyBlank32.ico")
                ElseIf Me._DisplayState = enDisplayState.Selected Then
                    For Each l As Level In Me.LevelsAttrVisible
                        l.IsSelected = True
                    Next l
                    Me.nd.Image = Me.Dimension.Cube.tm.ImageList32.Images("PBI_HierarchySelected32.ico")
                ElseIf Me._DisplayState = enDisplayState.Tristate Then
                    Me.nd.Image = Me.Dimension.Cube.tm.ImageList32.Images("PBI_HierarchyTristate32.ico")
                End If

            End Set
        End Property

        Public ReadOnly Property IsSelected As Boolean
            Get
                For Each l As clsTabularModel.Level In Me.LevelsAttrVisible
                    If l.IsSelected = True Then
                        Return True
                    End If
                Next l
                Return False
            End Get
        End Property



        Public ReadOnly Property LevelsAttrVisible As List(Of Level)
            Get
                Dim lstRes As New List(Of Level)
                For Each l As Level In Me.Levels
                    If l.LEVEL_IS_VISIBLE = True Then
                        If l.LEVEL_TYPE <> 1 Then
                            lstRes.Add(l)
                        End If
                    End If
                Next l
                Return lstRes
            End Get
        End Property

        Public Sub New(Dimension As Dimension)
            Me.Dimension = Dimension
            Me.Levels = New List(Of Level)
        End Sub




    End Class

    Public Class Level

        Public Property Hierarchy As Hierarchy

        Public Property CUBE_NAME As String
        Public Property DIMENSION_UNIQUE_NAME As String
        Public Property HIERARCHY_UNIQUE_NAME As String
        Public Property LEVEL_NAME As String
        Public Property LEVEL_UNIQUE_NAME As String
        Public Property LEVEL_CAPTION As String
        Public Property LEVEL_NUMBER As Integer
        Public Property LEVEL_TYPE As Integer
        Public Property DESCRIPTION As String
        Public Property LEVEL_IS_VISIBLE As Boolean
        Public Property LEVEL_ORDERING_PROPERTY As String
        Public Property LEVEL_DBTYPE As Integer
        Public Property LEVEL_ATTRIBUTE_HIERARCHY_NAME As String
        Public Property LEVEL_KEY_CARDINALITY As Integer
        Public Property LEVEL_ORIGIN As Integer

        Public Property FilterControlGUID As String

        'Public ReadOnly Property Connection As clsConnection
        '    Get
        '        Return Me.Hierarchy.Dimension.Cube.tm.Connection
        '    End Get
        'End Property

        Public ReadOnly Property UniName As String
            Get
                Return "'" & Me.Hierarchy.Dimension.DIMENSION_NAME & "'[" & Me.LEVEL_ATTRIBUTE_HIERARCHY_NAME & "]"
            End Get
        End Property

        Public ReadOnly Property TableName As String
            Get
                Return Me.Hierarchy.Dimension.DIMENSION_NAME
            End Get
        End Property

        Public ReadOnly Property FieldName As String
            Get
                Return Me.LEVEL_ATTRIBUTE_HIERARCHY_NAME
            End Get
        End Property

        Public ReadOnly Property Caption As String
            Get
                Return Me.LEVEL_CAPTION
            End Get
        End Property

        Public ReadOnly Property FormatString As String
            Get
                Return ""
            End Get
        End Property

        Public Property DaxFilter As String

        Public Property htSel As Hashtable

        Public Property SearchTerm As String

        Private _SelectionMode As enSelectionMode
        Public Property SelectionMode As clsTabularModel.enSelectionMode
            Get
                If Me._SelectionMode = 0 Then
                    Me._SelectionMode = enSelectionMode.AllSelected
                End If
                Return Me._SelectionMode
            End Get
            Set(value As clsTabularModel.enSelectionMode)
                Me._SelectionMode = value
            End Set
        End Property

        Public Property Sort As enSort

        Private _IsSelected As Boolean
        Public Property IsSelected As Boolean
            Get
                Return Me._IsSelected
            End Get
            Set(value As Boolean)
                Me._IsSelected = value
                If Me._IsSelected = True Then
                    Me.DisplayState = enDisplayState.Selected
                Else
                    Me.DisplayState = enDisplayState.Blank
                End If
            End Set
        End Property

        Public ReadOnly Property SiblingIsSelected As Boolean
            Get
                For Each l As Level In Me.LevelSiblings
                    If l.IsSelected = True Then
                        Return True
                    End If
                Next l
                Return False
            End Get
        End Property



        Public ReadOnly Property IsFiltered As Boolean
            Get
                Dim blnRes As Boolean = False
                If Me.SelectionMode = enSelectionMode.AllSelected Then
                    blnRes = False
                ElseIf Me.SelectionMode = enSelectionMode.AllSearch And Me.DaxFilter <> "" Then
                    blnRes = True
                ElseIf Me.SelectionMode = enSelectionMode.SelectMember Or Me.SelectionMode = enSelectionMode.DeSelectMember Then
                    If Me.htSel Is Nothing OrElse Me.htSel.Count = 0 Then
                        blnRes = False
                    Else
                        blnRes = True
                    End If
                End If
                Return blnRes
            End Get
        End Property


        Public ReadOnly Property FieldType As clsTabularModel.enFieldType
            Get
                Return clsTabularModel.enFieldType.Level
            End Get
        End Property

        Public ReadOnly Property DataType As enDataType
            Get
                Dim enRes As enDataType
                Select Case Me.LEVEL_DBTYPE
                    Case 0 : enRes = enDataType.Text
                    Case 2 To 6 : enRes = enDataType.Number
                    Case 7 : enRes = enDataType.DateTime
                    Case 8 : enRes = enDataType.Text
                    Case 9 To 10 : enRes = enDataType.Text
                    Case 11 : enRes = enDataType.Bool
                    Case 12 To 13 : enRes = enDataType.Text
                    Case 14 To 21 : enRes = enDataType.Number
                    Case 72 : enRes = enDataType.Text
                    Case 128 : enRes = enDataType.Text
                    Case 129 To 130 : enRes = enDataType.Text
                    Case 131 : enRes = enDataType.Number
                    Case 132 : enRes = enDataType.Text
                    Case 133 To 135 : enRes = enDataType.DateTime
                    Case 136 : enRes = enDataType.Text
                End Select
                Return enRes
            End Get
        End Property

        'Public Property FilterControl As ctrlFilter

        Private _LevelSiblings As List(Of Level)
        Public ReadOnly Property LevelSiblings As List(Of Level)
            Get
                If Me._LevelSiblings Is Nothing Then
                    Me._LevelSiblings = New List(Of Level)
                    For Each h As Hierarchy In Me.Hierarchy.Dimension.Hierarchies
                        If Not h Is Me.Hierarchy Then
                            For Each l As Level In h.Levels
                                If Not l Is Me Then
                                    If l.LEVEL_ATTRIBUTE_HIERARCHY_NAME = Me.LEVEL_ATTRIBUTE_HIERARCHY_NAME Or 1 = 2 Then
                                        Me._LevelSiblings.Add(l)
                                    End If
                                End If
                            Next l
                        End If
                    Next h
                End If
                Return Me._LevelSiblings
            End Get
        End Property


        Public nds As List(Of C1.Win.C1FlexGrid.Node)

        Public Enum enDisplayState
            Blank = 0
            Filtered = 1
            Selected = 2
            Tristate = 3
            SelectedAndFiltered = 4
            MouseOverSelected = 5
            MouseOverUnSelected = 6
            MouseOverSelectedAndFiltered = 7
        End Enum

        Private _BasicDisplayState As enDisplayState
        Private _DisplayState As enDisplayState
        Public Sub ResetDisplayState()
            Me.DisplayState = _BasicDisplayState
        End Sub

        Public Sub SetDisplayState()
            If Me.IsFiltered = True And Me.IsSelected = True Then
                Me.DisplayState = enDisplayState.SelectedAndFiltered
            ElseIf Me.IsFiltered = False And Me.IsSelected = True Then
                Me.DisplayState = enDisplayState.Selected
            ElseIf Me.IsFiltered = True And Me.IsSelected = False Then
                Me.DisplayState = enDisplayState.Filtered
            Else
                Me.DisplayState = enDisplayState.Blank
            End If
        End Sub

        Public Property DisplayState As enDisplayState
            Get
                Return _DisplayState
            End Get
            Set(value As enDisplayState)

                Me._DisplayState = value

                If Me.Hierarchy.HIERARCHY_ORIGIN = 6 Then
                    For Each nd As C1.Win.C1FlexGrid.Node In Me.nds
                        If Me._DisplayState = enDisplayState.Blank Then
                            nd.Image = Me.Hierarchy.Dimension.Cube.tm.ImageList32.Images("PBI_KeyColumnUnSelected32.ico")
                        ElseIf Me._DisplayState = enDisplayState.Selected Then
                            nd.Image = Me.Hierarchy.Dimension.Cube.tm.ImageList32.Images("PBI_KeyColumnSelected32.ico")
                        ElseIf Me._DisplayState = enDisplayState.SelectedAndFiltered Then
                            nd.Image = Me.Hierarchy.Dimension.Cube.tm.ImageList32.Images("PBI_KeyColumnSelAndFilt32.ico")
                        End If
                    Next nd
                Else
                    For Each nd As C1.Win.C1FlexGrid.Node In Me.nds
                        If Me._DisplayState = enDisplayState.Blank Then
                            nd.Image = Me.Hierarchy.Dimension.Cube.tm.ImageList32.Images("PBI_ColumnUnSelected32.ico")
                        ElseIf Me._DisplayState = enDisplayState.Selected Then
                            nd.Image = Me.Hierarchy.Dimension.Cube.tm.ImageList32.Images("PBI_ColumnSelected32.ico")
                        ElseIf Me._DisplayState = enDisplayState.SelectedAndFiltered Then
                            nd.Image = Me.Hierarchy.Dimension.Cube.tm.ImageList32.Images("PBI_ColumnSelAndFilt32.ico")
                        End If
                    Next nd
                End If


                For Each l As Level In Me.LevelSiblings
                    If l.DisplayState <> Me._DisplayState Then
                        l.DisplayState = Me._DisplayState
                    End If
                Next l




            End Set
        End Property

        Public Sub New(Hierarchy As Hierarchy)
            Me.Hierarchy = Hierarchy
            Me.nds = New List(Of C1.Win.C1FlexGrid.Node)
        End Sub

    End Class

    Public Class Members

        Public Property Level As clsTabularModel.Level
        Public Property Members As List(Of Object)

        Public Sub New(Level As clsTabularModel.Level, Members As List(Of Object))
            Me.Level = Level
            Me.Members = Members

        End Sub
    End Class


    Public Class Measure

        Public Property Dimension As Dimension
        Public Property CUBE_NAME As String
        Public Property MEASURE_NAME As String
        Public Property MEASURE_UNIQUE_NAME As String
        Public Property MEASURE_CAPTION As String
        Public Property MEASURE_AGGREGATOR As Integer
        Public Property DATA_TYPE As Integer
        Public Property NUMERIC_PRECISION As Integer
        Public Property NUMERIC_SCALE As Integer
        Public Property EXPRESSION As String
        Public Property MEASURE_IS_VISIBLE As Boolean
        Public Property MEASUREGROUP_NAME As String
        Public Property MEASURE_DISPLAY_FOLDER As String
        Public Property DEFAULT_FORMAT_STRING As String
        Public Property DisplayFolder As DisplayFolder
        Public nds As List(Of C1.Win.C1FlexGrid.Node)

        Public Property FilterControlGUID As String

        'Public ReadOnly Property Connection As clsConnection
        '    Get
        '        Return Me.Dimension.Cube.tm.Connection
        '    End Get
        'End Property

        Public ReadOnly Property UniName As String
            Get
                Return "'" & Me.Dimension.DIMENSION_NAME & "'[" & Me.MEASURE_NAME & "]"
            End Get
        End Property

        Public ReadOnly Property TableName As String
            Get
                Return Me.Dimension.DIMENSION_NAME
            End Get
        End Property

        Public ReadOnly Property FieldName As String
            Get
                Return Me.MEASURE_NAME
            End Get
        End Property

        Public ReadOnly Property Caption As String
            Get
                Return Me.MEASURE_CAPTION
            End Get
        End Property

        Public ReadOnly Property FormatString As String
            Get
                Return Me.DEFAULT_FORMAT_STRING
            End Get
        End Property


        Public Property DaxFilter As String
        Public Property htSel As Hashtable
        Public Property SearchTerm As String
        Private _SelectionMode As enSelectionMode
        Public Property SelectionMode As clsTabularModel.enSelectionMode
            Get
                If Me._SelectionMode = 0 Then
                    Me._SelectionMode = enSelectionMode.AllSelected
                End If
                Return Me._SelectionMode
            End Get
            Set(value As clsTabularModel.enSelectionMode)
                Me._SelectionMode = value
            End Set
        End Property

        Public Property Sort As enSort

        Private _IsSelected As Boolean
        Public Property IsSelected As Boolean
            Get
                Return _IsSelected
            End Get
            Set(value As Boolean)
                _IsSelected = value
                If _IsSelected = True Then
                    Me.DisplayState = enDisplayState.Selected
                Else
                    Me.DisplayState = enDisplayState.Blank
                End If
            End Set
        End Property




        Public ReadOnly Property IsFiltered As Boolean
            Get
                Dim blnRes As Boolean = False
                If Me.SelectionMode = enSelectionMode.AllSelected Then
                    blnRes = False
                ElseIf Me.SelectionMode = enSelectionMode.AllSearch And Me.DaxFilter <> "" Then
                    blnRes = True
                ElseIf Me.SelectionMode = enSelectionMode.SelectMember Or Me.SelectionMode = enSelectionMode.DeSelectMember Then
                    If Me.htSel Is Nothing OrElse Me.htSel.Count = 0 Then
                        blnRes = False
                    Else
                        blnRes = True
                    End If
                End If
                Return blnRes
            End Get
        End Property

        Public ReadOnly Property FieldType As clsTabularModel.enFieldType
            Get
                Return clsTabularModel.enFieldType.Measure
            End Get
        End Property

        Public ReadOnly Property DataType As clsTabularModel.enDataType
            Get
                Dim enRes As clsTabularModel.enDataType
                Select Case Me.DATA_TYPE
                    Case 0 : enRes = clsTabularModel.enDataType.Text
                    Case 2 To 6 : enRes = clsTabularModel.enDataType.Number
                    Case 7 : enRes = clsTabularModel.enDataType.DateTime
                    Case 8 : enRes = clsTabularModel.enDataType.Text
                    Case 9 To 10 : enRes = clsTabularModel.enDataType.Text
                    Case 11 : enRes = clsTabularModel.enDataType.Bool
                    Case 12 To 13 : enRes = clsTabularModel.enDataType.Text
                    Case 14 To 21 : enRes = clsTabularModel.enDataType.Number
                    Case 72 : enRes = clsTabularModel.enDataType.Text
                    Case 128 : enRes = clsTabularModel.enDataType.Text
                    Case 129 To 130 : enRes = clsTabularModel.enDataType.Text
                    Case 131 : enRes = clsTabularModel.enDataType.Number
                    Case 132 : enRes = clsTabularModel.enDataType.Text
                    Case 133 To 135 : enRes = clsTabularModel.enDataType.DateTime
                    Case 136 : enRes = clsTabularModel.enDataType.Text
                End Select
                Return enRes
            End Get
        End Property

        'Public Property FilterControl As ctrlFilter

        Public Enum enDisplayState
            Blank = 0
            Filtered = 1
            Selected = 2
            Tristate = 3
            SelectedAndFiltered = 4
            MouseOverSelected = 5
            MouseOverUnSelected = 6
            MouseOverSelectedAndFiltered = 7
        End Enum

        Private _BasicDisplayState As enDisplayState
        Private _DisplayState As enDisplayState
        Public Sub ResetDisplayState()
            Me.DisplayState = _BasicDisplayState
        End Sub

        Public Sub SetDisplayState()
            If Me.IsFiltered = True And Me.IsSelected = True Then
                Me.DisplayState = enDisplayState.SelectedAndFiltered
            ElseIf Me.IsFiltered = False And Me.IsSelected = True Then
                Me.DisplayState = enDisplayState.Selected
            ElseIf Me.IsFiltered = True And Me.IsSelected = False Then
                Me.DisplayState = enDisplayState.Filtered
            Else
                Me.DisplayState = enDisplayState.Blank
            End If
        End Sub

        Public Property DisplayState As enDisplayState
            Get
                Return _DisplayState
            End Get
            Set(value As enDisplayState)
                If value <> enDisplayState.MouseOverSelected And value <> enDisplayState.MouseOverUnSelected And value <> enDisplayState.MouseOverSelectedAndFiltered Then
                    _BasicDisplayState = value
                End If
                Me._DisplayState = value
                For Each nd As C1.Win.C1FlexGrid.Node In Me.nds
                    If Me._DisplayState = enDisplayState.Blank Then
                        nd.Image = Me.Dimension.Cube.tm.ImageList32.Images("PBI_MeasureUnSelected32.ico")
                    ElseIf Me._DisplayState = enDisplayState.Selected Then
                        nd.Image = Me.Dimension.Cube.tm.ImageList32.Images("PBI_MeasureChecked32.ico")
                    ElseIf Me._DisplayState = enDisplayState.SelectedAndFiltered Then
                        nd.Image = Me.Dimension.Cube.tm.ImageList32.Images("PBI_MeasureSelAndFilt32.ico")
                    End If
                Next nd
            End Set




        End Property

        Public Sub New(Dimension As Dimension)
            Me.Dimension = Dimension
            Me.nds = New List(Of C1.Win.C1FlexGrid.Node)
        End Sub

    End Class



#Region "Dataset"



    Private Sub GetDataSet()


        If Not Me.ds Is Nothing Then
            Me.ds.Dispose() : Me.ds = Nothing
        End If


        Dim dsRes As New DataSet
        Dim rec As Object = CreateObject("ADODB.RECORDSET")
        rec.open("Select * from $SYSTEM.MDSCHEMA_CUBES WHERE CUBE_NAME='" & Me.CubeName & "'", Me.conn, 0)
        Dim dt As DataTable = getDatatable(rec)
        dt.TableName = "CUBES"
        rec.close
        dsRes.Tables.Add(dt.Copy)
        dt.Dispose()

        rec.open("select * from $SYSTEM.MDSCHEMA_DIMENSIONS WHERE CUBE_NAME='" & Me.CubeName & "'", Me.conn, 0)
        dt = getDatatable(rec)
        dt.TableName = "DIMENSIONS"
        rec.close
        dsRes.Tables.Add(dt.Copy)
        dt.Dispose()

        rec.open("select * from $SYSTEM.MDSCHEMA_HIERARCHIES WHERE CUBE_NAME='" & Me.CubeName & "'", Me.conn, 0)
        dt = getDatatable(rec)
        dt.TableName = "HIERARCHIES"
        rec.close
        dsRes.Tables.Add(dt.Copy)
        dt.Dispose()

        rec.open("select * from $SYSTEM.MDSCHEMA_LEVELS WHERE CUBE_NAME='" & Me.CubeName & "'", Me.conn, 0)
        dt = getDatatable(rec)
        dt.TableName = "LEVELS"
        rec.close
        dsRes.Tables.Add(dt.Copy)
        dt.Dispose()

        rec.open("select * from $SYSTEM.MDSCHEMA_MEASURES WHERE CUBE_NAME='" & Me.CubeName & "'", Me.conn, 0)
        dt = getDatatable(rec)
        dt.TableName = "MEASURES"
        rec.close
        dsRes.Tables.Add(dt.Copy)
        dt.Dispose()

        rec.open("select * from $SYSTEM.MDSCHEMA_MEASUREGROUPS WHERE CUBE_NAME='" & Me.CubeName & "'", Me.conn, 0)
        dt = getDatatable(rec)
        dt.TableName = "MEASUREGROUPS"
        rec.close
        dsRes.Tables.Add(dt.Copy)
        dt.Dispose()

        rec.open("select * from $SYSTEM.MDSCHEMA_MEASUREGROUP_DIMENSIONS WHERE CUBE_NAME='" & Me.CubeName & "'", Me.conn, 0)
        dt = getDatatable(rec)
        dt.TableName = "MEASUREGROUP_DIMENSIONS"
        rec.close
        dsRes.Tables.Add(dt.Copy)
        dt.Dispose()

        rec.open("Select * From $SYSTEM.DBSCHEMA_CATALOGS", Me.conn, 0)
        dt = getDatatable(rec)
        dt.TableName = "DBSCHEMA_CATALOGS"
        rec.close
        dsRes.Tables.Add(dt.Copy)
        Dim strDB As String = dt.Rows(0)("CATALOG_NAME").ToString
        dt.Dispose()

        'rec.open("SELECT * FROM SYSTEMRESTRICTSCHEMA ($SYSTEM.DISCOVER_CSDL_METADATA, [CATALOG_NAME] = '" & strDB & "',[VERSION] = '4.0') ", Me.conn, 0)
        'This query can result in an error
        Try
            rec.open("SELECT * FROM SYSTEMRESTRICTSCHEMA ($SYSTEM.DISCOVER_CSDL_METADATA, [CATALOG_NAME] = '" & strDB & "') ", Me.conn, 0)
            dt = getDatatable(rec)
            dt.TableName = "CSDL_METADATA"
            rec.close
            dsRes.Tables.Add(dt.Copy)
            dt.Dispose()
        Catch ex As Exception
        End Try





        Dim pCols As DataColumn() = New DataColumn() {dsRes.Tables("CUBES").Columns("CUBE_NAME")}
        Dim cCols As DataColumn() = New DataColumn() {dsRes.Tables("DIMENSIONS").Columns("CUBE_NAME")}
        Dim r As New DataRelation("Cube_to_Hier", pCols, cCols, False)
        dsRes.Relations.Add(r)

        Dim pCols1 As DataColumn() = New DataColumn() {dsRes.Tables("DIMENSIONS").Columns("CUBE_NAME"), dsRes.Tables("DIMENSIONS").Columns("DIMENSION_UNIQUE_NAME")}
        Dim cCols1 As DataColumn() = New DataColumn() {dsRes.Tables("HIERARCHIES").Columns("CUBE_NAME"), dsRes.Tables("HIERARCHIES").Columns("DIMENSION_UNIQUE_NAME")}
        Dim r1 As New DataRelation("Dim_to_Hier", pCols1, cCols1, False)
        dsRes.Relations.Add(r1)

        Dim pCols2 = New DataColumn() {dsRes.Tables("HIERARCHIES").Columns("CUBE_NAME"), dsRes.Tables("HIERARCHIES").Columns("HIERARCHY_UNIQUE_NAME")}
        Dim cCols2 = New DataColumn() {dsRes.Tables("LEVELS").Columns("CUBE_NAME"), dsRes.Tables("LEVELS").Columns("HIERARCHY_UNIQUE_NAME")}
        Dim r2 As New DataRelation("Hier_to_Lev", pCols2, cCols2, False)
        dsRes.Relations.Add(r2)

        Dim pCols3 = New DataColumn() {dsRes.Tables("DIMENSIONS").Columns("CUBE_NAME"), dsRes.Tables("DIMENSIONS").Columns("DIMENSION_NAME")}
        Dim cCols3 = New DataColumn() {dsRes.Tables("MEASURES").Columns("CUBE_NAME"), dsRes.Tables("MEASURES").Columns("MEASUREGROUP_NAME")}
        Dim r3 As New DataRelation("Dim_to_Meas", pCols3, cCols3, False)
        dsRes.Relations.Add(r3)


        dsRes.Tables("HIERARCHIES").DefaultView.RowFilter = "HIERARCHY_DISPLAY_FOLDER<>''"
        For Each xr As DataRowView In dsRes.Tables("HIERARCHIES").DefaultView
            xr("HIERARCHY_DISPLAY_FOLDER") = _DisplayFolderPath(xr("HIERARCHY_DISPLAY_FOLDER").ToString)
        Next xr
        dsRes.Tables("HIERARCHIES").DefaultView.RowFilter = ""

        dsRes.Tables("MEASURES").DefaultView.RowFilter = "MEASURE_DISPLAY_FOLDER<>''"
        For Each xr As DataRowView In dsRes.Tables("MEASURES").DefaultView
            xr("MEASURE_DISPLAY_FOLDER") = _DisplayFolderPath(xr("MEASURE_DISPLAY_FOLDER").ToString)
        Next xr
        dsRes.Tables("MEASURES").DefaultView.RowFilter = ""

        Me.ds = dsRes.Copy
        dsRes.Dispose()


    End Sub


    Private Shared Function getDatatable(rec As Object) As DataTable
        Dim adap As New System.Data.OleDb.OleDbDataAdapter
        Dim dt As New DataTable
        adap.Fill(dt, rec)
        Return dt
    End Function

    Private Function _DisplayFolderPath(path As String) As String
        If path.Contains("\") = False Then
            Return path
        End If
        If path.StartsWith("\") Then path = path.Substring(1)
        If path.EndsWith("\") Then path = path.Substring(0, path.Length - 1)
        If path = "" Then path = " "
        Do While path.Contains("\\")
            path = path.Replace("\\", "\")
        Loop
        Return path
    End Function




    Public Function GetFormat(qc As clsQueryColumn) As String

        If Not Me.ds Is Nothing Then
            Try
                Dim xmlString As String = Me.ds.Tables("CSDL_METADATA")(0)(0)

                xmlString = xmlString.Substring(xmlString.IndexOf("<EntityType Name=""" & qc.EntityName & """>"))
                xmlString = xmlString.Substring(0, xmlString.IndexOf("</EntityType>"))
                If qc.FieldType = clsQueryColumn.enFieldType.Level Then
                    If qc.DataType = clsQueryColumn.enDataType.Bool Then
                        Return ""
                    End If
                    xmlString = xmlString.Substring(xmlString.IndexOf("<Property Name=""" & qc.ReferenceName & ""))
                    xmlString = xmlString.Substring(0, xmlString.IndexOf("</Property>"))
                    If xmlString.Contains("FormatString=") Then
                        xmlString = xmlString.Substring(xmlString.IndexOf("FormatString=") + 14)
                        xmlString = xmlString.Substring(0, xmlString.IndexOf(""""))
                        Return xmlString
                    End If
                ElseIf qc.FieldType = clsQueryColumn.enFieldType.Measure Then
                    If qc.DataType = clsQueryColumn.enDataType.Bool Then
                        Return ""
                    End If
                    xmlString = xmlString.Substring(xmlString.IndexOf("ReferenceName=""" & qc.ReferenceName & ""))
                    xmlString = xmlString.Substring(0, xmlString.IndexOf("</Property>"))
                    If xmlString.Contains("FormatString=") Then
                        xmlString = xmlString.Substring(xmlString.IndexOf("FormatString=") + 14)
                        xmlString = xmlString.Substring(0, xmlString.IndexOf(""""))
                        Return xmlString
                    End If
                End If
            Catch ex As Exception
                Return ""
            End Try
        End If

        Return ""

    End Function

#End Region

End Class
