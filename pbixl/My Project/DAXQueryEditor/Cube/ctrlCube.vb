Public Class ctrlCube

    Private Property pnlHeader As Panel
    Private Property lblHeader As Label
    Private Property pnlMain As Panel
    Private Property pnlException As Panel
    Private Property pnlLoading As Panel

    Private Property ctrlDaxQuery As ctrlDaxQuery
    Private Property pb As PictureBox
    Friend Property Err As Exception

    Friend Property tm As clsTabularModel

    Public Enum enctrlState
        loading = 1
        ready = 2
        exception = 3
    End Enum

    Private _ctrlState As enctrlState
    Public Property ctrlState As enctrlState
        Get
            Return _ctrlState
        End Get
        Set(value As enctrlState)
            Me._ctrlState = value
            If Me._ctrlState = enctrlState.ready Then
                Me.InvokeIfRequired(Sub()
                                        Me.pnlMain.Visible = True
                                        Me.pnlLoading.Visible = False
                                        Me.pnlException.Visible = False
                                        pb.Image = Me.ctrlDaxQuery.ImageList.Images("PowerBI_OK.ico")
                                    End Sub)
            ElseIf Me._ctrlState = enctrlState.loading Then
                Me.InvokeIfRequired(Sub()
                                        Me.pnlMain.Visible = False
                                        Me.pnlLoading.Visible = True
                                        Me.pnlException.Visible = False
                                        pb.Image = CType(My.Resources.Wheel, System.Drawing.Image)
                                    End Sub)
            Else
                Me.InvokeIfRequired(Sub()
                                        Me.ctrlDaxQuery.ctrlTable.ctrlState = ctrlTable.enctrlState.init
                                        Me.pnlMain.Visible = False
                                        Me.pnlLoading.Visible = False
                                        Me.pnlException.Visible = True
                                        Me.pnlException.Controls(0).Text = Me.Err.Message
                                        pb.Image = Me.ctrlDaxQuery.ImageList.Images("PowerBI_NotOK.ico")
                                    End Sub)
            End If



        End Set
    End Property


    Private _fg As C1.Win.C1FlexGrid.C1FlexGrid
    Public ReadOnly Property fg As C1.Win.C1FlexGrid.C1FlexGrid
        Get
            If Not Me._fg Is Nothing Then
                Return Me._fg
            End If

            Me._fg = New C1.Win.C1FlexGrid.C1FlexGrid
            With Me._fg

                .BeginUpdate()
                .Redraw = False
                .BorderStyle = C1.Win.C1FlexGrid.Util.BaseControls.BorderStyleEnum.None
                .Cols.Count = 1
                .Rows.Count = 0
                .Tree.Column = 0
                .Tree.Indent = 1
                .Rows.Fixed = 0
                .Cols.Fixed = 0
                .ExtendLastCol = True
                .Styles.Normal.Border.Style = C1.Win.C1FlexGrid.BorderStyleEnum.None
                .Styles.Normal.BackColor = Drawing.Color.WhiteSmoke
                .Styles.Normal.Border.Color = Drawing.Color.WhiteSmoke
                .Styles.EmptyArea.Border.Color = Drawing.Color.WhiteSmoke
                .Styles.EmptyArea.BackColor = Drawing.Color.WhiteSmoke
                .AllowEditing = False
                .Cols(0).TextAlign = C1.Win.C1FlexGrid.TextAlignEnum.LeftCenter
                .HighLight = C1.Win.C1FlexGrid.HighLightEnum.Always
                .Styles.Highlight.BackColor = Drawing.Color.LightGray
                .Styles.Highlight.ForeColor = Drawing.Color.Black
                .SelectionMode = C1.Win.C1FlexGrid.SelectionModeEnum.Row
                .FocusRect = C1.Win.C1FlexGrid.FocusRectEnum.None



                .Tree.NodeImageCollapsed = Me.ctrlDaxQuery.ImageList.Images("expand.ico")
                .Tree.NodeImageExpanded = Me.ctrlDaxQuery.ImageList.Images("down.ico")
                .Tree.Style = C1.Win.C1FlexGrid.TreeStyleFlags.Symbols

                .Anchor = 15

                .Redraw = True
                .EndUpdate()
            End With


            Return Me._fg
        End Get
    End Property



    Public Sub New(ctrlDaxQuery As ctrlDaxQuery)

        InitializeComponent()

        Me.ctrlDaxQuery = ctrlDaxQuery

        Me.BorderStyle = BorderStyle.FixedSingle

        Me.pnlHeader = New Panel
        With Me.pnlHeader
            .Location = New Point(0, 0)
            .Height = 20
            .Width = Me.Width
            .Anchor = 13
            .BackColor = Color.LightGray
        End With

        Me.pb = New PictureBox With {.Top = 0, .Left = 0, .Height = 20, .Width = 20, .SizeMode = PictureBoxSizeMode.Zoom}
        Me.pnlHeader.Controls.Add(pb)


        Me.lblHeader = New Label With {.Top = 3, .Left = 24, .Width = Me.pnlHeader.Width - 30, .Anchor = 13, .AutoEllipsis = True, .Text = "Conn Name"}
        Me.pnlHeader.Controls.Add(Me.lblHeader)

        Me.pnlMain = New Panel
        With Me.pnlMain
            .Location = New Point(0, Me.pnlHeader.Height)
            .Height = Me.Height - 20
            .Width = Me.Width
            .Anchor = 15
        End With

        Me.pnlLoading = New Panel
        With Me.pnlLoading
            .Location = New Point(0, Me.pnlHeader.Height)
            .Height = Me.Height - 20
            .Width = Me.Width
            .Anchor = 15
        End With
        Dim lblLoading As New Label With {.Top = 10, .Left = 10, .Text = "loading..", .Anchor = 5}
        Me.pnlLoading.Controls.Add(lblLoading)

        Me.pnlException = New Panel
        With Me.pnlException
            .Location = New Point(0, Me.pnlHeader.Height)
            .Height = Me.Height - 20
            .Width = Me.Width
            .Anchor = 15
        End With
        Dim lblException As New Label With {.Top = 10, .Left = 10, .Text = "", .Anchor = 5, .AutoSize = True, .MaximumSize = New Size(100, 0)}
        Me.pnlException.Controls.Add(lblException)


        Me.Controls.Add(Me.pnlHeader)
        Me.Controls.Add(Me.pnlLoading)
        Me.Controls.Add(Me.pnlException)
        Me.Controls.Add(Me.pnlMain)


        With Me.fg
            .Top = 4
            .Left = 4
            .Width = Me.pnlMain.Width - 8
            .Height = Me.pnlMain.Height - 8
        End With
        Me.pnlMain.Controls.Add(Me.fg)
        AddHandler Me.fg.MouseClick, AddressOf fg_MouseClick

        Me.lblHeader.Text = Me.ctrlDaxQuery.ConnName




    End Sub

    Private Sub fg_MouseClick(sender As Object, e As System.Windows.Forms.MouseEventArgs)

        Dim hti As C1.Win.C1FlexGrid.HitTestInfo = Me.fg.HitTest(e.X, e.Y)
        If hti.Row <= 0 OrElse hti.Row + 1 > Me.fg.Rows.Count Then
            Exit Sub
        End If

        Dim nd As C1.Win.C1FlexGrid.Node = Me.fg.Rows(hti.Row).Node
        If nd Is Nothing Then Exit Sub

        If hti.X - 4 > (nd.Level + 1) * Me.fg.Tree.Indent And hti.X - 4 < ((nd.Level + 1) * Me.fg.Tree.Indent) + 14 Then

            If Not TryCast(Me.fg.Rows(nd.GetCellRange.r1).UserData, clsTabularModel.Measure) Is Nothing Then
                If TryCast(Me.fg.Rows(nd.GetCellRange.r1).UserData, clsTabularModel.Measure).IsSelected Then
                    TryCast(Me.fg.Rows(nd.GetCellRange.r1).UserData, clsTabularModel.Measure).IsSelected = False
                Else
                    TryCast(Me.fg.Rows(nd.GetCellRange.r1).UserData, clsTabularModel.Measure).IsSelected = True
                End If
            ElseIf Not TryCast(Me.fg.Rows(nd.GetCellRange.r1).UserData, clsTabularModel.Level) Is Nothing Then
                If TryCast(Me.fg.Rows(nd.GetCellRange.r1).UserData, clsTabularModel.Level).IsSelected Then
                    TryCast(Me.fg.Rows(nd.GetCellRange.r1).UserData, clsTabularModel.Level).IsSelected = False
                    If TryCast(Me.fg.Rows(nd.GetCellRange.r1).UserData, clsTabularModel.Level).LevelSiblings.Count > 0 Then
                        For Each l As clsTabularModel.Level In TryCast(Me.fg.Rows(nd.GetCellRange.r1).UserData, clsTabularModel.Level).LevelSiblings
                            l.IsSelected = False
                        Next l
                    End If
                Else
                    TryCast(Me.fg.Rows(nd.GetCellRange.r1).UserData, clsTabularModel.Level).IsSelected = True
                    If TryCast(Me.fg.Rows(nd.GetCellRange.r1).UserData, clsTabularModel.Level).LevelSiblings.Count > 0 Then
                        For Each l As clsTabularModel.Level In TryCast(Me.fg.Rows(nd.GetCellRange.r1).UserData, clsTabularModel.Level).LevelSiblings
                            l.IsSelected = True
                        Next l
                    End If
                End If
            ElseIf Not TryCast(Me.fg.Rows(nd.GetCellRange.r1).UserData, clsTabularModel.Hierarchy) Is Nothing Then

                If TryCast(Me.fg.Rows(nd.GetCellRange.r1).UserData, clsTabularModel.Hierarchy).IsSelected = True Then
                    TryCast(Me.fg.Rows(nd.GetCellRange.r1).UserData, clsTabularModel.Hierarchy).DisplayState = clsTabularModel.Hierarchy.enDisplayState.Selected
                    For Each l As clsTabularModel.Level In TryCast(Me.fg.Rows(nd.GetCellRange.r1).UserData, clsTabularModel.Hierarchy).Levels
                        If l.IsSelected = True Then
                            l.IsSelected = False
                            If l.LevelSiblings.Count > 0 Then
                                For Each ls In l.LevelSiblings
                                    ls.IsSelected = False
                                Next ls
                            End If
                        End If
                    Next l
                Else
                    TryCast(Me.fg.Rows(nd.GetCellRange.r1).UserData, clsTabularModel.Hierarchy).DisplayState = clsTabularModel.Hierarchy.enDisplayState.Blank
                    For Each l As clsTabularModel.Level In TryCast(Me.fg.Rows(nd.GetCellRange.r1).UserData, clsTabularModel.Hierarchy).Levels
                        If l.IsSelected = False Then
                            l.IsSelected = True
                            If l.LevelSiblings.Count > 0 Then
                                For Each ls In l.LevelSiblings
                                    ls.IsSelected = True
                                Next ls
                            End If
                        End If
                    Next l
                End If
            ElseIf Not TryCast(Me.fg.Rows(nd.GetCellRange.r1).UserData, clsTabularModel.DisplayFolder) Is Nothing Then

                Dim blnSel As Boolean = False
                Dim lstF As List(Of Object) = Me.tm.GetAllObjectFields(TryCast(Me.fg.Rows(nd.GetCellRange.r1).UserData, clsTabularModel.DisplayFolder))
                For Each f In lstF
                    If f.IsSelected = True Then
                        blnSel = True
                        Exit For
                    End If
                Next f

                If blnSel = False Then
                    TryCast(Me.fg.Rows(nd.GetCellRange.r1).UserData, clsTabularModel.DisplayFolder).DisplayState = clsTabularModel.DisplayFolder.enDisplayState.Selected
                Else
                    TryCast(Me.fg.Rows(nd.GetCellRange.r1).UserData, clsTabularModel.DisplayFolder).DisplayState = clsTabularModel.DisplayFolder.enDisplayState.Blank
                End If
            ElseIf Not TryCast(Me.fg.Rows(nd.GetCellRange.r1).UserData, clsTabularModel.Dimension) Is Nothing Then

                Dim blnSel As Boolean = False
                Dim lstF As List(Of Object) = Me.tm.GetAllObjectFields(TryCast(Me.fg.Rows(nd.GetCellRange.r1).UserData, clsTabularModel.Dimension))
                For Each f In lstF
                    If f.IsSelected = True Then
                        blnSel = True
                        Exit For
                    End If
                Next f

                If blnSel = False Then
                    TryCast(Me.fg.Rows(nd.GetCellRange.r1).UserData, clsTabularModel.Dimension).DisplayState = clsTabularModel.Dimension.enDisplayState.Selected
                Else
                    TryCast(Me.fg.Rows(nd.GetCellRange.r1).UserData, clsTabularModel.Dimension).DisplayState = clsTabularModel.Dimension.enDisplayState.Blank
                End If
            End If


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
                ndS = Me.fg.Rows(Me.fg.Rows.Count - 1).Node
            End If
            For i = ndS.GetCellRange.r1 To ndP.GetCellRange.r1 Step -1
                If Not TryCast(Me.fg.Rows(i).UserData, clsTabularModel.DisplayFolder) Is Nothing Then
                    TryCast(Me.fg.Rows(i).UserData, clsTabularModel.DisplayFolder).SetDisplaystate()
                ElseIf Not TryCast(Me.fg.Rows(i).UserData, clsTabularModel.Hierarchy) Is Nothing Then
                    TryCast(Me.fg.Rows(i).UserData, clsTabularModel.Hierarchy).SetDisplayState()
                ElseIf Not TryCast(Me.fg.Rows(i).UserData, clsTabularModel.Dimension) Is Nothing Then
                    TryCast(Me.fg.Rows(i).UserData, clsTabularModel.Dimension).SetDisplayState()
                End If
            Next i



            Me.ctrlDaxQuery.RefreshPreview()



        End If








    End Sub



    Friend Sub ShowNavigation()


        If Me.Disposing = True Or Me.IsDisposed Then
            Exit Sub
        End If



        Me.tm = Me.ctrlDaxQuery.tm

        Me.InvokeIfRequired(Sub()

                                With Me.fg


                                    .Redraw = False
                                    .BeginUpdate()

                                    Dim cnd As C1.Win.C1FlexGrid.Node = Nothing
                                    Dim dnd As C1.Win.C1FlexGrid.Node = Nothing
                                    Dim dfd As C1.Win.C1FlexGrid.Node = Nothing
                                    Dim mnd As C1.Win.C1FlexGrid.Node = Nothing
                                    Dim hnd As C1.Win.C1FlexGrid.Node = Nothing
                                    Dim l_uh_nd As C1.Win.C1FlexGrid.Node = Nothing
                                    Dim l_ah_nd As C1.Win.C1FlexGrid.Node = Nothing

                                    Try
                                        .Rows.Count = 1
                                    Catch ex As Exception
                                        Exit Sub
                                    End Try


                                    .Rows.Count = 1
                                    .Rows(0).IsNode = True
                                    .Rows(0).Node.Level = 0
                                    cnd = .Rows(0).Node
                                    .Rows(0).UserData = Me.ctrlDaxQuery.CubeName
                                    .SetCellImage(cnd.GetCellRange.r1, 0, Me.ctrlDaxQuery.ImageList.Images("cube.ico"))
                                    .SetData(cnd.GetCellRange.r1, 0, Me.ctrlDaxQuery.CubeName)


                                    For Each c As clsTabularModel.Cube In tm.Cubes

                                        For Each d As clsTabularModel.Dimension In c.DimensionsSortedByTypeAndName

                                            If (d.DIMENSION_IS_VISIBLE = True Or d.DIMENSION_IS_MEASURE_TABLE = True) And d.DIMENSION_IS_MD_MEASURES = False Then

                                                dnd = cnd.AddNode(C1.Win.C1FlexGrid.NodeTypeEnum.LastChild, d.DIMENSION_CAPTION)
                                                .Rows(dnd.GetCellRange.r1).UserData = d
                                                d.nd = dnd

                                                d.DisplayState = clsTabularModel.Dimension.enDisplayState.Blank

#Region "DisplayFolder"




                                                For Each df As clsTabularModel.DisplayFolder In d.DisplayFolders
                                                    If df.dfParent Is Nothing Then

                                                        dfd = dnd.AddNode(C1.Win.C1FlexGrid.NodeTypeEnum.LastChild, df.Name)
                                                        .Rows(dfd.GetCellRange.r1).UserData = df
                                                        df.nd = dfd
                                                        df.DisplayState = clsTabularModel.DisplayFolder.enDisplayState.Blank
                                                    Else

                                                        dfd = df.dfParent.nd.AddNode(C1.Win.C1FlexGrid.NodeTypeEnum.LastChild, df.Name)
                                                        .Rows(dfd.GetCellRange.r1).UserData = df
                                                        df.nd = dfd
                                                        df.DisplayState = clsTabularModel.DisplayFolder.enDisplayState.Blank
                                                    End If

                                                    'Measures
                                                    For Each m As clsTabularModel.Measure In df.MeasuresSortedbyName
                                                        If m.MEASURE_IS_VISIBLE = True Then
                                                            mnd = df.nd.AddNode(C1.Win.C1FlexGrid.NodeTypeEnum.LastChild, m.MEASURE_CAPTION)
                                                            .Rows(mnd.GetCellRange.r1).UserData = m
                                                            m.nds.Add(mnd)
                                                            m.DisplayState = clsTabularModel.Measure.enDisplayState.Blank
                                                        End If
                                                    Next m

                                                    'User Hierarchies
                                                    For Each h As clsTabularModel.Hierarchy In df.Hierarchies
                                                        If h.HIERARCHY_IS_VISIBLE = True And h.HIERARCHY_ORIGIN = 1 Then

                                                            hnd = df.nd.AddNode(C1.Win.C1FlexGrid.NodeTypeEnum.LastChild, h.HIERARCHY_CAPTION)
                                                            .Rows(hnd.GetCellRange.r1).UserData = h
                                                            h.nd = hnd
                                                            h.DisplayState = clsTabularModel.Hierarchy.enDisplayState.Blank
                                                            For Each l As clsTabularModel.Level In h.Levels
                                                                If l.LEVEL_TYPE <> 1 Then
                                                                    l_uh_nd = hnd.AddNode(C1.Win.C1FlexGrid.NodeTypeEnum.LastChild, l.LEVEL_CAPTION)
                                                                    .Rows(l_uh_nd.GetCellRange.r1).UserData = l
                                                                    l.nds.Add(l_uh_nd)
                                                                    l.DisplayState = clsTabularModel.Level.enDisplayState.Blank
                                                                End If
                                                            Next l
                                                        End If
                                                    Next h

                                                    'Attribute Hierarchies
                                                    For Each h As clsTabularModel.Hierarchy In df.HierarchiesSortedbyOrdinal
                                                        If h.HIERARCHY_ORIGIN <> 1 Then
                                                            For Each l As clsTabularModel.Level In h.Levels
                                                                If l.LEVEL_TYPE <> 1 And l.LEVEL_IS_VISIBLE = True Then
                                                                    l_ah_nd = df.nd.AddNode(C1.Win.C1FlexGrid.NodeTypeEnum.LastChild, l.LEVEL_CAPTION)
                                                                    .Rows(l_ah_nd.GetCellRange.r1).UserData = l
                                                                    l.nds.Add(l_ah_nd)
                                                                    l.DisplayState = clsTabularModel.Level.enDisplayState.Blank
                                                                End If
                                                            Next l
                                                        End If
                                                    Next h

                                                Next df

#End Region

                                                'Measures
                                                For Each m As clsTabularModel.Measure In d.MeasuresSortedbyName
                                                    If m.MEASURE_IS_VISIBLE = True And m.DisplayFolder Is Nothing Then
                                                        mnd = dnd.AddNode(C1.Win.C1FlexGrid.NodeTypeEnum.LastChild, m.MEASURE_CAPTION)
                                                        .Rows(mnd.GetCellRange.r1).UserData = m
                                                        m.nds.Add(mnd)
                                                        m.DisplayState = clsTabularModel.Measure.enDisplayState.Blank
                                                    End If
                                                Next m


                                                'User Hierarchies
                                                For Each h As clsTabularModel.Hierarchy In d.Hierarchies
                                                    If h.HIERARCHY_IS_VISIBLE = True And h.HIERARCHY_ORIGIN = 1 Then

                                                        hnd = dnd.AddNode(C1.Win.C1FlexGrid.NodeTypeEnum.LastChild, h.HIERARCHY_CAPTION)
                                                        .Rows(hnd.GetCellRange.r1).UserData = h
                                                        h.nd = hnd
                                                        h.DisplayState = clsTabularModel.Hierarchy.enDisplayState.Blank

                                                        For Each l As clsTabularModel.Level In h.Levels
                                                            If l.LEVEL_TYPE <> 1 Then
                                                                l_uh_nd = hnd.AddNode(C1.Win.C1FlexGrid.NodeTypeEnum.LastChild, l.LEVEL_CAPTION)
                                                                .Rows(l_uh_nd.GetCellRange.r1).UserData = l
                                                                l.nds.Add(l_uh_nd)
                                                                l.DisplayState = clsTabularModel.Level.enDisplayState.Blank
                                                            End If
                                                        Next l
                                                    End If
                                                Next h

                                                'Attribute Hierarchies
                                                For Each h As clsTabularModel.Hierarchy In d.HierarchiesSortedbyOrdinal
                                                    If h.HIERARCHY_IS_VISIBLE = True And h.HIERARCHY_ORIGIN <> 1 And h.HIERARCHY_DISPLAY_FOLDER = "" Then
                                                        For Each l As clsTabularModel.Level In h.Levels
                                                            If l.LEVEL_TYPE <> 1 And l.LEVEL_IS_VISIBLE = True Then
                                                                l_ah_nd = dnd.AddNode(C1.Win.C1FlexGrid.NodeTypeEnum.LastChild, l.LEVEL_CAPTION)
                                                                .Rows(l_ah_nd.GetCellRange.r1).UserData = l
                                                                l.nds.Add(l_ah_nd)
                                                                l.DisplayState = clsTabularModel.Level.enDisplayState.Blank
                                                            End If
                                                        Next l
                                                    End If
                                                Next h





                                            End If

                                        Next d
                                    Next c

                                    For i As Integer = .Rows.Count - 1 To 0 Step -1
                                        If .Rows(i).Node.Level >= 1 Then
                                            .Rows(i).Node.Expanded = False
                                        End If
                                    Next i




                                    .EndUpdate()
                                    .Redraw = True



                                End With



                                Me.ctrlState = enctrlState.ready
                                End Sub)



    End Sub




End Class
