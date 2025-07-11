'Imports Microsoft.Office.Interop


Public Class ctrlFilter


    Enum enFieldType
        Level = 1
        Measure = 2
        iFunction = 3
    End Enum


    Enum enSort
        none = 0
        asc = 1
        desc = 2
    End Enum

    Enum enControlState
        Executing = 1
        Cancelled = 2
        IsError = 3
        NoResult = 4
        Preview = 5
        ImpMeasure = 6
    End Enum


    Private _Sort As enSort
    Public Property Sort As enSort
        Get
            If Me._Sort = Nothing Then
                Me._Sort = enSort.none
            End If
            Return Me._Sort
        End Get
        Set(value As enSort)
            Me._Sort = value
        End Set
    End Property

    Private ReadOnly Property IsImplicitCast As Boolean
        Get
            Return Me.QueryColumn.isImplicitCast
        End Get
    End Property

    Private ReadOnly Property RunAsync As Boolean
        Get
            Dim blnAsync As Boolean = False
            'If Not Me.PBIMeasure Is Nothing Then
            '    If Me.PBIMeasure.Dimension.Cube.tm.Connection.ConnType <> clsConnection.enConnType.WorkbookModel Then
            '        blnAsync = True
            '    End If
            'End If
            'If Not Me.PBILevel Is Nothing Then
            '    If Me.PBILevel.Hierarchy.Dimension.Cube.tm.Connection.ConnType <> clsConnection.enConnType.WorkbookModel Then
            '        blnAsync = True
            '    End If
            'End If
            Return True
            'Return blnAsync

        End Get
    End Property

    Private Enum enSelectionMode
        AllSelected = 1
        AllSearch = 2
        DeSelectMember = 3
        SelectMember = 4
    End Enum

    Private _SelectionMode As enSelectionMode
    Private Property SelectionMode As enSelectionMode
        Get
            If Me._SelectionMode <= 0 Then
                Return enSelectionMode.AllSelected
            Else
                Return Me._SelectionMode
            End If
        End Get
        Set(value As enSelectionMode)
            Me._SelectionMode = value
        End Set
    End Property



    Private ctsSource As System.Threading.CancellationTokenSource
    Private cts As System.Threading.CancellationToken

    Private _ControlState As enControlState
    Private Property ControlState As enControlState
        Get
            Return Me._ControlState
        End Get
        Set(value As enControlState)
            Me._ControlState = value

            If Not Me.objRec Is Nothing Then
                If Me.objRec.state = 1 Then
                    Me.objRec.close
                End If
                Me.objRec = Nothing
            End If


            If Me.IsDisposed OrElse Me.Disposing Then
                Return
            End If


            If value = enControlState.Executing Then
                Me.InvokeIfRequired(Sub()
                                        Me.lblApply.Enabled = False
                                        Me.txtBox.Enabled = False
                                        Me.fg.Visible = False
                                        Me.txtErr.Visible = False
                                        Me.pnlCanc.Visible = False
                                        Me.pnlExec.Visible = True
                                        Me.pnlExec.Refresh()
                                    End Sub)

            ElseIf value = enControlState.Cancelled Then
                Me.ctsSource.Cancel()
                Me.InvokeIfRequired(Sub()
                                        Me.lblApply.Enabled = True
                                        Me.txtBox.Enabled = True
                                        Me.fg.Visible = False
                                        Me.txtErr.Visible = False
                                        Me.pnlCanc.Visible = True
                                        Me.pnlExec.Visible = False
                                    End Sub)

            ElseIf value = enControlState.IsError Then
                Me.InvokeIfRequired(Sub()
                                        Me.lblApply.Enabled = True
                                        Me.txtBox.Enabled = True
                                        Me.fg.Visible = False
                                        Me.pnlCanc.Visible = False
                                        Me.pnlExec.Visible = False
                                        If Me.objErr Is Nothing Then
                                            Me.txtErr.Text = "An error occured."
                                        Else
                                            Me.txtErr.Text = Me.objErr.Message
                                        End If
                                        Me.txtErr.Visible = True
                                    End Sub)

            ElseIf value = enControlState.NoResult Then
                Me.InvokeIfRequired(Sub()
                                        Me.lblApply.Enabled = True
                                        Me.txtBox.Enabled = True
                                        Me.fg.Visible = False
                                        Me.pnlCanc.Visible = False
                                        Me.pnlExec.Visible = False
                                        Me.txtErr.Text = "No results."
                                        Me.txtErr.Visible = True
                                    End Sub)

            ElseIf value = enControlState.ImpMeasure Then
                Me.InvokeIfRequired(Sub()
                                        Me.lblApply.Enabled = True
                                        Me.txtBox.Enabled = True
                                        Me.fg.Visible = False
                                        Me.pnlCanc.Visible = False
                                        Me.pnlExec.Visible = False
                                        Me.txtErr.Visible = False
                                    End Sub)

            ElseIf value = enControlState.Preview Then
                Me.InvokeIfRequired(Sub()
                                        Me.lblApply.Enabled = True
                                        Me.txtBox.Enabled = True
                                        Me.fg.Visible = True
                                        Me.pnlCanc.Visible = False
                                        Me.pnlExec.Visible = False
                                        Me.txtErr.Visible = False
                                    End Sub)

            End If


        End Set
    End Property



    Friend Property DaxFilter As String
    Private Property txtErr As System.Windows.Forms.TextBox
    Private Property pnlExec As System.Windows.Forms.Panel
    Private Property lblCancel As System.Windows.Forms.LinkLabel
    Private Property pnlCanc As System.Windows.Forms.Panel

    Private Property lblApply As System.Windows.Forms.Label
    Private Property lblClear As System.Windows.Forms.Label
    Friend Property txtBox As System.Windows.Forms.TextBox

    Private Property pnlAZ As System.Windows.Forms.Panel
    Private Property btnAZ As System.Windows.Forms.Button
    Private Property lblAZ As System.Windows.Forms.Label

    Private Property pnlZA As System.Windows.Forms.Panel
    Private Property btnZA As System.Windows.Forms.Button
    Private Property lblZA As System.Windows.Forms.Label

    Private Property pnl1 As System.Windows.Forms.Panel
    Private Property pnl2 As System.Windows.Forms.Panel
    Private Property pnl3 As System.Windows.Forms.Panel

    Private Property pnlRe As System.Windows.Forms.Panel
    Private Property lblRe As System.Windows.Forms.Label

    Private Property pnlCl As System.Windows.Forms.Panel
    Private Property btnCl As System.Windows.Forms.Button
    Private Property lblCl As System.Windows.Forms.Label

    Private Property btnTy As System.Windows.Forms.Button
    Private Property btnInfo As System.Windows.Forms.Button

    Private ReadOnly Property Table As String
        Get
            If Me.FieldType = enFieldType.Level Then
                Return Me.PBILevel.Hierarchy.Dimension.DIMENSION_NAME
            Else
                Return Me.PBIMeasure.Dimension.DIMENSION_NAME
            End If
        End Get
    End Property

    Private ReadOnly Property Column As String
        Get
            If Me.FieldType = enFieldType.Level Then
                Return Me.PBILevel.LEVEL_NAME
            Else
                Return Me.PBIMeasure.MEASURE_NAME
            End If
        End Get
    End Property

    Private ReadOnly Property FieldType As enFieldType
        Get
            If Not Me.PBILevel Is Nothing Then
                Return enFieldType.Level
            Else
                Return enFieldType.Measure
            End If
        End Get
    End Property

    Enum enDataType
        Text = 1
        Number = 2
        DateTime = 3
        Bool = 4
    End Enum

    Private ReadOnly Property DataType As enDataType
        Get

            If Me.FieldType = enFieldType.Level Then
                Return Me.PBILevel.DataType
            Else
                Return Me.PBIMeasure.DataType
            End If

        End Get
    End Property

    Private Property _fg As C1.Win.C1FlexGrid.C1FlexGrid
    Private ReadOnly Property fg As C1.Win.C1FlexGrid.C1FlexGrid
        Get
            If Not Me._fg Is Nothing Then
                Return Me._fg
            End If

            Me._fg = New C1.Win.C1FlexGrid.C1FlexGrid
            With Me._fg
                .Visible = False
                .BeginUpdate()
                .Redraw = False
                .BorderStyle = C1.Win.C1FlexGrid.Util.BaseControls.BorderStyleEnum.FixedSingle
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

                .Cols(0).TextAlign = C1.Win.C1FlexGrid.TextAlignEnum.LeftCenter
                .HighLight = C1.Win.C1FlexGrid.HighLightEnum.Always
                .Styles.Highlight.BackColor = Drawing.Color.LightGray
                .Styles.Highlight.ForeColor = Drawing.Color.Black
                .SelectionMode = C1.Win.C1FlexGrid.SelectionModeEnum.Row
                .FocusRect = C1.Win.C1FlexGrid.FocusRectEnum.None



                .Redraw = True
                .EndUpdate()
            End With
            Return Me._fg
        End Get
    End Property

    Private Property PBILevel As clsTabularModel.Level
    Private Property PBIMeasure As clsTabularModel.Measure

    Private ReadOnly Property ConnectionString As String
        Get
            'If Not Me.PBILevel Is Nothing Then
            '    Return Me.PBILevel.Hierarchy.Dimension.Cube.tm.Connection.ConnString
            'ElseIf Not Me.PBIMeasure Is Nothing Then
            '    Return Me.PBIMeasure.Dimension.Cube.tm.Connection.ConnString
            'End If
            Return ""

        End Get
    End Property

    Public ReadOnly Property FieldUniName As String
        Get
            If Not Me.PBILevel Is Nothing Then
                Return Me.PBILevel.UniName
            ElseIf Not Me.PBIMeasure Is Nothing Then
                Return Me.PBIMeasure.UniName
            End If
            Return ""
        End Get
    End Property

    Public ReadOnly Property Caption As String
        Get
            If Not Me.PBILevel Is Nothing Then
                Return Me.PBILevel.LEVEL_CAPTION
            Else
                Return Me.PBIMeasure.MEASURE_CAPTION
            End If
            Return ""
        End Get
    End Property



    Private Property PBIField As Object
    Public Property GUID As String
    Private Property ctrlColumnHeader As ctrlColumnHeader

    Public Property QueryColumn As clsQueryColumn

    Public Sub New(PBIField As Object, qc As clsQueryColumn, ctrlColHdr As ctrlColumnHeader)
        InitializeComponent()

        Me.GUID = System.Guid.NewGuid.ToString
        Me.QueryColumn = qc

        Dim c As New List(Of Context)
        For Each _qc In Me.QueryColumn.Query.QueryColumns
            If Not _qc Is Me.QueryColumn Then
                If _qc.DaxFilter <> "" Then
                    Dim cn As New Context
                    With cn
                        .TableName = _qc.TableName
                        .FieldName = _qc.FieldName
                        .DataType = _qc.DataType
                        .FieldType = _qc.FieldType
                        .FilterExpression = _qc.DaxFilter
                    End With
                    c.Add(cn)
                End If
            End If
        Next _qc
        If c.Count > 0 Then
            'No context
            'Me.ContextExt = c
        End If

        Me.Width = Me.Width + 50


        Me.PBIField = PBIField
        Me.ctrlColumnHeader = ctrlColHdr

        If Not TryCast(Me.PBIField, clsTabularModel.Level) Is Nothing Then
            Me.PBILevel = Me.PBIField
        Else
            Me.PBIMeasure = Me.PBIField
        End If


        '        Private Property pnlAZ As System.Windows.Forms.Panel
        '   Private Property btnAZ As System.Windows.Forms.Button
        '  Private Property lblAZ As System.Windows.Forms.Label

        Me.pnlAZ = New System.Windows.Forms.Panel With {.Left = 0, .Width = Me.Width, .Top = 0, .Height = 24, .Anchor = 13, .BorderStyle = BorderStyle.None}
        Me.btnAZ = New System.Windows.Forms.Button With {.Left = 4, .Width = 20, .Height = 24, .Top = 0, .Anchor = 5, .FlatStyle = FlatStyle.Flat, .ImageAlign = ContentAlignment.MiddleLeft}
        Me.lblAZ = New System.Windows.Forms.Label With {.Left = 24, .Width = Me.Width - 30, .Height = 24, .Top = 0, .Anchor = 5, .FlatStyle = FlatStyle.Flat, .Text = "  Sort Smallest to Largest", .TextAlign = ContentAlignment.MiddleLeft}
        With btnAZ
            .FlatAppearance.BorderSize = 0
            .FlatAppearance.MouseOverBackColor = Color.Transparent
            .Image = Me.ctrlColumnHeader.ctrlTable.ctrlDaxQuery.ImageList.Images("Ascending.ico")
        End With
        Me.pnlAZ.Controls.Add(btnAZ)
        Me.pnlAZ.Controls.Add(lblAZ)
        Me.Controls.Add(pnlAZ)
        AddHandler Me.pnlAZ.Click, AddressOf AZ_Click
        AddHandler Me.btnAZ.Click, AddressOf AZ_Click
        AddHandler Me.lblAZ.Click, AddressOf AZ_Click




        Me.pnlZA = New System.Windows.Forms.Panel With {.Left = 0, .Width = Me.Width, .Top = 25, .Height = 24, .Anchor = 13, .BorderStyle = BorderStyle.None}
        Me.btnZA = New System.Windows.Forms.Button With {.Left = 4, .Width = 20, .Height = 24, .Top = 0, .Anchor = 5, .FlatStyle = FlatStyle.Flat, .ImageAlign = ContentAlignment.MiddleLeft}
        Me.lblZA = New System.Windows.Forms.Label With {.Left = 24, .Width = Me.Width - 24, .Height = 24, .Top = 0, .Anchor = 5, .FlatStyle = FlatStyle.Flat, .Text = "  Sort Largest to Smallest", .TextAlign = ContentAlignment.MiddleLeft}
        With btnZA
            .FlatAppearance.BorderSize = 0
            .FlatAppearance.MouseOverBackColor = Color.Transparent
            .Image = Me.ctrlColumnHeader.ctrlTable.ctrlDaxQuery.ImageList.Images("DescendingSort.ico")
        End With
        Me.pnlZA.Controls.Add(btnZA)
        Me.pnlZA.Controls.Add(lblZA)
        Me.Controls.Add(pnlZA)
        AddHandler Me.pnlZA.Click, AddressOf ZA_Click
        AddHandler Me.btnZA.Click, AddressOf ZA_Click
        AddHandler Me.lblZA.Click, AddressOf ZA_Click

        Me.pnl1 = New System.Windows.Forms.Panel With {.Left = 34, .Width = Me.Width - 40, .Top = 50, .Height = 1, .Anchor = 13, .BorderStyle = BorderStyle.None, .BackColor = Color.LightGray}
        Me.Controls.Add(pnl1)

        Me.pnlRe = New System.Windows.Forms.Panel With {.Left = 0, .Width = Me.Width, .Top = 50, .Height = 24, .Anchor = 13, .BorderStyle = BorderStyle.None}
        Me.lblRe = New System.Windows.Forms.Label With {.Left = 24, .Width = Me.Width - 24, .Height = 24, .Top = 0, .Anchor = 5, .FlatStyle = FlatStyle.Flat, .Text = "  Remove Column", .TextAlign = ContentAlignment.MiddleLeft}
        Me.pnlRe.Controls.Add(lblRe)
        Me.Controls.Add(pnlRe)
        AddHandler Me.pnlRe.Click, AddressOf RE_Click
        AddHandler Me.lblRe.Click, AddressOf RE_Click

        Me.pnl2 = New System.Windows.Forms.Panel With {.Left = 34, .Width = Me.Width - 40, .Top = 75, .Height = 1, .Anchor = 13, .BorderStyle = BorderStyle.None, .BackColor = Color.LightGray}
        Me.Controls.Add(pnl2)

        Me.pnlCl = New System.Windows.Forms.Panel With {.Left = 0, .Width = Me.Width, .Top = 75, .Height = 24, .Anchor = 13, .BorderStyle = BorderStyle.None}
        Me.btnCl = New System.Windows.Forms.Button With {.Left = 4, .Width = 20, .Height = 24, .Top = 0, .Anchor = 5, .FlatStyle = FlatStyle.Flat, .ImageAlign = ContentAlignment.MiddleLeft}
        Me.lblCl = New System.Windows.Forms.Label With {.Left = 24, .Width = Me.Width - 30, .Height = 24, .Top = 0, .Anchor = 5, .FlatStyle = FlatStyle.Flat, .Text = "  Clear Filter from """ & Me.PBIField.Caption & """", .TextAlign = ContentAlignment.MiddleLeft}
        With btnCl
            .FlatAppearance.BorderSize = 0
            .FlatAppearance.MouseOverBackColor = Color.Transparent
            .Image = Me.ctrlColumnHeader.ctrlTable.ctrlDaxQuery.ImageList.Images("DeleteFilter.ico")
        End With
        Me.pnlCl.Controls.Add(btnCl)
        Me.pnlCl.Controls.Add(lblCl)
        Me.Controls.Add(pnlCl)
        AddHandler Me.pnlCl.Click, AddressOf CL_Click
        AddHandler Me.btnCl.Click, AddressOf CL_Click
        AddHandler Me.lblCl.Click, AddressOf CL_Click


        Me.pnl3 = New System.Windows.Forms.Panel With {.Left = 34, .Width = Me.Width - 40, .Top = 100, .Height = 1, .Anchor = 13, .BorderStyle = BorderStyle.None, .BackColor = Color.LightGray}
        Me.Controls.Add(pnl3)


        Me.txtBox = New System.Windows.Forms.TextBox With {
        .Left = 34, .Top = 104, .Height = 18, .Width = Me.Width - 40, .Anchor = 13, .BorderStyle = Windows.Forms.BorderStyle.FixedSingle, .TextAlign = HorizontalAlignment.Left
        }
        AddHandler Me.txtBox.KeyPress, AddressOf Me.txtBox_KeyPress
        Me.Controls.Add(Me.txtBox)

        With Me.fg
            .Left = 34
            .Top = Me.txtBox.Top + Me.txtBox.Height + 5
            .Height = Me.Height - .Top - 20 - 10 + 4 - 10 + 10
            .Width = Me.Width - 40
            .Anchor = 15
            .Visible = True
        End With
        AddHandler Me.fg.CellChecked, AddressOf fg_CellChecked
        Me.Controls.Add(Me.fg)



        Me.lblClear = New System.Windows.Forms.Label With {
            .Left = Me.Width - 60 - 10 + 4, .Top = Me.fg.Top + Me.fg.Height + 2, .Anchor = 6, .Text = "Cancel", .Height = 20, .Width = 60, .BorderStyle = BorderStyle.FixedSingle, .TextAlign = ContentAlignment.MiddleCenter
        }
        AddHandler Me.lblClear.Click, AddressOf Me.lblClear_Click
        Me.Controls.Add(Me.lblClear)


        Me.lblApply = New System.Windows.Forms.Label With {
            .Left = Me.fg.Left, .Top = Me.fg.Top + Me.fg.Height + 2, .Anchor = 10, .Text = "Apply", .Height = 20, .Width = 60, .BorderStyle = BorderStyle.FixedSingle, .TextAlign = ContentAlignment.MiddleCenter
        }
        AddHandler Me.lblApply.Click, AddressOf Me.lblApply_Click
        Me.Controls.Add(Me.lblApply)

        Dim pnlExec As New System.Windows.Forms.Panel With {
        .Width = Me.fg.Width, .Height = Me.fg.Height, .Left = Me.fg.Left, .Top = Me.fg.Top, .Anchor = 15,
        .BackColor = Drawing.Color.WhiteSmoke, .Visible = True
        }
        Dim lblCancel As New System.Windows.Forms.LinkLabel With {.Text = "loading..", .Top = 5, .Left = 5}
        AddHandler lblCancel.Click, AddressOf Me.lblCancel_Click
        Me.lblCancel = lblCancel
        pnlExec.Controls.Add(lblCancel)
        Me.pnlExec = pnlExec
        Me.Controls.Add(pnlExec)

        Dim pnlCanc As New System.Windows.Forms.Panel With {
            .Width = Me.fg.Width, .Height = Me.fg.Height, .Left = Me.fg.Left, .Top = Me.fg.Top, .Anchor = 15,
            .BackColor = System.Drawing.Color.WhiteSmoke, .Visible = False
            }
        Dim lblCancMsg As New System.Windows.Forms.Label With {.Text = "cancelled.", .Top = 5, .Left = 5}
        pnlCanc.Controls.Add(lblCancMsg)
        Me.pnlCanc = pnlCanc
        Me.Controls.Add(pnlCanc)

        Dim txtErr As New System.Windows.Forms.TextBox With {
            .Width = Me.fg.Width, .Height = Me.fg.Height, .Left = Me.fg.Left, .Top = Me.fg.Top, .Anchor = 15,
            .BackColor = Drawing.Color.WhiteSmoke, .Text = "", .Multiline = True, .WordWrap = True, .Visible = False
        }
        txtErr.ReadOnly = True
        Me.txtErr = txtErr
        Me.Controls.Add(txtErr)


        Me.BorderStyle = Windows.Forms.BorderStyle.FixedSingle


        If Not Me.QueryColumn.htSel Is Nothing Then
            Me.htSel = Me.QueryColumn.htSel.Clone
        End If
        Me.Sort = Me.QueryColumn.Sort
        Me.SelectionMode = Me.QueryColumn.SelectionMode
        Me.Sort = Me.QueryColumn.Sort
        Me.txtBox.Text = Me.QueryColumn.SearchTerm
        Me.DaxFilter = Me.QueryColumn.DaxFilter


        Me.btnTy = New System.Windows.Forms.Button With {.Left = 8, .Width = 18, .Height = 18, .Top = 104, .Anchor = 5, .FlatStyle = FlatStyle.Flat, .ImageAlign = ContentAlignment.MiddleCenter}
        Me.btnTy.FlatAppearance.BorderSize = 0
        Me.btnTy.FlatAppearance.MouseOverBackColor = Me.BackColor
        If PBIField.DataType = clsQueryColumn.enDataType.Text Then
            Me.btnTy.Image = Me.ctrlColumnHeader.ctrlTable.ctrlDaxQuery.ImageList.Images("TypeText.ico")
        ElseIf PBIField.DataType = clsQueryColumn.enDataType.Number Then
            Me.btnTy.Image = Me.ctrlColumnHeader.ctrlTable.ctrlDaxQuery.ImageList.Images("TypeNumber.ico")
        ElseIf PBIField.DataType = clsQueryColumn.enDataType.DateTime Then
            Me.btnTy.Image = Me.ctrlColumnHeader.ctrlTable.ctrlDaxQuery.ImageList.Images("TypeDate.ico")
        ElseIf PBIField.DataType = clsQueryColumn.enDataType.Bool Then
            Me.btnTy.Image = Me.ctrlColumnHeader.ctrlTable.ctrlDaxQuery.ImageList.Images("TypeBool.ico")
        Else
            Me.btnTy.Image = Me.ctrlColumnHeader.ctrlTable.ctrlDaxQuery.ImageList.Images("TypeText.ico")
        End If

        Me.Controls.Add(Me.btnTy)
        AddHandler Me.btnTy.Click, AddressOf btnTy_Click

        Me.btnInfo = New System.Windows.Forms.Button With {.Left = 8, .Width = 18, .Height = 18, .Top = 104 + 24, .Anchor = 5, .FlatStyle = FlatStyle.Flat, .ImageAlign = ContentAlignment.MiddleCenter}
        Me.btnInfo.FlatAppearance.BorderSize = 0
        Me.btnInfo.FlatAppearance.MouseOverBackColor = Me.BackColor
        Me.btnInfo.Image = Me.ctrlColumnHeader.ctrlTable.ctrlDaxQuery.ImageList.Images("Info.ico")

        Me.Controls.Add(Me.btnInfo)
        AddHandler Me.btnInfo.Click, AddressOf Me.btnInfo_Click



        If Me.IsImplicitCast = True Then
            Me.ControlState = enControlState.ImpMeasure
            Me.btnTy.Image = Me.ctrlColumnHeader.ctrlTable.ctrlDaxQuery.ImageList.Images("TypeNumber.ico")
            Me.Height = Me.Height - 120
            Me.SelectionMode = enSelectionMode.AllSearch
        Else
            Me.LoadPreview()
        End If









    End Sub

    Private Sub btnTy_Click(sender As Object, e As EventArgs)
    End Sub


    Private Sub btnInfo_Click(sender As Object, e As EventArgs)
        Dim df As New dlgFilter


        Dim d As String = ""
        d += "Sample search terms" & vbCrLf
        d += "" & vbCrLf
        d += "2023" & vbCrLf
        d += "2023 or 2024" & vbCrLf
        d += "202308" & vbCrLf
        d += "year()>=2023" & vbCrLf
        d += "year()=2023 and month()=8" & vbCrLf
        d += "year()>=2020 and year()<=2023 and not 2021" & vbCrLf
        d += ">01-01-2021 and <31-12-2021" & vbCrLf
        d += "today()" & vbCrLf
        d += ">today()-10 and <today()" & vbCrLf
        d += ">=today()-1 and <=today() and hour()=9" & vbCrLf
        d += "(hour()=8 or hour()=9) and (minute()>=10 or minute()<=20)" & vbCrLf
        d += "20230801" & vbCrLf
        d += "blank()" & vbCrLf
        d += "not blank()" & vbCrLf

        Dim n As String = ""
        n += "Sample search terms" & vbCrLf
        n += "" & vbCrLf
        n += "1000" & vbCrLf
        n += ">1000 and <10000" & vbCrLf
        n += "<=1000 or (>10000 and <20000)" & vbCrLf
        n += "0 or blank()" & vbCrLf

        Dim b As String = ""
        b += "Sample search terms" & vbCrLf
        b += "" & vbCrLf
        b += "not False" & vbCrLf
        b += "blank() or false" & vbCrLf
        b += "true" & vbCrLf
        b += "not blank()" & vbCrLf


        Dim t As String = ""
        t += "Sample search terms" & vbCrLf
        t += "" & vbCrLf
        t += "Microsoft" & vbCrLf
        t += "Micro*" & vbCrLf
        t += "*soft" & vbCrLf
        t += "*Las Vegas" & vbCrLf
        t += """*Oregon""" & vbCrLf
        t += "Andrew Schmidt or ""*Oregon""" & vbCrLf
        t += "not blank()" & vbCrLf
        t += "Andre*midt" & vbCrLf
        t += "Andre* and *midt" & vbCrLf





        If Me.IsImplicitCast = True Then
            df.TextBox1.Text = n
        ElseIf Me.DataType = enDataType.DateTime Then
            df.TextBox1.Text = d
        ElseIf Me.DataType = enDataType.Number Then
            df.TextBox1.Text = n
        ElseIf Me.DataType = enDataType.Bool Then
            df.TextBox1.Text = b
        Else
            df.TextBox1.Text = t
        End If


        df.ShowDialog()


    End Sub


    Private Sub CL_Click(sender As Object, e As EventArgs)

        If Me.QueryColumn.Query.SelectionMode = clsQuery.enQuerySelectionMode.Member Then
            For Each qc As clsQueryColumn In Me.QueryColumn.Query.QueryColumns
                qc.SelectedMember.Clear()
            Next qc
            Me.QueryColumn.Query.SelectionMode = clsQuery.enQuerySelectionMode.Total
        End If


        For Each qc As clsQueryColumn In Me.ctrlColumnHeader.ctrlTable.ctrlDaxQuery.query.QueryColumns
            If qc.UniName.ToLower = Me.QueryColumn.UniName.ToLower Then
                qc.SelectionMode = clsQueryColumn.enSelectionMode.AllSelected
                qc.Sort = clsQueryColumn.enSort.none
                qc.SearchTerm = ""
                qc.DaxFilter = ""
                If Not qc.htSel Is Nothing Then
                    qc.htSel.Clear()
                    qc.htSel = Nothing
                End If
                qc.FilterControlGUID = Me.GUID
            End If
        Next qc





        'If Not Me.QueryColumn Is Nothing Then
        '    With Me.QueryColumn
        '        .SelectionMode = clsQueryColumn.enSelectionMode.AllSelected
        '        .Sort = clsQueryColumn.enSort.none
        '        .SearchTerm = ""
        '        .DaxFilter = ""
        '        If Not .htSel Is Nothing Then
        '            .htSel.Clear()
        '            .htSel = Nothing
        '        End If
        '        .FilterControlGUID = Me.GUID
        '    End With
        'End If

        Me.ctrlColumnHeader.ctrlTable.ctrlDaxQuery.FilterControl = Nothing
        Me.ctrlColumnHeader.ctrlTable.ctrlDaxQuery.RefreshPreview()
    End Sub

    Private Sub RE_Click(sender As Object, e As EventArgs)



        If Me.QueryColumn.Query.SelectionMode = clsQuery.enQuerySelectionMode.Member Then
            For Each qc As clsQueryColumn In Me.QueryColumn.Query.QueryColumns
                qc.SelectedMember.Clear()
            Next qc
            Me.QueryColumn.Query.SelectionMode = clsQuery.enQuerySelectionMode.Total
        End If



        Me.ctrlColumnHeader.ctrlTable.ctrlDaxQuery.tm.Cubes(0).DeSelect(Me.QueryColumn.UniName)

        If Not Me.QueryColumn Is Nothing Then
            For Each qc In Me.QueryColumn.Query.QueryColumns
                If qc Is Me.QueryColumn Then
                    Me.QueryColumn.Query.QueryColumns.Remove(qc)
                    Exit For
                End If
            Next qc
        End If



        For i As Integer = Me.ctrlColumnHeader.ctrlTable.fgT.Controls.Count - 1 To 0 Step -1
            If Not TryCast(Me.ctrlColumnHeader.ctrlTable.fgT.Controls.Item(i), ctrlColumnHeader) Is Nothing Then
                Me.ctrlColumnHeader.ctrlTable.fgT.Controls.RemoveAt(i)
            End If
        Next i



        Me.ctrlColumnHeader.ctrlTable.ctrlDaxQuery.FilterControl = Nothing
        Me.ctrlColumnHeader.ctrlTable.ctrlDaxQuery.RefreshPreview()



    End Sub

    Private Sub ZA_Click(sender As Object, e As EventArgs)

        For Each qc As clsQueryColumn In Me.ctrlColumnHeader.ctrlTable.ctrlDaxQuery.query.QueryColumns
            If qc.UniName.ToLower = Me.QueryColumn.UniName.ToLower Then
                qc.Sort = clsQueryColumn.enSort.desc
            Else
                qc.Sort = clsQueryColumn.enSort.none
            End If
        Next qc

        Me.ctrlColumnHeader.ctrlTable.ctrlDaxQuery.FilterControl = Nothing
        Me.ctrlColumnHeader.ctrlTable.ctrlDaxQuery.RefreshPreview()

    End Sub

    Private Sub AZ_Click(sender As Object, e As EventArgs)

        For Each qc As clsQueryColumn In Me.ctrlColumnHeader.ctrlTable.ctrlDaxQuery.query.QueryColumns
            If qc.UniName.ToLower = Me.QueryColumn.UniName.ToLower Then
                qc.Sort = clsQueryColumn.enSort.asc
            Else
                qc.Sort = clsQueryColumn.enSort.none
            End If
        Next qc

        Me.ctrlColumnHeader.ctrlTable.ctrlDaxQuery.FilterControl = Nothing
        Me.ctrlColumnHeader.ctrlTable.ctrlDaxQuery.RefreshPreview()

    End Sub


    Private Sub fg_CellChecked(sender As Object, e As C1.Win.C1FlexGrid.RowColEventArgs)

        If Me.ControlState <> enControlState.Preview Then
            Exit Sub
        End If

        Dim chk As C1.Win.C1FlexGrid.CheckEnum = Me.fg.Rows(e.Row).Node.Checked

        If e.Row = 0 Then

            Me.htSel.Clear()
            If chk = C1.Win.C1FlexGrid.CheckEnum.Checked Then
                Me.SelectionMode = enSelectionMode.DeSelectMember
                Me.fg.BeginUpdate()
                Me.fg.SetData(0, 0, "(Selected)")
                For i As Integer = 1 To Me.fg.Rows.Count - 1
                    Me.fg.Rows(i).Node.Checked = C1.Win.C1FlexGrid.CheckEnum.Checked
                Next i
                Me.fg.EndUpdate()
            ElseIf chk = C1.Win.C1FlexGrid.CheckEnum.Unchecked Then
                Me.SelectionMode = enSelectionMode.SelectMember
                Me.fg.BeginUpdate()
                Me.fg.SetData(0, 0, "(Selected)")
                For i As Integer = 1 To Me.fg.Rows.Count - 1
                    Me.fg.Rows(i).Node.Checked = C1.Win.C1FlexGrid.CheckEnum.Unchecked
                Next i
                Me.fg.EndUpdate()
            End If

        Else


            If Me.SelectionMode = enSelectionMode.AllSelected Or Me.SelectionMode = enSelectionMode.AllSearch Then
                Me.SelectionMode = enSelectionMode.DeSelectMember
                Me.htSel.Clear()
                Me.fg.Rows(0).Node.Checked = C1.Win.C1FlexGrid.CheckEnum.Grayed
                If Me.fg.Rows(e.Row).Node.Checked = C1.Win.C1FlexGrid.CheckEnum.Checked Then
                    Me.htSel.Add(Me.fg.Rows(e.Row).UserData, True)
                Else
                    Me.htSel.Add(Me.fg.Rows(e.Row).UserData, False)
                End If

            ElseIf Me.SelectionMode = enSelectionMode.DeSelectMember Then

                If chk = C1.Win.C1FlexGrid.CheckEnum.Unchecked Then
                    If Me.htSel.Contains(Me.fg.Rows(e.Row).UserData) = False Then
                        Me.htSel.Add(Me.fg.Rows(e.Row).UserData, False)
                    Else
                        Me.htSel(Me.fg.Rows(e.Row).UserData) = False
                    End If
                Else
                    If Me.htSel.Contains(Me.fg.Rows(e.Row).UserData) = True Then
                        Me.htSel.Remove(Me.fg.Rows(e.Row).UserData)
                    End If
                End If

                If htSel.Count > 0 Then
                    Me.fg.Rows(0).Node.Checked = C1.Win.C1FlexGrid.CheckEnum.Grayed
                Else
                    Me.fg.Rows(0).Node.Checked = C1.Win.C1FlexGrid.CheckEnum.Checked
                    If Me.txtBox.Text.Trim = "" Then
                        Me.SelectionMode = enSelectionMode.AllSelected
                    Else
                        Me.SelectionMode = enSelectionMode.AllSearch
                    End If
                End If

            ElseIf Me.SelectionMode = enSelectionMode.SelectMember Then

                If chk = C1.Win.C1FlexGrid.CheckEnum.Checked Then
                    If Me.htSel.Contains(Me.fg.Rows(e.Row).UserData) = False Then
                        Me.htSel.Add(Me.fg.Rows(e.Row).UserData, True)
                    Else
                        Me.htSel(Me.fg.Rows(e.Row).UserData) = True
                    End If
                Else
                    If Me.htSel.Contains(Me.fg.Rows(e.Row).UserData) = True Then
                        Me.htSel.Remove(Me.fg.Rows(e.Row).UserData)
                    End If
                End If

                If htSel.Count > 0 Then
                    Me.fg.Rows(0).Node.Checked = C1.Win.C1FlexGrid.CheckEnum.Grayed
                Else
                    Me.fg.Rows(0).Node.Checked = C1.Win.C1FlexGrid.CheckEnum.Unchecked
                End If

            End If


        End If

    End Sub






    Private Sub txtBox_KeyPress(sender As Object, e As System.Windows.Forms.KeyPressEventArgs)

        If e.KeyChar = ChrW(Keys.Enter) Then
            e.Handled = True

            If Me.QueryColumn.isImplicitCast = False Then
                Me.LoadPreview()
            End If

        End If

    End Sub


    Private Sub lblCancel_Click(sender As Object, e As EventArgs)

        Me.lblCancel.InvokeIfRequired(Sub()
                                          Me.lblCancel.Text = "cancelling.."
                                          Me.lblCancel.Refresh()
                                      End Sub)
        Me.ControlState = enControlState.Cancelled
    End Sub

    Private Sub lblClear_Click(sender As Object, e As EventArgs)

        Me.ctrlColumnHeader.ctrlTable.ctrlDaxQuery.FilterControl = Nothing

        Exit Sub



        If Me.QueryColumn.Query.SelectionMode = clsQuery.enQuerySelectionMode.Member Then
            For Each qc As clsQueryColumn In Me.QueryColumn.Query.QueryColumns
                qc.SelectedMember.Clear()
            Next qc
            Me.QueryColumn.Query.SelectionMode = clsQuery.enQuerySelectionMode.Total
        End If




        If Not Me.QueryColumn Is Nothing Then
            With Me.QueryColumn
                .SelectionMode = clsQueryColumn.enSelectionMode.AllSelected
                .Sort = clsQueryColumn.enSort.none
                .SearchTerm = ""
                .DaxFilter = ""
                If Not .htSel Is Nothing Then
                    .htSel.Clear()
                    .htSel = Nothing
                End If
                .FilterControlGUID = Me.GUID
            End With
        End If

        With Me.PBIField

        End With

        Me.ctrlColumnHeader.ctrlTable.ctrlDaxQuery.FilterControl = Nothing

        Me.ctrlColumnHeader.ctrlTable.ctrlDaxQuery.RefreshPreview()



    End Sub


    Private _errValid As Exception
    Public ReadOnly Property ValidExpression As Boolean
        Get
            Try
                Me._errValid = Nothing
                If Me.objConn Is Nothing Then
                    Me.objConn = CreateObject("ADODB.CONNECTION")
                    Me.objConn.connectionstring = Me.ctrlColumnHeader.ctrlTable.ctrlDaxQuery.ConnectionString
                End If
                If Me.objConn.state <> 1 Then
                    Me.objConn.open
                End If
                Me.objRec = CreateObject("ADODB.RECORDSET")
                Dim strCmd As String = Me.DaxStmnt
                Me.objRec.open(strCmd, Me.objConn)
                Me.objRec.close
                Return True
            Catch ex As Exception
                Me._errValid = ex
                Return False
            End Try
        End Get
    End Property



    Private Sub lblApply_Click(sender As Object, e As EventArgs)


        If Me.txtBox.Text.Trim <> "" Then
            Cursor.Current = Cursors.WaitCursor
            If Me.ValidExpression = False And Me.QueryColumn.FieldType <> clsQueryColumn.enFieldType.ImpMeasure Then
                Cursor.Current = Cursors.Default
                MsgBox("Please check the filter expression. " & Me._errValid.Message, MsgBoxStyle.Critical)
                Exit Sub
            End If
            Cursor.Current = Cursors.Default
        End If



        If Me.QueryColumn.Query.SelectionMode = clsQuery.enQuerySelectionMode.Member Then
            For Each qc As clsQueryColumn In Me.QueryColumn.Query.QueryColumns
                qc.SelectedMember.Clear()
            Next qc
            Me.QueryColumn.Query.SelectionMode = clsQuery.enQuerySelectionMode.Total
        End If

        With Me.QueryColumn

            .IsSelected = True
            .SelectionMode = Me.SelectionMode
            If Not Me.htSel Is Nothing Then
                If Me.htSel.Count = 0 Then
                    If Me.txtBox.Text.Trim <> "" Then
                        .SelectionMode = enSelectionMode.AllSearch
                    ElseIf Me.txtBox.Text.Trim = "" Then
                        .SelectionMode = clsQueryColumn.enSelectionMode.AllSelected
                    End If
                End If
            End If
            .htSel = Me.htSel.Clone
            .Sort = Me.Sort
            .SearchTerm = Me.txtBox.Text


            .DaxStmnt = Me.DaxStmnt
            .DaxFilter = Me.DaxFilter


            .FilterControlGUID = Me.GUID
        End With


        Me.ctrlColumnHeader.ctrlTable.ctrlDaxQuery.FilterControl = Nothing
        Me.ctrlColumnHeader.ctrlTable.ctrlDaxQuery.RefreshPreview()






    End Sub


    Private Sub LoadPreview()



        If Me.txtBox.Text.Trim <> "" And Me.SelectionMode = enSelectionMode.AllSelected Then
            Me.SelectionMode = enSelectionMode.AllSearch
        ElseIf Me.txtBox.Text.Trim = "" And Me.SelectionMode = enSelectionMode.AllSearch Then
            Me.SelectionMode = enSelectionMode.AllSelected
        End If

        If Me.RunAsync = True Then

            If Not Me.ctsSource Is Nothing Then
                If Me.ctsSource.IsCancellationRequested = False Then
                    Me.ctsSource.Cancel()
                End If
            End If
            Me.ctsSource = New System.Threading.CancellationTokenSource
            Me.cts = Me.ctsSource.Token
            Me.ControlState = enControlState.Executing
            Me.lblCancel.Text = "loading.."





            Me.GetPreview_async(Me.cts)

        Else



            If Me.FieldType = enFieldType.Level Then
                'Me.objConn = MyAddIn.GetWorkbook(Me.PBILevel.Hierarchy.Dimension.Cube.tm.Connection.WorkbookName).Model.DataModelConnection.ModelConnection.ADOConnection
            Else
                'Me.objConn = MyAddIn.GetWorkbook(Me.PBIMeasure.Dimension.Cube.tm.Connection.WorkbookName).Model.DataModelConnection.ModelConnection.ADOConnection
            End If
            Me.ControlState = enControlState.Executing
            Me.objRec = CreateObject("ADODB.RECORDSET")
            Me.GetPreview_sync(Me.objConn, Me.objRec, Nothing)

        End If




    End Sub

    Private __dt As DataTable
    Private Property dt As DataTable
        Get
            Return __dt
        End Get
        Set(value As DataTable)
            Me.__dt = value
            If Not Me.__dt Is Nothing Then
                If Me.__dt.Rows.Count > 0 Then

                    Me.RefreshPreview()

                Else
                    Me.ControlState = enControlState.NoResult
                End If
            End If
        End Set
    End Property

    Private _objErr As Exception
    Private Property objErr As Exception
        Get
            Return Me._objErr
        End Get
        Set(value As Exception)
            _objErr = value
            Me.ControlState = enControlState.IsError
        End Set
    End Property


    Private Property objConn As Object
    Private Property objRec As Object



    Private Property htSel As New Hashtable


    Private Sub GetPreview_async(cts As System.Threading.CancellationToken)
        Dim t = Task(Of Object).Factory.StartNew(Function()

                                                     Dim intProgess As Integer = 0
                                                     Dim blnCancelled As Boolean = False
                                                     Try

                                                         Dim blnStarted As Boolean = False

                                                         Me.cts.Register(Function()
                                                                             If blnStarted = True Then
                                                                                 blnCancelled = True
                                                                                 Try
                                                                                     If Not objRec Is Nothing Then
                                                                                         If objRec.state = 4 Then
                                                                                             objRec.cancel
                                                                                             objRec = Nothing
                                                                                         End If
                                                                                     End If
                                                                                     Me.dt = Nothing
                                                                                 Catch ex As Exception
                                                                                     Me.dt = Nothing
                                                                                 End Try
                                                                             End If
                                                                             blnStarted = False
                                                                             Return Nothing
                                                                         End Function)


                                                         blnStarted = True

                                                         intProgess = 1


                                                         Me.objRec = CreateObject("ADODB.Recordset")





                                                         If blnCancelled = True Then
                                                             Return Nothing
                                                         End If

                                                         intProgess = 2
                                                         Dim Stmnt As String = Me.DaxStmnt

                                                         'Debug.Print(Stmnt)

                                                         intProgess = 3
                                                         objRec.open(Stmnt, Me.ctrlColumnHeader.ctrlTable.ctrlDaxQuery.FilterConn, 0)
                                                         If blnCancelled = True Then
                                                             Return Nothing
                                                         End If


                                                         intProgess = 4
                                                         Dim adap As New System.Data.OleDb.OleDbDataAdapter
                                                         Dim _dt As New DataTable
                                                         adap.Fill(_dt, objRec)

                                                         If _dt.Rows.Count = 0 Then
                                                             Me.dt = _dt
                                                             Return Nothing
                                                         End If


                                                         If Not objRec Is Nothing Then
                                                             If objRec.state = 1 Then
                                                                 objRec.close
                                                             End If
                                                             objRec = Nothing
                                                         End If



                                                         If blnCancelled = True Then
                                                             Return Nothing
                                                         End If

                                                         Dim c As DataColumn = _dt.Columns.Add("chk", GetType(System.Boolean))
                                                         Dim cd As DataColumn = _dt.Columns.Add("disp", GetType(System.String))
                                                         c.SetOrdinal(0) : cd.SetOrdinal(1)

                                                         If Me.SelectionMode = enSelectionMode.AllSelected Then
                                                             For Each r As DataRow In _dt.Rows
                                                                 r(0) = True
                                                                 If r(2) Is DBNull.Value Then
                                                                     r(1) = "(Blank)"
                                                                 Else
                                                                     r(1) = r(2).ToString
                                                                 End If
                                                             Next r
                                                             If _dt.Rows.Count > 0 Then
                                                                 Dim sr As DataRow = _dt.NewRow
                                                                 sr(0) = True
                                                                 sr(1) = "(Select All)"
                                                                 _dt.Rows.InsertAt(sr, 0)
                                                             End If
                                                         ElseIf Me.SelectionMode = enSelectionMode.AllSearch Then
                                                             For Each r As DataRow In _dt.Rows
                                                                 r(0) = True
                                                                 If r(2) Is DBNull.Value Then
                                                                     r(1) = "(Blank)"
                                                                 Else
                                                                     r(1) = r(2).ToString
                                                                 End If
                                                             Next r
                                                             If _dt.Rows.Count > 0 Then
                                                                 Dim sr As DataRow = _dt.NewRow
                                                                 sr(0) = True
                                                                 sr(1) = "(Select All Search Results)"
                                                                 _dt.Rows.InsertAt(sr, 0)
                                                             End If
                                                         ElseIf Me.SelectionMode = enSelectionMode.DeSelectMember Then

                                                             Dim blnF As Boolean = False

                                                             For Each r As DataRow In _dt.Rows
                                                                 If Me.DataType <> enDataType.Bool Then
                                                                     blnF = False
                                                                     For Each k As DictionaryEntry In htSel
                                                                         If k.Key Is Nothing AndAlso IsDBNull(r(2)) = True Then
                                                                             blnF = True : Exit For
                                                                         Else
                                                                             If k.Key Is r(2) Then
                                                                                 blnF = True : Exit For
                                                                             End If
                                                                         End If
                                                                     Next k
                                                                 End If

                                                                 If htSel.ContainsKey(r(2)) = True OrElse blnF = True Then
                                                                     r(0) = False
                                                                 Else
                                                                     r(0) = True
                                                                 End If
                                                                 If r(2) Is DBNull.Value Then
                                                                     r(1) = "(Blank)"
                                                                 Else
                                                                     r(1) = r(2).ToString
                                                                 End If
                                                             Next r
                                                             If _dt.Rows.Count > 0 Then
                                                                 Dim sr As DataRow = _dt.NewRow
                                                                 If htSel.Count > 0 Then
                                                                     sr(0) = DBNull.Value
                                                                 Else
                                                                     sr(0) = True
                                                                 End If
                                                                 sr(1) = "(Selected)"
                                                                 _dt.Rows.InsertAt(sr, 0)
                                                             End If
                                                         ElseIf Me.SelectionMode = enSelectionMode.SelectMember Then

                                                             Dim blnF As Boolean = False
                                                             For Each r As DataRow In _dt.Rows
                                                                 If Me.DataType <> enDataType.Bool Then
                                                                     blnF = False
                                                                     For Each k As DictionaryEntry In htSel
                                                                         If k.Key Is Nothing AndAlso IsDBNull(r(2)) = True Then
                                                                             blnF = True : Exit For
                                                                         Else
                                                                             If k.Key Is r(2) Then
                                                                                 blnF = True : Exit For
                                                                             End If
                                                                         End If
                                                                     Next k
                                                                 End If

                                                                 If htSel.ContainsKey(r(2)) = True OrElse blnF = True Then
                                                                     r(0) = True
                                                                 Else
                                                                     r(0) = False
                                                                 End If

                                                                 If r(2) Is DBNull.Value Then
                                                                     r(1) = "(Blank)"
                                                                 Else
                                                                     r(1) = r(2).ToString
                                                                 End If
                                                             Next r
                                                             If _dt.Rows.Count > 0 Then
                                                                 Dim sr As DataRow = _dt.NewRow
                                                                 If htSel.Count > 0 Then
                                                                     sr(0) = DBNull.Value
                                                                 Else
                                                                     sr(0) = False
                                                                 End If
                                                                 sr(1) = "(Selected)"
                                                                 _dt.Rows.InsertAt(sr, 0)
                                                             End If
                                                         End If

                                                         If blnCancelled = True Then
                                                             Return Nothing
                                                         End If

                                                         Me.dt = _dt

                                                     Catch ex As Exception

                                                         If blnCancelled = True Then
                                                             Me.dt = Nothing
                                                             Me.ControlState = enControlState.Cancelled
                                                         Else
                                                             Me.dt = Nothing
                                                             If intProgess = 1 Then
                                                                 Me.objErr = New Exception("Connection error: " & ex.Message)
                                                             ElseIf intProgess = 2 Then
                                                                 Me.objErr = New Exception("Error thrown by parsing engine: " & ex.Message)
                                                             ElseIf intProgess = 3 Then
                                                                 Me.objErr = New Exception("Expression error: " & Me.DaxFilter)
                                                             ElseIf intProgess = 4 Then
                                                                 Me.objErr = New Exception("Error processing result: " & ex.Message)
                                                             Else
                                                                 Me.objErr = ex
                                                             End If
                                                         End If

                                                         Return Nothing
                                                     End Try

                                                     Return Nothing

                                                 End Function)



    End Sub


    Private Sub GetPreview_sync(objConn As Object, objRec As Object, cts As System.Threading.CancellationToken)


        Dim intProgess As Integer = 0
        Dim blnCancelled As Boolean = False

        Try


            Dim blnStarted As Boolean = False

            blnStarted = True

            intProgess = 1


            If blnCancelled = True Then
                Exit Sub
            End If

            intProgess = 2
            Dim Stmnt As String = Me.DaxStmnt


            intProgess = 3
            objRec.open(Stmnt, objConn, 0)
            If blnCancelled = True Then
                Exit Sub
            End If


            intProgess = 4
            Dim adap As New System.Data.OleDb.OleDbDataAdapter
            Dim _dt As New DataTable
            adap.Fill(_dt, objRec)

            If _dt.Rows.Count = 0 Then
                Me.dt = _dt
                Exit Sub
            End If


            If Not objRec Is Nothing Then
                If objRec.state = 1 Then
                    objRec.close
                End If
                objRec = Nothing
            End If



            If blnCancelled = True Then
                Exit Sub
            End If

            Dim c As DataColumn = _dt.Columns.Add("chk", GetType(System.Boolean))
            Dim cd As DataColumn = _dt.Columns.Add("disp", GetType(System.String))
            c.SetOrdinal(0) : cd.SetOrdinal(1)

            If Me.SelectionMode = enSelectionMode.AllSelected Then
                For Each r As DataRow In _dt.Rows
                    r(0) = True
                    If r(2) Is DBNull.Value Then
                        r(1) = "(Blank)"
                    Else
                        r(1) = r(2).ToString
                    End If
                Next r
                If _dt.Rows.Count > 0 Then
                    Dim sr As DataRow = _dt.NewRow
                    sr(0) = True
                    sr(1) = "(Select All)"
                    _dt.Rows.InsertAt(sr, 0)
                End If
            ElseIf Me.SelectionMode = enSelectionMode.AllSearch Then
                For Each r As DataRow In _dt.Rows
                    r(0) = True
                    If r(2) Is DBNull.Value Then
                        r(1) = "(Blank)"
                    Else
                        r(1) = r(2).ToString
                    End If
                Next r
                If _dt.Rows.Count > 0 Then
                    Dim sr As DataRow = _dt.NewRow
                    sr(0) = True
                    sr(1) = "(Select All Search Results)"
                    _dt.Rows.InsertAt(sr, 0)
                End If
            ElseIf Me.SelectionMode = enSelectionMode.DeSelectMember Then

                Dim blnF As Boolean = False

                For Each r As DataRow In _dt.Rows
                    If Me.DataType <> enDataType.Bool Then
                        blnF = False
                        For Each k As DictionaryEntry In htSel
                            If k.Key Is Nothing AndAlso IsDBNull(r(2)) = True Then
                                blnF = True : Exit For
                            Else
                                If k.Key Is r(2) Then
                                    blnF = True : Exit For
                                End If
                            End If
                        Next k
                    End If

                    If htSel.ContainsKey(r(2)) = True OrElse blnF = True Then
                        r(0) = False
                    Else
                        r(0) = True
                    End If
                    If r(2) Is DBNull.Value Then
                        r(1) = "(Blank)"
                    Else
                        r(1) = r(2).ToString
                    End If
                Next r
                If _dt.Rows.Count > 0 Then
                    Dim sr As DataRow = _dt.NewRow
                    If htSel.Count > 0 Then
                        sr(0) = DBNull.Value
                    Else
                        sr(0) = True
                    End If
                    sr(1) = "(Selected)"
                    _dt.Rows.InsertAt(sr, 0)
                End If
            ElseIf Me.SelectionMode = enSelectionMode.SelectMember Then

                Dim blnF As Boolean = False
                For Each r As DataRow In _dt.Rows
                    If Me.DataType <> enDataType.Bool Then
                        blnF = False
                        For Each k As DictionaryEntry In htSel
                            If k.Key Is Nothing AndAlso IsDBNull(r(2)) = True Then
                                blnF = True : Exit For
                            Else
                                If k.Key Is r(2) Then
                                    blnF = True : Exit For
                                End If
                            End If
                        Next k
                    End If

                    If htSel.ContainsKey(r(2)) = True OrElse blnF = True Then
                        r(0) = True
                    Else
                        r(0) = False
                    End If

                    If r(2) Is DBNull.Value Then
                        r(1) = "(Blank)"
                    Else
                        r(1) = r(2).ToString
                    End If
                Next r
                If _dt.Rows.Count > 0 Then
                    Dim sr As DataRow = _dt.NewRow
                    If htSel.Count > 0 Then
                        sr(0) = DBNull.Value
                    Else
                        sr(0) = False
                    End If
                    sr(1) = "(Selected)"
                    _dt.Rows.InsertAt(sr, 0)
                End If
            End If

            If blnCancelled = True Then
                Exit Sub
            End If

            Me.dt = _dt

        Catch ex As Exception

            If blnCancelled = True Then
                Me.dt = Nothing
                Me.ControlState = enControlState.Cancelled
            Else
                Me.dt = Nothing
                If intProgess = 1 Then
                    Me.objErr = New Exception("Connection error: " & ex.Message)
                ElseIf intProgess = 2 Then
                    Me.objErr = New Exception("Error thrown by parsing engine: " & ex.Message)
                ElseIf intProgess = 3 Then
                    Me.objErr = New Exception("Expression error: " & Me.DaxFilter)
                ElseIf intProgess = 4 Then
                    Me.objErr = New Exception("Error processing result: " & ex.Message)
                Else
                    Me.objErr = ex
                End If
            End If

        End Try


    End Sub



    Private Sub RefreshPreview()

        If Me.IsDisposed = True OrElse Me.Disposing = True Then
            Exit Sub
        End If

        Me.fg.InvokeIfRequired(Sub()
                                   Try
                                       Me.fg.Visible = False
                                       Me.txtErr.Visible = False

                                       If Me.dt Is Nothing OrElse Me.dt.Rows.Count = 0 Then
                                           Exit Sub
                                       End If

                                       With Me.fg
                                           .BeginUpdate()
                                           .Redraw = False

                                           .Rows.Count = Me.dt.Rows.Count
                                           For i As Integer = 0 To Me.dt.Rows.Count - 1
                                               .SetData(i, 0, Me.dt(i)(1))
                                               .Rows(i).IsNode = True
                                               If Me.dt(i)(0) Is DBNull.Value Then
                                                   .SetCellCheck(i, 0, C1.Win.C1FlexGrid.CheckEnum.Unchecked)
                                               ElseIf Me.dt(i)(0) = False Then
                                                   .SetCellCheck(i, 0, C1.Win.C1FlexGrid.CheckEnum.Unchecked)
                                               Else
                                                   .SetCellCheck(i, 0, C1.Win.C1FlexGrid.CheckEnum.Checked)
                                               End If
                                               .Rows(i).UserData = Me.dt(i)(2)
                                           Next i

                                           If Me.htSel.Count > 0 Then
                                               Me.fg.Rows(0).Node.Checked = C1.Win.C1FlexGrid.CheckEnum.Grayed
                                           End If


                                           .Redraw = True
                                           .EndUpdate()
                                       End With



                                       Me.ControlState = enControlState.Preview
                                   Catch ex As Exception
                                       Me.objErr = New Exception("unable to present result. " & ex.Message)
                                       Me.ControlState = enControlState.IsError
                                   End Try



                               End Sub
            )

    End Sub




    Private Class ColumnP

        Friend Property Search As String
        Friend Property P As String
        Friend Property StartPos As Integer
        Friend Property EndPos As Integer
        Friend Property fc As ctrlFilter
        Friend ReadOnly Property Length As Integer
            Get
                Return Me.EndPos - Me.StartPos + 1
            End Get
        End Property

        Friend Property Pos As Integer

        Private ReadOnly Property Is_Not As Boolean
            Get
                If Me.P.ToLower.Trim.StartsWith("not ") Then
                    Return True
                Else
                    Return False
                End If
            End Get
        End Property

        Private ReadOnly Property NotOp As String
            Get
                If Me.P.ToLower.Trim.StartsWith("not ") Then
                    Return "not "
                Else
                    Return ""
                End If
            End Get
        End Property


        Private ReadOnly Property Ope As String
            Get
                Dim strR As String = Me.P.Trim
                If Me.Is_Not = True Then
                    strR = strR.ToLower.Replace("not ", "").Trim
                End If

                If strR.StartsWith("<=") OrElse strR.StartsWith("<>") OrElse strR.StartsWith(">=") Then
                    Return strR.Substring(0, 2)
                End If

                If strR.StartsWith("<") OrElse strR.StartsWith("=") OrElse strR.StartsWith(">") Then
                    Return strR.Substring(0, 1)
                End If

                If strR.Trim.Replace(" ", "").ToLower.Contains("len()") _
                    Or strR.Trim.Replace(" ", "").ToLower.Contains("year()") _
                    Or strR.Trim.Replace(" ", "").ToLower.Contains("month()") _
                    Or strR.Trim.Replace(" ", "").ToLower.Contains("day()") _
                    Or strR.Trim.Replace(" ", "").ToLower.Contains("hour()") _
                    Or strR.Trim.Replace(" ", "").ToLower.Contains("minute()") _
                    Or strR.Trim.Replace(" ", "").ToLower.Contains("second()") Then

                    If strR.LastIndexOf("<=") > 0 Then
                        Return strR.Substring(strR.LastIndexOf("<="))
                    ElseIf strR.LastIndexOf("<>") > 0 Then
                        Return strR.Substring(strR.LastIndexOf("<>"))
                    ElseIf strR.LastIndexOf(">=") > 0 Then
                        Return strR.Substring(strR.LastIndexOf(">="))
                    ElseIf strR.LastIndexOf("<") > 0 Then
                        Return strR.Substring(strR.LastIndexOf("<"))
                    ElseIf strR.LastIndexOf("=") > 0 Then
                        Return strR.Substring(strR.LastIndexOf("="))
                    ElseIf strR.LastIndexOf(">") > 0 Then
                        Return strR.Substring(strR.LastIndexOf(">"))
                    End If

                End If



                Return ""
            End Get
        End Property

        Private ReadOnly Property Quotes As Boolean
            Get
                Dim strR As String = Me.P.Trim
                If Me.Is_Not = True Then
                    strR = strR.ToLower.Replace("not ", "").Trim
                End If

                If Me.Ope <> "" Then
                    strR = strR.ToLower.Replace(Me.Ope, "").Trim
                End If

                If strR.StartsWith("""") And strR.EndsWith("""") Then
                    Return True
                End If

                Return False

            End Get
        End Property

        Private ReadOnly Property searchCrit As String
            Get
                Dim strR As String = Me.P.Trim
                If Me.Is_Not = True Then
                    'strR = strR.ToLower.Replace("not ", "").Trim
                    If strR.IndexOf("not ", StringComparison.CurrentCultureIgnoreCase) >= 0 Then
                        strR = strR.Substring(strR.IndexOf("not ", StringComparison.CurrentCultureIgnoreCase) + 4).Trim
                    End If
                End If

                If Me.Ope <> "" Then
                    strR = strR.Trim.Substring(Me.Ope.Length)
                End If

                If strR.StartsWith("""") And strR.EndsWith("""") Then
                    strR = strR.Substring(1, strR.Length - 2)
                End If

                If Me.fc.DataType = enDataType.Number Then
                    Return strR.Replace(",", ".")
                Else
                    Return strR
                End If



            End Get
        End Property

        Private ReadOnly Property EightDigits As Boolean
            Get
                Dim strR As String = Me.P
                strR = strR.Replace(" ", "").ToLower.Trim

                If strR.Length < 8 Then Return False

                Dim strC() As Char = strR.ToCharArray
                For Each c In strC
                    If IsNumeric(c) = False Then
                        Return False
                    End If
                Next c

                strR = strR.Substring(strR.Length - 8, 8)
                Dim intY As Integer = Nothing
                If Integer.TryParse(strR, intY) = False Then
                    Return False
                Else
                    If intY >= 10000000 Then
                        Return True
                    Else
                        Return False
                    End If
                End If
            End Get
        End Property

        Private ReadOnly Property SixDigits As Boolean
            Get
                If Me.EightDigits = True Then
                    Return False
                End If

                Dim strR As String = Me.P
                strR = strR.Replace(" ", "").ToLower.Trim

                If strR.Length < 6 Then Return False

                Dim strC() As Char = strR.ToCharArray
                For Each c In strC
                    If IsNumeric(c) = False Then
                        Return False
                    End If
                Next c


                strR = strR.Substring(strR.Length - 6, 6)
                Dim intY As Integer = Nothing
                If Integer.TryParse(strR, intY) = False Then
                    Return False
                Else
                    If intY >= 100000 Then
                        Return True
                    Else
                        Return False
                    End If
                End If
            End Get
        End Property


        Private ReadOnly Property FourDigits As Boolean
            Get
                If Me.SixDigits = True OrElse Me.EightDigits = True Then
                    Return False
                End If

                Dim strR As String = Me.P
                strR = strR.Replace(" ", "").ToLower.Trim

                If strR.Length < 4 Then Return False

                Dim strC() As Char = Me.searchCrit.Replace(" ", "").Trim
                For Each c In strC
                    If IsNumeric(c) = False Then
                        Return False
                    End If
                Next c

                strR = strR.Substring(strR.Length - 4, 4)
                Dim intY As Integer = Nothing
                If Integer.TryParse(strR, intY) = False Then
                    Return False
                Else
                    If intY >= 1000 Then
                        Return True
                    Else
                        Return False
                    End If
                End If
            End Get
        End Property



        Private ReadOnly Property SecondFunc() As Boolean
            Get
                If Me.P.Trim.Replace(" ", "").ToLower.Contains("second(") Then
                    Return True
                Else
                    Return False
                End If
            End Get
        End Property

        Private ReadOnly Property MinuteFunc() As Boolean
            Get
                If Me.P.Trim.Replace(" ", "").ToLower.Contains("minute(") Then
                    Return True
                Else
                    Return False
                End If
            End Get
        End Property

        Private ReadOnly Property HourFunc() As Boolean
            Get
                If Me.P.Trim.Replace(" ", "").ToLower.Contains("hour(") Then
                    Return True
                Else
                    Return False
                End If
            End Get
        End Property

        Private ReadOnly Property DayFunc() As Boolean
            Get
                If Me.P.Trim.Replace(" ", "").ToLower.Contains("day(") Then
                    Return True
                Else
                    Return False
                End If
            End Get
        End Property

        Private ReadOnly Property MonthFunc() As Boolean
            Get
                If Me.P.Trim.Replace(" ", "").ToLower.Contains("month(") Then
                    Return True
                Else
                    Return False
                End If
            End Get
        End Property

        Private ReadOnly Property YearFunc() As Boolean
            Get
                If Me.P.Trim.Replace(" ", "").ToLower.Contains("year(") Then
                    Return True
                Else
                    Return False
                End If
            End Get
        End Property

        Private ReadOnly Property TodayFunc() As Boolean
            Get
                If Me.P.Trim.Replace(" ", "").ToLower.Contains("today(") Then
                    Return True
                Else
                    Return False
                End If
            End Get
        End Property

        Private ReadOnly Property LenFunc() As Boolean
            Get
                If Me.P.Trim.Replace(" ", "").ToLower.Contains("len(") Then
                    Return True
                Else
                    Return False
                End If
            End Get
        End Property


        Private ReadOnly Property BlankFunc() As Boolean
            Get
                If Me.searchCrit.Trim.Replace(" ", "").Contains("blank()") Then
                    Return True
                Else
                    Return False
                End If
            End Get
        End Property

        Private ReadOnly Property TrueFalse As Boolean
            Get
                If Me.searchCrit.Trim.ToLower.Contains("=true") Or Me.searchCrit.ToLower.Contains("= true") Or Me.searchCrit.ToLower.Contains("=false") Or Me.searchCrit.Contains("= false") Then
                    Return True
                Else
                    Return False
                End If
            End Get
        End Property

        Private ReadOnly Property pTrueFalse As String
            Get
                If Me.searchCrit.Trim.ToLower.Contains("=true") Or Me.searchCrit.ToLower.Contains("= true") Then
                    Return "true"
                ElseIf Me.searchCrit.ToLower.Contains("=false") Or Me.searchCrit.Contains("= false") Then
                    Return "false"
                Else
                    Return ""
                End If
            End Get
        End Property

        Private ReadOnly Property Asteriks As Boolean
            Get
                Dim strR As String = Me.P.Trim
                If Me.Is_Not = True Then
                    strR = strR.ToLower.Replace("not ", "").Trim
                End If

                If Me.Ope <> "" Then
                    strR = strR.ToLower.Replace(Me.Ope, "").Trim
                End If

                If strR.Contains("*") = True Then
                    If strR.Trim.StartsWith("*") OrElse strR.Trim.EndsWith("*") Then
                        Return True
                    End If
                End If

                Return False
            End Get
        End Property

        Private ReadOnly Property InBetAsteriks As Boolean
            Get
                Dim strR As String = Me.P.Trim
                If Me.Is_Not = True Then
                    strR = strR.ToLower.Replace("not ", "").Trim
                End If

                If Me.Ope <> "" Then
                    strR = strR.ToLower.Replace(Me.Ope, "").Trim
                End If

                If strR.Contains("*") = True Then
                    If strR.Trim.StartsWith("*") = False AndAlso strR.Trim.EndsWith("*") = False Then
                        Return True
                    End If
                End If

                Return False
            End Get
        End Property

        Private Function Delimiter(strR As String) As String
            Dim strD As String = ""
            Dim ctr As Integer = 0
            Dim strA = strR.ToCharArray
            For Each s In strA
                If IsNumeric(s) = False Then
                    ctr += 1
                    If strD = "" Then
                        strD = s
                    Else
                        If s <> strD Then
                            Return ""
                        End If
                    End If
                End If
            Next s
            If ctr <> 2 Then
                Return ""
            End If
            Return strD


        End Function


        Private ReadOnly Property IsDateTime As Boolean
            Get
                Dim strR As String = Me.P.Trim
                strR = strR.ToLower.Replace("<", "").Replace(">", "").Replace("=", "").Replace("not", "").Replace("""", "").Trim

                Dim d As DateTime = Nothing
                DateTime.TryParse(strR, d)
                If d = Nothing Then
                    Return False
                Else
                    Return True
                End If
                Return False

            End Get
        End Property


        Friend ReadOnly Property Dax As String
            Get
                Dim strR As String = ""


                If Me.fc.DataType = enDataType.Bool And Me.fc.IsImplicitCast = False Then

                    If Me.BlankFunc = True Then

                        If Me.TrueFalse = False Then
                            If Me.Ope = "" Then
                                If Me.Is_Not = False Then
                                    strR = "Isblank(" & Me.fc.ColUni & ")=True"
                                Else
                                    strR = "not Isblank(" & Me.fc.ColUni & ")=True"
                                End If
                            Else
                                If Me.Is_Not = False Then
                                    If Me.Ope = "<>" Then
                                        strR = "Isblank(" & Me.fc.ColUni & ")=false"
                                    ElseIf Me.Ope = "=" Then
                                        strR = "Isblank(" & Me.fc.ColUni & ")=true"
                                    End If
                                Else
                                    If Me.Ope = "<>" Then
                                        strR = "NOT Isblank(" & Me.fc.ColUni & ")=false"
                                    ElseIf Me.Ope = "=" Then
                                        strR = "NOT Isblank(" & Me.fc.ColUni & ")=true"
                                    End If
                                End If
                            End If

                        Else
                            If Me.pTrueFalse = "true" Then
                                If Me.Is_Not = False Then
                                    If Me.Ope = "" Or Me.Ope = "=" Then
                                        strR = "Isblank(" & Me.fc.ColUni & ")=True"
                                    ElseIf Me.Ope = "<>" Then
                                        strR = "Isblank(" & Me.fc.ColUni & ")=false"
                                    End If
                                Else
                                    If Me.Ope = "" Or Me.Ope = "=" Then
                                        strR = "NOT Isblank(" & Me.fc.ColUni & ")=True"
                                    ElseIf Me.Ope = "<>" Then
                                        strR = "NOT Isblank(" & Me.fc.ColUni & ")=false"
                                    End If
                                End If
                            Else
                                If Me.Is_Not = False Then
                                    If Me.Ope = "" Or Me.Ope = "=" Then
                                        strR = "Isblank(" & Me.fc.ColUni & ")=false"
                                    ElseIf Me.Ope = "<>" Then
                                        strR = "Isblank(" & Me.fc.ColUni & ")=true"
                                    End If
                                Else
                                    If Me.Ope = "" Or Me.Ope = "=" Then
                                        strR = "NOT Isblank(" & Me.fc.ColUni & ")=false"
                                    ElseIf Me.Ope = "<>" Then
                                        strR = "NOT Isblank(" & Me.fc.ColUni & ")=true"
                                    End If
                                End If
                            End If

                        End If


                    Else

                        If Me.Ope = "" Then
                            If Me.Is_Not = False Then
                                strR = Me.fc.ColUni & "=" & Me.searchCrit
                            Else
                                strR = "NOT " & Me.fc.ColUni & "=" & Me.searchCrit
                            End If

                        Else
                            If Me.Is_Not = False Then
                                strR = Me.fc.ColUni & Me.Ope & Me.searchCrit
                            Else
                                strR = "NOT " & Me.fc.ColUni & Me.Ope & Me.searchCrit
                            End If

                        End If

                    End If

                ElseIf Me.fc.DataType = enDataType.Text And Me.fc.IsImplicitCast = False Then

                    'blank() 
                    If Me.BlankFunc = True Then

                        If Me.TrueFalse = False Then
                            If Me.Ope = "" Then
                                If Me.Is_Not = False Then
                                    strR = "Isblank(" & Me.fc.ColUni & ")=True"
                                Else
                                    strR = "not Isblank(" & Me.fc.ColUni & ")=True"
                                End If
                            Else
                                If Me.Is_Not = False Then
                                    If Me.Ope = "<>" Then
                                        strR = "Isblank(" & Me.fc.ColUni & ")=false"
                                    ElseIf Me.Ope = "=" Then
                                        strR = "Isblank(" & Me.fc.ColUni & ")=true"
                                    End If
                                Else
                                    If Me.Ope = "<>" Then
                                        strR = "NOT Isblank(" & Me.fc.ColUni & ")=false"
                                    ElseIf Me.Ope = "=" Then
                                        strR = "NOT Isblank(" & Me.fc.ColUni & ")=true"
                                    End If
                                End If
                            End If

                        Else
                            If Me.pTrueFalse = "true" Then
                                If Me.Is_Not = False Then
                                    If Me.Ope = "" Or Me.Ope = "=" Then
                                        strR = "Isblank(" & Me.fc.ColUni & ")=True"
                                    ElseIf Me.Ope = "<>" Then
                                        strR = "Isblank(" & Me.fc.ColUni & ")=false"
                                    End If
                                Else
                                    If Me.Ope = "" Or Me.Ope = "=" Then
                                        strR = "NOT Isblank(" & Me.fc.ColUni & ")=True"
                                    ElseIf Me.Ope = "<>" Then
                                        strR = "NOT Isblank(" & Me.fc.ColUni & ")=false"
                                    End If
                                End If
                            Else
                                If Me.Is_Not = False Then
                                    If Me.Ope = "" Or Me.Ope = "=" Then
                                        strR = "Isblank(" & Me.fc.ColUni & ")=false"
                                    ElseIf Me.Ope = "<>" Then
                                        strR = "Isblank(" & Me.fc.ColUni & ")=true"
                                    End If
                                Else
                                    If Me.Ope = "" Or Me.Ope = "=" Then
                                        strR = "NOT Isblank(" & Me.fc.ColUni & ")=false"
                                    ElseIf Me.Ope = "<>" Then
                                        strR = "NOT Isblank(" & Me.fc.ColUni & ")=true"
                                    End If
                                End If
                            End If


                        End If

                        'len()
                    ElseIf Me.LenFunc Then

                        If Me.Is_Not = False Then
                            strR = "LEN(" & Me.fc.ColUni & ")" & Me.Ope
                        Else
                            strR = "NOT LEN(" & Me.fc.ColUni & ")" & Me.Ope
                        End If

                        'Ap*le
                    ElseIf Me.Quotes = False And Me.Asteriks = False And Me.InBetAsteriks = True And Me.Ope = "" Then
                        If Me.Is_Not = False Then
                            Dim intInd As Integer = Me.searchCrit.IndexOf("*")
                            strR = "LEFT(" & Me.fc.ColUni & "," & intInd & ")=""" & Me.searchCrit.Substring(0, intInd) &
                                """ && RIGHT(" & Me.fc.ColUni & "," & (Me.searchCrit.Length - 1 - intInd) & ")=""" & Me.searchCrit.Substring(intInd + 1) & """"
                        Else
                            Dim intInd As Integer = Me.searchCrit.IndexOf("*")
                            strR = "NOT LEFT(" & Me.fc.ColUni & "," & intInd & ")=""" & Me.searchCrit.Substring(0, intInd) &
                                """ && RIGHT(" & Me.fc.ColUni & "," & (Me.searchCrit.Length - 1 - intInd) & ")=""" & Me.searchCrit.Substring(intInd + 1) & """"
                        End If

                        '"Ap*le"
                    ElseIf Me.Quotes = True And Me.Asteriks = False And Me.InBetAsteriks = True And Me.Ope = "" Then
                        If Me.Is_Not = False Then
                            Dim intInd As Integer = Me.searchCrit.IndexOf("*")
                            strR = "LEFT(" & Me.fc.ColUni & "," & intInd & ")=""" & Me.searchCrit.Substring(0, intInd) &
                                """ && RIGHT(" & Me.fc.ColUni & "," & (Me.searchCrit.Length - 1 - intInd) & ")=""" & Me.searchCrit.Substring(intInd + 1) & """"
                        Else
                            Dim intInd As Integer = Me.searchCrit.IndexOf("*")
                            strR = "NOT LEFT(" & Me.fc.ColUni & "," & intInd & ")=""" & Me.searchCrit.Substring(0, intInd) &
                                """ && RIGHT(" & Me.fc.ColUni & "," & (Me.searchCrit.Length - 1 - intInd) & ")=""" & Me.searchCrit.Substring(intInd + 1) & """"
                        End If

                        'Apple
                    ElseIf Me.Quotes = False And Me.Asteriks = False And Me.Ope = "" Then

                        If Me.Is_Not = False Then
                            strR = "CONTAINSSTRING(" & Me.fc.ColUni & ",""" & Me.searchCrit & """)"
                        Else
                            strR = "NOT CONTAINSSTRING(" & Me.fc.ColUni & ",""" & Me.searchCrit & """)"
                        End If

                        '*Apple*
                    ElseIf Me.Quotes = False And Me.Asteriks = True And Me.Ope = "" And Me.searchCrit.StartsWith("*") And Me.searchCrit.EndsWith("*") Then
                        If Me.Is_Not = False Then
                            strR = "CONTAINSSTRING(" & Me.fc.ColUni & ",""" & Me.searchCrit.Substring(1, Me.searchCrit.Length - 2) & """)"
                        Else
                            strR = "NOT CONTAINSSTRING(" & Me.fc.ColUni & ",""" & Me.searchCrit.Substring(1, Me.searchCrit.Length - 2) & """)"
                        End If

                        '"*Apple*"
                    ElseIf Me.Quotes = True And Me.Asteriks = True And Me.Ope = "" And Me.searchCrit.StartsWith("*") And Me.searchCrit.EndsWith("*") Then
                        If Me.Is_Not = False Then
                            strR = "CONTAINSSTRING(" & Me.fc.ColUni & ",""" & Me.searchCrit.Substring(1, Me.searchCrit.Length - 2) & """)"
                        Else
                            strR = "NOT CONTAINSSTRING(" & Me.fc.ColUni & ",""" & Me.searchCrit.Substring(1, Me.searchCrit.Length - 2) & """)"
                        End If

                        '="*Apple*"
                    ElseIf Me.Quotes = True And Me.Asteriks = True And Me.Ope <> "" And Me.searchCrit.StartsWith("*") And Me.searchCrit.EndsWith("*") Then
                        If Me.Is_Not = False And Me.Ope = "=" Then
                            strR = "CONTAINSSTRING(" & Me.fc.ColUni & ",""" & Me.searchCrit.Substring(1, Me.searchCrit.Length - 2) & """)"
                        ElseIf Me.Is_Not = True And Me.Ope = "=" Then
                            strR = "NOT CONTAINSSTRING(" & Me.fc.ColUni & ",""" & Me.searchCrit.Substring(1, Me.searchCrit.Length - 2) & """)"
                        ElseIf Me.Is_Not = False And Me.Ope = "<>" Then
                            strR = "NOT CONTAINSSTRING(" & Me.fc.ColUni & ",""" & Me.searchCrit.Substring(1, Me.searchCrit.Length - 2) & """)"
                        ElseIf Me.Is_Not = True And Me.Ope = "<>" Then
                            strR = "CONTAINSSTRING(" & Me.fc.ColUni & ",""" & Me.searchCrit.Substring(1, Me.searchCrit.Length - 2) & """)"
                        End If

                        'Apple*
                    ElseIf Me.Quotes = False And Me.Asteriks = True And Me.Ope = "" And Me.searchCrit.StartsWith("*") = False And Me.searchCrit.EndsWith("*") Then
                        If Me.Is_Not = False Then
                            strR = "LEFT(" & Me.fc.ColUni & "," & (Me.searchCrit.Length - 1).ToString & ")=""" & Me.searchCrit.Substring(0, Me.searchCrit.Length - 1) & """"
                        Else
                            strR = "NOT LEFT(" & Me.fc.ColUni & "," & (Me.searchCrit.Length - 1).ToString & ")=""" & Me.searchCrit.Substring(0, Me.searchCrit.Length - 1) & """"
                        End If

                        '"Apple*"
                    ElseIf Me.Quotes = True And Me.Asteriks = True And Me.Ope = "" And Me.searchCrit.StartsWith("*") = False And Me.searchCrit.EndsWith("*") Then
                        If Me.Is_Not = False Then
                            strR = "LEFT(" & Me.fc.ColUni & "," & (Me.searchCrit.Length - 1).ToString & ")=""" & Me.searchCrit.Substring(0, Me.searchCrit.Length - 1) & """"
                        Else
                            strR = "NOT LEFT(" & Me.fc.ColUni & "," & (Me.searchCrit.Length - 1).ToString & ")=""" & Me.searchCrit.Substring(0, Me.searchCrit.Length - 1) & """"
                        End If

                        '="Apple*"
                    ElseIf Me.Quotes = True And Me.Asteriks = True And Me.Ope <> "" And Me.searchCrit.StartsWith("*") = False And Me.searchCrit.EndsWith("*") Then
                        If Me.Is_Not = False Then
                            strR = "LEFT(" & Me.fc.ColUni & "," & (Me.searchCrit.Length - 1).ToString & ")" & Me.Ope & """" & Me.searchCrit.Substring(0, Me.searchCrit.Length - 1) & """"
                        Else
                            strR = "NOT LEFT(" & Me.fc.ColUni & "," & (Me.searchCrit.Length - 1).ToString & ")" & Me.Ope & """" & Me.searchCrit.Substring(0, Me.searchCrit.Length - 1) & """"
                        End If

                        '*Apple
                    ElseIf Me.Quotes = False And Me.Asteriks = True And Me.Ope = "" And Me.searchCrit.StartsWith("*") = True And Me.searchCrit.EndsWith("*") = False Then
                        If Me.Is_Not = False Then
                            strR = "RIGHT(" & Me.fc.ColUni & "," & (Me.searchCrit.Length - 1).ToString & ")=""" & Me.searchCrit.Substring(1, Me.searchCrit.Length - 1) & """"
                        Else
                            strR = "NOT RIGHT(" & Me.fc.ColUni & "," & (Me.searchCrit.Length - 1).ToString & ")=""" & Me.searchCrit.Substring(1, Me.searchCrit.Length - 1) & """"
                        End If

                        '"*Apple"
                    ElseIf Me.Quotes = True And Me.Asteriks = True And Me.Ope = "" And Me.searchCrit.StartsWith("*") = True And Me.searchCrit.EndsWith("*") = False Then
                        If Me.Is_Not = False Then
                            strR = "RIGHT(" & Me.fc.ColUni & "," & (Me.searchCrit.Length - 1).ToString & ")=""" & Me.searchCrit.Substring(1, Me.searchCrit.Length - 1) & """"
                        Else
                            strR = "NOT RIGHT(" & Me.fc.ColUni & "," & (Me.searchCrit.Length - 1).ToString & ")=""" & Me.searchCrit.Substring(1, Me.searchCrit.Length - 1) & """"
                        End If

                        '="*Apple"
                    ElseIf Me.Quotes = True And Me.Asteriks = True And Me.Ope <> "" And Me.searchCrit.StartsWith("*") = True And Me.searchCrit.EndsWith("*") = False Then
                        If Me.Is_Not = False Then
                            strR = "RIGHT(" & Me.fc.ColUni & "," & (Me.searchCrit.Length - 1).ToString & ")" & Me.Ope & """" & Me.searchCrit.Substring(1, Me.searchCrit.Length - 1) & """"
                        Else
                            strR = "NOT RIGHT(" & Me.fc.ColUni & "," & (Me.searchCrit.Length - 1).ToString & ")" & Me.Ope & """" & Me.searchCrit.Substring(1, Me.searchCrit.Length - 1) & """"
                        End If

                        '"Apple"
                    ElseIf Me.Quotes = True And Me.Asteriks = False And Me.Ope = "" Then
                        If Me.Is_Not = False Then
                            strR = Me.fc.ColUni & "=" & """" & Me.searchCrit & """"
                        Else
                            strR = "NOT " & Me.fc.ColUni & "=" & """" & Me.searchCrit & """"
                        End If

                        '=Apple
                    ElseIf Me.Quotes = False And Me.Asteriks = False And Me.Ope <> "" Then
                        If Me.Is_Not = False Then
                            strR = Me.fc.ColUni & Me.Ope & """" & Me.searchCrit & """"
                        Else
                            strR = "NOT " & Me.fc.ColUni & Me.Ope & """" & Me.searchCrit & """"
                        End If

                        '="Apple"
                    ElseIf Me.Quotes = True And Me.Asteriks = False And Me.Ope <> "" Then
                        If Me.Is_Not = False Then
                            strR = Me.fc.ColUni & Me.Ope & """" & Me.searchCrit & """"
                        Else
                            strR = "NOT " & Me.fc.ColUni & Me.Ope & """" & Me.searchCrit & """"
                        End If

                    End If

                ElseIf Me.fc.DataType = enDataType.DateTime And Me.fc.IsImplicitCast = False Then

                    If Me.BlankFunc = True Then

                        If Me.TrueFalse = False Then
                            If Me.Ope = "" Then
                                If Me.Is_Not = False Then
                                    strR = "Isblank(" & Me.fc.ColUni & ")=True"
                                Else
                                    strR = "not Isblank(" & Me.fc.ColUni & ")=True"
                                End If
                            Else
                                If Me.Is_Not = False Then
                                    If Me.Ope = "<>" Then
                                        strR = "Isblank(" & Me.fc.ColUni & ")=false"
                                    ElseIf Me.Ope = "=" Then
                                        strR = "Isblank(" & Me.fc.ColUni & ")=true"
                                    End If
                                Else
                                    If Me.Ope = "<>" Then
                                        strR = "NOT Isblank(" & Me.fc.ColUni & ")=false"
                                    ElseIf Me.Ope = "=" Then
                                        strR = "NOT Isblank(" & Me.fc.ColUni & ")=true"
                                    End If
                                End If
                            End If

                        Else
                            If Me.pTrueFalse = "true" Then
                                If Me.Is_Not = False Then
                                    If Me.Ope = "" Or Me.Ope = "=" Then
                                        strR = "Isblank(" & Me.fc.ColUni & ")=True"
                                    ElseIf Me.Ope = "<>" Then
                                        strR = "Isblank(" & Me.fc.ColUni & ")=false"
                                    End If
                                Else
                                    If Me.Ope = "" Or Me.Ope = "=" Then
                                        strR = "NOT Isblank(" & Me.fc.ColUni & ")=True"
                                    ElseIf Me.Ope = "<>" Then
                                        strR = "NOT Isblank(" & Me.fc.ColUni & ")=false"
                                    End If
                                End If
                            Else
                                If Me.Is_Not = False Then
                                    If Me.Ope = "" Or Me.Ope = "=" Then
                                        strR = "Isblank(" & Me.fc.ColUni & ")=false"
                                    ElseIf Me.Ope = "<>" Then
                                        strR = "Isblank(" & Me.fc.ColUni & ")=true"
                                    End If
                                Else
                                    If Me.Ope = "" Or Me.Ope = "=" Then
                                        strR = "NOT Isblank(" & Me.fc.ColUni & ")=false"
                                    ElseIf Me.Ope = "<>" Then
                                        strR = "NOT Isblank(" & Me.fc.ColUni & ")=true"
                                    End If
                                End If
                            End If

                        End If

                        'len()
                    ElseIf Me.LenFunc Then

                        If Me.Is_Not = False Then
                            strR = "LEN(" & Me.fc.ColUni & ")" & Me.Ope
                        Else
                            strR = "NOT LEN(" & Me.fc.ColUni & ")" & Me.Ope
                        End If

                        'today()
                    ElseIf Me.TodayFunc Then

                        If Me.TrueFalse = False Then
                            If Me.Ope = "" Then
                                If Me.Is_Not = False Then
                                    strR = Me.fc.ColUni & "=" & Me.searchCrit
                                Else
                                    strR = "NOT " & Me.fc.ColUni & "=" & Me.searchCrit
                                End If
                            ElseIf Me.Ope <> "" Then
                                If Me.Is_Not = False Then
                                    strR = Me.fc.ColUni & Me.Ope & Me.searchCrit
                                Else
                                    strR = "NOT " & Me.fc.ColUni & Me.Ope & Me.searchCrit
                                End If
                            End If
                        Else
                            If Me.Ope = "" Then
                                If Me.Is_Not = False Then
                                    strR = Me.fc.ColUni & "=Today()=" & Me.pTrueFalse
                                Else
                                    strR = "NOT " & Me.fc.ColUni & "=Today()=" & Me.pTrueFalse
                                End If
                            ElseIf Me.Ope <> "" Then
                                If Me.Is_Not = False Then
                                    strR = Me.fc.ColUni & Me.Ope & "Today()=" & Me.pTrueFalse
                                Else
                                    strR = "NOT " & Me.fc.ColUni & Me.Ope & "Today()=" & Me.pTrueFalse
                                End If
                            End If
                        End If

                    ElseIf Me.YearFunc Then

                        If Me.Is_Not = False Then
                            strR = "YEAR(" & Me.fc.ColUni & ")" & Me.Ope
                        Else
                            strR = "NOT YEAR(" & Me.fc.ColUni & ")" & Me.Ope
                        End If

                    ElseIf Me.MonthFunc Then

                        If Me.Is_Not = False Then
                            strR = "MONTH(" & Me.fc.ColUni & ")" & Me.Ope
                        Else
                            strR = "NOT MONTH(" & Me.fc.ColUni & ")" & Me.Ope
                        End If

                    ElseIf Me.DayFunc Then

                        If Me.Is_Not = False Then
                            strR = "DAY(" & Me.fc.ColUni & ")" & Me.Ope
                        Else
                            strR = "NOT DAY(" & Me.fc.ColUni & ")" & Me.Ope
                        End If

                    ElseIf Me.HourFunc Then

                        If Me.Is_Not = False Then
                            strR = "HOUR(" & Me.fc.ColUni & ")" & Me.Ope
                        Else
                            strR = "NOT HOUR(" & Me.fc.ColUni & ")" & Me.Ope
                        End If

                    ElseIf Me.MinuteFunc Then

                        If Me.Is_Not = False Then
                            strR = "MINUTE(" & Me.fc.ColUni & ")" & Me.Ope
                        Else
                            strR = "NOT MINUTE(" & Me.fc.ColUni & ")" & Me.Ope
                        End If

                    ElseIf Me.SecondFunc Then

                        If Me.Is_Not = False Then
                            strR = "SECOND(" & Me.fc.ColUni & ")" & Me.Ope
                        Else
                            strR = "NOT SECOND(" & Me.fc.ColUni & ")" & Me.Ope
                        End If

                    ElseIf Me.FourDigits Then

                        If Me.Ope = "" Then
                            If Me.Is_Not = False Then
                                strR = "YEAR(" & Me.fc.ColUni & ")=" & Me.searchCrit
                            Else
                                strR = "NOT YEAR(" & Me.fc.ColUni & ")=" & Me.searchCrit
                            End If
                        Else
                            If Me.Is_Not = False Then
                                strR = "YEAR(" & Me.fc.ColUni & ")" & Me.Ope & Me.searchCrit
                            Else
                                strR = "NOT YEAR(" & Me.fc.ColUni & ")" & Me.Ope & Me.searchCrit
                            End If
                        End If

                    ElseIf Me.SixDigits Then

                        If Me.Ope = "" Then
                            If Me.Is_Not = False Then
                                strR = "YEAR(" & Me.fc.ColUni & ")*100+MONTH(" & Me.fc.ColUni & ")=" & Me.searchCrit
                            Else
                                strR = "NOT YEAR(" & Me.fc.ColUni & ")*100+MONTH(" & Me.fc.ColUni & ")=" & Me.searchCrit
                            End If
                        Else
                            If Me.Is_Not = False Then
                                strR = "YEAR(" & Me.fc.ColUni & ")*100+MONTH(" & Me.fc.ColUni & ")" & Me.Ope & Me.searchCrit
                            Else
                                strR = "NOT YEAR(" & Me.fc.ColUni & ")*100+MONTH(" & Me.fc.ColUni & ")" & Me.Ope & Me.searchCrit
                            End If
                        End If

                    ElseIf Me.EightDigits Then

                        If Me.Ope = "" Then
                            If Me.Is_Not = False Then
                                strR = "YEAR(" & Me.fc.ColUni & ")*10000+MONTH(" & Me.fc.ColUni & ")*100+DAY(" & Me.fc.ColUni & ")=" & Me.searchCrit
                            Else
                                strR = "NOT YEAR(" & Me.fc.ColUni & ")*10000+MONTH(" & Me.fc.ColUni & ")*100+DAY(" & Me.fc.ColUni & ")=" & Me.searchCrit
                            End If
                        Else
                            If Me.Is_Not = False Then
                                strR = "YEAR(" & Me.fc.ColUni & ")*10000+MONTH(" & Me.fc.ColUni & ")*100+DAY(" & Me.fc.ColUni & ")" & Me.Ope & Me.searchCrit
                            Else
                                strR = "NOT YEAR(" & Me.fc.ColUni & ")*10000+MONTH(" & Me.fc.ColUni & ")*100+DAY(" & Me.fc.ColUni & ")" & Me.Ope & Me.searchCrit
                            End If
                        End If

                    ElseIf Me.IsDateTime = True Then

                        If Me.Ope = "" Then
                            If Me.Is_Not = False Then
                                strR = Me.fc.ColUni & "=CONVERT(""" & Me.searchCrit & """,DATETIME)"
                            Else
                                strR = "NOT " & Me.fc.ColUni & "=CONVERT(""" & Me.searchCrit & """,DATETIME)"
                            End If
                        Else
                            If Me.Is_Not = False Then
                                strR = Me.fc.ColUni & Me.Ope & "CONVERT(""" & Me.searchCrit & """,DATETIME)"
                            Else
                                strR = "NOT " & Me.fc.ColUni & Me.Ope & "CONVERT(""" & Me.searchCrit & """,DATETIME)"
                            End If
                        End If

                    End If

                ElseIf Me.fc.DataType = enDataType.Number Or Me.fc.IsImplicitCast = True Then

                    If Me.BlankFunc = True Then

                        If Me.TrueFalse = False Then
                            If Me.Ope = "" Then
                                If Me.Is_Not = False Then
                                    strR = "Isblank(" & Me.fc.ColUni & ")=True"
                                Else
                                    strR = "not Isblank(" & Me.fc.ColUni & ")=True"
                                End If
                            Else
                                If Me.Is_Not = False Then
                                    If Me.Ope = "<>" Then
                                        strR = "Isblank(" & Me.fc.ColUni & ")=false"
                                    ElseIf Me.Ope = "=" Then
                                        strR = "Isblank(" & Me.fc.ColUni & ")=true"
                                    End If
                                Else
                                    If Me.Ope = "<>" Then
                                        strR = "NOT Isblank(" & Me.fc.ColUni & ")=false"
                                    ElseIf Me.Ope = "=" Then
                                        strR = "NOT Isblank(" & Me.fc.ColUni & ")=true"
                                    End If
                                End If
                            End If

                        Else
                            If Me.pTrueFalse = "true" Then
                                If Me.Is_Not = False Then
                                    If Me.Ope = "" Or Me.Ope = "=" Then
                                        strR = "Isblank(" & Me.fc.ColUni & ")=True"
                                    ElseIf Me.Ope = "<>" Then
                                        strR = "Isblank(" & Me.fc.ColUni & ")=false"
                                    End If
                                Else
                                    If Me.Ope = "" Or Me.Ope = "=" Then
                                        strR = "NOT Isblank(" & Me.fc.ColUni & ")=True"
                                    ElseIf Me.Ope = "<>" Then
                                        strR = "NOT Isblank(" & Me.fc.ColUni & ")=false"
                                    End If
                                End If
                            Else
                                If Me.Is_Not = False Then
                                    If Me.Ope = "" Or Me.Ope = "=" Then
                                        strR = "Isblank(" & Me.fc.ColUni & ")=false"
                                    ElseIf Me.Ope = "<>" Then
                                        strR = "Isblank(" & Me.fc.ColUni & ")=true"
                                    End If
                                Else
                                    If Me.Ope = "" Or Me.Ope = "=" Then
                                        strR = "NOT Isblank(" & Me.fc.ColUni & ")=false"
                                    ElseIf Me.Ope = "<>" Then
                                        strR = "NOT Isblank(" & Me.fc.ColUni & ")=true"
                                    End If
                                End If
                            End If

                        End If

                        'len()
                    ElseIf Me.LenFunc Then

                        If Me.Is_Not = False Then
                            strR = "LEN(" & Me.fc.ColUni & ")" & Me.Ope
                        Else
                            strR = "NOT LEN(" & Me.fc.ColUni & ")" & Me.Ope
                        End If

                    Else
                        If Me.Is_Not = False Then
                            If Me.Ope = "" Then
                                strR = Me.fc.ColUni & "=" & Me.searchCrit
                            Else
                                strR = Me.fc.ColUni & Me.Ope & Me.searchCrit
                            End If
                        Else
                            If Me.Ope = "" Then
                                strR = "NOT " & Me.fc.ColUni & "=" & Me.searchCrit
                            Else
                                strR = "NOT " & Me.fc.ColUni & Me.Ope & Me.searchCrit
                            End If
                        End If

                    End If

                End If


                Return " " & strR & " "
            End Get
        End Property




        Public Sub New()

        End Sub





    End Class

    Friend Function DaxStmnt() As String

        Dim strR As String = ""
        Dim strSort As String = ""

        If Me.Sort = enSort.none OrElse Me.Sort = enSort.asc Then
            strSort = "asc"
        Else
            strSort = "desc"
        End If


        If Me.txtBox.Text.Trim = "" AndAlso Me.ContextType = enContextType.LevelNon Then
            Me.DaxFilter = ""
            strR = "DEFINE " & vbCrLf
            strR += "var x0 = VALUES(" & Me.ColUni & ")" & vbCrLf
            strR += "EVALUATE" & Chr(13) & Chr(10) & vbCrLf
            strR += "TOPN(1001,x0," & Me.ColUni & "," & strSort & ") order by " & Me.ColUni & " " & strSort
            Return strR
        ElseIf Me.txtBox.Text.Trim = "" AndAlso Me.ContextType = enContextType.MeasureNon Then
            strR = "EVALUATE ROW(""" & Me.Column & """," & Me.ColUni & ")"
            Return strR
        End If


        Dim strSearch As String = Me.txtBox.Text

        Dim xE As Object = xExplicit_in(strSearch) : Dim htE As Hashtable = xE(0) : strSearch = xE(1)

        strSearch = xCleanSearchstring(strSearch)

        Dim strA() = strSearch.Split(New String() {"||", "&&"}, StringSplitOptions.None)
        Dim intL As Integer = 0
        Dim ctr As Integer = 0
        Dim lstCp As New List(Of ColumnP)
        For Each s In strA
            ctr += 1
            Dim p As String = s
            p = Me.xRemovePar(p)
            Dim cp As New ColumnP With {.Search = strSearch, .P = p, .StartPos = Me.getStart(p) + intL, .EndPos = Me.getEnd(p) + intL, .Pos = ctr, .fc = Me}
            lstCp.Add(cp)
            intL += p.Length + 2
        Next s

        For Each cp As ColumnP In lstCp
            strSearch = strSearch.Replace(" " & cp.P.Trim & " ", cp.Dax)
        Next cp

        strSearch = xExplicit_out(strSearch, htE)

        If strSearch.Trim = "" Then
            Me.DaxFilter = ""
        Else
            Me.DaxFilter = strSearch
        End If


        If Me.FieldType = enFieldType.Level Then

            If Me.ContextType = enContextType.LevelNon Then

                strR = "DEFINE" & vbCrLf
                strR += "var x0 = VALUES(" & Me.ColUni & ")"
                strR += "var x1 = FILTER(x0," & Me.DaxFilter & ")"
                strR += "EVALUATE" & vbCrLf
                strR += "TOPN(1001,x1," & Me.ColUni & "," & strSort & ") order by " & Me.ColUni & " " & strSort

            ElseIf Me.ContextType = enContextType.LevelLevel Then

                strR = "DEFINE" & vbCrLf
                Dim ctrF As Integer = 0
                For Each c As Context In Me.ContextLevelsFiltered
                    ctrF += 1
                    strR += "var f" & ctrF.ToString & " = FILTER(" & c.TableName & "," & c.FilterExpression & ")" & vbCrLf
                Next c

                strR += "var x0 = VALUES(" & Me.ColUni & ")" & vbCrLf
                If Me.DaxFilter <> "" Then
                    strR += "var x1 = FILTER(x0," & Me.DaxFilter & ")" & vbCrLf
                    strR += "var x2 = SUMMARIZECOLUMNS(" & Me.ColUni & ",x1,"
                Else
                    strR += "var x2 = SUMMARIZECOLUMNS(" & Me.ColUni & ",x0,"
                End If
                For i As Integer = 1 To ctrF
                    strR += "f" & i.ToString & ","
                Next i
                strR = strR.Substring(0, strR.Length - 1) & ")" & vbCrLf
                strR += "EVALUATE" & vbCrLf
                strR += "TOPN(1001,x2," & Me.ColUni & "," & strSort & ") order by " & Me.ColUni & " " & strSort

            ElseIf Me.ContextType = enContextType.LevelMeasure Then

                strR = "DEFINE" & vbCrLf

                strR += "var x0 = SUMMARIZECOLUMNS(" & Me.ColUni & ","
                For Each c As Context In Me.ContextLevels
                    strR += c.TableName & "[" & c.FieldName & "],"
                Next c
                strR = strR.Substring(0, strR.Length - 1) & ")" & vbCrLf

                strR += "var m0 = FILTER(x0,"
                For Each c As Context In Me.ContextMeasuresFiltered
                    strR += c.FilterExpression & " && "
                Next c
                strR = strR.Substring(0, strR.Length - 4) & ")" & vbCrLf

                strR += "var s0 = Distinct(SELECTCOLUMNS(m0,""" & Me.Column & """," & Me.ColUni & "))" & Chr(13) & Chr(10)

                strR += "EVALUATE" & Chr(13) & Chr(10)
                strR += "TOPN(1001,s0,[" & Me.Column & "]," & strSort & ") order by [" & Me.Column & "] " & strSort



            ElseIf Me.ContextType = enContextType.LevelLevelMeasure Then

                strR = "DEFINE" & vbCrLf
                Dim ctrF As Integer = 0
                For Each c As Context In Me.ContextLevelsFiltered
                    ctrF += 1
                    strR += "var f" & ctrF.ToString & " = FILTER(" & c.TableName & "," & c.FilterExpression & ")" & vbCrLf
                Next c

                strR += "var x0 = SUMMARIZECOLUMNS(" & Me.ColUni & ","
                For Each c As Context In Me.ContextLevels
                    strR += c.TableName & "[" & c.FieldName & "],"
                Next c
                For i As Integer = 1 To ctrF
                    strR += "f" & i.ToString & ","
                Next i
                strR = strR.Substring(0, strR.Length - 1) & ")" & vbCrLf

                strR += "var m0 = FILTER(x0,"
                For Each c As Context In Me.ContextMeasuresFiltered
                    strR += c.FilterExpression & " && "
                Next c
                strR = strR.Substring(0, strR.Length - 4) & ")" & vbCrLf

                strR += "var s0 = Distinct(SELECTCOLUMNS(m0,""" & Me.Column & """," & Me.ColUni & "))" & vbCrLf

                strR += "EVALUATE" & vbCrLf
                strR += "TOPN(1001,s0,[" & Me.Column & "]," & strSort & ") order by [" & Me.Column & "] " & strSort

            End If

        ElseIf Me.FieldType = enFieldType.Measure Then

            If Me.ContextType = enContextType.MeasureNon Then

                strR = "EVALUATE ROW(""" & Me.Column & """," & Me.ColUni & ")"

            ElseIf Me.ContextType = enContextType.MeasureLevel Then

                strR = "DEFINE" & Chr(13) & Chr(10)
                Dim ctrF As Integer = 0
                For Each c As Context In Me.ContextLevelsFiltered
                    ctrF += 1
                    strR += "var f" & ctrF.ToString & " = FILTER(" & c.TableName & "," & c.FilterExpression & ")" & vbCrLf
                Next c
                strR += "var c0 = SUMMARIZECOLUMNS("
                For Each c As Context In Me.ContextLevels
                    strR += c.TableName & "[" & c.FieldName & "],"
                Next c
                If ctrF > 0 Then
                    For i As Integer = 1 To ctrF
                        strR += "f" & i.ToString & ","
                    Next i
                End If
                strR += """" & Me.Column & """,IGNORE([" & Me.Column & "]))" & vbCrLf

                If Me.DaxFilter <> "" Then
                    strR += "var m0 = FILTER(c0," & Me.DaxFilter & ")" & Chr(13) & Chr(10)
                    strR += "var x1 = SELECTCOLUMNS(m0,""" & Me.Column & """,[" & Me.Column & "])" & vbCrLf
                Else
                    strR += "var x1 = SELECTCOLUMNS(c0,""" & Me.Column & """,[" & Me.Column & "])" & vbCrLf
                End If

                strR += "EVALUATE TOPN(1001,DISTINCT(x1),[" & Me.Column & "]," & strSort & ") order by [" & Me.Column & "] " & strSort

            ElseIf Me.ContextType = enContextType.MeasureMeasure Then

                strR = "EVALUATE ROW(""" & Me.Column & """," & Me.ColUni & ")"

            ElseIf Me.ContextType = enContextType.MeasureLevelMeasure Then

                strR = "DEFINE" & Chr(13) & Chr(10)
                Dim ctrF As Integer = 0
                For Each c As Context In Me.ContextLevelsFiltered
                    ctrF += 1
                    strR += "var f" & ctrF.ToString & " = FILTER(" & c.TableName & "," & c.FilterExpression & ")" & vbCrLf
                Next c
                strR += "var c0 = SUMMARIZECOLUMNS("
                For Each c As Context In Me.ContextLevels
                    strR += c.TableName & "[" & c.FieldName & "],"
                Next c
                If ctrF > 0 Then
                    For i As Integer = 1 To ctrF
                        strR += "f" & i.ToString & ","
                    Next i
                End If
                strR += """" & Me.Column & """,IGNORE([" & Me.Column & "]))" & vbCrLf
                strR += "var m0 = FILTER(c0,"
                If Me.DaxFilter <> "" Then
                    strR += Me.DaxFilter & " && "
                End If
                For Each c As Context In Me.ContextMeasuresFiltered
                    strR += c.FilterExpression & " && "
                Next c
                strR = strR.Substring(0, strR.Length - 4) & ")" & vbCrLf

                strR += "var x1 = SELECTCOLUMNS(m0,""" & Me.Column & """,[" & Me.Column & "])" & vbCrLf
                strR += "EVALUATE TOPN(1001,DISTINCT(x1),[" & Me.Column & "]," & strSort & ") order by [" & Me.Column & "] " & strSort


            End If

        End If


        Return strR


    End Function

    Private _ContextExt As List(Of Context)
    Public Property ContextExt As List(Of Context)
        Get
            Return _ContextExt
        End Get
        Set(value As List(Of Context))
            _ContextExt = value
        End Set
    End Property

    Private ReadOnly Property ContextLevelsFiltered As List(Of Context)
        Get
            Dim lstRes As New List(Of Context)
            For Each c As Context In Me.ContextExt
                If c.FieldType = enFieldType.Level AndAlso c.FilterExpression.Trim <> "" Then
                    lstRes.Add(c)
                End If
            Next c
            Return lstRes
        End Get
    End Property

    Private ReadOnly Property ContextLevels As List(Of Context)
        Get
            Dim lstRes As New List(Of Context)
            For Each c As Context In Me.ContextExt
                If c.FieldType = enFieldType.Level Then
                    lstRes.Add(c)
                End If
            Next c
            Return lstRes
        End Get
    End Property


    Private ReadOnly Property ContextMeasuresFiltered As List(Of Context)
        Get
            Dim lstRes As New List(Of Context)
            For Each c As Context In Me.ContextExt
                If c.FieldType = enFieldType.Measure AndAlso c.FilterExpression.Trim <> "" Then
                    lstRes.Add(c)
                End If
            Next c
            Return lstRes
        End Get
    End Property


    Private Enum enContextType
        MeasureNon = 1
        MeasureLevel = 2
        MeasureMeasure = 3
        MeasureLevelMeasure = 4
        LevelNon = 5
        LevelLevel = 6
        LevelMeasure = 7
        LevelLevelMeasure = 8
    End Enum

    Private ReadOnly Property ContextType As enContextType
        Get
            Dim blnMea As Boolean = False
            Dim blnLvl As Boolean = False

            If Me.FieldType = enFieldType.Level Then
                If Not Me.ContextExt Is Nothing Then
                    For Each c As Context In Me.ContextExt
                        If c.FieldType = enFieldType.Measure AndAlso c.FilterExpression.Trim <> "" Then
                            blnMea = True
                        ElseIf c.FieldType = enFieldType.Level AndAlso c.FilterExpression.Trim <> "" Then
                            blnLvl = True
                        End If
                    Next
                End If
                If Me.ContextExt Is Nothing Then
                    Return enContextType.LevelNon
                ElseIf blnMea = False And blnLvl = True Then
                    Return enContextType.LevelLevel
                ElseIf blnMea = True And blnLvl = False Then
                    Return enContextType.LevelMeasure
                ElseIf blnMea = True And blnLvl = True Then
                    Return enContextType.LevelLevelMeasure
                Else
                    Return enContextType.LevelNon
                End If
            End If

            If Me.FieldType = enFieldType.Measure Then
                If Me.ContextExt Is Nothing Then
                    Return enContextType.MeasureNon
                Else
                    For Each c As Context In Me.ContextExt
                        If c.FieldType = enFieldType.Level Then
                            blnLvl = True
                        ElseIf c.FieldType = enFieldType.Measure And c.FilterExpression.Trim <> "" Then
                            blnMea = True
                        End If
                    Next c
                    If blnLvl = True And blnMea = False Then
                        Return enContextType.MeasureLevel
                    ElseIf blnLvl = False And blnMea = True Then
                        Return enContextType.MeasureMeasure
                    ElseIf blnLvl = True And blnMea = True Then
                        Return enContextType.MeasureLevelMeasure
                    Else
                        Return enContextType.MeasureNon
                    End If
                End If
            End If

            Return Nothing

        End Get
    End Property

    Private ReadOnly Property ColUni As String
        Get
            If Me.FieldType = enFieldType.Level Then
                Return "'" & Me.Table & "'[" & Me.Column & "]"
            Else
                Return "[" & Me.Column & "]"
            End If

        End Get
    End Property

    Private Function xCleanSearchstring(strSearch As String) As String

        Dim strT As String = ""
        Dim intStart As Integer = strSearch.IndexOf("""")
        Dim intEnd As Integer = strSearch.LastIndexOf("""")
        Dim strGUID As String = System.Guid.NewGuid.ToString

        If (intStart <> intEnd) And intStart > -1 And intEnd > -1 Then
            strT = strSearch.Substring(intStart + 1, intEnd - intStart - 1)
            If strT.Length > 0 Then
                strSearch = strSearch.Replace("""" & strT & """", strGUID)
            End If
        End If

        strSearch = xClean(strSearch, "  ", " ")

        If (intStart <> intEnd) And intStart > -1 And intEnd > -1 Then
            strSearch = strSearch.Replace(strGUID, """" & strT & """")
        End If



        strSearch = xReplace(strSearch, " and and ", " && and ")
        strSearch = xReplace(strSearch, " and or", " && or ")
        strSearch = xReplace(strSearch, " or or", " || or ")
        strSearch = xReplace(strSearch, " or and", " || and ")
        strSearch = xReplace(strSearch, " and ", " && ")
        strSearch = xReplace(strSearch, ")and ", ") && ")
        strSearch = xReplace(strSearch, "and(", " && (")
        strSearch = xReplace(strSearch, " or ", " || ")
        strSearch = xReplace(strSearch, ")or ", ") || ")
        strSearch = xReplace(strSearch, " or(", " || (")

        If strSearch.EndsWith(")") = True AndAlso strSearch.EndsWith(" )") = False Then
            strSearch = strSearch.Substring(0, strSearch.Length - 1) & " )"
        End If
        For i As Integer = strSearch.Length - 1 To 0 Step -1
            If strSearch.Substring(i, 1) = ")" Then
                If strSearch.Substring(i - 1, 1) <> " " Then
                    strSearch = strSearch.Substring(0, i) & " " & strSearch.Substring(i)
                End If
            End If
        Next i

        If strSearch.StartsWith("(") = True AndAlso strSearch.StartsWith("( ") = False Then
            strSearch = strSearch.Substring(0, 1) & " " & strSearch.Substring(1)
        End If
        For i As Integer = strSearch.Length - 1 To 0 Step -1
            If strSearch.Substring(i, 1) = "(" Then
                If strSearch.Substring(i + 1, 1) <> " " Then
                    strSearch = strSearch.Substring(0, i + 1) & " " & strSearch.Substring(i + 1)
                End If
            End If
        Next i

        If strSearch.StartsWith(" ") = False Then strSearch = " " & strSearch
        If strSearch.EndsWith(" ") = False Then strSearch = strSearch & " "
        strSearch = Me.xCompFunc(strSearch, "blank( )", "blank()")
        strSearch = Me.xCompFunc(strSearch, "blank ( )", "blank()")
        strSearch = Me.xCompFunc(strSearch, "len( )", "len()")
        strSearch = Me.xCompFunc(strSearch, "len ( )", "len()")
        strSearch = Me.xCompFunc(strSearch, "len() =", "len()=")
        strSearch = Me.xCompFunc(strSearch, "= ", "=")
        strSearch = Me.xCompFunc(strSearch, "> ", ">")
        strSearch = Me.xCompFunc(strSearch, "< ", "<")
        strSearch = Me.xCompFunc(strSearch, ">= ", ">=")
        strSearch = Me.xCompFunc(strSearch, "<= ", "<=")
        strSearch = Me.xCompFunc(strSearch, "<> ", "<>")
        strSearch = Me.xCompFunc(strSearch, "now( )", "now()")
        strSearch = Me.xCompFunc(strSearch, "now ( )", "now()")
        strSearch = Me.xCompFunc(strSearch, "today( )", "today()")
        strSearch = Me.xCompFunc(strSearch, "today ( )", "today()")
        strSearch = Me.xCompFunc(strSearch, "year( )", "year()")
        strSearch = Me.xCompFunc(strSearch, "year ( )", "year()")
        strSearch = Me.xCompFunc(strSearch, "month( )", "month()")
        strSearch = Me.xCompFunc(strSearch, "month ( )", "month()")
        strSearch = Me.xCompFunc(strSearch, "day( )", "day()")
        strSearch = Me.xCompFunc(strSearch, "day ( )", "day()")
        strSearch = Me.xCompFunc(strSearch, "hour( )", "hour()")
        strSearch = Me.xCompFunc(strSearch, "hour ( )", "hour()")
        strSearch = Me.xCompFunc(strSearch, "minute( )", "minute()")
        strSearch = Me.xCompFunc(strSearch, "minute ( )", "minute()")
        strSearch = Me.xCompFunc(strSearch, "second( )", "second()")
        strSearch = Me.xCompFunc(strSearch, "second ( )", "second()")

        strSearch = Me.xCompFunc(strSearch, "not>", "not >")
        strSearch = Me.xCompFunc(strSearch, "not=", "not =")
        strSearch = Me.xCompFunc(strSearch, "not<", "not <")

        Return strSearch

    End Function

    Private Function xExplicit_in(strSearch As String) As Object

        Dim htE As Hashtable = Nothing
        Dim Ptn As String = "\"".*?"""
        For Each m As System.Text.RegularExpressions.Match In System.Text.RegularExpressions.Regex.Matches(strSearch, Ptn)
            If htE Is Nothing Then htE = New Hashtable
            Dim strM As String = m.ToString
            If htE.ContainsValue(strM) = False Then
                Dim strGUID As String = """" & System.Guid.NewGuid.ToString.Substring(0, 8) & """"
                htE.Add(strGUID, strM)
                strSearch = strSearch.Replace(strM, strGUID)
            End If
        Next m

        Return {htE, strSearch}

    End Function

    Private Function xExplicit_out(strS As String, htE As Hashtable) As String
        Dim strSearch As String = strS
        If Not htE Is Nothing Then
            For Each _key As DictionaryEntry In htE
                Dim strT As String = Me.ColUni & "=" & _key.Key

                If _key.Value.ToString.StartsWith("""*") = False And _key.Value.ToString.EndsWith("*""") = False And _key.Value.ToString.Contains("*") = True Then

                    Dim strVL As String = _key.Value.ToString.Replace("""", "")
                    strVL = """" & strVL.Substring(0, strVL.IndexOf("*")) & """"
                    Dim strVR As String = _key.Value.ToString.Replace("""", "")
                    strVR = """" & strVR.Substring(strVR.IndexOf("*") + 1) & """"
                    Dim strN As String = " LEFT(" & Me.ColUni & "," & strVL.Length - 2 & ")=" & strVL & " && RIGHT(" & Me.ColUni & "," & strVR.Length - 2 & ")=" & strVR
                    strSearch = strSearch.Replace(strT, strN)
                ElseIf _key.Value.ToString.StartsWith("""*") = False And _key.Value.ToString.EndsWith("*""") = False Then
                    strSearch = strSearch.Replace(_key.Key, _key.Value)
                ElseIf _key.Value.ToString.StartsWith("""*") = True And _key.Value.ToString.EndsWith("*""") = True Then
                    Dim strV As String = """" & _key.Value.ToString.Replace("""*", "").Replace("*""", "") & """"
                    Dim strN As String = " CONTAINSSTRING(" & Me.ColUni & "," & strV & ")"
                    strSearch = strSearch.Replace(strT, strN)
                ElseIf _key.Value.ToString.StartsWith("""*") = True And _key.Value.ToString.EndsWith("*""") = False Then
                    Dim strV As String = """" & _key.Value.ToString.Replace("""*", "")
                    Dim strN As String = " RIGHT(" & Me.ColUni & "," & strV.Length - 2.ToString & ")=" & strV
                    strSearch = strSearch.Replace(strT, strN)
                ElseIf _key.Value.ToString.StartsWith("""*") = False And _key.Value.ToString.EndsWith("*""") = True Then
                    Dim strV As String = _key.Value.ToString.Replace("*""", "") & """"
                    Dim strN As String = " LEFT(" & Me.ColUni & "," & strV.Length - 2.ToString & ")=" & strV
                    strSearch = strSearch.Replace(strT, strN)
                End If

            Next _key
            For Each _key As DictionaryEntry In htE
                strSearch = strSearch.Replace(_key.Key, _key.Value)
            Next _key
        End If

        Return strSearch

    End Function

    Private Function getStart(p As String) As Integer
        Dim intRes As Integer = 0
        For i As Integer = 0 To p.Length - 1
            If p.Substring(i, 1) = " " Then
                intRes += 1
            Else
                Exit For
            End If
        Next i
        Return intRes
    End Function

    Private Function getEnd(p As String) As Integer
        Dim intRes As Integer = p.Length
        For i As Integer = p.Length - 1 To 0 Step -1
            If p.Substring(i, 1) = " " Then
                intRes -= 1
            Else
                intRes -= 1
                Exit For
            End If
        Next i
        Return intRes
    End Function

    Private Function xClean(Result As String, opA As String, opB As String) As String

        If Result.Contains(Chr(13)) Then Result = Result.Replace(Chr(13), " ")
        If Result.Contains(Chr(10)) Then Result = Result.Replace(Chr(10), " ")

        For i As Integer = Result.Length - 1 To 0 Step -1
            If i + opA.Length <= Result.Length Then
                If Result.Substring(i, opA.Length).ToString.ToLower = opA.ToLower Then
                    Result = Result.Substring(0, i) & opB & Result.Substring(i + Len(opA))
                End If
            End If
        Next i
        If opB = " " AndAlso Result.EndsWith(" ") Then Result = Result.Substring(0, Result.Length - 1)
        If opB = " " AndAlso Result.StartsWith(" ") Then Result = Result.Substring(1, Result.Length - 1)


        Return Result
    End Function

    Private Function xReplace(Result As String, opA As String, opB As String) As String


        If Result = "" Then Return ""
        If Result.ToLower.Contains(opA.ToLower) = False Then
            Return Result
        End If

        For i As Integer = Result.Length - 1 To 0 Step -1
            If i + opA.Length <= Result.Length Then
                If Result.Substring(i, opA.Length).ToString.ToLower = opA.ToLower Then
                    Result = Result.Substring(0, i) & opB & Result.Substring(i + Len(opA))
                End If
            End If
        Next i

        Return Result
    End Function

    Private Function xCompFunc(strSearch As String, strFx As String, strRx As String) As String

        Dim strS As String = strSearch

        Do While strS.ToLower.Contains(strFx.ToLower)
            Dim intIndex As Integer = strS.IndexOf(strFx, StringComparison.CurrentCultureIgnoreCase)
            If intIndex = -1 Then
                Exit Do
            End If

            Dim strS1 As String = strS.Substring(0, intIndex) & strRx
            Dim strS2 As String = strS.Substring(intIndex + strFx.Length)
            strS = strS1 & strS2

        Loop

        Return strS

    End Function

    Private Function xRemovePar(p As String) As String


        If p.Contains("day()=<") Then
            Dim x As String = ""
        End If

        Dim ht As New Hashtable
        Dim strToken() As String = {"today(", "today (", "len(", "len (", "blank(", "blank (", "now(", "now (", "today(",
            "today (", "last(", "last (", "next(", "next (", "year(", "year (", "month(", "month (", "day(", "day (",
            "hour(", "hour (", "minute(", "minute (", "second(", "second ("}

        For Each t As String In strToken
            Do While p.IndexOf(t, StringComparison.CurrentCultureIgnoreCase) > 0
                Dim intStart As Integer = p.IndexOf(t, StringComparison.CurrentCultureIgnoreCase)
                Dim intEnd As Integer = -1

                For i As Integer = intStart + (t.Length - 1) + 1 To p.Length - 1
                    If p.Substring(i, 1) = ")" Then
                        intEnd = i
                        Exit For
                    End If
                Next i


                Dim strV As String = ""
                Dim strGUID As String = System.Guid.NewGuid.ToString
                If intEnd > intStart Then
                    strV = p.Substring(intStart, intEnd - intStart + 1)
                Else
                    strV = p.Substring(intStart, t.Length)
                End If
                ht.Add(strGUID, strV)
                p = p.Replace(strV, strGUID)
            Loop
        Next t

        p = p.Replace("(", " ")
        p = p.Replace(")", " ")


        For Each x In ht
            p = p.Replace(x.key, x.value)
        Next

        Return p

    End Function

    Private Sub ctrlFilter_Resize(sender As Object, e As EventArgs) Handles Me.Resize
        'If Not Me.lblCaption Is Nothing Then
        '    Me.lblCaption.Width = Me.Width - 30 - 24
        'End If
    End Sub

    Public Class Context

        Private _TableName As String
        Public Property TableName As String
            Get
                If Me._TableName.StartsWith("'") = False And Me._TableName.EndsWith("'") = False Then
                    Return "'" & Me._TableName & "'"
                Else
                    Return Me._TableName
                End If

            End Get
            Set(value As String)
                Me._TableName = value
            End Set
        End Property



        Public Property FieldName As String
        Public Property DataType As enDataType
        Public Property FieldType As enFieldType

        Private _FilterExpression As String
        Public Property FilterExpression As String
            Get
                If Me._FilterExpression Is Nothing Then
                    Return ""
                Else
                    Return Me._FilterExpression
                End If
            End Get
            Set(value As String)
                If Not value Is Nothing Then
                    Me._FilterExpression = value
                Else
                    Me._FilterExpression = ""
                End If

            End Set
        End Property


        Public Sub New()

        End Sub

    End Class

End Class
