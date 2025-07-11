

Public Class ctrlColumnHeader

    Friend ctrlTable As ctrlTable
    Friend btn As Button
    Friend lbl As Label
    Friend ColUniName As String

    Private qc As clsQueryColumn


    Public Sub New(ctrlTable As ctrlTable, qryCol As clsQueryColumn)

        InitializeComponent()

        If qryCol Is Nothing Then
            Exit Sub
            Me.Dispose()
        End If

        Me.qc = qryCol



        Me.ctrlTable = ctrlTable
        Me.ColUniName = Me.qc.UniName

        Me.btn = New Button With {.BackColor = Color.AliceBlue, .Top = 0,
            .Left = Me.Width - 16, .Width = 14, .Height = 14, .Anchor = 9,
            .Image = Me.ctrlTable.ctrlDaxQuery.ImageList.Images("qe_default.ico"),
            .ImageAlign = ContentAlignment.MiddleCenter, .FlatStyle = FlatStyle.Standard}
        If Me.qc.IsFiltered = True Then
            If Me.qc.Sort = clsQueryColumn.enSort.none Then
                Me.btn.Image = Me.ctrlTable.ctrlDaxQuery.ImageList.Images("qe_filter.ico")
            ElseIf Me.qc.Sort = clsQueryColumn.enSort.asc Then
                Me.btn.Image = Me.ctrlTable.ctrlDaxQuery.ImageList.Images("qe_filter_asc.ico")
            ElseIf Me.qc.Sort = clsQueryColumn.enSort.desc Then
                Me.btn.Image = Me.ctrlTable.ctrlDaxQuery.ImageList.Images("qe_filter_desc.ico")
            End If
        Else
            If Me.qc.Sort = clsQueryColumn.enSort.asc Then
                Me.btn.Image = Me.ctrlTable.ctrlDaxQuery.ImageList.Images("qe_default_asc.ico")
            ElseIf Me.qc.Sort = clsQueryColumn.enSort.desc Then
                Me.btn.Image = Me.ctrlTable.ctrlDaxQuery.ImageList.Images("qe_default_desc.ico")
            Else
                Me.btn.Image = Me.ctrlTable.ctrlDaxQuery.ImageList.Images("qe_default.ico")
            End If
        End If


        Me.btn.FlatStyle = FlatStyle.Flat
        Me.btn.FlatAppearance.BorderSize = 0
        Me.btn.FlatAppearance.MouseOverBackColor = Color.White
        Me.btn.FlatAppearance.MouseDownBackColor = Color.White

        AddHandler Me.btn.MouseMove, AddressOf Me.btn_MouseMove
        AddHandler Me.btn.Click, AddressOf Me.btn_Click

        Me.Controls.Add(btn)


        Dim strCaption As String = ""
        If Me.qc.FieldType = clsQueryColumn.enFieldType.Measure Then
            strCaption = Me.ctrlTable.ctrlDaxQuery.tm.GetMeasure(Me.qc.UniName).Caption
        ElseIf Me.qc.FieldType = clsQueryColumn.enFieldType.Level Then
            strCaption = Me.ctrlTable.ctrlDaxQuery.tm.GetLevel(Me.qc.UniName).Caption
        ElseIf Me.qc.FieldType = clsQueryColumn.enFieldType.ImpMeasure Then
            strCaption = Me.ctrlTable.ctrlDaxQuery.tm.GetLevel(Me.qc.UniName).Caption & "(" & Me.qc.iFunction.ToString & ")"
        End If

        Me.lbl = New Label With {.Text = strCaption, .Top = 0, .Left = 2 + 16 - 2, .Width = Me.Width - 12 - 16 + 2, .AutoEllipsis = True, .Anchor = 13}


        Dim pb As New PictureBox With {.Left = 0, .Top = 0, .Width = 16, .Height = 16}
        If Me.qc.FieldType = clsQueryColumn.enFieldType.Level Then
            pb.Image = My.Resources.ColumnNew.ToBitmap
        ElseIf Me.qc.FieldType = clsQueryColumn.enFieldType.Measure Then
            pb.Image = My.Resources.MeasureNewNew.ToBitmap
        ElseIf Me.qc.FieldType = clsQueryColumn.enFieldType.ImpMeasure Then
            pb.Image = My.Resources._Function.ToBitmap
        End If
        Me.Controls.Add(pb)
        AddHandler pb.Click, AddressOf pb_Click


        AddHandler lbl.MouseDown, AddressOf Me.lbl_MouseDown
        AddHandler lbl.MouseMove, AddressOf Me.lbl_MouseMove
        AddHandler lbl.QueryContinueDrag, AddressOf Me.lbl_QueryContinueDrag
        AddHandler lbl.DragDrop, AddressOf Me.lbl_DragDrop

        Me.Controls.Add(lbl)

        AddHandler lbl.DragOver, AddressOf lbl_DragOver

        Me.lbl.AllowDrop = True

    End Sub

    Private Sub pb_Click(sender As Object, e As EventArgs)

        If Me.qc.FieldType = clsQueryColumn.enFieldType.Measure Then
            Exit Sub
        End If

        Dim x As New Windows.Forms.ContextMenu
        x.MenuItems.Add("None", New EventHandler(AddressOf Me.FunctionChange))

        x.MenuItems.Add("Sum", New EventHandler(AddressOf Me.FunctionChange))
        If Me.qc.DataType = clsQueryColumn.enDataType.Text OrElse Me.qc.DataType = clsQueryColumn.enDataType.Bool Then
            x.MenuItems.Item(x.MenuItems.Count - 1).Visible = False
        End If

        x.MenuItems.Add("Average", New EventHandler(AddressOf Me.FunctionChange))
        If Me.qc.DataType = clsQueryColumn.enDataType.Text OrElse Me.qc.DataType = clsQueryColumn.enDataType.Bool Then
            x.MenuItems.Item(x.MenuItems.Count - 1).Visible = False
        End If

        x.MenuItems.Add("Minimum", New EventHandler(AddressOf Me.FunctionChange))
        If Me.qc.DataType = clsQueryColumn.enDataType.Bool Then
            x.MenuItems.Item(x.MenuItems.Count - 1).Visible = False
        End If

        x.MenuItems.Add("Maximum", New EventHandler(AddressOf Me.FunctionChange))
        If Me.qc.DataType = clsQueryColumn.enDataType.Bool Then
            x.MenuItems.Item(x.MenuItems.Count - 1).Visible = False
        End If

        x.MenuItems.Add("Count (Distinct)", New EventHandler(AddressOf Me.FunctionChange))

        x.MenuItems.Add("Count", New EventHandler(AddressOf Me.FunctionChange))
        If Me.qc.DataType = clsQueryColumn.enDataType.Bool Then
            x.MenuItems.Item(x.MenuItems.Count - 1).Visible = False
        End If

        x.MenuItems.Add("Standard Deviation", New EventHandler(AddressOf Me.FunctionChange))
        If Me.qc.DataType = clsQueryColumn.enDataType.Text OrElse Me.qc.DataType = clsQueryColumn.enDataType.Bool Then
            x.MenuItems.Item(x.MenuItems.Count - 1).Visible = False
        End If

        x.MenuItems.Add("Variance", New EventHandler(AddressOf Me.FunctionChange))
        If Me.qc.DataType = clsQueryColumn.enDataType.Text OrElse Me.qc.DataType = clsQueryColumn.enDataType.Bool Then
            x.MenuItems.Item(x.MenuItems.Count - 1).Visible = False
        End If

        x.MenuItems.Add("Median", New EventHandler(AddressOf Me.FunctionChange))
        If Me.qc.DataType = clsQueryColumn.enDataType.Text OrElse Me.qc.DataType = clsQueryColumn.enDataType.Bool OrElse Me.qc.DataType = clsQueryColumn.enDataType.DateTime Then
            x.MenuItems.Item(x.MenuItems.Count - 1).Visible = False
        End If

        x.MenuItems.Item(Me.qc.iFunction).Checked = True

        For Each _qc In Me.qc.Query.QueryColumns
            If Not _qc Is Me.qc Then
                If _qc.FieldType = clsQueryColumn.enFieldType.ImpMeasure AndAlso _qc.UniName = Me.qc.UniName Then
                    x.MenuItems.Item(_qc.iFunction).Visible = False
                End If
            End If
        Next _qc

        x.Show(Me, New Point(0, Me.Height))





    End Sub

    Private Sub FunctionChange(sender As Object, e As EventArgs)

        Dim blnUpdate As Boolean = True



        Me.qc.Query = Me.ctrlTable.ctrlDaxQuery.query



        If sender.index = 0 And Me.qc.FieldType = clsQueryColumn.enFieldType.Level Then
            Exit Sub
        End If


        If sender.index = Me.qc.iFunction Then
            Exit Sub
        End If

        If Me.qc.Query.SelectionMode = clsQuery.enQuerySelectionMode.Member Then
            For Each _qc As clsQueryColumn In Me.qc.Query.QueryColumns
                _qc.SelectedMember.Clear()
            Next _qc
            Me.qc.Query.SelectionMode = clsQuery.enQuerySelectionMode.Total
        End If

        If sender.index > 0 Then
            Me.qc.SelectionMode = clsQueryColumn.enSelectionMode.AllSearch
            Me.qc.Query.QueryColumnByGUID(Me.qc.GUID).SelectionMode = clsQueryColumn.enSelectionMode.AllSearch
        End If


        If Me.qc.FieldType = clsQueryColumn.enFieldType.Level Then

            Me.ctrlTable.ctrlDaxQuery.tm.Cubes(0).DeSelect(Me.qc.UniName)
            Me.qc.Query.QueryColumnByGUID(Me.qc.GUID).FieldType = clsQueryColumn.enFieldType.ImpMeasure
            Me.qc.Query.QueryColumnByGUID(Me.qc.GUID).iFunction = sender.index

        ElseIf Me.qc.FieldType = clsQueryColumn.enFieldType.ImpMeasure And sender.index > 0 And sender.index <> Me.qc.iFunction Then

            Me.qc.Query.QueryColumnByGUID(Me.qc.GUID).iFunction = sender.index

        ElseIf Me.qc.FieldType = clsQueryColumn.enFieldType.ImpMeasure And sender.index = 0 Then

            Dim blnExists As Boolean = False
            For Each _qc In Me.qc.Query.QueryColumns
                If _qc.FieldType = clsQueryColumn.enFieldType.Level And _qc.UniName.ToLower = Me.qc.UniName.ToLower Then
                    blnExists = True
                    Exit For
                End If
            Next _qc

            If blnExists = True Then
                For i As Integer = Me.qc.Query.QueryColumns.Count - 1 To 0 Step -1
                    If Me.qc.Query.QueryColumns.Item(i).FieldType = clsQueryColumn.enFieldType.ImpMeasure And Me.qc.Query.QueryColumns.Item(i).UniName.ToLower = Me.qc.UniName.ToLower And Me.qc.Query.QueryColumns.Item(i).iFunction = Me.qc.iFunction Then
                        Me.qc.Query.QueryColumns.RemoveAt(i)
                        blnUpdate = False
                        Exit For
                    End If
                Next i
            Else
                Me.ctrlTable.ctrlDaxQuery.tm.Cubes(0).SelectLevel(Me.qc.UniName)
                Me.qc.Query.QueryColumnByGUID(Me.qc.GUID).SearchTerm = ""
                Me.qc.Query.QueryColumnByGUID(Me.qc.GUID).DaxFilter = ""
                Me.qc.Query.QueryColumnByGUID(Me.qc.GUID).DaxStmnt = ""
                Me.qc.Query.QueryColumnByGUID(Me.qc.GUID).FieldType = clsQueryColumn.enFieldType.Level
                Me.qc.Query.QueryColumnByGUID(Me.qc.GUID).iFunction = 0
            End If



        End If

        Me.ctrlTable.ctrlDaxQuery.FilterControl = Nothing


        If blnUpdate = True Then
            Dim PBIField As Object = Me.ctrlTable.ctrlDaxQuery.tm.GetLevel(Me.qc.Query.QueryColumnByGUID(Me.qc.GUID).UniName)
            If PBIField Is Nothing Then
                PBIField = Me.ctrlTable.ctrlDaxQuery.tm.GetMeasure(Me.qc.Query.QueryColumnByGUID(Me.qc.GUID).UniName)
            End If
            Dim c As New ctrlFilter(PBIField, Me.qc.Query.QueryColumnByGUID(Me.qc.GUID), Me) With {
                .GUID = Me.qc.GUID
            }
            Me.ctrlTable.ctrlDaxQuery.FilterControl = c
            Me.ctrlTable.Controls.Add(c)
            If Me.qc.DataType = clsQueryColumn.enDataType.Text And Me.qc.isImplicitCast Then
                If IsNumeric(Me.qc.SearchTerm) = False Then
                    Me.qc.Query.QueryColumnByGUID(Me.qc.GUID).SearchTerm = ""
                    c.txtBox.Text = ""
                Else
                    c.txtBox.Text = Me.qc.SearchTerm
                End If
            End If
            c.Visible = False
            Me.qc.Query.QueryColumnByGUID(Me.qc.GUID).SearchTerm = c.txtBox.Text
            Me.qc.Query.QueryColumnByGUID(Me.qc.GUID).DaxStmnt = c.DaxStmnt
            Me.qc.Query.QueryColumnByGUID(Me.qc.GUID).DaxFilter = c.DaxFilter
            Me.ctrlTable.ctrlDaxQuery.query = Me.qc.Query
            Me.ctrlTable.ctrlDaxQuery.FilterControl = Nothing
        End If





        Me.ctrlTable.ctrlDaxQuery.RefreshPreview()




    End Sub



    Private Sub lbl_QueryContinueDrag(ByVal sender As Object, ByVal e As System.Windows.Forms.QueryContinueDragEventArgs)

        'If e.Action = DragAction.Cancel OrElse e.Action = DragAction.Drop Then
        'MsgBox(e.Action.ToString)
        'End If

    End Sub



    Private Sub lbl_MouseDown(sender As Object, e As MouseEventArgs)

        Me.ctrlTable.ctrlDaxQuery.FilterControl = Nothing
        Me.ctrlTable.ctrlDaxQuery.DragDropControl = Me.ctrlTable
        Me.ctrlTable.ctrlDaxQuery.DragDropObject = Me.qc


    End Sub

    Private Sub lbl_MouseMove(sender As Object, e As MouseEventArgs)

        Me.Cursor = Cursors.Default
        If e.Button <> MouseButtons.Left Then Exit Sub
        Me.lbl.DoDragDrop("FromTableColumn", DragDropEffects.Copy)
    End Sub

    Private Sub lbl_DragOver(sender As Object, e As DragEventArgs)

        Dim strSource As String = e.Data.GetData(DataFormats.StringFormat)
        If strSource = "FromTableMembers" AndAlso Me.ctrlTable.ctrlDaxQuery.DragDropControl Is Me.ctrlTable Then ' AndAlso Me.ColUniName <> TryCast(Me.ctrlTable.ctrlQuery.DragDropObject, clsTabularModel.Members).Level.UniName Then
            e.Effect = DragDropEffects.None
            Exit Sub
        End If

        e.Effect = DragDropEffects.Copy
        Me.ctrlTable.ctrlDaxQuery.DragOverColumn = Me
    End Sub

    Private Sub lbl_DragDrop(sender As Object, e As DragEventArgs)

        If Me.ctrlTable.ctrlDaxQuery.DragOverColumn Is Nothing Then
            Exit Sub
        End If

        If Not TryCast(Me.ctrlTable.ctrlDaxQuery.DragDropObject, clsTabularModel.Measure) Is Nothing Or Not TryCast(Me.ctrlTable.ctrlDaxQuery.DragDropObject, clsTabularModel.Level) Is Nothing Then
            If Me.ctrlTable.ctrlDaxQuery.DragDropObject Is Me.qc Then
                Me.ctrlTable.ctrlDaxQuery.DragOverColumn = Nothing
                Exit Sub
            End If
        End If


        'switch
        Dim strSource As String = e.Data.GetData(DataFormats.StringFormat)
        If strSource = "FromTableMembers" Then
            Exit Sub
        End If



        'Column reorder
        If Me.ctrlTable.ctrlDaxQuery.DragDropControl Is Me.ctrlTable Then


            Dim qcSource As clsQueryColumn = Me.ctrlTable.ctrlDaxQuery.DragDropObject
            Dim qcTarget As clsQueryColumn = Me.qc

            Dim intSourceCol As Integer = Me.ctrlTable.GetColumnIndex(qcSource)
            Dim intTargetCol As Integer = Me.ctrlTable.GetColumnIndex(qcTarget)

            If Me.ctrlTable.ctrlDaxQuery.query.QueryColumns.Count = 2 Then
                Me.ctrlTable.fgT.Cols(0).Move(1)
            Else
                If intSourceCol >= 0 And intTargetCol >= 0 Then
                    Me.ctrlTable.fgT.Cols(intSourceCol).Move(intTargetCol)
                End If
            End If


            Me.ctrlTable.SetOrdinals()
            For Each _qc In Me.ctrlTable.ctrlDaxQuery.query.QueryColumns
                If Not Me.qc.Query.QueryColumnByGUID(_qc.GUID) Is Nothing Then
                    Me.qc.Query.QueryColumnByGUID(_qc.GUID).Ordinal = _qc.Ordinal
                End If
            Next _qc
            Me.ctrlTable.ctrlDaxQuery.DragOverColumn = Nothing

        Else





        End If








    End Sub

    Private Sub btn_MouseMove(sender As Object, e As MouseEventArgs)
        Me.Cursor = Cursors.Default
    End Sub

    Private Sub btn_Click(sender As Object, e As EventArgs)

        'Dim qc As clsQueryColumn = Me.ctrlTable.ctrlDaxQuery.query.GetQueryColumn(Me.ColUniName)

        If Not Me.ctrlTable.ctrlDaxQuery.FilterControl Is Nothing Then
            If Me.ctrlTable.ctrlDaxQuery.FilterControl.GUID = Me.qc.GUID Then
                Me.ctrlTable.ctrlDaxQuery.FilterControl = Nothing
                Exit Sub
            End If
        End If


        Dim PBIField As Object = Me.ctrlTable.ctrlDaxQuery.tm.GetLevel(Me.qc.UniName)
        If PBIField Is Nothing Then
            PBIField = Me.ctrlTable.ctrlDaxQuery.tm.GetMeasure(Me.qc.UniName)
        End If

        Dim c As New ctrlFilter(PBIField, Me.qc, Me) With {
            .GUID = Me.qc.GUID
        }
        Me.ctrlTable.ctrlDaxQuery.FilterControl = c
        c.Left = Me.Parent.Parent.Parent.Left + Me.Parent.Parent.Left + Me.Parent.Left + Me.Left - 2
        c.Top = Me.Parent.Parent.Parent.Top + Me.Parent.Parent.Top + Me.Parent.Top + Me.Top + Me.Height + 2
        Me.ctrlTable.Controls.Add(c)
        c.Visible = True
        c.BringToFront()





    End Sub



End Class

