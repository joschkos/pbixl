

Public Class ctrlColumnHeader

    Friend ctrlTable As ctrlTable
    Friend btn As Button
    Friend lbl As Label
    Friend ColUniName As String

    Public Sub New(ctrlTable As ctrlTable, qc As clsQueryColumn)

        InitializeComponent()

        If qc Is Nothing Then
            Exit Sub
            Me.Dispose()
        End If


        Me.ctrlTable = ctrlTable
        Me.ColUniName = qc.UniName

        Me.btn = New Button With {.BackColor = Color.AliceBlue, .Top = 0,
            .Left = Me.Width - 16, .Width = 14, .Height = 14, .Anchor = 9,
            .Image = Me.ctrlTable.ctrlDaxQuery.ImageList.Images("qe_default.ico"),
            .ImageAlign = ContentAlignment.MiddleCenter, .FlatStyle = FlatStyle.Standard}
        If qc.IsFiltered = True Then
            If qc.Sort = clsQueryColumn.enSort.none Then
                Me.btn.Image = Me.ctrlTable.ctrlDaxQuery.ImageList.Images("qe_filter.ico")
            ElseIf qc.Sort = clsQueryColumn.enSort.asc Then
                Me.btn.Image = Me.ctrlTable.ctrlDaxQuery.ImageList.Images("qe_filter_asc.ico")
            ElseIf qc.Sort = clsQueryColumn.enSort.desc Then
                Me.btn.Image = Me.ctrlTable.ctrlDaxQuery.ImageList.Images("qe_filter_desc.ico")
            End If
        Else
            If qc.Sort = clsQueryColumn.enSort.asc Then
                Me.btn.Image = Me.ctrlTable.ctrlDaxQuery.ImageList.Images("qe_default_asc.ico")
            ElseIf qc.Sort = clsQueryColumn.enSort.desc Then
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
        If qc.FieldType = clsQueryColumn.enFieldType.Measure Then
            strCaption = Me.ctrlTable.ctrlDaxQuery.tm.GetMeasure(qc.UniName).Caption
        Else
            strCaption = Me.ctrlTable.ctrlDaxQuery.tm.GetLevel(qc.UniName).Caption
        End If

        Me.lbl = New Label With {.Text = strCaption, .Top = 0, .Left = 2, .Width = Me.Width - 12, .AutoEllipsis = True, .Anchor = 13}


        AddHandler lbl.MouseDown, AddressOf Me.lbl_MouseDown
        AddHandler lbl.MouseMove, AddressOf Me.lbl_MouseMove
        AddHandler lbl.QueryContinueDrag, AddressOf Me.lbl_QueryContinueDrag
        AddHandler lbl.DragDrop, AddressOf Me.lbl_DragDrop

        Me.Controls.Add(lbl)

        AddHandler lbl.DragOver, AddressOf lbl_DragOver

        Me.lbl.AllowDrop = True

    End Sub



    Private Sub lbl_QueryContinueDrag(ByVal sender As Object, ByVal e As System.Windows.Forms.QueryContinueDragEventArgs)

        'If e.Action = DragAction.Cancel OrElse e.Action = DragAction.Drop Then
        'MsgBox(e.Action.ToString)
        'End If

    End Sub



    Private Sub lbl_MouseDown(sender As Object, e As MouseEventArgs)

        Me.ctrlTable.ctrlDaxQuery.FilterControl = Nothing

        Me.ctrlTable.ctrlDaxQuery.DragDropControl = Me.ctrlTable
        Me.ctrlTable.ctrlDaxQuery.DragDropObject = Me.ColUniName

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
            If Me.ctrlTable.ctrlDaxQuery.DragDropObject.UniName = Me.ColUniName Then
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

            Dim qcSource As clsQueryColumn = Me.ctrlTable.ctrlDaxQuery.query.GetQueryColumn(Me.ctrlTable.ctrlDaxQuery.DragDropObject.UniName)
            Dim qcTarget As clsQueryColumn = Me.ctrlTable.ctrlDaxQuery.query.GetQueryColumn(Me.ColUniName)

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

            Me.ctrlTable.ctrlDaxQuery.DragOverColumn = Nothing



        Else





        End If








    End Sub

    Private Sub btn_MouseMove(sender As Object, e As MouseEventArgs)
        Me.Cursor = Cursors.Default
    End Sub

    Private Sub btn_Click(sender As Object, e As EventArgs)

        Dim qc As clsQueryColumn = Me.ctrlTable.ctrlDaxQuery.query.GetQueryColumn(Me.ColUniName)

        If Not Me.ctrlTable.ctrlDaxQuery.FilterControl Is Nothing Then
            If Me.ctrlTable.ctrlDaxQuery.FilterControl.GUID = qc.GUID Then
                Me.ctrlTable.ctrlDaxQuery.FilterControl = Nothing
                Exit Sub
            End If
        End If


        Dim PBIField As Object = Me.ctrlTable.ctrlDaxQuery.tm.GetLevel(qc.UniName)
        If PBIField Is Nothing Then
            PBIField = Me.ctrlTable.ctrlDaxQuery.tm.GetMeasure(qc.UniName)
        End If

        Dim c As New ctrlFilter(PBIField, qc, Me) With {
            .GUID = qc.GUID
        }
        Me.ctrlTable.ctrlDaxQuery.FilterControl = c
        c.Left = Me.Parent.Parent.Parent.Left + Me.Parent.Parent.Left + Me.Parent.Left + Me.Left - 2
        c.Top = Me.Parent.Parent.Parent.Top + Me.Parent.Parent.Top + Me.Parent.Top + Me.Top + Me.Height + 2
        Me.ctrlTable.Controls.Add(c)
        c.Visible = True
        c.BringToFront()





    End Sub



End Class

