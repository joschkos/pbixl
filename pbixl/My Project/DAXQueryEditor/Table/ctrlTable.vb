Public Class ctrlTable

    Friend ctrlDaxQuery As ctrlDaxQuery

    Private Property pnlHeader As Panel
    Private Property lblHeader As Label
    Private Property pnlMain As Panel
    Private Property pnlException As Panel
    Private Property btnRetry As Button
    Private Property pnlLoading As Panel
    Private Property lblCancel As LinkLabel

    Private Property lblSwitchPreview As LinkLabel

    Private Property pb As PictureBox
    Private Property pbC As PictureBox
    Friend Property Err As Exception
    Private Property txtException As TextBox

    Private Property objRec As Object
    Private Property objCmd As Object

    Private Property dtTotals As DataTable



    Private _fgT As C1.Win.C1FlexGrid.C1FlexGrid
    Public ReadOnly Property fgT As C1.Win.C1FlexGrid.C1FlexGrid
        Get
            If Not Me._fgT Is Nothing Then
                Return Me._fgT
            End If

            Me._fgT = New C1.Win.C1FlexGrid.C1FlexGrid
            With Me._fgT

                .BeginUpdate()
                .Redraw = False
                .BorderStyle = C1.Win.C1FlexGrid.Util.BaseControls.BorderStyleEnum.None
                .Cols.Count = 0
                .Rows.Count = 0

                .Cols.Fixed = 0
                .ExtendLastCol = False
                .Styles.Normal.Border.Style = C1.Win.C1FlexGrid.BorderStyleEnum.None
                .Styles.Normal.BackColor = Drawing.Color.WhiteSmoke
                .Styles.Normal.Border.Color = Drawing.Color.WhiteSmoke
                .Styles.EmptyArea.Border.Color = Drawing.Color.WhiteSmoke
                .Styles.EmptyArea.BackColor = Drawing.Color.WhiteSmoke
                .AllowEditing = False

                .HighLight = C1.Win.C1FlexGrid.HighLightEnum.Always
                .Styles.Highlight.BackColor = Drawing.Color.LightGray
                .Styles.Highlight.ForeColor = Drawing.Color.Black
                .SelectionMode = C1.Win.C1FlexGrid.SelectionModeEnum.Row
                .FocusRect = C1.Win.C1FlexGrid.FocusRectEnum.None

                .Styles.Frozen.BackColor = Drawing.Color.WhiteSmoke


                .AllowDragging = C1.Win.C1FlexGrid.AllowDraggingEnum.None
                .AllowSorting = C1.Win.C1FlexGrid.AllowSortingEnum.None

                .DropMode = C1.Win.C1FlexGrid.DropModeEnum.Manual

                .Anchor = 15

                .SelectionMode = C1.Win.C1FlexGrid.SelectionModeEnum.ListBox

                .DrawMode = C1.Win.C1FlexGrid.DrawModeEnum.OwnerDraw


                .Dock = DockStyle.Fill

                .AutoClipboard = True


                .Redraw = True
                .EndUpdate()
            End With

            AddHandler Me._fgT.OwnerDrawCell, AddressOf Me.fgT_OwnerDrawCell
            AddHandler Me._fgT.DragOver, AddressOf Me.fgT_DragOver


            Return Me._fgT
        End Get
    End Property

    Private Sub fgT_DragOver(sender As Object, e As DragEventArgs)

        If Me Is Me.ctrlDaxQuery.DragDropControl Then
            e.Effect = DragDropEffects.None
            Me.ctrlDaxQuery.DragOverColumn = Nothing
        Else
            Try
                Me.ctrlDaxQuery.DragOverColumn = TryCast(Me.fgT.GetUserData(0, Me.fgT.HitTest().Column), ctrlColumnHeader)
            Catch ex As Exception
                Me.ctrlDaxQuery.DragOverColumn = Nothing
            End Try

            e.Effect = DragDropEffects.Copy
        End If

    End Sub


    Private Sub fgT_OwnerDrawCell(sender As Object, e As C1.Win.C1FlexGrid.OwnerDrawCellEventArgs)


        If e.Row > 0 Then Exit Sub

        SyncLock Me.fgT
            If e.Row = 0 Then


                If Not Me.fgT.GetUserData(e.Row, e.Col) Is Nothing Then
                    Dim ch As ctrlColumnHeader = TryCast(Me.fgT.GetUserData(e.Row, e.Col), ctrlColumnHeader)
                    With ch
                        .Width = Me.fgT.GetCellRect(e.Row, e.Col).Width - 4
                        .Height = Me.fgT.GetCellRect(e.Row, e.Col).Height - 4
                        .Left = Me.fgT.GetCellRect(e.Row, e.Col).Left + 2
                        .Top = Me.fgT.GetCellRect(e.Row, e.Col).Top + 2
                    End With
                Else
                    Dim ch As New ctrlColumnHeader(Me, Me.fgT.Cols(e.Col).UserData) With {
                    .BackColor = Color.White,
                    .Width = Me.fgT.GetCellRect(e.Row, e.Col).Width - 4,
                    .Height = Me.fgT.GetCellRect(e.Row, e.Col).Height - 4,
                    .Left = Me.fgT.GetCellRect(e.Row, e.Col).Left + 2,
                    .Top = Me.fgT.GetCellRect(e.Row, e.Col).Top + 2
                }
                    fgT.Controls.Add(ch)
                    Me.fgT.SetUserData(e.Row, e.Col, ch)
                End If

            End If




        End SyncLock





    End Sub



    Public Enum enctrlState
        loading = 1
        ready = 2
        exception = 3
        init = 4
    End Enum


    Private _ctrlState As enctrlState
    Public Property ctrlState As enctrlState
        Get
            Return _ctrlState
        End Get
        Set(value As enctrlState)
            If Me.IsDisposed = True Or Me.Disposing = True Then
                Return
            End If

            Me._ctrlState = value
            If Me._ctrlState = enctrlState.ready Then

                Me.InvokeIfRequired(Sub()
                                        Me.pnlMain.Visible = True
                                        Me.pnlLoading.Visible = False
                                        Me.pnlException.Visible = False
                                        Me.pb.Image = Nothing
                                        Me.pbC.Image = Me.ctrlDaxQuery.ImageList.Images("Info.ico")
                                        Me.fgT.Visible = True

                                    End Sub)
            ElseIf Me._ctrlState = enctrlState.init Then
                Me.InvokeIfRequired(Sub()

                                        Me.pnlMain.Visible = True
                                        Me.pnlLoading.Visible = False
                                        Me.pnlException.Visible = False
                                        Me.pb.Image = Nothing
                                        Me.pbC.Image = Nothing
                                        Me.fgT.Visible = False


                                        'Dim intWidth As Integer = 0
                                        'For Each c As C1.Win.C1FlexGrid.Column In Me.fgT.Cols
                                        '    intWidth += c.WidthDisplay
                                        'Next c
                                        'If intWidth > 50 And intWidth < 1000 Then
                                        '    Me.Width = intWidth + 40
                                        'End If

                                    End Sub)
            ElseIf Me._ctrlState = enctrlState.loading Then
                Me.InvokeIfRequired(Sub()
                                        Me.lblCancel.Text = "loading.."
                                        Me.lblCancel.Enabled = True
                                        Me.pnlMain.Visible = False
                                        Me.pnlLoading.Visible = True
                                        Me.pnlException.Visible = False
                                        Me.pb.Image = CType(My.Resources.Wheel, System.Drawing.Image)
                                        Me.pbC.Image = Me.ctrlDaxQuery.ImageList.Images("Info.ico")

                                    End Sub)
            ElseIf Me._ctrlState = enctrlState.exception Then
                Me.InvokeIfRequired(Sub()

                                        If Me.Err.Source.ToLower.Trim = "adodb.recordset" Then
                                            Me.ctrlDaxQuery.ShowData = False
                                            Me.ctrlDaxQuery.RefreshPreview()
                                            Exit Sub
                                        End If


                                        Dim xLine As String = ""
                                        Try
                                            Dim trace = New Diagnostics.StackTrace(Me.Err, True)
                                            Dim line As String = Strings.Right(trace.ToString, 5)
                                            Dim method As String = ""
                                            Dim Xcont As Integer = 0

                                            For Each sf As StackFrame In trace.GetFrames
                                                Xcont = Xcont + 1
                                                method = method & Xcont & "- " & sf.GetMethod().ReflectedType.ToString & " " & sf.GetMethod().Name & vbCrLf
                                            Next sf

                                            xLine = line.ToString & vbCrLf & method
                                        Catch ex As Exception
                                            xLine = ex.Message
                                        End Try


                                        Me.pnlMain.Visible = False
                                        Me.pnlLoading.Visible = False
                                        Me.pnlException.Visible = True
                                        If Not Me.Err Is Nothing Then
                                            Me.txtException.Text = "Error: " & Me.Err.Message & vbCrLf & xLine
                                        Else
                                            Me.txtException.Text = "Error: unknown."
                                        End If
                                        Me.pb.Image = Me.ctrlDaxQuery.ImageList.Images("PowerBI_NotOK.ico")
                                        Me.pbC.Image = Me.ctrlDaxQuery.ImageList.Images("Info.ico")

                                    End Sub)
            End If



        End Set
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

        Me.pbC = New PictureBox With {.Top = 2, .Left = Me.pnlHeader.Width - 25, .Height = 16, .Width = 16, .SizeMode = PictureBoxSizeMode.Zoom, .Anchor = 9}
        AddHandler Me.pbC.MouseClick, AddressOf pbC_MouseClick
        Me.pnlHeader.Controls.Add(pbC)

        Me.lblHeader = New Label With {.Top = 3, .Left = 24, .Width = Me.pnlHeader.Width - 30, .Anchor = 13, .AutoEllipsis = True, .Text = "Query Name "} ' Me.ctrlQuery.ConnName}
        Me.pnlHeader.Controls.Add(Me.lblHeader)

        Me.pnlMain = New Panel
        With Me.pnlMain
            .Location = New Point(0, Me.pnlHeader.Height)
            .Height = Me.Height - 20
            .Width = Me.Width
            .Anchor = 15
        End With

        Me.pnlMain.Controls.Add(Me.fgT)

        Me.pnlLoading = New Panel
        With Me.pnlLoading
            .Location = New Point(0, Me.pnlHeader.Height)
            .Height = Me.Height - 20
            .Width = Me.Width
            .Anchor = 15
        End With

        Me.lblCancel = New System.Windows.Forms.LinkLabel With {.Text = "loading..", .Top = 50, .Left = 20, .Anchor = 5}
        Me.pnlLoading.Controls.Add(Me.lblCancel)
        AddHandler Me.lblCancel.Click, AddressOf Me.lblCancel_Click

        Me.lblSwitchPreview = New System.Windows.Forms.LinkLabel With {.Text = "Click here to refresh the preview..", .Top = 50, .Left = 20, .Anchor = 5, .Visible = False, .Width = 200}
        Me.fgT.Controls.Add(Me.lblSwitchPreview)
        AddHandler Me.lblSwitchPreview.Click, AddressOf Me.lblSwitchPreview_Click



        Me.pnlException = New Panel
        With Me.pnlException
            .Location = New Point(0, Me.pnlHeader.Height)
            .Height = Me.Height - 20
            .Width = Me.Width
            .Anchor = 15
        End With
        Me.txtException = New TextBox With {.Top = 10, .Left = 10, .Width = Me.Width - 20, .Height = Me.Height - 60, .Multiline = True, .ReadOnly = True, .Anchor = 15}
        Me.btnRetry = New Button With {.Top = Me.txtException.Top + Me.txtException.Height + 3, .Left = 10, .Enabled = True, .Anchor = 6, .Text = "Retry", .Width = .Width * 0.8}

        Me.pnlException.Controls.Add(Me.txtException)
        Me.pnlException.Controls.Add(Me.btnRetry)

        AddHandler Me.btnRetry.Click, AddressOf Me.btnRetry_Click

        Me.Controls.Add(Me.pnlHeader)
        Me.Controls.Add(Me.pnlLoading)
        Me.Controls.Add(Me.pnlException)
        Me.Controls.Add(Me.pnlMain)

        Me.lblHeader.Text = Me.ctrlDaxQuery.QueryName

    End Sub



    Private Sub btnRetry_Click()
        If Not Me.ctrlDaxQuery Is Nothing Then
            If Not Me.ctrlDaxQuery.ctsSource Is Nothing Then
                Me.ctrlDaxQuery.ctsSource.Cancel()
            End If
        End If

        Me.ctrlDaxQuery.ShowData = False
        Me.ctrlDaxQuery.RefreshPreview()
    End Sub

    Private Sub lblCancel_Click()
        If Not Me.ctrlDaxQuery Is Nothing Then
            If Not Me.ctrlDaxQuery.ctsSource Is Nothing Then
                Me.ctrlDaxQuery.ctsSource.Cancel()
            End If
        End If

        Me.ctrlDaxQuery.ShowData = False
        Me.ctrlDaxQuery.RefreshPreview()


    End Sub

    Private Sub lblSwitchPreview_Click()
        Me.ctrlDaxQuery.ShowData = True
        Me.ctrlDaxQuery.RefreshPreview()
    End Sub

    Private Sub pbC_MouseClick()

        If Me.ctrlDaxQuery.query Is Nothing Then
            Exit Sub
        End If

        Dim dlg As New dlgDax
        dlg.txtDax.Text = Me.ctrlDaxQuery.query.FormatDax(Me.ctrlDaxQuery.query.DAX(False), 4)

        dlg.ShowDialog()

    End Sub


    Friend Sub ShowTableColumns()

        Me.fgT.InvokeIfRequired(Sub()

                                    With Me.fgT

                                        .Redraw = False
                                        .BeginUpdate()

                                        Dim ctr As Integer = 0
                                        .Cols.Count = Me.ctrlDaxQuery.query.QueryColumns.Count
                                        .Rows.Count = 1
                                        For Each qc In ctrlDaxQuery.query.AllQueryColumnsSortedByOrdinal
                                            ctr += 1
                                            .SetData(0, ctr - 1, qc.UniName)
                                            .Cols(ctr - 1).UserData = qc

                                            If Not TryCast(.GetUserData(0, ctr - 1), ctrlColumnHeader) Is Nothing Then
                                                TryCast(.GetUserData(0, ctr - 1), ctrlColumnHeader).Dispose()
                                                .SetUserData(0, ctr - 1, Nothing)
                                            End If
                                        Next qc

                                        For i As Integer = Me.fgT.Controls.Count - 1 To 0 Step -1
                                            If Not TryCast(Me.fgT.Controls.Item(i), ctrlColumnHeader) Is Nothing Then
                                                Try
                                                    TryCast(Me.fgT.Controls.Item(i), ctrlColumnHeader).Dispose()
                                                    Me.fgT.Controls.RemoveAt(i)
                                                Catch ex As Exception

                                                End Try

                                            End If
                                        Next i



                                        .EndUpdate()
                                        .Redraw = True

                                    End With


                                End Sub)


    End Sub


    Private ColConns As New List(Of Object)


    Friend Sub RunQuery(cts As System.Threading.CancellationToken, q As clsQuery, blnShowData As Boolean)

        Dim blnCancelled As Boolean = False
        Dim blnStarted As Boolean = False
        cts.Register(Function()
                         If blnStarted = True Then
                             blnCancelled = True
                             Try
                                 If Not Me.objRec Is Nothing Then
                                     If Me.objRec.state = 4 Then
                                         Me.objRec.cancel
                                         'Debug.Print("canceled " & Now.Ticks.ToString)
                                         Me.objRec = Nothing
                                     End If
                                 End If
                             Catch ex As Exception

                             End Try
                         End If
                         blnStarted = False
                         Return Nothing
                     End Function)

        blnStarted = True
        If blnCancelled = True Then
            Exit Sub
        End If


        Dim blnPoolConn As Boolean = False
        Dim objConn As Object = Nothing
        If blnShowData = True Then
            For Each c In Me.ColConns
                If c.State = 1 Then
                    objConn = c
                    blnPoolConn = True
                    Exit For
                End If
            Next c
            If objConn Is Nothing Then
                objConn = CreateObject("ADODB.CONNECTION")
                objConn.ConnectionString = Me.ctrlDaxQuery.ConnectionString
                objConn.open
            End If
        End If

        If blnCancelled = True Then
            Exit Sub
        End If

        If blnCancelled = True Then
            Exit Sub
        End If

        If blnShowData = True Then
            Me.objRec = CreateObject("ADODB.RECORDSET")
            Me.objCmd = CreateObject("ADODB.COMMAND")
            'Dim strDax As String = q.DAXBaseTable(True, 1000, True)
            Dim strDax As String = q.DAX(True)
            'MsgBox(q.DAX(False))
            If strDax = "" Then
                Exit Sub
            End If
            Me.objCmd.activeconnection = objConn
            Me.objCmd.commandtext = strDax
            Me.objCmd.commandtimeout = 60
            Me.objRec.open(Me.objCmd,, 0, 1)
        End If

        If blnCancelled = True Then
            Exit Sub
        End If

        If Me.IsDisposed = False And Me.Disposing = False And blnCancelled = False Then

            Me.fgT.InvokeIfRequired(Sub()

                                        If blnCancelled = True Then
                                            Me.fgT.Visible = False : Exit Sub
                                        End If

                                        Me.fgT.Visible = False
                                        Me.fgT.Redraw = False
                                        Me.fgT.BeginUpdate()

                                        Try

                                            Dim ctr As Integer = 0
                                            Me.fgT.Cols.Count = Me.ctrlDaxQuery.query.QueryColumns.Count
                                            Me.fgT.Rows.Count = 1
                                            For Each qc In ctrlDaxQuery.query.AllQueryColumnsSortedByOrdinal
                                                ctr += 1
                                                Me.fgT.SetData(0, ctr - 1, qc.UniName)
                                                Me.fgT.Cols(ctr - 1).UserData = qc

                                                If Not TryCast(Me.fgT.GetUserData(0, ctr - 1), ctrlColumnHeader) Is Nothing Then
                                                    TryCast(Me.fgT.GetUserData(0, ctr - 1), ctrlColumnHeader).Dispose()
                                                    Me.fgT.SetUserData(0, ctr - 1, Nothing)
                                                End If
                                            Next qc

                                            If blnCancelled = True Then
                                                Me.fgT.Visible = False : Exit Sub
                                            End If


                                            For i As Integer = Me.fgT.Controls.Count - 1 To 0 Step -1
                                                If Not TryCast(Me.fgT.Controls.Item(i), ctrlColumnHeader) Is Nothing Then
                                                    Try
                                                        TryCast(Me.fgT.Controls.Item(i), ctrlColumnHeader).Dispose()
                                                        Me.fgT.Controls.RemoveAt(i)
                                                    Catch ex As Exception
                                                    End Try
                                                End If
                                            Next i


                                            If blnCancelled = True Then
                                                Me.fgT.Visible = False : Exit Sub
                                            End If

                                            If blnShowData = True Then
                                                If Me.objRec.RecordCount = 0 Then
                                                    Me.fgT.Rows.Count = 1
                                                    Me.fgT.Rows.Fixed = 1
                                                    Me.fgT.Rows.Frozen = 0
                                                ElseIf Me.objRec.RecordCount > 1000 Then
                                                    Me.fgT.Rows.Count = 1 + 1000
                                                    Me.fgT.Rows.Fixed = 1
                                                    Me.fgT.Rows.Frozen = 0
                                                Else
                                                    Me.fgT.Rows.Count = 1 + Me.objRec.RecordCount
                                                    Me.fgT.Rows.Fixed = 1
                                                    Me.fgT.Rows.Frozen = 0
                                                End If
                                            Else
                                                Me.fgT.Rows.Count = 1
                                                Me.fgT.Rows.Fixed = 1
                                                Me.fgT.Rows.Frozen = 0
                                            End If

                                            If blnShowData = True Then
                                                Me.lblSwitchPreview.Visible = False
                                            Else
                                                Dim blnVisible As Boolean = False
                                                For Each c In Me.fgT.Controls
                                                    If c Is Me.lblSwitchPreview Then
                                                        blnVisible = True : Exit For
                                                    End If
                                                Next c
                                                If blnVisible = False Then
                                                    Me.fgT.Controls.Add(Me.lblSwitchPreview)
                                                End If


                                                Me.lblSwitchPreview.Visible = True
                                            End If

                                            If blnCancelled = True Then
                                                Me.fgT.Visible = False : Exit Sub
                                            End If


                                            ctr = 0

                                            If blnShowData = True Then

                                                Do While Me.objRec.eof = False
                                                    ctr += 1
                                                    For i As Integer = 0 To Me.objRec.fields.count - 1
                                                        Me.fgT.SetData(ctr, i, Me.objRec.fields(i).value)
                                                    Next i
                                                    Me.objRec.movenext
                                                    If ctr = 1001 Then Exit Do
                                                    If blnCancelled = True Then Me.fgT.Visible = False : Exit Sub
                                                Loop

                                                For Each c As C1.Win.C1FlexGrid.Column In Me.fgT.Cols
                                                    If blnCancelled = True Then Me.fgT.Visible = False : Exit Sub
                                                    If Not TryCast(c.UserData, clsQueryColumn) Is Nothing Then
                                                        Dim x As String = Me.ctrlDaxQuery.tm.GetFormat(TryCast(c.UserData, clsQueryColumn))

                                                        If x.Trim = "" Then
                                                            Try
                                                                x = Me.ctrlDaxQuery.tm.GetMeasure(TryCast(c.UserData, clsQueryColumn).UniName).FormatString
                                                            Catch ex As Exception
                                                            End Try
                                                        End If

                                                        c.Format = x
                                                        c.StyleFixedDisplay.Format = x
                                                    End If

                                                Next c

                                            End If

                                        Catch ex As Exception
                                            'Debug.Print("Error " & ex.Message)

                                        End Try


                                        If Me.Disposing = False Then
                                            Me.fgT.EndUpdate()
                                            Me.fgT.Redraw = True
                                            Me.fgT.Visible = True
                                        End If




                                    End Sub)

        End If


        If Not Me.objRec Is Nothing Then
            If Me.objRec.State = 1 Then
                Me.objRec.close
            End If
            Me.objRec = Nothing
        End If
        If Not objConn Is Nothing Then
            If objConn.State = 1 Then
                If blnPoolConn = False Then
                    Me.ColConns.Add(objConn)
                End If
            End If
        End If



    End Sub




    Friend Function GetColumnIndex(qc As clsQueryColumn) As Integer
        Dim intRes As Integer = -1
        For i As Integer = 0 To Me.fgT.Cols.Count - 1
            If Not Me.fgT.Cols(i).UserData Is Nothing Then
                If Me.fgT.Cols(i).UserData Is qc Then
                    intRes = i
                    Exit For
                End If
            End If
        Next i
        Return intRes
    End Function

    Friend Sub SetOrdinals()

        For i As Integer = 0 To Me.fgT.Cols.Count - 1
            If Not TryCast(Me.fgT.Cols(i).UserData, clsQueryColumn) Is Nothing Then
                TryCast(Me.fgT.Cols(i).UserData, clsQueryColumn).Ordinal = i
            End If
        Next i

    End Sub

End Class
