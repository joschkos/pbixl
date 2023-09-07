Imports System.Runtime.CompilerServices

Public Class ctrlDaxQuery

    Friend Property ConnName As String
    Friend Property QueryName As String

    Friend Property ConnectionString As String
    Friend Property CubeName As String
    Friend Property query As clsQuery

    Friend Property tm As clsTabularModel
    Friend Property ctrlCube As ctrlCube
    Friend Property ctrlTable As ctrlTable

    Friend Property Async As Boolean

    Friend Property ShowData As Boolean





    Private _FilterConn As Object
    Friend ReadOnly Property FilterConn As Object
        Get
            If Me._FilterConn Is Nothing Then
                Me._FilterConn = CreateObject("ADODB.Connection")
                Me._FilterConn.connectionstring = Me.ConnectionString
            End If
            If Me._FilterConn.State <> 1 Then
                Me._FilterConn.open
            End If
            Return Me._FilterConn
        End Get
    End Property




    Private Property split As Windows.Forms.SplitContainer

    Protected Overrides ReadOnly Property CreateParams() As CreateParams
        Get
            Dim cp As CreateParams = MyBase.CreateParams
            cp.ExStyle = cp.ExStyle Or &H2000000
            Return cp
        End Get
    End Property 'CreateParams

    Private _DragDropObject As Object
    Friend Property DragDropObject As Object
        Get
            Return _DragDropObject
        End Get
        Set(value As Object)

            If TypeOf value Is String Then
                Dim f As Object = Me.tm.GetMeasure(value)
                If f Is Nothing Then f = Me.tm.GetLevel(value)
                Me._DragDropObject = f
            Else
                Me._DragDropObject = value
            End If

        End Set
    End Property

    Private _ctrFilter As ctrlFilter
    Friend Property FilterControl As ctrlFilter
        Get
            Return _ctrFilter
        End Get
        Set(value As ctrlFilter)
            If Me._ctrFilter Is value Then
                Exit Property
            ElseIf value Is Nothing And Me._ctrFilter Is Nothing Then
                Exit Property
            Else
                If Not Me._ctrFilter Is Nothing Then
                    Me._ctrFilter.Visible = False
                    Me._ctrFilter.Dispose()
                End If
            End If
            _ctrFilter = value
        End Set
    End Property

    Friend Property DragDropControl As Object

    Private _DragOverColumn As Object
    Friend Property DragOverColumn As Object
        Get
            Return _DragOverColumn
        End Get
        Set(value As Object)
            If value Is Nothing AndAlso Not Me._DragOverColumn Is Nothing Then
                TryCast(Me._DragOverColumn, ctrlColumnHeader).lbl.Font = New Font(TryCast(Me._DragOverColumn, ctrlColumnHeader).lbl.Font, FontStyle.Regular)
                Me._DragOverColumn = Nothing
            ElseIf value Is Nothing Then
                Me._DragOverColumn = Nothing
            ElseIf Me._DragOverColumn Is Nothing Then
                Me._DragOverColumn = value
                TryCast(Me._DragOverColumn, ctrlColumnHeader).lbl.Font = New Font(TryCast(Me._DragOverColumn, ctrlColumnHeader).lbl.Font, FontStyle.Underline)
            ElseIf Not value Is Me._DragOverColumn Then
                TryCast(Me.DragOverColumn, ctrlColumnHeader).lbl.Font = New Font(TryCast(Me.DragOverColumn, ctrlColumnHeader).lbl.Font, FontStyle.Regular)
                Me._DragOverColumn = value
                TryCast(Me._DragOverColumn, ctrlColumnHeader).lbl.Font = New Font(TryCast(Me._DragOverColumn, ctrlColumnHeader).lbl.Font, FontStyle.Underline)
            End If
        End Set
    End Property





    Private _conn As Object
    Public Property conn As Object
        Get
            Return _conn
        End Get
        Set(value As Object)
            _conn = value
        End Set
    End Property


    Public Sub New(ConnName As String, QueryName As String, ConnectionString As String, CubeName As String, query As clsQuery, Async As Boolean, ShowData As Boolean)

        InitializeComponent()



        Me.ConnName = ConnName
        Me.QueryName = QueryName

        Me.ConnectionString = ConnectionString
        Me.CubeName = CubeName
        Me.query = query
        Me.Async = Async
        Me.ShowData = ShowData

        Me.BackColor = Color.White
        Me.split = New SplitContainer
        Me.split.Dock = DockStyle.Fill

        Me.Controls.Add(split)

        Me.split.FixedPanel = FixedPanel.Panel1

        Me.ctrlCube = New ctrlCube(Me)
        Me.ctrlCube.Dock = DockStyle.Fill
        Me.split.Panel1.Controls.Add(Me.ctrlCube)
        Me.ctrlCube.ctrlState = ctrlCube.enctrlState.loading

        Me.ctrlTable = New ctrlTable(Me)
        Me.ctrlTable.Dock = DockStyle.Fill
        Me.split.Panel2.Controls.Add(Me.ctrlTable)
        Me.ctrlTable.ctrlState = ctrlTable.enctrlState.init



        Me.Init()



    End Sub

    Friend ctsCubeSource As System.Threading.CancellationTokenSource
    Friend ctsCube As System.Threading.CancellationToken

    Friend ctsSource As System.Threading.CancellationTokenSource
    Friend cts As System.Threading.CancellationToken

    Friend RunningQueryGUID As String

    Friend Sub RefreshPreview()

        Me.ctrlTable.ctrlState = ctrlTable.enctrlState.loading

        If Me.query Is Nothing Then
            Me.query = New clsQuery
            Me.query.CubeName = Me.CubeName
        End If


        Me.UpdateQueryColumns()



        If Me.query.QueryColumns.Count = 0 Then
            If Not Me.ctsSource Is Nothing Then
                Me.ctsSource.Cancel()
            End If


            Me.ctrlTable.fgT.Cols.Count = 0
            Me.ctrlTable.fgT.Rows.Count = 0
            For i As Integer = Me.ctrlTable.fgT.Controls.Count - 1 To 0 Step -1
                Me.ctrlTable.fgT.Controls.RemoveAt(i)
            Next i

            Me.ctrlTable.ctrlState = ctrlTable.enctrlState.init
            Exit Sub
        End If



        If Not Me.ctsSource Is Nothing Then
            Me.ctsSource.cancel
        End If

        Me.ctsSource = New System.Threading.CancellationTokenSource
        Me.cts = Me.ctsSource.Token




        Dim t = Task(Of Object).Factory.StartNew(Function()
                                                     Try


                                                         Dim q As clsQuery = Me.query.Clone


                                                         Me.ctrlTable.ctrlState = ctrlTable.enctrlState.loading
                                                         Me.RunningQueryGUID = System.Guid.NewGuid.ToString

                                                         Me.ctrlTable.RunQuery(cts, q, Me.ShowData)

                                                         Me.ctrlTable.ctrlState = ctrlTable.enctrlState.ready


                                                         Return Nothing
                                                     Catch ex As Exception
                                                         If cts.IsCancellationRequested = True Then
                                                             Me.ctrlTable.Err = ex
                                                             'Me.ctrlTable.ctrlState = ctrlTable.enctrlState.exception
                                                             Me.ShowData = False
                                                             Me.RefreshPreview()
                                                         Else
                                                             Me.ctrlTable.Err = ex
                                                             Me.ctrlTable.ctrlState = ctrlTable.enctrlState.exception
                                                         End If

                                                         Return Nothing
                                                     End Try
                                                 End Function)




    End Sub


    Private blnStarted As Boolean
    Private blnCancelled As Boolean

    Private Sub Init()

        Me.blnStarted = True
        Me.blnCancelled = False

        Me.ctsCubeSource = New System.Threading.CancellationTokenSource
        Me.ctsCube = Me.ctsCubeSource.Token

        Me.ctsCube.Register(Function()
                                If blnStarted = True Then
                                    blnCancelled = True
                                    'Debug.Print("cancelled" & Now.Ticks.ToString)
                                End If
                                blnStarted = False
                                Return Nothing
                            End Function)

        Dim t = Task(Of Object).Factory.StartNew(Function()
                                                     Try
                                                         Me.GetConnection()
                                                         If blnCancelled = True Then Return Nothing
                                                         Me.GetTabularModel()
                                                         If blnCancelled = True Then Return Nothing
                                                         If Me.Disposing OrElse Me.IsDisposed Then
                                                             Return Nothing
                                                         End If
                                                         Me.ctrlCube.ShowNavigation()
                                                         If blnCancelled = True Then Return Nothing

                                                         Me.ctrlCube.ctrlState = ctrlCube.enctrlState.ready
                                                         If Me.query Is Nothing Then
                                                             Me.ctrlTable.ctrlState = ctrlTable.enctrlState.init
                                                         Else
                                                             If blnCancelled = True Then Return Nothing
                                                             Me.ctrlTable.ctrlState = ctrlTable.enctrlState.loading

                                                             Dim q As clsQuery = Me.query.Clone
                                                             Me.query = Nothing
                                                             Me.query = New clsQuery()
                                                             With Me.query
                                                                 .GUID = q.GUID
                                                                 .FilterControlVisible = False
                                                                 .FilterControlGUID = ""
                                                                 .ConnectionString = q.ConnectionString
                                                                 .ConnectionName = q.ConnectionName
                                                                 .Command = q.Command
                                                                 .AddMissingItems = q.AddMissingItems
                                                                 .CubeName = q.CubeName
                                                                 .SelectionMode = q.SelectionMode
                                                             End With
                                                             If blnCancelled = True Then Return Nothing

                                                             Dim strErrMsg As String = ""
                                                             Dim intErrCtr As Integer = 0

                                                             For Each qc As clsQueryColumn In q.AllQueryColumnsSortedByOrdinal

                                                                 Dim _qc As New clsQueryColumn
                                                                 With _qc
                                                                     .BlankSel = qc.BlankSel
                                                                     .DataType = qc.DataType
                                                                     .DaxFilter = qc.DaxFilter
                                                                     .DaxStmnt = qc.DaxStmnt
                                                                     .FieldName = qc.FieldName
                                                                     .FieldType = qc.FieldType
                                                                     .FilterControlGUID = ""
                                                                     .GUID = qc.GUID
                                                                     If Not qc.htSel Is Nothing Then
                                                                         .htSel = qc.htSel.Clone
                                                                     End If
                                                                     .Ordinal = qc.Ordinal
                                                                     .Query = q
                                                                     .SearchTerm = qc.SearchTerm
                                                                     If Not qc.SelectedMember Is Nothing AndAlso qc.SelectedMember.Count > 0 Then
                                                                         .SelectedMember.AddRange(qc.SelectedMember)
                                                                     End If
                                                                     .SelectionMode = qc.SelectionMode
                                                                     .Sort = qc.Sort
                                                                     .TableName = qc.TableName
                                                                     .UniName = qc.UniName
                                                                     .IsSelected = True
                                                                 End With

                                                                 If blnCancelled = True Then Return Nothing


                                                                 If _qc.FieldType = clsQueryColumn.enFieldType.Level Then
                                                                     Dim l As clsTabularModel.Level = Me.tm.GetLevel(_qc.UniName)
                                                                     If Not l Is Nothing AndAlso l.DataType = qc.DataType Then
                                                                         l.IsSelected = True
                                                                         Me.query.QueryColumns.Add(_qc)
                                                                     Else
                                                                         intErrCtr += 1
                                                                         If l Is Nothing Then
                                                                             If intErrCtr <= 5 Then
                                                                                 strErrMsg += "Column " & qc.UniName & " removed from query because it is missing from model." & vbCrLf & vbCrLf
                                                                             End If
                                                                         Else
                                                                             If intErrCtr <= 5 Then
                                                                                 strErrMsg += "Column " & qc.UniName & " removed from query because due to data type mismatch." & vbCrLf & vbCrLf
                                                                             End If
                                                                         End If
                                                                     End If
                                                                 Else
                                                                     Dim m As clsTabularModel.Measure = Me.tm.GetMeasure(_qc.UniName)
                                                                     If Not m Is Nothing AndAlso m.DataType = qc.DataType Then
                                                                         m.IsSelected = True
                                                                         Me.query.QueryColumns.Add(_qc)
                                                                     Else
                                                                         intErrCtr += 1
                                                                         If m Is Nothing Then
                                                                             If intErrCtr <= 5 Then
                                                                                 strErrMsg += "Column " & qc.UniName & " removed from query because it is missing from model." & vbCrLf & vbCrLf
                                                                             End If
                                                                         Else
                                                                             If intErrCtr <= 5 Then
                                                                                 strErrMsg += "Column " & qc.UniName & " removed from query because due to data type mismatch." & vbCrLf & vbCrLf
                                                                             End If
                                                                         End If
                                                                     End If

                                                                 End If
                                                             Next qc

                                                             If strErrMsg <> "" Then
                                                                 If intErrCtr > 5 Then
                                                                     strErrMsg = strErrMsg & (intErrCtr - 5).ToString & " more columns removed."
                                                                 End If
                                                                 MsgBox(strErrMsg, vbCritical)
                                                             End If


                                                             If blnCancelled = True Then Return Nothing

                                                             'Set Display state
                                                             For i = Me.ctrlCube.fg.Rows.Count - 1 To 1 Step -1
                                                                 If Not TryCast(Me.ctrlCube.fg.Rows(i).UserData, clsTabularModel.DisplayFolder) Is Nothing Then
                                                                     TryCast(Me.ctrlCube.fg.Rows(i).UserData, clsTabularModel.DisplayFolder).SetDisplaystate()
                                                                 ElseIf Not TryCast(Me.ctrlCube.fg.Rows(i).UserData, clsTabularModel.Hierarchy) Is Nothing Then
                                                                     TryCast(Me.ctrlCube.fg.Rows(i).UserData, clsTabularModel.Hierarchy).SetDisplayState()
                                                                 ElseIf Not TryCast(Me.ctrlCube.fg.Rows(i).UserData, clsTabularModel.Dimension) Is Nothing Then
                                                                     TryCast(Me.ctrlCube.fg.Rows(i).UserData, clsTabularModel.Dimension).SetDisplayState()
                                                                 End If
                                                             Next i


                                                             'MsgBox("removed columns")



                                                             If blnCancelled = True Then Return Nothing
                                                             Me.RefreshPreview()

                                                         End If


                                                         Return Nothing
                                                     Catch ex As Exception
                                                         Me.ctrlCube.Err = ex
                                                         Me.ctrlCube.ctrlState = ctrlCube.enctrlState.exception

                                                         Return Nothing
                                                     End Try
                                                 End Function)


    End Sub

    Private Sub GetConnection()
        Dim objconn As Object = CreateObject("ADODB.CONNECTION")
        objconn.connectionstring = ConnectionString
        objconn.open
        Me.conn = objconn
    End Sub

    Private Sub GetTabularModel()
        Me.tm = New clsTabularModel(Me.CubeName, Me.ImageList32, Me.ImageList48, Me)
    End Sub

    Private Sub UpdateQueryColumns()
        Dim lstF As List(Of Object) = Me.tm.GetAllObjectFields(Me.tm.Cubes(0))
        Dim ctr As Integer = 0

        Dim htKeys As New Hashtable
        For Each f In lstF
            If htKeys.ContainsKey(f.UniName.ToString.ToLower) = False Then
                htKeys.Add(f.UniName.ToString.ToLower, Nothing)
                If f.IsSelected = True Then

                    If Me.query.ColumnIsInQuery(f.UniName) = False Then
                        Dim qc As New clsQueryColumn(Me.query, f.UniName, f.TableName, f.FieldName)
                        With qc
                            .Ordinal = 99
                            .IsSelected = True
                            .SelectionMode = f.SelectionMode
                            .DaxFilter = f.DaxFilter
                            .SearchTerm = f.SearchTerm
                            If Not f.htSel Is Nothing Then
                                .htSel = f.htSel.Clone
                            End If
                            .FilterControlGUID = f.FilterControlGUID
                            .DataType = f.DataType
                            .FieldType = f.FieldType
                            .TableName = f.TableName
                            .FieldName = f.FieldName
                            .UniName = f.UniName
                            .Sort = f.Sort
                        End With
                        Me.query.AddColumn(qc, 99999)

                    End If
                Else
                    If Me.query.ColumnIsInQuery(f.UniName) = True Then
                        Me.query.RemoveColumn(f.uniname)
                    End If
                End If

            End If
        Next f
    End Sub

End Class

Public Module Extensions
    <Extension()>
    Public Sub InvokeIfRequired(ByVal Control As Windows.Forms.Control, ByVal Method As [Delegate], ByVal ParamArray Parameters As Object())
        If Parameters Is Nothing OrElse
                Parameters.Length = 0 Then Parameters = Nothing
        If Control.InvokeRequired = True Then
            If Not Control Is Nothing And Control.IsDisposed = False And Control.Disposing = False Then
                Try
                    Control.Invoke(Method, Parameters)
                Catch ex As Exception

                End Try

            End If
        Else
            Method.DynamicInvoke(Parameters)
        End If
    End Sub
End Module