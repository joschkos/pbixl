Public Class frmMain


    Private pnlMain As System.Windows.Forms.Panel
    Friend btnCancel As System.Windows.Forms.Button
    Friend btnOK As System.Windows.Forms.Button
    Private ctrlDaxQuery As ctrlDaxQuery

    Public Sub CancelQuery()
        If Not Me.ctrlDaxQuery.ctsSource Is Nothing Then
            Try
                Me.ctrlDaxQuery.ctsSource.Cancel()
                Me.ctrlDaxQuery.conn.Dispose
                Me.ctrlDaxQuery.conn = Nothing
            Catch ex As Exception
            End Try
        End If
        If Not Me.ctrlDaxQuery.ctsCubeSource Is Nothing Then
            Try
                Me.ctrlDaxQuery.ctsCubeSource.Cancel()
                Me.ctrlDaxQuery.conn.Dispose
                Me.ctrlDaxQuery.conn = Nothing
            Catch ex As Exception
            End Try
        End If
    End Sub




    Public ReadOnly Property DAX As String
        Get
            Return Me.ctrlDaxQuery.query.DAX(False)
        End Get
    End Property

    Public ReadOnly Property Query As clsQuery
        Get
            Return Me.ctrlDaxQuery.query
        End Get
    End Property

    Public Property ConnName As String
    Public Property QueryName As String



    Public Sub New(strConn As String, CubeName As String, q As clsQuery, ConnName As String, QueryName As String)

        ' This call is required by the designer.
        InitializeComponent()

        Me.ConnName = ConnName
        Me.QueryName = QueryName


        ' Add any initialization after the InitializeComponent() call.
        Me.Width = 1032
        Me.Height = Me.Width / 1.618

        Me.btnCancel = New System.Windows.Forms.Button With {.Text = "Cancel", .DialogResult = DialogResult.Cancel}
        With Me.btnCancel
            .Top = Me.ClientSize.Height - Me.btnCancel.Height - 10
            .Left = Me.ClientSize.Width - Me.btnCancel.Width - 10
            .Anchor = 10
        End With
        Me.Controls.Add(Me.btnCancel)
        AddHandler Me.btnCancel.Click, AddressOf btnCancel_Click

        Me.CancelButton = Me.btnCancel

        Me.btnOK = New System.Windows.Forms.Button With {.Text = "OK", .DialogResult = DialogResult.OK}
        With Me.btnOK
            .Top = Me.ClientSize.Height - Me.btnCancel.Height - 10
            .Left = Me.ClientSize.Width - Me.btnCancel.Width - 10 - Me.btnOK.Width - 5
            .Anchor = 10
        End With
        AddHandler Me.btnOK.Click, AddressOf btnOK_Click

        Me.Controls.Add(Me.btnOK)

        Me.pnlMain = New System.Windows.Forms.Panel
        With Me.pnlMain
            .BackColor = Color.LightBlue
            .Top = 5
            .Left = 5
            .Height = Me.ClientSize.Height - 50
            .Width = Me.ClientSize.Width - 10
            .Anchor = 15
        End With
        Me.Controls.Add(Me.pnlMain)

        Me.StartPosition = FormStartPosition.CenterScreen



        Me.ctrlDaxQuery = New ctrlDaxQuery(Me.ConnName, Me.QueryName, strConn, CubeName, q, True, True, Me.btnOK)


        Me.ctrlDaxQuery.Dock = DockStyle.Fill
        Me.pnlMain.Controls.Add(Me.ctrlDaxQuery)






    End Sub

    Private Sub btnCancel_Click()
        Me.CancelQuery()
        Me.DialogResult = DialogResult.Cancel
        Me.Close()
    End Sub

    Private Sub btnOK_Click()
        Me.CancelQuery()
        Me.DialogResult = DialogResult.OK
        Me.Close()
    End Sub


End Class