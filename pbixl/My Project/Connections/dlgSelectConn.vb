Imports System.Windows.Forms
Imports Excel = Microsoft.Office.Interop.Excel

Public Class dlgSelectConn

    Private fg As New C1.Win.C1FlexGrid.C1FlexGrid
    Private Connections As clsConnections
    Private wb As Excel.Workbook


    Public ReadOnly Property SelectedConn As clsConnections.clsConnection
        Get
            Return TryCast(Me.fg.Rows(Me.fg.Selection.r1).UserData, clsConnections.clsConnection)
        End Get
    End Property

    Public Sub New(wb As Excel.Workbook)

        InitializeComponent()

        Me.wb = wb

        Me.fg = New C1.Win.C1FlexGrid.C1FlexGrid
        With Me.fg

            .Top = 5
            .Left = 5
            .Width = Me.ClientSize.Width - 10
            .Height = Me.ClientSize.Height - 10 - Me.OK_Button.Height - 20
            .Anchor = 15
            .Rows.Fixed = 1
            .Rows.Count = 1
            .Cols.Fixed = 0
            .Cols.Count = 1
            .ExtendLastCol = True

            .BackColor = Color.WhiteSmoke

            .Styles.EmptyArea.Border.Color = Drawing.Color.WhiteSmoke
            .Styles.EmptyArea.BackColor = Drawing.Color.WhiteSmoke

            .Styles.Normal.Border.Style = C1.Win.C1FlexGrid.BorderStyleEnum.None
            .Styles.Normal.BackColor = Drawing.Color.WhiteSmoke
            .Styles.Normal.Border.Color = Drawing.Color.WhiteSmoke

            .HighLight = C1.Win.C1FlexGrid.HighLightEnum.Always
            .Styles.Highlight.BackColor = Drawing.Color.LightGray
            .Styles.Highlight.ForeColor = Drawing.Color.Black
            .SelectionMode = C1.Win.C1FlexGrid.SelectionModeEnum.Row
            .FocusRect = C1.Win.C1FlexGrid.FocusRectEnum.None

            .Styles.Fixed.BackColor = Drawing.Color.WhiteSmoke

            .Tree.Column = 0
            .Tree.Style = C1.Win.C1FlexGrid.TreeStyleFlags.SimpleLeaf


            .SetData(0, 0, "Connections")


            .AllowEditing = False
            .AllowDelete = False
            .AllowSorting = False

            .Redraw = False

        End With

        AddHandler Me.fg.DoubleClick, AddressOf fg_DoubleClick


        Me.Controls.Add(Me.fg)



        Me.Connections = New clsConnections(Me.wb)
        Me.Connections.Refresh()

        For Each c As clsConnections.clsConnection In Me.Connections.Connections
            Me.fg.Rows.Count += 1
            Me.fg.Rows(Me.fg.Rows.Count - 1).IsNode = True
            Me.fg.Rows(Me.fg.Rows.Count - 1).UserData = c
            c.nd = Me.fg.Rows(Me.fg.Rows.Count - 1).Node
            c.TestConnection()
        Next c

        Me.fg.Redraw = True



    End Sub

    Private Sub fg_DoubleClick()
        If Me.fg.Selection.r1 >= 1 Then
            Dim c As clsConnections.clsConnection = TryCast(Me.fg.Rows(Me.fg.Selection.r1).UserData, clsConnections.clsConnection)
            If c.ConnState = clsConnections.clsConnection.enConnState.OK Then
                Me.DialogResult = DialogResult.OK
                Me.Close()
            End If
        End If

    End Sub

    Private Sub OK_Button_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles OK_Button.Click

        If Me.fg.Selection.r1 >= 1 Then
            Dim c As clsConnections.clsConnection = TryCast(Me.fg.Rows(Me.fg.Selection.r1).UserData, clsConnections.clsConnection)
            If Not c.ConnState = clsConnections.clsConnection.enConnState.OK Then
                Exit Sub
            End If
        End If

        Me.DialogResult = System.Windows.Forms.DialogResult.OK
        Me.Close()
    End Sub

    Private Sub Cancel_Button_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Cancel_Button.Click
        Me.DialogResult = System.Windows.Forms.DialogResult.Cancel
        Me.Close()
    End Sub

    Private Sub BtnRefresh_Click(sender As Object, e As EventArgs) Handles btnRefresh.Click



        Me.fg.Rows.Count = 1

        Me.Connections = Nothing

        Me.Connections = New clsConnections(Me.wb)
        Me.Connections.Refresh()

        For Each c As clsConnections.clsConnection In Me.Connections.Connections
            Me.fg.Rows.Count += 1
            Me.fg.Rows(Me.fg.Rows.Count - 1).IsNode = True
            Me.fg.Rows(Me.fg.Rows.Count - 1).UserData = c
            c.nd = Me.fg.Rows(Me.fg.Rows.Count - 1).Node
            c.TestConnection()
        Next c

    End Sub
End Class
