Imports System.Windows.Forms
Imports Excel = Microsoft.Office.Interop.Excel

Public Class dlgSelectConn

    Private fg As New C1.Win.C1FlexGrid.C1FlexGrid
    Private Connections As clsConnections
    Private wb As Excel.Workbook

    Private ts As ToolStrip
    Private sb As ToolStripSplitButton
    Private pl As Windows.Forms.Panel

    Public Enum enumDialogMode
        Table = 1
        Pivot = 2
    End Enum

    Private enDialogMode As enumDialogMode
    Public Property DialogMode As enumDialogMode
        Get
            Return enDialogMode
        End Get
        Set(value As enumDialogMode)
            enDialogMode = value
            If Me.enDialogMode = enumDialogMode.Table Then
                sb.Image = Me.ImageList.Images(9)
            Else
                sb.Image = Me.ImageList.Images(10)
            End If



        End Set
    End Property





    Public ReadOnly Property SelectedConn As clsConnections.clsConnection
        Get
            Return TryCast(Me.fg.Rows(Me.fg.Selection.r1).UserData, clsConnections.clsConnection)
        End Get
    End Property

    Public Sub New(wb As Excel.Workbook, DialogMode As enumDialogMode)

        InitializeComponent()

        Me.wb = wb



        Me.ts = New ToolStrip
        Me.sb = New ToolStripSplitButton()
        Me.sb.DropDownItems.Add("Table")
        Me.sb.DropDownItems.Item(0).Image = Me.ImageList.Images(9)
        Me.sb.DropDownItems.Add("Pivot Table")
        Me.sb.DropDownItems.Item(1).Image = Me.ImageList.Images(10)
        ts.Items.Add(sb)
        ts.LayoutStyle = ToolStripLayoutStyle.Flow
        ts.AutoSize = False
        ts.Width = Me.OK_Button.Width
        ts.Height = Me.OK_Button.Height + 5
        sb.AutoSize = False
        sb.Width = ts.Width
        sb.Height = ts.Height
        sb.BackColor = Me.OK_Button.BackColor
        ts.BackColor = Me.OK_Button.BackColor
        ts.Dock = DockStyle.None
        ts.Top = -5
        ts.ShowItemToolTips = False
        sb.Text = "OK"



        sb.AutoToolTip = False
        sb.ToolTipText = "click here"


        ts.ShowItemToolTips = False




        Me.DialogMode = DialogMode


        Dim pl As New Windows.Forms.Panel With {.Width = Me.OK_Button.Width, .Height = Me.OK_Button.Height - 2, .Top = Me.OK_Button.Top + 1, .Left = Me.OK_Button.Left}
        pl.BorderStyle = BorderStyle.FixedSingle
        pl.Controls.Add(ts)

        pl.Anchor = 10

        ts.BackColor = SystemColors.ControlLight


        AddHandler sb.ButtonClick, AddressOf sb_ButtonClick
        AddHandler sb.DropDownItemClicked, AddressOf sb_DropDownItemClicked
        AddHandler sb.DropDownOpening, AddressOf sb_DropDownOpening



        Me.OK_Button.Visible = False
        'Me.TableLayoutPanel1.Controls.Add(pl)
        Me.Controls.Add(pl)





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

            .BorderStyle = C1.Win.C1FlexGrid.Util.BaseControls.BorderStyleEnum.Light3D

            .SetData(0, 0, "Connections")


            .AllowEditing = False
            .AllowDelete = False
            .AllowSorting = False

            .Redraw = False

        End With

        AddHandler Me.fg.DoubleClick, AddressOf fg_DoubleClick


        Me.Controls.Add(Me.fg)

        Me.RefreshConnections()







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

        Me.RefreshConnections()


    End Sub

    Private Sub RefreshConnections()

        Me.fg.Redraw = False
        Me.fg.Rows.Count = 1


        Me.Connections = Nothing
        Me.Connections = New clsConnections(Me.wb)
        Me.Connections.Refresh()

        For Each c As clsConnections.clsConnection In Me.Connections.Connections
            If c.ConnType = clsConnections.clsConnection.enConnType.PBIDesktop Then
                Me.fg.Rows.Count += 1
                Me.fg.Rows(Me.fg.Rows.Count - 1).IsNode = True
                Me.fg.Rows(Me.fg.Rows.Count - 1).UserData = c
                c.nd = Me.fg.Rows(Me.fg.Rows.Count - 1).Node
                c.TestConnection()
            End If
        Next c
        For Each c As clsConnections.clsConnection In Me.Connections.Connections
            If c.ConnType = clsConnections.clsConnection.enConnType.PowerBI Then
                Me.fg.Rows.Count += 1
                Me.fg.Rows(Me.fg.Rows.Count - 1).IsNode = True
                Me.fg.Rows(Me.fg.Rows.Count - 1).UserData = c
                c.nd = Me.fg.Rows(Me.fg.Rows.Count - 1).Node
                c.TestConnection()
            End If
        Next c
        For Each c As clsConnections.clsConnection In Me.Connections.Connections
            If c.ConnType = clsConnections.clsConnection.enConnType.TabularSvr Then
                Me.fg.Rows.Count += 1
                Me.fg.Rows(Me.fg.Rows.Count - 1).IsNode = True
                Me.fg.Rows(Me.fg.Rows.Count - 1).UserData = c
                c.nd = Me.fg.Rows(Me.fg.Rows.Count - 1).Node
                c.TestConnection()
            End If
        Next c


        For Each c As clsConnections.clsConnection In Me.Connections.Connections
            If c.ConnType = clsConnections.clsConnection.enConnType.QueryObject Then
            End If
        Next c






        Me.fg.Redraw = True

    End Sub



    Private Sub sb_ButtonClick(sender As Object, e As EventArgs)
        If Me.fg.Selection.r1 >= 1 Then
            Dim c As clsConnections.clsConnection = TryCast(Me.fg.Rows(Me.fg.Selection.r1).UserData, clsConnections.clsConnection)
            If Not c.ConnState = clsConnections.clsConnection.enConnState.OK Then
                Exit Sub
            End If
        End If

        Me.DialogResult = System.Windows.Forms.DialogResult.OK
        Me.Close()
    End Sub

    Private Sub sb_DropDownItemClicked(sender As Object, e As ToolStripItemClickedEventArgs)
        'MsgBox(e.ClickedItem.Text)
        If e.ClickedItem.Text = "Table" Then
            Me.DialogMode = enumDialogMode.Table
        Else
            Me.DialogMode = enumDialogMode.Pivot
        End If


    End Sub

    Private Sub sb_DropDownOpening(sender As Object, e As EventArgs)
        If Me.DialogMode = enumDialogMode.Table Then
            Me.sb.DropDownItems(0).Select()
        Else
            Me.sb.DropDownItems(1).Select()
        End If


    End Sub

End Class
