Imports Excel = Microsoft.Office.Interop.Excel

Public Class clsConnections

    Private ReadOnly lstConns As List(Of clsConnection)

    Public ReadOnly Property Conns As List(Of clsConnection)
        Get
            Return Me.lstConns
        End Get
    End Property

    Public Sub New()
        Me.lstConns = New List(Of clsConnection)
    End Sub

    Private _lstConns As List(Of Object)
    Public Function GetConnection(ConnString As String) As Object

        If Me._lstConns Is Nothing Then
            Me._lstConns = New List(Of Object)
        End If

        For i As Integer = Me._lstConns.Count - 1 To 0 Step -1
            If Me._lstConns.Item(i).ConnectionString = ConnString Then
                Return Me._lstConns.Item(i)
            End If
        Next i

        SyncLock Me._lstConns
            Dim conn As Object = CreateObject("ADODB.CONNECTION")
            conn.connectionstring = ConnString
            Me._lstConns.Add(conn)
            Return conn
        End SyncLock

        Return Nothing

    End Function

    Public Function GetConnectionString(WorkbookName As String, ConnName As String) As String

        Dim strR As String = ""
        For Each c As clsConnection In Me.Conns
            If c.Name.ToLower = ConnName.ToLower AndAlso c.WorkbookName.ToLower = WorkbookName.ToLower Then
                Return c.ConnectionString
            End If
        Next c

        Return ""


    End Function

    Public dtRefresh As DateTime
    Public Sub Refresh(WorkbookName As String)
        Dim wb As Excel.Workbook = MyAddIn.GetWorkbook(WorkbookName)
        Me._Refresh(wb)
    End Sub

    Public Sub Refresh(wb As Excel.Workbook)
        Me._Refresh(wb)
    End Sub

    Private Sub _Refresh(wb As Excel.Workbook)
        If (Now - dtRefresh).TotalSeconds < 5 Then
            Exit Sub
        End If
        Me.InitRefresh()
        Me.GetWorkbookConn(wb)
        Me.GetServerConns(wb)
        Me.GetDesktopConns()
        Me.Clean()
        Me.dtRefresh = Now
    End Sub

    Private Sub InitRefresh()
        For Each c As clsConnection In Me.lstConns
            c.Seen = False
        Next c
    End Sub

    Private Sub Clean()
        For i As Integer = Me.lstConns.Count - 1 To 0 Step -1
            If Me.lstConns.Item(i).Seen = False Or Me.lstConns.Item(i).IsLiveConn = True Then
                Me.lstConns.RemoveAt(i)
            End If
        Next i
    End Sub

    Private Sub GetDesktopConns()

        Dim lstPorts As New List(Of String)
        Try
            Dim tcps() As MIB_TCPROW_OWNER_PID = TCPConns()
            For Each proc As System.Diagnostics.Process In System.Diagnostics.Process.GetProcessesByName("msmdsrv")
                If proc.SessionId > 0 Then
                    Dim intProcID As Integer = proc.Id
                    For Each row As MIB_TCPROW_OWNER_PID In tcps
                        If row.PID = intProcID Then
                            Dim tcp As TcpConnection = MIB_ROW_To_TCP(row)
                            Dim strPort As String = tcp.localPort.ToString
                            lstPorts.Add(strPort)
                        End If
                    Next row
                End If
            Next proc
            tcps = Nothing
        Catch ex As Exception
        End Try
        For Each p In lstPorts
            Dim c As clsConnection = Me.GetDesktopConn(p)
            If c Is Nothing Then
                c = New clsConnection With {.ConnType = clsConnection.enConnType.PBIDesktop, .Port = p}
                If c.IsLiveConn = False Then
                    Me.lstConns.Add(c)
                End If
            Else
                c.Seen = True
            End If


        Next p

    End Sub



    Private Sub GetServerConns(wb As Excel.Workbook)

        For Each wc As Excel.WorkbookConnection In wb.Connections
            Dim strConn As String = ""
            Try
                If wc.Name <> "ThisWorkbookDataModel" And wc.Name.ToString.StartsWith("WorksheetConnection_") = False _
                    And wc.Description.ToString.ToLower.Trim.StartsWith("pbixl:") = False _
                    And wc.Description.ToString.ToLower.Trim.StartsWith("pbixl,") = False _
                    And wc.Description.ToLower.StartsWith("temp") = False Then
                    strConn = wc.OLEDBConnection.Connection.ToString
                    If strConn.ToLower.Trim.Contains("$embedded$") = True Then
                        strConn = ""
                    End If
                Else
                    strConn = ""
                End If
            Catch ex As Exception
                strConn = ""
            End Try

            If strConn.ToLower.Replace(" ", "").Contains("provider=msolap") Then

                Dim c As clsConnection = Me.GetServerConn(wb.Name, wc.Name)
                If c Is Nothing Then
                    If strConn.ToLower.Trim.Contains("powerbi.com") Then
                        Me.lstConns.Add(New clsConnection With {.ConnType = clsConnection.enConnType.PBIService, .WorkbookName = wb.Name, .Name = wc.Name, .ConnectionString = strConn})
                    Else
                        Me.lstConns.Add(New clsConnection With {.ConnType = clsConnection.enConnType.TabularServer, .WorkbookName = wb.Name, .Name = wc.Name, .ConnectionString = strConn})
                    End If
                Else
                    c.Seen = True
                    c.ConnectionString = strConn
                    If strConn.ToLower.Trim.Contains("powerbi.com") Then
                        c.ConnType = clsConnection.enConnType.PBIService
                    Else
                        c.ConnType = clsConnection.enConnType.TabularServer
                    End If
                End If
            End If

        Next wc


    End Sub

    Private Function GetDesktopConn(Port As String) As clsConnection
        For Each c In Me.lstConns
            If c.ConnType = clsConnection.enConnType.PBIDesktop AndAlso c.Port = Port Then
                Return c
            End If
        Next c
        Return Nothing
    End Function

    Private Function GetServerConn(wbName As String, ConnName As String) As clsConnection
        For Each c In Me.lstConns
            If c.WorkbookName.ToLower = wbName.ToLower AndAlso c.Name.ToLower = ConnName.ToLower Then
                If c.ConnType = clsConnection.enConnType.PBIService OrElse c.ConnType = clsConnection.enConnType.TabularServer Then
                    Return c
                End If
            End If
        Next c
        Return Nothing
    End Function


    Private Sub GetWorkbookConn(wb As Excel.Workbook)
        For Each c In Me.lstConns
            If c.ConnType = clsConnection.enConnType.PowerPivot AndAlso c.WorkbookName.ToLower = wb.Name.ToLower Then
                If wb.Model.ModelTables.Count = 0 Then
                    Me.lstConns.Remove(c)
                Else
                    c.Seen = True
                End If
                Exit Sub
            End If
        Next c

        If wb.Model.ModelTables.Count > 0 Then
            Me.lstConns.Add(New clsConnection With {.Name = "ThisWorkbookDataModel", .ConnType = clsConnection.enConnType.PowerPivot, .WorkbookName = wb.Name})
        End If

    End Sub




    Public Class clsConnection


        Public Property Seen As Boolean

        Public Sub New()
            Me.Seen = True
        End Sub

        Enum enConnType
            PowerPivot = 1
            PBIDesktop = 2
            PBIService = 3
            TabularServer = 4
        End Enum
        Public Property ConnType As enConnType
        Public Property WorkbookName As String
        Public Property Name As String
        Public Property Port As String

        Friend Property _ConnAlias As String

        Private _ConnectionString As String
        Public Property ConnectionString As String
            Get
                If Me.ConnType = enConnType.PBIDesktop Then
                    Return "Provider=MSOLAP;Data Source=localhost:" & Me.Port
                ElseIf Me.ConnType = enConnType.PowerPivot Then
                    Return ""
                Else
                    Return Me._ConnectionString
                End If
            End Get
            Set(value As String)
                Me._ConnectionString = value
            End Set
        End Property


        Public ReadOnly Property IsLiveConn As Boolean
            Get
                Try
                    If Me.ConnType <> enConnType.PBIDesktop Then
                        Return False
                    End If

                    Dim blnRes As Boolean = True

                    Dim conn As Object = CreateObject("ADODB.CONNECTION")
                    conn.ConnectionString = "Provider=MSOLAP;Data Source=localhost:" & Me.Port
                    conn.open
                    Dim rec As Object = CreateObject("ADODB.RECORDSET")
                    rec.open("Select * from $system.DBSCHEMA_CATALOGS", conn, 0)
                    Do While rec.eof = False
                        blnRes = False
                        Exit Do
                    Loop
                    rec.close : rec = Nothing
                    conn.close
                    Return blnRes
                Catch ex As Exception
                    Return True
                End Try
            End Get
        End Property


        Public ReadOnly Property ConnAlias As String
            Get
                Try
                    If Me.ConnType <> enConnType.PBIDesktop Then
                        Return Me.Name
                    End If

                    Dim strN As String = ""
                    Dim conn As Object = CreateObject("ADODB.CONNECTION")
                    conn.ConnectionString = "Provider=MSOLAP;Data Source=localhost:" & Me.Port
                    conn.open
                    Dim rec As Object = CreateObject("ADODB.RECORDSET")
                    rec.open("Select * from $system.MDSCHEMA_MEASURES WHERE MEASURE_NAME='pbixl'", conn, 0)
                    Do While rec.eof = False
                        strN = rec.fields("EXPRESSION").value.ToString.Replace("""", "")
                        Exit Do
                    Loop
                    rec.close : rec = Nothing
                    conn.close
                    Me._ConnAlias = strN
                Catch ex As Exception
                    Me._ConnAlias = ""
                End Try

                If Me._ConnAlias = "" Then
                    Return "localhost:" & Me.Port
                Else
                    Return Me._ConnAlias
                End If


            End Get
        End Property



    End Class



End Class
