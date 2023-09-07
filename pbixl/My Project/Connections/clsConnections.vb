﻿Imports Excel = Microsoft.Office.Interop.Excel

Public Class clsConnections

    Public Connections As List(Of clsConnection)
    Private wb As Excel.Workbook

    Public Sub New(wb As Excel.Workbook)
        Me.wb = wb
        Me.Connections = New List(Of clsConnection)
    End Sub

    Public Sub Refresh()

        Dim lstGuids As New List(Of String)
        Dim qm As New clsQryMgr(Me.wb)
        For Each q As clsQuery In qm.Queries
            lstGuids.Add(q.GUID.Substring(0, 8))
        Next q
        qm = Nothing


        For Each wc As Excel.WorkbookConnection In Me.wb.Connections
            Dim c As clsConnection = New clsConnection(wc.Name, wc.OLEDBConnection.Connection, "", wc.Description)
            If c.ConnectionString.ToLower.Trim.Contains("msolap") Then
                Dim blnFound As Boolean = False

                For Each s As String In lstGuids
                    If wc.Description.Contains(s) Then
                        blnFound = True
                        Exit For
                    End If
                Next s


                If blnFound = False Then
                    Me.Connections.Add(c)
                End If

            End If
        Next wc

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
            Dim c As New clsConnection("", "", p, "")
            Me.Connections.Add(c)
        Next p

    End Sub


    Public Class clsConnection

        Public Property Port As String
        Public Property Description As String
        Public Property Err As Exception
        Public Property CubeName As String
        Public Property LastSchemaUpdate As DateTime
        Public Property LastDataUpdate As DateTime
        Public Property ConnAlias As String
        Public Property Cubes As List(Of clsCube)

        Private _Name As String
        Public Property Name As String
            Get
                If Me.ConnType = enConnType.PBIDesktop And Me.ConnAlias = "" Then
                    Return "Port:" & Me.Port.ToString
                ElseIf Me.ConnType = enConnType.PBIDesktop And Me.ConnAlias <> "" Then
                    Return Me.ConnAlias & " (Port:" & Me.Port & ")"
                Else
                    Return Me._Name
                End If
            End Get
            Set(value As String)
                Me._Name = value
            End Set
        End Property

        Public ReadOnly Property BASECUBENAME As String
            Get
                For Each c As clsCube In Me.Cubes
                    If c.BASE_CUBE_NAME.ToString = "" Then
                        Return c.NAME
                    End If
                Next c
                Return Nothing
            End Get
        End Property



        Private _nd As C1.Win.C1FlexGrid.Node
        Public Property nd As C1.Win.C1FlexGrid.Node
            Get
                Return _nd
            End Get
            Set(value As C1.Win.C1FlexGrid.Node)


                Try
                    Me._nd = value
                    Me._nd.Row.Grid.SetData(Me.nd.GetCellRange.r1, 0, Me.Name)
                    Me._nd.Image = Me.Image
                    Me._nd.Row.Grid.Rows(_nd.GetCellRange.r1).Visible = True

                    If Me.ConnType = enConnType.PBIDesktop And Me.Cubes.Count = 0 Then
                        Me._nd.Row.Grid.Rows(Me._nd.GetCellRange.r1).Visible = False
                    End If
                Catch ex As Exception

                End Try






            End Set
        End Property

        Public Enum enConnState
            OK = 1
            NotOK = 2
            Unknown = 3
        End Enum
        Public Property ConnState As enConnState

        Public Enum enConnType
            PowerBI = 1
            PBIDesktop = 2
            TabularSvr = 3
        End Enum

        Public ReadOnly Property ConnType As enConnType
            Get
                If Me.Port <> "" Then
                    Return enConnType.PBIDesktop
                ElseIf Me._ConnectionString.ToLower.Contains("powerbi.com") Then
                    Return enConnType.PowerBI
                Else
                    Return enConnType.TabularSvr
                End If
            End Get
        End Property

        Public ReadOnly Property Image As Image
            Get
                If Me.ConnType = enConnType.PBIDesktop And Me.ConnState = enConnState.Unknown Then
                    Return TryCast(Me._nd.Row.Grid.Parent, dlgSelectConn).ImageList.Images("PBIDesktop_Unknown.ico")
                ElseIf Me.ConnType = enConnType.PowerBI And Me.ConnState = enConnState.Unknown Then
                    Return TryCast(Me._nd.Row.Grid.Parent, dlgSelectConn).ImageList.Images("PowerBI_Unknown.ico")
                ElseIf Me.ConnType = enConnType.TabularSvr And Me.ConnState = enConnState.Unknown Then
                    Return TryCast(Me._nd.Row.Grid.Parent, dlgSelectConn).ImageList.Images("TabularSvr_Unknown.ico")

                ElseIf Me.ConnType = enConnType.PBIDesktop And Me.ConnState = enConnState.OK Then
                    Return TryCast(Me._nd.Row.Grid.Parent, dlgSelectConn).ImageList.Images("PBIDesktop_OK.ico")
                ElseIf Me.ConnType = enConnType.PowerBI And Me.ConnState = enConnState.OK Then
                    Return TryCast(Me._nd.Row.Grid.Parent, dlgSelectConn).ImageList.Images("PowerBI_OK.ico")
                ElseIf Me.ConnType = enConnType.TabularSvr And Me.ConnState = enConnState.OK Then
                    Return TryCast(Me._nd.Row.Grid.Parent, dlgSelectConn).ImageList.Images("TabularSvr_OK.ico")
                ElseIf Me.ConnType = enConnType.PBIDesktop And Me.ConnState = enConnState.NotOK Then
                    Return TryCast(Me._nd.Row.Grid.Parent, dlgSelectConn).ImageList.Images("PBIDesktop_NotOK.ico")
                ElseIf Me.ConnType = enConnType.PowerBI And Me.ConnState = enConnState.NotOK Then
                    Return TryCast(Me._nd.Row.Grid.Parent, dlgSelectConn).ImageList.Images("PowerBI_NotOK.ico")
                ElseIf Me.ConnType = enConnType.TabularSvr And Me.ConnState = enConnState.NotOK Then
                    Return TryCast(Me._nd.Row.Grid.Parent, dlgSelectConn).ImageList.Images("TabularSvr_NotOK.ico")

                End If
                Return Nothing
            End Get
        End Property


        Private _ConnectionString As String
        Public Property ConnectionString
            Get
                If Me.ConnType = enConnType.PBIDesktop Then
                    Return "Provider=MSOLAP;Data Source=localhost:" & Me.Port
                Else
                    Return Me._ConnectionString.Substring(6)
                End If
            End Get
            Set(value)
                Me._ConnectionString = value
            End Set
        End Property



        Public Sub New(Name As String, ConnectionString As String, Port As String, Description As String)

            Me.Name = Name
            Me.Description = Description
            Me.ConnectionString = ConnectionString
            Me.Port = Port
            Me.ConnState = enConnState.Unknown
            Me.Cubes = New List(Of clsCube)
        End Sub

        Public Sub TestConnection()

            Dim t = Task(Of Object).Factory.StartNew(Function()
                                                         Try
                                                             Dim conn As Object = CreateObject("ADODB.CONNECTION")
                                                             conn.connectionstring = Me.ConnectionString
                                                             conn.open

                                                             Dim rec As Object = CreateObject("ADODB.RECORDSET")
                                                             rec.open("Select * from $system.MDSCHEMA_CUBES", conn, 0)

                                                             Me.ConnAlias = ""
                                                             Me.Cubes = New List(Of clsCube)

                                                             Do While rec.eof = False
                                                                 Me.Cubes.Add(New clsCube(rec.fields("CUBE_NAME").value.ToString, rec.fields("CUBE_CAPTION").value.ToString,
                                                                                          rec.fields("BASE_CUBE_NAME").value.ToString))
                                                                 rec.movenext
                                                             Loop


                                                             rec.close

                                                             If Me.ConnType = enConnType.PBIDesktop Then
                                                                 rec.open("Select * from $system.MDSCHEMA_MEASURES WHERE MEASURE_NAME='pbixl'", conn, 0)
                                                                 Do While rec.eof = False
                                                                     Me.ConnAlias = rec.fields("EXPRESSION").value.ToString.Replace("""", "")
                                                                     Exit Do
                                                                 Loop
                                                                 rec.close : rec = Nothing
                                                             End If

                                                             conn.close : conn = Nothing

                                                             If Me.ConnType = enConnType.PBIDesktop And Me.Cubes.Count = 0 Then
                                                                 Me.ConnState = enConnState.NotOK
                                                             Else
                                                                 Me.ConnState = enConnState.OK
                                                             End If

                                                             Me.nd = Me.nd
                                                             Return Nothing
                                                         Catch ex As Exception

                                                             'retry
                                                             If Me.ConnType = enConnType.PBIDesktop And Me.Cubes.Count > 0 And Me.Err Is Nothing Then
                                                                 System.Threading.Thread.Sleep(1000)
                                                                 Me.TestConnection()
                                                             End If


                                                             Me.Err = ex
                                                             Me.ConnState = enConnState.NotOK
                                                             Me.nd = Me.nd
                                                             Return Nothing
                                                         End Try
                                                     End Function)



        End Sub

    End Class


    Public Class clsCube

        Public Property NAME As String
        Public Property CAPTION As String
        Public Property BASE_CUBE_NAME As String

        Public Sub New(NAME As String, CAPTION As String, BASE_CUBE_NAME As String)
            Me.NAME = NAME
            Me.CAPTION = CAPTION
            Me.BASE_CUBE_NAME = BASE_CUBE_NAME
        End Sub
    End Class



End Class
