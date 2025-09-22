Imports System.Threading
Imports ExcelDna.Integration
Imports Excel = Microsoft.Office.Interop.Excel

Public Module pbixl_DaxQuery



    Friend lstWbConns As List(Of (WbName As String, ConnName As String, ConnString As String))

    Friend ConnectionsX As clsConnectionsX


    <ExcelFunction(Description:="pbixl", Category:="pbixl", IsMacroType:=False, IsVolatile:=False, IsHidden:=False, Name:="pbixl.DAX")>
    Public Function pbixl_DaxQuery(Connection As Object, Statement As Object, Hint As Object) As Object

        If TypeOf Statement Is ExcelDna.Integration.ExcelMissing Then
            Return "# pbixl Error: function expects parameter for connection and statement."
        End If



        Dim ConnString As String = ""
        Dim Stmnt As String = ""
        Dim FuncHint As String = ""

        If ConnectionsX Is Nothing Then
            ConnectionsX = New clsConnectionsX
        End If

        If lstWbConns Is Nothing Then lstWbConns = New List(Of (WbName As String, ConnName As String, ConnString As String))

        If TypeOf Connection Is ExcelDna.Integration.ExcelEmpty OrElse TypeOf Connection Is ExcelDna.Integration.ExcelMissing Then
            ConnString = ""
        Else
            ConnString = Connection
        End If

        If TypeOf Statement Is ExcelDna.Integration.ExcelEmpty OrElse TypeOf Statement Is ExcelDna.Integration.ExcelMissing Then
            Stmnt = ""
        Else
            Stmnt = Statement
        End If

        If Stmnt = "" Then
            Return "#pbixl Error: Statment missing."
        End If


        If TypeOf Hint Is ExcelDna.Integration.ExcelEmpty OrElse TypeOf Hint Is ExcelDna.Integration.ExcelMissing Then
            FuncHint = ""
        Else
            FuncHint = Hint
        End If


        If ConnString = "" Then
            If ConnectionsX.PortCount = 0 Then
                Return "#pbixl Error: No open PBI Workbook found."
            ElseIf ConnectionsX.PortCount > 1 Then
                Return "#pbixl Error: " & ConnectionsX.PortCount & " PBI Workbooks found. Please ensure only one Workbook is open. Or provide a names."
            Else
                ConnString = "Provider=MSOLAP;Data Source=localhost:" & ConnectionsX.DefaultPort
                Dim res As Object = ExcelTaskUtil.RunTask("GetQuery", New Object() {ConnString, Statement, FuncHint, Nothing}, Function(ct) pbixl_DAX_Async(ConnString, Statement, FuncHint, ct))
                If res Is Nothing Then
                    Return ""
                ElseIf res.Equals(ExcelDna.Integration.ExcelError.ExcelErrorNA) Then
                    Return ExcelDna.Integration.ExcelError.ExcelErrorGettingData
                Else
                    If Not TryCast(res, Exception) Is Nothing Then
                        Return TryCast(res, Exception).Message
                    Else
                        If IsDBNull(res) Then
                            Return ExcelDna.Integration.ExcelError.ExcelErrorNull
                        Else
                            Return res
                        End If
                    End If
                End If
            End If
        End If

        If IsNumeric(ConnString) = True Then
            If ConnectionsX.HasPort(ConnString) = True Then
                ConnString = "Provider=MSOLAP;Data Source=localhost:" & ConnString
            Else
                Return "#pbixl Error: No PBI Workbook found on port " & ConnString
            End If
        ElseIf ConnString.ToLower.StartsWith("localhost:") AndAlso ConnString.ToLower.Length > 10 AndAlso isnumeric(ConnString.Substring(10)) Then
            If ConnectionsX.HasPort(ConnString.Substring(10)) = True Then
                ConnString = "Provider=MSOLAP;Data Source=localhost:" & ConnString.Substring(10)
            Else
                Return "#pbixl Error: No PBI Workbook found on port " & ConnString
            End If
        Else
            Dim strPort As String = ConnectionsX.GetPortNumber(ConnString)
            If strPort <> "" Then
                ConnString = "Provider=MSOLAP;Data Source=localhost:" & strPort
            Else
                If ConnString.Trim = "" Then
                    Return "#pbixl Error: Please Provide a name for the connection."
                End If

                Dim caller As ExcelReference = XlCall.Excel(XlCall.xlfCaller)
                Dim wsName As String = XlCall.Excel(XlCall.xlSheetNm, caller)
                Dim wb As Excel.Workbook = Nothing
                For Each wb In MyAddin.App.Workbooks
                    If wsName.ToLower.StartsWith("[" & wb.Name.ToLower & "]") Then
                        Exit For
                    End If
                Next wb


                For i As Integer = lstWbConns.Count - 1 To 0 Step -1
                    If lstWbConns.Item(i).WbName.ToLower = wb.Name.ToLower Then
                        lstWbConns.RemoveAt(i)
                    End If
                Next i
                For Each wc As Excel.WorkbookConnection In wb.Connections
                    Try
                        If wc.OLEDBConnection.Connection.ToString.ToLower.Contains("msolap") Then
                            lstWbConns.Add((wb.Name, wc.Name, wc.OLEDBConnection.Connection))
                        End If
                    Catch ex As Exception
                    End Try
                Next wc

                If lstWbConns.Count = 0 Then
                    Return "#pbixl Error: No connections found in this workbook."
                End If


                Dim blnFound As Boolean = False
                For Each c In lstWbConns
                    If c.ConnName.ToLower = ConnString.ToLower Then
                        ConnString = c.ConnString
                        blnFound = True
                        Exit For
                    End If
                Next c

                If blnFound = False Then
                    Return "#pbixl Error: No connection found."
                End If


            End If
        End If


        Dim _res As Object = ExcelTaskUtil.RunTask("GetQuery", New Object() {ConnString, Statement, FuncHint, Nothing}, Function(ct) pbixl_DAX_Async(ConnString, Statement, FuncHint, ct))

        If _res Is Nothing Then
            Return ""
        ElseIf _res.Equals(ExcelDna.Integration.ExcelError.ExcelErrorNA) Then
            Return ExcelDna.Integration.ExcelError.ExcelErrorGettingData
        Else
            If Not TryCast(_res, Exception) Is Nothing Then
                Return TryCast(_res, Exception).Message
            Else
                If IsDBNull(_res) Then
                    Return ExcelDna.Integration.ExcelError.ExcelErrorNull
                Else
                    Return _res
                End If
            End If
        End If




    End Function


    Private Function pbixl_DAX_Async(strConn As String, strDax As String, Hint As String, ct As CancellationToken) As Task(Of Object)
        Dim task1 = Task(Of Object).Factory.StartNew(Function()
                                                         Try

                                                             Dim blnNoHeader As Boolean = False
                                                             If Hint.ToString.ToLower.Contains("noheader") Then
                                                                 blnNoHeader = True
                                                             End If

                                                             Dim blnTranspose As Boolean = False
                                                             If Hint.ToString.ToLower.Contains("transpose") Then
                                                                 blnTranspose = True
                                                             End If

                                                             Dim blnDone As Boolean = False
                                                             Dim rec As Object = Nothing

                                                             ct.Register(Function()
                                                                             If blnDone = False Then
                                                                                 Try
                                                                                     If Not rec Is Nothing Then
                                                                                         If rec.state = 4 Then
                                                                                             rec.cancel
                                                                                         ElseIf rec.state = 1 Then
                                                                                             rec.close
                                                                                         End If
                                                                                         rec = Nothing
                                                                                     End If
                                                                                 Catch ex As Exception
                                                                                     rec = Nothing
                                                                                 End Try
                                                                                 'Return "!!Canceled"
                                                                             End If
                                                                             rec = Nothing
                                                                             'Return "!!Not Canceled"
                                                                             Return Nothing
                                                                         End Function)


                                                             Dim objRes As Object = Nothing

                                                             Try

                                                                 If strConn.ToLower.Trim.StartsWith("oledb;") Then
                                                                     strConn = strConn.Substring(6)
                                                                 End If

                                                                 'Dim c As Object = MyAddin.ConnMgr.GetADOConnection(strConn)
                                                                 Dim c As Object = CreateObject("ADODB.CONNECTION")
                                                                 c.connectionstring = strConn
                                                                 c.open

                                                                 Try
                                                                     If c.state <> 1 Then
                                                                         c.open
                                                                     End If
                                                                 Catch ex As Exception
                                                                     c = Nothing
                                                                     Throw (New Exception("can not obtain connection"))
                                                                 End Try


                                                                 rec = CreateObject("ADODB.RECORDSET")
                                                                 rec.open(strDax, c)

                                                                 Dim strRes(,)

                                                                 If blnTranspose = False Then
                                                                     Dim ctr As Integer = -1
                                                                     If rec.recordcount = 0 Then
                                                                         If blnNoHeader = False Then
                                                                             ReDim strRes(rec.recordcount, rec.fields.count - 1)
                                                                             For i As Integer = 0 To rec.fields.count - 1
                                                                                 strRes(0, i) = rec.fields(i).name
                                                                             Next i
                                                                         Else
                                                                             ReDim strRes(0, 0)
                                                                             strRes(0, 0) = ""
                                                                         End If
                                                                     ElseIf blnNoHeader = True Then
                                                                         ctr = -1
                                                                         ReDim strRes(rec.recordcount - 1, rec.fields.count - 1)
                                                                         Do While rec.eof = False
                                                                             ctr += 1
                                                                             For i As Integer = 0 To rec.fields.count - 1
                                                                                 If rec.fields(i).value Is DBNull.Value Then
                                                                                     strRes(ctr, i) = ""
                                                                                 Else
                                                                                     strRes(ctr, i) = rec.fields(i).value
                                                                                 End If
                                                                             Next i
                                                                             rec.movenext
                                                                         Loop
                                                                     Else
                                                                         ctr = 0
                                                                         ReDim strRes(rec.recordcount, rec.fields.count - 1)
                                                                         For i As Integer = 0 To rec.fields.count - 1
                                                                             strRes(0, i) = rec.fields(i).name
                                                                         Next i
                                                                         Do While rec.eof = False
                                                                             ctr += 1
                                                                             For i As Integer = 0 To rec.fields.count - 1
                                                                                 If rec.fields(i).value Is DBNull.Value Then
                                                                                     strRes(ctr, i) = ""
                                                                                 Else
                                                                                     strRes(ctr, i) = rec.fields(i).value
                                                                                 End If
                                                                             Next i
                                                                             rec.movenext
                                                                         Loop
                                                                     End If



                                                                 Else
                                                                     Dim ctr As Integer = -1
                                                                     If rec.recordcount = 0 Then
                                                                         If blnNoHeader = False Then
                                                                             ReDim strRes(rec.fields.count - 1, rec.recordcount)
                                                                             For i As Integer = 0 To rec.fields.count - 1
                                                                                 strRes(i, 0) = rec.fields(i).name
                                                                             Next i
                                                                         Else
                                                                             ReDim strRes(0, 0)
                                                                             strRes(0, 0) = ""
                                                                         End If
                                                                     ElseIf blnNoHeader = True Then
                                                                         ctr = -1
                                                                         ReDim strRes(rec.fields.count - 1, rec.recordcount - 1)
                                                                         Do While rec.eof = False
                                                                             ctr += 1
                                                                             For i As Integer = 0 To rec.fields.count - 1
                                                                                 If rec.fields(i).value Is DBNull.Value Then
                                                                                     strRes(i, ctr) = ""
                                                                                 Else
                                                                                     strRes(i, ctr) = rec.fields(i).value
                                                                                 End If
                                                                             Next i
                                                                             rec.movenext
                                                                         Loop
                                                                     Else
                                                                         ctr = 0
                                                                         ReDim strRes(rec.fields.count - 1, rec.recordcount)
                                                                         For i As Integer = 0 To rec.fields.count - 1
                                                                             strRes(i, 0) = rec.fields(i).name
                                                                         Next i
                                                                         Do While rec.eof = False
                                                                             ctr += 1
                                                                             For i As Integer = 0 To rec.fields.count - 1
                                                                                 If rec.fields(i).value Is DBNull.Value Then
                                                                                     strRes(i, ctr) = ""
                                                                                 Else
                                                                                     strRes(i, ctr) = rec.fields(i).value
                                                                                 End If
                                                                             Next i
                                                                             rec.movenext
                                                                         Loop
                                                                     End If



                                                                 End If

                                                                 rec.close : rec = Nothing
                                                                 Return strRes



                                                             Catch ex As Exception
                                                                 Throw ex
                                                             End Try

                                                             blnDone = True
                                                             Return objRes

                                                         Catch xex As Exception
                                                             Return xex
                                                         End Try
                                                     End Function)
        Return task1
    End Function







End Module



Friend Class clsConnectionsX

    Dim lstPorts As List(Of (PortName As String, CubeName As String, ConnAlias As String))
    Public LastRefresh As Date

    Public ReadOnly Property PortCount As Integer
        Get
            If Now().Subtract(Me.LastRefresh).TotalMilliseconds > 1000 Then
                Me.RefreshPorts()
            End If
            Return Me.lstPorts.Count
        End Get
    End Property

    Public ReadOnly Property HasPort(Port As String) As Boolean
        Get
            If Now().Subtract(Me.LastRefresh).TotalMilliseconds > 1000 Then
                Me.RefreshPorts()
            End If
            For Each p In Me.lstPorts
                If p.PortName.Trim = Port.Trim Then
                    Return True
                End If
            Next p
            Return False
        End Get
    End Property

    Public ReadOnly Property GetPortNumber(Name As String) As String
        Get
            If Now().Subtract(Me.LastRefresh).TotalMilliseconds > 1000 Then
                Me.RefreshPorts()
            End If
            For Each p In Me.lstPorts
                If p.ConnAlias.ToLower = Name.ToLower Then
                    Return p.PortName
                End If
            Next p
            Return ""
        End Get
    End Property



    Public ReadOnly Property DefaultPort As String
        Get

            If Now().Subtract(Me.LastRefresh).TotalMilliseconds > 1000 Then
                Me.RefreshPorts()
            End If
            If Me.lstPorts.Count = 1 Then
                Return lstPorts.Item(0).PortName
            Else
                Return lstPorts.Item(0).PortName
            End If

        End Get
    End Property


    Public Sub New()

        Me.lstPorts = New List(Of (PortName As String, CubeName As String, ConnAlias As String))
        Me.RefreshPorts()

        'Debug.Print(Me.lstPorts.Count)

    End Sub

    Private Sub RefreshPorts()

        'Debug.Print("Refresh")

        Me.lstPorts.Clear()


        Dim objWMI As Object
        Dim objProcess As Object
        Dim strProcessName As String
        Dim strCmd As String = ""
        Dim strPort As String = ""
        Dim strCube As String = ""
        Dim strAlias As String = ""
        strProcessName = "msmdsrv.exe"
        objWMI = GetObject("winmgmts:\\.\root\cimv2")
        For Each objProcess In objWMI.ExecQuery("SELECT * FROM Win32_Process WHERE Name='" & strProcessName & "'")
            strCmd = objProcess.CommandLine.ToString
            If strCmd <> "" Then
                For i = Len(strCmd) - 1 To 0 Step -1
                    If Mid(strCmd, i, 1) = """" Then
                        strCmd = strCmd.Substring(i).Replace("""", "")
                        If System.IO.File.Exists(strCmd & "\msmdsrv.port.txt") Then
                            strPort = My.Computer.FileSystem.ReadAllText(strCmd & "\msmdsrv.port.txt", System.Text.Encoding.Unicode).ToString
                            Dim conn As Object = CreateObject("ADODB.CONNECTION")
                            conn.connectionstring = "Provider=MSOLAP;Data Source=localhost:" & strPort
                            conn.open
                            Dim rec As Object = CreateObject("ADODB.RECORDSET")
                            rec.open("Select * from $system.MDSCHEMA_CUBES", conn)
                            strCube = ""
                            Do While rec.eof = False
                                If rec.Fields("BASE_CUBE_NAME").Value.ToString.Trim = "" Then
                                    strCube = rec.fields("CUBE_NAME").Value.ToString
                                End If
                                rec.movenext
                            Loop
                            rec.close

                            strAlias = ""
                            rec.open("Select * from $system.MDSCHEMA_MEASURES where MEASURE_NAME='pbixl'")
                            Do While rec.eof = False
                                strAlias = rec.fields("EXPRESSION").Value.ToString
                                If strAlias.StartsWith("""") Then strAlias = strAlias.Substring(1)
                                If strAlias.EndsWith("""") Then strAlias = strAlias.Substring(0, strAlias.Length - 1)
                                Exit Do
                            Loop

                            conn.close

                            If strCube.Trim <> "" Then
                                Me.lstPorts.Add((strPort, strCube, strAlias))
                            End If

                        End If
                        Exit For
                    End If
                Next i
            End If
        Next objProcess

        Me.LastRefresh = Now


    End Sub


End Class

