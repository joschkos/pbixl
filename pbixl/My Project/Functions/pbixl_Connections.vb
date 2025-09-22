Imports ExcelDna.Integration
'Imports ExcelDna.Integration
'Imports System.Runtime.InteropServices
'Imports ExcelDna.Integration.CustomUI
Imports Excel = Microsoft.Office.Interop.Excel
'Imports System.Runtime.CompilerServices

Public Module pbixl_Connections

    Private Connections As clsConnections


    <ExcelFunction(Description:="pbixl", Category:="pbixl", IsMacroType:=True, IsVolatile:=False, IsHidden:=False, Name:="pbixl.Connections")>
    Public Function pbixl_Connections(Options As String) As Object

        If ExcelDnaUtil.IsInFunctionWizard = True Then
            Return ""
        End If


        Dim caller As ExcelReference = XlCall.Excel(XlCall.xlfCaller)
        Dim wsName As String = XlCall.Excel(XlCall.xlSheetNm, caller)
        Dim wb As Excel.Workbook = Nothing
        Dim lstConns As New List(Of (Type As String, Name As String, DataSource As String))
        For Each wb In MyAddin.App.Workbooks
            If wsName.ToLower.StartsWith("[" & wb.Name.ToLower & "]") Then
                Exit For
            End If
        Next wb
        For Each wc As Excel.WorkbookConnection In wb.Connections
            Try
                If wc.OLEDBConnection.Connection.ToString.ToLower.StartsWith("oledb") Then
                    Dim bx As New OleDb.OleDbConnectionStringBuilder(wc.OLEDBConnection.Connection)
                    Dim strDataSource As String = ""
                    Dim strApp As String = ""
                    Dim strAppName As String = ""
                    If bx.ContainsKey("App") Then strApp = bx.Item("App")
                    If bx.ContainsKey("Application Name") Then strAppName = bx.Item("Application Name")
                    If bx.ContainsKey("Data Source") Then strDataSource = bx.Item("Data Source")
                    If bx.ContainsKey("DataSource") Then strDataSource = bx.Item("DataSource")

                    If strApp.ToLower.StartsWith("pbixl") = False And strAppName.ToLower.StartsWith("pbixl") = False Then
                        If strDataSource.ToLower.StartsWith("pbiazure") Then
                            lstConns.Add(("PBI Service", wc.Name, bx.Item("Data Source")))
                        Else
                            lstConns.Add(("Tabular Server", wc.Name, bx.Item("Data Source")))
                        End If
                    End If
                End If

            Catch ex As Exception

            End Try
        Next wc




        Dim lstPorts As New List(Of (Type As String, Name As String, DataSource As String))
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
                                lstPorts.Add(("PBI Desktop", strAlias, strPort))
                            End If



                        End If
                        Exit For
                    End If
                Next i
            End If
        Next objProcess

        Dim strR(lstPorts.Count + lstConns.Count, 2)
        strR(0, 0) = "Type"
        strR(0, 1) = "Name"
        strR(0, 2) = "Data Source"
        Dim ctr As Integer = 0
        For Each p In lstPorts
            ctr += 1
            strR(ctr, 0) = p.Type
            strR(ctr, 1) = p.Name
            strR(ctr, 2) = "localhost:" & p.DataSource
        Next p
        For Each c In lstConns
            ctr += 1
            strR(ctr, 0) = c.Type
            strR(ctr, 1) = c.Name
            strR(ctr, 2) = c.DataSource
        Next


        Return strR



    End Function


End Module
