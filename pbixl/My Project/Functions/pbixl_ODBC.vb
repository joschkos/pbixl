
Imports System.Threading
Imports ExcelDna.Integration
Imports Excel = Microsoft.Office.Interop.Excel

Public Module pbixl_ODBC

    <ExcelFunction(Description:="pbixl ODBC", Category:="pbixl", IsMacroType:=False, IsVolatile:=False, IsHidden:=False, Name:="pbixl.ODBC")>
    Public Function pbixl_ODBC(Connection As Object, Statement As Object, Hint As Object) As Object

        If TypeOf Connection Is ExcelDna.Integration.ExcelMissing Or TypeOf Connection Is ExcelDna.Integration.ExcelEmpty Then
            Return "#pbixl Error: Please provide a server name, database name, and a query."
        End If

        If TypeOf Statement Is ExcelDna.Integration.ExcelMissing Or TypeOf Statement Is ExcelDna.Integration.ExcelEmpty Then
            Return "#pbixl Error: Please provide a server name, database name, and a query."
        End If

        If TypeOf Hint Is ExcelDna.Integration.ExcelMissing Or TypeOf Hint Is ExcelDna.Integration.ExcelEmpty Then
            Hint = ""
        End If


        Dim strConn As String = Connection.ToString
        Dim strStatement As String = Statement.ToString
        Dim strHint As String = Hint.ToString

        Dim res As Object = ExcelTaskUtil.RunTask("GetQuery", New Object() {strConn, Statement, strHint, Nothing}, Function(ct) pbixl_ODBC_Async(strConn, strStatement, strHint, ct))

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


    End Function


    Private Function pbixl_ODBC_Async(strConn As String, strStatement As String, strHint As String, ct As CancellationToken) As Task(Of Object)
        Dim task1 = Task(Of Object).Factory.StartNew(Function()
                                                         Try

                                                             Dim blnNoHeader As Boolean = False
                                                             If strHint.ToString.ToLower.Contains("noheader") Then
                                                                 blnNoHeader = True
                                                             End If



                                                             Dim blnTranspose As Boolean = False


                                                             Dim blnDone As Boolean = False
                                                             Dim conn As Odbc.OdbcConnection = Nothing
                                                             Dim cmd As Odbc.OdbcCommand = Nothing
                                                             Dim dr As Odbc.OdbcDataReader = Nothing


                                                             ct.Register(Function()
                                                                             If blnDone = False Then
                                                                                 Try
                                                                                     If Not cmd Is Nothing Then
                                                                                         cmd.Cancel()
                                                                                     End If
                                                                                     cmd = Nothing
                                                                                 Catch ex As Exception
                                                                                     cmd = Nothing
                                                                                 End Try
                                                                                 'Return "!!Canceled"
                                                                             End If
                                                                             cmd = Nothing
                                                                             'Return "!!Not Canceled"
                                                                             Return Nothing
                                                                         End Function)


                                                             Dim objRes As Object = Nothing

                                                             Try

                                                                 If strConn.ToLower.Trim.StartsWith("oledb;") Then
                                                                     strConn = strConn.Substring(6)
                                                                 End If

                                                                 conn = New Odbc.OdbcConnection
                                                                 conn.ConnectionString = strConn

                                                                 Try
                                                                     conn.Open()
                                                                 Catch ex As Exception
                                                                 End Try
                                                                 If conn.State.ToString.ToLower <> "open" Then
                                                                     conn.Open()
                                                                 End If

                                                                 cmd = New Odbc.OdbcCommand(strStatement, conn)
                                                                 dr = cmd.ExecuteReader
                                                                 Dim dt As New DataTable()
                                                                 Try
                                                                     dt.Load(dr)
                                                                 Catch ex As Exception
                                                                     If TypeOf (ex) Is System.Data.ConstraintException Then
                                                                         For Each c As DataColumn In dt.Columns
                                                                             If c.DataType = System.Type.GetType("System.String") Then
                                                                                 c.MaxLength = -1
                                                                             End If
                                                                         Next c
                                                                         Try
                                                                             dt.Load(dr)
                                                                         Catch xex As Exception
                                                                             Return "#pbixl Error: " & dt.GetErrors()(0).RowError.ToString & " ColumnName:" & dt.Columns(0).ColumnName & " Unique:" & dt.Columns(0).Unique.ToString & " Type:" & dt.Columns(0).DataType.ToString & " Max Length:" & dt.Columns(0).MaxLength
                                                                         End Try
                                                                     Else
                                                                         Throw (ex)
                                                                     End If


                                                                 End Try

                                                                 dr.Close()

                                                                 Dim strRes(,)
                                                                 If blnNoHeader = True Then
                                                                     ReDim strRes(dt.Rows.Count - 1, dt.Columns.Count - 1)
                                                                 Else
                                                                     ReDim strRes(dt.Rows.Count - 1 + 1, dt.Columns.Count - 1)
                                                                     For i As Integer = 0 To dt.Columns.Count - 1
                                                                         strRes(0, i) = dt.Columns(i).ColumnName
                                                                     Next i
                                                                 End If


                                                                 Dim intPos As Integer = If(blnNoHeader = False, 1, 0)
                                                                 For i As Integer = 0 To dt.Rows.Count - 1
                                                                     Dim itemArray As Object() = dt.Rows(i).ItemArray
                                                                     For j As Integer = 0 To dt.Columns.Count - 1
                                                                         If itemArray(j) Is DBNull.Value Then
                                                                             strRes(i + intPos, j) = ""
                                                                         Else
                                                                             strRes(i + intPos, j) = itemArray(j)
                                                                         End If
                                                                     Next j
                                                                 Next i

                                                                 dt.Dispose()
                                                                 cmd.Dispose()
                                                                 conn.Close()

                                                                 blnDone = True

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







