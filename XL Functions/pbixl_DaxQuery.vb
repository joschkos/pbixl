Imports System.Threading
Imports ExcelDna.Integration
Imports Excel = Microsoft.Office.Interop.Excel

Public Module pbixl_DaxQuery

    <ExcelFunction(Description:="pbixl", Category:="pbixl", IsMacroType:=True, IsVolatile:=False, IsHidden:=False, Name:="pbixl.DAX")>
    Public Function pbixl_Query(ConnName As Object, Statement As Object, Hint As Object) As Object

        Dim blnNoHeader As Boolean = False
        If TypeOf Hint Is ExcelDna.Integration.ExcelMissing Then
            Hint = ""
            blnNoHeader = False
        Else
            If Hint.ToString.ToLower.Contains("noheader") Then
                blnNoHeader = True
            End If
        End If

        Dim blnTranspose As Boolean = False
        If TypeOf Hint Is ExcelDna.Integration.ExcelMissing Then
            Hint = ""
            blnTranspose = False
        Else
            If Hint.ToString.ToLower.Contains("transpose") Then
                blnTranspose = True
            End If
        End If


        If MyAddIn.DynamicArray = False Then
            Return "Dynamic Arrays Not supported by this version of Excel."
        End If

        If TypeOf ConnName Is ExcelDna.Integration.ExcelMissing Then
            Return "Connection missing"
        End If
        If ConnName.ToString.Trim = "" Then
            Return "Connection missing"
        End If
        If TypeOf Statement Is ExcelMissing Then
            Return "Statement missing"
        End If
        If Statement.ToString.Trim = "" Then
            Return "Statement missing"
        End If

        If ExcelDnaUtil.IsInFunctionWizard = True Then
            Return ""
        End If

        Dim wb As Excel.Workbook = MyAddIn.GetWorkbook(XlCall.Excel(XlCall.xlfGetWorkbook, 16))

        If ConnName.ToString.ToLower.Trim = "thisworkbookdatamodel" Then

            Dim rec As Object = CreateObject("ADODB.RECORDSET")

            Try


                Dim connWb As Object = wb.Model.DataModelConnection.ModelConnection.ADOConnection
                rec.open(Statement, connWb)

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

                connWb = Nothing
                rec.close : rec = Nothing

                Return strRes



            Catch ex As Exception
                If Not rec Is Nothing Then
                    rec = Nothing
                End If
            End Try
        End If


        Dim connString As Object = MyAddIn.Connections.GetConnectionString(wb.Name, ConnName)
        If connString = "" Then
            MyAddIn.Connections.Refresh(wb)
        End If
        connString = MyAddIn.Connections.GetConnectionString(wb.Name, ConnName)
        If connString.ToString.Trim = "" Then
            Return "Connection Not found."
        End If


        Dim res As Object = ExcelTaskUtil.RunTask("GetQuery", New Object() {connString, Statement, Hint, Nothing}, Function(ct) pbixl_DAX_Async(wb.Name, connString, Statement, Hint, ct))


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



    Private Function pbixl_DAX_Async(strWbName As String, strConn As String, strDax As String, Hint As String, ct As CancellationToken) As Task(Of Object)
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

                                                                 Dim c As Object = MyAddIn.Connections.GetConnection(strConn)


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
