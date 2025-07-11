Imports Microsoft.Office.Interop

Public Class clsQryMgr

    Public GlobalQueries As List(Of clsQuery)

    Public Sub AddGlobalQuery(nq As clsQuery)
        For i As Integer = Me.GlobalQueries.Count - 1 To 0 Step -1
            If Me.GlobalQueries.Item(i).GUID = nq.GUID Then
                Me.GlobalQueries.Item(i) = nq.Clone
                Exit Sub
            End If
        Next i
        Me.GlobalQueries.Add(nq.Clone)
    End Sub




    Public Sub New()
        Me.GlobalQueries = New List(Of clsQuery)
    End Sub


    Public Sub SetPorts(wb As Excel.Workbook)


        'MsgBox("setPorts")



        '2 PBI Desktop Connections conflict?
        Dim connections As New clsConnections(wb)
        connections.Refresh()

        For Each cx In connections.PBIConnections
            'MsgBox(cx.Cubes.Count & " xxx")
        Next cx






        For Each c As clsConnections.clsConnection In connections.PBIConnections
            c.TestSync()
        Next c

        For Each c As clsConnections.clsConnection In connections.PBIConnections
            For Each _c As clsConnections.clsConnection In connections.PBIConnections
                If Not c Is _c AndAlso _c.ConnAlias.ToLower = c.ConnAlias.ToLower Then
                    c.IsDuplicate = True
                End If
            Next _c
        Next c

        'MsgBox("here we are")



        Dim lstwc As New List(Of Excel.WorkbookConnection)
        For Each _wb As Excel.Workbook In MyAddin.App.Workbooks
            For Each _wc As Excel.WorkbookConnection In _wb.Connections
                Try
                    If Not _wc.OLEDBConnection.Connection Is Nothing Then
                        Dim strConn As String = _wc.OLEDBConnection.Connection
                        If strConn.ToLower.StartsWith("oledb;") Then
                            strConn = strConn.Substring(6)
                        End If
                        Dim strAppName As String = ""
                        Dim bx As New OleDb.OleDbConnectionStringBuilder(strConn)
                        strAppName = bx.Item("Application Name").ToString

                        If strAppName.ToLower.StartsWith("pbixl") And bx.ConnectionString.ToLower.Replace(" ", "").Contains("datasource=localhost:") Then

                            If strAppName.ToLower.StartsWith("pbixlpivottblunnamed pbi desktop") Then

                                For Each c As clsConnections.clsConnection In connections.PBIConnections
                                    If c.ConnAlias = "" And c.IsDuplicate = False Then
                                        If bx.Item("Data Source") <> "localhost:" & c.Port Then
                                            bx.Item("Data Source") = "localhost:" & c.Port
                                            _wc.OLEDBConnection.Connection = "OLEDB;" & bx.ConnectionString
                                        End If
                                        Exit For
                                    End If
                                Next c

                            ElseIf strAppName.ToLower.StartsWith("pbixlpivottbl") Then
                                Dim strName As String = strAppName.Substring(13)
                                For Each c As clsConnections.clsConnection In connections.PBIConnections
                                    If c.ConnAlias.ToLower = strName.ToLower And c.IsDuplicate = False Then
                                        If bx.Item("Data Source") <> "localhost:" & c.Port Then
                                            bx.Item("Data Source") = "localhost:" & c.Port
                                            _wc.OLEDBConnection.Connection = "OLEDB;" & bx.ConnectionString
                                        End If
                                        Exit For
                                    End If
                                Next c

                            ElseIf strAppName.ToLower.StartsWith("pbixl") And strAppName.ToLower.EndsWith("unnamed pbi desktop") Then
                                For Each c As clsConnections.clsConnection In connections.PBIConnections
                                    If c.ConnAlias = "" And c.IsDuplicate = False Then
                                        If bx.Item("Data Source") <> "localhost:" & c.Port Then
                                            bx.Item("Data Source") = "localhost:" & c.Port
                                            _wc.OLEDBConnection.Connection = "OLEDB;" & bx.ConnectionString
                                        End If
                                        Exit For
                                    End If
                                Next c

                            ElseIf strAppName.ToLower.StartsWith("pbixl") And strAppName.ToLower.Substring(13) <> "" Then
                                For Each c As clsConnections.clsConnection In connections.PBIConnections
                                    If c.ConnAlias.ToLower = strAppName.Substring(13).ToLower And c.IsDuplicate = False Then
                                        If bx.Item("Data Source") <> "localhost:" & c.Port Then
                                            bx.Item("Data Source") = "localhost:" & c.Port
                                            _wc.OLEDBConnection.Connection = "OLEDB;" & bx.ConnectionString
                                        End If
                                        Exit For
                                    End If
                                Next c
                            End If

                        End If

                    End If
                Catch ex As Exception

                End Try
            Next _wc
        Next _wb















        'Dim c As clsConnections.clsConnection = Nothing
        'For Each _c As clsConnections.clsConnection In connections.Connections
        '    If _c.ConnType = clsConnections.clsConnection.enConnType.PBIDesktop Then
        '        Try
        '            _c.TestSync()
        '            If _c.ConnState = clsConnections.clsConnection.enConnState.OK Then
        '                c = _c
        '                Exit For
        '            End If
        '        Catch ex As Exception
        '        End Try
        '    End If
        'Next _c

        'If Not c Is Nothing Then
        '    For Each _wb As Excel.Workbook In MyAddin.App.Workbooks
        '        For Each _ws As Excel.Worksheet In _wb.Worksheets
        '            For Each _lo As Excel.ListObject In _ws.ListObjects
        '                Try
        '                    If _lo.QueryTable.WorkbookConnection.OLEDBConnection.Connection.ToString.ToLower.Replace(" ", "").Contains("datasource=localhost:") Then

        '                        Dim bx As New OleDb.OleDbConnectionStringBuilder(_lo.QueryTable.WorkbookConnection.OLEDBConnection.Connection.ToString.Substring(6))
        '                        bx.DataSource = "localhost:" & c.Port
        '                        _lo.QueryTable.WorkbookConnection.OLEDBConnection.Connection = "OLEDB;" & bx.ConnectionString

        '                    End If
        '                Catch ex As Exception
        '                End Try
        '            Next _lo
        '        Next _ws

        '        For Each wc As Excel.WorkbookConnection In _wb.Connections
        '            If wc.Description.ToLower.Contains("pbixl") AndAlso wc.OLEDBConnection.Connection.ToString.ToLower.Replace(" ", "").Contains("datasource=localhost:") Then

        '                Dim bx As New OleDb.OleDbConnectionStringBuilder(wc.OLEDBConnection.Connection.ToString.Substring(6))
        '                bx.DataSource = "localhost:" & c.Port
        '                wc.OLEDBConnection.Connection = "OLEDB;" & bx.ConnectionString

        '            End If



        '        Next wc


        '    Next _wb
        'End If


        'copy paste pbi queries
        Dim lstObj As New pbixlListObjects
        For Each _wb As Excel.Workbook In MyAddin.App.Workbooks
            For Each _ws As Excel.Worksheet In _wb.Worksheets
                For Each _lo As Excel.ListObject In _ws.ListObjects

                    Dim strConn As String = ""
                    Dim strGUID As String = ""

                    Try
                        strConn = _lo.QueryTable.Connection
                        If strConn.ToLower.StartsWith("oledb;") AndAlso strConn.Replace(" ", "").ToLower.Contains("provider=msolap") Then
                            strConn = strConn.Substring(6)
                        Else
                            strConn = ""
                        End If
                    Catch ex As Exception
                        strConn = ""
                    End Try

                    If strConn <> "" Then
                        Try
                            Dim bx As New OleDb.OleDbConnectionStringBuilder(strConn)
                            strGUID = bx.Item("Application Name").ToString
                        Catch ex As Exception
                            strGUID = ""
                        End Try
                    End If

                    If strGUID.ToLower.StartsWith("pbixl") AndAlso strGUID.Trim.Length >= 13 Then



                        Dim q As clsQuery = Me.GetQueryByGUIDRework(_wb, strGUID.Substring(5, 8))
                        lstObj.ListObjs.Add(New pbixlListObjects.clsListObj With {.conn = strConn, .lo = _lo, .query = q, .wb = _wb, .ws = _ws})


                        'lstObj.Add(New pbixlListObject With {.conn = strConn, .lo = _lo, .query = q, .wb = _wb, .ws = _ws})

                    End If




                Next _lo
            Next _ws
        Next _wb



        For Each pbiLo As pbixlListObjects.clsListObj In lstObj.ListObjs
            If pbiLo.query Is Nothing Then
                If pbiLo.ObjGUID <> "" Then
                    For Each _pbilo In lstObj.ListObjs
                        If Not _pbilo.query Is Nothing And _pbilo.ObjGUID = pbiLo.ObjGUID Then
                            Dim q As clsQuery = Me.GetQueryByGUIDRework(_pbilo.wb, _pbilo.ObjGUID.Substring(5, 8))

                            If Not q Is Nothing Then
                                q.GUID = System.Guid.NewGuid.ToString
                                Dim bx As New OleDb.OleDbConnectionStringBuilder(_pbilo.conn)
                                bx.Item("Application Name") = "pbixl" & q.GUID.ToString.Substring(0, 8)
                                bx.Item("App") = "pbixl" & q.GUID.ToString.Substring(0, 8)
                                pbiLo.lo.QueryTable.Connection = "OLEDB;" & bx.ConnectionString
                                Me.SaveQuery(pbiLo.wb, q)
                            End If

                        End If
                    Next _pbilo
                End If
            End If
        Next pbiLo

        For Each pbilo As pbixlListObjects.clsListObj In lstObj.ListObjs
            If lstObj.isD(pbilo) = True Then
                Dim q As clsQuery = Me.GetQueryByGUIDRework(pbilo.wb, pbilo.ObjGUID.Substring(5))

                If Not q Is Nothing Then
                    q.GUID = System.Guid.NewGuid.ToString
                    Dim bx As New OleDb.OleDbConnectionStringBuilder(pbilo.conn)
                    bx.Item("Application Name") = "pbixl" & q.GUID.ToString.Substring(0, 8)
                    bx.Item("App") = "pbixl" & q.GUID.ToString.Substring(0, 8)
                    pbilo.lo.QueryTable.Connection = "OLEDB;" & bx.ConnectionString
                    Me.SaveQuery(pbilo.wb, q)
                    pbilo.conn = bx.ConnectionString
                End If

            End If
        Next pbilo


        For Each pbilo As pbixlListObjects.clsListObj In lstObj.ListObjs
            If pbilo.query Is Nothing Then
                If pbilo.ObjGUID <> "" Then
                    For Each _wb As Excel.Workbook In MyAddin.App.Workbooks
                        Dim q As clsQuery = Me.GetQueryByGUIDRework(_wb, pbilo.ObjGUID.Substring(5, 8))

                        If Not q Is Nothing Then
                            q.GUID = System.Guid.NewGuid.ToString
                            Dim bx As New OleDb.OleDbConnectionStringBuilder(pbilo.conn)
                            bx.Item("Application Name") = "pbixl" & q.GUID.ToString.Substring(0, 8)
                            bx.Item("App") = "pbixl" & q.GUID.ToString.Substring(0, 8)
                            pbilo.lo.QueryTable.Connection = "OLEDB;" & bx.ConnectionString
                            Me.SaveQuery(pbilo.wb, q)
                            pbilo.conn = bx.ConnectionString
                        End If

                    Next _wb
                End If
            End If
        Next pbilo

        lstObj = Nothing





    End Sub

    Private Class pbixlListObjects

        Public Property ListObjs As List(Of clsListObj)

        Public Function isD(lo As pbixlListObjects.clsListObj) As Boolean
            For Each l As pbixlListObjects.clsListObj In ListObjs
                If lo.x <> l.x Then
                    If lo.ObjGUID = l.ObjGUID Then
                        Return True
                    End If
                End If
            Next l
            Return False
        End Function


        Public Sub New()
            Me.ListObjs = New List(Of clsListObj)
        End Sub

        Friend Class clsListObj
            Public Property x As String
            Public Property wb As Excel.Workbook
            Public Property ws As Excel.Worksheet
            Public Property lo As Excel.ListObject
            Public Property conn As String
            Public Property query As clsQuery

            Public Sub New()
                x = System.Guid.NewGuid.ToString
            End Sub

            Public ReadOnly Property ObjGUID As String
                Get
                    Dim strR As String = Me.conn
                    Try
                        If strR.ToLower.StartsWith("oledb;") Then
                            strR = strR.Substring(6)
                        End If

                        Dim bx As New OleDb.OleDbConnectionStringBuilder(strR)
                        strR = bx.Item("Application Name").ToString
                    Catch ex As Exception
                        strR = ""
                    End Try
                    Return strR
                End Get
            End Property







        End Class


    End Class



    Public Function GetQueryByGUIDRework(wb As Excel.Workbook, strGUID As String) As clsQuery

        If strGUID Is Nothing OrElse strGUID.ToString = "" Then
            Return Nothing
        End If



        Dim qRes As clsQuery = Nothing
        For Each q As clsQuery In Me.WorkbookQueries(wb)


            If strGUID.Contains(q.GUID.Substring(0, 8)) Then

                qRes = q

                For Each qc As clsQueryColumn In qRes.QueryColumns
                    If qc.BlankSel = True OrElse qc.lstSel.Count > 0 Then

                        Dim blnSel As Boolean = True
                        If qc.SelectionMode = clsQueryColumn.enSelectionMode.SelectMember Then
                            blnSel = True
                        Else
                            blnSel = False
                        End If

                        qc.htSel = New Hashtable
                        If qc.BlankSel = True Then
                            qc.htSel.Add(DBNull.Value, blnSel)
                        End If

                        For Each l In qc.lstSel
                            qc.htSel.Add(l, blnSel)
                        Next l


                    End If

                Next qc

                Exit For
            End If
        Next q

        If Not qRes Is Nothing Then
            Me.AddGlobalQuery(qRes)
        End If

        Return qRes

    End Function


    Public ReadOnly Property WorkbookQueries(wb As Excel.Workbook) As List(Of clsQuery)
        Get

            Dim lstRes As New List(Of clsQuery)

            For Each p As Microsoft.Office.Core.CustomXMLPart In wb.CustomXMLParts

                If p.NamespaceURI = "http://pbixl.com/queries" Then

                    Dim qXML As String = Me.qXML(p.XML)
                    Dim q As clsQuery = GetQueryObject(qXML)
                    If Not q Is Nothing Then
                        Me.AddGlobalQuery(q)
                        lstRes.Add(q)
                    End If
                End If
            Next p




            Return lstRes
        End Get
    End Property



    Public Sub SaveQuery(wb As Excel.Workbook, q As clsQuery)



        RemoveQuery(wb, q.GUID)
        AddGlobalQuery(q)

        For Each qc As clsQueryColumn In q.QueryColumns
            If Not qc.htSel Is Nothing Then
                qc.lstSel = New List(Of Object)
                qc.BlankSel = False
                For Each k In qc.htSel.Keys
                    If k Is DBNull.Value Then
                        qc.BlankSel = True
                    ElseIf k.ToString = "" Then
                        qc.lstSel.Add("")
                    Else
                        qc.lstSel.Add(k)
                    End If
                Next k

            End If
        Next qc

        Dim xmlString As String =
                                    "<?xml version=""1.0"" encoding=""utf-8"" ?>" &
                                    "<query xmlns=""http://pbixl.com/queries"">" &
                                    "<base64>" & Me.Base64(q.GetSerializeString) & "</base64>" &
                                    "</query>"

        Dim p As Microsoft.Office.Core.CustomXMLPart = wb.CustomXMLParts.Add(xmlString)



    End Sub


    Public Function Zip(text As String) As String
        Dim buffer As Byte() = System.Text.Encoding.Unicode.GetBytes(text)
        Using ms As New System.IO.MemoryStream
            Using zipStream As New System.IO.Compression.GZipStream(ms, System.IO.Compression.CompressionMode.Compress, True)
                zipStream.Write(buffer, 0, buffer.Length)
            End Using
            ms.Position = 0
            Dim compressed As Byte() = New Byte(ms.Length - 1) {}
            ms.Read(compressed, 0, compressed.Length)
            Dim gzBuffer As Byte() = New Byte(compressed.Length + 3) {}
            System.Buffer.BlockCopy(compressed, 0, gzBuffer, 4, compressed.Length)
            System.Buffer.BlockCopy(BitConverter.GetBytes(buffer.Length), 0, gzBuffer, 0, 4)
            Return Convert.ToBase64String(gzBuffer)
        End Using

    End Function

    Public Sub xRemoveQueryWithWhiteList(wb As Excel.Workbook, WhiteList As List(Of String))


        For i As Integer = wb.CustomXMLParts.Count To 1 Step -1

            Try
                Dim p As Microsoft.Office.Core.CustomXMLPart = wb.CustomXMLParts.Item(i)

                Dim qx As String = Me.qXML(p.XML)
                Dim qo As clsQuery = GetQueryObject(qx)

                If Not qo Is Nothing Then

                    Dim blnWhite As Boolean = False
                    For Each s As String In WhiteList
                        If s.Contains(qo.GUID.Substring(0, 8)) Then
                            blnWhite = True
                            Exit For
                        End If
                    Next s

                    If blnWhite = False Then
                        p.Delete()
                    End If

                End If
            Catch ex As Exception

            End Try

        Next i


    End Sub



    Private Sub RemoveQuery(wb As Excel.Workbook, GUID As String)
        For Each p As Microsoft.Office.Core.CustomXMLPart In wb.CustomXMLParts

            If p.NamespaceURI = "http://pbixl.com/queries" Then
                Dim qx As String = Me.qXML(p.XML)
                Dim qo As clsQuery = GetQueryObject(qx)
                If Not qo Is Nothing Then
                    If GUID = qo.GUID Then
                        p.Delete()
                        Exit Sub
                    End If
                End If
            End If
        Next p
    End Sub


    Public Sub RemoveQuery(wb As Excel.Workbook, q As clsQuery)
        For Each p As Microsoft.Office.Core.CustomXMLPart In wb.CustomXMLParts
            If p.NamespaceURI = "http://pbixl.com/queries" Then
                Dim qx As String = Me.qXML(p.XML)
                Dim qo As clsQuery = GetQueryObject(qx)
                If Not qo Is Nothing Then
                    If q.GUID = qo.GUID Then
                        p.Delete()
                        Exit Sub
                    End If
                End If
            End If
        Next p
    End Sub

    Private Function GetQueryObject(XML As String) As clsQuery
        Try
            Dim sr As New System.IO.StringReader(XML)
            Dim q As New clsQuery
            Dim x As New System.Xml.Serialization.XmlSerializer(q.GetType)
            q = CType(x.Deserialize(sr), clsQuery)
            Return q
        Catch ex As Exception
            Return Nothing
        End Try
    End Function

    Private Function qXML(XML64 As String) As String

        XML64 = XML64.Substring(XML64.IndexOf("<base64>") + 8)
        XML64 = XML64.Substring(0, XML64.LastIndexOf("</base64>"))

        Dim bt As Byte() = System.Convert.FromBase64String(XML64)
        Dim ms As New System.IO.MemoryStream
        ms.Write(bt, 0, bt.Length)
        ms.Seek(0, System.IO.SeekOrigin.Begin)
        Dim bf As New System.Runtime.Serialization.Formatters.Binary.BinaryFormatter
        Dim x As String = bf.Deserialize(ms)

        bf = Nothing
        ms.Dispose() : ms = Nothing

        Return x

    End Function

    Private Function Base64(objString As String) As String
        Dim strR As String = ""

        Dim ms As New System.IO.MemoryStream
        Dim bf As New System.Runtime.Serialization.Formatters.Binary.BinaryFormatter
        bf.Serialize(ms, objString)
        Dim bt As Byte() = ms.ToArray()
        strR = System.Convert.ToBase64String(bt)

        bf = Nothing
        ms.Dispose() : ms = Nothing
        Return strR
    End Function


End Class