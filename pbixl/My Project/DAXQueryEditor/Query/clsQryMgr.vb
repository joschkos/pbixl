Imports Microsoft.Office.Interop

Public Class clsQryMgr

    Public ReadOnly Property Queries As List(Of clsQuery)
        Get
            Return Me.WorkbookQueries
        End Get
    End Property

    Public Property wb As Excel.Workbook

    Public Sub New(wb As Excel.Workbook)
        Me.wb = wb
    End Sub



    Public Function GetQueryByGUIDRework(strGUID As String) As clsQuery


        Dim qRes As clsQuery = Nothing
        For Each q As clsQuery In Me.Queries
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
        Return qRes

    End Function

    Public ReadOnly Property WorkbookQueries As List(Of clsQuery)
        Get
            Dim lstRes As New List(Of clsQuery)

            For Each p As Microsoft.Office.Core.CustomXMLPart In Me.wb.CustomXMLParts
                If p.NamespaceURI = "http://pbixl.com/queries" Then

                    Dim qXML As String = Me.qXML(p.XML)
                    Dim q As clsQuery = GetQueryObject(qXML)
                    If Not q Is Nothing Then
                        lstRes.Add(q)
                    End If
                End If
            Next p
            Return lstRes
        End Get
    End Property



    Public Sub SaveQuery(q As clsQuery)

        RemoveQuery(q.GUID)

        For Each qc As clsQueryColumn In q.QueryColumns
            If Not qc.htSel Is Nothing Then
                qc.lstSel = New List(Of Object)
                qc.BlankSel = False
                For Each k In qc.htSel.Keys
                    If k Is DBNull.Value Then
                        qc.BlankSel = True
                    ElseIf k = "" Then
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

        Dim p As Microsoft.Office.Core.CustomXMLPart = Me.wb.CustomXMLParts.Add(xmlString)
        Me.Queries.Add(q)

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

    Public Sub RemoveQueryWithWhiteList(WhiteList As List(Of String))


        For i As Integer = Me.wb.CustomXMLParts.Count To 1 Step -1

            Try
                Dim p As Microsoft.Office.Core.CustomXMLPart = Me.wb.CustomXMLParts.Item(i)
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



    Public Sub RemoveQuery(GUID As String)
        For Each p As Microsoft.Office.Core.CustomXMLPart In Me.wb.CustomXMLParts
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


    Public Sub RemoveQuery(q As clsQuery)
        For Each p As Microsoft.Office.Core.CustomXMLPart In Me.wb.CustomXMLParts
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

    Public Function GetQueryObject(XML As String) As clsQuery
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