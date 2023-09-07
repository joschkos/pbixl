Imports System.Xml.Serialization


<Serializable>
Public Class clsQuery

    Public Enum enQuerySelectionMode
        Unknown = 0
        Total = 2
        Member = 3
    End Enum

    Public Property SelectionMode As enQuerySelectionMode

    Public Property CubeName As String


    Public ReadOnly Property FilterExpressions As List(Of String)
        Get
            Dim lstRes As New List(Of String)
            If Me.LevelCount = 0 Then
                Return lstRes
            End If



            If Me.SelectionMode = enQuerySelectionMode.Member Then
                For Each qc As clsQueryColumn In Me.QueryColumns
                    If qc.SelectedMember.Count > 0 Then
                        lstRes.Add(qc.DaxFilterTable)
                    End If
                Next qc
                Return lstRes
            End If

            If Me.SelectionMode = enQuerySelectionMode.Total Then
                If Me.InnerFilter = True Then
                    If Me.LevelFilter = True And Me.MeasureFilter = False Then
                        Dim strR As String = "var f_" & Me.GUID.Substring(0, 8) & " = SUMMARIZECOLUMNS(" & Chr(13) & Chr(10)
                        Dim strC As String = ""
                        Dim strF As String = ""
                        For Each qc As clsQueryColumn In Me.QueryColumns
                            If qc.DaxFilter <> "" Then
                                strC += qc.UniName & "," & Chr(13) & Chr(10)
                                strF += "Filter('" & qc.TableName & "'," & qc.DaxFilter & ")," & Chr(13) & Chr(10)
                            End If
                        Next qc
                        strF = strF.Substring(0, strF.Length - 3)
                        lstRes.Add(strR & strC & strF & ")")
                    ElseIf Me.MeasureFilter = True Then
                        Dim strR As String = "var f_" & Me.GUID.Substring(0, 8) & " = FILTER(SUMMARIZECOLUMNS(" & Chr(13) & Chr(10)
                        For Each qc As clsQueryColumn In Me.QueryColumns
                            If qc.FieldType = clsQueryColumn.enFieldType.Level Then
                                strR += qc.UniName & "," & Chr(13) & Chr(10)
                            End If
                        Next qc
                        If Me.LevelFilter = True Then
                            For Each qc As clsQueryColumn In Me.QueryColumns
                                If qc.DaxFilter <> "" AndAlso qc.FieldType = clsQueryColumn.enFieldType.Level Then
                                    strR += "Filter('" & qc.TableName & "'," & qc.DaxFilter & ")," & Chr(13) & Chr(10)
                                End If
                            Next qc
                        End If
                        strR = strR.Substring(0, strR.Length - 3) & ")," & Chr(13) & Chr(10)

                        For Each qc As clsQueryColumn In Me.QueryColumns
                            If qc.FieldType = clsQueryColumn.enFieldType.Measure Then
                                strR += qc.DaxFilter & " && "
                            End If
                        Next qc
                        strR = strR.Substring(0, strR.Length - 4) & ")"
                        lstRes.Add(strR)
                    End If

                End If
            End If








            Return lstRes
        End Get
    End Property


    Public Property GUID As String
    Public Property QueryColumns As List(Of clsQueryColumn)

    Public Property ConnectionName As String
    Public Property ConnectionString As String
    Public Property Command As String


    Public Property FilterControlVisible As Boolean
    Public Property FilterControlGUID As String


    Public Property AddMissingItems As Boolean

    Public Property SelectedHash As New List(Of Integer)


    Public Property CurrentExternalQueries As New List(Of clsQuery)
    Public Property PreviousExternalQueries As New List(Of clsQuery)

    Public ReadOnly Property ExternalQueriesChanged As Boolean
        Get
            Dim blnRes As Boolean = False

            If (Not Me.CurrentExternalQueries Is Nothing And Me.PreviousExternalQueries Is Nothing) Or (Me.CurrentExternalQueries Is Nothing And Not Me.PreviousExternalQueries Is Nothing) Then
                Return True
            End If

            If Me.CurrentExternalQueries Is Nothing And Me.PreviousExternalQueries Is Nothing Then
                Return False
            End If

            If Me.CurrentExternalQueries.Count <> Me.PreviousExternalQueries.Count Then
                Return True
            End If

            Me.CurrentExternalQueries.Sort(Function(x, y) x.GUID.CompareTo(y.GUID))
            Me.PreviousExternalQueries.Sort(Function(x, y) x.GUID.CompareTo(y.GUID))

            For i As Integer = 0 To Me.CurrentExternalQueries.Count - 1
                If Me.CurrentExternalQueries.Item(i).GUID <> Me.PreviousExternalQueries.Item(i).GUID Then
                    Return True
                End If
            Next i

            Dim strC As String = ""
            For i As Integer = 0 To Me.CurrentExternalQueries.Count - 1
                If Me.CurrentExternalQueries.Item(i).FilterExpressions.Count <> Me.PreviousExternalQueries.Item(i).FilterExpressions.Count Then
                    Return True
                Else
                    Me.CurrentExternalQueries.Item(i).FilterExpressions.Sort()
                    Me.PreviousExternalQueries.Item(i).FilterExpressions.Sort()
                    For j As Integer = 0 To Me.CurrentExternalQueries.Item(i).FilterExpressions.Count - 1
                        If Me.CurrentExternalQueries.Item(i).FilterExpressions(j) <> Me.PreviousExternalQueries.Item(i).FilterExpressions(j) Then
                            Return True
                        End If
                    Next j
                End If

            Next i

            Return False


        End Get
    End Property



    Public ReadOnly Property IsOutdated As Boolean
        Get
            Dim blnRes As Boolean = False

            If (Not Me.CurrentExternalQueries Is Nothing And Me.PreviousExternalQueries Is Nothing) Or (Me.CurrentExternalQueries Is Nothing And Not Me.PreviousExternalQueries Is Nothing) Then
                Return True
            End If

            If Me.CurrentExternalQueries Is Nothing And Me.PreviousExternalQueries Is Nothing Then
                Return False
            End If

            If Me.CurrentExternalQueries.Count <> Me.PreviousExternalQueries.Count Then
                Return True
            End If

            For Each q As clsQuery In Me.CurrentExternalQueries
                Dim blnFound As Boolean = False
                For Each _q As clsQuery In Me.PreviousExternalQueries
                    If q.GUID = _q.GUID Then

                        blnFound = True : Exit For
                    End If
                Next _q
                If blnFound = False Then Return True
            Next q





            Return blnRes

        End Get
    End Property





    Public Function Clone() As clsQuery

        Dim q As New clsQuery
        With q
            .GUID = Me.GUID
            .FilterControlVisible = False
            .FilterControlGUID = ""
            .ConnectionString = Me.ConnectionString
            .ConnectionName = Me.ConnectionName
            .Command = Me.Command
            .AddMissingItems = Me.AddMissingItems
            .CubeName = Me.CubeName
            .SelectionMode = Me.SelectionMode

            For Each qc In Me.QueryColumns
                Dim _qc As New clsQueryColumn()
                With _qc
                    .BlankSel = qc.BlankSel
                    .lstSel = qc.lstSel
                    If Not qc.htSel Is Nothing Then
                        .htSel = qc.htSel.Clone
                    End If
                    .DataType = qc.DataType
                    .DaxFilter = qc.DaxFilter
                    .DaxStmnt = qc.DaxStmnt
                    .FieldName = qc.FieldName
                    .FieldType = qc.FieldType
                    .FilterControlGUID = ""
                    .GUID = qc.GUID
                    .Ordinal = qc.Ordinal
                    .Query = q
                    .SearchTerm = qc.SearchTerm
                    If Not qc.SelectedMember Is Nothing AndAlso qc.SelectedMember.Count > 0 Then
                        .SelectedMember.AddRange(qc.SelectedMember)
                    End If
                    .SelectionMode = qc.SelectionMode
                    .Sort = qc.Sort
                    .TableName = qc.TableName
                    .UniName = qc.UniName


                End With

                q.QueryColumns.Add(_qc)


            Next qc
        End With

        Return q

    End Function



    Public Function GetSerializeString() As String

        Dim strR As String = ""
        Dim XmlSerializer As New XmlSerializer(Me.GetType())
        Dim txtWriter As New System.IO.StringWriter
        XmlSerializer.Serialize(txtWriter, Me)
        strR = txtWriter.ToString
        txtWriter.Dispose() : txtWriter = Nothing
        XmlSerializer = Nothing
        Return strR

    End Function


    Public ReadOnly Property QueryDefaultName As String
        Get
            Dim strRes As String = "Table"
            For Each c In Me.QueryColumns
                strRes = c.TableName
                Exit For
            Next c
            Return strRes
        End Get
    End Property

    Public ReadOnly Property DAXs0(TechCaption As Boolean) As String
        Get

            Dim strD As String = ""
            Dim strF As String = ""
            Dim lstfe As New List(Of String)

            If Me.QueryColumns.Count = 0 Then
                strD = "DEFINE var s0 = ROW(""pbixl nothing selected"",Blank())"
                Return strD
            End If


            If Me.CurrentExternalQueries.Count > 0 Then
                For Each q As clsQuery In Me.CurrentExternalQueries
                    For Each fe As String In q.FilterExpressions
                        strF += fe
                        lstfe.Add(fe.Substring(4, 10))
                    Next fe
                Next q
            End If


            If Me.LevelFilter = False And Me.MeasureFilter = False Then

                If strF <> "" Then
                    strD = "DEFINE" & Chr(13) & Chr(10) & strF & Chr(13) & Chr(10)
                Else
                    strD = "DEFINE" & Chr(13) & Chr(10)
                End If
                If Me.AddMissingItems = True Then
                    strD += "var x0 = ADDMISSINGITEMS("
                    For Each qc As clsQueryColumn In Me.AllQueryColumnsSortedByOrdinal
                        If qc.FieldType = clsQueryColumn.enFieldType.Level Then
                            strD += qc.UniName & ","
                        End If
                    Next qc
                    strD += Chr(13) & Chr(10) & "SUMMARIZECOLUMNS(" & Chr(13) & Chr(10)
                Else
                    strD += "var x0= SUMMARIZECOLUMNS(" & Chr(13) & Chr(10)
                End If
                For Each c As clsQueryColumn In Me.QueryColumns
                    If c.FieldType = clsQueryColumn.enFieldType.Level Then
                        strD = strD & c.UniName & "," & Chr(13) & Chr(10)
                    End If
                Next c
                If lstfe.Count > 0 Then
                    For Each fe In lstfe
                        strD += fe & "," & Chr(13) & Chr(10)
                    Next fe
                End If
                For Each c As clsQueryColumn In Me.QueryColumns
                    If c.FieldType = clsQueryColumn.enFieldType.Measure Then
                        strD = strD & """" & c.TechCaption & """," & c.UniName & "," & Chr(13) & Chr(10)
                    End If
                Next c
                strD = strD.Substring(0, strD.Length - 3)
                strD += ")" & Chr(13) & Chr(10)


                If Me.AddMissingItems = True Then
                    strD += ","
                    For Each qc As clsQueryColumn In Me.AllQueryColumnsSortedByOrdinal
                        If qc.FieldType = clsQueryColumn.enFieldType.Level Then
                            strD += qc.UniName & ","
                        End If
                    Next qc
                    strD = strD.Substring(0, strD.Length - 1) & ")"
                End If


                strD += "var s0 = SELECTCOLUMNS(x0,"
                For Each c As clsQueryColumn In Me.QueryColumns
                    If c.FieldType = clsQueryColumn.enFieldType.Level Then
                        strD += """" & c.TechCaption & """," & c.UniName & ","
                    Else
                        strD += """" & c.TechCaption & """,[" & c.TechCaption & "],"
                    End If
                Next c
                strD = strD.Substring(0, strD.Length - 1) & ")" & Chr(13) & Chr(10)

                'Debug.Print(strD)
                Return strD


            End If


            If Me.LevelFilter = True And Me.MeasureFilter = False Then

                If strF <> "" Then
                    strD = "DEFINE" & Chr(13) & Chr(10) & strF & Chr(13) & Chr(10)
                Else
                    strD = "DEFINE" & Chr(13) & Chr(10)
                End If

                Dim ctrF As Integer = 0
                For Each c As clsQueryColumn In Me.QueryColumns
                    If c.FieldType = clsQueryColumn.enFieldType.Level AndAlso c.DaxFilter <> "" Then
                        ctrF += 1
                        strD += "var f" & ctrF.ToString & "=Filter('" & c.TableName & "'," & c.DaxFilter & ")" & Chr(13) & Chr(10)
                    End If
                Next c

                If Me.AddMissingItems = True Then
                    strD += "var x0 = ADDMISSINGITEMS("
                    For Each qc As clsQueryColumn In Me.AllQueryColumnsSortedByOrdinal
                        If qc.FieldType = clsQueryColumn.enFieldType.Level Then
                            strD += qc.UniName & ","
                        End If
                    Next qc
                    strD += Chr(13) & Chr(10) & "SUMMARIZECOLUMNS(" & Chr(13) & Chr(10)
                Else
                    strD += "var x0= SUMMARIZECOLUMNS(" & Chr(13) & Chr(10)
                End If

                For Each c As clsQueryColumn In Me.QueryColumns
                    If c.FieldType = clsQueryColumn.enFieldType.Level Then
                        strD = strD & c.UniName & "," & Chr(13) & Chr(10)
                    End If
                Next c
                If lstfe.Count > 0 Then
                    For Each fe In lstfe
                        strD += fe & "," & Chr(13) & Chr(10)
                    Next fe
                End If
                For i = 1 To ctrF
                    strD += "f" & i.ToString & "," & Chr(13) & Chr(10)
                Next i
                For Each c As clsQueryColumn In Me.QueryColumns
                    If c.FieldType = clsQueryColumn.enFieldType.Measure Then
                        strD = strD & """" & c.TechCaption & """," & c.UniName & "," & Chr(13) & Chr(10)
                    End If
                Next c
                strD = strD.Substring(0, strD.Length - 3)
                strD += ")" & Chr(13) & Chr(10)

                If Me.AddMissingItems = True Then
                    strD += ","
                    For Each qc As clsQueryColumn In Me.AllQueryColumnsSortedByOrdinal
                        If qc.FieldType = clsQueryColumn.enFieldType.Level Then
                            strD += qc.UniName & "," & Chr(13) & Chr(10)
                        End If
                    Next qc
                    If lstfe.Count > 0 Then
                        For Each fe In lstfe
                            strD += fe & "," & Chr(13) & Chr(10)
                        Next fe
                    End If
                    For i = 1 To ctrF
                        strD += "f" & i.ToString & "," & Chr(13) & Chr(10)
                    Next i
                    strD = strD.Substring(0, strD.Length - 3) & ")"
                End If

                strD += Chr(13) & Chr(10) & "var s0 = SELECTCOLUMNS(x0,"
                For Each c As clsQueryColumn In Me.QueryColumns
                    If c.FieldType = clsQueryColumn.enFieldType.Level Then
                        strD += """" & c.TechCaption & """," & c.UniName & ","
                    Else
                        strD += """" & c.TechCaption & """,[" & c.TechCaption & "],"
                    End If
                Next c
                strD = strD.Substring(0, strD.Length - 1) & ")" & Chr(13) & Chr(10)

                Return strD

            End If


            If Me.LevelFilter = True And Me.MeasureFilter = True Then

                If strF <> "" Then
                    strD = "DEFINE" & Chr(13) & Chr(10) & strF & Chr(13) & Chr(10)
                Else
                    strD = "DEFINE" & Chr(13) & Chr(10)
                End If

                Dim ctrF As Integer = 0
                For Each c As clsQueryColumn In Me.QueryColumns
                    If c.FieldType = clsQueryColumn.enFieldType.Level AndAlso c.DaxFilter <> "" Then
                        ctrF += 1
                        strD += "var f" & ctrF.ToString & "=Filter('" & c.TableName & "'," & c.DaxFilter & ")" & Chr(13) & Chr(10)
                    End If
                Next c

                If Me.AddMissingItems = True Then
                    strD += "var x0 = ADDMISSINGITEMS("
                    For Each qc As clsQueryColumn In Me.AllQueryColumnsSortedByOrdinal
                        If qc.FieldType = clsQueryColumn.enFieldType.Level Then
                            strD += qc.UniName & ","
                        End If
                    Next qc
                    strD += Chr(13) & Chr(10) & "SUMMARIZECOLUMNS(" & Chr(13) & Chr(10)
                Else
                    strD += "var x0= SUMMARIZECOLUMNS(" & Chr(13) & Chr(10)
                End If

                For Each c As clsQueryColumn In Me.QueryColumns
                    If c.FieldType = clsQueryColumn.enFieldType.Level Then
                        strD = strD & c.UniName & "," & Chr(13) & Chr(10)
                    End If
                Next c
                If lstfe.Count > 0 Then
                    For Each fe In lstfe
                        strD += fe & "," & Chr(13) & Chr(10)
                    Next fe
                End If
                For i = 1 To ctrF
                    strD += "f" & i.ToString & "," & Chr(13) & Chr(10)
                Next i
                For Each c As clsQueryColumn In Me.QueryColumns
                    If c.FieldType = clsQueryColumn.enFieldType.Measure Then
                        strD = strD & """" & c.TechCaption & """," & c.UniName & "," & Chr(13) & Chr(10)
                    End If
                Next c
                strD = strD.Substring(0, strD.Length - 3)
                strD += ")" & Chr(13) & Chr(10)

                If Me.AddMissingItems = True Then
                    strD += ","
                    For Each qc As clsQueryColumn In Me.AllQueryColumnsSortedByOrdinal
                        If qc.FieldType = clsQueryColumn.enFieldType.Level Then
                            strD += qc.UniName & "," & Chr(13) & Chr(10)
                        End If
                    Next qc
                    If lstfe.Count > 0 Then
                        For Each fe In lstfe
                            strD += fe & "," & Chr(13) & Chr(10)
                        Next fe
                    End If
                    For i = 1 To ctrF
                        strD += "f" & i.ToString & "," & Chr(13) & Chr(10)
                    Next i
                    strD = strD.Substring(0, strD.Length - 3) & ")"
                End If

                strD += Chr(13) & Chr(10) & "var m0 = Filter(x0,"
                For Each c As clsQueryColumn In Me.QueryColumns
                    If c.FieldType = clsQueryColumn.enFieldType.Measure Then
                        If c.DaxFilter <> "" Then
                            strD += c.DaxFilter & " && "
                        End If
                    End If
                Next c
                strD = strD.Substring(0, strD.Length - 4) & ")" & Chr(13) & Chr(10)

                strD += "var s0 = SELECTCOLUMNS(m0,"
                For Each c As clsQueryColumn In Me.QueryColumns
                    If c.FieldType = clsQueryColumn.enFieldType.Level Then
                        strD += """" & c.TechCaption & """," & c.UniName & ","
                    Else
                        strD += """" & c.TechCaption & """,[" & c.TechCaption & "],"
                    End If
                Next c
                strD = strD.Substring(0, strD.Length - 1) & ")" & Chr(13) & Chr(10)

                Return strD

            End If


            If Me.LevelFilter = False And Me.MeasureFilter = True Then


                If strF <> "" Then
                    strD = "DEFINE" & Chr(13) & Chr(10) & strF & Chr(13) & Chr(10)
                Else
                    strD = "DEFINE" & Chr(13) & Chr(10)
                End If

                If Me.AddMissingItems = True Then
                    strD += "var x0 = ADDMISSINGITEMS("
                    For Each qc As clsQueryColumn In Me.AllQueryColumnsSortedByOrdinal
                        If qc.FieldType = clsQueryColumn.enFieldType.Level Then
                            strD += qc.UniName & ","
                        End If
                    Next qc
                    strD += Chr(13) & Chr(10) & "SUMMARIZECOLUMNS(" & Chr(13) & Chr(10)
                Else
                    strD += "var x0= SUMMARIZECOLUMNS(" & Chr(13) & Chr(10)
                End If

                For Each c As clsQueryColumn In Me.QueryColumns
                    If c.FieldType = clsQueryColumn.enFieldType.Level Then
                        strD = strD & c.UniName & "," & Chr(13) & Chr(10)
                    End If
                Next c
                If lstfe.Count > 0 Then
                    For Each fe In lstfe
                        strD += fe & "," & Chr(13) & Chr(10)
                    Next fe
                End If
                For Each c As clsQueryColumn In Me.QueryColumns
                    If c.FieldType = clsQueryColumn.enFieldType.Measure Then
                        strD = strD & """" & c.TechCaption & """," & c.UniName & "," & Chr(13) & Chr(10)
                    End If
                Next c
                strD = strD.Substring(0, strD.Length - 3)
                strD += ")" & Chr(13) & Chr(10)

                If Me.AddMissingItems = True Then
                    strD += ","
                    For Each qc As clsQueryColumn In Me.AllQueryColumnsSortedByOrdinal
                        If qc.FieldType = clsQueryColumn.enFieldType.Level Then
                            strD += qc.UniName & ","
                        End If
                    Next qc
                    If lstfe.Count > 0 Then
                        For Each fe In lstfe
                            strD += fe & "," & Chr(13) & Chr(10)
                        Next fe
                    End If
                    strD = strD.Substring(0, strD.Length - 1) & ")"
                End If

                strD += Chr(13) & Chr(10) & "var m0 = Filter(x0,"
                For Each c As clsQueryColumn In Me.QueryColumns
                    If c.FieldType = clsQueryColumn.enFieldType.Measure Then
                        If c.DaxFilter <> "" Then
                            strD += c.DaxFilter & " && "
                        End If
                    End If
                Next c
                strD = strD.Substring(0, strD.Length - 4) & ")" & Chr(13) & Chr(10)

                strD += "var s0 = SELECTCOLUMNS(m0,"
                For Each c As clsQueryColumn In Me.QueryColumns
                    If c.FieldType = clsQueryColumn.enFieldType.Level Then
                        strD += """" & c.TechCaption & """," & c.UniName & ","
                    Else
                        strD += """" & c.TechCaption & """,[" & c.TechCaption & "],"
                    End If
                Next c
                strD = strD.Substring(0, strD.Length - 1) & ")" & Chr(13) & Chr(10)

                Return strD

            End If

            Return ""


        End Get
    End Property

    Public ReadOnly Property DAXBaseTable(Preview As Boolean, nRows As Integer, TechCaption As Boolean) As String
        Get


            Try

                Dim lstfi As New List(Of String)
                Dim ctrfi As Integer = 0
                For Each qc As clsQueryColumn In Me.QueryColumns
                    If qc.FieldType = clsQueryColumn.enFieldType.Level Then
                        If qc.IsFiltered = True Then
                            ctrfi += 1
                            lstfi.Add("var f" & ctrfi.ToString & " = FILTER('" & qc.TableName & "'," & qc.DaxFilter & ")")
                        End If
                    End If
                Next qc


                Dim lstfe As New List(Of String)
                For Each q As clsQuery In Me.CurrentExternalQueries
                    For Each s As String In q.DAXFilterExport
                        lstfe.Add(s)
                    Next s
                Next q

                If Me.SelectionMode = enQuerySelectionMode.Member Then
                    If Me.DAXSelectedMember <> "" Then
                        For Each s As String In Me.DAXSelectedMemberLst
                            'lstfe.Add(s)
                        Next s
                    End If
                End If


                For i = 0 To lstfe.Count - 1
                    lstfe.Item(i) = "var e" & i.ToString & " = " & lstfe.Item(i).Substring(9)
                Next i

                Dim strD As String = "DEFINE " & vbCrLf
                For Each s As String In lstfe
                    strD += s & vbCrLf
                Next s
                For Each s As String In lstfi
                    strD += s & vbCrLf
                Next

                strD += "var x0 = " & vbCrLf

                If Me.AddMissingItems = True Then
                    strD += "ADDMISSINGITEMS("
                    For Each qc As clsQueryColumn In Me.QueryColumns
                        If qc.FieldType = clsQueryColumn.enFieldType.Level Then
                            strD += qc.UniName & ","
                        End If
                    Next qc
                End If
                strD += "SUMMARIZECOLUMNS("
                For Each qc As clsQueryColumn In Me.QueryColumns
                    If qc.FieldType = clsQueryColumn.enFieldType.Level Then
                        strD += qc.UniName & "," & vbCrLf
                    End If
                Next qc
                For Each s In lstfi
                    strD += s.Substring(0, 8).Trim.Replace("var", "").Replace("=", "").Trim & ","
                Next s
                For Each s In lstfe
                    strD += s.Substring(0, 8).Trim.Replace("var", "").Replace("=", "").Trim & ","
                Next s
                For Each qc As clsQueryColumn In Me.QueryColumns
                    If qc.FieldType = clsQueryColumn.enFieldType.Measure Then
                        strD += """" & qc.UniName & """," & qc.UniName & "," & vbCrLf
                    End If
                Next qc
                If strD.EndsWith(",") Then strD = strD.Substring(0, strD.Length - 1)
                If strD.EndsWith(vbCrLf) Then strD = strD.Substring(0, strD.Length - 2)
                If strD.EndsWith(",") Then strD = strD.Substring(0, strD.Length - 1)
                If strD.EndsWith(vbCrLf) Then strD = strD.Substring(0, strD.Length - 2)
                strD += ")" & vbCrLf
                If Me.AddMissingItems = True Then
                    strD += ","
                    For Each qc As clsQueryColumn In Me.QueryColumns
                        If qc.FieldType = clsQueryColumn.enFieldType.Level Then
                            strD += qc.UniName & ","
                        End If
                    Next qc
                    For Each s In lstfi
                        strD += s.Substring(0, 8).Trim.Replace("var", "").Replace("=", "").Trim & ","
                    Next s
                    strD = strD.Substring(0, strD.Length - 1) & ")"
                End If



                If Me.MeasureFilter = True Then
                    strD += "var f_m0 = Filter(x0,"
                    For Each qc In Me.QueryColumns
                        If qc.FieldType = clsQueryColumn.enFieldType.Measure Then
                            If qc.IsFiltered = True Then
                                'strD += "('" & qc.TableName & "'" & qc.DaxFilter.Trim & ") && "
                                strD += "(" & qc.DaxFilter.Trim & ") && "
                            End If
                        End If
                    Next qc
                    strD = strD.Substring(0, strD.Length - 4) & ")" & vbCrLf
                Else
                    strD += "var f_m0 = Filter(x0,true)" & vbCrLf
                End If

                strD += "var s0 = SelectColumns(f_m0," & vbCrLf
                For Each qc As clsQueryColumn In Me.AllQueryColumnsSortedByOrdinal
                    If TechCaption = False Then
                        strD += """" & qc.ReferenceName & """," & qc.UniName & "," & vbCrLf
                    Else
                        strD += """" & qc.TechCaption & """," & qc.UniName & "," & vbCrLf
                    End If
                Next qc

                strD = strD.Substring(0, strD.Length - 3) & ")" & vbCrLf


                If Me.QueryColumns.Count = 1 Then
                    If Me.QueryColumns.Item(0).Sort = clsQueryColumn.enSort.none Then
                        Me.QueryColumns.Item(0).Sort = clsQueryColumn.enSort.asc
                    End If
                    If TechCaption = False Then
                        strD += "var x1 = TOPN(1000,s0,[" & Me.QueryColumns.Item(0).ReferenceName & "]," & Me.QueryColumns.Item(0).Sort.ToString & ")"
                    Else
                        strD += "var x1 = TOPN(1000,s0,[" & Me.QueryColumns.Item(0).TechCaption & "]," & Me.QueryColumns.Item(0).Sort.ToString & ")"
                    End If
                Else
                    If TechCaption = False Then
                        strD += "var x1 = TOPN(1000,s0,[" & Me.OrderColumn.ReferenceName & "]," & Me.OrderColumn.Sort.ToString & ","
                    Else
                        strD += "var x1 = TOPN(1000,s0,[" & Me.OrderColumn.TechCaption & "]," & Me.OrderColumn.Sort.ToString & ","
                    End If
                    For Each qc In Me.QueryColumns
                        If Not qc Is Me.OrderColumn Then
                            If TechCaption = False Then
                                strD += "[" & qc.ReferenceName & "],false,"
                            Else
                                strD += "[" & qc.TechCaption & "],false,"
                            End If
                        End If
                    Next qc
                    strD = strD.Substring(0, strD.Length - 1) & ")"
                End If

                If Me.QueryColumns.Count = 1 Then
                    If TechCaption = False Then
                        strD += vbCrLf & "EVALUATE x1 ORDER BY [" & Me.OrderColumn.ReferenceName & "] " & Me.OrderColumn.Sort.ToString
                    Else
                        strD += vbCrLf & "EVALUATE x1 ORDER BY [" & Me.OrderColumn.TechCaption & "] " & Me.OrderColumn.Sort.ToString
                    End If
                Else
                    If TechCaption = False Then
                        strD += vbCrLf & "EVALUATE x1 ORDER BY [" & Me.OrderColumn.ReferenceName & "] " & Me.OrderColumn.Sort.ToString & ","
                    Else
                        strD += vbCrLf & "EVALUATE x1 ORDER BY [" & Me.OrderColumn.TechCaption & "] " & Me.OrderColumn.Sort.ToString & ","
                    End If
                    For Each qc As clsQueryColumn In Me.AllQueryColumnsSortedByOrdinal
                        If Not qc Is Me.OrderColumn Then
                            If TechCaption = False Then
                                strD += "[" & qc.ReferenceName & "] asc,"
                            Else
                                strD += "[" & qc.TechCaption & "] asc,"
                            End If


                        End If
                    Next qc
                    strD = strD.Substring(0, strD.Length - 1) & ""
                End If


                Me.cDAXBaseTable = strD
                'Debug.Print(strD)
                Return strD

            Catch ex As Exception
                Return ""
            End Try

        End Get
    End Property







    Public ReadOnly Property DAXPreview(TechCaption As Boolean, nRows As Integer)
        Get
            Dim strD As String = Me.DAXs0(TechCaption) & Chr(13) & Chr(10)

            If Not Me.OrderColumn Is Nothing Then

                strD += "var x1 = TOPN(" & nRows.ToString & ",s0," & Chr(13) & Chr(10)
                strD += "[" & Me.OrderColumn.TechCaption & "]," & Me.OrderColumn.Sort.ToString & Chr(13) & Chr(10) & ","
                For Each qc In Me.QueryColumns
                    If Not qc Is Me.OrderColumn Then
                        strD += "[" & qc.TechCaption & "],False" & Chr(13) & Chr(10) & ","
                    End If
                Next qc
                strD = strD.Substring(0, strD.Length - 1) & ")" & Chr(13) & Chr(10)
                strD += "EVALUATE x1 ORDER BY [" & Me.OrderColumn.TechCaption & "] " & Me.OrderColumn.Sort.ToString

            Else

                strD += "var x1 = TOPN(" & nRows.ToString & ",s0,"
                For Each qc In Me.QueryColumns
                    strD += "[" & qc.TechCaption & "],False" & Chr(13) & Chr(10) & ","
                Next qc
                strD = strD.Substring(0, strD.Length - 1) & ")" & Chr(13) & Chr(10)
                strD += "EVALUATE x1"

            End If

            Return strD
        End Get
    End Property

    Public ReadOnly Property DaxSelectionTotal(TechCaption As Boolean) As String
        Get

            If Me.QueryColumns.Count = 0 Then
                Return "DEFINE var s0 = ROW(""pbixl Nothing selected"",Blank())"
            End If

            Dim ctr As Integer = 0
            Dim strD As String = "DEFINE" & Chr(13) & Chr(10)
            For Each qc As clsQueryColumn In Me.QueryColumns
                If qc.FieldType = clsQueryColumn.enFieldType.Level Then

                    Dim df As String = qc.DaxFilterTable
                    If df <> "" Then ctr += 1
                    strD += df


                End If
            Next qc

            For Each eq As clsQuery In Me.CurrentExternalQueries
                For Each qc As clsQueryColumn In eq.QueryColumns

                    Dim df As String = qc.DaxFilterTable
                    If df <> "" Then ctr += 1
                    strD += df

                Next qc
            Next eq




            strD += "EVALUATE" & Chr(13) & Chr(10)
            If Me.QueryColumns.Count = 1 Then
                strD += "ROW(""TechCaption"",""" & Me.QueryColumns.Item(0).TechCaption & """,""value"",countrows(f_" & Me.QueryColumns.Item(0).TechCaption & "))" & Chr(13) & Chr(10)
            Else
                strD += "UNION(" & Chr(13) & Chr(10)
                For Each qc As clsQueryColumn In Me.QueryColumns
                    If qc.FieldType = clsQueryColumn.enFieldType.Level Then
                        strD += "ROW(""TechCaption"",""" & qc.TechCaption & """,""value"",countrows(f_" & qc.TechCaption & "))," & Chr(13) & Chr(10)
                    End If
                Next qc
                For Each qc As clsQueryColumn In Me.QueryColumns
                    If qc.FieldType = clsQueryColumn.enFieldType.Measure Then
                        strD += "ROW(""TechCaption"",""" & qc.TechCaption & """,""value"",calculate([" & qc.FieldName & "],"
                        For Each _qc As clsQueryColumn In Me.QueryColumns
                            If _qc.FieldType = clsQueryColumn.enFieldType.Level Then
                                strD += "f_" & _qc.TechCaption & ","
                            End If
                        Next _qc
                        For Each eq As clsQuery In Me.CurrentExternalQueries
                            For Each _qc As clsQueryColumn In eq.QueryColumns
                                If _qc.FieldType = clsQueryColumn.enFieldType.Level Then
                                    strD += "f_" & _qc.TechCaption & ","
                                End If
                            Next _qc
                        Next eq



                        strD = strD.Substring(0, strD.Length - 1) & "))" & "," & Chr(13) & Chr(10)
                    End If
                Next qc
                If strD.EndsWith("," & Chr(13) & Chr(10)) Then
                    strD = strD.Substring(0, strD.Length - 3)
                ElseIf strD.EndsWith(Chr(13) & Chr(10)) Then
                    strD = strD.Substring(0, strD.Length - 2)
                End If

                strD += Chr(13) & Chr(10) & ")"
            End If



            If ctr = 0 Then
                Return ""
            Else
                Return strD
            End If


        End Get
    End Property

    Private _DaxSelectedMemberLst As List(Of String)
    Public ReadOnly Property DAXSelectedMemberLst As List(Of String)
        Get
            Dim strX As String = Me.DAXSelectedMember
            Return Me._DaxSelectedMemberLst
        End Get
    End Property


    Public ReadOnly Property DAXSelectedMember As String
        Get
            Me._DaxSelectedMemberLst = New List(Of String)



            Dim ctrF As Integer = 0

            Dim strF As String = ""
            Dim strL As String = ""
            For Each qc In Me.QueryColumns
                strL = ""
                If qc.SelectedMember.Count > 0 Then
                    ctrF += 1
                    strF += "var f_" & ctrF.ToString & " = Filter(Values(" & qc.UniName & ")," & qc.UniName & " In {"
                    strL = "var fx = Filter(Values(" & qc.UniName & ")," & qc.UniName & " In {"
                    If qc.DataType = clsQueryColumn.enDataType.Text Then
                        For Each m In qc.SelectedMember
                            If IsDBNull(m) Then
                                strF += "blank(),"
                                strL += "blank(),"
                            Else
                                strF += """" & m & ""","
                                strL += """" & m & ""","
                            End If
                        Next m
                        strF = strF.Substring(0, strF.Length - 1) & "})" & Chr(13) & Chr(10)
                        strL = strL.Substring(0, strL.Length - 1) & "})" & Chr(13) & Chr(10)
                        Me._DaxSelectedMemberLst.Add(strL)
                    ElseIf qc.DataType = clsQueryColumn.enDataType.DateTime Then
                        For Each m In qc.SelectedMember
                            If IsDBNull(m) Then
                                strF += "blank(),"
                                strL += "blank(),"
                            Else
                                Dim strR As String = "(Date(" & DateAndTime.Year(m) & "," & DateAndTime.Month(m) & "," & DateAndTime.Day(m) & ") + " _
                                & "TIME(" & DateAndTime.Hour(m) & "," & DateAndTime.Minute(m) & "," & DateAndTime.Second(m) & "))"
                                strF += strR & ","
                                strL += strR & ","
                            End If

                        Next m
                        strF = strF.Substring(0, strF.Length - 1) & "})" & Chr(13) & Chr(10)
                        strL = strL.Substring(0, strL.Length - 1) & "})" & Chr(13) & Chr(10)
                        Me._DaxSelectedMemberLst.Add(strL)
                    ElseIf qc.DataType = clsQueryColumn.enDataType.Bool Then
                        For Each m In qc.SelectedMember
                            If IsDBNull(m) Then
                                strF += "blank(),"
                                strL += "blank(),"
                            Else
                                If m.ToString.ToLower = "true" Then
                                    strF += "True,"
                                    strL += "True,"
                                Else
                                    strF += "False,"
                                    strL += "False,"
                                End If
                            End If
                        Next m
                        strF = strF.Substring(0, strF.Length - 1) & "})" & Chr(13) & Chr(10)
                        strL = strL.Substring(0, strL.Length - 1) & "})" & Chr(13) & Chr(10)
                        Me._DaxSelectedMemberLst.Add(strL)
                    Else
                        For Each m In qc.SelectedMember
                            If IsDBNull(m) Then
                                strF += "blank(),"
                                strL += "blank(),"
                            Else
                                strF += m & ","
                                strL += m & ","
                            End If
                        Next m
                        strF = strF.Substring(0, strF.Length - 1) & "})" & Chr(13) & Chr(10)
                        strL = strL.Substring(0, strL.Length - 1) & "})" & Chr(13) & Chr(10)
                        Me._DaxSelectedMemberLst.Add(strL)
                    End If

                End If
            Next qc


            Return strF



        End Get
    End Property


    Public ReadOnly Property DAXFilterExport As List(Of String)
        Get
            Dim lst As New List(Of String)
            Dim strX As String = ""

            If Me.SelectionMode = enQuerySelectionMode.Member Then

                For Each s As String In Me.DAXSelectedMemberLst
                    lst.Add(s)
                    strX += s & vbCrLf
                Next s
                Me.cDAXFilterExport = strX
                Return lst
            End If


            If Me.LevelFilter = False And Me.MeasureFilter = False Then
                Me.cDAXFilterExport = ""
                Return lst
            End If

            Dim ctrf As Integer = 0
            If Me.LevelFilter = True And Me.MeasureFilter = False Then
                For Each qc As clsQueryColumn In Me.QueryColumns
                    If qc.IsFiltered = True Then
                        ctrf += 1
                        lst.Add("var f" & ctrf.ToString & " = FILTER(" & qc.TableName & "," & qc.DaxFilter & ")")
                        strX += lst.Item(lst.Count - 1) & vbCrLf
                    End If
                Next qc
                Me.cDAXFilterExport = strX
                Return lst
            End If

            If Me.MeasureFilter = True Then
                ctrf += 1
                Dim strD As String = "var f" & ctrf.ToString & " = FILTER(" & Chr(13) & Chr(10)

                strD += "summarizecolumns("
                For Each qc As clsQueryColumn In Me.QueryColumns
                    If qc.FieldType = clsQueryColumn.enFieldType.Level Then
                        strD += qc.UniName & "," & Chr(13) & Chr(10)
                    End If
                Next qc
                For Each qc As clsQueryColumn In Me.QueryColumns
                    If qc.FieldType = clsQueryColumn.enFieldType.Level Then
                        If qc.IsFiltered = True Then
                            strD += "Filter(Values(" & qc.UniName & ")," & qc.DaxFilter & ")," & Chr(13) & Chr(10)
                        End If
                    End If
                Next qc
                For Each qc As clsQueryColumn In Me.QueryColumns
                    If qc.FieldType = clsQueryColumn.enFieldType.Measure Then
                        strD += """" & qc.TechCaption & """," & qc.UniName & "," & Chr(13) & Chr(10)
                    End If
                Next qc
                strD = strD.Substring(0, strD.Length - 3) & ")"
                strD += ","
                For Each qc In Me.QueryColumns
                    If qc.FieldType = clsQueryColumn.enFieldType.Measure Then
                        If qc.IsFiltered = True Then
                            strD += qc.DaxFilter.Replace("[" & qc.ReferenceName & "]", "[" & qc.TechCaption & "]") & ","
                        End If
                    End If
                Next qc
                strD = strD.Substring(0, strD.Length - 1) & ")"
                lst.Add(strD)
                Me.cDAXFilterExport = strD
                Return lst
            End If

            Me.cDAXFilterExport = ""
            Return lst

        End Get
    End Property



    Public Property cDAXTableTotal As String
    Public Property cDAXFilterExport As String
    Public Property cDAXBaseTable As String


    Public ReadOnly Property DAXTableTotal(TechCaption As Boolean) As String
        Get

            Dim lstfi As New List(Of String)
            Dim ctrfi As Integer = 0
            For Each qc As clsQueryColumn In Me.QueryColumns
                If qc.FieldType = clsQueryColumn.enFieldType.Level Then
                    If qc.IsFiltered = True Then
                        ctrfi += 1
                        lstfi.Add("var f" & ctrfi.ToString & " = FILTER(" & qc.TableName & "," & qc.DaxFilter & ")")
                    End If
                End If
            Next qc


            Dim lstfe As New List(Of String)
            For Each q As clsQuery In Me.CurrentExternalQueries
                For Each s As String In q.DAXFilterExport
                    lstfe.Add(s)
                Next s
            Next q

            If Me.SelectionMode = enQuerySelectionMode.Member Then
                If Me.DAXSelectedMember <> "" Then
                    For Each s As String In Me.DAXSelectedMemberLst
                        lstfe.Add(s)
                    Next s
                End If
            End If

            For i = 0 To lstfe.Count - 1
                lstfe.Item(i) = "var e" & i.ToString & " = " & lstfe.Item(i).Substring(9)
            Next i

            Dim strD As String = "DEFINE" & vbCrLf
            For Each s As String In lstfe
                strD += s & vbCrLf
            Next s
            For Each s As String In lstfi
                strD += s & vbCrLf
            Next s



            strD += "var x0 = " & vbCrLf
            Dim strS As String = "SUMMARIZECOLUMNS(" & vbCrLf
            For Each qc In Me.QueryColumns
                If qc.FieldType = clsQueryColumn.enFieldType.Level Then
                    strS += qc.UniName & "," & vbCrLf
                End If
            Next qc
            For Each s As String In lstfe
                strS += s.Substring(0, s.IndexOf("=")).Trim.Replace("var", "").Replace("=", "").Trim & ","
            Next s
            For Each s As String In lstfi
                strS += s.Substring(0, s.IndexOf("=")).Trim.Replace("var", "").Replace("=", "").Trim & ","
            Next s
            If lstfe.Count > 0 Or lstfi.Count > 0 Then
                strS += vbCrLf
            End If
            For Each qc In Me.QueryColumns
                If qc.FieldType = clsQueryColumn.enFieldType.Measure Then
                    strS += """" & qc.TechCaption & """," & qc.UniName & "," & vbCrLf
                End If
            Next qc




            If Me.MeasureFilter = False Then
                strS = strS.Substring(0, strS.Length - 3) & ")"
            Else
                strS = "Filter(" & strS
                For Each qc In Me.QueryColumns
                    If qc.FieldType = clsQueryColumn.enFieldType.Measure Then
                        If qc.IsFiltered = True Then
                            If strS.Contains(qc.TechCaption) = False Then
                                strS += """" & qc.TechCaption & """," & qc.UniName & "," & vbCrLf
                            End If
                        End If
                    End If
                Next qc
                If strS.EndsWith("," & vbCrLf) Then
                    strS = strS.Substring(0, strS.Length - 3) & vbCrLf
                End If

                strS += "),"
                For Each qc In Me.QueryColumns
                    If qc.FieldType = clsQueryColumn.enFieldType.Measure Then
                        If qc.IsFiltered = True Then
                            strS += qc.DaxFilter & " && "
                        End If
                    End If
                Next qc
                strS = strS.Substring(0, strS.Length - 4) & ")"
            End If
            strD = strD & strS

            Dim lstR As New List(Of String)
            For Each qc In Me.QueryColumns
                If qc.FieldType = clsQueryColumn.enFieldType.Level Then
                    Dim strR As String = ""
                    strR = "row(""TechCaption"",""" & qc.TechCaption & """,""value"",countrows(summarize(x0," & qc.UniName & ")))" & Chr(13) & Chr(10)
                    lstR.Add(strR)
                End If
            Next qc
            For Each qc In Me.QueryColumns
                If qc.FieldType = clsQueryColumn.enFieldType.Measure Then
                    Dim strR As String = ""
                    If Me.MeasureFilter = True Then
                        strR = "row(""TechCaption"",""" & qc.TechCaption & """,""value"",calculate(" & qc.UniName & ",x0,"
                    Else
                        strR = "row(""TechCaption"",""" & qc.TechCaption & """,""value"",calculate(" & qc.UniName & ","
                    End If
                    For Each s As String In lstfe
                        strR += s.Substring(0, s.IndexOf("=")).Trim.Replace("var", "").Replace("=", "").Trim & ","
                    Next s
                    For Each s As String In lstfi
                        strR += s.Substring(0, s.IndexOf("=")).Trim.Replace("var", "").Replace("=", "").Trim & ","
                    Next s
                    strR = strR.Substring(0, strR.Length - 1) & "))"
                    lstR.Add(strR)
                End If
            Next qc

            strD += vbCrLf
            If lstR.Count = 1 Then
                strD += "EVALUATE " & lstR.Item(0)
            Else
                strD += "EVALUATE UNION(" & vbCrLf
                For Each s As String In lstR
                    strD += s & ","
                Next s
                strD = strD.Substring(0, strD.Length - 1) & vbCrLf & ")"
            End If

            Me.cDAXTableTotal = strD
            Return strD


        End Get
    End Property

    Public ReadOnly Property DAXTableInnerTotal(TechCaption As Boolean) As String
        Get

            Dim strD As String = Me.DAXs0(TechCaption) & Chr(13) & Chr(10)

            If Me.QueryColumns.Count > 1 Then
                strD += "EVALUATE UNION(" & Chr(13) & Chr(10)
                For Each qc In Me.QueryColumns
                    If qc.FieldType = clsQueryColumn.enFieldType.Level Then
                        strD += "row(""TechCaption"",""" & qc.TechCaption & """,""value"",countrows(summarize(s0,[" & qc.TechCaption & "])))," & Chr(13) & Chr(10)
                    End If
                    If qc.FieldType = clsQueryColumn.enFieldType.Measure Then
                        strD += "row(""TechCaption"",""" & qc.TechCaption & """,""value"",calculate(" & qc.UniName & ",s0))," & Chr(13) & Chr(10)
                    End If
                Next qc
                strD = strD.Substring(0, strD.Length - 3) & ")"
            Else
                strD += "EVALUATE " & Chr(13) & Chr(10)
                Dim qc As clsQueryColumn = Me.QueryColumns.Item(0)
                If qc.FieldType = clsQueryColumn.enFieldType.Level Then
                    strD += "row(""TechCaption"",""" & qc.TechCaption & """,""value"",countrows(summarize(s0,[" & qc.TechCaption & "])))," & Chr(13) & Chr(10)
                End If
                If qc.FieldType = clsQueryColumn.enFieldType.Measure Then
                    strD += "row(""TechCaption"",""" & qc.TechCaption & """,""value"",calculate(" & qc.UniName & ",s0))," & Chr(13) & Chr(10)
                End If
                strD = strD.Substring(0, strD.Length - 3) & ""
            End If

            Return strD

        End Get
    End Property

    Public Function FormatDax(strDax As String, intL As Integer) As String

        Dim strT As String = New String(" ", intL)


        Dim l() As String = Split(strDax, vbCrLf)
        Dim r As New List(Of String)
        For Each s As String In l

            If s.Trim.ToUpper.StartsWith("DEFINE") Then
                s = s.Trim
            ElseIf s.Trim.ToUpper.StartsWith("VAR ") Then
                s = strT & s.Trim
            ElseIf s.Trim.ToUpper.StartsWith("FILTER") Then
                s = strT & strT & s.Trim
            ElseIf s.Trim.ToUpper.StartsWith("SUMMARIZE") Then
                s = strT & strT & s.Trim
            ElseIf s.Trim.ToUpper.StartsWith("SELECTCOL") Then
                s = strT & strT & s.Trim
            ElseIf s.Trim.ToUpper.StartsWith(")") Then
                s = strT & strT & s.Trim
            ElseIf s.Trim.ToUpper.StartsWith("TOPN") Then
                s = strT & strT & s.Trim
            ElseIf s.Trim.ToUpper.StartsWith("EVALUATE") Then
                s = vbCrLf & s.Trim
            ElseIf s.Trim.ToLower.StartsWith("x1") Then
                s = s.Trim
            ElseIf s.Trim.ToLower.StartsWith("s0") Then
                s = s.Trim
            ElseIf s.Trim.ToUpper.StartsWith("ORDER BY") Then
                s = s.Trim
            Else
                s = strT & strT & strT & s.Trim
            End If


            r.Add(s)
        Next s

        Dim strRes As String = ""
        For Each s In r
            strRes += s & vbCrLf
        Next s





        Return strRes
    End Function




    Public ReadOnly Property DAX(blnPreview As Boolean) As String
        Get

            If Me.QueryColumns.Count = 0 Then
                Dim strD As String = "EVALUATE ROW(""pbixl nothing selected"",Blank())"
                Return strD
            End If




            If Me.LevelFilter = False And Me.MeasureFilter = False Then

                Dim strD As String = "DEFINE" & vbCrLf
                strD += "VAR x0 = " & vbCrLf & "SUMMARIZECOLUMNS(" & vbCrLf
                For Each c As clsQueryColumn In Me.QueryColumns
                    If c.FieldType = clsQueryColumn.enFieldType.Level Then
                        strD = strD & c.UniName & "," & vbCrLf
                    End If
                Next c
                For Each c As clsQueryColumn In Me.QueryColumns
                    If c.FieldType = clsQueryColumn.enFieldType.Measure Then
                        strD = strD & """" & c.FieldName & """," & c.UniName & "," & vbCrLf
                    End If
                Next c
                strD = strD.Substring(0, strD.Length - 3)
                strD += vbCrLf & ")" & vbCrLf

                strD += "VAR s0 = " & vbCrLf & "SELECTCOLUMNS(" & vbCrLf & "x0," & vbCrLf
                For Each c As clsQueryColumn In Me.AllQueryColumnsSortedByOrdinal 'In Me.QueryColumns
                    If c.FieldType = clsQueryColumn.enFieldType.Level Then
                        strD += """" & c.UniName & """," & c.UniName & "," & vbCrLf
                    Else
                        strD += """" & c.FieldName & """," & c.UniName & "," & vbCrLf
                    End If
                Next c
                strD = strD.Substring(0, strD.Length - 3) & vbCrLf & ")" & vbCrLf

                If Not Me.OrderColumn Is Nothing And blnPreview = True Then
                    strD += "VAR x1 = " & vbCrLf & "TOPN(1000,s0," & Me.OrderColumn.UniName & "," & Me.OrderColumn.Sort.ToString & ")" & vbCrLf
                    strD += "EVALUATE x1 " & vbCrLf & "ORDER BY " & Me.OrderColumn.UniName & " " & Me.OrderColumn.Sort.ToString.ToUpper
                ElseIf Me.OrderColumn Is Nothing And blnPreview = True Then
                    strD += "VAR x1 = " & vbCrLf & "TOPN(1000,s0)" & vbCrLf
                    strD += "EVALUATE" & vbCrLf & " x1"
                ElseIf Not Me.OrderColumn Is Nothing And blnPreview = False Then
                    strD += "EVALUATE " & vbCrLf & "s0 " & vbCrLf & "ORDER BY " & Me.OrderColumn.UniName & " " & Me.OrderColumn.Sort.ToString.ToUpper
                ElseIf Me.OrderColumn Is Nothing And blnPreview = False Then
                    strD += "EVALUATE " & vbCrLf & "s0"
                End If

                Return strD

            End If

            If Me.LevelFilter = True And Me.MeasureFilter = False Then

                Dim strD As String = "DEFINE" & vbCrLf

                Dim ctrF As Integer = 0
                For Each c As clsQueryColumn In Me.QueryColumns
                    If c.FieldType = clsQueryColumn.enFieldType.Level AndAlso c.DaxFilter <> "" Then
                        ctrF += 1
                        strD += "VAR f" & ctrF.ToString & " = " & vbCrLf & "FILTER('" & c.TableName & "'," & c.DaxFilter & ")" & vbCrLf
                    End If
                Next c

                strD += "VAR x0 = " & vbCrLf & " SUMMARIZECOLUMNS(" & vbCrLf
                For Each c As clsQueryColumn In Me.QueryColumns
                    If c.FieldType = clsQueryColumn.enFieldType.Level Then
                        strD += c.UniName & "," & vbCrLf
                    End If
                Next c
                For i As Integer = 1 To ctrF
                    strD += "f" & i.ToString & "," & vbCrLf
                Next i
                For Each c As clsQueryColumn In Me.QueryColumns
                    If c.FieldType = clsQueryColumn.enFieldType.Measure Then
                        strD += """" & c.FieldName & """," & c.UniName & "," & vbCrLf
                    End If
                Next c
                strD = strD.Substring(0, strD.Length - 3) & vbCrLf & ")" & vbCrLf

                strD += "VAR s0 = " & vbCrLf & "SELECTCOLUMNS( " & vbCrLf & "x0," & vbCrLf
                For Each c As clsQueryColumn In Me.AllQueryColumnsSortedByOrdinal
                    If c.FieldType = clsQueryColumn.enFieldType.Level Then
                        strD += """" & c.UniName & """," & c.UniName & "," & vbCrLf
                    Else
                        strD += """" & c.FieldName & """," & c.UniName & "," & vbCrLf
                    End If
                Next c
                strD = strD.Substring(0, strD.Length - 3) & vbCrLf & ")" & vbCrLf

                If Not Me.OrderColumn Is Nothing And blnPreview = True Then
                    strD += "VAR x1 = " & vbCrLf & "TOPN(1000,s0," & Me.OrderColumn.UniName & "," & Me.OrderColumn.Sort.ToString & ")" & vbCrLf
                    strD += "EVALUATE x1 " & vbCrLf & "ORDER BY " & Me.OrderColumn.UniName & " " & Me.OrderColumn.Sort.ToString.ToUpper
                ElseIf Me.OrderColumn Is Nothing And blnPreview = True Then
                    strD += "VAR x1 = " & vbCrLf & "TOPN(1000,s0)" & vbCrLf
                    strD += "EVALUATE " & vbCrLf & "x1"
                ElseIf Not Me.OrderColumn Is Nothing And blnPreview = False Then
                    strD += "EVALUATE " & vbCrLf & "s0 " & vbCrLf & "ORDER BY " & Me.OrderColumn.UniName & " " & Me.OrderColumn.Sort.ToString.ToUpper
                ElseIf Me.OrderColumn Is Nothing And blnPreview = False Then
                    strD += "EVALUATE " & vbCrLf & "s0"
                End If

                Return strD

            End If

            If Me.LevelFilter = True And Me.MeasureFilter = True Then

                Dim strD As String = "DEFINE" & vbCrLf

                Dim ctrF As Integer = 0
                For Each c As clsQueryColumn In Me.QueryColumns
                    If c.FieldType = clsQueryColumn.enFieldType.Level AndAlso c.DaxFilter <> "" Then
                        ctrF += 1
                        strD += "VAR f" & ctrF.ToString & "=" & vbCrLf & "FILTER('" & c.TableName & "'," & c.DaxFilter & ")" & vbCrLf
                    End If
                Next c

                strD += "VAR x0 = " & vbCrLf & "SUMMARIZECOLUMNS(" & vbCrLf
                For Each c As clsQueryColumn In Me.QueryColumns
                    If c.FieldType = clsQueryColumn.enFieldType.Level Then
                        strD += c.UniName & "," & vbCrLf
                    End If
                Next c
                For i As Integer = 1 To ctrF
                    strD += "f" & i.ToString & "," & vbCrLf
                Next i
                For Each c As clsQueryColumn In Me.QueryColumns
                    If c.FieldType = clsQueryColumn.enFieldType.Measure Then
                        strD += """" & c.FieldName & """," & c.UniName & "," & vbCrLf
                    End If
                Next c
                strD = strD.Substring(0, strD.Length - 3) & vbCrLf & ")" & vbCrLf

                strD += "VAR m0 = " & vbCrLf & "FILTER(x0,"
                For Each c As clsQueryColumn In Me.QueryColumns
                    If c.FieldType = clsQueryColumn.enFieldType.Measure Then
                        If c.DaxFilter <> "" Then
                            strD += c.DaxFilter & " && "
                        End If
                    End If
                Next c
                strD = strD.Substring(0, strD.Length - 4) & ")" & vbCrLf

                strD += "VAR s0 = " & vbCrLf & "SELECTCOLUMNS( " & vbCrLf & "m0," & vbCrLf
                For Each c As clsQueryColumn In Me.AllQueryColumnsSortedByOrdinal
                    If c.FieldType = clsQueryColumn.enFieldType.Level Then
                        strD += """" & c.UniName & """," & c.UniName & "," & vbCrLf
                    Else
                        strD += """" & c.FieldName & """," & c.UniName & "," & vbCrLf
                    End If

                Next c
                strD = strD.Substring(0, strD.Length - 3) & vbCrLf & ")" & vbCrLf

                If Not Me.OrderColumn Is Nothing And blnPreview = True Then
                    strD += "VAR x1 = " & vbCrLf & "TOPN(1000,s0," & Me.OrderColumn.UniName & "," & Me.OrderColumn.Sort.ToString & ")" & vbCrLf
                    strD += "EVALUATE " & vbCrLf & "x1 " & vbCrLf & "ORDER BY " & Me.OrderColumn.UniName & " " & Me.OrderColumn.Sort.ToString.ToUpper
                ElseIf Me.OrderColumn Is Nothing And blnPreview = True Then
                    strD += "VAR x1 = " & vbCrLf & "TOPN(1000,s0)" & vbCrLf
                    strD += "EVALUATE " & vbCrLf & "x1"
                ElseIf Not Me.OrderColumn Is Nothing And blnPreview = False Then
                    strD += "EVALUATE " & vbCrLf & "s0 " & vbCrLf & "ORDER BY " & Me.OrderColumn.UniName & " " & Me.OrderColumn.Sort.ToString.ToUpper
                ElseIf Me.OrderColumn Is Nothing And blnPreview = False Then
                    strD += "EVALUATE " & vbCrLf & "s0"
                End If

                Return strD

            End If

            If Me.LevelFilter = False And Me.MeasureFilter = True Then

                Dim strD As String = "DEFINE" & vbCrLf

                strD += "VAR x0 = " & vbCrLf & "SUMMARIZECOLUMNS(" & vbCrLf
                For Each c As clsQueryColumn In Me.QueryColumns
                    If c.FieldType = clsQueryColumn.enFieldType.Level Then
                        strD += c.UniName & "," & vbCrLf
                    End If
                Next c
                For Each c As clsQueryColumn In Me.QueryColumns
                    If c.FieldType = clsQueryColumn.enFieldType.Measure Then
                        strD += """" & c.FieldName & """," & c.UniName & "," & vbCrLf
                    End If
                Next c
                strD = strD.Substring(0, strD.Length - 3) & vbCrLf & ")"
                strD += vbCrLf

                strD += "VAR m0 = " & vbCrLf & "FILTER(x0,"
                For Each c As clsQueryColumn In Me.QueryColumns
                    If c.FieldType = clsQueryColumn.enFieldType.Measure Then
                        If c.DaxFilter <> "" Then
                            strD += c.DaxFilter & " && "
                        End If
                    End If
                Next c
                strD = strD.Substring(0, strD.Length - 4) & ")" & vbCrLf

                strD += "VAR s0 = " & vbCrLf & "SELECTCOLUMNS( " & vbCrLf & "m0," & vbCrLf
                For Each c As clsQueryColumn In Me.AllQueryColumnsSortedByOrdinal
                    If c.FieldType = clsQueryColumn.enFieldType.Level Then
                        strD += """" & c.UniName & """," & c.UniName & "," & vbCrLf
                    Else
                        strD += """" & c.FieldName & """," & c.UniName & "," & vbCrLf
                    End If
                Next c
                strD = strD.Substring(0, strD.Length - 3) & vbCrLf & ")" & vbCrLf

                If Not Me.OrderColumn Is Nothing And blnPreview = True Then
                    strD += "VAR x1 = " & vbCrLf & "TOPN(1000,s0," & Me.OrderColumn.UniName & "," & Me.OrderColumn.Sort.ToString & ")" & vbCrLf
                    strD += "EVALUATE " & vbCrLf & "x1 " & vbCrLf & "ORDER BY " & Me.OrderColumn.UniName & " " & Me.OrderColumn.Sort.ToString.ToUpper
                ElseIf Me.OrderColumn Is Nothing And blnPreview = True Then
                    strD += "VAR x1 = " & vbCrLf & "TOPN(1000,s0)" & vbCrLf
                    strD += "EVALUATE " & vbCrLf & "x1"
                ElseIf Not Me.OrderColumn Is Nothing And blnPreview = False Then
                    strD += "EVALUATE " & vbCrLf & "s0 " & vbCrLf & "ORDER BY " & Me.OrderColumn.UniName & " " & Me.OrderColumn.Sort.ToString.ToUpper
                ElseIf Me.OrderColumn Is Nothing And blnPreview = False Then
                    strD += "EVALUATE " & vbCrLf & "s0"
                End If

                Return strD

            End If

            Return ""

        End Get
    End Property




    'Public ReadOnly Property DAX(blnPreview As Boolean) As String
    '    Get

    '        If Me.QueryColumns.Count = 0 Then
    '            Dim strD As String = "EVALUATE ROW(""pbixl Nothing selected"",Blank())"
    '            Return strD
    '        End If



    '        If Me.LevelFilter = False And Me.MeasureFilter = False Then

    '            Dim strD As String = "DEFINE" & Chr(13) & Chr(10)
    '            strD += "var x0 = SUMMARIZECOLUMNS(" & Chr(13) & Chr(10)
    '            For Each c As clsQueryColumn In Me.QueryColumns
    '                If c.FieldType = clsQueryColumn.enFieldType.Level Then
    '                    strD = strD & c.UniName & "," & Chr(13) & Chr(10)
    '                End If
    '            Next c
    '            For Each c As clsQueryColumn In Me.QueryColumns
    '                If c.FieldType = clsQueryColumn.enFieldType.Measure Then
    '                    strD = strD & """" & c.FieldName & """," & c.UniName & "," & Chr(13) & Chr(10)
    '                End If
    '            Next c
    '            strD = strD.Substring(0, strD.Length - 3)
    '            strD += ")" & Chr(13) & Chr(10)

    '            strD += "var s0 = SELECTCOLUMNS(x0,"
    '            For Each c As clsQueryColumn In Me.QueryColumns
    '                If c.FieldType = clsQueryColumn.enFieldType.Level Then
    '                    strD += """" & c.TechCaption & """," & c.UniName & ","
    '                Else
    '                    strD += """" & c.TechCaption & """," & c.UniName & ","
    '                End If
    '            Next c
    '            strD = strD.Substring(0, strD.Length - 1) & ")" & Chr(13) & Chr(10)

    '            If Not Me.OrderColumn Is Nothing And blnPreview = True Then

    '                strD += "var x1 = TOPN(1000,s0," & Chr(13) & Chr(10)
    '                strD += "[" & Me.OrderColumn.TechCaption & "]," & Me.OrderColumn.Sort.ToString & Chr(13) & Chr(10) & ","
    '                For Each qc In Me.QueryColumns
    '                    If Not qc Is Me.OrderColumn Then
    '                        strD += "[" & qc.TechCaption & "],False" & Chr(13) & Chr(10) & ","
    '                    End If
    '                Next qc
    '                strD = strD.Substring(0, strD.Length - 1) & ")" & Chr(13) & Chr(10)
    '                strD += "EVALUATE x1 ORDER BY [" & Me.OrderColumn.TechCaption & "] " & Me.OrderColumn.Sort.ToString

    '            ElseIf Me.OrderColumn Is Nothing And blnPreview = True Then

    '                strD += "var x1 = TOPN(1000,s0,"
    '                For Each qc In Me.QueryColumns
    '                    strD += "[" & qc.TechCaption & "],False" & Chr(13) & Chr(10) & ","
    '                Next qc
    '                strD = strD.Substring(0, strD.Length - 1) & ")" & Chr(13) & Chr(10)
    '                strD += "EVALUATE x1"

    '            ElseIf Not Me.OrderColumn Is Nothing And blnPreview = False Then
    '                strD += "EVALUATE s0 ORDER BY [" & Me.OrderColumn.TechCaption & "] " & Me.OrderColumn.Sort.ToString
    '            ElseIf Me.OrderColumn Is Nothing And blnPreview = False Then
    '                strD += "EVALUATE s0"
    '            End If

    '            Return strD

    '        End If

    '        If Me.LevelFilter = True And Me.MeasureFilter = False Then

    '            Dim strD As String = "DEFINE" & Chr(13) & Chr(10)

    '            Dim ctrF As Integer = 0
    '            For Each c As clsQueryColumn In Me.QueryColumns
    '                If c.FieldType = clsQueryColumn.enFieldType.Level AndAlso c.DaxFilter <> "" Then
    '                    ctrF += 1
    '                    strD += "var f" & ctrF.ToString & "=Filter('" & c.TableName & "'," & c.DaxFilter & ")" & Chr(13) & Chr(10)
    '                End If
    '            Next c

    '            strD += "var x0 = SUMMARIZECOLUMNS(" & Chr(13) & Chr(10)
    '            For Each c As clsQueryColumn In Me.QueryColumns
    '                If c.FieldType = clsQueryColumn.enFieldType.Level Then
    '                    strD += c.UniName & ","
    '                End If
    '            Next c
    '            For i As Integer = 1 To ctrF
    '                strD += "f" & i.ToString & ","
    '            Next i
    '            For Each c As clsQueryColumn In Me.QueryColumns
    '                If c.FieldType = clsQueryColumn.enFieldType.Measure Then
    '                    strD += """" & c.FieldName & """," & c.UniName & ","
    '                End If
    '            Next c
    '            strD = strD.Substring(0, strD.Length - 1) & ")" & Chr(13) & Chr(10)

    '            strD += "var s0 = SELECTCOLUMNS(x0,"
    '            For Each c As clsQueryColumn In Me.QueryColumns
    '                If c.FieldType = clsQueryColumn.enFieldType.Level Then
    '                    strD += """" & c.TechCaption & """," & c.UniName & ","
    '                Else
    '                    strD += """" & c.TechCaption & """," & c.UniName & ","
    '                End If
    '            Next c
    '            strD = strD.Substring(0, strD.Length - 1) & ")" & Chr(13) & Chr(10)

    '            If Not Me.OrderColumn Is Nothing And blnPreview = True Then

    '                strD += "var x1 = TOPN(1000,s0," & Chr(13) & Chr(10)
    '                strD += "[" & Me.OrderColumn.TechCaption & "]," & Me.OrderColumn.Sort.ToString & Chr(13) & Chr(10) & ","

    '                For Each qc In Me.QueryColumns
    '                    If Not qc Is Me.OrderColumn Then
    '                        strD += "[" & qc.TechCaption & "],false" & Chr(13) & Chr(10) & ","
    '                    End If
    '                Next qc
    '                strD = strD.Substring(0, strD.Length - 1) & ")" & Chr(13) & Chr(10)
    '                strD += "EVALUATE x1 ORDER BY [" & Me.OrderColumn.TechCaption & "] " & Me.OrderColumn.Sort.ToString


    '            ElseIf Me.OrderColumn Is Nothing And blnPreview = True Then

    '                strD += "var x1 = TOPN(1000,s0,"
    '                For Each qc In Me.QueryColumns
    '                    strD += "[" & qc.TechCaption & "],false" & Chr(13) & Chr(10) & ","
    '                Next qc
    '                strD = strD.Substring(0, strD.Length - 1) & ")" & Chr(13) & Chr(10)
    '                strD += "EVALUATE x1"

    '            ElseIf Not Me.OrderColumn Is Nothing And blnPreview = False Then
    '                strD += "EVALUATE s0 ORDER BY [" & Me.OrderColumn.TechCaption & "] " & Me.OrderColumn.Sort.ToString
    '            ElseIf Me.OrderColumn Is Nothing And blnPreview = False Then
    '                strD += "EVALUATE s0"
    '            End If

    '            Return strD

    '        End If

    '        If Me.LevelFilter = True And Me.MeasureFilter = True Then

    '            Dim strD As String = "DEFINE" & Chr(13) & Chr(10)

    '            Dim ctrF As Integer = 0
    '            For Each c As clsQueryColumn In Me.QueryColumns
    '                If c.FieldType = clsQueryColumn.enFieldType.Level AndAlso c.DaxFilter <> "" Then
    '                    ctrF += 1
    '                    strD += "var f" & ctrF.ToString & "=Filter('" & c.TableName & "'," & c.DaxFilter & ")" & Chr(13) & Chr(10)
    '                End If
    '            Next c

    '            strD += "var x0 = SUMMARIZECOLUMNS(" & Chr(13) & Chr(10)
    '            For Each c As clsQueryColumn In Me.QueryColumns
    '                If c.FieldType = clsQueryColumn.enFieldType.Level Then
    '                    strD += c.UniName & ","
    '                End If
    '            Next c
    '            For i As Integer = 1 To ctrF
    '                strD += "f" & i.ToString & ","
    '            Next i
    '            For Each c As clsQueryColumn In Me.QueryColumns
    '                If c.FieldType = clsQueryColumn.enFieldType.Measure Then
    '                    strD += """" & c.FieldName & """," & c.UniName & ","
    '                End If
    '            Next c
    '            strD = strD.Substring(0, strD.Length - 1) & ")" & Chr(13) & Chr(10)

    '            strD += "var m0 = Filter(x0,"
    '            For Each c As clsQueryColumn In Me.QueryColumns
    '                If c.FieldType = clsQueryColumn.enFieldType.Measure Then
    '                    If c.DaxFilter <> "" Then
    '                        strD += c.DaxFilter & " && "
    '                    End If
    '                End If
    '            Next c
    '            strD = strD.Substring(0, strD.Length - 4) & ")" & Chr(13) & Chr(10)

    '            strD += "var s0 = SELECTCOLUMNS(m0,"
    '            For Each c As clsQueryColumn In Me.QueryColumns
    '                If c.FieldType = clsQueryColumn.enFieldType.Level Then
    '                    strD += """" & c.TechCaption & """," & c.UniName & ","
    '                Else
    '                    strD += """" & c.TechCaption & """," & c.UniName & ","
    '                End If

    '            Next c
    '            strD = strD.Substring(0, strD.Length - 1) & ")" & Chr(13) & Chr(10)

    '            If Not Me.OrderColumn Is Nothing And blnPreview = True Then

    '                strD += "var x1 = TOPN(1000,s0," & Chr(13) & Chr(10)
    '                strD += "[" & Me.OrderColumn.TechCaption & "]," & Me.OrderColumn.Sort.ToString & Chr(13) & Chr(10) & ","
    '                For Each qc In Me.QueryColumns
    '                    If Not qc Is Me.OrderColumn Then
    '                        strD += "[" & qc.TechCaption & "],false" & Chr(13) & Chr(10) & ","
    '                    End If
    '                Next qc
    '                strD = strD.Substring(0, strD.Length - 1) & ")" & Chr(13) & Chr(10)
    '                strD += "EVALUATE x1 ORDER BY [" & Me.OrderColumn.TechCaption & "] " & Me.OrderColumn.Sort.ToString

    '            ElseIf Me.OrderColumn Is Nothing And blnPreview = True Then

    '                strD += "var x1 = TOPN(1000,s0,"
    '                For Each qc In Me.QueryColumns
    '                    strD += "[" & qc.TechCaption & "],false" & Chr(13) & Chr(10) & ","
    '                Next qc
    '                strD = strD.Substring(0, strD.Length - 1) & ")" & Chr(13) & Chr(10)
    '                strD += "EVALUATE x1"

    '            ElseIf Not Me.OrderColumn Is Nothing And blnPreview = False Then
    '                strD += "EVALUATE s0 ORDER BY [" & Me.OrderColumn.TechCaption & "] " & Me.OrderColumn.Sort.ToString
    '            ElseIf Me.OrderColumn Is Nothing And blnPreview = False Then
    '                strD += "EVALUATE s0"
    '            End If

    '            Return strD

    '        End If

    '        If Me.LevelFilter = False And Me.MeasureFilter = True Then

    '            Dim strD As String = "DEFINE" & Chr(13) & Chr(10)

    '            strD += "var x0 = SUMMARIZECOLUMNS(" & Chr(13) & Chr(10)
    '            For Each c As clsQueryColumn In Me.QueryColumns
    '                If c.FieldType = clsQueryColumn.enFieldType.Level Then
    '                    strD += c.UniName & ","
    '                End If
    '            Next c
    '            For Each c As clsQueryColumn In Me.QueryColumns
    '                If c.FieldType = clsQueryColumn.enFieldType.Measure Then
    '                    strD += """" & c.FieldName & """," & c.UniName & ","
    '                End If
    '            Next c
    '            strD = strD.Substring(0, strD.Length - 1) & ")"
    '            strD += Chr(13) & Chr(10)

    '            strD += "var m0 = Filter(x0,"
    '            For Each c As clsQueryColumn In Me.QueryColumns
    '                If c.FieldType = clsQueryColumn.enFieldType.Measure Then
    '                    If c.DaxFilter <> "" Then
    '                        strD += c.DaxFilter & " && "
    '                    End If
    '                End If
    '            Next c
    '            strD = strD.Substring(0, strD.Length - 4) & ")" & Chr(13) & Chr(10)

    '            strD += "var s0 = SELECTCOLUMNS(m0,"
    '            For Each c As clsQueryColumn In Me.QueryColumns
    '                If c.FieldType = clsQueryColumn.enFieldType.Level Then
    '                    strD += """" & c.TechCaption & """," & c.UniName & ","
    '                Else
    '                    strD += """" & c.TechCaption & """," & c.UniName & ","
    '                End If
    '            Next c
    '            strD = strD.Substring(0, strD.Length - 1) & ")" & Chr(13) & Chr(10)

    '            If Not Me.OrderColumn Is Nothing And blnPreview = True Then

    '                strD += "var x1 = TOPN(1000,s0," & Chr(13) & Chr(10)
    '                strD += "[" & Me.OrderColumn.TechCaption & "]," & Me.OrderColumn.Sort.ToString & Chr(13) & Chr(10) & ","
    '                For Each qc In Me.QueryColumns
    '                    If Not qc Is Me.OrderColumn Then
    '                        strD += "[" & qc.TechCaption & "],false" & Chr(13) & Chr(10) & ","
    '                    End If
    '                Next qc
    '                strD = strD.Substring(0, strD.Length - 1) & ")" & Chr(13) & Chr(10)
    '                strD += "EVALUATE x1 ORDER BY [" & Me.OrderColumn.TechCaption & "] " & Me.OrderColumn.Sort.ToString

    '            ElseIf Me.OrderColumn Is Nothing And blnPreview = True Then

    '                strD += "var x1 = TOPN(1000,s0,"
    '                For Each qc In Me.QueryColumns
    '                    strD += "[" & qc.TechCaption & "],false" & Chr(13) & Chr(10) & ","
    '                Next qc
    '                strD = strD.Substring(0, strD.Length - 1) & ")" & Chr(13) & Chr(10)
    '                strD += "EVALUATE x1"

    '            ElseIf Not Me.OrderColumn Is Nothing And blnPreview = False Then
    '                strD += "EVALUATE s0 ORDER BY [" & Me.OrderColumn.TechCaption & "] " & Me.OrderColumn.Sort.ToString
    '            ElseIf Me.OrderColumn Is Nothing And blnPreview = False Then
    '                strD += "EVALUATE s0"
    '            End If

    '            Return strD

    '        End If

    '        Return ""

    '    End Get
    'End Property

    Public ReadOnly Property InnerFilter As Boolean
        Get
            If LevelFilter = True OrElse MeasureFilter = True Then
                Return True
            Else
                Return False
            End If
        End Get
    End Property


    Public ReadOnly Property LevelFilter As Boolean
        Get
            For Each c As clsQueryColumn In Me.QueryColumns
                If c.FieldType = clsQueryColumn.enFieldType.Level Then
                    If c.DaxFilter.Trim <> "" Then
                        Return True
                    End If
                End If
            Next c
            Return False
        End Get
    End Property

    Public ReadOnly Property MeasureFilter As Boolean
        Get
            For Each c As clsQueryColumn In Me.QueryColumns
                If c.FieldType = clsQueryColumn.enFieldType.Measure Then
                    If c.DaxFilter.Trim <> "" Then
                        Return True
                    End If
                End If
            Next c
            Return False
        End Get
    End Property

    Public ReadOnly Property OrderColumn As clsQueryColumn
        Get
            For Each c As clsQueryColumn In Me.QueryColumns
                If c.FieldType = clsQueryColumn.enFieldType.Measure Then
                    If c.Sort <> clsQueryColumn.enSort.none Then
                        Return c
                    End If
                ElseIf c.FieldType = clsQueryColumn.enFieldType.Level Then
                    If c.Sort <> clsQueryColumn.enSort.none Then
                        Return c
                    End If
                End If
            Next c

            For Each qc In Me.AllQueryColumnsSortedByOrdinal
                qc.Sort = clsQueryColumn.enSort.asc
                Return qc
            Next qc


            Return Nothing
        End Get
    End Property

    Public ReadOnly Property MeasureCount As Integer
        Get
            Dim ctr As Integer = 0
            For Each c As clsQueryColumn In Me.QueryColumns
                If c.FieldType = clsQueryColumn.enFieldType.Measure Then
                    ctr += 1
                End If
            Next c
            Return ctr
        End Get
    End Property

    Public ReadOnly Property LevelCount As Integer
        Get
            Dim ctr As Integer = 0
            For Each c As clsQueryColumn In Me.QueryColumns
                If c.FieldType = clsQueryColumn.enFieldType.Level Then
                    ctr += 1
                End If
            Next c
            Return ctr
        End Get
    End Property


    Public Function AllQueryColumnsSortedByOrdinal() As List(Of clsQueryColumn)
        Me.QueryColumns.Sort(Function(x As clsQueryColumn, y As clsQueryColumn) (x.Ordinal).CompareTo(y.Ordinal))
        Return Me.QueryColumns
    End Function






    Public Sub New()

        Me.FilterControlVisible = False
        Me.FilterControlGUID = ""

        'Me.IsInternal = False
        Me.AddMissingItems = False

        Me.GUID = System.Guid.NewGuid.ToString
        Me.QueryColumns = New List(Of clsQueryColumn)

    End Sub

    Public Function GetQueryColumn(UniName As String) As clsQueryColumn
        For Each _c As clsQueryColumn In Me.QueryColumns
            If _c.UniName.ToLower = UniName.ToLower Then
                Return _c
            End If
        Next _c
        Return Nothing
    End Function

    Public Function ColumnIsInQuery(UniName As String) As Boolean
        For Each c As clsQueryColumn In Me.QueryColumns
            If c.UniName.ToLower = UniName.ToLower AndAlso c.IsSelected = True Then
                Return True
            End If
        Next c
        Return False
    End Function

    Public Sub RemoveColumn(UniName As String)
        For i As Integer = Me.QueryColumns.Count - 1 To 0 Step -1
            If Me.QueryColumns.Item(i).UniName.ToLower = UniName.ToLower Then
                Me.QueryColumns.RemoveAt(i)
            End If
        Next i
    End Sub

    Public Sub AddColumn(qc As clsQueryColumn, AfterOrdinal As Integer)

        For Each _qc In Me.QueryColumns
            If _qc.Ordinal > AfterOrdinal Then
                _qc.Ordinal += 1000000
            End If
        Next _qc
        qc.Ordinal = 100000
        Me.QueryColumns.Add(qc)

        Dim ctr As Integer = -1
        For Each _qc In Me.AllQueryColumnsSortedByOrdinal
            ctr += 1
            _qc.Ordinal = ctr
        Next _qc

    End Sub




End Class


<Serializable>
Public Class clsQueryColumn

    Enum enFieldType
        Level = 1
        Measure = 2
    End Enum

    Enum enDataType
        Text = 1
        Number = 2
        DateTime = 3
        Bool = 4
    End Enum

    Public Enum enSelectionMode
        AllSelected = 1
        AllSearch = 2
        DeSelectMember = 3
        SelectMember = 4
    End Enum

    Public Enum enSort
        none = 0
        asc = 1
        desc = 2
    End Enum

    Public Property GUID As String
    Public Property UniName As String

    Public ReadOnly Property EntityName As String
        Get
            Return Me.UniName.Substring(0, UniName.IndexOf("[") - 1).Replace("'", "")
        End Get
    End Property

    Public ReadOnly Property ReferenceName As String
        Get
            Return Me.UniName.Substring(UniName.IndexOf("[") + 1).Replace("]", "")
        End Get
    End Property

    Public ReadOnly Property TechCaption As String
        Get
            Return Me.GUID.Substring(0, 8)
        End Get
    End Property


    Public Property TableName As String
    Public Property FieldName As String

    Friend Property Query As clsQuery

    Public Property FieldType As enFieldType
    Public Property DataType As enDataType
    Public Property SelectionMode As enSelectionMode


    Private Property __Sort As enSort
    Public Property Sort As enSort
        Get
            Return __Sort
        End Get
        Set(value As enSort)
            Me.__Sort = value
        End Set
    End Property





    Public Property SearchTerm As String

    Public Property Ordinal As Integer

    Public Property FilterControlGUID As String

    Friend Property htSel As Hashtable

    Private _BlankSel As Boolean
    Public Property BlankSel As Boolean
        Get
            Return Me._BlankSel
        End Get
        Set(value As Boolean)
            Me._BlankSel = value
        End Set
    End Property

    Private _lstSel As List(Of Object)
    Public Property lstSel As List(Of Object)
        Get
            Return Me._lstSel
        End Get
        Set(value As List(Of Object))
            Me._lstSel = value
        End Set
    End Property

    Public ReadOnly Property IsFiltered As Boolean
        Get
            Dim blnRes As Boolean = False
            If Me.SelectionMode = enSelectionMode.AllSelected Then
                blnRes = False
            ElseIf Me.SelectionMode = enSelectionMode.AllSearch And Me.DaxFilter <> "" Then
                blnRes = True
            ElseIf Me.SelectionMode = enSelectionMode.SelectMember Or Me.SelectionMode = enSelectionMode.DeSelectMember Then
                If Me.htSel Is Nothing OrElse Me.htSel.Count = 0 Then
                    blnRes = False
                Else
                    blnRes = True
                End If
            End If
            Return blnRes
        End Get
    End Property

    Public Property DaxStmnt As String


    Private _DaxFilter As String
    Public Property DaxFilter As String
        Get
            If Me.SelectionMode = enSelectionMode.AllSearch Then
                Return Me._DaxFilter
            ElseIf Me.SelectionMode = enSelectionMode.SelectMember OrElse Me.SelectionMode = enSelectionMode.DeSelectMember Then
                If Me.htSel Is Nothing OrElse Me.htSel.Count = 0 Then
                    Return ""
                End If

                Dim strR As String = ""
                If Me.SelectionMode = enSelectionMode.SelectMember Then
                    strR = Me.UniName & " IN {"
                Else
                    strR = " NOT" & Me.UniName & " IN {"
                End If

                For Each s In Me.htSel
                    If Me.DataType = enDataType.Bool Then
                        If s.key Is DBNull.Value Then
                            strR += "BLANK(),"
                        ElseIf s.key = True Then
                            strR += "true,"
                        Else
                            strR += "false,"
                        End If
                    ElseIf Me.DataType = enDataType.DateTime Then
                        If s.key Is DBNull.Value Then
                            strR += "BLANK(),"
                        Else
                            strR += "(Date(" & DateAndTime.Year(s.key) & "," & DateAndTime.Month(s.key) & "," & DateAndTime.Day(s.key) & ") + " _
                            & "TIME(" & DateAndTime.Hour(s.key) & "," & DateAndTime.Minute(s.key) & "," & DateAndTime.Second(s.key) & ")),"
                        End If
                    ElseIf Me.DataType = enDataType.Number Then
                        If s.key Is DBNull.Value Then
                            strR += "BLANK(),"
                        Else
                            strR += s.key.ToString.Replace(",", ".") & ","
                        End If
                    ElseIf Me.DataType = enDataType.Text Then
                        If s.key Is DBNull.Value Then
                            strR += "BLANK(),"
                        Else
                            strR += """" & s.key & ""","
                        End If
                    End If
                Next s
                strR = strR.Substring(0, strR.Length - 1) & "}"
                Return strR
            End If
            Return ""
        End Get
        Set(value As String)
            Me._DaxFilter = value
        End Set
    End Property

    Public ReadOnly Property DaxFilterTable As String
        Get
            Dim strD As String = ""
            If Me.FieldType = enFieldType.Level Then
                strD += "var fx = Filter(Values(" & Me.UniName & ")," & Me.UniName & " IN {"
                If Me.DataType = clsQueryColumn.enDataType.Text Then
                    For Each m In Me.SelectedMember
                        If IsDBNull(m) Then
                            strD += "blank(),"
                        Else
                            strD += """" & m & ""","
                        End If
                    Next m
                    strD = strD.Substring(0, strD.Length - 1) & "})" & Chr(13) & Chr(10)
                ElseIf Me.DataType = clsQueryColumn.enDataType.DateTime Then
                    For Each m In Me.SelectedMember
                        If IsDBNull(m) Then
                            strD += "blank(),"
                        Else
                            Dim strR As String = "(Date(" & DateAndTime.Year(m) & "," & DateAndTime.Month(m) & "," & DateAndTime.Day(m) & ") + " _
                            & "TIME(" & DateAndTime.Hour(m) & "," & DateAndTime.Minute(m) & "," & DateAndTime.Second(m) & "))"
                            strD += strR & ","
                        End If

                    Next m
                    strD = strD.Substring(0, strD.Length - 1) & "})" & Chr(13) & Chr(10)
                ElseIf Me.DataType = clsQueryColumn.enDataType.Bool Then
                    For Each m In Me.SelectedMember
                        If IsDBNull(m) Then
                            strD += "blank(),"
                        Else
                            If m.ToString.ToLower = "true" Then
                                strD += "true,"
                            Else
                                strD += "false,"
                            End If
                        End If
                    Next m
                    strD = strD.Substring(0, strD.Length - 1) & "})" & Chr(13) & Chr(10)
                Else
                    For Each m In Me.SelectedMember
                        If IsDBNull(m) Then
                            strD += "blank(),"
                        Else
                            strD += m & ","
                        End If
                    Next m
                    strD = strD.Substring(0, strD.Length - 1) & "})" & Chr(13) & Chr(10)
                End If
            End If
            Return strD
        End Get
    End Property




    Public Property IsSelected As Boolean


    Public Property SelectedMember As New List(Of Object)

    Sub New()

    End Sub


    Public Sub New(UniName As String)
        Me.SelectionMode = enSelectionMode.AllSelected
        Me.Sort = enSort.none

        Me.UniName = UniName
        Me.TableName = UniName.Substring(0, UniName.IndexOf("[") - 1).Replace("'", "")
        Me.FieldName = UniName.Substring(UniName.IndexOf("[")).Replace("[", "").Replace("]", "")

        Me.IsSelected = False
        Me.GUID = System.Guid.NewGuid.ToString

    End Sub

    Public Sub New(Query As clsQuery, UniName As String, TableName As String, FieldName As String)




        Me.Query = Query
        Me.UniName = UniName
        Me.TableName = TableName
        Me.FieldName = FieldName

        Me.IsSelected = False
        Me.GUID = System.Guid.NewGuid.ToString

    End Sub





End Class