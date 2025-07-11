Imports ExcelDna.Integration
Imports System.Runtime.InteropServices
Imports ExcelDna.Integration.CustomUI
Imports Excel = Microsoft.Office.Interop.Excel
Imports System.Runtime.CompilerServices

Friend Class MyAddin

    Implements IExcelAddIn

    Public Shared QryMgr As New clsQryMgr

    Public Shared WithEvents App As Excel.Application = ExcelDna.Integration.ExcelDnaUtil.Application


    Public Shared Sub Application_SelectionChange(Sheet As Excel.Worksheet, Target As Excel.Range) Handles App.SheetSelectionChange
    End Sub


    Public Shared Sub Application_WorkbookBeforeSave() Handles App.WorkbookBeforeSave

        For Each wb As Excel.Workbook In App.Workbooks
            For Each ws As Excel.Worksheet In wb.Worksheets
                For Each lo As Excel.ListObject In ws.ListObjects

                    Try
                        Dim strGUID As String = ""
                        Dim strConn As String = lo.QueryTable.WorkbookConnection.OLEDBConnection.Connection
                        If strConn.ToLower.StartsWith("oledb;") Then
                            strConn = strConn.Substring(6)
                        End If

                        Dim bx As New OleDb.OleDbConnectionStringBuilder(strConn)
                        strGUID = bx.Item("Application Name").ToString.ToLower.Replace("pbixl", "").Substring(0, 8)


                        Dim q As clsQuery = MyAddin.QryMgr.GetQueryByGUIDRework(wb, strGUID)

                        If q Is Nothing Then
                            For Each _q In MyAddin.QryMgr.GlobalQueries
                                If _q.GUID.StartsWith(strGUID) Then
                                    MyAddin.QryMgr.SaveQuery(wb, _q)
                                    Exit For
                                End If
                            Next _q
                        End If

                    Catch ex As Exception

                    End Try

                Next lo
            Next ws
        Next wb



    End Sub


    Public Shared Sub Application_WorkbookOpen(Wb As Excel.Workbook) Handles App.WorkbookOpen
        Try
            MyAddin.QryMgr.SetPorts(Wb)
        Catch ex As Exception
        End Try

        Try
            For Each q In MyAddin.QryMgr.WorkbookQueries(Wb)
                MyAddin.QryMgr.AddGlobalQuery(q)
            Next q
        Catch ex As Exception
        End Try

    End Sub


    Public Sub AutoOpen() Implements IExcelAddIn.AutoOpen
        'Throw New NotImplementedException()
    End Sub

    Public Sub AutoClose() Implements IExcelAddIn.AutoClose
        'Throw New NotImplementedException()
    End Sub

    Private Shared intBitness As Integer = 0
    Public Shared ReadOnly Property Bitness As Integer
        Get
            If intBitness = 0 Then

                Dim vHinstance As Object
                Try
                    vHinstance = MyAddin.App.Hinstance
                    intBitness = 32
                Catch ex As Exception
                    intBitness = 64
                End Try
            End If
            Return intBitness
        End Get
    End Property



End Class

<ComVisible(True)>
Public Class EFRRibbon
    Inherits ExcelRibbon

    Public Overrides Sub OnBeginShutdown(ByRef custom As Array)
        MyBase.OnBeginShutdown(custom)
    End Sub

    Public Overrides Function GetCustomUI(uiName As String) As String

        If MyAddin.App.Version < "16.0" Then
            'MsgBox("This AddIn does not support Excel version lower 16.0")
            'Return ""
        End If

        Dim strXML As String = ""
        strXML += "<customUI xmlns='http://schemas.microsoft.com/office/2006/01/customui' loadImage='GetImage'>"
        strXML += "<ribbon>"
        strXML += "<tabs>"
        strXML += "<tab id='pbixl' label='pbixl'>"
        strXML += "<group id='pbixl_Grp1'>"
        'strXML += "<button id='btnTable' label='Table' image='pbiTable' size='large' onAction='OnPBITable'/>"
        'strXML += "<button id='btnPivot' label='Pivot' image='pbiPivot' size='large' onAction='OnPBIPivot'/>"
        strXML += "<button id='btnConn' label='pbixl' image='pbiConnections' size='large' onAction='OnPBIConn'/>"
        strXML += "</group>"
        strXML += "</tab>"
        strXML += "</tabs>"
        strXML += "</ribbon>"
        strXML += "</customUI>"
        Return strXML

    End Function

    Public Function GetImage(ByVal imageName As String) As System.Drawing.Image
        If imageName = "pbixl" Then
            Return My.Resources.pbixl.ToBitmap
        ElseIf imageName = "pbiXLTable" Then
            Return My.Resources.pbiXLTable.ToBitmap
        ElseIf imageName = "pbiTable" Then
            Return My.Resources.PbiTable.ToBitmap
        ElseIf imageName = "pbiPivot" Then
            Return My.Resources.PbiPivot.ToBitmap
        ElseIf imageName = "pbiConnections" Then
            Return My.Resources.pbiConnections.ToBitmap
        End If
        Return Nothing
    End Function

    Public Sub OnPBIConn(control As ExcelDna.Integration.CustomUI.IRibbonControl)


        If MyAddin.App.ActiveWorkbook Is Nothing Then
            Dim _wb As Excel.Workbook = MyAddin.App.Workbooks.Add
            _wb.Activate()
        End If

        Try
            MyAddin.QryMgr.SetPorts(MyAddin.App.ActiveWorkbook)
        Catch ex As Exception

        End Try



        Dim ce As Excel.Range = MyAddin.App.ActiveCell
        Dim sr As Excel.Range = MyAddin.App.Selection
        Dim ws As Excel.Worksheet = sr.Worksheet
        Dim wb As Excel.Workbook = Nothing
        Dim lo As Excel.ListObject = Nothing
        Dim pt As Excel.PivotTable = Nothing


        For Each _wb As Excel.Workbook In MyAddin.App.Workbooks
            For Each _ws As Excel.Worksheet In _wb.Worksheets
                If _ws Is ws Then
                    wb = _wb
                    Exit For
                End If
            Next _ws
            If Not wb Is Nothing Then Exit For
        Next _wb

        If wb Is Nothing Then
            Exit Sub
        End If

        For Each _lo As Excel.ListObject In ws.ListObjects
            If Not MyAddin.App.Intersect(_lo.Range, ce) Is Nothing OrElse Not MyAddin.App.Intersect(_lo.Range, sr) Is Nothing Then
                lo = _lo
                Exit For
            End If
        Next _lo

        For Each _pt As Excel.PivotTable In ws.PivotTables
            If Not MyAddin.App.Intersect(_pt.TableRange2, ce) Is Nothing OrElse Not MyAddin.App.Intersect(_pt.TableRange2, sr) Is Nothing Then
                pt = _pt
                Exit For
            End If
        Next _pt


        If pt Is Nothing And lo Is Nothing Then

            Dim d As New dlgSelectConn(wb, dlgSelectConn.enumDialogMode.Table)
            If d.ShowDialog = DialogResult.OK Then

                If d.SelectedConn Is Nothing Then
                    d.Close() : d.Dispose()
                    Exit Sub
                End If

                Dim strAppName As String = ""
                If d.SelectedConn.ConnType = clsConnections.clsConnection.enConnType.PBIDesktop Then
                    If d.SelectedConn.ConnAlias Is Nothing OrElse d.SelectedConn.ConnAlias = "" Then
                        strAppName = "pbixl00000000Unnamed PBI Desktop"
                    Else
                        strAppName = "pbixl00000000" & d.SelectedConn.ConnAlias
                    End If
                Else
                    strAppName = "pbixl00000000"
                End If


                If d.DialogMode = dlgSelectConn.enumDialogMode.Pivot Then

                    Dim blnAdd As Boolean = False

                    Dim strCube As String = d.SelectedConn.BASECUBENAME
                    Dim strConnName As String = d.SelectedConn.ConnType.ToString
                    Dim wc As Excel.WorkbookConnection = Nothing

                    If d.SelectedConn.ConnType <> clsConnections.clsConnection.enConnType.PBIDesktop Then
                        strAppName = d.SelectedConn.Name
                    Else
                        strConnName = d.SelectedConn.Name
                    End If


                    Dim bx As New OleDb.OleDbConnectionStringBuilder(d.SelectedConn.ConnectionString)
                    If d.SelectedConn.ConnType = clsConnections.clsConnection.enConnType.PBIDesktop Then
                        If d.SelectedConn.ConnAlias = "" Then
                            bx.Item("Application Name") = "pbixlpivottblUnnamed PBI Desktop"
                        Else
                            bx.Item("Application Name") = "pbixlpivottbl" & d.SelectedConn.ConnAlias
                        End If
                    End If

                    Try

                        If d.SelectedConn.ConnType <> clsConnections.clsConnection.enConnType.PBIDesktop Then
                            For Each _wc As Excel.WorkbookConnection In wb.Connections
                                If _wc.Name = d.SelectedConn.Name Then
                                    wc = _wc
                                    Exit For
                                End If
                            Next _wc
                        End If

                        Dim strConn As String = ""
                        Dim strAlias As String = ""
                        If d.SelectedConn.ConnType = clsConnections.clsConnection.enConnType.PBIDesktop Then
                            For Each _wc As Excel.WorkbookConnection In wb.Connections

                                Try
                                    strConn = _wc.OLEDBConnection.Connection
                                Catch ex As Exception

                                End Try
                                If strConn <> "" AndAlso strConn.ToLower.StartsWith("oledb;") Then
                                    Dim xbx As New OleDb.OleDbConnectionStringBuilder(strConn.Substring(6))
                                    If xbx.Item("Application Name").ToString.ToLower.StartsWith("pbixlpivottbl") Then
                                        strAlias = xbx.Item("Application Name").ToString.Substring(13)
                                    End If
                                End If

                                If d.SelectedConn.ConnAlias Is Nothing OrElse d.SelectedConn.ConnAlias.ToString.Trim = "" Then
                                    If strAlias.ToLower = "unnamed pbi desktop" Then
                                        wc = _wc
                                        Exit For
                                    End If
                                Else
                                    If strAlias.ToLower = d.SelectedConn.ConnAlias.ToLower Then
                                        wc = _wc
                                        Exit For
                                    End If
                                End If




                            Next _wc
                        End If


                        If wc Is Nothing Then
                            blnAdd = True
                            If wc Is Nothing And d.SelectedConn.ConnType <> clsConnections.clsConnection.enConnType.PBIDesktop Then
                                wc = wb.Connections.Add2(strConnName, strAppName, "OLEDB;" & bx.ConnectionString, strCube, 1)
                            ElseIf wc Is Nothing And d.SelectedConn.ConnType = clsConnections.clsConnection.enConnType.PBIDesktop Then
                                If d.SelectedConn.ConnAlias Is Nothing OrElse d.SelectedConn.ConnAlias.ToString = "" Then
                                    wc = wb.Connections.Add2("PBI Desktop Unnamed", "PBI Desktop Pivot Table", "OLEDB;" & bx.ConnectionString, strCube, 1)
                                Else
                                    wc = wb.Connections.Add2("PBI Desktop " & d.SelectedConn.ConnAlias, " PBI Desktop Pivot Table", "OLEDB;" & bx.ConnectionString, strCube, 1)
                                End If
                            End If
                        End If

                        Try
                            Dim x As Object = CreateObject("ADODB.CONNECTION")
                            x.open(wb.Connections(wc.Name).OLEDBConnection.Connection.ToString.Substring(6))
                            x.close
                        Catch ex As Exception
                        End Try

                        wb.PivotCaches.Create(Excel.XlPivotTableSourceType.xlExternal, SourceData:=wb.Connections(wc.Name)).CreatePivotTable(TableDestination:=ce, TableName:=wc.Name & System.Guid.NewGuid.ToString.Substring(0, 8))

                    Catch ex As Exception

                        If Not wc Is Nothing AndAlso blnAdd = True Then
                            Try
                                wc.Delete()
                            Catch xex As Exception
                            End Try
                        End If
                        Try
                            d.Close()
                            d.Dispose()
                        Catch xex As Exception
                        End Try
                        MsgBox(ex.Message, MsgBoxStyle.Critical)

                    End Try






                Else
                    Dim strCubeName As String = d.SelectedConn.BASECUBENAME
                    Dim f As New frmMain(d.SelectedConn.ConnectionString, strCubeName, Nothing, d.SelectedConn.Name, "New..")



                    If f.ShowDialog = DialogResult.OK Then


                        If f.Query Is Nothing OrElse f.Query.QueryColumns.Count = 0 Then
                            Exit Sub
                        End If


                        Dim q As clsQuery = f.Query.Clone
                        Dim c As String = d.SelectedConn.ConnectionString

                        q.GUID = System.Guid.NewGuid.ToString

                        If d.SelectedConn.ConnType = clsConnections.clsConnection.enConnType.PBIDesktop Then

                            If d.SelectedConn.ConnAlias Is Nothing OrElse d.SelectedConn.ConnAlias.ToString.Trim = "" Then
                                strAppName = "pbixl" & q.GUID.Substring(0, 8) & "Unnamed PBI Desktop"
                            Else
                                strAppName = "pbixl" & q.GUID.Substring(0, 8) & d.SelectedConn.ConnAlias
                            End If
                        Else
                            strAppName = "pbixl" & q.GUID.Substring(0, 8) & d.SelectedConn.Name
                        End If



                        d.Dispose()
                        f.CancelQuery()
                        f.Dispose()



                        If c.ToLower.StartsWith("oledb;") Then
                            c = c.Substring(6)
                        End If


                        Dim sb As New OleDb.OleDbConnectionStringBuilder(c)
                        sb("Application Name") = strAppName



                        MyAddin.QryMgr.SaveQuery(wb, q)

                        lo = ws.ListObjects.AddEx(Excel.XlListObjectSourceType.xlSrcExternal, Source:="OLEDB;" & sb.ConnectionString, Destination:=ce)
                        lo.QueryTable.ListObject.Name = GetTableName(wb, q.QueryDefaultName)
                        lo.QueryTable.WorkbookConnection.Description = "pbixl " & strAppName.Substring(strAppName.LastIndexOf(":") + 1)
                        lo.QueryTable.CommandText = q.DAX(False)
                        lo.QueryTable.BackgroundQuery = True


                        Try
                            Dim x As Object = CreateObject("ADODB.CONNECTION")
                            x.open(sb.ConnectionString)
                            x.close
                        Catch ex As Exception
                        End Try

                        Try
                            lo.Refresh()
                        Catch ex As Exception
                            MsgBox(ex.Message, vbCritical)
                        End Try



                    Else
                        f.CancelQuery()
                        f.Dispose()
                    End If
                End If









            Else
                d.Close()
                d.Dispose()
            End If

        ElseIf Not lo Is Nothing Then

            Try
                If TryCast(lo, Excel.ListObject) Is Nothing Then
                    Exit Sub
                End If
            Catch ex As Exception
            End Try

            Try
                If TryCast(lo.QueryTable, Excel.QueryTable) Is Nothing Then
                    MsgBox("Please place cursor in an empty worksheet cell.", vbInformation)
                    Exit Sub
                End If
            Catch ex As Exception
                MsgBox("Please place cursor in an empty worksheet cell.", vbInformation)
                Exit Sub
            End Try


            If lo.QueryTable.Refreshing Then
                MsgBox("table " & lo.Name & " is refreshing", MsgBoxStyle.Critical)
                Exit Sub
            End If

            Dim strGUID As String = ""
            Dim strConn As String = lo.QueryTable.WorkbookConnection.OLEDBConnection.Connection
            If strConn.ToLower.StartsWith("oledb;") Then
                strConn = strConn.Substring(6)
            End If

            Dim bx As New OleDb.OleDbConnectionStringBuilder(strConn)
            strGUID = bx.Item("Application Name").ToString.ToLower.Replace("pbixl", "")

            If strGUID.Length <> 8 Then
                Try
                    strGUID = lo.QueryTable.WorkbookConnection.Description
                Catch ex As Exception
                    strGUID = lo.Summary
                End Try
                If strGUID = "" Then
                    strGUID = lo.Summary
                End If
            End If

            Dim q As clsQuery = MyAddin.QryMgr.GetQueryByGUIDRework(wb, strGUID)


            If q Is Nothing Then
                MsgBox("No associated query found for Table " & lo.Name & ".", vbCritical)
                Exit Sub
            End If


            Dim strConnName As String = lo.QueryTable.WorkbookConnection.Name
            Dim strQueryName As String = lo.Name

            Dim f As New frmMain(lo.QueryTable.WorkbookConnection.OLEDBConnection.Connection.ToString.Substring(6), q.CubeName, q, strConnName, strQueryName)
            If f.ShowDialog = DialogResult.OK Then

                If f.Query.QueryColumns.Count = 0 Then
                    Exit Sub
                End If



                Dim qn As clsQuery = f.Query.Clone
                f.CancelQuery()
                f.Dispose()

                qn.GUID = f.Query.GUID
                MyAddin.QryMgr.SaveQuery(wb, qn)

                Try
                    Dim x As Object = CreateObject("ADODB.CONNECTION")
                    x.open(lo.QueryTable.WorkbookConnection.OLEDBConnection.Connection.ToString.Substring(6))
                    x.close
                Catch ex As Exception
                End Try

                lo.QueryTable.CommandText = qn.DAX(False)
                Try
                    lo.QueryTable.Refresh()
                Catch ex As Exception
                    MsgBox(ex.Message, MsgBoxStyle.Critical)
                End Try

            Else
                f.CancelQuery()
                f.Dispose()
            End If



        ElseIf Not pt Is Nothing Then

            Try
                Dim strConn As String = pt.PivotCache.Connection
                If strConn.ToLower.StartsWith("oledb;") Then
                    strConn = strConn.Substring(6)
                End If

                Try
                    Dim x As Object = CreateObject("ADODB.CONNECTION")
                    x.open(strConn)
                    x.close
                Catch ex As Exception
                End Try


                Dim bx As New OleDb.OleDbConnectionStringBuilder(strConn)
                Dim strAppName = bx.Item("Application Name").ToString.ToLower.Replace("pbixl", "")


            Catch ex As Exception
                Exit Sub
            End Try


        End If



    End Sub




    Private Function GetTableName(wb As Excel.Workbook, strTableName As String) As String
        Dim strRes As String = strTableName
        Dim ctr As Integer = 0
        Do While Me.TableNameExist(wb, strRes) = True
            ctr += 1
            strRes = strTableName & ctr.ToString
        Loop
        Return strRes
    End Function

    Private Function TableNameExist(wb As Excel.Workbook, strTableName As String) As Boolean
        Dim blnRes As Boolean = False
        Dim ws As Excel.Worksheet = Nothing
        Dim lo As Excel.ListObject = Nothing
        For Each ws In wb.Worksheets
            For Each lo In ws.ListObjects
                If lo.Name.ToLower.Trim = strTableName.ToLower.Trim Then
                    Return True
                End If
            Next lo
        Next ws
        Return blnRes
    End Function

    Private Function getListObjxect() As Excel.ListObject

        Dim c As Excel.Range = MyAddin.App.ActiveCell
        Dim sr As Excel.Range = MyAddin.App.Selection
        Dim ws As Excel.Worksheet = sr.Worksheet
        Dim wb As Excel.Workbook = Nothing
        Dim lo As Excel.ListObject = Nothing
        For Each _wb As Excel.Workbook In MyAddin.App.Workbooks
            For Each _ws As Excel.Worksheet In _wb.Worksheets
                If _ws Is ws Then
                    wb = _wb
                    Exit For
                End If
            Next _ws
            If Not wb Is Nothing Then Exit For
        Next _wb

        For Each _lo As Excel.ListObject In ws.ListObjects
            If Not MyAddin.App.Intersect(_lo.Range, c) Is Nothing OrElse Not MyAddin.App.Intersect(_lo.Range, sr) Is Nothing Then
                lo = _lo
                Exit For
            End If
        Next _lo

        Return lo

    End Function




End Class