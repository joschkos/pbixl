Imports ExcelDna.Integration
Imports System.Runtime.InteropServices
Imports ExcelDna.Integration.CustomUI
Imports Excel = Microsoft.Office.Interop.Excel
Imports System.Runtime.CompilerServices

Friend Class MyAddin




    Implements IExcelAddIn

    Public Shared WithEvents App As Excel.Application = ExcelDna.Integration.ExcelDnaUtil.Application

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
        Dim strXML As String = ""
        strXML += "<customUI xmlns='http://schemas.microsoft.com/office/2006/01/customui' loadImage='GetImage'>"
        strXML += "<ribbon>"
        strXML += "<tabs>"
        strXML += "<tab id='pbixl' label='pbixl'>"
        strXML += "<group id='pbixl_Grp1'>"
        strXML += "<button id='btnNav' label='pbixl' image='pbiXLTable' size='large' onAction='OnPBI'/>"
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
        End If
        Return Nothing
    End Function


    Public Sub onPBI(control As ExcelDna.Integration.CustomUI.IRibbonControl)

        Dim ce As Excel.Range = MyAddin.App.ActiveCell
        Dim sr As Excel.Range = MyAddin.App.Selection
        Dim ws As Excel.Worksheet = sr.Worksheet
        Dim wb As Excel.Workbook = Nothing
        Dim lo As Excel.ListObject = Nothing



        Try



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
            Else
                Dim qryMgr As New clsQryMgr(wb)
                Dim lstGuid As New List(Of String)
                For Each _ws As Excel.Worksheet In wb.Worksheets
                    For Each _lo As Excel.ListObject In _ws.ListObjects
                        Try
                            lstGuid.Add(_lo.QueryTable.WorkbookConnection.Description.Substring(0, 8))
                        Catch ex As Exception
                        End Try
                    Next _lo
                Next _ws
                If lstGuid.Count > 0 Then
                    qryMgr.RemoveQueryWithWhiteList(lstGuid)
                End If
            End If

            For Each _lo As Excel.ListObject In ws.ListObjects
                If Not MyAddin.App.Intersect(_lo.Range, ce) Is Nothing OrElse Not MyAddin.App.Intersect(_lo.Range, sr) Is Nothing Then
                    lo = _lo
                    Exit For
                End If
            Next _lo

            If lo Is Nothing Then 'New

                Dim d As New dlgSelectConn(wb)
                If d.ShowDialog = DialogResult.OK Then

                    'If d.SelectedConn.ConnType = clsConnections.clsConnection.enConnType.PBIDesktop Then
                    'End If


                    Dim strCubeName As String = d.SelectedConn.BASECUBENAME
                    Dim f As New frmMain(d.SelectedConn.ConnectionString, strCubeName, Nothing, d.SelectedConn.Name, "New..")
                    If f.ShowDialog = DialogResult.OK Then

                        Dim q As clsQuery = f.Query.Clone
                        Dim c As String = d.SelectedConn.ConnectionString

                        d.Dispose()
                        f.CancelQuery()
                        f.Dispose()

                        q.GUID = System.Guid.NewGuid.ToString

                        Dim qm As New clsQryMgr(wb)
                        qm.SaveQuery(q)

                        With ws.ListObjects.Add(Excel.XlListObjectSourceType.xlSrcExternal, Source:="OLEDB;" & c, Destination:=ce).QueryTable
                            .ListObject.Name = GetTableName(wb, q.QueryDefaultName)
                            .WorkbookConnection.Description = q.GUID.Substring(0, 8)
                            .CommandText = q.DAX(False)
                            .BackgroundQuery = True
                            .Refresh()
                        End With
                    Else
                        f.CancelQuery()
                        f.Dispose()
                    End If
                Else
                    d.Close()
                    d.Dispose()
                End If
            Else 'edit

                Try
                    If TryCast(lo, Excel.ListObject) Is Nothing Then
                        Exit Sub
                    End If
                Catch ex As Exception

                End Try



                If lo.QueryTable.Refreshing Then
                    MsgBox("table " & lo.Name & " is refreshing", MsgBoxStyle.Critical)
                    Exit Sub
                End If

                Dim qm As New clsQryMgr(wb)
                Dim q As clsQuery = qm.GetQueryByGUIDRework(lo.QueryTable.WorkbookConnection.Description)

                If q Is Nothing Then
                    MsgBox("No associated query found for Table " & lo.Name & ".", vbCritical)
                    Exit Sub
                End If

                Dim strConnName As String = lo.QueryTable.WorkbookConnection.Name
                Dim strQueryName As String = lo.Name

                Dim f As New frmMain(lo.QueryTable.WorkbookConnection.OLEDBConnection.Connection.ToString.Substring(6), q.CubeName, q, strConnName, strQueryName)
                If f.ShowDialog = DialogResult.OK Then
                    Dim qn As clsQuery = f.Query.Clone
                    f.CancelQuery()
                    f.Dispose()
                    qn.GUID = System.Guid.NewGuid.ToString
                    qm.SaveQuery(qn)
                    lo.QueryTable.WorkbookConnection.Description = qn.GUID.Substring(0, 8)
                    lo.QueryTable.CommandText = qn.DAX(False)
                    lo.QueryTable.Refresh()
                Else
                    f.CancelQuery()
                    f.Dispose()
                End If


            End If

        Catch ex As Exception
            If Not lo Is Nothing Then
                MsgBox("The table is not associated to any pbixl query.", MsgBoxStyle.Critical)
            Else
                MsgBox(ex.Message, MsgBoxStyle.Critical)
            End If

            
        End Try


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