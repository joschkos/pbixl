Imports ExcelDna.Integration
Imports System.Runtime.InteropServices
Imports ExcelDna.Integration.CustomUI
Imports Excel = Microsoft.Office.Interop.Excel
Imports System.Runtime.CompilerServices

Public Module Extensions
    <Extension()>
    Public Sub InvokeIfRequired(ByVal Control As Windows.Forms.Control, ByVal Method As [Delegate], ByVal ParamArray Parameters As Object())
        If Parameters Is Nothing OrElse
            Parameters.Length = 0 Then Parameters = Nothing
        If Control.InvokeRequired = True Then
            Control.Invoke(Method, Parameters)
        Else
            Method.DynamicInvoke(Parameters)
        End If
    End Sub
End Module


Friend Class MyAddIn

    Implements IExcelAddIn

    Public Shared objDragDrop As Object
    Public Shared blnDragDropMode As Boolean
    Public Shared strDragDropGUID As String
    Public Shared htDropTarget As Hashtable

    Public Shared DragJob As DragJob

    Public Shared Connections As New clsConnections


    Private Shared intBitness As Integer = 0
    Public Shared ReadOnly Property Bitness As Integer
        Get
            If intBitness = 0 Then

                Dim vHinstance As Object
                Try
                    vHinstance = MyAddIn.App.Hinstance
                    intBitness = 32
                Catch ex As Exception
                    intBitness = 64
                End Try
            End If
            Return intBitness
        End Get
    End Property

    Public Shared ReadOnly Property DynamicArray As Boolean
        Get
            Return ExcelDnaUtil.SupportsDynamicArrays
        End Get
    End Property

    Public Shared WithEvents App As Excel.Application = ExcelDna.Integration.ExcelDnaUtil.Application

    Private Shared _ActiveCTP As ctrlCTP
    Public Shared Property ActiveCTP As ctrlCTP
        Get
            Return MyAddIn._ActiveCTP
        End Get
        Set(value As ctrlCTP)
            MyAddIn._ActiveCTP = value
        End Set
    End Property

    Public Shared Function GetWorkbook(WorkbookName As String) As Excel.Workbook
        For Each wb As Excel.Workbook In MyAddIn.App.Workbooks
            If wb.Name.ToLower = WorkbookName.ToString.ToLower Then
                Return wb
            End If
        Next wb
        Return Nothing
    End Function

    Public Shared Sub Application_SheetChange(Sheet As Object, Target As Excel.Range) Handles App.SheetChange
    End Sub

    Public Shared Sub Application_SheetActivate(Sheet As Object) Handles App.SheetActivate
    End Sub

    Public Shared Sub Application_SheetBeforeRightClick(Sheet As Object, target As Excel.Range, ByRef cancel As Boolean) Handles App.SheetBeforeRightClick
    End Sub

    Public Shared Sub Application_SelectionChange(Sheet As Object, Target As Excel.Range) Handles App.SheetSelectionChange
    End Sub

    Private Shared Sub SetActiveQuery(Sheet As Object, Target As Excel.Range)
    End Sub

    Public Shared Sub Application_Workbook_BeforSave() Handles App.WorkbookBeforeSave
    End Sub

    Public Shared Sub Application_Workbook_BeforeClose() Handles App.WorkbookBeforeClose
    End Sub

    Public Shared Sub Application_Workbook_Open() Handles App.WorkbookOpen
    End Sub


    <DllImport("user32")>
    Public Shared Function SetFocus(ByVal hWnd As IntPtr) As IntPtr
    End Function

    <DllImport("user32")>
    Public Shared Function SetForegroundWindow(ByVal hWnd As IntPtr) As IntPtr
    End Function

    <DllImport("user32")>
    Public Shared Function SetActiveWindow(ByVal hWnd As IntPtr) As IntPtr
    End Function


    Public Sub AutoOpen() Implements ExcelDna.Integration.IExcelAddIn.AutoOpen
    End Sub

    Public Sub AutoClose() Implements ExcelDna.Integration.IExcelAddIn.AutoClose
    End Sub


End Class


<ComVisible(True)>
Public Class EFRRibbon
    Inherits ExcelRibbon

    Public Shared lstCTPs As New List(Of CustomTaskPane)

    Public Shared Sub SetCTPVisble()
        If lstCTPs.Count = 0 Then
            MyAddIn.ActiveCTP = Nothing
            Exit Sub
        End If
        For i As Integer = lstCTPs.Count - 1 To 0 Step -1
            If Not lstCTPs.Item(i) Is Nothing Then
                If lstCTPs.Item(i).Visible = True Then
                    MyAddIn.ActiveCTP = TryCast(lstCTPs.Item(i).ContentControl, ctrlCTP)
                    Exit Sub
                End If
            End If
        Next i
        MyAddIn.ActiveCTP = Nothing
    End Sub

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
        strXML += "<button id='btnNav' label='pbixl' image='pbixl' size='large' onAction='OnShowCTP'/>"
        'strXML += "<button id='btnRef' label='Refresh' image='refresh' size='large' onAction='OnRefresh'/>"
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
        ElseIf imageName = "refresh" Then
            Return My.Resources.Refresh_32.ToBitmap
        End If
        Return Nothing
    End Function

    Public Sub OnRefresh(control As ExcelDna.Integration.CustomUI.IRibbonControl)
        MsgBox("refresh connections and tables, pivot tables, functions")
    End Sub

    Public Sub OnShowCTP(control As ExcelDna.Integration.CustomUI.IRibbonControl)

        Dim oldCI As System.Globalization.CultureInfo = System.Threading.Thread.CurrentThread.CurrentCulture
        System.Threading.Thread.CurrentThread.CurrentCulture = New System.Globalization.CultureInfo("en-US")
        MyAddIn.App.Cursor = Excel.XlMousePointer.xlWait

        Try
            If MyAddIn.App.ActiveWorkbook Is Nothing Then
                Dim wb As Excel.Workbook = MyAddIn.App.Workbooks.Add()
                wb.Activate()
            End If


            For i As Integer = lstCTPs.Count - 1 To 0 Step -1
                Try
                    Dim x As Excel.Workbook = TryCast(lstCTPs.Item(i).ContentControl, ctrlCTP).Workbook
                Catch ex As Exception
                    lstCTPs.RemoveAt(i)
                End Try
            Next i

            Dim ctp As CustomTaskPane = Nothing

            Try
                For Each c As CustomTaskPane In lstCTPs
                    If TryCast(c.ContentControl, ctrlCTP).Workbook Is MyAddIn.App.ActiveWorkbook Then
                        ctp = c
                        Exit For
                    End If
                Next c
            Catch ex As Exception
                For i As Integer = lstCTPs.Count - 1 To 0 Step -1
                    If lstCTPs.Item(i) Is Nothing Then
                        lstCTPs.RemoveAt(i)
                    End If
                Next i
            End Try

            If ctp Is Nothing Then

                ctp = CustomTaskPaneFactory.CreateCustomTaskPane(Type.GetType("pbixl.ctrlCTP"), "pbixl")
                ctp.Visible = False
                TryCast(ctp.ContentControl, ctrlCTP).Workbook = MyAddIn.App.ActiveWorkbook
                TryCast(ctp.ContentControl, ctrlCTP).CTP = ctp

                AddHandler ctp.DockPositionStateChange, AddressOf ctp_DockPositionStateChange
                AddHandler ctp.VisibleStateChange, AddressOf ctp_VisibleStateChange

                lstCTPs.Add(ctp)



                ctp.DockPosition = MsoCTPDockPosition.msoCTPDockPositionLeft

                Dim intCtpWidth As Integer = 0
                Try
                    'Integer.TryParse(MyAddIn.Repository.GetQYSetting("CTP_width"), intCtpWidth)
                Catch ex As Exception
                    intCtpWidth = -1
                End Try

                If intCtpWidth > 200 And intCtpWidth < 400 Then
                    ctp.Width = intCtpWidth
                Else
                    ctp.Width = 200
                End If

                ctp.Visible = True
            Else
                If ctp.Visible = True Then
                    ctp.Visible = False
                Else
                    ctp.DockPosition = MsoCTPDockPosition.msoCTPDockPositionLeft
                    ctp.Visible = True
                End If
            End If

        Catch ex As Exception

            MsgBox("pbixl: Error while creating the Custom Task Pane:" & ex.Message, vbCritical)


        End Try


        MyAddIn.App.Cursor = Excel.XlMousePointer.xlDefault
        System.Threading.Thread.CurrentThread.CurrentCulture = oldCI
        SetCTPVisble


    End Sub

    Public Shared Sub ctp_VisibleStateChange(ctp As CustomTaskPane)
        If ctp.Visible = True Then
            Dim ctrl As ctrlCTP
            ctrl = CType(ctp.ContentControl, ctrlCTP)
        End If
        SetCTPVisble()
    End Sub

    Public Shared Sub ctp_DockPositionStateChange(ctp As CustomTaskPane)
        SetCTPVisble()
    End Sub

End Class


Public Class DragJob

    Enum JobTypeEnum
        PBIConn = 1
        PBIColumnToTable = 2
        PBIMeasureToTable = 3
        PBIDisplayFolderToTable = 4
        PBIDimensionToTable = 5
        PBIHierarchyToTable = 6
        PBIMemberToTable = 7
    End Enum


    Public Property JobType As JobTypeEnum
    Public Property objSource As Object
    Public Property Workbook As Excel.Workbook

    Public Sub New(JobType As JobTypeEnum, objSource As Object, WorkBook As Excel.Workbook)
        Me.JobType = JobType
        Me.objSource = objSource
        Me.Workbook = WorkBook
    End Sub
End Class