Imports System.Windows.Forms


<Runtime.InteropServices.ComVisible(True)>
Public Class ctrlCTP

    Public Property CTP As ExcelDna.Integration.CustomUI.CustomTaskPane

    Private _Workbook As Microsoft.Office.Interop.Excel.Workbook
    Friend Property Workbook As Microsoft.Office.Interop.Excel.Workbook
        Get
            If Me._Workbook Is Nothing Then
                Return Nothing
            Else
                For Each wb As Microsoft.Office.Interop.Excel.Workbook In MyAddIn.App.Workbooks
                    If wb.Name.ToLower.Trim = Me._Workbook.Name.ToLower.Trim Then
                        Return Me._Workbook
                    End If
                Next wb
            End If
            Return Nothing
        End Get
        Set(value As Microsoft.Office.Interop.Excel.Workbook)
            Me._Workbook = value
            InitControl()
        End Set
    End Property

    Public Sub New()
        InitializeComponent()
    End Sub

    Private Sub InitControl()
        Me.BorderStyle = BorderStyle.None
    End Sub






End Class
