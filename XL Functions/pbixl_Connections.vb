Imports ExcelDna.Integration
Imports Excel = Microsoft.Office.Interop.Excel


Public Module pbixl_Connections

    <ExcelFunction(Description:="pbixl", Category:="pbixl", IsMacroType:=True, IsVolatile:=False, IsHidden:=False, Name:="pbixl.Connections")>
    Public Function pbixl_Connections() As Object

        If MyAddIn.DynamicArray = False Then
            Return "Dynamic Arrays not supported by this version of Excel."
        End If

        If ExcelDnaUtil.IsInFunctionWizard = True Then
            Return ""
        End If

        Dim wb As Excel.Workbook = MyAddIn.GetWorkbook(XlCall.Excel(XlCall.xlfGetWorkbook, 16))
        MyAddIn.Connections.Refresh(wb)

        Dim objRes(0, 0)
        objRes(0, 0) = "Name"

        If MyAddIn.Connections.Conns.Count = 0 Then
            Return objRes
        Else
            ReDim objRes(MyAddIn.Connections.Conns.Count - 1, 0)
            For i As Integer = 0 To MyAddIn.Connections.Conns.Count - 1
                If MyAddIn.Connections.Conns.Item(i).ConnType = clsConnections.clsConnection.enConnType.PBIDesktop Then
                    objRes(i, 0) = MyAddIn.Connections.Conns.Item(i).ConnAlias
                Else
                    objRes(i, 0) = MyAddIn.Connections.Conns.Item(i).Name
                End If

            Next i
            Return objRes
        End If

    End Function

End Module
