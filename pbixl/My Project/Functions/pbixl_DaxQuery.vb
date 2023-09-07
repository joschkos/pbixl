Imports System.Threading
Imports ExcelDna.Integration
Imports Excel = Microsoft.Office.Interop.Excel

Public Module pbixl_DaxQuery


    <ExcelFunction(Description:="pbixl", Category:="pbixl", IsMacroType:=False, IsVolatile:=True, IsHidden:=False, Name:="pbixl.DAX")>
    Public Function pbixl_Query(ConnName As Object, Statement As Object, Hint As Object) As Object

        Return "Hello My Friend" & Now.ToLongTimeString


    End Function



End Module


