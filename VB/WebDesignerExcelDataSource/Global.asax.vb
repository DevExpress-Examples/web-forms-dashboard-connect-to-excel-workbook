Imports System
Imports System.Collections.Generic
Imports System.Web.Hosting
Imports DevExpress.DashboardCommon
Imports DevExpress.DashboardWeb
Imports DevExpress.DataAccess.Excel

Namespace WebDesignerExcelDataSource
    Public Class [Global]
        Inherits System.Web.HttpApplication

        Private Sub Application_Start(ByVal sender As Object, ByVal e As EventArgs)
            Dim dashboardFileStorage As New DashboardFileStorage("~/App_Data/Dashboards")
            DashboardService.SetDashboardStorage(dashboardFileStorage)

            ' Creates an Excel data source and selects the specific cell range from the SalesPerson worksheet.
            Dim excelDataSource As New DashboardExcelDataSource("Excel Data Source")
            excelDataSource.FileName = HostingEnvironment.MapPath("~/App_Data/ExcelDataSource.xlsx")
            Dim worksheetSettings As New ExcelWorksheetSettings("SalesPerson", "A1:L2000")
            excelDataSource.SourceOptions = New ExcelSourceOptions(worksheetSettings)

            ' Specifies the fields that will be available for the created data source.
            Dim schemaProvider As IExcelSchemaProvider = _
                TryCast(excelDataSource.GetService(GetType(IExcelSchemaProvider)), IExcelSchemaProvider)
            Dim availableFields() As FieldInfo = schemaProvider.GetSchema(excelDataSource.FileName, _
                                                                          Nothing, _
                                                                          ExcelDocumentFormat.Xlsx, _
                                                                          excelDataSource.SourceOptions, _
                                                                          System.Threading.CancellationToken.None)
            Dim fieldsToSelect As New List(Of String)() From {"CategoryName", "ProductName", _
                                                              "Country", "Quantity", "Extended Price"}
            For Each field As FieldInfo In availableFields
                If fieldsToSelect.Contains(field.Name) Then
                    excelDataSource.Schema.Add(field)
                Else
                    field.Selected = False
                    excelDataSource.Schema.Add(field)
                End If
            Next field

            ' Adds the created data source to the data source storage.
            Dim dataSourceStorage As New DataSourceInMemoryStorage()
            dataSourceStorage.RegisterDataSource("excelDataSource", excelDataSource.SaveToXml())
            DashboardService.SetDataSourceStorage(dataSourceStorage)

            AddHandler DashboardService.DataApi.ConfigureDataConnection, AddressOf DataApi_ConfigureDataConnection
        End Sub

        Private Sub DataApi_ConfigureDataConnection(ByVal sender As Object, _
                                                    ByVal e As ServiceConfigureDataConnectionEventArgs)
            If e.DataSourceName = "Excel Data Source" Then
                CType(e.ConnectionParameters, ExcelDataSourceConnectionParameters).FileName = _
                    HostingEnvironment.MapPath("~/App_Data/ExcelDataSource.xlsx")
            End If
        End Sub

        Protected Sub Session_Start(ByVal sender As Object, ByVal e As EventArgs)

        End Sub

        Protected Sub Application_BeginRequest(ByVal sender As Object, ByVal e As EventArgs)

        End Sub

        Protected Sub Application_AuthenticateRequest(ByVal sender As Object, ByVal e As EventArgs)

        End Sub

        Protected Sub Application_Error(ByVal sender As Object, ByVal e As EventArgs)

        End Sub

        Protected Sub Session_End(ByVal sender As Object, ByVal e As EventArgs)

        End Sub

        Protected Sub Application_End(ByVal sender As Object, ByVal e As EventArgs)

        End Sub
    End Class
End Namespace