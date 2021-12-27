Imports System
Imports System.Collections.Generic
Imports System.Web.Hosting
Imports DevExpress.DashboardCommon
Imports DevExpress.DashboardWeb
Imports DevExpress.DataAccess.Excel

Namespace WebDesignerExcelDataSource
	Partial Public Class [Default]
		Inherits System.Web.UI.Page

		Protected Sub Page_Load(ByVal sender As Object, ByVal e As EventArgs)
			Dim dashboardFileStorage As New DashboardFileStorage("~/App_Data/Dashboards")
			ASPxDashboard1.SetDashboardStorage(dashboardFileStorage)

			' Creates an Excel data source and selects the specific cell range from the SalesPerson worksheet.
			Dim excelDataSource As New DashboardExcelDataSource("Excel Data Source")
			excelDataSource.ConnectionName = "xlsProducts"
			excelDataSource.FileName = HostingEnvironment.MapPath("~/App_Data/ExcelDataSource.xlsx")
			Dim worksheetSettings As New ExcelWorksheetSettings("SalesPerson", "A1:L2000")
			excelDataSource.SourceOptions = New ExcelSourceOptions(worksheetSettings)

			' Specifies the fields that will be available for the created data source.
			Dim schemaProvider As IExcelSchemaProvider = TryCast(excelDataSource.GetService(GetType(IExcelSchemaProvider)), IExcelSchemaProvider)
			Dim availableFields() As FieldInfo = schemaProvider.GetSchema(excelDataSource.FileName, Nothing, ExcelDocumentFormat.Xlsx, excelDataSource.SourceOptions, System.Threading.CancellationToken.None)
			Dim fieldsToSelect As New List(Of String)() From {"CategoryName", "ProductName", "Country", "Quantity", "Extended Price"}
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
			ASPxDashboard1.SetDataSourceStorage(dataSourceStorage)
		End Sub

		Protected Sub ASPxDashboard1_ConfigureDataConnection(ByVal sender As Object, ByVal e As ConfigureDataConnectionWebEventArgs)
			If e.ConnectionName = "xlsProducts" Then
				CType(e.ConnectionParameters, ExcelDataSourceConnectionParameters).FileName = HostingEnvironment.MapPath("~/App_Data/ExcelDataSource.xlsx")
			End If
		End Sub
	End Class
End Namespace