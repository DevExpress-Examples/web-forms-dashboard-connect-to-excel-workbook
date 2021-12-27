using System;
using System.Collections.Generic;
using System.Web.Hosting;
using DevExpress.DashboardCommon;
using DevExpress.DashboardWeb;
using DevExpress.DataAccess.Excel;

namespace WebDesignerExcelDataSource {
    public partial class Default : System.Web.UI.Page {
        protected void Page_Load(object sender, EventArgs e) {
            DashboardFileStorage dashboardFileStorage = new DashboardFileStorage("~/App_Data/Dashboards");
            ASPxDashboard1.SetDashboardStorage(dashboardFileStorage);

            // Creates an Excel data source and selects the specific cell range from the SalesPerson worksheet.
            DashboardExcelDataSource excelDataSource = new DashboardExcelDataSource("Excel Data Source");
            excelDataSource.ConnectionName = "xlsProducts";
            excelDataSource.FileName = HostingEnvironment.MapPath(@"~/App_Data/ExcelDataSource.xlsx");
            ExcelWorksheetSettings worksheetSettings = new ExcelWorksheetSettings("SalesPerson", "A1:L2000");
            excelDataSource.SourceOptions = new ExcelSourceOptions(worksheetSettings);

            // Specifies the fields that will be available for the created data source.
            IExcelSchemaProvider schemaProvider = excelDataSource.GetService(typeof(IExcelSchemaProvider))
                as IExcelSchemaProvider;
            FieldInfo[] availableFields = schemaProvider.GetSchema(excelDataSource.FileName, null,
                ExcelDocumentFormat.Xlsx, excelDataSource.SourceOptions, System.Threading.CancellationToken.None);
            List<string> fieldsToSelect = new List<string>() { "CategoryName", "ProductName", "Country", "Quantity",
                "Extended Price"};
            foreach (FieldInfo field in availableFields) {
                if (fieldsToSelect.Contains(field.Name)) {
                    excelDataSource.Schema.Add(field);
                } else {
                    field.Selected = false;
                    excelDataSource.Schema.Add(field);
                }
            }

            // Adds the created data source to the data source storage.
            DataSourceInMemoryStorage dataSourceStorage = new DataSourceInMemoryStorage();
            dataSourceStorage.RegisterDataSource("excelDataSource", excelDataSource.SaveToXml());
            ASPxDashboard1.SetDataSourceStorage(dataSourceStorage);
        }

        protected void ASPxDashboard1_ConfigureDataConnection(object sender, ConfigureDataConnectionWebEventArgs e) {
            if (e.ConnectionName == "xlsProducts") {
                ((ExcelDataSourceConnectionParameters)e.ConnectionParameters).FileName =
                    HostingEnvironment.MapPath(@"~/App_Data/ExcelDataSource.xlsx");
            }
        }
    }
}