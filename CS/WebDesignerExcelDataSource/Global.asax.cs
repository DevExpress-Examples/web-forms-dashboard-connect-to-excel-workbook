using System;
using System.Collections.Generic;
using System.Web.Hosting;
using DevExpress.DashboardCommon;
using DevExpress.DashboardWeb;
using DevExpress.DataAccess.Excel;

namespace WebDesignerExcelDataSource
{
    public class Global : System.Web.HttpApplication
    {
        void Application_Start(object sender, EventArgs e)
        {            
            DashboardFileStorage dashboardFileStorage = new DashboardFileStorage("~/App_Data/Dashboards");
            DashboardService.SetDashboardStorage(dashboardFileStorage);

            // Creates an Excel data source and selects the specific cell range from the SalesPerson worksheet.
            DashboardExcelDataSource excelDataSource = new DashboardExcelDataSource("Excel Data Source");
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
                }
                else {
                    field.Selected = false;
                    excelDataSource.Schema.Add(field);
                }
            }

            // Adds the created data source to the data source storage.
            DataSourceInMemoryStorage dataSourceStorage = new DataSourceInMemoryStorage();
            dataSourceStorage.RegisterDataSource("excelDataSource", excelDataSource.SaveToXml());
            DashboardService.SetDataSourceStorage(dataSourceStorage);

            DashboardService.DataApi.ConfigureDataConnection += 
                new ServiceConfigureDataConnectionEventHandler(DataApi_ConfigureDataConnection);
        }

        void DataApi_ConfigureDataConnection(object sender, ServiceConfigureDataConnectionEventArgs e)
        {
            if (e.DataSourceName == "Excel Data Source")
	        {
                ((ExcelDataSourceConnectionParameters)e.ConnectionParameters).FileName = 
                    HostingEnvironment.MapPath(@"~/App_Data/ExcelDataSource.xlsx");
	        } 
        }

        protected void Session_Start(object sender, EventArgs e)
        {

        }

        protected void Application_BeginRequest(object sender, EventArgs e)
        {

        }

        protected void Application_AuthenticateRequest(object sender, EventArgs e)
        {

        }

        protected void Application_Error(object sender, EventArgs e)
        {

        }

        protected void Session_End(object sender, EventArgs e)
        {

        }

        protected void Application_End(object sender, EventArgs e)
        {

        }
    }
}