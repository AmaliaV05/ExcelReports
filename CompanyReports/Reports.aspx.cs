using Syncfusion.XlsIO;
using System;
using System.Data;
using System.Data.SqlClient;
using System.Web.UI;

namespace CompanyReports
{
    public partial class Reports : Page
    {
        protected void Page_Load(object sender, EventArgs e)
        {

        }

        protected void OnButtonClickedGetRfmAnalysisData(object sender, EventArgs args)
        {
            var countryFilter = DisplayCountry.SelectedValue;
            var todayDate = DateTime.Now.ToString("dd/MM/yyyy");
            using (ExcelEngine excelEngine = new ExcelEngine())
            {
                IApplication application = excelEngine.Excel;
                application.DefaultVersion = ExcelVersion.Excel2016;
                IWorkbook workbook = application.Workbooks.Create(1);
                IWorksheet worksheet = workbook.Worksheets[0];
                SetWorksheetStyle(workbook, worksheet);
                SetTableBordersRfmAnalysis(worksheet);
                SetWorksheetHeaderRfmAnalysis(worksheet, todayDate, countryFilter);
                SetWorksheetBodyRfmAnalysis(worksheet, countryFilter);
                workbook.SaveAs("RFM Analysis - " + todayDate + ".xlsx", Response, ExcelDownloadType.Open, ExcelHttpContentType.Excel2016);
                workbook.Close();
                excelEngine.Dispose();
            }
        }

        protected void OnButtonClickedGetProductPareto(object sender, EventArgs args)
        {
            var todayDate = DateTime.Now.ToString("dd/MM/yyyy");
            using (ExcelEngine excelEngine = new ExcelEngine())
            {
                IApplication application = excelEngine.Excel;
                application.DefaultVersion = ExcelVersion.Excel2016;
                IWorkbook workbook = application.Workbooks.Create(1);
                IWorksheet worksheet = workbook.Worksheets[0];
                SetWorksheetStyle(workbook, worksheet);
                SetTableBordersProductPareto(worksheet);
                SetWorksheetHeaderProductPareto(worksheet, todayDate);
                SetWorksheetBodyProductPareto(worksheet);
                workbook.SaveAs("Products - " + todayDate + ".xlsx", Response, ExcelDownloadType.Open, ExcelHttpContentType.Excel2016);
                workbook.Close();
                excelEngine.Dispose();
            }
        }

        private void SetWorksheetStyle(IWorkbook workbook, IWorksheet worksheet)
        {
            IStyle titleStyle = workbook.Styles.Add("TitleStyle");
            titleStyle.Font.Bold = true;
            titleStyle.HorizontalAlignment = ExcelHAlign.HAlignCenter;
            titleStyle.VerticalAlignment = ExcelVAlign.VAlignCenter;
            worksheet.Range["A3"].CellStyle = titleStyle;

            IStyle subtitleStyle = workbook.Styles.Add("SubtitleStyle");
            subtitleStyle.Font.Italic = true;
            subtitleStyle.HorizontalAlignment = ExcelHAlign.HAlignCenter;
            subtitleStyle.VerticalAlignment = ExcelVAlign.VAlignCenter;
            worksheet.Range["C4"].CellStyle = subtitleStyle;
        }

        private void SetTableBordersRfmAnalysis(IWorksheet worksheet)
        {
            worksheet.Range["A7:G7"].BorderAround(ExcelLineStyle.Medium);
        }

        private void SetWorksheetHeaderRfmAnalysis(IWorksheet worksheet, string date, string country)
        {
            worksheet.Range["A3:D3"].Merge();
            worksheet.Range["A3"].Text = "RFM Analysis " + date;
            worksheet.Range["C4"].Text = country;
        }

        private void SetWorksheetBodyRfmAnalysis(IWorksheet worksheet, string country)
        {
            var data = GetRFMAnalysisData(country);
            worksheet.ImportDataTable(data, true, 7, 1);
            worksheet.UsedRange.AutofitColumns();
        }

        private DataTable GetRFMAnalysisData(string country)
        {
            string constr = @"Data Source=.\SQLEXPRESS;Initial Catalog='AdventureWorks2019';Integrated Security=True";
            using (var connection = new SqlConnection(constr))
            {
                var query = "with Dataset as (" +
               "select CustomerID, SalesOrderID, OrderDate, TotalDue " +
               "from Sales.SalesOrderHeader " +
               "inner join Sales.SalesTerritory on SalesTerritory.TerritoryID = SalesOrderHeader.TerritoryID " +
               "where SalesTerritory.Name = '" + country + "' and SalesOrderHeader.Status = 5" +
               ")," +
               "Order_Summary as (" +
               "select CustomerID, SalesOrderID, OrderDate, sum(TotalDue) as Total_Sales " +
               "from Dataset " +
               "group by CustomerID, SalesOrderID, OrderDate" +
               ") " +
               "select t1.CustomerID, " +
               "datediff(day, (select max(OrderDate) from Order_Summary where CustomerID = t1.CustomerID), (select max(OrderDate) from Order_Summary)) as Recency, " +
               "count(t1.SalesOrderID) as Frequency, " +
               "sum(t1.Total_Sales) as Monetary, " +
               "ntile(10) over(order by datediff(day, (select max(OrderDate) from Order_Summary where CustomerID = t1.CustomerID), (select max(OrderDate) from Order_Summary)) desc) as R, " +
               "ntile(10) over(order by count(t1.SalesOrderID) asc) as F, " +
               "ntile(10) over(order by sum(t1.Total_Sales) asc) as M " +
               "from Order_Summary t1 " +
               "group by t1.CustomerID " +
               "order by 1, 3 desc;";
                using (var command = new SqlCommand(query))
                {
                    using (var dataAdapter = new SqlDataAdapter())
                    {
                        command.Connection = connection;
                        dataAdapter.SelectCommand = command;
                        using (var dt = new DataTable())
                        {
                            dataAdapter.Fill(dt);
                            return dt;
                        }
                    }
                }
            }
        }

        private void SetTableBordersProductPareto(IWorksheet worksheet)
        {
            worksheet.Range["A7:D7"].BorderAround(ExcelLineStyle.Medium);
        }

        private void SetWorksheetHeaderProductPareto(IWorksheet worksheet, string date)
        {
            worksheet.Range["A3:D3"].Merge();
            worksheet.Range["A3"].Text = "Products which bring 80% of sales " + date;
        }

        private void SetWorksheetBodyProductPareto(IWorksheet worksheet)
        {
            var data = GetProductParetoRuleData();
            worksheet.ImportDataTable(data, true, 7, 1);
            worksheet.UsedRange.AutofitColumns();
        }

        private DataTable GetProductParetoRuleData()
        {
            string constr = @"Data Source=.\SQLEXPRESS;Initial Catalog='AdventureWorks2019';Integrated Security=True";
            using (var connection = new SqlConnection(constr))
            {
                var query = "with product_wise_sales as (" +
                    "select Product.ProductNumber, sum(SalesOrderDetail.LineTotal) as product_sales " +
                    "from Sales.SalesOrderDetail " +
                    "inner join Production.Product on Product.ProductID = SalesOrderDetail.ProductID " +
                    "inner join Sales.SalesOrderHeader on SalesOrderDetail.SalesOrderID = SalesOrderHeader.SalesOrderID " +
                    "where SalesOrderHeader.Status = 5 " +
                    "group by Product.ProductNumber " +
                    ")," +
                    "calc_sales as (" +
                    "select product_wise_sales.ProductNumber, product_sales, " +
                    "sum(product_sales) over(order by product_sales desc rows between unbounded preceding and 0 preceding) as running_sales, " +
                    "0.8 * sum(product_sales) over() as total_sales " +
                    "from product_wise_sales" +
                    ") " +
                    "select* from calc_sales where running_sales <= total_sales";
                using (var command = new SqlCommand(query))
                {
                    using (var dataAdapter = new SqlDataAdapter())
                    {
                        command.Connection = connection;
                        dataAdapter.SelectCommand = command;
                        using (var dt = new DataTable())
                        {
                            dataAdapter.Fill(dt);
                            return dt;
                        }
                    }
                }
            }
        }
    }
}