using Syncfusion.XlsIO;
using System;
using System.Data;
using System.Data.SqlClient;
using System.IO;
using System.Reflection;
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
                Assembly assembly = Assembly.GetCallingAssembly();
                const string RfmAnalysis_File_Path = "CompanyReports.Scripts.Queries.RfmAnalysis.sql";
                Stream resourceStream = assembly.GetManifestResourceStream(RfmAnalysis_File_Path);
                using (var reader = new StreamReader(resourceStream))
                {
                    var sqlScript = reader.ReadToEnd();
                    using (var command = new SqlCommand(sqlScript))
                    {
                        using (var dataAdapter = new SqlDataAdapter())
                        {
                            command.Connection = connection;
                            command.CommandType = CommandType.Text;
                            command.Parameters.Add("@Country", SqlDbType.NChar).Value = country;
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
                Assembly assembly = Assembly.GetCallingAssembly();
                const string ParetoRuleAnalysis_File_Path = "CompanyReports.Scripts.Queries.ParetoRuleAnalysis.sql";
                Stream resourceStream = assembly.GetManifestResourceStream(ParetoRuleAnalysis_File_Path);
                using (var reader = new StreamReader(resourceStream))
                {
                    var sqlScript = reader.ReadToEnd();
                    using (var command = new SqlCommand(sqlScript))
                    {
                        using (var dataAdapter = new SqlDataAdapter())
                        {
                            command.Connection = connection;
                            command.CommandType = CommandType.Text;                            
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
}