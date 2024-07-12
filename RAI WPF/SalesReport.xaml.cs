using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows;
using System.Windows.Controls;
using System.Windows.Data;
using System.Windows.Documents;
using System.Windows.Input;
using System.Windows.Media;
using System.Windows.Media.Imaging;
using System.Windows.Shapes;
using System.Data;
using System.Data.SqlClient;
using System.Collections.ObjectModel;
using OfficeOpenXml;
using OfficeOpenXml.Style;
using Microsoft.WindowsAPICodePack.Dialogs;

namespace RAI_WPF
{

    /// <summary>
    /// Interaction logic for UserWindow.xaml
    /// </summary>
    public partial class SalesReport : Window
    {

        public SalesReport()
        {
            InitializeComponent();
        }
        private void ReportButton_Click(object sender, RoutedEventArgs e)
        {
            ErrorLabel.Content = "";
            try
            {

                DateTime fromDate = new DateTime();
                DateTime toDate = new DateTime();

                if (!dtFromDate.SelectedDate.HasValue)
                {
                    ErrorLabel.Content = "Please enter a valid date";
                    return;
                }
                else
                    fromDate = dtFromDate.SelectedDate.Value;
                if (!dtToDate.SelectedDate.HasValue)
                {
                    ErrorLabel.Content = "Please enter a valid date";
                    return;
                }
                else
                    toDate = dtToDate.SelectedDate.Value;

                DataSet report = RunsStoredProc.RunStoredProc("TableFunding_SalesReport", "FromDate", fromDate.ToString(), "ToDate", toDate.ToString(), "", "", "", "", "", "", "", "", "", 0);
                ExcelPackage.LicenseContext = LicenseContext.NonCommercial;

                //tableFunding.Delete();
                using (ExcelPackage pck = new ExcelPackage())
                {
                    ExcelWorksheet ws = pck.Workbook.Worksheets.Add("Sales Report");

                    ws.Cells["A4"].LoadFromDataTable(report.Tables[0], true);
                    ws.Cells["A2"].Value = "Sales Report";
                    ws.Cells["A3"].Formula = "=Today()";

                    ws.Cells["A3"].Style.Numberformat.Format = "yyyy-mm-dd";

                    ws.Cells["F:G"].Style.Numberformat.Format = "#,##0.00";

                    ws.Cells["H:H"].Style.Numberformat.Format = "yyyy-mm-dd";

                    ws.Cells["A2"].Style.Font.Size = 18;
                    ws.Cells["A3"].Style.Font.Size = 18;

                    ws.Cells[4, 1, 4, 38].Style.Font.Size = 14;


                    ws.Cells[4, 1, 4, 38].Style.Fill.PatternType = ExcelFillStyle.Solid;

                    ws.Cells[4, 1, 4, 38].Style.Fill.BackgroundColor.SetColor(System.Drawing.Color.DarkGray);

                    ws.Cells[4, 1, 4, 38].Style.Font.Color.SetColor(System.Drawing.Color.White);


                    string client = "";
                    string investor = "";
                    int clientStartRow = 0;
                    int clientInvestorStartRow = 0;
                    string rowType = "";
                    //double totalOpenMortgageBalance = 0;
                    //double totalOpenAdvanceBalance = 0;
                    int dataStartRow = 5;
                    int dataEndRow = 0;
                    for (int row = dataStartRow; row <= 1000; row++)
                    {
                        if (ws.Cells[row, 1].Value == null)
                        {
                            break;
                        }

                        if (row == dataStartRow)
                        {
                            client = ws.Cells[row, 1].Text;
                            investor = ws.Cells[row, 3].Text;
                            clientStartRow = row;
                            clientInvestorStartRow = row;
                        }

                        rowType = ws.Cells[row, 14].Text;
                        if (rowType == "Grand Total")
                        {
                            ws.Cells[row, 4].Value = "";
                            ws.Cells[row, 6].Formula = "=SUBTOTAL(9,F" + dataStartRow + ":F" + (row - 1) + ")";
                            ws.Cells[row, 7].Formula = "=SUBTOTAL(9,G" + dataStartRow + ":G" + (row - 1) + ")";

                            ws.Cells[row, 8].Value = "";
                            ws.Cells[row, 1, row, 11].Style.Font.Bold = true;
                            ws.Cells[row, 1, row, 11].Style.Font.Size = 14;

                            ws.Cells[row, 1, row, 11].Style.Fill.PatternType = ExcelFillStyle.Solid;

                            ws.Cells[row, 1, row, 11].Style.Fill.BackgroundColor.SetColor(System.Drawing.Color.DarkGray);

                            ws.Cells[row, 1, row, 11].Style.Font.Color.SetColor(System.Drawing.Color.White);

                        }

                        if (rowType == "Client Total")
                        {
                            ws.Cells[row, 4].Value = "";
                            ws.Cells[row, 6].Formula = "=SUBTOTAL(9,F" + clientStartRow + ":F" + (row - 1) + ")";
                            ws.Cells[row, 7].Formula = "=SUBTOTAL(9,G" + clientStartRow + ":G" + (row - 1) + ")";

                            ws.Cells[row, 8].Value = "";
                            ws.Cells[row, 9].Value = "";
                            ws.Cells[row, 10].Value = "";
                            ws.Cells[row, 11].Value = "";
                            ws.Cells[row, 1, row, 11].Style.Font.Bold = true;
                            ws.Cells[row, 1, row, 11].Style.Font.Size = 12;

                            ws.Cells[row, 1, row, 11].Style.Fill.PatternType = ExcelFillStyle.Solid;

                            ws.Cells[row, 1, row, 11].Style.Fill.BackgroundColor.SetColor(System.Drawing.Color.LightGray);
                        }

                        if (rowType == "Client Investor Total")
                        {
                            ws.Cells[row, 4].Value = "";
                            ws.Cells[row, 6].Formula = "=SUBTOTAL(9,F" + clientInvestorStartRow + ":F" + (row - 1) + ")";
                            ws.Cells[row, 7].Formula = "=SUBTOTAL(9,G" + clientInvestorStartRow + ":G" + (row - 1) + ")";

                            ws.Cells[row, 8].Value = "";
                            ws.Cells[row, 9].Value = "";
                            ws.Cells[row, 10].Value = "";
                            ws.Cells[row, 11].Value = "";
                            ws.Cells[row, 1, row, 11].Style.Font.Bold = true;
                            ws.Cells[row, 1, row, 11].Style.Font.Size = 13;

                            ws.Cells[row, 1, row, 11].Style.Fill.PatternType = ExcelFillStyle.Solid;

                            ws.Cells[row, 1, row, 11].Style.Fill.BackgroundColor.SetColor(System.Drawing.Color.Gray);
                        }
                        if (rowType == "Detail")
                        {


                        }
                        if (ws.Cells[row, 1].Text != client && rowType == "Detail")
                        {
                            investor = "";
                            client = ws.Cells[row, 1].Text;
                            clientStartRow = row;
                        }
                        if (ws.Cells[row, 3].Text != investor && rowType == "Detail")
                        {
                            investor = ws.Cells[row, 3].Text;
                            clientInvestorStartRow = row;
                        }

                        dataEndRow = row;
                    }

                    ExcelRange dataRange = ws.Cells[dataStartRow - 1, 1, dataEndRow, 10];
                    dataRange.AutoFilter = true;

                    ws.Cells["A:P"].AutoFitColumns();
                    ws.Column(6).Width = 20;
                    ws.Column(7).Width = 20;

                    CommonOpenFileDialog dialog = new CommonOpenFileDialog();
                    dialog.InitialDirectory = "C:\\Users";
                    dialog.IsFolderPicker = true;

                    string fileName;

                    fileName = "";

                    if (dialog.ShowDialog() == CommonFileDialogResult.Ok)
                    {
                        fileName = dialog.FileName + "\\" + "SalesReport.xlsx";

                    }



                    //System.IO.FileInfo tableFunding = new System.IO.FileInfo(MyVariables.excelReportFilePath + "SalesReport.xlsx");

                    System.IO.FileInfo tableFunding = new System.IO.FileInfo(fileName);


                    tableFunding.Delete();

                    pck.SaveAs(tableFunding);
                    System.Diagnostics.Process.Start(fileName);
                }
            }
            catch (Exception ex)
            {
                ErrorLabel.Content = ex.Message;
                ErrorLabel.Foreground = new SolidColorBrush(Colors.Red);
            }
        }
           
    }
}
