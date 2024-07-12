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
    public partial class AgingReportWindow : Window
    {

        public AgingReportWindow()
        {
            InitializeComponent();
        }
        private void ReportButton_Click(object sender, RoutedEventArgs e)
        {
            ErrorLabel.Content = "";
            try
            {

                DateTime asofDate = new DateTime();

                if (!dtAsofDate.SelectedDate.HasValue)
                {
                    ErrorLabel.Content = "Please enter a valid date";
                    return;
                }
                else
                    asofDate = dtAsofDate.SelectedDate.Value;
                DataSet report = RunsStoredProc.RunStoredProc("TableFunding_AgingReport2", "AsofDate", asofDate.ToString(), "", "", "", "", "", "", "", "", "", "", "", 0);
                ExcelPackage.LicenseContext = LicenseContext.NonCommercial;


                using (ExcelPackage pck = new ExcelPackage())
                {
                    ExcelWorksheet ws = pck.Workbook.Worksheets.Add("Aging Report");

                    ws.Cells["A4"].LoadFromDataTable(report.Tables[0], true);
                    ws.Cells["A2"].Value = "Master Receivable Report";
                    ws.Cells["A3"].Value = asofDate.ToShortDateString();

                    ws.Cells["A3"].Style.Numberformat.Format = "yyyy-mm-dd";

                    ws.Cells["E:F"].Style.Numberformat.Format = "#,##0.00";

                    ws.Cells["G:G"].Style.Numberformat.Format = "yyyy-mm-dd";
                    ws.Cells["H:H"].Style.Numberformat.Format = "#,##0";
                    ws.Cells["I:I"].Style.Numberformat.Format = "#,##0.00";
                    ws.Cells["J:J"].Style.Numberformat.Format = "#0.00%";

                    ws.Cells["L:L"].Style.Numberformat.Format = "yyyy-mm-dd";
                    ws.Cells["A2"].Style.Font.Size = 18;
                    ws.Cells["A3"].Style.Font.Size = 18;

                    ws.Cells["A4"].Style.Font.Size = 14;
                    ws.Cells["B4"].Style.Font.Size = 14;
                    ws.Cells["C4"].Style.Font.Size = 14;
                    ws.Cells["D4"].Style.Font.Size = 14;
                    ws.Cells["E4"].Style.Font.Size = 14;
                    ws.Cells["F4"].Style.Font.Size = 14;
                    ws.Cells["G4"].Style.Font.Size = 14;
                    ws.Cells["H4"].Style.Font.Size = 14;
                    ws.Cells["I4"].Style.Font.Size = 14;
                    ws.Cells["J4"].Style.Font.Size = 14;
                    ws.Cells["K4"].Style.Font.Size = 14;
                    ws.Cells["L4"].Style.Font.Size = 14;


                    ws.Cells["A4"].Style.Fill.PatternType = ExcelFillStyle.Solid;
                    ws.Cells["B4"].Style.Fill.PatternType = ExcelFillStyle.Solid;
                    ws.Cells["C4"].Style.Fill.PatternType = ExcelFillStyle.Solid;
                    ws.Cells["D4"].Style.Fill.PatternType = ExcelFillStyle.Solid;
                    ws.Cells["E4"].Style.Fill.PatternType = ExcelFillStyle.Solid;
                    ws.Cells["F4"].Style.Fill.PatternType = ExcelFillStyle.Solid;
                    ws.Cells["G4"].Style.Fill.PatternType = ExcelFillStyle.Solid;
                    ws.Cells["H4"].Style.Fill.PatternType = ExcelFillStyle.Solid;
                    ws.Cells["I4"].Style.Fill.PatternType = ExcelFillStyle.Solid;
                    ws.Cells["J4"].Style.Fill.PatternType = ExcelFillStyle.Solid;
                    ws.Cells["K4"].Style.Fill.PatternType = ExcelFillStyle.Solid;
                    ws.Cells["L4"].Style.Fill.PatternType = ExcelFillStyle.Solid;

                    ws.Cells["A4"].Style.Fill.BackgroundColor.SetColor(System.Drawing.Color.DarkGray);
                    ws.Cells["B4"].Style.Fill.BackgroundColor.SetColor(System.Drawing.Color.DarkGray);
                    ws.Cells["C4"].Style.Fill.BackgroundColor.SetColor(System.Drawing.Color.DarkGray);
                    ws.Cells["D4"].Style.Fill.BackgroundColor.SetColor(System.Drawing.Color.DarkGray);
                    ws.Cells["E4"].Style.Fill.BackgroundColor.SetColor(System.Drawing.Color.DarkGray);
                    ws.Cells["F4"].Style.Fill.BackgroundColor.SetColor(System.Drawing.Color.DarkGray);
                    ws.Cells["G4"].Style.Fill.BackgroundColor.SetColor(System.Drawing.Color.DarkGray);
                    ws.Cells["H4"].Style.Fill.BackgroundColor.SetColor(System.Drawing.Color.DarkGray);
                    ws.Cells["I4"].Style.Fill.BackgroundColor.SetColor(System.Drawing.Color.DarkGray);
                    ws.Cells["J4"].Style.Fill.BackgroundColor.SetColor(System.Drawing.Color.DarkGray);
                    ws.Cells["K4"].Style.Fill.BackgroundColor.SetColor(System.Drawing.Color.DarkGray);
                    ws.Cells["L4"].Style.Fill.BackgroundColor.SetColor(System.Drawing.Color.DarkGray);

                    ws.Cells["A4"].Style.Font.Color.SetColor(System.Drawing.Color.White);
                    ws.Cells["B4"].Style.Font.Color.SetColor(System.Drawing.Color.White);
                    ws.Cells["C4"].Style.Font.Color.SetColor(System.Drawing.Color.White);
                    ws.Cells["D4"].Style.Font.Color.SetColor(System.Drawing.Color.White);
                    ws.Cells["E4"].Style.Font.Color.SetColor(System.Drawing.Color.White);
                    ws.Cells["F4"].Style.Font.Color.SetColor(System.Drawing.Color.White);
                    ws.Cells["G4"].Style.Font.Color.SetColor(System.Drawing.Color.White);
                    ws.Cells["H4"].Style.Font.Color.SetColor(System.Drawing.Color.White);
                    ws.Cells["I4"].Style.Font.Color.SetColor(System.Drawing.Color.White);
                    ws.Cells["J4"].Style.Font.Color.SetColor(System.Drawing.Color.White);
                    ws.Cells["K4"].Style.Font.Color.SetColor(System.Drawing.Color.White);
                    ws.Cells["L4"].Style.Font.Color.SetColor(System.Drawing.Color.White);


                    ws.Cells["M4"].Value = "";
                    ws.Cells["N4"].Value = "";
                    ws.Cells["O4"].Value = "";

                    string client = "";
                    string entity = "";
                    int clientStartRow = 0;
                    int clientEntityStartRow = 0;
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
                            entity = ws.Cells[row, 2].Text;
                            clientStartRow = row;
                            clientEntityStartRow = row;
                        }

                        rowType = ws.Cells[row, 13].Text;
                        if (rowType == "Grand Total")
                        {
                            ws.Cells[row, 5].Formula = "=SUBTOTAL(9,E" + dataStartRow + ":E" + (row - 1) + ")";
                            ws.Cells[row, 6].Formula = "=SUBTOTAL(9,F" + dataStartRow + ":F" + (row - 1) + ")";
                            ws.Cells[row, 7].Value = "";
                            ws.Cells[row, 9].Value = "";
                            ws.Cells[row, 10].Value = "";
                            ws.Cells[row, 11].Value = "";
                            ws.Cells[row, 12].Value = "";
                            ws.Cells[row, 13].Value = "";

                            ws.Cells[row, 14].Value = "";
                            ws.Cells[row, 15].Value = "";
                            ws.Cells[row, 1].Style.Font.Bold = true;
                            ws.Cells[row, 1].Style.Font.Size = 14;
                            ws.Cells[row, 5].Style.Font.Size = 14;
                            ws.Cells[row, 5].Style.Font.Bold = true;
                            ws.Cells[row, 6].Style.Font.Bold = true;
                            ws.Cells[row, 6].Style.Font.Size = 14;


                            ws.Cells[row, 1].Style.Fill.PatternType = ExcelFillStyle.Solid;
                            ws.Cells[row, 2].Style.Fill.PatternType = ExcelFillStyle.Solid;
                            ws.Cells[row, 3].Style.Fill.PatternType = ExcelFillStyle.Solid;
                            ws.Cells[row, 4].Style.Fill.PatternType = ExcelFillStyle.Solid;
                            ws.Cells[row, 5].Style.Fill.PatternType = ExcelFillStyle.Solid;
                            ws.Cells[row, 6].Style.Fill.PatternType = ExcelFillStyle.Solid;
                            ws.Cells[row, 7].Style.Fill.PatternType = ExcelFillStyle.Solid;
                            ws.Cells[row, 8].Style.Fill.PatternType = ExcelFillStyle.Solid;
                            ws.Cells[row, 9].Style.Fill.PatternType = ExcelFillStyle.Solid;
                            ws.Cells[row, 10].Style.Fill.PatternType = ExcelFillStyle.Solid;
                            ws.Cells[row, 11].Style.Fill.PatternType = ExcelFillStyle.Solid;
                            ws.Cells[row, 12].Style.Fill.PatternType = ExcelFillStyle.Solid;


                            ws.Cells[row, 1].Style.Fill.BackgroundColor.SetColor(System.Drawing.Color.DarkGray);
                            ws.Cells[row, 2].Style.Fill.BackgroundColor.SetColor(System.Drawing.Color.DarkGray);
                            ws.Cells[row, 3].Style.Fill.BackgroundColor.SetColor(System.Drawing.Color.DarkGray);
                            ws.Cells[row, 4].Style.Fill.BackgroundColor.SetColor(System.Drawing.Color.DarkGray);
                            ws.Cells[row, 5].Style.Fill.BackgroundColor.SetColor(System.Drawing.Color.DarkGray);
                            ws.Cells[row, 6].Style.Fill.BackgroundColor.SetColor(System.Drawing.Color.DarkGray);
                            ws.Cells[row, 7].Style.Fill.BackgroundColor.SetColor(System.Drawing.Color.DarkGray);
                            ws.Cells[row, 8].Style.Fill.BackgroundColor.SetColor(System.Drawing.Color.DarkGray);
                            ws.Cells[row, 9].Style.Fill.BackgroundColor.SetColor(System.Drawing.Color.DarkGray);
                            ws.Cells[row, 10].Style.Fill.BackgroundColor.SetColor(System.Drawing.Color.DarkGray);
                            ws.Cells[row, 11].Style.Fill.BackgroundColor.SetColor(System.Drawing.Color.DarkGray);
                            ws.Cells[row, 12].Style.Fill.BackgroundColor.SetColor(System.Drawing.Color.DarkGray);

                            ws.Cells[row, 1].Style.Font.Color.SetColor(System.Drawing.Color.White);
                            ws.Cells[row, 2].Style.Font.Color.SetColor(System.Drawing.Color.White);
                            ws.Cells[row, 3].Style.Font.Color.SetColor(System.Drawing.Color.White);
                            ws.Cells[row, 4].Style.Font.Color.SetColor(System.Drawing.Color.White);
                            ws.Cells[row, 5].Style.Font.Color.SetColor(System.Drawing.Color.White);
                            ws.Cells[row, 6].Style.Font.Color.SetColor(System.Drawing.Color.White);
                            ws.Cells[row, 7].Style.Font.Color.SetColor(System.Drawing.Color.White);
                            ws.Cells[row, 8].Style.Font.Color.SetColor(System.Drawing.Color.White);
                            ws.Cells[row, 9].Style.Font.Color.SetColor(System.Drawing.Color.White);
                            ws.Cells[row, 10].Style.Font.Color.SetColor(System.Drawing.Color.White);
                            ws.Cells[row, 11].Style.Font.Color.SetColor(System.Drawing.Color.White);
                            ws.Cells[row, 12].Style.Font.Color.SetColor(System.Drawing.Color.White);



                        }
                        if (rowType == "Client Total")
                        {
                            ws.Cells[row, 5].Formula = "=SUBTOTAL(9,E" + clientStartRow + ":E" + (row - 1) + ")";
                            ws.Cells[row, 6].Formula = "=SUBTOTAL(9,F" + clientStartRow + ":F" + (row - 1) + ")";
                            ws.Cells[row, 7].Value = "";
                            ws.Cells[row, 9].Value = "";
                            ws.Cells[row, 10].Value = "";
                            ws.Cells[row, 11].Value = "";
                            ws.Cells[row, 12].Value = "";
                            ws.Cells[row, 13].Value = "";
                            ws.Cells[row, 14].Value = "";
                            ws.Cells[row, 15].Value = "";
                            ws.Cells[row, 1].Style.Font.Bold = true;
                            ws.Cells[row, 1].Style.Font.Size = 13;
                            ws.Cells[row, 1].Style.Font.Bold = true;
                            ws.Cells[row, 5].Style.Font.Bold = true;
                            ws.Cells[row, 5].Style.Font.Size = 13;
                            ws.Cells[row, 6].Style.Font.Bold = true;
                            ws.Cells[row, 6].Style.Font.Size = 13;

                            ws.Cells[row, 1].Style.Fill.PatternType = ExcelFillStyle.Solid;
                            ws.Cells[row, 2].Style.Fill.PatternType = ExcelFillStyle.Solid;
                            ws.Cells[row, 3].Style.Fill.PatternType = ExcelFillStyle.Solid;
                            ws.Cells[row, 4].Style.Fill.PatternType = ExcelFillStyle.Solid;
                            ws.Cells[row, 5].Style.Fill.PatternType = ExcelFillStyle.Solid;
                            ws.Cells[row, 6].Style.Fill.PatternType = ExcelFillStyle.Solid;
                            ws.Cells[row, 7].Style.Fill.PatternType = ExcelFillStyle.Solid;
                            ws.Cells[row, 8].Style.Fill.PatternType = ExcelFillStyle.Solid;
                            ws.Cells[row, 9].Style.Fill.PatternType = ExcelFillStyle.Solid;
                            ws.Cells[row, 10].Style.Fill.PatternType = ExcelFillStyle.Solid;
                            ws.Cells[row, 11].Style.Fill.PatternType = ExcelFillStyle.Solid;
                            ws.Cells[row, 12].Style.Fill.PatternType = ExcelFillStyle.Solid;


                            ws.Cells[row, 1].Style.Fill.BackgroundColor.SetColor(System.Drawing.Color.Gray);
                            ws.Cells[row, 2].Style.Fill.BackgroundColor.SetColor(System.Drawing.Color.Gray);
                            ws.Cells[row, 3].Style.Fill.BackgroundColor.SetColor(System.Drawing.Color.Gray);
                            ws.Cells[row, 4].Style.Fill.BackgroundColor.SetColor(System.Drawing.Color.Gray);
                            ws.Cells[row, 5].Style.Fill.BackgroundColor.SetColor(System.Drawing.Color.Gray);
                            ws.Cells[row, 6].Style.Fill.BackgroundColor.SetColor(System.Drawing.Color.Gray);
                            ws.Cells[row, 7].Style.Fill.BackgroundColor.SetColor(System.Drawing.Color.Gray);
                            ws.Cells[row, 8].Style.Fill.BackgroundColor.SetColor(System.Drawing.Color.Gray);
                            ws.Cells[row, 9].Style.Fill.BackgroundColor.SetColor(System.Drawing.Color.Gray);
                            ws.Cells[row, 10].Style.Fill.BackgroundColor.SetColor(System.Drawing.Color.Gray);
                            ws.Cells[row, 11].Style.Fill.BackgroundColor.SetColor(System.Drawing.Color.Gray);
                            ws.Cells[row, 12].Style.Fill.BackgroundColor.SetColor(System.Drawing.Color.Gray);
                        }

                        if (rowType == "Client Entity Total")
                        {
                            ws.Cells[row, 5].Formula = "=SUBTOTAL(9,E" + clientEntityStartRow + ":E" + (row - 1) + ")";
                            ws.Cells[row, 6].Formula = "=SUBTOTAL(9,F" + clientEntityStartRow + ":F" + (row - 1) + ")";
                            ws.Cells[row, 7].Value = "";
                            ws.Cells[row, 9].Value = "";
                            ws.Cells[row, 10].Value = "";
                            ws.Cells[row, 11].Value = "";
                            ws.Cells[row, 12].Value = "";
                            ws.Cells[row, 13].Value = "";
                            ws.Cells[row, 14].Value = "";
                            ws.Cells[row, 15].Value = "";
                            ws.Cells[row, 1].Style.Font.Bold = true;
                            ws.Cells[row, 1].Style.Font.Size = 12;
                            ws.Cells[row, 1].Style.Font.Bold = true;
                            ws.Cells[row, 2].Style.Font.Size = 12;
                            ws.Cells[row, 2].Style.Font.Bold = true;
                            ws.Cells[row, 5].Style.Font.Bold = true;
                            ws.Cells[row, 5].Style.Font.Size = 12;
                            ws.Cells[row, 6].Style.Font.Bold = true;
                            ws.Cells[row, 6].Style.Font.Size = 12;

                            ws.Cells[row, 1].Style.Fill.PatternType = ExcelFillStyle.Solid;
                            ws.Cells[row, 2].Style.Fill.PatternType = ExcelFillStyle.Solid;
                            ws.Cells[row, 3].Style.Fill.PatternType = ExcelFillStyle.Solid;
                            ws.Cells[row, 4].Style.Fill.PatternType = ExcelFillStyle.Solid;
                            ws.Cells[row, 5].Style.Fill.PatternType = ExcelFillStyle.Solid;
                            ws.Cells[row, 6].Style.Fill.PatternType = ExcelFillStyle.Solid;
                            ws.Cells[row, 7].Style.Fill.PatternType = ExcelFillStyle.Solid;
                            ws.Cells[row, 8].Style.Fill.PatternType = ExcelFillStyle.Solid;
                            ws.Cells[row, 9].Style.Fill.PatternType = ExcelFillStyle.Solid;
                            ws.Cells[row, 10].Style.Fill.PatternType = ExcelFillStyle.Solid;
                            ws.Cells[row, 11].Style.Fill.PatternType = ExcelFillStyle.Solid;
                            ws.Cells[row, 12].Style.Fill.PatternType = ExcelFillStyle.Solid;


                            ws.Cells[row, 1].Style.Fill.BackgroundColor.SetColor(System.Drawing.Color.LightGray);
                            ws.Cells[row, 2].Style.Fill.BackgroundColor.SetColor(System.Drawing.Color.LightGray);
                            ws.Cells[row, 3].Style.Fill.BackgroundColor.SetColor(System.Drawing.Color.LightGray);
                            ws.Cells[row, 4].Style.Fill.BackgroundColor.SetColor(System.Drawing.Color.LightGray);
                            ws.Cells[row, 5].Style.Fill.BackgroundColor.SetColor(System.Drawing.Color.LightGray);
                            ws.Cells[row, 6].Style.Fill.BackgroundColor.SetColor(System.Drawing.Color.LightGray);
                            ws.Cells[row, 7].Style.Fill.BackgroundColor.SetColor(System.Drawing.Color.LightGray);
                            ws.Cells[row, 8].Style.Fill.BackgroundColor.SetColor(System.Drawing.Color.LightGray);
                            ws.Cells[row, 9].Style.Fill.BackgroundColor.SetColor(System.Drawing.Color.LightGray);
                            ws.Cells[row, 10].Style.Fill.BackgroundColor.SetColor(System.Drawing.Color.LightGray);
                            ws.Cells[row, 11].Style.Fill.BackgroundColor.SetColor(System.Drawing.Color.LightGray);
                            ws.Cells[row, 12].Style.Fill.BackgroundColor.SetColor(System.Drawing.Color.LightGray);

                        }

                        if (rowType == "Detail")
                        {
                            ws.Cells[row, 8].Formula = "=A3-G" + row.ToString();
                            ws.Cells[row, 13].Value = "";
                            ws.Cells[row, 14].Value = "";
                            ws.Cells[row, 15].Value = "";

                        }
                        if (ws.Cells[row, 1].Text != client && rowType == "Detail")
                        {
                            entity = "";
                            client = ws.Cells[row, 1].Text;
                            clientStartRow = row;
                        }
                        if (ws.Cells[row, 2].Text != entity && rowType == "Detail")
                        {
                            entity = ws.Cells[row, 2].Text;
                            clientEntityStartRow = row;
                        }

                        dataEndRow = row;
                    }

                    ExcelRange dataRange = ws.Cells[dataStartRow - 1, 1, dataEndRow, 12];
                    dataRange.AutoFilter = true;

                    ws.Cells["A:L"].AutoFitColumns();
                    ExcelPackage.LicenseContext = LicenseContext.NonCommercial;
                    CommonOpenFileDialog dialog = new CommonOpenFileDialog();
                    dialog.InitialDirectory = "C:\\Users";
                    dialog.IsFolderPicker = true;

                    string fileName;

                    fileName = "";

                    if (dialog.ShowDialog() == CommonFileDialogResult.Ok)
                    {
                        fileName = dialog.FileName + "\\" + "AgingReport.xlsx";

                    }

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
