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
    public partial class CollectionsReport : Window
    {

        public CollectionsReport()
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

                DataSet report = RunsStoredProc.RunStoredProc("TableFunding_CollectionReport", "FromDate", fromDate.ToString(), "ToDate", toDate.ToString(), "", "", "", "", "", "", "", "", "", 0);
                ExcelPackage.LicenseContext = LicenseContext.NonCommercial;


                using (ExcelPackage pck = new ExcelPackage())
                {
                    ExcelWorksheet ws = pck.Workbook.Worksheets.Add("Collections Report");

                    ws.Cells["A4"].LoadFromDataTable(report.Tables[0], true);
                    ws.Cells["A2"].Value = "Collections Report";
                    ws.Cells["C3"].Value = toDate.ToShortDateString();
                    ws.Cells["B3"].Value = "To";
                    ws.Cells["A3"].Value = fromDate.ToShortDateString();


                    ws.Cells["A3"].Style.Numberformat.Format = "yyyy-mm-dd";
                    ws.Cells["C3"].Style.Numberformat.Format = "yyyy-mm-dd";

                    ws.Cells["E:E"].Style.Numberformat.Format = "#,##0.00";
                    ws.Cells["F:G"].Style.Numberformat.Format = "#,##0.00";

                    ws.Cells["H:H"].Style.Numberformat.Format = "#0.00%";
                    ws.Cells["I:I"].Style.Numberformat.Format = "#,##0";
                    ws.Cells["J:J"].Style.Numberformat.Format = "yyyy-mm-dd";

                    ws.Cells["K:N"].Style.Numberformat.Format = "#,##0.00";

                    ws.Cells["A2"].Style.Font.Size = 18;
                    ws.Cells["A3"].Style.Font.Size = 18;
                    ws.Cells["B3"].Style.Font.Size = 18;
                    ws.Cells["C3"].Style.Font.Size = 18;

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
                    ws.Cells["M4"].Style.Font.Size = 14;
                    ws.Cells["N4"].Style.Font.Size = 14;


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
                    ws.Cells["M4"].Style.Fill.PatternType = ExcelFillStyle.Solid;
                    ws.Cells["N4"].Style.Fill.PatternType = ExcelFillStyle.Solid;

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
                    ws.Cells["M4"].Style.Fill.BackgroundColor.SetColor(System.Drawing.Color.DarkGray);
                    ws.Cells["N4"].Style.Fill.BackgroundColor.SetColor(System.Drawing.Color.DarkGray);

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
                    ws.Cells["M4"].Style.Font.Color.SetColor(System.Drawing.Color.White);
                    ws.Cells["N4"].Style.Font.Color.SetColor(System.Drawing.Color.White);


                    ws.Cells["O4"].Value = "";
                    ws.Cells["P4"].Value = "";
                    ws.Cells["Q4"].Value = "";

                    //string client = "";
                    string entity = "";
                    int clientStartRow = 0;
                    int clientEntityStartRow = 0;
                    int entityStartRow = 0;
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
                            //client = ws.Cells[row, 1].Text;
                            entity = ws.Cells[row, 1].Text;
                            entityStartRow = row;
                            //clientStartRow = row;
                            //clientEntityStartRow = row;
                        }

                        rowType = ws.Cells[row, 15].Text;
                        if (rowType == "Grand Total")
                        {
                            ws.Cells[row, 5].Formula = "=SUBTOTAL(9,E" + dataStartRow + ":E" + (row - 1) + ")";
                            ws.Cells[row, 6].Formula = "=SUBTOTAL(9,F" + dataStartRow + ":F" + (row - 1) + ")";
                            ws.Cells[row, 7].Formula = "=SUBTOTAL(9,G" + dataStartRow + ":G" + (row - 1) + ")";
                            ws.Cells[row, 8].Value = "";
                            ws.Cells[row, 9].Value = "";
                            ws.Cells[row, 10].Value = "";
                            ws.Cells[row, 11].Formula = "=SUBTOTAL(9,K" + dataStartRow + ":K" + (row - 1) + ")";
                            ws.Cells[row, 12].Formula = "=SUBTOTAL(9,L" + dataStartRow + ":L" + (row - 1) + ")";
                            ws.Cells[row, 13].Formula = "=SUBTOTAL(9,M" + dataStartRow + ":M" + (row - 1) + ")";
                            ws.Cells[row, 14].Formula = "=SUBTOTAL(9,N" + dataStartRow + ":N" + (row - 1) + ")";
                            ws.Cells[row, 15].Value = "";
                            ws.Cells[row, 16].Value = "";
                            ws.Cells[row, 17].Value = "";

                            ws.Cells[row, 1].Style.Font.Bold = true;
                            ws.Cells[row, 1].Style.Font.Size = 14;
                            ws.Cells[row, 4].Style.Font.Size = 14;
                            ws.Cells[row, 4].Style.Font.Bold = true;
                            ws.Cells[row, 5].Style.Font.Bold = true;
                            ws.Cells[row, 5].Style.Font.Size = 14;
                            ws.Cells[row, 6].Style.Font.Bold = true;
                            ws.Cells[row, 6].Style.Font.Size = 14;
                            ws.Cells[row, 7].Style.Font.Bold = true;
                            ws.Cells[row, 7].Style.Font.Size = 14;
                            ws.Cells[row, 11].Style.Font.Bold = true;
                            ws.Cells[row, 11].Style.Font.Size = 14;
                            ws.Cells[row, 12].Style.Font.Bold = true;
                            ws.Cells[row, 12].Style.Font.Size = 14;
                            ws.Cells[row, 13].Style.Font.Bold = true;
                            ws.Cells[row, 13].Style.Font.Size = 14;
                            ws.Cells[row, 14].Style.Font.Bold = true;
                            ws.Cells[row, 14].Style.Font.Size = 14;


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
                            ws.Cells[row, 13].Style.Fill.PatternType = ExcelFillStyle.Solid;
                            ws.Cells[row, 14].Style.Fill.PatternType = ExcelFillStyle.Solid;


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
                            ws.Cells[row, 13].Style.Fill.BackgroundColor.SetColor(System.Drawing.Color.DarkGray);
                            ws.Cells[row, 14].Style.Fill.BackgroundColor.SetColor(System.Drawing.Color.DarkGray);

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
                            ws.Cells[row, 13].Style.Font.Color.SetColor(System.Drawing.Color.White);
                            ws.Cells[row, 14].Style.Font.Color.SetColor(System.Drawing.Color.White);



                        }
                        //if (rowType == "Client Total")
                        //{
                        //    ws.Cells[row, 5].Formula = "=SUBTOTAL(9,E" + clientStartRow + ":E" + (row - 1) + ")";
                        //    ws.Cells[row, 6].Formula = "=SUBTOTAL(9,F" + clientStartRow + ":F" + (row - 1) + ")";
                        //    ws.Cells[row, 7].Value = "";
                        //    ws.Cells[row, 9].Value = "";
                        //    ws.Cells[row, 10].Value = "";
                        //    ws.Cells[row, 11].Value = "";
                        //    ws.Cells[row, 12].Value = "";
                        //    ws.Cells[row, 13].Value = ""; ws.Cells[row, 1].Style.Font.Bold = true;
                        //    ws.Cells[row, 1].Style.Font.Size = 13;
                        //    ws.Cells[row, 1].Style.Font.Bold = true;
                        //    ws.Cells[row, 5].Style.Font.Bold = true;
                        //    ws.Cells[row, 5].Style.Font.Size = 13;
                        //    ws.Cells[row, 6].Style.Font.Bold = true;
                        //    ws.Cells[row, 6].Style.Font.Size = 13;

                        //    ws.Cells[row, 1].Style.Fill.PatternType = ExcelFillStyle.Solid;
                        //    ws.Cells[row, 2].Style.Fill.PatternType = ExcelFillStyle.Solid;
                        //    ws.Cells[row, 3].Style.Fill.PatternType = ExcelFillStyle.Solid;
                        //    ws.Cells[row, 4].Style.Fill.PatternType = ExcelFillStyle.Solid;
                        //    ws.Cells[row, 5].Style.Fill.PatternType = ExcelFillStyle.Solid;
                        //    ws.Cells[row, 6].Style.Fill.PatternType = ExcelFillStyle.Solid;
                        //    ws.Cells[row, 7].Style.Fill.PatternType = ExcelFillStyle.Solid;
                        //    ws.Cells[row, 8].Style.Fill.PatternType = ExcelFillStyle.Solid;
                        //    ws.Cells[row, 9].Style.Fill.PatternType = ExcelFillStyle.Solid;
                        //    ws.Cells[row, 10].Style.Fill.PatternType = ExcelFillStyle.Solid;


                        //    ws.Cells[row, 1].Style.Fill.BackgroundColor.SetColor(System.Drawing.Color.Gray);
                        //    ws.Cells[row, 2].Style.Fill.BackgroundColor.SetColor(System.Drawing.Color.Gray);
                        //    ws.Cells[row, 3].Style.Fill.BackgroundColor.SetColor(System.Drawing.Color.Gray);
                        //    ws.Cells[row, 4].Style.Fill.BackgroundColor.SetColor(System.Drawing.Color.Gray);
                        //    ws.Cells[row, 5].Style.Fill.BackgroundColor.SetColor(System.Drawing.Color.Gray);
                        //    ws.Cells[row, 6].Style.Fill.BackgroundColor.SetColor(System.Drawing.Color.Gray);
                        //    ws.Cells[row, 7].Style.Fill.BackgroundColor.SetColor(System.Drawing.Color.Gray);
                        //    ws.Cells[row, 8].Style.Fill.BackgroundColor.SetColor(System.Drawing.Color.Gray);
                        //    ws.Cells[row, 9].Style.Fill.BackgroundColor.SetColor(System.Drawing.Color.Gray);
                        //    ws.Cells[row, 10].Style.Fill.BackgroundColor.SetColor(System.Drawing.Color.Gray);
                        //}
                        if (rowType == "Entity Total")
                        {
                            ws.Cells[row, 5].Formula = "=SUBTOTAL(9,E" + entityStartRow + ":E" + (row - 1) + ")";
                            ws.Cells[row, 6].Formula = "=SUBTOTAL(9,F" + entityStartRow + ":F" + (row - 1) + ")";
                            ws.Cells[row, 7].Formula = "=SUBTOTAL(9,G" + entityStartRow + ":G" + (row - 1) + ")";

                            ws.Cells[row, 8].Value = "";
                            ws.Cells[row, 9].Value = "";
                            ws.Cells[row, 10].Value = "";
                            ws.Cells[row, 11].Formula = "=SUBTOTAL(9,K" + dataStartRow + ":K" + (row - 1) + ")";
                            ws.Cells[row, 12].Formula = "=SUBTOTAL(9,L" + dataStartRow + ":L" + (row - 1) + ")";
                            ws.Cells[row, 13].Formula = "=SUBTOTAL(9,M" + dataStartRow + ":M" + (row - 1) + ")";
                            ws.Cells[row, 14].Formula = "=SUBTOTAL(9,N" + dataStartRow + ":N" + (row - 1) + ")";
                            ws.Cells[row, 15].Value = "";
                            ws.Cells[row, 16].Value = "";
                            ws.Cells[row, 17].Value = "";

                            ws.Cells[row, 1].Style.Font.Bold = true;
                            ws.Cells[row, 1].Style.Font.Size = 13;
                            ws.Cells[row, 1].Style.Font.Bold = true;
                            ws.Cells[row, 4].Style.Font.Bold = true;
                            ws.Cells[row, 4].Style.Font.Size = 13;
                            ws.Cells[row, 5].Style.Font.Bold = true;
                            ws.Cells[row, 5].Style.Font.Size = 13;
                            ws.Cells[row, 6].Style.Font.Bold = true;
                            ws.Cells[row, 6].Style.Font.Size = 13;
                            ws.Cells[row, 7].Style.Font.Bold = true;
                            ws.Cells[row, 7].Style.Font.Size = 13;
                            ws.Cells[row, 8].Style.Font.Bold = true;
                            ws.Cells[row, 8].Style.Font.Size = 13;
                            ws.Cells[row, 9].Style.Font.Bold = true;
                            ws.Cells[row, 9].Style.Font.Size = 13;
                            ws.Cells[row, 10].Style.Font.Bold = true;
                            ws.Cells[row, 10].Style.Font.Size = 13;
                            ws.Cells[row, 11].Style.Font.Bold = true;
                            ws.Cells[row, 11].Style.Font.Size = 13;
                            ws.Cells[row, 12].Style.Font.Bold = true;
                            ws.Cells[row, 12].Style.Font.Size = 13;
                            ws.Cells[row, 13].Style.Font.Bold = true;
                            ws.Cells[row, 13].Style.Font.Size = 13;
                            ws.Cells[row, 14].Style.Font.Bold = true;
                            ws.Cells[row, 14].Style.Font.Size = 13;

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
                            ws.Cells[row, 13].Style.Fill.PatternType = ExcelFillStyle.Solid;
                            ws.Cells[row, 14].Style.Fill.PatternType = ExcelFillStyle.Solid;


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
                            ws.Cells[row, 13].Style.Fill.BackgroundColor.SetColor(System.Drawing.Color.Gray);
                            ws.Cells[row, 14].Style.Fill.BackgroundColor.SetColor(System.Drawing.Color.Gray);
                        }

                        //if (rowType == "Client Entity Total")
                        //{
                        //    ws.Cells[row, 5].Formula = "=SUBTOTAL(9,E" + clientEntityStartRow + ":E" + (row - 1) + ")";
                        //    ws.Cells[row, 6].Formula = "=SUBTOTAL(9,F" + clientEntityStartRow + ":F" + (row - 1) + ")";
                        //    ws.Cells[row, 7].Value = "";
                        //    ws.Cells[row, 9].Value = "";
                        //    ws.Cells[row, 10].Value = "";
                        //    ws.Cells[row, 11].Value = "";
                        //    ws.Cells[row, 12].Value = "";
                        //    ws.Cells[row, 13].Value = ""; ws.Cells[row, 1].Style.Font.Bold = true;
                        //    ws.Cells[row, 1].Style.Font.Size = 12;
                        //    ws.Cells[row, 1].Style.Font.Bold = true;
                        //    ws.Cells[row, 2].Style.Font.Size = 12;
                        //    ws.Cells[row, 2].Style.Font.Bold = true;
                        //    ws.Cells[row, 5].Style.Font.Bold = true;
                        //    ws.Cells[row, 5].Style.Font.Size = 12;
                        //    ws.Cells[row, 6].Style.Font.Bold = true;
                        //    ws.Cells[row, 6].Style.Font.Size = 12;

                        //    ws.Cells[row, 1].Style.Fill.PatternType = ExcelFillStyle.Solid;
                        //    ws.Cells[row, 2].Style.Fill.PatternType = ExcelFillStyle.Solid;
                        //    ws.Cells[row, 3].Style.Fill.PatternType = ExcelFillStyle.Solid;
                        //    ws.Cells[row, 4].Style.Fill.PatternType = ExcelFillStyle.Solid;
                        //    ws.Cells[row, 5].Style.Fill.PatternType = ExcelFillStyle.Solid;
                        //    ws.Cells[row, 6].Style.Fill.PatternType = ExcelFillStyle.Solid;
                        //    ws.Cells[row, 7].Style.Fill.PatternType = ExcelFillStyle.Solid;
                        //    ws.Cells[row, 8].Style.Fill.PatternType = ExcelFillStyle.Solid;
                        //    ws.Cells[row, 9].Style.Fill.PatternType = ExcelFillStyle.Solid;
                        //    ws.Cells[row, 10].Style.Fill.PatternType = ExcelFillStyle.Solid;


                        //    ws.Cells[row, 1].Style.Fill.BackgroundColor.SetColor(System.Drawing.Color.LightGray);
                        //    ws.Cells[row, 2].Style.Fill.BackgroundColor.SetColor(System.Drawing.Color.LightGray);
                        //    ws.Cells[row, 3].Style.Fill.BackgroundColor.SetColor(System.Drawing.Color.LightGray);
                        //    ws.Cells[row, 4].Style.Fill.BackgroundColor.SetColor(System.Drawing.Color.LightGray);
                        //    ws.Cells[row, 5].Style.Fill.BackgroundColor.SetColor(System.Drawing.Color.LightGray);
                        //    ws.Cells[row, 6].Style.Fill.BackgroundColor.SetColor(System.Drawing.Color.LightGray);
                        //    ws.Cells[row, 7].Style.Fill.BackgroundColor.SetColor(System.Drawing.Color.LightGray);
                        //    ws.Cells[row, 8].Style.Fill.BackgroundColor.SetColor(System.Drawing.Color.LightGray);
                        //    ws.Cells[row, 9].Style.Fill.BackgroundColor.SetColor(System.Drawing.Color.LightGray);
                        //    ws.Cells[row, 10].Style.Fill.BackgroundColor.SetColor(System.Drawing.Color.LightGray);

                        //}

                        if (rowType == "Detail")
                        {
                            ws.Cells[row, 15].Value = "";
                            ws.Cells[row, 16].Value = "";
                            ws.Cells[row, 17].Value = "";

                        }
                        //if (ws.Cells[row, 1].Text != client && rowType == "Detail")
                        //{
                        //    entity = "";
                        //    client = ws.Cells[row, 1].Text;
                        //    clientStartRow = row;
                        //}
                        if (ws.Cells[row, 1].Text != entity && rowType == "Detail")
                        {
                            entity = ws.Cells[row, 1].Text;
                            entityStartRow = row;
                        }

                        dataEndRow = row;
                    }

                    ExcelRange dataRange = ws.Cells[dataStartRow - 1, 1, dataEndRow, 9];
                    dataRange.AutoFilter = true;

                    ws.Cells["A:N"].AutoFitColumns();
                    ExcelPackage.LicenseContext = LicenseContext.NonCommercial;
                    CommonOpenFileDialog dialog = new CommonOpenFileDialog();
                    dialog.InitialDirectory = "C:\\Users";
                    dialog.IsFolderPicker = true;

                    string fileName;

                    fileName = "";

                    if (dialog.ShowDialog() == CommonFileDialogResult.Ok)
                    {
                        fileName = dialog.FileName + "\\" + "CollectionReport.xlsx";

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
