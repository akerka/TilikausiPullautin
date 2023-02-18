using System;
using System.IO;
using System.Globalization;
using OfficeOpenXml;
using System.Drawing;
using OfficeOpenXml.Style;

class Pullautin {
    static void Main(string[] args) {
        while (true)
        {
            // Ask for the desired year
            Console.WriteLine("Kirjoita mille vuodelle tämä tehdään ja paina \"Enter\".");
    
            var input = Console.ReadLine();
            int year = 0;
            DateTime requested_year = DateTime.UnixEpoch;

            if (input != null)
            {
                if (!int.TryParse(input, out year))
                {
                    Console.WriteLine("\nJos tällä kertaa antaasit numeroina vuosiluvun...\n\n");
                    continue;
                }
                if (year < DateTime.Now.Year - 2)
                {
                    Console.WriteLine(string.Format("\nTaitaa olla myöhästä tehdä tilikauden kirjanpitoa vuodelle {0}. Koitahan uudestaan.\n\n", year));
                    continue;
                }

                try
                {
                    requested_year = new DateTime(year, 1, 1);
                }
                catch (Exception)
                {
                    Console.WriteLine("\nOof, ei tollasta vuosilukua pysty edes parsimaan...");
                    Console.WriteLine("Yritäs uudestaan.\n\n");
                    continue;
                }
            }

            Console.WriteLine("Kirjoita kenelle tämä tulee ja paina \"Enter\".");
            string name = Console.ReadLine();

            Console.WriteLine("\nTyöstetään...\n");

            // Define filename and make sure the intended path for it exists
            string documents_path_with_subfolder = string.Format(@"{0}\Vip-Hius", Environment.GetFolderPath(Environment.SpecialFolder.MyDocuments));
            if (!Directory.Exists(documents_path_with_subfolder))
            {
                Directory.CreateDirectory(documents_path_with_subfolder);
            }
            string filename = string.Format("{0}_tilikausi_{1}.xlsx", name, year);
            string path = Path.Combine(documents_path_with_subfolder, filename);

            if (File.Exists(path))
            {
                Console.WriteLine("Tälle nimelle ja vuodelle on jo tilikausi-dokumentti olemassa! \nKäytä toista vuosilukua tai nimeä, tai poista tai siirrä olemassa oleva dokumentti toiseen kansioon ennen kuin yrität uudestaan.\n\n\n");
                continue;
            }

            ExcelPackage.LicenseContext = LicenseContext.NonCommercial;
            var package = new ExcelPackage(new FileInfo(path));

            // Used for finding the monthly summary cells for yearly summary page
            Dictionary<string, int> monthlyTotalRows = new Dictionary<string, int>();

            CultureInfo cultureInfo = new CultureInfo("fi-FI");

            int a = 1;
            int b = 2;
            int c = 3;
            int d = 4;
            int e = 5;
            int f = 6;

            // iterate months, adding the dates, days, and formulas to the corresponding cells
            for (int i = 0; i < 12; i++)
            {
                string month = cultureInfo.TextInfo.ToTitleCase(requested_year.ToString("MMMM", cultureInfo));
                Console.WriteLine(month);
                ExcelWorksheet worksheet = package.Workbook.Worksheets.Add(month);

                // Sheet title
                worksheet.Cells["C1"].Style.HorizontalAlignment = ExcelHorizontalAlignment.Left;
                worksheet.Cells["C1"].Value = "Vip-Hius " + name;
                worksheet.Cells["C1"].Style.Font.Bold = true;
                worksheet.Cells["C1"].Style.Font.Italic = true;
                worksheet.Cells["C1"].Style.Font.UnderLine = true;
                worksheet.Cells["C1"].Style.Font.Size = 15;
                worksheet.Cells["E2"].Value = string.Format("{0} {1}", month, year);

                worksheet.Cells["C3"].Value = "Alv 24%";
                worksheet.Cells["D3"].Value = "Alv 24%";

                // Sheet headers
                worksheet.Cells["A4"].Value = "Pvm";
                worksheet.Cells["A4"].Style.Font.Bold = true;
                worksheet.Columns[a].Style.HorizontalAlignment = ExcelHorizontalAlignment.Center;
                worksheet.Columns[a].Width = 4.5;
                worksheet.Cells["B4"].Value = "Päivä";
                worksheet.Cells["B4"].Style.Font.Bold = true;
                worksheet.Columns[b].Style.HorizontalAlignment = ExcelHorizontalAlignment.Center;
                worksheet.Columns[b].Width = 5;
                worksheet.Cells["C4"].Value = "Työt";
                worksheet.Cells["C4"].Style.Font.Bold = true;
                worksheet.Columns[c].Width = 12;
                worksheet.Cells["D4"].Value = "Aineet";
                worksheet.Cells["D4"].Style.Font.Bold = true;
                worksheet.Columns[d].Width = 12;
                worksheet.Cells["E4"].Value = "Viikkokoonti";
                worksheet.Cells["E4"].Style.Font.Bold = true;

                // Weekly summary headers
                worksheet.Cells["E5"].Value = "Työt";
                worksheet.Cells["E5"].Style.Font.Bold = true;
                worksheet.Columns[e].Width = 12;
                worksheet.Cells["F5"].Value = "Aineet";
                worksheet.Cells["F5"].Style.Font.Bold = true;
                worksheet.Columns[f].Width = 12;

                int daysInMonth = DateTime.DaysInMonth(requested_year.Year, requested_year.Month);
                int summaryOffset = 0;

                int rowOffset = 6;
                int dayNumberOffset = 1;

                int rowCounter = rowOffset;

                int lastSummaryRow = 0;

                List<string> summaryRowsTyöt = new List<string>();
                List<string> summaryRowsAineet = new List<string>();

                for (; rowCounter < rowOffset + daysInMonth + summaryOffset; rowCounter++)
                {
                    worksheet.Cells[rowCounter, a].Value = rowCounter - rowOffset + dayNumberOffset - summaryOffset;
                    worksheet.Cells[rowCounter, a].Style.Font.Bold = true;
                    worksheet.Cells[rowCounter, b].Value = new CultureInfo("fi-FI").DateTimeFormat.GetDayName(requested_year.AddDays(rowCounter - rowOffset - summaryOffset).DayOfWeek).Substring(0,2);
                    worksheet.Cells[rowCounter, b].Style.Font.Bold = true;

                    worksheet.Cells[rowCounter, c].Style.Numberformat.Format = "0.00€";
                    worksheet.Cells[rowCounter, c].Value = 0;
                    worksheet.Cells[rowCounter, d].Style.Numberformat.Format = "0.00€";
                    worksheet.Cells[rowCounter, d].Value = 0;

                    if (requested_year.AddDays(rowCounter - rowOffset - summaryOffset).DayOfWeek == DayOfWeek.Sunday)
                    {
                        // add empty row to the calculations for weekly summary row
                        rowCounter++;
                        summaryOffset++;

                        // Add summary formulas for this week
                        // Työt
                        worksheet.Cells[rowCounter, e].Formula = string.Format("=SUM(C{0}:C{1})", rowCounter - 1, rowCounter - (rowCounter - lastSummaryRow) + 1);
                        worksheet.Cells[rowCounter, e].Style.Font.Bold = true;
                        worksheet.Cells[rowCounter, e].Style.Border.BorderAround(ExcelBorderStyle.Medium);
                        worksheet.Cells[rowCounter, e].Style.Numberformat.Format = "0.00€";
                        summaryRowsTyöt.Add("C" + rowCounter);

                        // Aineet
                        worksheet.Cells[rowCounter, f].Formula = string.Format("=SUM(D{0}:D{1})", rowCounter - 1, rowCounter - (rowCounter - lastSummaryRow) + 1);
                        worksheet.Cells[rowCounter, f].Style.Font.Bold = true;
                        worksheet.Cells[rowCounter, f].Style.Border.BorderAround(ExcelBorderStyle.Medium);
                        worksheet.Cells[rowCounter, f].Style.Numberformat.Format = "0.00€";
                        summaryRowsAineet.Add("D" + rowCounter);

                        AddHorizontalLineAbove(ref worksheet, rowCounter, 1, 4);

                        lastSummaryRow = rowCounter;
                    }
                }

                if (rowCounter > lastSummaryRow + 1)
                {
                    // Add last weekly total summary formulas
                    // Työt
                    worksheet.Cells[rowCounter, e].Formula = string.Format("=SUM(C{0}:C{1})", rowCounter - 1, rowCounter - (rowCounter - lastSummaryRow) + 1);
                    worksheet.Cells[rowCounter, e].Style.Font.Bold = true;
                    worksheet.Cells[rowCounter, e].Style.Border.BorderAround(ExcelBorderStyle.Medium);
                    worksheet.Cells[rowCounter, e].Style.Numberformat.Format = "0.00€";
                    summaryRowsTyöt.Add("C" + rowCounter);

                    // Aineet
                    worksheet.Cells[rowCounter, f].Formula = string.Format("=SUM(D{0}:D{1})", rowCounter - 1, rowCounter - (rowCounter - lastSummaryRow) + 1);
                    worksheet.Cells[rowCounter, f].Style.Font.Bold = true;
                    worksheet.Cells[rowCounter, f].Style.Border.BorderAround(ExcelBorderStyle.Medium);
                    worksheet.Cells[rowCounter, f].Style.Numberformat.Format = "0.00€";
                    summaryRowsAineet.Add("D" + rowCounter);

                    AddHorizontalLineAbove(ref worksheet, rowCounter, 1, 4);
                }

                rowCounter += 2;

                worksheet.Cells[rowCounter, e].Value = "Kuukausikoonti";
                worksheet.Cells[rowCounter, e].Style.Font.Bold = true;

                rowCounter++;

                // Montly summary headers
                worksheet.Cells[rowCounter, e].Value = "Työt";
                worksheet.Cells[rowCounter, e].Style.Font.Bold = true;
                worksheet.Cells[rowCounter, f].Value = "Aineet";
                worksheet.Cells[rowCounter, f].Style.Font.Bold = true;

                rowCounter++;

                // Add month total summary formulas
                // Työt
                worksheet.Cells[rowCounter, e].Formula = string.Format("={0}", string.Join('+', summaryRowsTyöt.Select(x => "E" + x.Substring(1, x.Length - 1))));
                worksheet.Cells[rowCounter, e].Style.Font.Bold = true;
                worksheet.Cells[rowCounter, e].Style.Border.BorderAround(ExcelBorderStyle.Medium);
                worksheet.Cells[rowCounter, e].Style.Numberformat.Format = "0.00€";
                
                // Aineet
                worksheet.Cells[rowCounter, f].Formula = string.Format("={0}", string.Join('+', summaryRowsTyöt.Select(x => "F" + x.Substring(1, x.Length - 1))));
                worksheet.Cells[rowCounter, f].Style.Font.Bold = true;
                worksheet.Cells[rowCounter, f].Style.Border.BorderAround(ExcelBorderStyle.Medium);
                worksheet.Cells[rowCounter, f].Style.Numberformat.Format = "0.00€";

                monthlyTotalRows.Add(month, rowCounter);

                AddHorizontalLineAbove(ref worksheet, rowCounter, 1, 4);

                rowCounter++;

                // Alv summary
                worksheet.Cells[rowCounter, e].Value = "Alv 24%";
                worksheet.Cells[rowCounter, f].Value = "Alv 24%";

                rowCounter++;
                worksheet.Cells[rowCounter, e].Style.Numberformat.Format = "0.00€";
                worksheet.Cells[rowCounter, e].Formula = string.Format("=E{0}*0.24", monthlyTotalRows[month]);
                worksheet.Cells[rowCounter, f].Style.Numberformat.Format = "0.00€";
                worksheet.Cells[rowCounter, f].Formula = string.Format("=F{0}*0.24", monthlyTotalRows[month]);

                // Save the current worksheet before continuing.
                package.Save();


                requested_year = requested_year.AddMonths(1);
            }

            

            // Add year summary sheet
            ExcelWorksheet yearSummaryWorksheet = package.Workbook.Worksheets.Add("Vuosikoonti");

            // Year summary sheet title
            yearSummaryWorksheet.Cells["B1"].Style.HorizontalAlignment = ExcelHorizontalAlignment.Left;
            yearSummaryWorksheet.Cells["B1"].Value = "Vip-Hius " + name;
            yearSummaryWorksheet.Cells["B1"].Style.Font.Bold = true;
            yearSummaryWorksheet.Cells["B1"].Style.Font.Italic = true;
            yearSummaryWorksheet.Cells["B1"].Style.Font.UnderLine = true;
            yearSummaryWorksheet.Cells["B1"].Style.Font.Size = 15;

            yearSummaryWorksheet.Cells["A4"].Value = "Vuosikoonti " + requested_year.Year;
            yearSummaryWorksheet.Cells["A4"].Style.Font.Bold = true;
            yearSummaryWorksheet.Cells["A4"].Style.Font.Size = 12;
            yearSummaryWorksheet.Columns[a].Width = 15;

            yearSummaryWorksheet.Cells["B6"].Value = "Työt";
            yearSummaryWorksheet.Cells["B6"].Style.Font.Bold = true;
            yearSummaryWorksheet.Columns[b].Width = 12;
            yearSummaryWorksheet.Cells["C6"].Value = "Aineet";
            yearSummaryWorksheet.Cells["C6"].Style.Font.Bold = true;
            yearSummaryWorksheet.Columns[c].Width = 12;

            int row = 7;

            // Year summary
            foreach (KeyValuePair<string, int> pair in monthlyTotalRows)
            {
                AddHorizontalLineAbove(ref yearSummaryWorksheet, row, 1, 3);

                // month name
                yearSummaryWorksheet.Cells[row, a].Value = pair.Key;
                yearSummaryWorksheet.Cells[row, a].Style.Font.Bold = true;

                // Työt total
                yearSummaryWorksheet.Cells[row, b].Formula = string.Format("={0}!E{1}", pair.Key, pair.Value);
                yearSummaryWorksheet.Cells[row, b].Style.Font.Bold = true;
                yearSummaryWorksheet.Cells[row, b].Style.Numberformat.Format = "0.00€";

                // Aineet total
                yearSummaryWorksheet.Cells[row, c].Formula = string.Format("={0}!F{1}", pair.Key, pair.Value);
                yearSummaryWorksheet.Cells[row, c].Style.Font.Bold = true;
                yearSummaryWorksheet.Cells[row, c].Style.Numberformat.Format = "0.00€";

                row++;
            }
            AddHorizontalLineAbove(ref yearSummaryWorksheet, row, a, c);

            int yearTotalsRow = 21;

            // Year totals
            yearSummaryWorksheet.Cells["A20"].Value = "Koko Vuosi";
            yearSummaryWorksheet.Cells["A20"].Style.Font.Bold = true;
            yearSummaryWorksheet.Cells["B20"].Value = "Työt";
            yearSummaryWorksheet.Cells["B20"].Style.Font.Bold = true;
            yearSummaryWorksheet.Cells["C20"].Value = "Aineet";
            yearSummaryWorksheet.Cells["C20"].Style.Font.Bold = true;

            yearSummaryWorksheet.Cells[yearTotalsRow, b].Formula = "=SUM(B7:B18)";
            yearSummaryWorksheet.Cells[yearTotalsRow, b].Style.Font.Bold = true;
            yearSummaryWorksheet.Cells[yearTotalsRow, b].Style.Numberformat.Format = "0.00€";
            yearSummaryWorksheet.Cells[yearTotalsRow, c].Formula = "=SUM(C7:C18)";
            yearSummaryWorksheet.Cells[yearTotalsRow, c].Style.Font.Bold = true;
            yearSummaryWorksheet.Cells[yearTotalsRow, c].Style.Numberformat.Format = "0.00€";

            // Alv 24%
            yearSummaryWorksheet.Cells["B23"].Value = "Alv 24%";
            yearSummaryWorksheet.Cells["C23"].Value = "Alv 24%";

            yearSummaryWorksheet.Cells["B24"].Style.Numberformat.Format = "0.00€";
            yearSummaryWorksheet.Cells["B24"].Formula = string.Format("=B{0}*0.24", yearTotalsRow);
            yearSummaryWorksheet.Cells["C24"].Style.Numberformat.Format = "0.00€";
            yearSummaryWorksheet.Cells["C24"].Formula = string.Format("=C{0}*0.24", yearTotalsRow);

            package.Save();

            // Yearly summary on the last page, calculated based on the values in the monthly tabs

            Console.WriteLine(string.Format("Tiedosto tallennettu sijaintiin\"{0}\"", path));
            Console.WriteLine("Paina \"Q\" sulkeaksesi sovelluksen, tai \"Enter\" tehdäksesi uuden dokumentin.");


            // Ask user for quit or another round.
            ConsoleKey key = ConsoleKey.NoName;

            while (key != ConsoleKey.Enter && key != ConsoleKey.Q)
            {
                key = Console.ReadKey().Key;
            }

            if (key == ConsoleKey.Enter)
            {
                continue;
            }
            else if (key == ConsoleKey.Q)
            {
                return;
            }
        }
    }

    static void AddHorizontalLineAbove(ref ExcelWorksheet worksheet, int targetRow, int fromCol, int toCol)
    {
        for (int j = fromCol; j <= toCol; j++)
        {
            worksheet.Cells[targetRow, j].Style.Border.Top.Style = ExcelBorderStyle.Thin;
        }
    }
}

