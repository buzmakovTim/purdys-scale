using System;
using System.Collections.Generic;
using System.Linq;
using System.Data;
using System.Drawing;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;
using System.IO.Ports;
using System.IO;
using Excel = Microsoft.Office.Interop.Excel;
using ExcelLibrary.SpreadSheet;
using System.Threading;

namespace PortMainScaleTest
{
    public static class SaveData
    {
        static int today = 0;
        static int yestarday = -1;
        static int tomorrow = 1;


        // Saving Data Start
        public static bool saveData(ShiftRun shiftRun, ManualScaleWeightCollection manualScaleWeightCollection, PLScaleWeightCollection pLScaleWeightCollection)
        {
            Microsoft.Office.Interop.Excel._Application app = new Microsoft.Office.Interop.Excel.Application();
            Microsoft.Office.Interop.Excel._Workbook workbook = app.Workbooks.Add(Type.Missing);
            Microsoft.Office.Interop.Excel._Worksheet worksheet = null;
            worksheet = workbook.Sheets["Sheet1"];
            worksheet = workbook.ActiveSheet;
            worksheet.Name = "PackLineScale";

            // Formating Data
            Excel.Range formatRange;


            // How many cells will be formated for output Data. +2 (To skip 2 lanes from the top)
            int formatTo = 0;
            if (shiftRun.PlCount >= shiftRun.ManualCount)
            {
                formatTo = shiftRun.PlCount + 2;
            }
            else
            {
                formatTo = shiftRun.ManualCount + 2;
            }



            formatRange = worksheet.get_Range("a1", "e1");
            worksheet.get_Range("a1", "b1").Merge(false);
            worksheet.get_Range("d1", "e1").Merge(false);
            formatRange.EntireRow.Font.Bold = true;


            worksheet.Cells[1, 1] = "Running Weight";
            worksheet.Cells[1, 4] = "Adjusted";

            // Date Formating Manual Scale
            formatRange = worksheet.get_Range("e2", "e" + formatTo);
            formatRange.NumberFormat = "MMMM dd yyyy    h:mm:ss AM/PM";
            formatRange.ColumnWidth = 29;

            // Date Formating PL Scale 
            formatRange = worksheet.get_Range("b2", "b" + formatTo);
            formatRange.NumberFormat = "MMMM dd yyyy    h:mm:ss AM/PM";
            formatRange.ColumnWidth = 29;

            // Weight Formating Manual Scale
            formatRange = worksheet.get_Range("d2", "d" + formatTo);
            formatRange.NumberFormat = "#,###.0";  // was #,###.0

            // Weight Formating PL Scale
            formatRange = worksheet.get_Range("a2", "a" + formatTo);
            formatRange.NumberFormat = "#,###.0";  // was #,###.0


            // Choosing format style
            if (shiftRun.saveToNEWformat != true) // Save to NEW style format 
            {

                //              NEW output format
                //
                // Table Data output Formating
                // Primary data

                formatRange = worksheet.get_Range("f1", "f1");
                formatRange.ColumnWidth = 15;
                formatRange = worksheet.get_Range("g1", "g1");
                formatRange.ColumnWidth = 23;
                formatRange = worksheet.get_Range("h1", "ag1");
                formatRange.ColumnWidth = 15;
                //formatRange.HorizontalAlignment();

                worksheet.Cells[1, 6] = "Shift";                                                                // Title Shift
                worksheet.Cells[2, 6] = shiftRun.Shift;                                                         // Data Shift              

                worksheet.Cells[1, 7].ColumnWidth = 25;
                worksheet.Cells[1, 7] = "Date";                                                                 // Title Date
                worksheet.Cells[2, 7] = DateTime.Now.ToString("dddd, dd MMMM yyyy");                            // Data Date

                worksheet.Cells[1, 8] = "Packline";                                                             // Title Packline Number
                worksheet.Cells[2, 8] = ($"{ shiftRun.Location.ToString()} PackLine - { shiftRun.PackLineNumber.ToString()}");  // Data Packline Number 

                worksheet.Cells[1, 9] = "SKU";                                                                  // Title SKU
                worksheet.Cells[2, 9] = shiftRun.Sku;                                                           // Data SKU

                worksheet.Cells[1, 10] = "Count";                                                               // Title Count
                worksheet.Cells[2, 10] = shiftRun.PlCount.ToString();                                           // Data Count

                worksheet.Cells[1, 11] = "Adjusted Count";                                                      // Title Adjusted Count
                worksheet.Cells[2, 11] = shiftRun.ManualCount.ToString();                                       // Data Adjusted Count

                worksheet.Cells[1, 12] = "Count Avg g.";                                                           // Title Count Avg
                if (shiftRun.PlCount > 0)
                {
                    worksheet.Cells[2, 12] = (shiftRun.AverageWeightCount / shiftRun.PlCount).ToString("0.0");      // Data Count Avg              
                }
                else
                {
                    worksheet.Cells[2, 12] = 0;
                }
                worksheet.Cells[1, 13] = "Adjusted Avg g.";                                                        // Title Adjusted Avg
                if (shiftRun.ManualCount > 0)
                {
                    worksheet.Cells[2, 13] = (shiftRun.AverageWeightAdjusted / shiftRun.ManualCount).ToString("0.0");// Data Adjusted Avg           
                }
                else
                {
                    worksheet.Cells[2, 13] = 0;
                }
                worksheet.Cells[1, 14] = "Total Avg g.";                                                           // Title Total Avg
                if (shiftRun.ManualCount > 0 || shiftRun.PlCount > 0)
                {
                    worksheet.Cells[2, 14] = (shiftRun.AverageWeight / (shiftRun.ManualCount + shiftRun.PlCount)).ToString("0.0");// Data Total Avg
                }
                else
                {
                    worksheet.Cells[2, 14] = 0;
                }
                worksheet.Cells[1, 15] = "Target Weight g.";                                                    // Title Target Weight
                worksheet.Cells[2, 15] = shiftRun.TargetWeight.ToString("0.0");                               // Data Target Weight

                // Additional data
                worksheet.Cells[1, 16] = "Given Less g.";                                                       // Title Given Less (weight)
                worksheet.Cells[2, 16] = shiftRun.LessWeight.ToString("0.0");                                        // Data Given Less (weight)

                worksheet.Cells[1, 17] = "Given Target g.";                                                     // Title Given Target (weight)
                worksheet.Cells[2, 17] = shiftRun.TargetWeight.ToString("0.0");                                      // Data Given Target (weight)

                worksheet.Cells[1, 18] = "Given Heavy g.";                                                      // Title Given Heavy (weight)
                worksheet.Cells[2, 18] = shiftRun.HeavyWeight.ToString("0.0");                                       // Data Given Heavy (weight)

                worksheet.Cells[1, 19] = "Count Less";                                                          // Title Count Less
                worksheet.Cells[2, 19] = shiftRun.LessCount;                                                    // Data Count Less

                worksheet.Cells[1, 20] = "Count Target";                                                        // Title Count Target
                worksheet.Cells[2, 20] = shiftRun.TargetCount;                                                  // Data Count Target

                worksheet.Cells[1, 21] = "Count Heavy";                                                         // Title Count Heavy
                worksheet.Cells[2, 21] = shiftRun.HeavyCount;                                                   // Data Count Heavy

                worksheet.Cells[1, 22] = "Less Avg g.";                                                         // Title Less Avg g.
                if (shiftRun.LessCount > 0)
                {
                    worksheet.Cells[2, 22] = (shiftRun.AverageWeightLess / shiftRun.LessCount).ToString("0.0");   // Data Less Avg g.     
                }
                else
                {
                    worksheet.Cells[2, 22] = 0;
                }

                worksheet.Cells[1, 23] = "Target Avg g.";                                                       // Title Target Avg g.
                if (shiftRun.TargetCount > 0)
                {
                    worksheet.Cells[2, 23] = (shiftRun.AverageWeightTarget / shiftRun.TargetCount).ToString("0.0");// Data Target Avg g.      
                }
                else
                {
                    worksheet.Cells[2, 23] = 0;
                }

                worksheet.Cells[1, 24] = "Heavy Avg g.";                                                        // Title Heavy Avg g.
                if (shiftRun.HeavyCount > 0)
                {
                    worksheet.Cells[2, 24] = (shiftRun.AverageWeightHeavy / shiftRun.HeavyCount).ToString("0.0"); // Data Heavy Avg g.    
                }
                else
                {
                    worksheet.Cells[2, 24] = 0;
                }

                worksheet.Cells[1, 25] = "Given Away KG";                                                         // Title Given Away KG
                worksheet.Cells[2, 25] = (shiftRun.kgGivingAway * 0.001).ToString("0.000");                    // Data Given Away KG

                worksheet.Cells[1, 26] = "Given Away %";                                                          // Title Given Away %
                worksheet.Cells[2, 26] = shiftRun.percentageGivingAway.ToString("0.0 %");                         // Data Given Away KG

                worksheet.Cells[1, 27] = "Units p/h Target";                                                      // Title UPH Target
                worksheet.Cells[2, 27] = shiftRun.productivityTarget.ToString();                                  // Data UPH Target

                worksheet.Cells[1, 28] = "Units p/h Actual";                                                      // Title UPH Actual
                worksheet.Cells[2, 28] = shiftRun.productivityActual.ToString();                                  // Data UPH Actual

                worksheet.Cells[1, 29] = "Running Efficiency";                                                     // Running Efficiency
                worksheet.Cells[2, 29] = (shiftRun.runningEfficiency * 0.01).ToString("0.0 %");                    // Data Running Efficiency  %

                worksheet.Cells[1, 30] = "Expected Efficiency";                                                    // Running Efficiency
                worksheet.Cells[2, 30] = (shiftRun.expectedEfficiency * 0.01).ToString("0.0 %");                   // Data Expected Efficiency  %

                worksheet.Cells[1, 31] = "Staff Required";                                                        // Total People Expected
                worksheet.Cells[2, 31] = shiftRun.StaffNumberRequired;                                            // Total People Expected

                worksheet.Cells[1, 32] = "Staff Actual";                                                          // Total People Actual
                worksheet.Cells[2, 32] = shiftRun.StaffNumberActual;                                              // Total People Actual

                worksheet.Cells[1, 33] = "Productivity";                                                           // Productivity Run
                worksheet.Cells[2, 33] = (shiftRun.productivityRun * 0.01).ToString("0.0 %");                      // Productivity Run data %

                worksheet.Cells[1, 34] = "Run Time Seconds";                                                       // Total Run Time in Seconds
                worksheet.Cells[2, 34] = shiftRun.runningTimeInSeconds;                                            // Total Run Time in Seconds

                worksheet.Cells[1, 35] = "Stop Time Seconds";                                                      // Total Stop Time
                worksheet.Cells[2, 35] = shiftRun.totalBreakTimeInSeconds;                                         // Total Stop Time


                // Ends Table Data output Formating
                //
                //
            }
            if (shiftRun.saveToNEWformat != false) // Save to old style format
            {
                //Pack Line N formating
                worksheet.get_Range("h1", "i2").Merge(false);
                worksheet.get_Range("h1").Font.Size = 18;
                worksheet.get_Range("h1", "i2").HorizontalAlignment = 3;
                worksheet.get_Range("h1", "i2").VerticalAlignment = 2;
                worksheet.Cells[1, 8] = ($"PackLine - { shiftRun.PackLineNumber.ToString()}"); // Packline Number


                //Count Adjust SKU Weights formating
                formatRange = worksheet.get_Range("h3", "h" + formatTo);
                formatRange.ColumnWidth = 15;
                formatRange = worksheet.get_Range("i3", "i" + formatTo);
                formatRange.ColumnWidth = 15;

                worksheet.Cells[3, 8] = "COUNT";
                worksheet.Cells[3, 9] = "ADJUSTED";

                worksheet.get_Range("h4", "h5").Merge(false);
                worksheet.get_Range("h4").Font.Size = 18;
                worksheet.get_Range("h3", "h4").HorizontalAlignment = 3;
                worksheet.get_Range("h4", "h4").VerticalAlignment = 2;

                worksheet.Cells[4, 8] = shiftRun.PlCount.ToString(); // Count PL Scale data

                worksheet.get_Range("i4", "i5").Merge(false);
                worksheet.get_Range("i4").Font.Size = 18;
                worksheet.get_Range("i3", "i4").HorizontalAlignment = 3;
                worksheet.get_Range("i4", "i4").VerticalAlignment = 2;

                worksheet.Cells[4, 9] = shiftRun.ManualCount.ToString(); // Count Manual Scale data 

                //SKU Formating
                worksheet.get_Range("h7", "h8").Merge(false);
                worksheet.get_Range("i7", "i8").Merge(false);
                worksheet.get_Range("h7", "i7").Font.Size = 18;
                worksheet.get_Range("h7", "h8").HorizontalAlignment = 3;
                worksheet.get_Range("h7", "h8").VerticalAlignment = 2;
                worksheet.get_Range("i7", "i8").HorizontalAlignment = 3;
                worksheet.get_Range("i7", "i8").VerticalAlignment = 2;
                worksheet.get_Range("h7", "i7").Font.Color = System.Drawing.ColorTranslator.ToOle(System.Drawing.Color.Green);
                //worksheet.get_Range("h5", "i5").Interior.Color = System.Drawing.
                //ColorTranslator.ToOle(System.Drawing.Color.Yellow);
                worksheet.Cells[7, 8] = "SKU";
                worksheet.Cells[7, 9] = shiftRun.Sku; // SKU Data


                // Less Target Heavy Average Formating
                worksheet.get_Range("h9", "i13").Font.Size = 14;
                worksheet.get_Range("i9", "i11").NumberFormat = "#,###.0"; // Formating weight 100.0 g

                worksheet.get_Range("h9", "i9").Interior.Color = System.Drawing. // Less Background color
                ColorTranslator.ToOle(System.Drawing.Color.Yellow); // Less Background color
                worksheet.Cells[9, 8] = "Less:";
                worksheet.Cells[9, 9] = shiftRun.LessWeight.ToString();
                worksheet.Cells[9, 10] = shiftRun.LessCount.ToString(); // Total Less count

                worksheet.get_Range("h10", "i10").Interior.Color = System.Drawing. // Target Background color
                ColorTranslator.ToOle(System.Drawing.Color.Green); // Target Background color
                worksheet.Cells[10, 8] = "Target:";
                worksheet.Cells[10, 9] = shiftRun.TargetWeight.ToString();
                worksheet.Cells[10, 10] = shiftRun.TargetCount.ToString(); // Total Target Count

                worksheet.get_Range("h11", "i11").Interior.Color = System.Drawing. // Heavy Background color
                ColorTranslator.ToOle(System.Drawing.Color.Red); // Heavy Background color
                worksheet.Cells[11, 8] = "Heavy:";
                worksheet.Cells[11, 9] = shiftRun.HeavyWeight.ToString();
                worksheet.Cells[11, 10] = shiftRun.HeavyCount.ToString(); // Total Heavy count

                //worksheet.get_Range("h13", "i13").Interior.Color = System.Drawing. // Average Background color
                //ColorTranslator.ToOle(System.Drawing.Color.Purple); // Average Background color
                worksheet.Cells[13, 8] = "Average:";
                worksheet.Cells[13, 9] = (shiftRun.AverageWeight / (shiftRun.ManualCount + shiftRun.PlCount)).ToString("0.0 g");



                //Line for exporting
                worksheet.get_Range("r8", "x8").Interior.Color = System.Drawing. // Background color
                ColorTranslator.ToOle(System.Drawing.Color.Red); // Background color

                worksheet.Cells[6, 18] = "Date";
                worksheet.Cells[8, 18] = DateTime.Now.ToString("dddd, dd MMMM yyyy");
                worksheet.Cells[6, 19] = "SKU";
                worksheet.Cells[8, 19] = worksheet.Cells[7, 9];
                worksheet.Cells[6, 20] = "Count";
                worksheet.Cells[8, 20] = worksheet.Cells[4, 8];
                worksheet.Cells[6, 21] = "Adj Count";
                worksheet.Cells[8, 21] = worksheet.Cells[4, 9];
                worksheet.Cells[6, 22] = "Average";
                worksheet.Cells[8, 22] = worksheet.Cells[13, 9];
                worksheet.Cells[6, 23] = "Target";
                worksheet.Cells[8, 23] = worksheet.Cells[10, 9];
                worksheet.Cells[6, 24] = "Packline N";
                worksheet.Cells[8, 24] = ($"PackLine - { shiftRun.PackLineNumber.ToString()}");
                // End Line for exporting
            }

            // Manual weight to Exel all data from Collection ShiftRun
            int rowWeightManual = 2;
            int colWeightManual = 4;
            foreach (ManualScaleWeight data in manualScaleWeightCollection)
            {
                if (data.WeightManualScale != 0.0)
                {
                    rowWeightManual++;
                    worksheet.Cells[rowWeightManual, colWeightManual] = (decimal)data.WeightManualScale;
                }
            }

            // Manual Date to Exel all data from Collection ShiftRun
            int rowDateManual = 2;
            int colDateManual = 5;
            foreach (ManualScaleWeight data in manualScaleWeightCollection)
            {
                if (data.WeightManualScale != 0.0)
                {
                    rowDateManual++;
                    worksheet.Cells[rowDateManual, colDateManual] = data.dateTime;
                }
            }


            // PL weight to Exel all data from Collection ShiftRun
            int rowWeightPL = 2;
            int colWeightPL = 1;
            foreach (PLScaleWeight data in pLScaleWeightCollection)
            {
                if (data.WeightPLScale != 0.0)
                {
                    rowWeightPL++;
                    worksheet.Cells[rowWeightPL, colWeightPL] = (decimal)data.WeightPLScale;
                }
            }

            // PL Date to Exel all data from Collection ShiftRun
            int rowDatePL = 2;
            int colDatePL = 2;
            foreach (PLScaleWeight data in pLScaleWeightCollection)
            {
                if (data.WeightPLScale != 0.0)
                {
                    rowDatePL++;
                    worksheet.Cells[rowDatePL, colDatePL] = data.dateTime;
                }
            }


            // File name generating
            string dateTime = DateTime.Now.ToString("MMMM dd yyyy  h mm tt");
            string fileName = shiftRun.Location + "_PL" + shiftRun.PackLineNumber + "_" + shiftRun.Sku + "@" + dateTime + ".xlsx";

            // Path generating   \\hedgehog\Syteline\PacklineScaleData\2020\May\FileName.xlsx"
            string year = DateTime.Now.Year.ToString();
            string month = String.Format("{0:MMMM}", DateTime.Now);
            //string pathToSave = @"\\hedgehog\Syteline\PacklineScaleData\", "year", "\TestFile.xlsx";
            string[] path = { @"\\hedgehog\Syteline\PacklineScaleData\", year, month, fileName };
            string[] pathForEmail = { @"\\hedgehog\Syteline\PacklineScaleData\", year, month };

            string fullPathForEmail = Path.Combine(pathForEmail);
            string fullPath = Path.Combine(path);

            // Trying to save data
            try
            {
                workbook.SaveAs(fullPath);


                // Check if file has been saved. Show message if saved successfully 
                if (System.IO.File.Exists(fullPath))
                {
                    workbook.SaveAs(fileName); // Save to Documents Folder even submission is successful (BackUP)
                    app.Quit();

                    MessageBox.Show("Data has been saved successfully!", "Confirmation", MessageBoxButtons.OK, MessageBoxIcon.Information);

                    return true;
                }
                return false;
                app.Quit();
            }
            // If Access denied Data will be save in Document folder
            catch (Exception ex)
            {
                workbook.SaveAs(fileName); // Save to Documents Folder even submission is successful (BackUP)
                app.Quit();
                MessageBox.Show("Access denied! Data will be save in Documents folder \nNO action require.\nIT has been notified about it", "Message", MessageBoxButtons.OK, MessageBoxIcon.Error);

                ThreadPool.QueueUserWorkItem(state => SendEmail.sendEmailToHDAboutBackUpData(shiftRun, fileName, fullPathForEmail));
                //SendEmail.sendEmailToHDAboutBackUpData(shiftRun, fileName, fullPathForEmail);
                return true;
            }

        } // Save data End





        // Save Data for Dayly Report start
        public static bool saveDataforDayRepor(ShiftRun shiftRun, ManualScaleWeightCollection manualScaleWeightCollection, PLScaleWeightCollection pLScaleWeightCollection)
        {

            string fullPath;
            if (shiftRun.timsTesting)
            {
                fullPath = getDaylyReportFileName(tomorrow); // If testing
            }
            else
            {
                fullPath = getDaylyReportFileName(today);
            }

            FileInfo daylyReport = new FileInfo(fullPath);

            if (!daylyReport.Exists) // If File not Exist YET Try to create and add info
            {
                //MessageBox.Show("file doesn't exists!", "Confirmation", MessageBoxButtons.OK, MessageBoxIcon.Information);

                // File not Exist Creating the file
                //createDaylyReportFile();

                if (createDaylyReportFile(shiftRun)) // If file created successfully we can add data
                {
                    //Adding to File all data
                    addDataToDaylyReportFile(fullPath, shiftRun, manualScaleWeightCollection, pLScaleWeightCollection);
                }
            }
            else  // If File Exist just to add info
            {


                try // Let's check if this file Open and currently in used by other PL. Let's delay and try to save in 5 seconds.
                {
                    Stream s = File.Open(fullPath, FileMode.Open, FileAccess.Read, FileShare.None); // File not in use we can open it
                    s.Close();

                    //Adding to File all data
                    addDataToDaylyReportFile(fullPath, shiftRun, manualScaleWeightCollection, pLScaleWeightCollection);

                    
                }
                catch (Exception) // File is OPEN. lets wait till it's closed and then do save
                {
                    shiftRun.isDelaySaving = true;
                    //Console.WriteLine("This saving with delay");
                    Task.Delay(15000).ContinueWith(t => addDataToDaylyReportFile(fullPath, shiftRun, manualScaleWeightCollection, pLScaleWeightCollection));
                }






                //formatRange = worksheet.get_Range("a1", "e1");

                //Excel.Range cell = (Excel.Range)formatRange.Cells[2, 1];
                //int i = 2;
                //if (xlRange.Cells[i, 1] != null)
                //{
                //    i++;
                //}
                //xlRange.Cells[i, 1] = shiftRun.Sku;
                //Console.WriteLine("Number of ROWS - " + xlRange.Rows.Count);

                //for (int i = 1; i <= xlRange.Rows.Count+1; i++)
                //{
                //if (xlRange.Cells[i, 1] == null)
                //{
                //int i = xlRange.Rows.Count;
                //xlRange.Cells[i+1, 1] = shiftRun.Sku;
                //xlRange.Cells[1, 2] = shiftRun.Sku;
                //break;
                //}
                //}
                //string testing = xlWorksheet.Cells[1, 2].Value.ToString();


                //if (!testing.Equals(""))
                //{
                //   MessageBox.Show("Tes I can read this Cell" + testing, "Confirmation", MessageBoxButtons.OK, MessageBoxIcon.Information);
                //   xlRange.Cells[1, 2] = "SHIFT22";

                //}



                //xlWorkbook.Save();
                //xlApp.Quit();
            }



            return true;

        }// Save Data for Dayly Report End




        // Create Dayly Report Excel File Start
        public static bool createDaylyReportFile(ShiftRun shiftRun)
        {
            Microsoft.Office.Interop.Excel._Application app = new Microsoft.Office.Interop.Excel.Application();
            Microsoft.Office.Interop.Excel._Workbook workbook = app.Workbooks.Add(Type.Missing);
            Microsoft.Office.Interop.Excel._Worksheet worksheet = null;
            worksheet = workbook.Sheets["Sheet1"];
            worksheet = workbook.ActiveSheet;
            worksheet.Name = "Daily_Report";

            // Formating Data
            Excel.Range xlRange = worksheet.UsedRange;

            xlRange.EntireRow.Font.Bold = true;

            //xlRange = worksheet.get_Range("A1", "A1");
            //xlRange.ColumnWidth = 29;
            //worksheet.Cells[1, 1] = DateTime.Now.ToString("dddd, dd MMMM yyyy");

            xlRange = worksheet.get_Range("A1", "A1");
            xlRange.ColumnWidth = 14;
            worksheet.Cells[1, 1] = "Date";

            xlRange = worksheet.get_Range("B1", "B1");
            xlRange.ColumnWidth = 6;
            worksheet.Cells[1, 2] = "Shift";

            xlRange = worksheet.get_Range("C1", "C1");
            xlRange.ColumnWidth = 13;
            worksheet.Cells[1, 3] = "Packline";

            xlRange = worksheet.get_Range("D1", "D1");
            xlRange.ColumnWidth = 14;
            worksheet.Cells[1, 4] = "SKU";

            xlRange = worksheet.get_Range("E1", "E1");
            xlRange.ColumnWidth = 0;
            worksheet.Cells[1, 5] = "Description";

            xlRange = worksheet.get_Range("F1", "F1");
            xlRange.ColumnWidth = 10;
            worksheet.Cells[1, 6] = "Start Time";

            xlRange = worksheet.get_Range("G1", "G1");
            xlRange.ColumnWidth = 10;
            worksheet.Cells[1, 7] = "Finish Time";

            xlRange = worksheet.get_Range("H1", "H1");
            xlRange.ColumnWidth = 10;
            worksheet.Cells[1, 8] = "Run Time";

            xlRange = worksheet.get_Range("I1", "I1");
            xlRange.ColumnWidth = 10;
            worksheet.Cells[1, 9] = "Idle Time";

            xlRange = worksheet.get_Range("J1", "J1");
            xlRange.ColumnWidth = 8;
            worksheet.Cells[1, 10] = "Count";

            xlRange = worksheet.get_Range("K1", "K1");
            xlRange.ColumnWidth = 9;
            worksheet.Cells[1, 11] = "Adjusted";

            xlRange = worksheet.get_Range("L1", "L1");
            xlRange.ColumnWidth = 13;
            worksheet.Cells[1, 12] = "Given Less g.";

            xlRange = worksheet.get_Range("M1", "M1");
            xlRange.ColumnWidth = 13;
            worksheet.Cells[1, 13] = "Given Target g.";

             xlRange = worksheet.get_Range("N1", "N1");
            xlRange.ColumnWidth = 15;
            worksheet.Cells[1, 14] = "Given Heavy g.";

             xlRange = worksheet.get_Range("O1", "O1");
            xlRange.ColumnWidth = 16;
            worksheet.Cells[1, 15] = "Total Avg g."; 

             xlRange = worksheet.get_Range("P1", "P1");
            xlRange.ColumnWidth = 13;
            worksheet.Cells[1, 16] = "Given Away KG";

            xlRange = worksheet.get_Range("Q1", "Q1");
            xlRange.ColumnWidth = 13;
            worksheet.Cells[1, 17] = "Given Away %";

             xlRange = worksheet.get_Range("R1", "R1");
            xlRange.ColumnWidth = 14;
            worksheet.Cells[1, 18] = "Units p/h Target";

             xlRange = worksheet.get_Range("S1", "S1");
            xlRange.ColumnWidth = 14;
            worksheet.Cells[1, 19] = "Units p/ h Actual"; 

             xlRange = worksheet.get_Range("T1", "T1");
            xlRange.ColumnWidth = 15;
            worksheet.Cells[1, 20] = "Running Efficiency";

            xlRange = worksheet.get_Range("U1", "U1");
            xlRange.ColumnWidth = 16;
            worksheet.Cells[1, 21] = "Expected Efficiency";

            xlRange = worksheet.get_Range("V1", "V1");
            xlRange.ColumnWidth = 15;
            worksheet.Cells[1, 22] = "Staff Required";

            xlRange = worksheet.get_Range("W1", "W1");
            xlRange.ColumnWidth = 15;
            worksheet.Cells[1, 23] = "Staff Actual";

            xlRange = worksheet.get_Range("X1", "X1");
            xlRange.ColumnWidth = 12;
            worksheet.Cells[1, 24] = "Productivity";




            // Trying to save data
            try
            {
                if (shiftRun.timsTesting) // If Testing we are creating Fila for Tommotow's date so not to screw up existing file
                {
                    workbook.SaveAs(getDaylyReportFileName(today + 1));
                }
                else 
                {
                    workbook.SaveAs(getDaylyReportFileName(today));
                }

                app.Quit();

                return true;

            }
            // If Access denied Data will be save in Document folder
            catch (Exception ex)
            {
                //workbook.SaveAs(fileName); // Save to Documents Folder even submission is successful (BackUP)
                app.Quit();
                //MessageBox.Show("Access denied! Data will be save in Documents folder \nNO action require.\nIT has been notified about it", "Message", MessageBoxButtons.OK, MessageBoxIcon.Error);

                //ThreadPool.QueueUserWorkItem(state => SendEmail.sendEmailToHDAboutBackUpData(shiftRun, fileName, fullPathForEmail));
                //SendEmail.sendEmailToHDAboutBackUpData(shiftRun, fileName, fullPathForEmail);
                return true;
            }

        } // Create Dayly Report Excel File End


        // Add data to dayly Report File start
        public static void addDataToDaylyReportFile(string fileName, ShiftRun shiftRun, ManualScaleWeightCollection manualScaleWeightCollection, PLScaleWeightCollection pLScaleWeightCollection)
        {
            Excel.Application xlApp = new Excel.Application();
            Excel.Workbook xlWorkbook = xlApp.Workbooks.Open(fileName);
            Excel._Worksheet xlWorksheet = xlWorkbook.Sheets[1];
            Excel.Range xlRange = xlWorksheet.UsedRange;

            int i = xlRange.Rows.Count + 1; // Shift to one row Down

            
            xlRange.Cells[i, 3] = shiftRun.Location + " - " + shiftRun.PackLineNumber;
            xlRange.Cells[i, 4] = shiftRun.Sku;
            xlRange.Cells[i, 4].Style.HorizontalAlignment = Microsoft.Office.Interop.Excel.XlHAlign.xlHAlignLeft; // Alighment Left
            xlRange.Cells[i, 5] = shiftRun.PlCount;

            xlWorksheet.Cells[i, 1] = DateTime.Now.ToString("yyyy-MM-dd");
            xlWorksheet.Cells[i, 2] = shiftRun.Shift;
            xlWorksheet.Cells[i, 3] = shiftRun.Location + " Packline - " + shiftRun.PackLineNumber;
            xlWorksheet.Cells[i, 4] = shiftRun.Sku;
            xlWorksheet.Cells[i, 5] = ""; // PlaceHolder for future Description

            xlWorksheet.Cells[i, 6] = shiftRun.startTime.ToString("h:mm tt");  // Start Time
            xlWorksheet.Cells[i, 7] = DateTime.Now.ToString("h:mm tt");  // Finish Time


            TimeSpan runTime = TimeSpan.FromSeconds(shiftRun.runningTimeInSeconds);
            xlWorksheet.Cells[i, 8] = runTime.Hours + ":" + runTime.Minutes + ":" + runTime.Seconds; // Totoal Run Time
            xlWorksheet.Cells[i, 40] = shiftRun.runningTimeInSeconds;  // Store it just for Report later

            TimeSpan breakTime = TimeSpan.FromSeconds(shiftRun.totalBreakTimeInSeconds);
            xlWorksheet.Cells[i, 9] = breakTime.Hours + ":" + breakTime.Minutes + ":" + breakTime.Seconds; // Totoal Idle Time
            xlWorksheet.Cells[i, 41] = shiftRun.totalBreakTimeInSeconds;  // Store it just for Report later




            xlWorksheet.Cells[i, 3].Style.HorizontalAlignment = Microsoft.Office.Interop.Excel.XlHAlign.xlHAlignLeft; // Alighment Left

            

            

            

            
            xlWorksheet.Cells[i, 10] = shiftRun.PlCount;  // PL Count

            xlWorksheet.Cells[i, 11] = shiftRun.ManualCount;  // Adjusted

            xlWorksheet.Cells[i, 12] = shiftRun.LessWeight.ToString("0.0");  // Given Less g.

            xlWorksheet.Cells[i, 13] = shiftRun.TargetWeight.ToString("0.0");  // Given Target g.

            xlWorksheet.Cells[i, 14] = shiftRun.HeavyWeight.ToString("0.0");  // Given Heavy g.


            if (shiftRun.ManualCount > 0 || shiftRun.PlCount > 0)
            {
                xlWorksheet.Cells[i, 15] = (shiftRun.AverageWeight / (shiftRun.ManualCount + shiftRun.PlCount)).ToString("0.0"); // Data Total Avg
            }
            else
            {
                xlWorksheet.Cells[i, 15] = 0;
            }

            xlWorksheet.Cells[i, 16] = (shiftRun.kgGivingAway * 0.001).ToString("0.000");  // Given Away KG

            xlWorksheet.Cells[i, 17] = shiftRun.percentageGivingAway.ToString("0.0 %"); // Given Away %

            xlWorksheet.Cells[i, 18] = shiftRun.productivityTarget.ToString();  // Units p/h Target

            xlWorksheet.Cells[i, 19] = shiftRun.productivityActual.ToString();  // Units p/h Actual

            xlWorksheet.Cells[i, 20] = (shiftRun.runningEfficiency * 0.01).ToString("0.0 %");  //  Running Efficiency

            xlWorksheet.Cells[i, 21] = (shiftRun.expectedEfficiency * 0.01).ToString("0.0 %");  //  Expected Efficiency

            xlWorksheet.Cells[i, 22] = shiftRun.StaffNumberRequired;  //  Staff Required

            xlWorksheet.Cells[i, 23] = shiftRun.StaffNumberActual;  //  Staff Actual

            xlWorksheet.Cells[i, 24] = (shiftRun.productivityRun * 0.01).ToString("0.0 %");  // Productivity Run %



            try // Trying to save
            {
                xlWorkbook.Save();
                xlApp.Quit();
            }
            catch (Exception ex)
            {
                xlApp.Quit();
            }
        }
        // Add data to dayly Report File start



        // Check if file Dayly Report exist when start RUN. Create if not exist. Start
        public static void checkIfFileDaylyReportExist(ShiftRun shiftRun)
        {
            FileInfo dailyReport;

            if (shiftRun.timsTesting)
            {
                dailyReport = new FileInfo(getDaylyReportFileName(today + 1));
            }
            else 
            {
                dailyReport = new FileInfo(getDaylyReportFileName(today));
            }

            if (!dailyReport.Exists) // If File not Exist YET Try to create and add info
            {
                createDaylyReportFile(shiftRun);
            }
        }
        // Check if file Dayly Report exist when start RUN. Create if not exist. End






        //Get the File name for Dayly Report Start
        public static string getDaylyReportFileName(int day)
        {

            // File name generating
            string dateTime = DateTime.Now.AddDays(day).ToString("MMMM dd yyyy");
            
            string fileName = dateTime + "_DAILY_REPORT" + ".xlsx";

            string[] path = { @"\\hedgehog\Syteline\PacklineScaleData\ScaleProject\ScaleApp(Tim)\TEMP\", fileName }; // Temporary store at this Location till end of the day
            
            string fileNameFullPath = Path.Combine(path);

            return fileNameFullPath;
        }
        //Get the File name for Dayly Report End












        // Add data to dayly Report File start
        public static bool generateDailyReportFile(ShiftRun shiftRun)
        {
            FileInfo dailyReport;
            int dayToCheck;
            if (shiftRun.timsTesting)
            {
                dailyReport = new FileInfo(getDaylyReportFileName(tomorrow)); // For Tims Testing generating Tomorrow's report
                dayToCheck = tomorrow;
            }
            else
            {
                dailyReport = new FileInfo(getDaylyReportFileName(yestarday)); // Yesterday as we generate report next day
                dayToCheck = yestarday;
            }


            if (dailyReport.Exists) // If File Exist Generate the report
            {
                Excel.Application xlApp = new Excel.Application();
                Excel.Workbook xlWorkbook = xlApp.Workbooks.Open(getDaylyReportFileName(dayToCheck));
                Excel._Worksheet xlWorksheet = xlWorkbook.Sheets[1];
                Excel.Range xlRange = xlWorksheet.UsedRange;
            
            // I'm gonna put it to Try block
            try // Trying to add and to save data
            {

                //xlRange = xlWorksheet.get_Range("A", "S");
                //xlRange.EntireRow.Font.Bold = true;

            int i = xlRange.Rows.Count + 1; // Shift to one row Down
            int x = xlRange.Rows.Count; // For formulas
            int y = xlRange.Rows.Count + 7; // For Summary





            //Summary Layout Set
            

            xlWorksheet.Cells[y + 1, 4].Font.Bold = true;
            xlWorksheet.Cells[y + 2, 4].Font.Bold = true;
            xlWorksheet.Cells[y + 3, 4].Font.Bold = true;
            xlWorksheet.Cells[y + 4, 4].Font.Bold = true;
            xlWorksheet.Cells[y+1, 4] = "KW Packline - 1";
            xlWorksheet.Cells[y+2, 4] = "CH Packline - 1";
            xlWorksheet.Cells[y+3, 4] = "CH Packline - 2";
            xlWorksheet.Cells[y+4, 4] = "CH Packline - 3";


            string startRange = "A1";
            string endRange = "S"+x;
            Excel.Range currentRange = (Excel.Range)xlWorksheet.get_Range(startRange, endRange);
            currentRange.Style.HorizontalAlignment = Microsoft.Office.Interop.Excel.XlHAlign.xlHAlignLeft;




            xlWorksheet.Cells[y, 6].Font.Bold = true;
            xlWorksheet.Cells[y, 7].Font.Bold = true;
            xlWorksheet.Cells[y, 8].Font.Bold = true;
            xlWorksheet.Cells[y, 9].Font.Bold = true;
            xlWorksheet.Cells[y, 10].Font.Bold = true;
            xlWorksheet.Cells[y, 11].Font.Bold = true;
            xlWorksheet.Cells[y, 12].Font.Bold = true;
            xlWorksheet.Cells[y, 13].Font.Bold = true;
            xlWorksheet.Cells[y, 14].Font.Bold = true; 
            xlWorksheet.Cells[y, 15].Font.Bold = true;
            xlWorksheet.Cells[y, 16].Font.Bold = true;
            xlWorksheet.Cells[y, 17].Font.Bold = true;
            xlWorksheet.Cells[y, 6] = "Runs"; // Total Runs per Line
            //xlWorksheet.Cells[y, 5].HorizontalAlignment = Microsoft.Office.Interop.Excel.XlHAlign.xlHAlignLeft;
            xlWorksheet.Cells[y, 7] = "Run Time"; // Total Run Time per Line
            xlWorksheet.Cells[y, 8] = "Idle Time"; // Total Idle Time per Line
            xlWorksheet.Cells[y, 9] = "Productivity"; // Total Count per Line
            xlWorksheet.Cells[y, 10] = "Count"; // Total Count per Line
            xlWorksheet.Cells[y, 11] = "Adjusted"; // Total Adjusted per Line
            xlWorksheet.Cells[y, 12] = "Given Away KG"; // Total Given Away KG per Line
            xlWorksheet.Cells[y, 13] = "Given Away %"; // Total Given Away % per Line
            xlWorksheet.Cells[y, 14] = "Running Efficiency"; // Total Running Efficiency per Line
            xlWorksheet.Cells[y, 15] = "Expected Efficiency"; // Total Expected Efficiency per Line
            xlWorksheet.Cells[y, 16] = "Staff Required"; // Total Staff required
            xlWorksheet.Cells[y, 17] = "Staff Actual"; // Total Staff actual


            xlWorksheet.get_Range("E"+(y+1), "Q"+(y+4)).Interior.Color = System.Drawing. // Less Background color
            ColorTranslator.ToOle(System.Drawing.Color.FromArgb(210, 207, 245)); // Less Background color


            // Total Runs Per PL start
            int kwPl1 = 0;
            int chPl1 = 0;
            int chPl2 = 0;
            int chPl3 = 0;
            for (int rowN = 2; rowN < xlRange.Rows.Count + 1; rowN++)
            {
                string compareForRuns = xlWorksheet.Cells[rowN, 3].Value.ToString();

                if (compareForRuns.Equals("KW Packline - 1"))
                {
                    kwPl1++;
                }
                if (compareForRuns.Equals("CH Packline - 1"))
                {
                    chPl1++;
                }
                if (compareForRuns.Equals("CH Packline - 2"))
                {
                    chPl2++;
                }
                if (compareForRuns.Equals("CH Packline - 3"))
                {
                    chPl3++;
                }
            }
            xlWorksheet.Cells[y + 1, 6] = kwPl1; // KW Packline 1 
            xlWorksheet.Cells[y + 2, 6] = chPl1; // CH Packline 1
            xlWorksheet.Cells[y + 3, 6] = chPl2; // CH Packline 2
            xlWorksheet.Cells[y + 4, 6] = chPl3; // CH Packline 3
            //Runs End

            // Run Time start
            int dtKwPl1Seconds = 0;
            int dtChPl1Seconds = 0;
            int dtChPl2Seconds = 0;
            int dtChPl3Seconds = 0;
            for (int rowN = 2; rowN < xlRange.Rows.Count + 1; rowN++)
            {
                string compareForRunTime = xlWorksheet.Cells[rowN, 3].Value.ToString();

                if (compareForRunTime.Equals("KW Packline - 1"))
                {
                    dtKwPl1Seconds +=Convert.ToInt32(xlWorksheet.Cells[rowN, 40].Value);
                }
                if (compareForRunTime.Equals("CH Packline - 1"))
                {
                    dtChPl1Seconds +=Convert.ToInt32(xlWorksheet.Cells[rowN, 40].Value);
                }
                if (compareForRunTime.Equals("CH Packline - 2"))
                {
                    dtChPl2Seconds +=Convert.ToInt32(xlWorksheet.Cells[rowN, 40].Value);
                }
                if (compareForRunTime.Equals("CH Packline - 3"))
                {
                    dtChPl3Seconds +=Convert.ToInt32(xlWorksheet.Cells[rowN, 40].Value);
                }

            }
            TimeSpan runTime = TimeSpan.FromSeconds(dtKwPl1Seconds);
            xlWorksheet.Cells[y + 1, 7] = runTime.Hours + ":" + runTime.Minutes + ":" + runTime.Seconds; // Totoal Run Time  KW Packline 1
            TimeSpan runTime1 = TimeSpan.FromSeconds(dtChPl1Seconds);
            xlWorksheet.Cells[y + 2, 7] = runTime1.Hours + ":" + runTime1.Minutes + ":" + runTime1.Seconds; // Totoal Run Time CH Packline 1
            TimeSpan runTime2 = TimeSpan.FromSeconds(dtChPl2Seconds);
            xlWorksheet.Cells[y + 3, 7] = runTime2.Hours + ":" + runTime2.Minutes + ":" + runTime2.Seconds; // Totoal Run Time CH Packline 2
            TimeSpan runTime3 = TimeSpan.FromSeconds(dtChPl3Seconds);
            xlWorksheet.Cells[y + 4, 7] = runTime3.Hours + ":" + runTime3.Minutes + ":" + runTime3.Seconds; // Totoal Run Time CH Packline 3
            // Run Time end


            // Idle Time per PL start
            int dtKwPl1SecondsBreak = 0;
            int dtChPl1SecondsBreak = 0;
            int dtChPl2SecondsBreak = 0;
            int dtChPl3SecondsBreak = 0;
            for (int rowN = 2; rowN < xlRange.Rows.Count + 1; rowN++)
            {
                string compareForRunTime = xlWorksheet.Cells[rowN, 3].Value.ToString();

                if (compareForRunTime.Equals("KW Packline - 1"))
                {
                    dtKwPl1SecondsBreak +=Convert.ToInt32(xlWorksheet.Cells[rowN, 41].Value);
                }
                if (compareForRunTime.Equals("CH Packline - 1"))
                {
                    dtChPl1SecondsBreak +=Convert.ToInt32(xlWorksheet.Cells[rowN, 41].Value);
                }
                if (compareForRunTime.Equals("CH Packline - 2"))
                {
                    dtChPl2SecondsBreak +=Convert.ToInt32(xlWorksheet.Cells[rowN, 41].Value);
                }
                if (compareForRunTime.Equals("CH Packline - 3"))
                {
                    dtChPl3SecondsBreak +=Convert.ToInt32(xlWorksheet.Cells[rowN, 41].Value);
                }

            }
            TimeSpan breakTime = TimeSpan.FromSeconds(dtKwPl1SecondsBreak);
            xlWorksheet.Cells[y + 1, 8] = breakTime.Hours + ":" + breakTime.Minutes + ":" + breakTime.Seconds; // Totoal Run Time  KW Packline 1
            TimeSpan breakTime1 = TimeSpan.FromSeconds(dtChPl1SecondsBreak);
            xlWorksheet.Cells[y + 2, 8] = breakTime1.Hours + ":" + breakTime1.Minutes + ":" + breakTime1.Seconds; // Totoal Run Time CH Packline 1
            TimeSpan breakTime2 = TimeSpan.FromSeconds(dtChPl2SecondsBreak);
            xlWorksheet.Cells[y + 3, 8] = breakTime2.Hours + ":" + breakTime2.Minutes + ":" + breakTime2.Seconds; // Totoal Run Time CH Packline 2
            TimeSpan breakTime3 = TimeSpan.FromSeconds(dtChPl3SecondsBreak);
            xlWorksheet.Cells[y + 4, 8] = breakTime3.Hours + ":" + breakTime3.Minutes + ":" + breakTime3.Seconds; // Totoal Run Time CH Packline 3

            // Idle Time per PL end



            // Productivity stat



            // Productivity end



            //Count start
            double kwPl1Count = 0;
            double chPl1Count = 0;
            double chPl2Count = 0;
            double chPl3Count = 0;
            for (int rowN = 2; rowN < xlRange.Rows.Count + 1; rowN++)
            {
                string compareForRuns = xlWorksheet.Cells[rowN, 3].Value.ToString();

                if (compareForRuns.Equals("KW Packline - 1"))
                {
                    kwPl1Count += xlWorksheet.Cells[rowN, 10].Value;
                }
                if (compareForRuns.Equals("CH Packline - 1"))
                {
                    chPl1Count += xlWorksheet.Cells[rowN, 10].Value;
                }
                if (compareForRuns.Equals("CH Packline - 2"))
                {
                    chPl2Count += xlWorksheet.Cells[rowN, 10].Value;
                }
                if (compareForRuns.Equals("CH Packline - 3"))
                {
                    chPl3Count += xlWorksheet.Cells[rowN, 10].Value;
                }

            }
            xlWorksheet.Cells[y + 1, 10] = kwPl1Count; // KW Packline 1 
            xlWorksheet.Cells[y + 2, 10] = chPl1Count; // CH Packline 1
            xlWorksheet.Cells[y + 3, 10] = chPl2Count; // CH Packline 2
            xlWorksheet.Cells[y + 4, 10] = chPl3Count; // CH Packline 3
            //Count End


            //Adjusted start
            double kwPl1Adjusted = 0;
            double chPl1Adjusted = 0;
            double chPl2Adjusted = 0;
            double chPl3Adjusted = 0;
            for (int rowN = 2; rowN < xlRange.Rows.Count + 1; rowN++)
            {
                string compareForRuns = xlWorksheet.Cells[rowN, 3].Value.ToString();

                if (compareForRuns.Equals("KW Packline - 1"))
                {
                    kwPl1Adjusted += xlWorksheet.Cells[rowN, 11].Value;
                }
                if (compareForRuns.Equals("CH Packline - 1"))
                {
                    chPl1Adjusted += xlWorksheet.Cells[rowN, 11].Value;
                }
                if (compareForRuns.Equals("CH Packline - 2"))
                {
                    chPl2Adjusted += xlWorksheet.Cells[rowN, 11].Value;
                }
                if (compareForRuns.Equals("CH Packline - 3"))
                {
                    chPl3Adjusted += xlWorksheet.Cells[rowN, 11].Value;
                }

            }
            xlWorksheet.Cells[y + 1, 11] = kwPl1Adjusted; // KW Packline 1 
            xlWorksheet.Cells[y + 2, 11] = chPl1Adjusted; // CH Packline 1
            xlWorksheet.Cells[y + 3, 11] = chPl2Adjusted; // CH Packline 2
            xlWorksheet.Cells[y + 4, 11] = chPl3Adjusted; // CH Packline 3
            //Adjusted End


            //Given away KG start
            double kwPl1GivenAwayKG = 0;
            double chPl1GivenAwayKG = 0;
            double chPl2GivenAwayKG = 0;
            double chPl3GivenAwayKG = 0;
            for (int rowN = 2; rowN < xlRange.Rows.Count + 1; rowN++)
            {
                string compareForRuns = xlWorksheet.Cells[rowN, 3].Value.ToString();

                if (compareForRuns.Equals("KW Packline - 1"))
                {
                    kwPl1GivenAwayKG += xlWorksheet.Cells[rowN, 16].Value;
                }
                if (compareForRuns.Equals("CH Packline - 1"))
                {
                    chPl1GivenAwayKG += xlWorksheet.Cells[rowN, 16].Value;
                }
                if (compareForRuns.Equals("CH Packline - 2"))
                {
                    chPl2GivenAwayKG += xlWorksheet.Cells[rowN, 16].Value;
                }
                if (compareForRuns.Equals("CH Packline - 3"))
                {
                    chPl3GivenAwayKG += xlWorksheet.Cells[rowN, 16].Value;
                }

            }
            xlWorksheet.Cells[y + 1, 12] = kwPl1GivenAwayKG; // KW Packline 1 
            xlWorksheet.Cells[y + 2, 12] = chPl1GivenAwayKG; // CH Packline 1
            xlWorksheet.Cells[y + 3, 12] = chPl2GivenAwayKG; // CH Packline 2
            xlWorksheet.Cells[y + 4, 12] = chPl3GivenAwayKG; // CH Packline 3
            //Given away KG End


            //Staff Required Start
            int kwPl1StaffRequired = 0;
            int chPl1StaffRequired = 0;
            int chPl2StaffRequired = 0;
            int chPl3StaffRequired = 0;
            for (int rowN = 2; rowN < xlRange.Rows.Count + 1; rowN++)
            {
                string compareForRuns = xlWorksheet.Cells[rowN, 3].Value.ToString();

                if (compareForRuns.Equals("KW Packline - 1"))
                {
                    kwPl1StaffRequired += xlWorksheet.Cells[rowN, 22].Value;
                }
                if (compareForRuns.Equals("CH Packline - 1"))
                {
                    chPl1StaffRequired += xlWorksheet.Cells[rowN, 22].Value;
                }
                if (compareForRuns.Equals("CH Packline - 2"))
                {
                    chPl2StaffRequired += xlWorksheet.Cells[rowN, 22].Value;
                }
                if (compareForRuns.Equals("CH Packline - 3"))
                {
                    chPl3StaffRequired += xlWorksheet.Cells[rowN, 22].Value;
                }

            }
            xlWorksheet.Cells[y + 1, 16] = kwPl1StaffRequired; // KW Packline 1 
            xlWorksheet.Cells[y + 2, 16] = chPl1StaffRequired; // CH Packline 1
            xlWorksheet.Cells[y + 3, 16] = chPl2StaffRequired; // CH Packline 2
            xlWorksheet.Cells[y + 4, 16] = chPl3StaffRequired; // CH Packline 3
                                                               //Staff RequiActual
            //Staff Actual Start
            int kwPl1StaffActual = 0;
            int chPl1StaffActual = 0;
            int chPl2StaffActual = 0;
            int chPl3StaffActual = 0;
            for (int rowN = 2; rowN < xlRange.Rows.Count + 1; rowN++)
            {
                string compareForRuns = xlWorksheet.Cells[rowN, 3].Value.ToString();

                if (compareForRuns.Equals("KW Packline - 1"))
                {
                    kwPl1StaffActual += xlWorksheet.Cells[rowN, 23].Value;
                }
                if (compareForRuns.Equals("CH Packline - 1"))
                {
                    chPl1StaffActual += xlWorksheet.Cells[rowN, 23].Value;
                }
                if (compareForRuns.Equals("CH Packline - 2"))
                {
                    chPl2StaffActual += xlWorksheet.Cells[rowN, 23].Value;
                }
                if (compareForRuns.Equals("CH Packline - 3"))
                {
                    chPl3StaffActual += xlWorksheet.Cells[rowN, 23].Value;
                }

            }
            xlWorksheet.Cells[y + 1, 17] = kwPl1StaffActual; // KW Packline 1 
            xlWorksheet.Cells[y + 2, 17] = chPl1StaffActual; // CH Packline 1
            xlWorksheet.Cells[y + 3, 17] = chPl2StaffActual; // CH Packline 2
            xlWorksheet.Cells[y + 4, 17] = chPl3StaffActual; // CH Packline 3
            //Staff Actual End


                    // Total Calculation 


            xlWorksheet.Cells[y + 5, 6].Font.Bold = true;
            xlWorksheet.Cells[y + 5, 6].Font.Size = 12;
            xlWorksheet.Cells[y + 5, 7].Font.Bold = true;
            xlWorksheet.Cells[y + 5, 7].Font.Size = 12;
            xlWorksheet.Cells[y + 5, 8].Font.Bold = true;
            xlWorksheet.Cells[y + 5, 8].Font.Size = 12;
            xlWorksheet.Cells[y + 5, 6] = Convert.ToInt32(xlWorksheet.Cells[y+1, 6].Value)+ Convert.ToInt32(xlWorksheet.Cells[y + 2, 6].Value)+ Convert.ToInt32(xlWorksheet.Cells[y + 3, 6].Value)+Convert.ToInt32(xlWorksheet.Cells[y + 4, 6].Value); // Totoal Run Time for all Runs
            xlWorksheet.Cells[y + 5, 7] = "=SUM(H2:H" + x + ")"; // Totoal Run Time for all Runs
            xlWorksheet.Cells[y + 5, 8] = "=SUM(I2:I" + x + ")"; // Totoal Idle Time for all Stops

            xlWorksheet.Cells[y + 5, 10].Font.Bold = true;
            xlWorksheet.Cells[y + 5, 10].Font.Size = 12;
            xlWorksheet.Cells[y + 5, 10] = "=SUM(J2:J" + x + ")"; // Totoal Count all Runs

            xlWorksheet.Cells[y + 5, 11].Font.Bold = true;
            xlWorksheet.Cells[y + 5, 11].Font.Size = 12;
            xlWorksheet.Cells[y + 5, 11] = "=SUM(K2:K" + x + ")"; // Totoal Count all Runs

            xlWorksheet.Cells[y + 5, 12].Font.Bold = true;
            xlWorksheet.Cells[y + 5, 12].Font.Size = 12;
            xlWorksheet.Cells[y + 5, 12] = "=SUM(P2:P" + x + ")"; // Totoal Given away KG all Runs

            xlWorksheet.Cells[y + 5, 16].Font.Size = 12;
            xlWorksheet.Cells[y + 5, 16] = "=SUM(V2:V" + x + ")"; // Totoal Staff Required all Runs
            xlWorksheet.Cells[y + 5, 16].Font.Bold = true;

            xlWorksheet.Cells[y + 5, 17].Font.Size = 12;
            xlWorksheet.Cells[y + 5, 17] = "=SUM(W2:W" + x + ")"; // Totoal Staff Actual all Runs
            xlWorksheet.Cells[y + 5, 17].Font.Bold = true;


                    // Calculate total Give Away % and Running Efficiency % start
                    double totalGiveAwayPercentage = 0.0;
                    double totalRunningEfficiency = 0.0;
                    double totalExpectedEfficiency = 0.0;
                    double totalProductivityPercentage = 0.0;

                    int totalCount = 0;
                    int addToTotalCount = 0;
                    for (int rowX = 2; rowX < x + 1; rowX++)
                    {
                        if (xlWorksheet.Cells[rowX, 10].Value >= xlWorksheet.Cells[rowX, 11].Value)
                        {
                            addToTotalCount = Convert.ToInt32(xlWorksheet.Cells[rowX, 10].Value);
                        }
                        else
                        {
                            addToTotalCount = Convert.ToInt32(xlWorksheet.Cells[rowX, 11].Value);
                        }
                        totalCount += addToTotalCount;

                    }

                    for (int rowX = 2; rowX <= x; rowX++)
                    {
                        if (xlWorksheet.Cells[rowX, 10].Value >= xlWorksheet.Cells[rowX, 11].Value)
                        {
                            // Total Given Away %
                            // Total Running Efficiency %
                            totalGiveAwayPercentage += (((xlWorksheet.Cells[rowX, 10].Value * 100.0) / totalCount) * 0.01) * xlWorksheet.Cells[rowX, 17].Value; //   (((count * 100) / totalCount) * 0.01 )* giveAway%
                            totalRunningEfficiency += (((xlWorksheet.Cells[rowX, 10].Value * 100.0) / totalCount) * 0.01) * xlWorksheet.Cells[rowX, 20].Value; //   (((count * 100) / totalCount) * 0.01 )* runninEfficiency%
                            totalExpectedEfficiency += (((xlWorksheet.Cells[rowX, 10].Value * 100.0) / totalCount) * 0.01) * xlWorksheet.Cells[rowX, 21].Value; //   (((count * 100) / totalCount) * 0.01 )* expectedEfficiency%
                            totalProductivityPercentage += (((xlWorksheet.Cells[rowX, 10].Value * 100.0) / totalCount) * 0.01) * xlWorksheet.Cells[rowX, 24].Value; //   (((count * 100) / totalCount) * 0.01 )* productivityPercentage%
                        }
                        else // If Adjusted is primary count
                        {
                            // Total Given Away %
                            // Total Running Efficiency %
                            totalGiveAwayPercentage += (((xlWorksheet.Cells[rowX, 11].Value * 100.0) / totalCount) * 0.01) * xlWorksheet.Cells[rowX, 17].Value; //   (((count * 100) / totalCount) * 0.01 )* giveAway%
                            totalRunningEfficiency += (((xlWorksheet.Cells[rowX, 11].Value * 100.0) / totalCount) * 0.01) * xlWorksheet.Cells[rowX, 20].Value; //   (((count * 100) / totalCount) * 0.01 )* runninEfficiency%
                            totalExpectedEfficiency += (((xlWorksheet.Cells[rowX, 11].Value * 100.0) / totalCount) * 0.01) * xlWorksheet.Cells[rowX, 21].Value; //   (((count * 100) / totalCount) * 0.01 )* expectedEfficiency%
                            totalProductivityPercentage += (((xlWorksheet.Cells[rowX, 11].Value * 100.0) / totalCount) * 0.01) * xlWorksheet.Cells[rowX, 24].Value; //   (((count * 100) / totalCount) * 0.01 )* productivityPercentage%
                        }
                    }


                    xlWorksheet.Cells[y + 5, 9].Font.Bold = true;
                    xlWorksheet.Cells[y + 5, 9].Font.Size = 12;
                    xlWorksheet.Cells[y + 5, 9] = totalProductivityPercentage.ToString("0.00 %"); // Totoal Productivity % all Runs

                    xlWorksheet.Cells[y + 5, 13].Font.Bold = true;
                    xlWorksheet.Cells[y + 5, 13].Font.Size = 12;
                    xlWorksheet.Cells[y + 5, 13] = totalGiveAwayPercentage.ToString("0.00 %"); // Totoal Given away % all Runs

                    xlWorksheet.Cells[y + 5, 14].Font.Bold = true;
                    xlWorksheet.Cells[y + 5, 14].Font.Size = 12;
                    xlWorksheet.Cells[y + 5, 14] = totalRunningEfficiency.ToString("0.00 %"); // Totoal Running Efficiency % all Runs

                    xlWorksheet.Cells[y + 5, 15].Font.Bold = true;
                    xlWorksheet.Cells[y + 5, 15].Font.Size = 12;
                    xlWorksheet.Cells[y + 5, 15] = totalExpectedEfficiency.ToString("0.00 %"); // Totoal Expected Efficiency % all Runs
                    // Calculate total Give Away % and Running Efficiency % End

                    // Calculate total Give Away % and total Running Efficiency % for KW Packline 1 start
                    totalGiveAwayPercentage = 0.0;
                    totalRunningEfficiency = 0.0;
                    totalExpectedEfficiency = 0.0;
                    totalProductivityPercentage = 0.0;
                    totalCount = 0;
                    addToTotalCount = 0;
                    for (int rowX = 2; rowX < x + 1; rowX++)
                    {
                        if (xlWorksheet.Cells[rowX, 3].Value.ToString().Equals("KW Packline - 1"))
                        {
                            if (xlWorksheet.Cells[rowX, 10].Value >= xlWorksheet.Cells[rowX, 11].Value)
                            {
                                addToTotalCount = Convert.ToInt32(xlWorksheet.Cells[rowX, 10].Value);
                            }
                            else
                            {
                                addToTotalCount = Convert.ToInt32(xlWorksheet.Cells[rowX, 11].Value);
                            }
                            totalCount += addToTotalCount;
                        }
                    }

                    for (int rowX = 2; rowX <= x; rowX++)
                    {
                        if (xlWorksheet.Cells[rowX, 3].Value.ToString().Equals("KW Packline - 1"))
                        {
                            if (xlWorksheet.Cells[rowX, 10].Value >= xlWorksheet.Cells[rowX, 11].Value)
                            {
                                totalGiveAwayPercentage += (((xlWorksheet.Cells[rowX, 10].Value * 100.0) / totalCount) * 0.01) * xlWorksheet.Cells[rowX, 17].Value; //   (((count * 100) / totalCount) * 0.01 )* giveAway%
                                totalRunningEfficiency += (((xlWorksheet.Cells[rowX, 10].Value * 100.0) / totalCount) * 0.01) * xlWorksheet.Cells[rowX, 20].Value;  //   (((count * 100) / totalCount) * 0.01 )* runninEfficiency%
                                totalExpectedEfficiency += (((xlWorksheet.Cells[rowX, 10].Value * 100.0) / totalCount) * 0.01) * xlWorksheet.Cells[rowX, 21].Value;  //   (((count * 100) / totalCount) * 0.01 )* expectedEfficiency%
                                totalProductivityPercentage += (((xlWorksheet.Cells[rowX, 10].Value * 100.0) / totalCount) * 0.01) * xlWorksheet.Cells[rowX, 24].Value;  //   (((count * 100) / totalCount) * 0.01 )* productivityPercentage%
                            }
                            else // If Adjusted is primary count
                            {
                                totalGiveAwayPercentage += (((xlWorksheet.Cells[rowX, 11].Value * 100.0) / totalCount) * 0.01) * xlWorksheet.Cells[rowX, 17].Value; //   (((count * 100) / totalCount) * 0.01 )* giveAway%
                                totalRunningEfficiency += (((xlWorksheet.Cells[rowX, 11].Value * 100.0) / totalCount) * 0.01) * xlWorksheet.Cells[rowX, 20].Value; //   (((count * 100) / totalCount) * 0.01 )* runninEfficiency%
                                totalExpectedEfficiency += (((xlWorksheet.Cells[rowX, 11].Value * 100.0) / totalCount) * 0.01) * xlWorksheet.Cells[rowX, 21].Value; //   (((count * 100) / totalCount) * 0.01 )* expectedEfficiency%
                                totalProductivityPercentage += (((xlWorksheet.Cells[rowX, 11].Value * 100.0) / totalCount) * 0.01) * xlWorksheet.Cells[rowX, 24].Value;  //   (((count * 100) / totalCount) * 0.01 )* productivityPercentage%
                            }
                        }
                    }

                    xlWorksheet.Cells[y + 1, 9] = totalProductivityPercentage.ToString("0.00 %"); // Totoal productivity % all Runs
                    xlWorksheet.Cells[y + 1, 13] = totalGiveAwayPercentage.ToString("0.00 %"); // Totoal Given away % all Runs
                    xlWorksheet.Cells[y + 1, 14] = totalRunningEfficiency.ToString("0.00 %"); // Totoal Running Efficiency % all Runs
                    xlWorksheet.Cells[y + 1, 15] = totalExpectedEfficiency.ToString("0.00 %"); // Totoal Expected Efficiency % all Runs

                    // Calculate total Give Away % and total Running Efficiency % for CH Packline 1 start
                    totalGiveAwayPercentage = 0.0;
                    totalRunningEfficiency = 0.0;
                    totalExpectedEfficiency = 0.0;
                    totalProductivityPercentage = 0.0;
                    totalCount = 0;
                    addToTotalCount = 0;
                    for (int rowX = 2; rowX < x + 1; rowX++)
                    {
                        if (xlWorksheet.Cells[rowX, 3].Value.ToString().Equals("CH Packline - 1"))
                        {
                            if (xlWorksheet.Cells[rowX, 10].Value >= xlWorksheet.Cells[rowX, 11].Value)
                            {
                                addToTotalCount = Convert.ToInt32(xlWorksheet.Cells[rowX, 10].Value);
                            }
                            else
                            {
                                addToTotalCount = Convert.ToInt32(xlWorksheet.Cells[rowX, 11].Value);
                            }
                            totalCount += addToTotalCount;
                        }
                    }

                    for (int rowX = 2; rowX <= x; rowX++)
                    {
                        if (xlWorksheet.Cells[rowX, 3].Value.ToString().Equals("CH Packline - 1"))
                        {
                            if (xlWorksheet.Cells[rowX, 10].Value >= xlWorksheet.Cells[rowX, 11].Value)
                            {
                                totalGiveAwayPercentage += (((xlWorksheet.Cells[rowX, 10].Value * 100.0) / totalCount) * 0.01) * xlWorksheet.Cells[rowX, 17].Value; //   (((count * 100) / totalCount) * 0.01 )* giveAway%
                                totalRunningEfficiency += (((xlWorksheet.Cells[rowX, 10].Value * 100.0) / totalCount) * 0.01) * xlWorksheet.Cells[rowX, 20].Value;  //   (((count * 100) / totalCount) * 0.01 )* runninEfficiency%
                                totalExpectedEfficiency += (((xlWorksheet.Cells[rowX, 10].Value * 100.0) / totalCount) * 0.01) * xlWorksheet.Cells[rowX, 21].Value;  //   (((count * 100) / totalCount) * 0.01 )* expectedEfficiency%
                                totalProductivityPercentage += (((xlWorksheet.Cells[rowX, 10].Value * 100.0) / totalCount) * 0.01) * xlWorksheet.Cells[rowX, 24].Value;  //   (((count * 100) / totalCount) * 0.01 )* productivityPercentage%
                            }
                            else // If Adjusted is primary count
                            {
                                totalGiveAwayPercentage += (((xlWorksheet.Cells[rowX, 11].Value * 100.0) / totalCount) * 0.01) * xlWorksheet.Cells[rowX, 17].Value; //   (((count * 100) / totalCount) * 0.01 )* giveAway%
                                totalRunningEfficiency += (((xlWorksheet.Cells[rowX, 11].Value * 100.0) / totalCount) * 0.01) * xlWorksheet.Cells[rowX, 20].Value; //   (((count * 100) / totalCount) * 0.01 )* runninEfficiency%
                                totalExpectedEfficiency += (((xlWorksheet.Cells[rowX, 11].Value * 100.0) / totalCount) * 0.01) * xlWorksheet.Cells[rowX, 21].Value; //   (((count * 100) / totalCount) * 0.01 )* expectedEfficiency%
                                totalProductivityPercentage += (((xlWorksheet.Cells[rowX, 11].Value * 100.0) / totalCount) * 0.01) * xlWorksheet.Cells[rowX, 24].Value;  //   (((count * 100) / totalCount) * 0.01 )* productivityPercentage%
                            }
                        }
                    }
                    xlWorksheet.Cells[y + 2, 9] = totalProductivityPercentage.ToString("0.00 %"); // Totoal productivity % all Runs
                    xlWorksheet.Cells[y + 2, 13] = totalGiveAwayPercentage.ToString("0.00 %"); // Totoal Given away % all Runs
                    xlWorksheet.Cells[y + 2, 14] = totalRunningEfficiency.ToString("0.00 %"); // Totoal Running Efficiency % all Runs
                    xlWorksheet.Cells[y + 2, 15] = totalExpectedEfficiency.ToString("0.00 %"); // Totoal Expected Efficiency % all Runs


                    // Calculate total Give Away % and total Running Efficiency % for CH Packline 2 start
                    totalGiveAwayPercentage = 0.0;
                    totalRunningEfficiency = 0.0;
                    totalExpectedEfficiency = 0.0;
                    totalProductivityPercentage = 0.0;
                    totalCount = 0;
                    addToTotalCount = 0;
                    for (int rowX = 2; rowX < x + 1; rowX++)
                    {
                        if (xlWorksheet.Cells[rowX, 3].Value.ToString().Equals("CH Packline - 2"))
                        {
                            if (xlWorksheet.Cells[rowX, 10].Value >= xlWorksheet.Cells[rowX, 11].Value)
                            {
                                addToTotalCount = Convert.ToInt32(xlWorksheet.Cells[rowX, 10].Value);
                            }
                            else
                            {
                                addToTotalCount = Convert.ToInt32(xlWorksheet.Cells[rowX, 11].Value);
                            }
                            totalCount += addToTotalCount;
                        }
                    }

                    for (int rowX = 2; rowX <= x; rowX++)
                    {
                        if (xlWorksheet.Cells[rowX, 3].Value.ToString().Equals("CH Packline - 2"))
                        {
                            if (xlWorksheet.Cells[rowX, 10].Value >= xlWorksheet.Cells[rowX, 11].Value)
                            {
                                totalGiveAwayPercentage += (((xlWorksheet.Cells[rowX, 10].Value * 100.0) / totalCount) * 0.01) * xlWorksheet.Cells[rowX, 17].Value; //   (((count * 100) / totalCount) * 0.01 )* giveAway%
                                totalRunningEfficiency += (((xlWorksheet.Cells[rowX, 10].Value * 100.0) / totalCount) * 0.01) * xlWorksheet.Cells[rowX, 20].Value;  //   (((count * 100) / totalCount) * 0.01 )* runninEfficiency%
                                totalExpectedEfficiency += (((xlWorksheet.Cells[rowX, 10].Value * 100.0) / totalCount) * 0.01) * xlWorksheet.Cells[rowX, 21].Value;  //   (((count * 100) / totalCount) * 0.01 )* expectedEfficiency%
                                totalProductivityPercentage += (((xlWorksheet.Cells[rowX, 10].Value * 100.0) / totalCount) * 0.01) * xlWorksheet.Cells[rowX, 24].Value;  //   (((count * 100) / totalCount) * 0.01 )* productivityPercentage%
                            }
                            else // If Adjusted is primary count
                            {
                                totalGiveAwayPercentage += (((xlWorksheet.Cells[rowX, 11].Value * 100.0) / totalCount) * 0.01) * xlWorksheet.Cells[rowX, 17].Value; //   (((count * 100) / totalCount) * 0.01 )* giveAway%
                                totalRunningEfficiency += (((xlWorksheet.Cells[rowX, 11].Value * 100.0) / totalCount) * 0.01) * xlWorksheet.Cells[rowX, 20].Value; //   (((count * 100) / totalCount) * 0.01 )* runninEfficiency%
                                totalExpectedEfficiency += (((xlWorksheet.Cells[rowX, 11].Value * 100.0) / totalCount) * 0.01) * xlWorksheet.Cells[rowX, 21].Value; //   (((count * 100) / totalCount) * 0.01 )* expectedEfficiency%
                                totalProductivityPercentage += (((xlWorksheet.Cells[rowX, 11].Value * 100.0) / totalCount) * 0.01) * xlWorksheet.Cells[rowX, 24].Value;  //   (((count * 100) / totalCount) * 0.01 )* productivityPercentage%
                            }
                        }
                    }
                    xlWorksheet.Cells[y + 3, 9] = totalProductivityPercentage.ToString("0.00 %"); // Totoal productivity % all Runs
                    xlWorksheet.Cells[y + 3, 13] = totalGiveAwayPercentage.ToString("0.00 %"); // Totoal Given away % all Runs
                    xlWorksheet.Cells[y + 3, 14] = totalRunningEfficiency.ToString("0.00 %"); // Totoal Running Efficiency % all Runs
                    xlWorksheet.Cells[y + 3, 15] = totalExpectedEfficiency.ToString("0.00 %"); // Totoal Expected Efficiency % all Runs

                    // Calculate total Give Away % and total Running Efficiency % for CH Packline 3 start
                    totalGiveAwayPercentage = 0.0;
                    totalRunningEfficiency = 0.0;
                    totalExpectedEfficiency = 0.0;
                    totalProductivityPercentage = 0.0;
                    totalCount = 0;
                    addToTotalCount = 0;
                    for (int rowX = 2; rowX < x + 1; rowX++)
                    {
                        if (xlWorksheet.Cells[rowX, 3].Value.ToString().Equals("CH Packline - 3"))
                        {
                            if (xlWorksheet.Cells[rowX, 10].Value >= xlWorksheet.Cells[rowX, 11].Value)
                            {
                                addToTotalCount = Convert.ToInt32(xlWorksheet.Cells[rowX, 10].Value);
                            }
                            else
                            {
                                addToTotalCount = Convert.ToInt32(xlWorksheet.Cells[rowX, 11].Value);
                            }
                            totalCount += addToTotalCount;
                        }
                    }

                    for (int rowX = 2; rowX <= x; rowX++)
                    {
                        if (xlWorksheet.Cells[rowX, 3].Value.ToString().Equals("CH Packline - 3"))
                        {
                            if (xlWorksheet.Cells[rowX, 10].Value >= xlWorksheet.Cells[rowX, 11].Value)
                            {
                                totalGiveAwayPercentage += (((xlWorksheet.Cells[rowX, 10].Value * 100.0) / totalCount) * 0.01) * xlWorksheet.Cells[rowX, 17].Value; //   (((count * 100) / totalCount) * 0.01 )* giveAway%
                                totalRunningEfficiency += (((xlWorksheet.Cells[rowX, 10].Value * 100.0) / totalCount) * 0.01) * xlWorksheet.Cells[rowX, 20].Value;  //   (((count * 100) / totalCount) * 0.01 )* runninEfficiency%
                                totalExpectedEfficiency += (((xlWorksheet.Cells[rowX, 10].Value * 100.0) / totalCount) * 0.01) * xlWorksheet.Cells[rowX, 21].Value;  //   (((count * 100) / totalCount) * 0.01 )* expectedEfficiency%
                                totalProductivityPercentage += (((xlWorksheet.Cells[rowX, 10].Value * 100.0) / totalCount) * 0.01) * xlWorksheet.Cells[rowX, 24].Value;  //   (((count * 100) / totalCount) * 0.01 )* productivityPercentage%
                            }
                            else // If Adjusted is primary count
                            {
                                totalGiveAwayPercentage += (((xlWorksheet.Cells[rowX, 11].Value * 100.0) / totalCount) * 0.01) * xlWorksheet.Cells[rowX, 17].Value; //   (((count * 100) / totalCount) * 0.01 )* giveAway%
                                totalRunningEfficiency += (((xlWorksheet.Cells[rowX, 11].Value * 100.0) / totalCount) * 0.01) * xlWorksheet.Cells[rowX, 20].Value; //   (((count * 100) / totalCount) * 0.01 )* runninEfficiency%
                                totalExpectedEfficiency += (((xlWorksheet.Cells[rowX, 11].Value * 100.0) / totalCount) * 0.01) * xlWorksheet.Cells[rowX, 21].Value; //   (((count * 100) / totalCount) * 0.01 )* expectedEfficiency%
                                totalProductivityPercentage += (((xlWorksheet.Cells[rowX, 11].Value * 100.0) / totalCount) * 0.01) * xlWorksheet.Cells[rowX, 24].Value;  //   (((count * 100) / totalCount) * 0.01 )* productivityPercentage%
                            }
                        }
                    }
                    xlWorksheet.Cells[y + 4, 9] = totalProductivityPercentage.ToString("0.00 %"); // Totoal productivity % all Runs
                    xlWorksheet.Cells[y + 4, 13] = totalGiveAwayPercentage.ToString("0.00 %"); // Totoal Given away % all Runs
                    xlWorksheet.Cells[y + 4, 14] = totalRunningEfficiency.ToString("0.00 %"); // Totoal Running Efficiency % all Runs
                    xlWorksheet.Cells[y + 4, 15] = totalExpectedEfficiency.ToString("0.00 %"); // Totoal Expected Efficiency % all Runs
                    // Calculate total Give Away % End




                    xlWorkbook.Save();
            xlApp.Quit();
            return true;
                    //MessageBox.Show("Report Has been generated", "Confirmation", MessageBoxButtons.OK, MessageBoxIcon.Information);
                }
            catch (Exception ex) // If Exeption thrown Send email to Tim with details and 
            {
                xlApp.Quit();
                return false;
                ThreadPool.QueueUserWorkItem(state => SendEmail.sendEmailToTim(shiftRun, "Exeption while trying to generate a Daily Report "+ ex));
            }

            }
            else // If File not Exist
            {
                return false;
                //MessageBox.Show("File not Exist, nothing to generate", "Message", MessageBoxButtons.OK, MessageBoxIcon.Error);
            }
        }
        // Add data to dayly Report File start
    }
}   
