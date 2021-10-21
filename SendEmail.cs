using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using Outlook = Microsoft.Office.Interop.Outlook;
using Office = Microsoft.Office.Core;
using System.Windows.Forms;
using System.Drawing;
using System.IO;
using System.Drawing.Imaging;
using System.Reflection;

namespace PortMainScaleTest
{
    public static class SendEmail
    {
        static int today = 0;
        static int yestarday = -1;
        static int tomorrow = 1;

        // Send HD support request if IT HELP button has pressed
        public static bool sendEmailToHD(ShiftRun shiftRun)
        {

            try
            {
                Outlook._Application _app = new Outlook.Application();
                Outlook.MailItem mail = (Outlook.MailItem)_app.CreateItem(Outlook.OlItemType.olMailItem);
                mail.To = "it_helpdesk@purdys.com";
                mail.CC = "tim_b@purdys.com";
                mail.Subject = "IT support request from " + shiftRun.Location + " Pack line - " + shiftRun.PackLineNumber;
                mail.Body = "User from " + shiftRun.Location+ " Pack line - " + shiftRun.PackLineNumber +" requested IT assistance for Scale Machine";
                mail.Importance = Outlook.OlImportance.olImportanceHigh;
                ((Outlook._MailItem)mail).Send();

                Logger.INFO("IT support request has been sent (sendEmailToHD)");

                MessageBox.Show("IT support request has been sent\nWe will help you ASAP", "Message", MessageBoxButtons.OK, MessageBoxIcon.Information);
            }
            catch (Exception ex)
            {
                Logger.ERROR("Exception thrown while sending IT Help Desk request (sendEmailToHD): " + ex);
                //MessageBox.Show(ex.Message, "Message", MessageBoxButtons.OK, MessageBoxIcon.Error);
            }

            return true;
        }

        // Send Email to Tim_b if something went wrong
        public static bool sendEmailToTim(ShiftRun shiftRun, string theReason)
        {
            try
            {
                Outlook._Application _app = new Outlook.Application();
                Outlook.MailItem mail = (Outlook.MailItem)_app.CreateItem(Outlook.OlItemType.olMailItem);
                mail.To = "tim_b@purdys.com";
                mail.Subject = "Issue with Scale App" + shiftRun.Location + " Pack line - " + shiftRun.PackLineNumber;
                mail.Body = theReason;
                mail.Importance = Outlook.OlImportance.olImportanceHigh;
                ((Outlook._MailItem)mail).Send();

                Logger.INFO("Email to Tim_B has been sent (sendEmailToTim)");
                //MessageBox.Show("IT support request has been sent\nWe will help you ASAP", "Message", MessageBoxButtons.OK, MessageBoxIcon.Information);
            }
            catch (Exception ex)
            {
                Logger.ERROR("Exception thrown while sending Email to Tim_B (sendEmailToTim) " + ex);
            }

            return true;
        }

        // Send Email about BarCode NOT matching
        public static bool sendEmailBarCodeNotMatching(ShiftRun shiftRun)
        {
            try
            {
                Outlook._Application _app = new Outlook.Application();
                Outlook.MailItem mail = (Outlook.MailItem)_app.CreateItem(Outlook.OlItemType.olMailItem);
                mail.To = shiftRun.barCodeEmailNotificationList;
                mail.CC = shiftRun.barCodeEmailNotificationListCC;
                mail.Subject = "Bar code at " + shiftRun.Location + " Pack line - " + shiftRun.PackLineNumber + " - ERROR!";

                mail.Body = "Please check Packaging at " + shiftRun.Location + " Pack Line - " + shiftRun.PackLineNumber +
                            "\n BabarCode not matching";


                mail.Importance = Outlook.OlImportance.olImportanceHigh;
                ((Outlook._MailItem)mail).Send();

                Logger.INFO("Email to Tim_B has been sent (sendEmailToTim)");
                //MessageBox.Show("IT support request has been sent\nWe will help you ASAP", "Message", MessageBoxButtons.OK, MessageBoxIcon.Information);
            }
            catch (Exception ex)
            {
                Logger.ERROR("Exception thrown while sending Email to Tim_B (sendEmailToTim) " + ex);
            }

            return true;
        }

        // Send email to QA if running average down below between Less and Target
        public static void sendEmailToQA(ShiftRun shiftRun)
        {
            try
            {
                string dateTime = DateTime.Now.ToString("MMMM dd yyyy  h mm tt");

                Outlook._Application _app = new Outlook.Application();
                Outlook.MailItem mail = (Outlook.MailItem)_app.CreateItem(Outlook.OlItemType.olMailItem);

                mail.To = "quality_assurance@purdys.com";
                mail.CC = "productionsupervisors@purdys.com; william_h@purdys.com; tim_b@purdys.com";

                //mail.To = "tim_b";

                string averageToDisplay = (0.0).ToString("0.0 g");

                if (shiftRun.AverageWeightDynamic != 0.0)
                {
                    averageToDisplay = (shiftRun.AverageWeightDynamic).ToString("0.0 g"); // If Dynamic average not 0 yet.
                }
                else
                {
                    averageToDisplay = (shiftRun.AverageWeight).ToString("0.0 g"); // If Dynamic average 0 yet. Show total average
                }


                mail.Subject = "ALERT | " + shiftRun.Location + " Pack line - " + shiftRun.PackLineNumber + " | Low Avg weight";
                mail.Body = dateTime +
                            "\n\nAlert from "+ shiftRun.Location + " Pack Line: " + shiftRun.PackLineNumber + 
                            "\n\nRunning SKU: " + shiftRun.Sku +

                            "\n\nWeight" +
                            "\nLess: " + shiftRun.LessWeight +
                            "\nTarget: " + shiftRun.TargetWeight +
                            "\nHeavy: " + shiftRun.HeavyWeight +
                            "\n\nCurrent count: " + (shiftRun.PlCount + shiftRun.ManualCount) +
                            //"\n\nAverage weight: " + (shiftRun.AverageWeight / (shiftRun.ManualCount + shiftRun.PlCount)).ToString("0.0 g");
                            "\n\nRunning Average weight: " + averageToDisplay + // Should shows a dynamic average if it's not 0 
                            "\n\nScreenshot attached";

                int screenLeft = SystemInformation.VirtualScreen.Left;
                int screenTop = SystemInformation.VirtualScreen.Top;
                int screenWidth = SystemInformation.VirtualScreen.Width;
                int screenHeight = SystemInformation.VirtualScreen.Height;

             // Take screen shot and attach to Email start
                using (Bitmap bitmap = new Bitmap(screenWidth, screenHeight))
                {
                    using (Graphics g = Graphics.FromImage(bitmap))
                    {
                        g.CopyFromScreen(screenLeft, screenTop, 0, 0, bitmap.Size);
                        //g.CopyFromScreen(new Point(0, 0), new Point(0, 0), Screen.PrimaryScreen.WorkingArea.Size);
                    }
                    bitmap.Save(@"\\hedgehog\Syteline\PacklineScaleData\Screen.jpeg", ImageFormat.Jpeg);
                }

                string[] path = { @"\\hedgehog\Syteline\PacklineScaleData\Screen.jpeg"};
               
                string fullPath = System.IO.Path.Combine(path);

                mail.Attachments.Add(fullPath, Outlook.OlAttachmentType.olByValue);
             // Take screen shot and attach to Email end

                mail.Importance = Outlook.OlImportance.olImportanceHigh;
                ((Outlook._MailItem)mail).Send();

                Logger.INFO("Email to QA has been sent (sendEmailToQA)");
            }
            catch (Exception ex)
            {
                Logger.ERROR("Exception thrown while sending Email to QA (sendEmailToQA)" + ex);
            }
        }

        // Send Help Desk Request about data Not been saved in proper place. Saved in Document on local machine.
        // Saved file has to be move manually from Documents to \\hedgehog\Syteline\PacklineScaleData\year\month
        public static void sendEmailToHDAboutBackUpData(ShiftRun shiftRun, string fileName, string fullPath)
        {
            try
            {
                Outlook._Application _app = new Outlook.Application();
                Outlook.MailItem mail = (Outlook.MailItem)_app.CreateItem(Outlook.OlItemType.olMailItem);
                mail.To = "it_helpdesk@purdys.com";
                mail.CC = "tim_b@purdys.com";
                mail.Subject = "Action REQUIRED - request from " + shiftRun.Location + " Pack line - " + shiftRun.PackLineNumber;
                mail.Body = "File:  " + fileName +
                    "\n\nNOT been saved in "+fullPath+"  due to access issue" +
                    "\n(possibly password for scaleuser1 changed and no network access to this folder)" +
                    "\nFile saved in Documents folder on the local machine" +
                    "\n\nPlease manually move this file to " + fullPath +
                    "\n\nThank you";
                mail.Importance = Outlook.OlImportance.olImportanceNormal;
                ((Outlook._MailItem)mail).Send();

                Logger.INFO("Send Help Desk Request about data Not been saved in proper place (sendEmailToHDAboutBackUpData)");
            }
            catch (Exception ex)
            {
                Logger.ERROR("Exception thrown while sending Email about data Not been saved in proper place (sendEmailToHDAboutBackUpData)" + ex);
            }
        }

        // Missing ExcelLibrary.dll  Sending HD request
        public static void sendEmailToHDAboutMissingDLL(ShiftRun shiftRun)
        {
            try
            {
                Outlook._Application _app = new Outlook.Application();
                Outlook.MailItem mail = (Outlook.MailItem)_app.CreateItem(Outlook.OlItemType.olMailItem);
                mail.To = "it_helpdesk@purdys.com";
                mail.CC = "tim_b@purdys.com";
                mail.Subject = "Action REQUIRED - request from " + shiftRun.Location + " Pack line - " + shiftRun.PackLineNumber;
                mail.Body = "ExcelLibrary.dll is missing!" +
                    "\n\nPlease copy this file ExcelLibrary.dll FROM: \\\\hedgehog\\Syteline\\PacklineScaleData\\ScaleProject\\ScaleApp(Tim)" +
                    "\nTO: Desktop or to where the Scale application located" +
                    "\n\nIt requires for data saving" +
                    "\n\nThank you";
                mail.Importance = Outlook.OlImportance.olImportanceHigh;
                ((Outlook._MailItem)mail).Send();

                Logger.INFO("Email has been sent about Missing ExcelLibrary.dll  Sending HD request (sendEmailToHDAboutMissingDLL)");
            }
            catch (Exception ex)
            {
                Logger.ERROR("Exception thrown while sending Email about Missing ExcelLibrary.dll (sendEmailToHDAboutMissingDLL)" + ex);
            }
        }


        // Send Daily Report at 6am start
        public static void sendDailyReport(ShiftRun shiftRun)
        {
            // File name generating
            string dateTime;
            if (shiftRun.timsTesting)
            {
                dateTime = DateTime.Now.AddDays(tomorrow).ToString("MMMM dd yyyy"); // Check the file that created Tomorrow as it's testing mode
            }
            else
            {
                dateTime = DateTime.Now.AddDays(yestarday).ToString("MMMM dd yyyy"); // Check the file that created Yesterday
            }

            string fileName = dateTime + "_DAILY_REPORT" + ".xlsx";
            string[] path = { @"\\hedgehog\Syteline\PacklineScaleData\ScaleProject\ScaleApp(Tim)\TEMP\", fileName }; // Temporary store at this Location till end of the day

            string fileNameFullPath = Path.Combine(path);

            // For moving File
            string year = DateTime.Now.Year.ToString();
            string month = String.Format("{0:MMMM}", DateTime.Now);
            //string pathToSave = @"\\hedgehog\Syteline\PacklineScaleData\", "year", "\TestFile.xlsx";
            string[] path2 = { @"\\hedgehog\Syteline\PacklineScaleData\Daily_Reports\", year, fileName };
            string fileNameFullPathDistination = Path.Combine(path2);
            //MessageBox.Show(fileNameFullPath, "Confirmation", MessageBoxButtons.OK, MessageBoxIcon.Information);


            FileInfo dailyReport = new FileInfo(fileNameFullPath);

            if (dailyReport.Exists) // If File exist we can send it now
            {
                try
                {
                    Outlook._Application _app = new Outlook.Application();
                    Outlook.MailItem mail = (Outlook.MailItem)_app.CreateItem(Outlook.OlItemType.olMailItem);

                    if (!shiftRun.timsTesting) // If Not testing send a regular report
                    {
                        //mail.To = "elena_s@purdys.com";
                        mail.To = "productionsupervisors@purdys.com";
                        mail.CC = "william_h@purdys.com; tim_b@purdys.com";
                    }
                    else // If testing send to myself only
                    {
                        mail.To = "tim_b@purdys.com";
                    }
                    
                    mail.Subject = "PL Daily Report " + dateTime;
                    mail.Body = "PL Daily Report is attached" +
                        "\n\nAll reports are available here: \\\\hedgehog\\Syteline\\PacklineScaleData\\Daily_Reports";

                    mail.Attachments.Add(fileNameFullPath, Outlook.OlAttachmentType.olByValue); // Attach File and Send

                    mail.Importance = Outlook.OlImportance.olImportanceNormal;
                    ((Outlook._MailItem)mail).Send();

                    Logger.INFO("Daily Report has been sent (sendDailyReport)");

                    
                    // Move file after Sending Email if not testing mode
                    //if(!shiftRun.timsTesting)
                    File.Move(fileNameFullPath, fileNameFullPathDistination);
                    Logger.INFO("Daily Report file has been moved to: " + fileNameFullPathDistination +"  (sendDailyReport)");
                }
                catch (Exception ex)
                {
                    Logger.ERROR("Exception thrown while sending Daily Report (sendDailyReport)" + ex);
                }
            }
            else
            {
                Logger.INFO("Daily Report already has been sent from Different PL (sendDailyReport)");
            }
        }// Send Daily Report at 6am start
    }
}
