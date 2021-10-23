using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;
using System.IO.Ports;
using System.Threading;
using System.Runtime.Serialization;
using System.Runtime.InteropServices;

namespace PortMainScaleTest
{


    public partial class Settings : Form
    {

        //
        // Rounded corners for Form START
        //
        [DllImport("Gdi32.dll", EntryPoint = "CreateRoundRectRgn")]
        private static extern IntPtr CreateRoundRectRgn
        (
            int nLeftRect,     // x-coordinate of upper-left corner
            int nTopRect,      // y-coordinate of upper-left corner
            int nRightRect,    // x-coordinate of lower-right corner
            int nBottomRect,   // y-coordinate of lower-right corner
            int nWidthEllipse, // width of ellipse
            int nHeightEllipse // height of ellipse
        );
        //
        // Rounded corners for Form END
        //

        public string PLportNumber  { get; set; }
        public string ManualportNumber { get; set; }

        ShiftRun shiftRun;

        public static SerialPort ManualScaleSerialPort { get; set; }
        public static SerialPort PLScaleSerialPort { get; set; }

        public Settings(SerialPort serial, SerialPort serial2, ShiftRun shiftRun)
        {
            this.shiftRun = shiftRun;
            InitializeComponent();
            ManualScaleSerialPort = serial;
            PLScaleSerialPort = serial2;

            Region = System.Drawing.Region.FromHrgn(CreateRoundRectRgn(0, 0, Width, Height, 20, 20)); // Rounded corners for Form
        }

        private void Settings_Load(object sender, EventArgs e)
        {

            // Adding to combo box Available PL numbers
            comboBoxPLNumber.Items.Add(1);
            comboBoxPLNumber.Items.Add(2);
            comboBoxPLNumber.Items.Add(3);
            //comboBoxPLNumber.Items.Add(4);
            //comboBoxPLNumber.Items.Add(5);

            comboBoxCHKW.Items.Add("CH");
            comboBoxCHKW.Items.Add("KW");

            //Load form with saved settings for PackLine N
            comboBoxPLNumber.SelectedItem = Properties.Settings.Default.packLineNumber;
            comboBoxCHKW.SelectedItem = Properties.Settings.Default.locationCHKW;

            //Barcode checker Count and Minute
            comboBoxBarcodeCount.Items.Add(10);
            comboBoxBarcodeCount.Items.Add(30);
            comboBoxBarcodeCount.Items.Add(60);
            comboBoxBarcodeCount.Items.Add(100);
            comboBoxBarcodeCount.Items.Add(120);

            comboBoxBarcodeMinutesCount.Items.Add(15);
            comboBoxBarcodeMinutesCount.Items.Add(30);
            comboBoxBarcodeMinutesCount.Items.Add(45);
            comboBoxBarcodeMinutesCount.Items.Add(60);
            comboBoxBarcodeMinutesCount.Items.Add(120);


            checkBoxBarcode.Checked = Properties.Settings.Default.isBarcodeChecker;
            checkBoxIsEveryCount.Checked = Properties.Settings.Default.isCheckAtCount;
            checkBoxIsEveryMinute.Checked = Properties.Settings.Default.isCheckAtTime;

            comboBoxBarcodeCount.Text = Properties.Settings.Default.barcodeCheckerCount.ToString();
            comboBoxBarcodeMinutesCount.Text = Properties.Settings.Default.barCodeCheckEveryNumberMinutes.ToString();



            for (int i = 0; i < 23; i++)
            comboBoxHour.Items.Add(i);

            for (int i = 0; i < 60; i++)
            comboBoxMinute.Items.Add(i);


            PLScaleSerialPort.Close();
            ManualScaleSerialPort.Close();

            string[] ports = SerialPort.GetPortNames();
            comboBoxPL.Items.AddRange(ports);
            // comboBoxPL.SelectedIndex = 0;  Old settings
            comboBoxPL.Text = Properties.Settings.Default.PLCOMsettings;
            comboBoxManual.Items.AddRange(ports);
            comboBoxManual.Text = Properties.Settings.Default.ManualCOMsettings;

            // For auto report
            comboBoxHour.Text = Properties.Settings.Default.sendReportAtHour.ToString();
            comboBoxMinute.Text = Properties.Settings.Default.sendReportAtMinute.ToString();

            // Email To and CC
            richTextBoxEmailTo.Text = Properties.Settings.Default.barCodeEmailNotificationList;
            richTextBoxEmailToCC.Text = Properties.Settings.Default.barCodeEmailNotificationListCC;

            if (shiftRun.autoGenerateReport == true)
            {
                checkBoxAutoReport.Checked = true;
            }
            else
            {
                checkBoxAutoReport.Checked = false;
            }

            PLportNumber = Properties.Settings.Default.PLCOMsettings; // Load default settings when settings open
            ManualportNumber = Properties.Settings.Default.ManualCOMsettings; // Load default settings when settings open

            
            if (shiftRun.boxOverSize == true)
            {
                checkBoxSaveToNewFormat.Checked = true;
            }
            if (shiftRun.boxOverSize == false)
            {
                checkBoxSaveToNewFormat.Checked = false;
            }

            // For testing Buttons enable disable
            if (shiftRun.timsTesting == true)
            {
                checkBoxTimsTesting.Checked = true; 
            }
            else
            {
                checkBoxTimsTesting.Checked = false; 
            }

        }

        private void buttonSettingsOK_Click(object sender, EventArgs e)
        {


            // Save Application settings for Location and PL number



            try
            {
                shiftRun.PackLineNumber = Convert.ToInt32(comboBoxPLNumber.SelectedItem);
                Properties.Settings.Default.locationCHKW = comboBoxCHKW.Text;
                shiftRun.Location = comboBoxCHKW.Text;
                Properties.Settings.Default.packLineNumber = Convert.ToInt32(comboBoxPLNumber.Text);
                Properties.Settings.Default.barcodeCheckerCount = Convert.ToInt32(comboBoxBarcodeCount.Text);
                Properties.Settings.Default.barCodeCheckEveryNumberMinutes = Convert.ToInt32(comboBoxBarcodeMinutesCount.Text);

                shiftRun.barCodeCheckAtCount = Convert.ToInt32(comboBoxBarcodeCount.Text);
                shiftRun.barCodeCheckEveryNumberMinutes = Convert.ToInt32(comboBoxBarcodeMinutesCount.Text);
                shiftRun.isBarcodeChecker = checkBoxBarcode.Checked;

                shiftRun.barCodeEmailNotificationList = richTextBoxEmailTo.Text;
                shiftRun.barCodeEmailNotificationListCC = richTextBoxEmailToCC.Text;

                //
                // Next box will be checked at
                //
                    if (shiftRun.PlCount >= shiftRun.ManualCount)
                    {
                        shiftRun.nextCheckAt = shiftRun.PlCount + shiftRun.barCodeCheckAtCount;
                    }
                    else
                    {
                        shiftRun.nextCheckAt = shiftRun.ManualCount + shiftRun.barCodeCheckAtCount;
                    }
                //
                //
                //

                if (shiftRun.Location == "")
                {
                    throw new IllegalArgumentException();
                }

                DialogResult = DialogResult.OK;
            }
            catch (Exception ex) // Location or PL not Set 
            {
                MessageBox.Show("Location or PL number not selected", "Error!", MessageBoxButtons.OK, MessageBoxIcon.Error);
                DialogResult = DialogResult.OK; // Might be temporary
            }
            //Properties.Settings.Default.Save();

            // Save Application settings for ports and for Save format
            Properties.Settings.Default.PLCOMsettings = comboBoxPL.Text;
            Properties.Settings.Default.ManualCOMsettings = comboBoxManual.Text;
            Properties.Settings.Default.checkBoxSaveToNewFormat = checkBoxSaveToNewFormat.Checked;
            Properties.Settings.Default.isBarcodeChecker = checkBoxBarcode.Checked;
            Properties.Settings.Default.isCheckAtCount = checkBoxIsEveryCount.Checked;
            Properties.Settings.Default.isCheckAtTime = checkBoxIsEveryMinute.Checked;
            
            Properties.Settings.Default.barCodeEmailNotificationList = richTextBoxEmailTo.Text;
            Properties.Settings.Default.barCodeEmailNotificationListCC = richTextBoxEmailToCC.Text;
            Properties.Settings.Default.Save();
            
            PLportNumber = comboBoxPL.Text;
            ManualportNumber = comboBoxManual.Text;

            //For reporting
            Properties.Settings.Default.sendReportAtHour = Convert.ToInt32(comboBoxHour.Text);
            Properties.Settings.Default.sendReportAtMinute = Convert.ToInt32(comboBoxMinute.Text);

            if (checkBoxAutoReport.Checked == true)
            {
                Properties.Settings.Default.autoGenerateReport = true;
            }
            else
            {
                Properties.Settings.Default.autoGenerateReport = false;
            }
        }

    private void buttonSettingsCancel_Click(object sender, EventArgs e)
        {
            DialogResult = DialogResult.Cancel;
        }


        // check box to chose if we gonna save data to NEW format
        private void checkBox1_CheckedChanged(object sender, EventArgs e)
        {
            if (checkBoxSaveToNewFormat.Checked)
            {
                shiftRun.boxOverSize = true; // New format chosen
            }
            if (!checkBoxSaveToNewFormat.Checked)
            {
                shiftRun.boxOverSize = false; // Old format chosen
            }
        }

        private void checkBoxTimsTesting_CheckedChanged(object sender, EventArgs e)
        {
            if (checkBoxTimsTesting.Checked)
            {
                shiftRun.timsTesting = true; // New format chosen
            }
            if (!checkBoxTimsTesting.Checked)
            {
                shiftRun.timsTesting = false; // Old format chosen
            }
        }

        private void generateReport_Click(object sender, EventArgs e)
        {
            SaveData.generateDailyReportFile(shiftRun);
        }
        private void sendDailyReport_Click(object sender, EventArgs e)
        {
            ThreadPool.QueueUserWorkItem(state => SendEmail.sendDailyReport(shiftRun));
        }

        private void sendDailyReportAndSend_Click(object sender, EventArgs e)
        {
            shiftRun.isGenerateAndSend = true;

        }

        //BarCode checker OFF/ON
        private void checkBoxBarcode_CheckedChanged(object sender, EventArgs e)
        {
            if (checkBoxBarcode.Checked) {
                shiftRun.isBarcodeChecker = true; //BarCode checker is ON
            }
            if (!checkBoxBarcode.Checked)
            {
                shiftRun.isBarcodeChecker = false; //BarCode checker is OFF
            }
        }

        //Check Every N count OFF/ON
        private void checkBoxIsEveryCount_CheckedChanged(object sender, EventArgs e)
        {
            if (checkBoxIsEveryCount.Checked) {
                shiftRun.isCheckAtCount = true; // Check at count is ON
            }
            if (!checkBoxIsEveryCount.Checked)
            {
                shiftRun.isCheckAtCount = false; // Check at count is OFF
            }
        }

        //Check Every Minute OFF/ON
        private void checkBoxIsEveryMinute_CheckedChanged(object sender, EventArgs e)
        {
            if (checkBoxIsEveryMinute.Checked)
            {
                shiftRun.isCheckAtTime = true; // Check every Minute is ON
            }
            if (!checkBoxIsEveryMinute.Checked)
            {
                shiftRun.isCheckAtTime = false; // Check every Minute is OFF
            }
        }
    }

    [Serializable]
    internal class IllegalArgumentException : Exception
    {
        public IllegalArgumentException()
        {
        }

        public IllegalArgumentException(string message) : base(message)
        {
        }

        public IllegalArgumentException(string message, Exception innerException) : base(message, innerException)
        {
        }

        protected IllegalArgumentException(SerializationInfo info, StreamingContext context) : base(info, context)
        {
        }
    }
}
