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

namespace PortMainScaleTest
{


    public partial class Settings : Form
    {
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

            //Barcode checker
            comboBoxBarcodeCount.Items.Add(10);
            comboBoxBarcodeCount.Items.Add(30);
            comboBoxBarcodeCount.Items.Add(60);
            comboBoxBarcodeCount.Items.Add(100);
            comboBoxBarcodeCount.Items.Add(120);

            comboBoxBarcodeCountType.Items.Add("ea");
            comboBoxBarcodeCountType.Items.Add("min");

            checkBoxBarcode.Checked = Properties.Settings.Default.isBarcodeChecker;
            comboBoxBarcodeCount.Text = Properties.Settings.Default.barcodeCheckerCount.ToString();
            comboBoxBarcodeCountType.SelectedItem = Properties.Settings.Default.barcodeCheckerCountType;


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
            
            Properties.Settings.Default.barcodeCheckerCountType = comboBoxBarcodeCountType.Text;
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
