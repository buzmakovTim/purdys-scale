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
using System.IO;
using System.Threading;

//using Microsoft.Office.Interop.Excel;
using Excel = Microsoft.Office.Interop.Excel;

using ExcelLibrary.SpreadSheet; // In order to Save to Excel file
using System.Windows.Forms.DataVisualization.Charting;
using System.Security.Cryptography.X509Certificates;

namespace PortMainScaleTest
{
    public partial class Scale : Form
    {
        // For auto stop
        int stopInSeconds;

        //int countOneHour; // When set a new hour
        int currentHour;
        int numberOfErrorsWhenEmailSend;

        const int NUM_OF_HOURS = 12;
        

        // For Timer
        //bool isTimerActive = false;
        //int timeH;
        //int timeM;
        //int timeS;
        //int breakTimeH;
        //int breakTimeM;
        //int breakTimeS;

        // For moving Form
        int mov;
        int movX;
        int movY;

        public bool dataSaved;

        static int countPLScale; // First box count PL Scale
        static int countManualScale; // First box count MAnual Scale



        string[][] uphCollection = new string[NUM_OF_HOURS][]; // Create a 2D collection of UPH for each hour 12 MAX plus time stamp

        UnitsPerHour unitsPerHour = new UnitsPerHour(); // Units Per Hour object created
        ShiftRun shiftRun = new ShiftRun(); // ShiftRun object created
        DataValidator validate = new DataValidator();
        ManualScaleWeightCollection manualScaleWeightCollection = new ManualScaleWeightCollection();
        PLScaleWeightCollection pLScaleWeightCollection = new PLScaleWeightCollection();




        TimerRunning timerRunning = new TimerRunning(); // Timer Class TESTING

        static Queue<double> dynamicAverageCollection = new Queue<double>(); // Collection for Dynamic average

        public Scale()
        {
            InitializeComponent();
            ManualScaleSerialPort.DataReceived += new SerialDataReceivedEventHandler(ManualScale_DataReceived);
            PLScaleSerialPort.DataReceived += new SerialDataReceivedEventHandler(PLScale_DataReceived);
        }


        // Reset All data Start
        public void resetAllDataMembers()
        {
            //resetTimer(); // Reset timer
            //resetBreakTimer(); // Reset Break
            //isTimerActive = false; // Timer not running

            dataSaved = true; // Will set to false once START button pressed
            countPLScale = 1; // First box count PL Scale
            countManualScale = 1; // First box count Manual Scale

            // Set UPH collection to 0
            for (int i = 0; i < uphCollection.Length; i++)
            {
                uphCollection[i] = new string[2];
                uphCollection[i][0] = "";
                uphCollection[i][1] = "";
            }

            currentHour = 0;
            numberOfErrorsWhenEmailSend = 15; // Send email to Tim after 15 errors 

            shiftRun.DataFromPLScale = "";
            shiftRun.DataFromManualScale = "";
            shiftRun.PlCount = 0;
            shiftRun.StaffNumberRequired = 1;
            shiftRun.StaffNumberActual = 1;
            shiftRun.ManualCount = 0;
            shiftRun.Sku = "";
            shiftRun.Shift = "";
            shiftRun.LessWeight = 0;
            shiftRun.TargetWeight = 0;
            shiftRun.HeavyWeight = 0;
            shiftRun.PackLineNumber = Properties.Settings.Default.packLineNumber;
            shiftRun.Location = Properties.Settings.Default.locationCHKW;
            shiftRun.LessCount = 0;
            shiftRun.TargetCount = 0;
            shiftRun.HeavyCount = 0;
            shiftRun.AverageWeight = 0;
            shiftRun.AverageWeightDynamic = 0;
            shiftRun.AverageWeightCount = 0;
            shiftRun.AverageWeightAdjusted = 0;
            shiftRun.AverageWeightLess = 0;
            shiftRun.AverageWeightTarget = 0;
            shiftRun.AverageWeightHeavy = 0;
            shiftRun.runningTimeInSeconds = 0; // Running time in Minutes
            shiftRun.totalBreakTimeInSeconds = 0; // Break time in Minutes
            shiftRun.isBreak = false; // Is break set to false
            shiftRun.Running = false;
            shiftRun.timsTesting = false;
            shiftRun.saveToNEWformat = Properties.Settings.Default.checkBoxSaveToNewFormat;
            shiftRun.Warning = false;
            shiftRun.errorCountPL = 0; // Error count to 0
            shiftRun.errorCountPL = 0; // Error count to 0
            labelErrorPL.Text = shiftRun.errorCountPL.ToString(); // Label errors
            labelErrorManual.Text = shiftRun.errorCountPL.ToString(); // Label errors 
            labelStaffRequired.Text = "0";
            labelStaffActual.Text = "0";


            shiftRun.isDelaySaving = false;

            //Productivity target
            shiftRun.productivityRun = 0;
            shiftRun.productivityTarget = 0;
            shiftRun.productivityActual = 0;
            shiftRun.runningEfficiency = 0;
            shiftRun.expectedEfficiency = 0;
            labelProdTargetData.Text = shiftRun.productivityTarget.ToString("0");
            labelProdActualData.Text = shiftRun.productivityActual.ToString("0");
            labelUPHPercentage.Text = shiftRun.runningEfficiency.ToString("0.0 %"); // Show Running Efficiency Percentage
            expectedEfficiencyLabel.Text = shiftRun.errorCountManual.ToString("0.0 %"); // Show Ecxpected Efficiency Percentage

            labelProdTargetData.ForeColor = Color.Green;
            labelProdActualData.ForeColor = Color.White;
            labelUPHPercentage.ForeColor = Color.White;


            // Give awaya
            shiftRun.kgGivingAway = 0;
            shiftRun.percentageGivingAway = 0;
            labelGiveAwayData.ForeColor = Color.White;
            labelGiveAwayPercentage.ForeColor = Color.White;

            labelGiveAwayData.Text = (shiftRun.kgGivingAway * 0.001).ToString("0.0 kg"); // Show Giving away data
            labelGiveAwayPercentage.Text = shiftRun.percentageGivingAway.ToString("0.0 %"); // Show Giving away data in %

            shiftRun.emailToQAsent = false;
            

            labelSKUData.Text = shiftRun.Sku; // Default SKU 
            labelLessData.Text = shiftRun.LessWeight.ToString(); // Default Valuers for Weight
            labelTargetData.Text = shiftRun.TargetWeight.ToString(); // Default Valuers for Weight
            labelHeavyData.Text = shiftRun.HeavyWeight.ToString(); // Default Valuers for Weight

            labelCountLess.Text = shiftRun.LessCount.ToString(); // Count Less Default Value
            labelCountTarget.Text = shiftRun.TargetCount.ToString(); // Count Target Default Value
            labelCountHeavy.Text = shiftRun.HeavyCount.ToString(); // Count Heavy Default Value

            // Avg. Weight
            labelDynamicAvgWeight.ForeColor = Color.White;
            labelAverageWeight.ForeColor = Color.White;
            labelAverageWeight.Text = shiftRun.AverageWeight.ToString(); // Average weight Default value
            labelDynamicAvgWeight.Text = shiftRun.AverageWeightDynamic.ToString(); // Dynamic Average weight Default value
            labelAvgLess.Text = shiftRun.AverageWeightLess.ToString(); // Less Average weight Default value
            labelAvgTarget.Text = shiftRun.AverageWeightTarget.ToString(); // Target Average weight Default value
            labelAvgHeavy.Text = shiftRun.AverageWeightHeavy.ToString(); // Heavy Average weight Default value
            labelAverageCount.Text = shiftRun.AverageWeightCount.ToString(); // Count Average weight Default value
            labelAverageAdjusted.Text = shiftRun.AverageWeightAdjusted.ToString(); // Adjusted Average weight Default value

            labelAdjusted.Text = shiftRun.ManualCount.ToString(); // Starting Count For PL Scale Value 0
            labelCount.Text = shiftRun.PlCount.ToString(); // // Starting Count Adjusted Scale Value 0

            label_manualData.Text = string.Empty;
            label_plData.Text = string.Empty;

            // Oversize box
            shiftRun.boxOverSize = false; // For oversize boxes
            labelOversizeBox.Visible = false; // Message about Box oversize

            // Timer lable update
            labelRunningTimeS.Text = String.Format("{0:00}", 0);
            labelRunningTimeM.Text = String.Format("{0:00}.", 0);
            labelRunningTimeH.Text = String.Format("{0:00}:", 0);

            labelBreakTimeS.Text = String.Format("{0:00}", 0);
            labelBreakTimeM.Text = String.Format("{0:00}.", 0);
            labelBreakTimeH.Text = String.Format("{0:00}:", 0);

            // Show break time
            labelBreakTimeS.Visible = false;
            labelBreakTimeM.Visible = false;
            labelBreakTimeH.Visible = false;
            labelBreak.Visible = false;

            shiftRun.isTimerON = false;
            labelShift.Visible = false;
            //countOneHour = 0;

            // Buttons 
            buttonStart.Enabled = true;
            buttonStart.Text = "START";
            buttonStart.BackColor = Color.DarkGreen;
            buttonStop.Enabled = false;
            buttonStop.Text = "STOP";
            buttonStop.BackColor = Color.DarkGray;

            // Also settings has this
            buttonClear.Visible = false;
            buttontest.Visible = false;
            buttonSet.Visible = false;

            // Charts clear
            chartWeight.Series["Count"].Points.Clear();
            chartUPH.Series["UPH"].Points.Clear();

            labelPackLineNumberData.Text = shiftRun.PackLineNumber.ToString(); // PackLine when app starts
            labelCHorKW.Text = shiftRun.Location; // Location when app starts

            // For reporting
            shiftRun.autoGenerateReport = Properties.Settings.Default.autoGenerateReport; // Use preconfigured settings
            shiftRun.sendReportAtHour = Properties.Settings.Default.sendReportAtHour; // Use preconfigured settings
            shiftRun.sendReportAtMinute = Properties.Settings.Default.sendReportAtMinute; // Use preconfigured settings

            //Bar Code Checker
            labelBarcode.Text = "";
            shiftRun.isBarcodeChecker = Properties.Settings.Default.isBarcodeChecker; // by default it's off or Whatever settings saved
            shiftRun.barCode = "";              
            shiftRun.isBarCodeMatch = true;
            shiftRun.barCodeCheckAtCount = Properties.Settings.Default.barcodeCheckerCount;
            shiftRun.barCodeCountType = Properties.Settings.Default.barcodeCheckerCountType; // Ea. or Min
            shiftRun.nextCheckAt = shiftRun.barCodeCheckAtCount; // Originally checking at Default value
            shiftRun.barCodeEmailNotificationList = Properties.Settings.Default.barCodeEmailNotificationList; // Email List
            shiftRun.barCodeEmailNotificationListCC = Properties.Settings.Default.barCodeEmailNotificationListCC; // CC Email List
            //shiftRun.nextCheckAt = 10; // Originally checking at Default value

        }
        // Reset All data End

        public void Form1_Load(object sender, EventArgs e)
        {


            Logger.INFO("Application has been started");   // Logging App Started

            //timerRunning.startTimer(); // Start Timer
            timerTime.Start(); // Timer
            labelTime.Text = DateTime.Now.ToLongTimeString(); // Time on the right top corner
            labelDate.Text = DateTime.Now.ToString("MMM dd yyyy"); // Date on the right top corner

            // Creating new ShiftRun object

            checkExcelLibrary(); // Check if ExcelLibriry.dll Exist

            resetAllDataMembers(); // All data members with default values

            PLportN.Text = Properties.Settings.Default.PLCOMsettings;
            Manualport.Text = Properties.Settings.Default.ManualCOMsettings;

            PlConnected(false);
            ManualConnected(false);

            setBinding();
            setupDataGridViewPL();
            setupDataGridViewManual();


            try
            {
                PLScaleSerialPort.PortName = Properties.Settings.Default.PLCOMsettings; // Use saved settings for COM port
                PLScaleSerialPort.BaudRate = 9600;
                PLScaleSerialPort.Parity = Parity.None;
                PLScaleSerialPort.StopBits = StopBits.One;
                PLScaleSerialPort.DataBits = 8;
                PLScaleSerialPort.Open();
                PlConnected(true);

                // Warning Message
                shiftRun.Warning = false;
                warningMessage(shiftRun.Warning);
            }
            catch (Exception ex)
            {
                //MessageBox.Show(ex.Message, "Message", MessageBoxButtons.OK, MessageBoxIcon.Error);
                PlConnected(false);
                shiftRun.Warning = true;

                // Warning Message
                shiftRun.Warning = true;
                warningMessage(shiftRun.Warning);
                
                Logger.ERROR("Catch Block COM port for PL scale can't be open (FORM LOAD)");   // COM ports can't be open
            }

            try
            {
                ManualScaleSerialPort.PortName = Properties.Settings.Default.ManualCOMsettings;   // Use saved settings for COM port
                ManualScaleSerialPort.BaudRate = 9600;
                ManualScaleSerialPort.Parity = Parity.None;
                ManualScaleSerialPort.StopBits = StopBits.One;
                ManualScaleSerialPort.DataBits = 8;
                ManualScaleSerialPort.Open();
                ManualConnected(true);

                // Warning Message
                shiftRun.Warning = false;
                warningMessage(shiftRun.Warning);
            }
            catch (Exception ex)
            {
                //MessageBox.Show(ex.Message, "Message", MessageBoxButtons.OK, MessageBoxIcon.Error);
                ManualConnected(false);

                // Warning Message
                shiftRun.Warning = true;
                warningMessage(shiftRun.Warning);

                Logger.ERROR("Catch Block COM port for Manual scale can't be open (FORM LOAD)");   // COM ports can't be open
            }
        }

        // Binding
        public void setBinding()
        {
            dataGridViewManualScale.AutoGenerateColumns = false;
            dataGridViewPLScale.AutoGenerateColumns = false;


            dataGridViewManualScale.DataSource = manualScaleWeightCollection;  // Add data from collection to Data Grid View
            dataGridViewPLScale.DataSource = pLScaleWeightCollection;  // Add data from collection to Data Grid View

        }

        // Data Grid View for PL scale
        private void setupDataGridViewPL() // Data Grid View for Pack Line scale data
        {
            // configure for readonly 
            dataGridViewPLScale.SelectionMode = DataGridViewSelectionMode.FullRowSelect;
            dataGridViewPLScale.MultiSelect = false;
            dataGridViewPLScale.AllowUserToAddRows = false;
            dataGridViewPLScale.EditMode = DataGridViewEditMode.EditProgrammatically;
            dataGridViewPLScale.AllowUserToOrderColumns = false;
            dataGridViewPLScale.AllowUserToResizeColumns = false;
            dataGridViewPLScale.AllowUserToResizeRows = false;
            dataGridViewPLScale.ColumnHeadersDefaultCellStyle.Font = new Font(DataGridView.DefaultFont, FontStyle.Bold);

            //dataGridViewPLScale.AutoResizeColumns();

            // add columns Count
            DataGridViewTextBoxColumn count = new DataGridViewTextBoxColumn();
            count.Name = "Count";
            count.DataPropertyName = "Count";
            count.HeaderText = "Count";
            count.Width = 65;
            count.HeaderCell.Style.Alignment = DataGridViewContentAlignment.MiddleCenter;
            count.DefaultCellStyle.Alignment = DataGridViewContentAlignment.MiddleCenter;
            count.SortMode = DataGridViewColumnSortMode.NotSortable;
            dataGridViewPLScale.Columns.Add(count);

            // add columns Weight
            DataGridViewTextBoxColumn weight = new DataGridViewTextBoxColumn();
            weight.Name = "WeightPLScale";
            weight.DataPropertyName = "WeightPLScale";
            weight.HeaderText = "Weight";
            weight.Width = 80;
            weight.HeaderCell.Style.Alignment = DataGridViewContentAlignment.MiddleCenter;
            weight.DefaultCellStyle.Alignment = DataGridViewContentAlignment.MiddleCenter;
            weight.DefaultCellStyle.Format = "0.0 g";
            weight.SortMode = DataGridViewColumnSortMode.NotSortable;
            dataGridViewPLScale.Columns.Add(weight);


            // add columns DateTime
            DataGridViewTextBoxColumn dateTime = new DataGridViewTextBoxColumn();
            dateTime.Name = "DateTime";
            dateTime.DataPropertyName = "dateTime";
            dateTime.HeaderText = "Date and Time";
            //dateTime.Width = 160;
            dateTime.HeaderCell.Style.Alignment = DataGridViewContentAlignment.MiddleCenter;
            dateTime.DefaultCellStyle.Alignment = DataGridViewContentAlignment.MiddleCenter;
            dateTime.DefaultCellStyle.Format = "MMMM dd yyyy   h:mm:ss tt";
            dateTime.SortMode = DataGridViewColumnSortMode.NotSortable;
            dataGridViewPLScale.Columns.Add(dateTime);
            dataGridViewPLScale.Columns[2].AutoSizeMode = DataGridViewAutoSizeColumnMode.Fill; // Data to fill 100% wight

        }

        // Data Grid View for Manual scale
        private void setupDataGridViewManual() // Data Grid View for Manual scale data
        {
            // configure for readonly 
            dataGridViewManualScale.SelectionMode = DataGridViewSelectionMode.FullRowSelect;
            dataGridViewManualScale.MultiSelect = false;
            dataGridViewManualScale.AllowUserToAddRows = false;
            dataGridViewManualScale.EditMode = DataGridViewEditMode.EditProgrammatically;
            dataGridViewManualScale.AllowUserToOrderColumns = false;
            dataGridViewManualScale.AllowUserToResizeColumns = false;
            dataGridViewManualScale.AllowUserToResizeRows = false;
            dataGridViewManualScale.ColumnHeadersDefaultCellStyle.Font = new Font(DataGridView.DefaultFont, FontStyle.Bold);




            // add columns Count
            DataGridViewTextBoxColumn count = new DataGridViewTextBoxColumn();
            count.Name = "Count";
            count.DataPropertyName = "Count";
            count.HeaderText = "Count";
            count.Width = 65;
            count.HeaderCell.Style.Alignment = DataGridViewContentAlignment.MiddleCenter;
            count.DefaultCellStyle.Alignment = DataGridViewContentAlignment.MiddleCenter;
            count.SortMode = DataGridViewColumnSortMode.NotSortable;
            dataGridViewManualScale.Columns.Add(count);

            // add columns Weight
            DataGridViewTextBoxColumn weight = new DataGridViewTextBoxColumn();
            weight.Name = "WeightManualScale";
            weight.DataPropertyName = "WeightManualScale";
            weight.HeaderText = "Weight";
            weight.Width = 80;
            weight.HeaderCell.Style.Alignment = DataGridViewContentAlignment.MiddleCenter;
            weight.DefaultCellStyle.Alignment = DataGridViewContentAlignment.MiddleCenter;
            weight.DefaultCellStyle.Format = "0.0 g";
            weight.SortMode = DataGridViewColumnSortMode.NotSortable;
            dataGridViewManualScale.Columns.Add(weight);

            // add columns DateTime
            DataGridViewTextBoxColumn dateTime = new DataGridViewTextBoxColumn();
            dateTime.Name = "DateTime";
            dateTime.DataPropertyName = "dateTime";
            dateTime.HeaderText = "Date and Time";
            //dateTime.Width = 160;
            dateTime.HeaderCell.Style.Alignment = DataGridViewContentAlignment.MiddleCenter;
            dateTime.DefaultCellStyle.Alignment = DataGridViewContentAlignment.MiddleCenter;
            dateTime.DefaultCellStyle.Format = "MMMM dd yyyy   h:mm:ss tt";
            dateTime.SortMode = DataGridViewColumnSortMode.NotSortable;
            dataGridViewManualScale.Columns.Add(dateTime);
            dataGridViewManualScale.Columns[2].AutoSizeMode = DataGridViewAutoSizeColumnMode.Fill; // Data to fill 100% wight

        }

        // START button
        private void buttonOpen_Click(object sender, EventArgs e)
        {
            
            try
            {
                if (!PLScaleSerialPort.IsOpen) // Making sure Port is Open
                    PLScaleSerialPort.Open();

                if (!ManualScaleSerialPort.IsOpen) // Making sure Port is Open
                    ManualScaleSerialPort.Open();

                // Warning Message
                shiftRun.Warning = false;
                warningMessage(shiftRun.Warning);
            }
            catch (Exception ex) //Application Starts with warnings
            {
                //MessageBox.Show("Scale not connected\nPlease contact to IT", "Error", MessageBoxButtons.OK, MessageBoxIcon.Error);
                // Warning Message
                shiftRun.Warning = true;
                warningMessage(shiftRun.Warning);

                Logger.WARN("Catch Block App starts with Warning Message (START Button)");
            }

            //Update date
            labelDate.Text = DateTime.Now.ToString("MMM dd yyyy"); // Date on the right top corner



            // Set SKU target Weight
            shiftRun.Sku = labelSKUData.Text;


            shiftRun.LessWeight = Convert.ToDouble(labelLessData.Text.ToString().Replace("g", "")); // .ToString().Replace("g", "") remove .g 
            shiftRun.TargetWeight = Convert.ToDouble(labelTargetData.Text.ToString().Replace("g", "")); // .ToString().Replace("g", "") remove .g
            shiftRun.HeavyWeight = Convert.ToDouble(labelHeavyData.Text.ToString().Replace("g", "")); // .ToString().Replace("g", "") remove .g

            // Check if Weight all set otherwise invoke SET WEIGHT method 
            if (shiftRun.LessWeight == 0 || shiftRun.TargetWeight == 0 || shiftRun.HeavyWeight == 0 || shiftRun.Sku.Equals("") || shiftRun.Shift.Equals(""))
            {
                this.Invoke(new EventHandler(buttonSet_Click));
            }
            else // All set start RUNNING
            {
                if (timerRunning != null)
                {
                    timerRunning = null;
                    timerRunning = new TimerRunning();
                }

                timerRunning.startTimer();
                shiftRun.isBreak = true; // Break set so not counting right away
                timerRunning.setIsBreak(true); // Break set so not counting right away
                buttonStart.Enabled = false;
                labelShift.Visible = true;
                labelShift.Text = shiftRun.Shift;
                labelBarcode.Text = "Barcode - " + shiftRun.barCode;

                dataSaved = false; // Prep for saving
                //labelCHorKW.Text = shiftRun.Location; // Location when app starts
                
                buttonStart.Text = "Running";
                buttonStart.BackColor = Color.DarkGray;
                buttonStop.Enabled = true;
                buttonStop.BackColor = Color.Firebrick;
                //buttonSave.Enabled = false;
                shiftRun.Running = true;

                labelStaffRequired.Text = shiftRun.StaffNumberRequired.ToString();
                labelStaffActual.Text = shiftRun.StaffNumberActual.ToString();

                ThreadPool.QueueUserWorkItem(state => SaveData.checkIfFileDaylyReportExist(shiftRun)); // Check if file for Daly Report created. if not. Create

                // Check for oversize box. True if match SKU for oversize

                if (shiftRun.Sku.Equals("26411") || shiftRun.Sku.Equals("12602") || shiftRun.Sku.Equals("12503X") || shiftRun.Sku.Equals("12503Y"))
                {
                    shiftRun.boxOverSize = true; // This SKU is Oversize Box
                    labelOversizeBox.Visible = true; // Show the message about Oversize

                    Logger.INFO("SKU with oversize boxes (START Button)");
                }

                // Logg all info about starting
                Logger.INFO("");
                Logger.INFO("");
                Logger.INFO("");
                Logger.INFO("");
                Logger.INFO("NEW RUN STARTING");
                Logger.INFO("");
                Logger.INFO("");
                Logger.INFO("SKU: " + shiftRun.Sku);
                Logger.INFO("Barcode: " + shiftRun.barCode);
                Logger.INFO("SHIFT: " + shiftRun.Shift);
                Logger.INFO("Location: "+shiftRun.Location + " PL - " + shiftRun.PackLineNumber);
                Logger.INFO("Less: " + shiftRun.LessWeight);
                Logger.INFO("Target: " + shiftRun.TargetWeight);
                Logger.INFO("Heavy: " + shiftRun.HeavyWeight);
                Logger.INFO("Units p/hr Target: " + shiftRun.productivityTarget);
                Logger.INFO("Staff Required: " + shiftRun.StaffNumberRequired);
                Logger.INFO("Staff Actual: " + shiftRun.StaffNumberActual);
                Logger.INFO("");
                Logger.INFO("");

            }
        }

        // PL Scale Data Received
        private void PLScale_DataReceived(object sender, SerialDataReceivedEventArgs e)
        {
            try
            {
                if (PLScaleSerialPort.IsOpen)
                {
                    System.Threading.Thread.Sleep(500);  // Delay to be able to receive all data from the SLOW com port.
                    shiftRun.DataFromPLScale = PLScaleSerialPort.ReadExisting();


                    // Format delete all unwonted spaces
                    var sb = new StringBuilder(shiftRun.DataFromPLScale.Length);
                    foreach (char i in shiftRun.DataFromPLScale)
                        if (i != '\n' && i != '\r' && i != '\t')
                            sb.Append(i);


                    shiftRun.DataFromPLScale = sb.ToString();

                    if (shiftRun.Running) // Update Data only if Running (making sure Ports are open)
                        this.Invoke(new EventHandler(UpdateData));

                }
            }
            catch (Exception ex)
            {
                //MessageBox.Show(ex.Message, "Message", MessageBoxButtons.OK, MessageBoxIcon.Error); // Do not show this message App might be frozen
                // Warning Message
                shiftRun.Warning = true;
                warningMessage(shiftRun.Warning);
                shiftRun.DataFromPLScale = "";

                Logger.ERROR("Exeption thrown! (PLScale_DataReceived): " + ex);
            }
        }

        // Manual Scale Data Received
        private void ManualScale_DataReceived(object sender, SerialDataReceivedEventArgs e)
        {
            try
            {
                if (ManualScaleSerialPort.IsOpen)
                {

                    System.Threading.Thread.Sleep(500);  // Delay to be able to receive all data from the SLOW com port.
                    shiftRun.DataFromManualScale = ManualScaleSerialPort.ReadExisting();

                    // Format delete all unwonted spaces
                    var sb = new StringBuilder(shiftRun.DataFromManualScale.Length);
                    foreach (char i in shiftRun.DataFromManualScale)
                        if (i != '\n' && i != '\r' && i != '\t')
                            sb.Append(i);


                    shiftRun.DataFromManualScale = sb.ToString();


                    if (shiftRun.Running) // Update Data only if Running (making sure Ports are open)
                        this.Invoke(new EventHandler(UpdateData));
                }
            }
            catch (Exception ex)
            {
                //MessageBox.Show(ex.Message, "Message", MessageBoxButtons.OK, MessageBoxIcon.Error); // Do not show this message App might be frozen
                // Warning Message
                shiftRun.Warning = true;
                warningMessage(shiftRun.Warning);
                shiftRun.DataFromManualScale = "";

                Logger.ERROR("Exeption thrown! (ManualScale_DataReceived): " + ex);
            }
        }



        // Update Data Shows on the screen adding data to Collection
        static double dynamicAvg = 0.0; // temp collecting weight from the scale for Dynamic Avg
        private void UpdateData(object sender, EventArgs e)
        {
            //isTimerActive = true; // Start timer
            stopInSeconds = 0;
            timerRunning.setIsBreak(false);
            shiftRun.isBreak = false;
            shiftRun.isTimerON = true;
            labelBreakTimeS.Text = String.Format("{0:00}", 0);
            labelBreakTimeM.Text = String.Format("{0:00}.", 0);
            labelBreakTimeH.Text = String.Format("{0:00}:", 0);



            // If Data comes from PL scale
            if (shiftRun.DataFromPLScale != "" && shiftRun.boxOverSize == false)
            {
                double weightFromPLScaleAfterValidatin = validate.validateWeight(shiftRun.DataFromPLScale); // Store data From Scale in temp VAR
                string rawDataFromPLScaleNotFormated = shiftRun.DataFromPLScale;
                shiftRun.DataFromPLScale = ""; // I don't need this so I we can remove data

                // Heavy  If data Heavier than middle between target and heavy BUT not havier than target + 50%    
                if (weightFromPLScaleAfterValidatin > (shiftRun.TargetWeight + ((shiftRun.HeavyWeight - shiftRun.TargetWeight) / 10)) && weightFromPLScaleAfterValidatin < (shiftRun.TargetWeight + (shiftRun.TargetWeight / 2)))
                {
                    pLScaleWeightCollection.Add(new PLScaleWeight { WeightPLScale = weightFromPLScaleAfterValidatin, dateTime = DateTime.Now, Count = countPLScale });
                    shiftRun.PlCount++; // Total PL scale count update
                    shiftRun.HeavyCount++; // Heavy count update
                    labelCount.Text = shiftRun.PlCount.ToString(); // // PL Scale Value update
                    
                    label_plData.Text = rawDataFromPLScaleNotFormated; // Straight from scale not formated and NOT converted to Double

                    //shiftRun.DataFromPLScale = ""; // I don't need this so I we can remove data

                    labelCountHeavy.Text = shiftRun.HeavyCount.ToString(); // Update total Heavy count
                    
                    countPLScale++; // Get count ready for next box

                    shiftRun.AverageWeight += weightFromPLScaleAfterValidatin; // Collecting all weights for average calculation

                    shiftRun.AverageWeightHeavy += weightFromPLScaleAfterValidatin; // Heavy weights for average

                    shiftRun.AverageWeightCount += weightFromPLScaleAfterValidatin; // weights for average COUNT

                    dynamicAvg = weightFromPLScaleAfterValidatin; // temp collecting weight from the scale for Dynamic Avg

                    //Calculate Giving Away 
                    shiftRun.kgGivingAway += (weightFromPLScaleAfterValidatin - shiftRun.TargetWeight);
                }

                // If data heavier than Target + (50% out of Heavy) add to collection as a heavy
                if (weightFromPLScaleAfterValidatin >= (shiftRun.TargetWeight + (shiftRun.HeavyWeight / 2)))
                {
                    pLScaleWeightCollection.Add(new PLScaleWeight { WeightPLScale = shiftRun.HeavyWeight, dateTime = DateTime.Now, Count = countPLScale });
                    shiftRun.PlCount++; // Total PL scale count update
                    shiftRun.HeavyCount++; // Heavy count update
                    labelCount.Text = shiftRun.PlCount.ToString(); // // PL Scale Value update
                    label_plData.Text = rawDataFromPLScaleNotFormated; // Straight from scale not formated and NOT converted to Double

                    labelCountHeavy.Text = shiftRun.HeavyCount.ToString(); // Update total Heavy count

                    Logger.DEBUG("Weigth Error HEAVY collection (Data comes from PL scale) RAW Data: " + rawDataFromPLScaleNotFormated);

                    countPLScale++; // Get count ready for next box
                    
                    //shiftRun.DataFromPLScale = ""; // I don't need this so I we can remove data

                    shiftRun.AverageWeight += shiftRun.HeavyWeight; // Collecting all weights for average calculation

                    shiftRun.AverageWeightHeavy += shiftRun.HeavyWeight; // Heavy weights for average

                    shiftRun.AverageWeightCount += shiftRun.HeavyWeight; // weights for average COUNT

                    dynamicAvg = shiftRun.HeavyWeight; // temp collecting weight from the scale for Dynamic Avg

                    //Calculate Giving Away 
                    shiftRun.kgGivingAway += (shiftRun.HeavyWeight - shiftRun.TargetWeight);

                    //
                    //
                    // NED TO ADD ERROR COLLECTIN AS THIS WILL RICH HERE IF SCALE OUTPUTS WRONG DATA
                    //
                    //
                    shiftRun.errorCountPL++; // add PL error
                }

                // Target If data equal to Target or less than middle between target and heavy  
                if (weightFromPLScaleAfterValidatin >= shiftRun.TargetWeight && weightFromPLScaleAfterValidatin <= (shiftRun.TargetWeight + ((shiftRun.HeavyWeight - shiftRun.TargetWeight) / 10)))
                {
                    pLScaleWeightCollection.Add(new PLScaleWeight { WeightPLScale = weightFromPLScaleAfterValidatin, dateTime = DateTime.Now, Count = countPLScale });
                    shiftRun.PlCount++; // Total PL scale count update
                    shiftRun.TargetCount++; // Target count update
                    labelCount.Text = shiftRun.PlCount.ToString(); // // PL Scale Value update
                    label_plData.Text = rawDataFromPLScaleNotFormated; // Straight from scale not formated and NOT converted to Double

                    labelCountTarget.Text = shiftRun.TargetCount.ToString(); // Update total Target count
                    
                    //shiftRun.DataFromPLScale = ""; // I don't need this so I we can remove data
                    
                    countPLScale++; // Get count ready for next box

                    shiftRun.AverageWeight += weightFromPLScaleAfterValidatin; // Collecting all weights for average calculation

                    shiftRun.AverageWeightTarget += weightFromPLScaleAfterValidatin; // Target weights for average

                    shiftRun.AverageWeightCount += weightFromPLScaleAfterValidatin; // weights for average COUNT

                    dynamicAvg = weightFromPLScaleAfterValidatin; // temp collecting weight from the scale for Dynamic Avg

                    //Calculate Giving Away 
                    shiftRun.kgGivingAway += (weightFromPLScaleAfterValidatin - shiftRun.TargetWeight);
                }

                // Less If Data less than Target but grater than target - 50%
                if (weightFromPLScaleAfterValidatin < shiftRun.TargetWeight && weightFromPLScaleAfterValidatin > (shiftRun.TargetWeight - (shiftRun.TargetWeight / 2)))
                {
                    pLScaleWeightCollection.Add(new PLScaleWeight { WeightPLScale = weightFromPLScaleAfterValidatin, dateTime = DateTime.Now, Count = countPLScale });
                    shiftRun.PlCount++; // Total PL scale count update
                    shiftRun.LessCount++; // Less count update
                    labelCount.Text = shiftRun.PlCount.ToString(); // // PL Scale Value update
                    label_plData.Text = rawDataFromPLScaleNotFormated; // Straight from scale not formated and NOT converted to Double

                    labelCountLess.Text = shiftRun.LessCount.ToString(); // Update total Less count
                    
                    //shiftRun.DataFromPLScale = ""; // I don't need this so I we can remove data

                    countPLScale++; // Get count ready for next box

                    shiftRun.AverageWeight += weightFromPLScaleAfterValidatin; // Collecting all weights for average calculation

                    shiftRun.AverageWeightLess += weightFromPLScaleAfterValidatin; // Less weights for average

                    shiftRun.AverageWeightCount += weightFromPLScaleAfterValidatin; // weights for average COUNT

                    dynamicAvg = weightFromPLScaleAfterValidatin; // temp collecting weight from the scale for Dynamic Avg    

                    //Calculate Giving Away 
                    shiftRun.kgGivingAway += (weightFromPLScaleAfterValidatin - shiftRun.TargetWeight);
                }


                // If data from the scale less than target - 50% . Add count and add ptoduct as a Less weight

                if (weightFromPLScaleAfterValidatin <= (shiftRun.TargetWeight - (shiftRun.TargetWeight / 2)))
                {
                    pLScaleWeightCollection.Add(new PLScaleWeight { WeightPLScale = shiftRun.LessWeight, dateTime = DateTime.Now, Count = countPLScale });
                    shiftRun.PlCount++; // Total PL scale count update
                    shiftRun.LessCount++; // Less count update
                    labelCount.Text = shiftRun.PlCount.ToString(); // // PL Scale Value update
                    label_plData.Text = rawDataFromPLScaleNotFormated; // Straight from scale not formated and NOT converted to Double

                    Logger.DEBUG("Weigth Error LESS collection (Data comes from PL scale) RAW Data: " + rawDataFromPLScaleNotFormated);

                    labelCountLess.Text = shiftRun.LessCount.ToString(); // Update total Less count
                    
                    //shiftRun.DataFromPLScale = ""; // I don't need this so I we can remove data

                    countPLScale++; // Get count ready for next box

                    shiftRun.AverageWeight += shiftRun.LessWeight; // Collecting all weights for average calculation

                    shiftRun.AverageWeightLess += shiftRun.LessWeight; // Less weights for average

                    shiftRun.AverageWeightCount += shiftRun.LessWeight; // weights for average COUNT

                    dynamicAvg = shiftRun.LessWeight; // temp collecting weight from the scale for Dynamic Avg    

                    //Calculate Giving Away 
                    shiftRun.kgGivingAway += (shiftRun.LessWeight - shiftRun.TargetWeight);

                    //
                    //
                    // NED TO ADD ERROR COLLECTIN AS THIS WILL RICH HERE IF SCALE OUTPUTS WRONG DATA
                    //
                    //
                    shiftRun.errorCountPL++; // add PL error
                }

            }
            if (shiftRun.DataFromManualScale != "") // If Data comes from Manual scale
            {
                double weightFromManualScaleAfterValidatin = validate.validateWeight(shiftRun.DataFromManualScale); // Store data From Scale in temp VAR 
                string rawDataFromManualScaleNotFormated = shiftRun.DataFromManualScale;
                shiftRun.DataFromManualScale = ""; // I don't need this so I we can remove data

                //Less
                // Adjusted If data between Less and Target
                if (weightFromManualScaleAfterValidatin >= shiftRun.LessWeight && weightFromManualScaleAfterValidatin < shiftRun.TargetWeight)
                {
                    manualScaleWeightCollection.Add(new ManualScaleWeight { WeightManualScale = weightFromManualScaleAfterValidatin, dateTime = DateTime.Now, Count = countManualScale });
                    shiftRun.ManualCount++; // Total Manual scale count update
                    shiftRun.LessCount++; // Target count update
                    labelAdjusted.Text = shiftRun.ManualCount.ToString(); // Adjusted Scale Value update
                    label_manualData.Text = rawDataFromManualScaleNotFormated; // Straight from scale not formated and NOT converted to Double

                    labelCountLess.Text = shiftRun.LessCount.ToString(); // Update total Heavy count
                    
                    //shiftRun.DataFromManualScale = ""; // I don't need this so I we can remove data

                    countManualScale++; // Get count ready for next box

                    shiftRun.AverageWeight += weightFromManualScaleAfterValidatin; // Collecting all weights for average calculation

                    shiftRun.AverageWeightLess += weightFromManualScaleAfterValidatin; // Target weights for average

                    shiftRun.AverageWeightAdjusted += weightFromManualScaleAfterValidatin; // weights for average Adjusted

                    dynamicAvg = weightFromManualScaleAfterValidatin; // temp collecting weight from the scale for Dynamic Avg    

                    //Calculate Giving Away 
                    shiftRun.kgGivingAway += (weightFromManualScaleAfterValidatin - shiftRun.TargetWeight);
                }

                //Target
                // Adjusted If data Target of less than Middle between Heavy and Target
                if (weightFromManualScaleAfterValidatin >= shiftRun.TargetWeight && weightFromManualScaleAfterValidatin <= (shiftRun.TargetWeight + ((shiftRun.HeavyWeight - shiftRun.TargetWeight) / 10)))
                {
                    manualScaleWeightCollection.Add(new ManualScaleWeight { WeightManualScale = weightFromManualScaleAfterValidatin, dateTime = DateTime.Now, Count = countManualScale });
                    shiftRun.ManualCount++; // Total Manual scale count update
                    shiftRun.TargetCount++; // Target count update
                    labelAdjusted.Text = shiftRun.ManualCount.ToString(); // Adjusted Scale Value update
                    label_manualData.Text = rawDataFromManualScaleNotFormated; // Straight from scale not formated and NOT converted to Double

                    labelCountTarget.Text = shiftRun.TargetCount.ToString(); // Update total Heavy count
                    
                    //shiftRun.DataFromManualScale = ""; // I don't need this so I we can remove data
                    
                    countManualScale++; // Get count ready for next box

                    shiftRun.AverageWeight += weightFromManualScaleAfterValidatin; // Collecting all weights for average calculation

                    shiftRun.AverageWeightTarget += weightFromManualScaleAfterValidatin; // Target weights for average

                    shiftRun.AverageWeightAdjusted += weightFromManualScaleAfterValidatin; // weights for average Adjusted

                    dynamicAvg = weightFromManualScaleAfterValidatin; // temp collecting weight from the scale for Dynamic Avg

                    //Calculate Giving Away 
                    shiftRun.kgGivingAway += (weightFromManualScaleAfterValidatin - shiftRun.TargetWeight);
                }

                //Heavy
                // Adjusted If data heavier than Middle between Heavy and Target and Less than Target + 50% 
                if (weightFromManualScaleAfterValidatin > (shiftRun.TargetWeight + ((shiftRun.HeavyWeight - shiftRun.TargetWeight) / 10)) && weightFromManualScaleAfterValidatin < (shiftRun.TargetWeight + (shiftRun.TargetWeight / 2)))
                {
                    manualScaleWeightCollection.Add(new ManualScaleWeight { WeightManualScale = weightFromManualScaleAfterValidatin, dateTime = DateTime.Now, Count = countManualScale });
                    shiftRun.ManualCount++; // Total Manual scale count update
                    shiftRun.HeavyCount++; // Heavy count update
                    labelAdjusted.Text = shiftRun.ManualCount.ToString(); // Adjusted Scale Value update
                    label_manualData.Text = rawDataFromManualScaleNotFormated; // Straight from scale not formated and NOT converted to Double

                    labelCountHeavy.Text = shiftRun.HeavyCount.ToString(); // Update total Heavy count
                    
                    //shiftRun.DataFromManualScale = ""; // I don't need this so I we can remove data

                    countManualScale++; // Get count ready for next box

                    shiftRun.AverageWeight += weightFromManualScaleAfterValidatin; // Collecting all weights for average calculation

                    shiftRun.AverageWeightHeavy += weightFromManualScaleAfterValidatin; // Heavy weights for average

                    shiftRun.AverageWeightAdjusted += weightFromManualScaleAfterValidatin; // weights for average Adjusted

                    dynamicAvg = weightFromManualScaleAfterValidatin; // temp collecting weight from the scale for Dynamic Avg

                    //Calculate Giving Away 
                    shiftRun.kgGivingAway += (weightFromManualScaleAfterValidatin - shiftRun.TargetWeight);
                }

                //Heavy Error collection
                // Adjusted If data heavier than Middle between Heavy and Target and Less than Target + 50% 
                if (weightFromManualScaleAfterValidatin >= (shiftRun.TargetWeight + (shiftRun.TargetWeight / 2)))
                {
                    manualScaleWeightCollection.Add(new ManualScaleWeight { WeightManualScale = shiftRun.HeavyWeight, dateTime = DateTime.Now, Count = countManualScale });
                    shiftRun.ManualCount++; // Total Manual scale count update
                    shiftRun.HeavyCount++; // Heavy count update
                    labelAdjusted.Text = shiftRun.ManualCount.ToString(); // Adjusted Scale Value update
                    label_manualData.Text = rawDataFromManualScaleNotFormated; // Straight from scale not formated and NOT converted to Double

                    Logger.DEBUG("Weigth Error HEAVY collection (Data comes from Manual scale) RAW Data: " + rawDataFromManualScaleNotFormated);

                    labelCountHeavy.Text = shiftRun.HeavyCount.ToString(); // Update total Heavy count
                    
                    //shiftRun.DataFromManualScale = ""; // I don't need this so I we can remove data

                    countManualScale++; // Get count ready for next box

                    shiftRun.AverageWeight += shiftRun.HeavyWeight; // Collecting all weights for average calculation

                    shiftRun.AverageWeightHeavy += shiftRun.HeavyWeight; // Heavy weights for average

                    shiftRun.AverageWeightAdjusted += shiftRun.HeavyWeight; // weights for average Adjusted

                    dynamicAvg = shiftRun.HeavyWeight; // temp collecting weight from the scale for Dynamic Avg

                    //Calculate Giving Away 
                    shiftRun.kgGivingAway += (shiftRun.HeavyWeight - shiftRun.TargetWeight);

                    //
                    //
                    // NED TO ADD ERROR COLLECTIN AS THIS WILL RICH HERE IF SCALE OUTPUTS WRONG DATA
                    //
                    //
                    shiftRun.errorCountManual++; // add error count
                }

            }

            // Errors Label update
            labelErrorPL.Text = shiftRun.errorCountPL.ToString(); // PL Errors
            labelErrorManual.Text = shiftRun.errorCountManual.ToString(); // Manual Errors

            // UPH Acual Count and Display
            if (shiftRun.runningTimeInSeconds > 0 && (shiftRun.PlCount > 0 || shiftRun.ManualCount > 0))
            {
                if (shiftRun.PlCount >= shiftRun.ManualCount) // If most boxes from PL  
                {
                    shiftRun.productivityActual = unitsPerHour.unitsPerHourOveral(shiftRun.runningTimeInSeconds, shiftRun.PlCount);
                    //shiftRun.productivityActual = shiftRun.PlCount * 60 / shiftRun.runningTimeInMin; // If most boxes from PL
                }
                else
                {
                    shiftRun.productivityActual = unitsPerHour.unitsPerHourOveral(shiftRun.runningTimeInSeconds, shiftRun.ManualCount);
                    //shiftRun.productivityActual = shiftRun.ManualCount * 60 / shiftRun.runningTimeInMin; // If most boxes from Manua scale
                }
            }


            // Running Efficiency label color 
            if (shiftRun.productivityActual >= shiftRun.productivityTarget)
            {
                labelProdActualData.ForeColor = Color.Green;
                labelUPHPercentage.ForeColor = Color.Green;
            }
            else
            {
                labelProdActualData.ForeColor = Color.Red;
                labelUPHPercentage.ForeColor = Color.Red;
            }

            labelProdActualData.Text = shiftRun.productivityActual.ToString("0");

            // If Proguctivity Target set 
            if (shiftRun.productivityTarget > 0)
                shiftRun.runningEfficiency = shiftRun.productivityActual * 100 / shiftRun.productivityTarget; // Calculate Running Efficiency 

            labelUPHPercentage.Text = (shiftRun.runningEfficiency * 0.01).ToString("0.0 %"); // Show Running Efficiency Percentage

            // If Staff number set
            if (shiftRun.StaffNumberActual > 1 && shiftRun.StaffNumberRequired > 1)
                shiftRun.expectedEfficiency = shiftRun.StaffNumberActual * 100 / shiftRun.StaffNumberRequired; // Calculating Expected Efficiency

            expectedEfficiencyLabel.Text = (shiftRun.expectedEfficiency * 0.01).ToString("0.0 %"); // Show Expected Efficiency


            //Test UPH by hours
            //if (shiftRun.runningTimeInMin < 60)
            //{
            //   labelFirstHour.Text = unitsPerHour.unitsPerHourLastHour(shiftRun.runningTimeInMin, shiftRun.PlCount).ToString("0");
            //}
            //if (shiftRun.runningTimeInMin >= 60)
            //{
            //    labelSecondHour.Text = unitsPerHour.unitsPerHourLastHour(shiftRun.runningTimeInMin, shiftRun.PlCount).ToString("0");
            //}



            int hour = shiftRun.runningTimeInSeconds / 3600;  // It was 60 Now 3600 as now we count in Seconds. Calculating which hour is running right now 

            if (hour > currentHour)
            {
                currentHour = hour;
                unitsPerHour.setNextHour(true); 
            }

            switch (hour)
            {
                case 0:
                    if (shiftRun.PlCount >= shiftRun.ManualCount)
                    {
                        uphCollection[0][0] = unitsPerHour.unitsPerHourLastHour(shiftRun.runningTimeInSeconds, shiftRun.PlCount).ToString();
                        uphCollection[0][1] = DateTime.Now.ToString("h:mm tt");
                        //Console.WriteLine("Switch statement 0 - " + hour);
                    }
                    else
                    {
                        uphCollection[0][0] = unitsPerHour.unitsPerHourLastHour(shiftRun.runningTimeInSeconds, shiftRun.ManualCount).ToString();
                        uphCollection[0][1] = DateTime.Now.ToString("h:mm tt");
                    }
                    break;
                case 1:
                    if (shiftRun.PlCount >= shiftRun.ManualCount)
                    {
                        uphCollection[1][0] = unitsPerHour.unitsPerHourLastHour(shiftRun.runningTimeInSeconds, shiftRun.PlCount).ToString();
                        uphCollection[1][1] = DateTime.Now.ToString("h:mm tt");
                        //Console.WriteLine("Switch statement 1 - " + hour);
                    }
                    else
                    {
                        uphCollection[1][0] = unitsPerHour.unitsPerHourLastHour(shiftRun.runningTimeInSeconds, shiftRun.ManualCount).ToString();
                        uphCollection[1][1] = DateTime.Now.ToString("h:mm tt");
                    }
                    break;
                case 2:
                    if (shiftRun.PlCount >= shiftRun.ManualCount)
                    {
                        uphCollection[2][0] = unitsPerHour.unitsPerHourLastHour(shiftRun.runningTimeInSeconds, shiftRun.PlCount).ToString();
                        uphCollection[2][1] = DateTime.Now.ToString("h:mm tt");
                        //Console.WriteLine("Switch statement 2 - " + hour);
                    }
                    else
                    {
                        uphCollection[2][0] = unitsPerHour.unitsPerHourLastHour(shiftRun.runningTimeInSeconds, shiftRun.ManualCount).ToString();
                        uphCollection[2][1] = DateTime.Now.ToString("h:mm tt");
                    }
                    break;
                case 3:
                    if (shiftRun.PlCount >= shiftRun.ManualCount)
                    {
                        uphCollection[3][0] = unitsPerHour.unitsPerHourLastHour(shiftRun.runningTimeInSeconds, shiftRun.PlCount).ToString();
                        uphCollection[3][1] = DateTime.Now.ToString("h:mm tt");
                        //Console.WriteLine("Switch statement 3 - " + hour);
                    }
                    else
                    {
                        uphCollection[3][0] = unitsPerHour.unitsPerHourLastHour(shiftRun.runningTimeInSeconds, shiftRun.ManualCount).ToString();
                        uphCollection[3][1] = DateTime.Now.ToString("h:mm tt");
                    }
                    break;
                case 4:
                    if (shiftRun.PlCount >= shiftRun.ManualCount)
                    {
                        uphCollection[4][0] = unitsPerHour.unitsPerHourLastHour(shiftRun.runningTimeInSeconds, shiftRun.PlCount).ToString();
                        uphCollection[4][1] = DateTime.Now.ToString("h:mm tt");
                        //Console.WriteLine("Switch statement 4 - " + hour);
                    }
                    else
                    {
                        uphCollection[4][0] = unitsPerHour.unitsPerHourLastHour(shiftRun.runningTimeInSeconds, shiftRun.ManualCount).ToString();
                        uphCollection[4][1] = DateTime.Now.ToString("h:mm tt");
                    }
                    break;
                case 5:
                    if (shiftRun.PlCount >= shiftRun.ManualCount)
                    {
                        uphCollection[5][0] = unitsPerHour.unitsPerHourLastHour(shiftRun.runningTimeInSeconds, shiftRun.PlCount).ToString();
                        uphCollection[5][1] = DateTime.Now.ToString("h:mm tt");
                        //Console.WriteLine("Switch statement 5 - " + hour);
                    }
                    else
                    {
                        uphCollection[5][0] = unitsPerHour.unitsPerHourLastHour(shiftRun.runningTimeInSeconds, shiftRun.ManualCount).ToString();
                        uphCollection[5][1] = DateTime.Now.ToString("h:mm tt");
                    }
                    break;
                case 6:
                    if (shiftRun.PlCount >= shiftRun.ManualCount)
                    {
                        uphCollection[6][0] = unitsPerHour.unitsPerHourLastHour(shiftRun.runningTimeInSeconds, shiftRun.PlCount).ToString();
                        uphCollection[6][1] = DateTime.Now.ToString("h:mm tt");
                    }
                    else
                    {
                        uphCollection[6][0] = unitsPerHour.unitsPerHourLastHour(shiftRun.runningTimeInSeconds, shiftRun.ManualCount).ToString();
                        uphCollection[6][1] = DateTime.Now.ToString("h:mm tt");
                    }
                    break;
                case 7:
                    if (shiftRun.PlCount >= shiftRun.ManualCount)
                    {
                        uphCollection[7][0] = unitsPerHour.unitsPerHourLastHour(shiftRun.runningTimeInSeconds, shiftRun.PlCount).ToString();
                        uphCollection[7][1] = DateTime.Now.ToString("h:mm tt");
                    }
                    else
                    {
                        uphCollection[7][0] = unitsPerHour.unitsPerHourLastHour(shiftRun.runningTimeInSeconds, shiftRun.ManualCount).ToString();
                        uphCollection[7][1] = DateTime.Now.ToString("h:mm tt");
                    }
                    break;
                case 8:
                    if (shiftRun.PlCount >= shiftRun.ManualCount)
                    {
                        uphCollection[8][0] = unitsPerHour.unitsPerHourLastHour(shiftRun.runningTimeInSeconds, shiftRun.PlCount).ToString();
                        uphCollection[8][1] = DateTime.Now.ToString("h:mm tt");
                    }
                    else
                    {
                        uphCollection[8][0] = unitsPerHour.unitsPerHourLastHour(shiftRun.runningTimeInSeconds, shiftRun.ManualCount).ToString();
                        uphCollection[8][1] = DateTime.Now.ToString("h:mm tt");
                    }
                    break;
                default:
                    // Default case should neve be executed
                    break;
            }



            //Test UPH by hours TEST END


            // Show Give Away weight on the screen
            if (shiftRun.kgGivingAway > 0)
            {
                labelGiveAwayData.ForeColor = Color.Red;
                labelGiveAwayData.Text = (shiftRun.kgGivingAway * 0.001).ToString("0.0 kg"); // Show Giving away data in KG

                //Percentage Give Away
                shiftRun.percentageGivingAway = (((((shiftRun.PlCount + shiftRun.ManualCount) * (shiftRun.AverageWeight / (shiftRun.ManualCount + shiftRun.PlCount))) * 100) / ((shiftRun.PlCount + shiftRun.ManualCount) * shiftRun.TargetWeight)) - 100) * 0.01; // Calculate Give Away Percentage

                // New idea FOR CAlCULATING a percentageGiveAway
                //shiftRun.percentageGivingAway = (shiftRun.kgGivingAway * 100) / ((shiftRun.PlCount + shiftRun.ManualCount) * shiftRun.TargetWeight);
                
                labelGiveAwayPercentage.ForeColor = Color.Red; // Label color
                labelGiveAwayPercentage.Text = shiftRun.percentageGivingAway.ToString("0.0 %"); // Show Giving away data in %
            }
            if (shiftRun.kgGivingAway < 0)
            {
                labelGiveAwayData.ForeColor = Color.Green;
                labelGiveAwayData.Text = (shiftRun.kgGivingAway * 0.001).ToString("0.0 kg"); // Show Giving away data in KG

                //Percentage Give Away
                shiftRun.percentageGivingAway = (((((shiftRun.PlCount + shiftRun.ManualCount) * (shiftRun.AverageWeight / (shiftRun.ManualCount + shiftRun.PlCount))) * 100) / ((shiftRun.PlCount + shiftRun.ManualCount) * shiftRun.TargetWeight)) - 100) * 0.01; // Calculate Give Away Percentage
                labelGiveAwayPercentage.ForeColor = Color.Green; // Label color
                labelGiveAwayPercentage.Text = shiftRun.percentageGivingAway.ToString("0.0 %"); // Show Giving away data in %
            }



            // Total Average shows on the screen

            // Color For Average If Less
            if ((shiftRun.AverageWeight / (shiftRun.ManualCount + shiftRun.PlCount)) < shiftRun.TargetWeight)
            {
                labelAverageWeight.ForeColor = Color.Goldenrod;
            }
            // Color For Average If Target
            if ((shiftRun.AverageWeight / (shiftRun.ManualCount + shiftRun.PlCount)) >= shiftRun.TargetWeight && (shiftRun.AverageWeight / (shiftRun.ManualCount + shiftRun.PlCount)) <= (shiftRun.TargetWeight + ((shiftRun.HeavyWeight - shiftRun.TargetWeight) / 10)))
            {
                labelAverageWeight.ForeColor = Color.Green;
            }
            // Color For Average If Heavy
            if ((shiftRun.AverageWeight / (shiftRun.ManualCount + shiftRun.PlCount)) > (shiftRun.TargetWeight + ((shiftRun.HeavyWeight - shiftRun.TargetWeight) / 10)))
            {
                labelAverageWeight.ForeColor = Color.Red;
            }


            dynamicAvg = calculateDynamicAvg(dynamicAvg); // Calculate Daynamic Avg  Last 100 Boxes

            // Color For Dynamic Average If Less
            if (dynamicAvg != 0.0)
            {
                if (dynamicAvg < shiftRun.TargetWeight)
                {
                    labelDynamicAvgWeight.Text = dynamicAvg.ToString("0.0 g");
                    labelDynamicAvgWeight.ForeColor = Color.Goldenrod;
                }
                // Color For Dynamic Average If Target
                if (dynamicAvg >= shiftRun.TargetWeight && (shiftRun.AverageWeight / (shiftRun.ManualCount + shiftRun.PlCount)) <= (shiftRun.TargetWeight + ((shiftRun.HeavyWeight - shiftRun.TargetWeight) / 10)))
                {
                    labelDynamicAvgWeight.Text = dynamicAvg.ToString("0.0 g");
                    labelDynamicAvgWeight.ForeColor = Color.Green;
                }
                // Color For Dynamic Average If Heavy
                if (dynamicAvg > (shiftRun.TargetWeight + ((shiftRun.HeavyWeight - shiftRun.TargetWeight) / 10)))
                {
                    labelDynamicAvgWeight.Text = dynamicAvg.ToString("0.0 g");
                    labelDynamicAvgWeight.ForeColor = Color.Red;
                }
            }
            else
            {
                labelDynamicAvgWeight.Text = dynamicAvg.ToString("0.0 g");
                labelDynamicAvgWeight.ForeColor = Color.White;
            }

            //dynamicAvg = 0.0;


            labelAverageWeight.Text = (shiftRun.AverageWeight / (shiftRun.ManualCount + shiftRun.PlCount)).ToString("0.0 g");


            if (shiftRun.PlCount > 0)
                labelAverageCount.Text = (shiftRun.AverageWeightCount / shiftRun.PlCount).ToString("0.0 g");
            if (shiftRun.ManualCount > 0)
                labelAverageAdjusted.Text = (shiftRun.AverageWeightAdjusted / shiftRun.ManualCount).ToString("0.0 g");
            if (shiftRun.LessCount > 0)
                labelAvgLess.Text = (shiftRun.AverageWeightLess / shiftRun.LessCount).ToString("0.0 g"); // Less Average shows on the screen
            if (shiftRun.TargetCount > 0)
                labelAvgTarget.Text = (shiftRun.AverageWeightTarget / shiftRun.TargetCount).ToString("0.0 g"); // Target Average shows on the screen
            if (shiftRun.HeavyCount > 0)
                labelAvgHeavy.Text = (shiftRun.AverageWeightHeavy / shiftRun.HeavyCount).ToString("0.0 g"); // Heavy Average shows on the screen




            // Char for weight load Data start

            chartWeight.Series["Count"].Points.Clear();
            //chartWeight.Series["Count"]["LabelStyle"] = "Bottom";

            //chartWeight.ChartAreas[0].BackColor = Color.Blue;

            chartWeight.ChartAreas[0].AxisX.MajorGrid.Enabled = false;
            chartWeight.ChartAreas[0].AxisX.MinorGrid.Enabled = false;
            chartWeight.ChartAreas[0].AxisY.MajorGrid.Enabled = false;
            chartWeight.ChartAreas[0].AxisY.MinorGrid.Enabled = false;
            chartWeight.ChartAreas[0].AxisX.MajorTickMark.Enabled = false;
            chartWeight.ChartAreas[0].AxisY.MajorTickMark.Enabled = false;
            chartWeight.ChartAreas[0].AxisX.MinorTickMark.Enabled = false;
            chartWeight.ChartAreas[0].AxisY.MinorTickMark.Enabled = false;


            chartWeight.ChartAreas[0].AxisX.LineColor = Color.White;
            chartWeight.ChartAreas[0].AxisX.MajorGrid.LineColor = Color.White;
            chartWeight.ChartAreas[0].AxisX.LabelStyle.ForeColor = Color.White;

            chartWeight.ChartAreas[0].AxisY.LineColor = Color.White;
            chartWeight.ChartAreas[0].AxisY.MajorGrid.LineColor = Color.White;
            chartWeight.ChartAreas[0].AxisY.LabelStyle.ForeColor = Color.White;


            chartWeight.Series["Count"].Points.Add(shiftRun.LessCount);
            chartWeight.Series["Count"].Points[0].Color = Color.Goldenrod;
            chartWeight.Series["Count"].Points[0].LegendText = "Less";
            chartWeight.Series["Count"].Points[0].AxisLabel = "Less";
            chartWeight.Series["Count"].Points[0].Label = shiftRun.LessCount.ToString();

            chartWeight.Series["Count"].Points.Add(shiftRun.TargetCount);
            chartWeight.Series["Count"].Points[1].Color = Color.ForestGreen;
            chartWeight.Series["Count"].Points[1].LegendText = "Target";
            chartWeight.Series["Count"].Points[1].AxisLabel = "Target";
            chartWeight.Series["Count"].Points[1].Label = shiftRun.TargetCount.ToString();

            chartWeight.Series["Count"].Points.Add(shiftRun.HeavyCount);
            chartWeight.Series["Count"].Points[2].Color = Color.Red;
            chartWeight.Series["Count"].Points[2].LegendText = "Heavy";
            chartWeight.Series["Count"].Points[2].AxisLabel = "Heavy";
            chartWeight.Series["Count"].Points[2].Label = shiftRun.HeavyCount.ToString();

            // Char for weight load Data end



            // Char for UPH load Data start

            chartUPH.Series["UPH"].Points.Clear();

            chartUPH.ChartAreas[0].AxisX.MajorGrid.Enabled = false;
            chartUPH.ChartAreas[0].AxisX.MinorGrid.Enabled = false;
            chartUPH.ChartAreas[0].AxisY.MajorGrid.Enabled = false;
            chartUPH.ChartAreas[0].AxisY.MinorGrid.Enabled = false;
            chartUPH.ChartAreas[0].AxisX.MajorTickMark.Enabled = false;
            chartUPH.ChartAreas[0].AxisY.MajorTickMark.Enabled = false;
            chartUPH.ChartAreas[0].AxisX.MinorTickMark.Enabled = false;
            chartUPH.ChartAreas[0].AxisY.MinorTickMark.Enabled = false;

            chartUPH.ChartAreas[0].BorderColor = Color.Red;


            chartUPH.ChartAreas[0].AxisX.LineColor = Color.White;
            chartUPH.ChartAreas[0].AxisX.MajorGrid.LineColor = Color.White;
            chartUPH.ChartAreas[0].AxisX.LabelStyle.ForeColor = Color.White;

            chartUPH.ChartAreas[0].AxisY.LineColor = Color.White;
            chartUPH.ChartAreas[0].AxisY.MajorGrid.LineColor = Color.White;
            chartUPH.ChartAreas[0].AxisY.LabelStyle.ForeColor = Color.White;


            chartUPH.ChartAreas[0].AxisX.IsStartedFromZero = true;
            chartUPH.ChartAreas[0].AxisY.IsStartedFromZero = true;

            chartUPH.Series["UPH"].Points.AddXY(0, 0);
            chartUPH.Series["UPH"].Points[0].LegendText = "Started at " + timerRunning.getStartintTime().ToString("h:mm tt"); // Time stamp
            chartUPH.Series["UPH"].Points[0].AxisLabel = "Started at " + timerRunning.getStartintTime().ToString("h:mm tt"); // Time stamp


            for (int i = 0; i < uphCollection.Length; i++)
            {
                if (uphCollection[i][0] != "")
                {
                    //Console.WriteLine("uphCollection[" + i+ "]  = " + uphCollection[i]);
                    double uphPercentage = 0.0;

                    if (shiftRun.productivityTarget > 0) // Add to chart if Target set only
                    {
                        uphPercentage = Int32.Parse(uphCollection[i][0]) * 100 / shiftRun.productivityTarget;
                        Console.WriteLine("uphCollection calculated: " + uphPercentage);

                        
                        //chartUPH.ChartAreas[0].AxisX.IsStartedFromZero = true;

                        if (Int32.Parse(uphCollection[i][0]) >= shiftRun.productivityTarget)  // Target or above target Shows as a green color
                        {
                            chartUPH.Series["UPH"].Points.Add(uphPercentage * 0.01);
                            chartUPH.Series["UPH"].Points[i + 1].Color = Color.Green;
                            chartUPH.Series["UPH"].Points[i + 1].LegendText = uphCollection[i][1]; // Time stamp
                            chartUPH.Series["UPH"].Points[i + 1].AxisLabel = uphCollection[i][1]; // Time stamp
                            chartUPH.Series["UPH"].Points[i + 1].Label = (uphPercentage * 0.01).ToString("0 %");
                        }
                        else // Less than a Target  Shows as a red color
                        {
                            chartUPH.Series["UPH"].Points.Add(uphPercentage * 0.01);
                            chartUPH.Series["UPH"].Points[i + 1].Color = Color.Red;
                            chartUPH.Series["UPH"].Points[i + 1].LegendText = uphCollection[i][1]; // Time stamp
                            chartUPH.Series["UPH"].Points[i + 1].AxisLabel = uphCollection[i][1]; // Time stamp
                            chartUPH.Series["UPH"].Points[i + 1].Label = (uphPercentage * 0.01).ToString("0 %");
                        }

                    }

                }

            }


            // Char for weight load Data end




            // Select last added row to DataGrid PL Scale 
            if (dataGridViewPLScale != null)
            {
                if (dataGridViewPLScale.Rows.Count > 0)
                {
                    dataGridViewPLScale.Rows[dataGridViewPLScale.Rows.Count - 1].Selected = true;
                    dataGridViewPLScale.FirstDisplayedScrollingRowIndex = dataGridViewPLScale.Rows.Count - 1;
                }
            }
            // Select last added row to DataGrid Manual Scale
            if (dataGridViewManualScale != null)
            {
                if (dataGridViewManualScale.Rows.Count > 0)
                {
                    dataGridViewManualScale.Rows[dataGridViewManualScale.Rows.Count - 1].Selected = true;
                    dataGridViewManualScale.FirstDisplayedScrollingRowIndex = dataGridViewManualScale.Rows.Count - 1;
                }
            }


            // Check Avg And if it below Min Send a Email to QA
            checkAvgAndSendEmail(shiftRun);
            // Check Error Number and send Email to Tim about it
            checkErrorsAndSendEmailToTim(shiftRun);

            //
            // BARCODE CHECKER START
            //
            //Open BarCode checker Window if condition is true

            if (shiftRun.PlCount >= shiftRun.ManualCount) // For mostly from Main scale
            {

                if (shiftRun.nextCheckAt == shiftRun.PlCount && shiftRun.isBarcodeChecker == true)
                {
                    ThreadPool.QueueUserWorkItem(state => barCodeWindowShow());
                    //barCodeWindowShow(); // Execute pop up window function
                }

            }
            else { // For mostly from Manual scale

                if (shiftRun.ManualCount == shiftRun.PlCount && shiftRun.isBarcodeChecker == true)
                {
                    ThreadPool.QueueUserWorkItem(state => barCodeWindowShow());
                    //barCodeWindowShow(); // Execute pop up window function
                }
            }
            

            //
            // BARCODE CHECKER END
            //
        }


        // Button STOP 
        private void buttonStop_Click(object sender, EventArgs e)
        {
            DialogResult dialog = MessageBox.Show("Finish and save data?: YES \nContinue running: NO", "Action Required!", MessageBoxButtons.YesNo, MessageBoxIcon.Warning);
            if (dialog == DialogResult.Yes)
            {
                shiftRun.totalBreakTimeInSeconds = timerRunning.getTotalBreakTimeInSeconds();
                shiftRun.startTime = timerRunning.getStartintTime(); // Get Start Time
                ProductivityCalculator.calculateProductivity(shiftRun); // Calculate ptoductivity Run

                // SAVE info to Logs
                Logger.INFO("");
                Logger.INFO("");
                Logger.INFO("SAVING DATA");
                Logger.INFO("");
                Logger.INFO("");
                Logger.INFO("SKU: " + shiftRun.Sku);
                Logger.INFO("Barcode: " + shiftRun.barCode);
                Logger.INFO("Shift: " + shiftRun.Shift);
                Logger.INFO("Location: " + shiftRun.Location + " PL - " + shiftRun.PackLineNumber);

                Logger.INFO("Start Time: " + timerRunning.getStartintTime().ToString("h:mm tt"));
                Logger.INFO("Finish Time: " + DateTime.Now.ToString("h:mm tt"));
                Logger.INFO("Run Time: " + timerRunning.getRunInSeconds());
                Logger.INFO("Idle Time: " + timerRunning.getTotalBreakTimeInSeconds());
                Logger.INFO("Count: " + shiftRun.PlCount);
                Logger.INFO("Adjusted: " + shiftRun.ManualCount);
                Logger.INFO("Less: " + shiftRun.LessWeight);
                Logger.INFO("Target: " + shiftRun.TargetWeight);
                Logger.INFO("Heavy: " + shiftRun.HeavyWeight);

                if (shiftRun.ManualCount > 0 || shiftRun.PlCount > 0)
                {
                    Logger.INFO("Total Avg: " + (shiftRun.AverageWeight / (shiftRun.ManualCount + shiftRun.PlCount)).ToString("0.0")); // Data Total Avg
                }

                Logger.INFO("Given Away KG: " + (shiftRun.kgGivingAway * 0.001).ToString("0.000"));
                Logger.INFO("Given Away %: " + shiftRun.percentageGivingAway.ToString("0.0 %"));
                Logger.INFO("Units p/hr Target: " + shiftRun.productivityTarget);
                Logger.INFO("Units p/hr Actual: " + shiftRun.productivityActual);
                Logger.INFO("Running Efficiency: " + (shiftRun.runningEfficiency * 0.01).ToString("0.0 %"));
                Logger.INFO("Expected Efficiency: " + (shiftRun.expectedEfficiency * 0.01).ToString("0.0 %"));
                Logger.INFO("Staff Required: " + shiftRun.StaffNumberRequired);
                Logger.INFO("Staff Actual: " + shiftRun.StaffNumberActual);
                Logger.INFO("Productivity: " + (shiftRun.productivityRun * 0.01).ToString("0.0 %"));
                Logger.INFO("");
                Logger.INFO("PL Errors: " + shiftRun.errorCountPL);
                Logger.INFO("Manual Errors: " + shiftRun.errorCountManual);
                Logger.INFO("");
                Logger.INFO("////////////////////////////////////////////////////////////////////////////////////");


                if (checkExcelLibrary()) // If .DLL file not missing save data
                {
                    shiftRun.totalBreakTimeInSeconds = timerRunning.getTotalBreakTimeInSeconds(); // Total Break Time in seconds
                    buttonStop.Text = "Saving";
                    buttonStop.BackColor = Color.DarkGray;
                    buttonStart.Enabled = false;
                    //buttonSave.Enabled = false;
                    shiftRun.Running = false;

                    dataSave(); // Save data

                    buttonStart.Enabled = true;
                    buttonStart.Text = "Start";
                    buttonStart.BackColor = Color.DarkGreen;
                    buttonStop.Enabled = true;
                    buttonStop.Text = "STOP";
                    this.Invoke(new EventHandler(buttonClear_Click)); // Clear and ready for new Product
                }
                else // If .dll missing and can't automatically copy
                {
                    Logger.ERROR("Data can't save as ExcelLibrary.dll missing (Button STOP)");
                    MessageBox.Show("Data can't be saved till issue resolve \nPlease wait IT support", "ERROR", MessageBoxButtons.OK, MessageBoxIcon.Error);
                }
            }
            else if (dialog == DialogResult.No)
            {
                //buttonStart.Enabled = false;
                //buttonStart.Text = "Running";
                buttonStop.Enabled = true;
                buttonStop.BackColor = Color.Firebrick;
                //buttonSave.Enabled = true;
                shiftRun.Running = true;
            }
        }

        // When Form closed Close all open Ports
        private void Form1_FormClosed(object sender, FormClosedEventArgs e)
        {
            if (PLScaleSerialPort.IsOpen)
                PLScaleSerialPort.Close(); // Closing Port

            if (ManualScaleSerialPort.IsOpen)
                ManualScaleSerialPort.Close(); // Closing Port

            Logger.INFO("Form Closed");
        }


        private void timer1_Tick(object sender, EventArgs e)
        {
            labelTime.Text = DateTime.Now.ToLongTimeString();
            timerTime.Start();

            //Generate and send report
            if (shiftRun.timsTesting == true && shiftRun.isGenerateAndSend == true)
            {
                sendDailyReport();
            }


            // Daily report Sending at 6 am check time and if it'a 6 invoce a sendDailyReport funtion 
            if (shiftRun.autoGenerateReport == true && shiftRun.isDailyReportSent == false && DateTime.Now.Hour == shiftRun.sendReportAtHour && DateTime.Now.Minute == getAproperMinuteToSend())
            {
                shiftRun.isDailyReportSent = true;
                sendDailyReport();
            }

            if (DateTime.Now.Hour > 22 && DateTime.Now.Minute > 58)
            {
                shiftRun.isDailyReportSent = false;
            }
            


            //Console.WriteLine(timerRunning.runninTimeInSeconds());


            if (shiftRun.Running && shiftRun.isBreak == false && shiftRun.isTimerON == true)
            {
                stopInSeconds++;


                labelBreakTimeS.Visible = false;
                labelBreakTimeM.Visible = false;
                labelBreakTimeH.Visible = false;
                labelBreak.Visible = false;


                if (timerRunning.runninTimeInSeconds() > 0)
                {
                    shiftRun.runningTimeInSeconds = timerRunning.runninTimeInSeconds(); // Update Total Time in seconds

                    TimeSpan t = TimeSpan.FromSeconds(timerRunning.runninTimeInSeconds());

                    labelRunningTimeS.Text = String.Format("{0:00}", t.Seconds);
                    labelRunningTimeM.Text = String.Format("{0:00}.", t.Minutes);
                    labelRunningTimeH.Text = String.Format("{0:00}:", t.Hours);

                    //Console.WriteLine("Total seconds " + t.TotalSeconds);
                }
            }

            

            if (stopInSeconds >= 60) // Should be a 60  Switch to break after one minut not running
            {
                stopInSeconds = 0;
                //isTimerActive = false; // Stop Running Time if more then a min no product comming
                shiftRun.isTimerON = false; // Stop Running 
                shiftRun.isBreak = true; // Start break Time
                timerRunning.setIsBreak(true);
                
                //Console.WriteLine("Break Statement Reached");

            }

            if (shiftRun.isBreak && timerRunning.breakTimeInSeconds() > 0) // If break
            {
                // Show break time
                labelBreakTimeS.Visible = true;
                labelBreakTimeM.Visible = true;
                labelBreakTimeH.Visible = true;
                labelBreak.Visible = true;

                TimeSpan t1 = TimeSpan.FromSeconds(timerRunning.breakTimeInSeconds());

                //Console.WriteLine("Break time " + timerRunning.breakTimeInSeconds());

                labelBreakTimeS.Text = String.Format("{0:00}", t1.Seconds);
                labelBreakTimeM.Text = String.Format("{0:00}.", t1.Minutes);
                labelBreakTimeH.Text = String.Format("{0:00}:", t1.Hours);

                if (t1.Hours >= 2) // Save Data if not running for long time and stop Runing
                {
                    shiftRun.isBreak = false;
                    dataSave();
                }
            }

        }


        // Set SKU Weight 
        private void buttonSet_Click(object sender, EventArgs e)
        {
            SetSKUWeight setSKUWeight = new SetSKUWeight();
            
            buttonStart.Enabled = false; // Disable START button while SET UP window is opent
            buttonStart.BackColor = Color.DarkGray;

            setSKUWeight.ShiftRun = shiftRun;

            if (setSKUWeight.ShowDialog() == DialogResult.OK)
            {
                // Update Labels
                //labelPackLineNumberData.Text = shiftRun.PackLineNumber.ToString();
                labelSKUData.Text = shiftRun.Sku;
                labelLessData.Text = shiftRun.LessWeight.ToString("0.0 g");
                labelTargetData.Text = shiftRun.TargetWeight.ToString("0.0 g");
                labelHeavyData.Text = shiftRun.HeavyWeight.ToString("0.0 g");
                labelProdTargetData.Text = shiftRun.productivityTarget.ToString("0");

                // Invoke Start after Set click Start
                this.Invoke(new EventHandler(buttonOpen_Click));
            }
            else
            {
                setSKUWeight.Dispose();
                buttonStart.Enabled = true; 
                buttonStart.BackColor = Color.DarkGreen;
            }

            
        }


        // Clear function 
        private void buttonClear_Click(object sender, EventArgs e)
        {

            DeletAllObjects(); // Delete all objects

            // Creating a new objects
            unitsPerHour = new UnitsPerHour();
            timerRunning = new TimerRunning();
            shiftRun = new ShiftRun();
            validate = new DataValidator();
            manualScaleWeightCollection = new ManualScaleWeightCollection();
            pLScaleWeightCollection = new PLScaleWeightCollection();
            dynamicAverageCollection = new Queue<double>(); // Collection for Dynamic average

            // ShiftRun object assigning default values

            resetAllDataMembers();

            setBinding();
            setupDataGridViewPL();
            setupDataGridViewManual();

            Logger.INFO("");
            Logger.INFO("");
            Logger.INFO("All data members have been reset to Defaul values. Ready for new RUN");
            Logger.INFO("");
            Logger.INFO("");
        }

        // Delete all objects
        private void DeletAllObjects()
        {
            
            shiftRun = null;
            validate = null;
            manualScaleWeightCollection = null;
            pLScaleWeightCollection = null;
            dynamicAverageCollection = null;

            dataGridViewManualScale.DataSource = null; // DataGridView set to null
            dataGridViewManualScale.Rows.Clear(); // DataGridView remove all rows
            dataGridViewManualScale.Columns.Clear(); // DataGridView remove all columns
            dataGridViewManualScale.Refresh(); // DataGridView refresh

            dataGridViewPLScale.DataSource = null; // DataGridView set to null
            dataGridViewPLScale.Rows.Clear(); // DataGridView remove all rows
            dataGridViewPLScale.Columns.Clear(); // DataGridView remove all columns
            dataGridViewPLScale.Refresh(); // DataGridView refresh

            timerRunning = null; // Delete timer object
            
        }


        // Just for testing
        private void buttontest_Click(object sender, EventArgs e)
        {
            // Test Manual scale



            //shiftRun.DataFromManualScale = "G3000500.5g6";
            //this.Invoke(new EventHandler(UpdateData));

            //// Test PL scale

            // For scale error check
            //shiftRun.DataFromPLScale = "EB-41857<";
            //this.Invoke(new EventHandler(UpdateData));
            
            
            shiftRun.DataFromPLScale = "G3000501.0g6";
            this.Invoke(new EventHandler(UpdateData));


        }


        // When Click Close application
        // Check if data been saved, otherwise ask to save data before closing
        private void Scale_FormClosing(object sender, FormClosingEventArgs e)
        {
            if (dataSaved == true) // If data saved - close the application
            {
                DialogResult dialog = MessageBox.Show("Do you really want to close the program?", "Exit", MessageBoxButtons.YesNo, MessageBoxIcon.Warning);
                if (dialog == DialogResult.Yes)
                {
                    Logger.INFO("");
                    Logger.INFO("");
                    Logger.INFO("Application Closed. Data Been saved");
                    Logger.INFO("");
                    Logger.INFO("");
                    e.Cancel = false;
                }
                else if (dialog == DialogResult.No)
                {
                    e.Cancel = true;
                }
            }
            else // If Data not saved yet Ask to save it
            {
                DialogResult dialog = MessageBox.Show("Data NOT saved yet! \nDo you want to save before closing?", "Data NOT saved!", MessageBoxButtons.YesNo, MessageBoxIcon.Error);
                if (dialog == DialogResult.Yes)
                {
                    //this.Invoke(new EventHandler(buttonSave_Click));
                    dataSave();

                    Logger.INFO("");
                    Logger.INFO("");
                    Logger.INFO("Application Closed. Data Been saved");
                    Logger.INFO("");
                    Logger.INFO("");
                    e.Cancel = true;
                }
                else if (dialog == DialogResult.No) // Data not SAVED!!! And Application will be closed!!! WARNING  (DATA will be saved in LOGS)
                {
                    // SAVE info to Logs only
                    Logger.INFO("");
                    Logger.INFO("");
                    Logger.INFO("DATA NOT BEEN SAVED JUST IN LOGS");
                    Logger.INFO("");
                    Logger.INFO("");
                    Logger.INFO("SKU: " + shiftRun.Sku);
                    Logger.INFO("Barcode: " + shiftRun.barCode);
                    Logger.INFO("Shift: " + shiftRun.Shift);
                    Logger.INFO("Location: " + shiftRun.Location + " PL - " + shiftRun.PackLineNumber);

                    Logger.INFO("Start Time: " + timerRunning.getStartintTime().ToString("h:mm tt"));
                    Logger.INFO("Finish Time: " + DateTime.Now.ToString("h:mm tt"));
                    Logger.INFO("Run Time: " + timerRunning.getRunInSeconds());
                    Logger.INFO("Idle Time: " + timerRunning.getTotalBreakTimeInSeconds());
                    Logger.INFO("Count: " + shiftRun.PlCount);
                    Logger.INFO("Adjusted: " + shiftRun.ManualCount);
                    Logger.INFO("Less: " + shiftRun.LessWeight);
                    Logger.INFO("Target: " + shiftRun.TargetWeight);
                    Logger.INFO("Heavy: " + shiftRun.HeavyWeight);

                    if (shiftRun.ManualCount > 0 || shiftRun.PlCount > 0)
                    {
                        Logger.INFO("Total Avg: " + (shiftRun.AverageWeight / (shiftRun.ManualCount + shiftRun.PlCount)).ToString("0.0")); // Data Total Avg
                    }

                    Logger.INFO("Given Away KG: " + (shiftRun.kgGivingAway * 0.001).ToString("0.000"));
                    Logger.INFO("Given Away %: " + shiftRun.percentageGivingAway.ToString("0.0 %"));
                    Logger.INFO("Units p/hr Target: " + shiftRun.productivityTarget);
                    Logger.INFO("Units p/hr Actual: " + shiftRun.productivityActual);
                    Logger.INFO("Running Efficiency: " + (shiftRun.runningEfficiency * 0.01).ToString("0.0 %"));
                    Logger.INFO("Expected Efficiency: " + (shiftRun.expectedEfficiency * 0.01).ToString("0.0 %"));
                    Logger.INFO("Staff Required: " + shiftRun.StaffNumberRequired);
                    Logger.INFO("Staff Actual: " + shiftRun.StaffNumberActual);
                    Logger.INFO("");
                    Logger.INFO("PL Errors: " + shiftRun.errorCountPL);
                    Logger.INFO("Manual Errors: " + shiftRun.errorCountManual);
                    Logger.INFO("");
                    Logger.INFO("////////////////////////////////////////////////////////////////////////////////////");

                    Logger.INFO("");
                    Logger.INFO("");
                    Logger.INFO("Application Closed. Data NOT Been saved. Just in LOGS");
                    Logger.INFO("");
                    Logger.INFO("");

                    e.Cancel = false;
                }
            }

        }

        // Send IT support Request
        private void buttonIThelp_Click(object sender, EventArgs e)
        {
            Logger.WARN("Button IT support has been pressed");

            buttonIThelp.Enabled = false;
            buttonIThelp.Text = "Sending";

            try
            {
                buttonIThelp.Enabled = ThreadPool.QueueUserWorkItem(state => SendEmail.sendEmailToHD(shiftRun)); // Send email using separate thread
                buttonIThelp.Text = "IT Help";
            }
            catch (Exception ex)
            {
                Logger.ERROR("Exception thrown while IT support button been pressed "+ ex);
                MessageBox.Show(ex.Message, "Message", MessageBoxButtons.OK, MessageBoxIcon.Error);
                buttonIThelp.Enabled = true;
                buttonIThelp.Text = "IT Help";
            }
        }

        // Check Avg and if it below between Minimum and Target send Email to QA team
        private void checkAvgAndSendEmail(ShiftRun shiftRun)
        {
            if (shiftRun.emailToQAsent != true)
            {
                if ((shiftRun.PlCount + shiftRun.ManualCount) >= 101 && (dynamicAvg < (shiftRun.LessWeight + ((shiftRun.TargetWeight - shiftRun.LessWeight) / 2))))
                //if ((shiftRun.PlCount + shiftRun.ManualCount) >= 100 && (shiftRun.AverageWeight / (shiftRun.ManualCount + shiftRun.PlCount)) < (shiftRun.LessWeight + ((shiftRun.TargetWeight - shiftRun.LessWeight) / 2)))
                {
                    shiftRun.emailToQAsent = true;

                    ThreadPool.QueueUserWorkItem(state => SendEmail.sendEmailToQA(shiftRun)); // Send email using separate thread
                    //SendEmail.sendEmailToQA(shiftRun); // old way
                }
            }
        }


        // Check Errors N and send Email to Tim Start
        private void checkErrorsAndSendEmailToTim(ShiftRun shiftRun)
        {
            
            if (shiftRun.errorCountPL > numberOfErrorsWhenEmailSend || shiftRun.errorCountManual > numberOfErrorsWhenEmailSend)
            {
                string reason = shiftRun.Location + "Packline "+ shiftRun.PackLineNumber+"  Running with errors more than 15. Please check the scale";
                ThreadPool.QueueUserWorkItem(state => SendEmail.sendEmailToTim(shiftRun, reason)); // Send email using separate thread
                numberOfErrorsWhenEmailSend += 100;
            }
        }// Check Errors N and send Email to Tim Start

        // Settings 
        private void buttonSettinrgs_Click(object sender, EventArgs e)
        {
            Settings settings = new Settings(ManualScaleSerialPort, PLScaleSerialPort, shiftRun);



            if (settings.ShowDialog() == DialogResult.OK)
            {

                labelPackLineNumberData.Text = shiftRun.PackLineNumber.ToString();
                labelCHorKW.Text = shiftRun.Location;
                
                // For reporting
                shiftRun.autoGenerateReport = Properties.Settings.Default.autoGenerateReport; // Use preconfigured settings
                shiftRun.sendReportAtHour = Properties.Settings.Default.sendReportAtHour; // Use preconfigured settings
                shiftRun.sendReportAtMinute = Properties.Settings.Default.sendReportAtMinute; // Use preconfigured settings


                // For Testing ENABLE DISABLE buttons
                if (shiftRun.timsTesting == false)
                {
                    buttonClear.Visible = false;
                    buttontest.Visible = false;
                    buttonSet.Visible = false;
                }
                else
                {
                    buttonClear.Visible = true;
                    buttontest.Visible = true;
                    buttonSet.Visible = true;
                }


                try
                {
                    PLScaleSerialPort.PortName = Properties.Settings.Default.PLCOMsettings; // Use saved settings for COM port
                    PLScaleSerialPort.BaudRate = 9600;
                    PLScaleSerialPort.Parity = Parity.None;
                    PLScaleSerialPort.StopBits = StopBits.One;
                    PLScaleSerialPort.DataBits = 8;
                    PLScaleSerialPort.Open();
                    PlConnected(true);


                    ManualScaleSerialPort.PortName = Properties.Settings.Default.ManualCOMsettings;   // Use saved settings for COM port
                    ManualScaleSerialPort.BaudRate = 9600;
                    ManualScaleSerialPort.Parity = Parity.None;
                    ManualScaleSerialPort.StopBits = StopBits.One;
                    ManualScaleSerialPort.DataBits = 8;
                    ManualScaleSerialPort.Open();
                    ManualConnected(true);

                    // Warning Message NOT SHOW
                    shiftRun.Warning = false;
                    warningMessage(shiftRun.Warning);
                }
                catch (Exception ex)
                {
                    // Warning Message SHOW
                    shiftRun.Warning = true;
                    warningMessage(shiftRun.Warning);
                    MessageBox.Show(ex.Message, "Message", MessageBoxButtons.OK, MessageBoxIcon.Error);
                    ManualConnected(false);
                }

                PLportN.Text = settings.PLportNumber;
                Manualport.Text = settings.ManualportNumber;

            }

            settings.Dispose();
        }

        // Connected Labels
        private void PlConnected(bool connected)
        {
            if (connected)
            {
                labelPlConnected.Text = "Connected";
                labelPlConnected.ForeColor = System.Drawing.Color.Green;
            }
            else
            {
                labelPlConnected.Text = "Disconnected";
                labelPlConnected.ForeColor = System.Drawing.Color.Red;
            }
        }

        private void ManualConnected(bool connected)
        {
            if (connected)
            {
                LabelManualConnected.Text = "Connected";
                LabelManualConnected.ForeColor = System.Drawing.Color.Green;
            }
            else
            {
                LabelManualConnected.Text = "Disconnected";
                LabelManualConnected.ForeColor = System.Drawing.Color.Red;
            }
        }


        // Saving Data
        private void dataSave()
        {
            shiftRun.Running = false;
            buttonStop.Enabled = false;
            shiftRun.isBreak = false; // Break time count STOP
            if (shiftRun.PlCount > 0 || shiftRun.ManualCount > 0) // Try to save if only run not empty 
            {
                //shiftRun.startTime = timerRunning.getStartintTime(); // Get Start Time
                SaveData.saveDataforDayRepor(shiftRun, manualScaleWeightCollection, pLScaleWeightCollection); // Testing Dayly report

                // If Data saving with delay
                if (shiftRun.isDelaySaving)
                {
                    Console.WriteLine("ALL DATA WILL BE WITH DELAY SEVING");
                    dataSaved = SaveData.saveData(shiftRun, manualScaleWeightCollection, pLScaleWeightCollection); // Try to save and Return true if data saved
                    this.Invoke(new EventHandler(buttonClear_Click)); // Clear and ready for new Product 
                    //Task.Delay(25000).ContinueWith(t => this.Invoke(new EventHandler(buttonClear_Click))); // Clear and ready for new Product 
                }
                else
                {
                    dataSaved = SaveData.saveData(shiftRun, manualScaleWeightCollection, pLScaleWeightCollection); // Try to save and Return true if data saved
                    this.Invoke(new EventHandler(buttonClear_Click)); // Clear and ready for new Product 
                }
            }
            else
            {
                dataSaved = true;
                this.Invoke(new EventHandler(buttonClear_Click)); // Clear and ready for new Product 
            }
            
            //this.Invoke(new EventHandler(buttonClear_Click)); // Clear and ready for new Product 
        }

        // Exit X hover
        private void labelExit_MouseHover(object sender, EventArgs e)
        {
            labelExit.ForeColor = System.Drawing.Color.White;
        }

        private void labelExit_MouseLeave(object sender, EventArgs e)
        {
            labelExit.ForeColor = System.Drawing.Color.BlueViolet;
        }


        // Closing Link
        private void labelExit_Click(object sender, EventArgs e)
        {
            if (dataSaved == true) // If data saved - close the application
            {
                DialogResult dialog = MessageBox.Show("Do you really want to close the program?", "Exit", MessageBoxButtons.YesNo, MessageBoxIcon.Warning);
                if (dialog == DialogResult.Yes)
                {
                    //e.Cancel = false;
                    this.Close();
                }
                else if (dialog == DialogResult.No)
                {
                    //e.Cancel = true;
                }
            }
            else // If Data not saved yet Ask to save it
            {
                DialogResult dialog = MessageBox.Show("Data NOT saved yet! \nDo you want to save before closing?", "Data NOT saved!", MessageBoxButtons.YesNo, MessageBoxIcon.Error);
                if (dialog == DialogResult.Yes)
                {
                    //this.Invoke(new EventHandler(buttonSave_Click));
                    dataSave();
                    //e.Cancel = true;
                }
                else if (dialog == DialogResult.No)
                {
                    this.Close();
                    //e.Cancel = false;
                }
            }
        }

        // Minimize Form
        private void label8_Click(object sender, EventArgs e)
        {
            this.WindowState = FormWindowState.Minimized;
        }

        private void label8_MouseHover(object sender, EventArgs e)
        {
            labelMinimize.ForeColor = System.Drawing.Color.White;
        }

        private void label8_MouseLeave(object sender, EventArgs e)
        {
            labelMinimize.ForeColor = System.Drawing.Color.BlueViolet;
        }


        // WARNING! message
        // Worning Message If Scale not connected
        public void warningMessage(bool set, string message = "WARNING!\nPlease Press IT Help\nScales NOT connected")
        {
            //labelWarning.Text = "WARNING!\nPlease Press IT Help\nScales NOT connected";
            //labelWarning.Text = "WARNING!\nPlease Check Packaging.\nBarCode NOT Matching";
            labelWarning.Text = message;

            if (set == true) // Warning shows 
            {
                this.Invoke(new EventHandler(timer2_Tick)); // Invoke flashing Warning 
            }
            else // All good no warning
            {
                timerWarninMessage.Stop();
                labelWarning.Visible = false;
            }
        }
        // WARNING! message
        // Flashing warning message
        private void timer2_Tick(object sender, EventArgs e)
        {
            timerWarninMessage.Interval = 1200;
            timerWarninMessage.Enabled = true;
            timerWarninMessage.Start();

            if (labelWarning.Visible == true)
            {
                labelWarning.Visible = false;
            }
            else
            {
                labelWarning.Visible = true;
                labelWarning.BackColor = Color.White;
            }
        }

        // Check for ExelLibriary
        public bool checkExcelLibrary()
        {
            string currentPath = Directory.GetCurrentDirectory();

            // If ExcelLibrary.dll not Exist will be copy from here
            string sourcePath = @"\\hedgehog\Syteline\PacklineScaleData\ScaleProject\ScaleApp(Tim)\ExcelLibrary.dll";

            currentPath += @"\ExcelLibrary.dll";

            try
            {
                // Check if File exist
                if (System.IO.File.Exists(currentPath))
                {
                    // All Good File exist!

                    //MessageBox.Show("ExcelLibrary.dll exist! \nCurrent path is: " + currentPath, "Confirmation", MessageBoxButtons.OK, MessageBoxIcon.Information);
                    return true;
                }
                else
                {
                    // File not exist Will be copy 
                    //MessageBox.Show("Copying File Current Path "+currentPath, "Info", MessageBoxButtons.OK, MessageBoxIcon.Information);

                    System.IO.File.Copy(sourcePath, currentPath, true); // Copying File
                    return true;
                }
            }
            catch (Exception e)
            {
                // Create a HD tiket about Missing Dll file
                if (!shiftRun.Location.Equals(""))
                {
                    ThreadPool.QueueUserWorkItem(state => SendEmail.sendEmailToHDAboutMissingDLL(shiftRun)); // Send email in separate thread
                    //SendEmail.sendEmailToHDAboutMissingDLL(shiftRun);
                }

                // Can not copy ExcelLibrary.dll
                MessageBox.Show("Missing ExcelLibrary.dll\nIT will be notify about it now", "ERROR", MessageBoxButtons.OK, MessageBoxIcon.Error);
                return false;
            }

        }


        ///// Moving the Form
        private void panelTop_MouseDown(object sender, MouseEventArgs e)
        {
            mov = 1;
            movX = e.X;
            movY = e.Y;
        }

        private void panelTop_MouseMove(object sender, MouseEventArgs e)
        {
            if (mov == 1)
            {
                this.SetDesktopLocation(MousePosition.X - movX, MousePosition.Y - movY);
            }
        }

        private void panelTop_MouseUp(object sender, MouseEventArgs e)
        {
            mov = 0;
        }
        ///// END Moving the FORM


        // DoubleClick Top panel Maximized window
        private void panelTop_DoubleClick(object sender, EventArgs e)
        {
            if (this.WindowState == FormWindowState.Maximized)
            {
                this.WindowState = FormWindowState.Normal;
            }
            else
            {
                this.WindowState = FormWindowState.Maximized;
            }
        }


        // Method to calculate Dynamic Average
        private double calculateDynamicAvg(double weighToAdd)
        {
            double toDisplay = 0.0;

            if (weighToAdd != 0)
            {
                dynamicAverageCollection.Enqueue(weighToAdd);
            }

            if (dynamicAverageCollection.Count >= 100)
            {
                double temp = 0.0;
                foreach (var i in dynamicAverageCollection.ToArray())
                    temp += i;

                shiftRun.AverageWeightDynamic = temp / dynamicAverageCollection.Count;

                toDisplay = shiftRun.AverageWeightDynamic;
                dynamicAverageCollection.Dequeue();
            }
            return toDisplay;
        }


        
        // Check if it time to send a Daily Report
        public void sendDailyReport()
        {
            shiftRun.isGenerateAndSend = false;

            //MessageBox.Show("Report will be generated in " + delayforSending, "Confirmation", MessageBoxButtons.OK, MessageBoxIcon.Information);

            //First Generate report
            
            //Task.Delay(delayforSending).ContinueWith(t => ThreadPool.QueueUserWorkItem(state => isReportGenerated = SaveData.generateDailyReportFile(shiftRun)));
            SaveData.generateDailyReportFile(shiftRun);

            //Then Send email with 15 sesonds delay
            Task.Delay(15000).ContinueWith(t => ThreadPool.QueueUserWorkItem(state => SendEmail.sendDailyReport(shiftRun))); // Send Daily report in separate thread
            
        }

        // Get a Different send minute for each PL
        public int getAproperMinuteToSend()
        {
            int minute = 0;

            string checkLocation = shiftRun.Location + shiftRun.PackLineNumber;
            if (checkLocation.Equals("KW1"))
            {
                minute = shiftRun.sendReportAtMinute + 1;
            }
            if (checkLocation.Equals("CH1"))
            {
                minute = shiftRun.sendReportAtMinute + 2;
            }
            if (checkLocation.Equals("CH2"))
            {
                minute = shiftRun.sendReportAtMinute + 3;
            }
            if (checkLocation.Equals("CH3"))
            {
                minute = shiftRun.sendReportAtMinute + 4;
            }

            return minute;
        }


        //Bar Code checker Function to Prompt a bar code scanning window START
        public void barCodeWindowShow() {

            BarcodeCheckerForm barcodeCheckerForm = new BarcodeCheckerForm(shiftRun);

            if(barcodeCheckerForm.ShowDialog() == DialogResult.OK)
            {
                
                if (shiftRun.PlCount >= shiftRun.ManualCount)
                {
                    shiftRun.nextCheckAt = shiftRun.PlCount + shiftRun.barCodeCheckAtCount;
                }
                else
                {
                    shiftRun.nextCheckAt = shiftRun.ManualCount + shiftRun.barCodeCheckAtCount;
                }

                barcodeCheckerForm.Dispose();
            }


            // If BarCode not matching Show Error message and Notifying Management by Email.
            if (shiftRun.isBarCodeMatch == false)
            {

                //
                // SEND EMAIL START
                ThreadPool.QueueUserWorkItem(state => SendEmail.sendEmailBarCodeNotMatching(shiftRun));
                // SEND EMAIL END
                //

                warningMessage(true, "WARNING!\nPlease Check Packaging.\nBarCode NOT Matching");

                // If BarCode not match we will ask for checking again in after 5 boxes.
                if (shiftRun.PlCount >= shiftRun.ManualCount) {
                    shiftRun.nextCheckAt = shiftRun.PlCount + 5;
                } else {
                    shiftRun.ManualCount = shiftRun.ManualCount + 5;
                }
                

            }
            else {

                warningMessage(false);
            }

        }
        //Bar Code checker Function to Prompt a bar code scanning window END
    }


}
