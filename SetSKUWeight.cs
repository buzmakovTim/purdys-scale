using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;
using System.Runtime.InteropServices;

namespace PortMainScaleTest
{
    public partial class SetSKUWeight : Form
    {

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

        public ShiftRun ShiftRun { get; set; }
        public SetSKUWeight()
        {
            

            InitializeComponent();
            Region = System.Drawing.Region.FromHrgn(CreateRoundRectRgn(0, 0, Width, Height, 20, 20));

            if (Properties.Settings.Default.isBarcodeChecker)
            {
                textBoxBarcode.Visible = true;
                labelSetBacCode.Visible = true;
                labelBarcode6.Visible = true;
            }
            else
            {
                textBoxBarcode.Visible = false;
                labelSetBacCode.Visible = false;
                labelBarcode6.Visible = false;
            }

        }

        

        // Click START
        private void buttonSaveSetSKU_Click(object sender, EventArgs e)
        {
            // Need to validate if Data is empty

            //ShiftRun.PackLineNumber = Convert.ToInt32(comboBoxPLNumber.SelectedItem);
            ShiftRun.Sku = textBoxSKU.Text.ToUpper();

            try // Check if Weight input correctly
            {
                ShiftRun.LessWeight = Convert.ToDouble(textBoxLess.Text);
                ShiftRun.TargetWeight = Convert.ToDouble(textBoxTarget.Text);
                ShiftRun.HeavyWeight = Convert.ToDouble(textBoxHeavy.Text);
                ShiftRun.productivityTarget = Convert.ToInt32(textBoxProdTarget.Text);
                ShiftRun.barCode = textBoxBarcode.Text;

                // Selecting Number of people working
                // And validating that min N of people positive number and not less than 1
                int requiredPeople = Convert.ToInt32(comboBoxPplExpected.Text);
                int actualPeople = Convert.ToInt32(comboBoxPplActual.Text);

                if (requiredPeople >= 1 && actualPeople >= 1)
                {
                    ShiftRun.StaffNumberRequired = requiredPeople;
                    ShiftRun.StaffNumberActual = actualPeople;
                }
                else {
                    ShiftRun.StaffNumberRequired = 1;
                    ShiftRun.StaffNumberActual = 1;
                }
                

                if (!ShiftRun.Shift.Equals(""))
                {
                    if (ShiftRun.LessWeight <= ShiftRun.TargetWeight)
                    {
                        if (ShiftRun.TargetWeight <= ShiftRun.HeavyWeight)
                        {
                            if (ShiftRun.productivityTarget >= 10) // Min for Units p/hr 
                            {
                                DialogResult = DialogResult.OK;
                            }
                            else
                            {
                                MessageBox.Show("Units p/hr target can't be less than 10", "Warning!", MessageBoxButtons.OK, MessageBoxIcon.Warning);
                            }
                        }
                        else
                        {
                            MessageBox.Show("TARGET cannot be greater than HEAVY", "Warning!", MessageBoxButtons.OK, MessageBoxIcon.Warning);
                        }
                    }
                    else
                    {
                        MessageBox.Show("LESS cannot be greater than TARGET", "Warning!", MessageBoxButtons.OK, MessageBoxIcon.Warning);
                    }
                    
                }
                else
                {
                    if(ShiftRun.Shift.Equals(""))
                    MessageBox.Show("SHIFT must be selected\nPlease select one option", "Warning!", MessageBoxButtons.OK, MessageBoxIcon.Warning);
                }

            }
            catch (Exception ex) // Weight Input Incorrect or KPI not set
            {
                MessageBox.Show("Weight input incorrect or Units p/hr not set correctly", "Error!", MessageBoxButtons.OK, MessageBoxIcon.Error);
            }

        }

        private void buttonCancelSKUset_Click(object sender, EventArgs e)
        {
            DialogResult = DialogResult.Cancel;
        }

        private void SetSKUWeight_Load(object sender, EventArgs e)
        {
            // Adding to combo box number of people working
            int i = 1;
            while (i < 21)
            {
                comboBoxPplExpected.Items.Add(i);
                comboBoxPplActual.Items.Add(i);
                i++;
            }
            //comboBoxPplExpected.SelectedIndex = 0;
            //comboBoxPplActual.SelectedIndex = 0;




            ////Load form with saved settings for PackLine N
            //comboBoxPLNumber.SelectedItem = Properties.Settings.Default.packLineNumber;
            //comboBoxCHKW.SelectedItem = Properties.Settings.Default.locationCHKW;

            textBoxSKU.Text = ShiftRun.Sku;
            textBoxLess.Text = ShiftRun.LessWeight.ToString();
            textBoxTarget.Text = ShiftRun.TargetWeight.ToString();
            textBoxHeavy.Text = ShiftRun.HeavyWeight.ToString();
        }

        private void textBoxSKU_Click(object sender, EventArgs e)
        {
            textBoxSKU.SelectAll(); // Select all data after select textBox
        }

        private void textBoxLess_Click(object sender, EventArgs e)
        {
            textBoxLess.SelectAll(); // Select all data after select textBox
        }

        private void textBoxTarget_Click(object sender, EventArgs e)
        {
            textBoxTarget.SelectAll(); // Select all data after select textBox
        }

        private void textBoxHeavy_Click(object sender, EventArgs e)
        {
            textBoxHeavy.SelectAll(); // Select all data after select textBox
        }

        private void textBoxProdTarget_Click(object sender, EventArgs e)
        {
            textBoxProdTarget.SelectAll(); // Select all data after select textBox
        }


        private void buttonAM_Click(object sender, EventArgs e)
        {
            buttonAM.Enabled = false;
            buttonPM.Enabled = true;
            buttonGRV.Enabled = true;
            ShiftRun.Shift = "AM";
            pictureBoxAM.Visible = true;
            pictureBoxPM.Visible = false;
            pictureBoxGRV.Visible = false;
        }
        private void buttonPM_Click(object sender, EventArgs e)
        {
            buttonAM.Enabled = true;
            buttonPM.Enabled = false;
            buttonGRV.Enabled = true;
            ShiftRun.Shift = "PM";
            pictureBoxAM.Visible = false;
            pictureBoxPM.Visible = true;
            pictureBoxGRV.Visible = false;
        }

        private void buttonGRV_Click(object sender, EventArgs e)
        {
            buttonAM.Enabled = true;
            buttonPM.Enabled = true;
            buttonGRV.Enabled = false;
            ShiftRun.Shift = "Graveyard";
            pictureBoxAM.Visible = false;
            pictureBoxPM.Visible = false;
            pictureBoxGRV.Visible = true;
        }
    }
}
