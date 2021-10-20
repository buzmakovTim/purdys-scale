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
//using System.Windows.Input;

namespace PortMainScaleTest
{
    public partial class BarcodeCheckerForm : Form
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


        Timer timer1;
        ShiftRun shiftRun;
        public BarcodeCheckerForm(ShiftRun shiftRun)
        {
            InitializeComponent();
            this.shiftRun = shiftRun;
            Region = System.Drawing.Region.FromHrgn(CreateRoundRectRgn(0, 0, Width, Height, 20, 20));

            //Timer for flickering background color 
            timer1 = new Timer();
            timer1.Interval = 1000;
            timer1.Tick += new EventHandler(timer1_Tick);
            timer1.Start();

        }
       
        //
        private void BarCodeChecker_Load(object sender, EventArgs e)
        {

        }

        private void timer1_Tick(object sender, EventArgs e)
        {
            if (this.BackColor == Color.Firebrick)
            {
                this.BackColor = Color.White;
                label1.ForeColor = Color.Firebrick;
            }
            else {
                this.BackColor = Color.Firebrick;
                label1.ForeColor = Color.White;
            }


            //labelText.Text = counter.ToString();
        }


        private void textBoxBarCodeChecker_KeyUp(object sender, KeyEventArgs e)
        {
            if (e.KeyValue == 13)
            {
                if (shiftRun.barCode.Equals(textBoxBarCodeChecker.Text))
                {
                    DialogResult = DialogResult.OK;
                    this.shiftRun.isBarCodeMatch = true; // Barcode match nothing to alert
                    timer1.Stop();
                }
                else
                {
                    DialogResult = DialogResult.OK;
                    this.shiftRun.isBarCodeMatch = false; // Barcode NOT matching  !!!NEEDS TO BE ALERTED!!!
                    timer1.Stop();
                }
            }

        }

        //if (shiftRun.BarCode.Equals(textBoxBarCodeChecker.Text))
        //{
        //    DialogResult = DialogResult.OK;
        //}
        //else
        //{
        //    DialogResult = DialogResult.Cancel;
        //}
    }
}
