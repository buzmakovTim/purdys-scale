using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;
//using System.Windows.Input;

namespace PortMainScaleTest
{
    public partial class BarcodeCheckerForm : Form
    {

        ShiftRun shiftRun;
        public BarcodeCheckerForm(ShiftRun shiftRun)
        {
            InitializeComponent();
            this.shiftRun = shiftRun;
        }

        //
        private void BarCodeChecker_Load(object sender, EventArgs e)
        {

        }

       

        private void textBoxBarCodeChecker_KeyUp(object sender, KeyEventArgs e)
        {
            if (e.KeyValue == 13)
            {
                if (shiftRun.BarCode.Equals(textBoxBarCodeChecker.Text))
                {
                    DialogResult = DialogResult.OK;
                    this.shiftRun.isBarCodeMatch = true; // Barcode match nothing to alert
                }
                else
                {
                    DialogResult = DialogResult.OK;
                    this.shiftRun.isBarCodeMatch = false; // Barcode NOT matching  !!!NEEDS TO BE ALERTED!!!
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
