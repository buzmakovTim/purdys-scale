using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace PortMainScaleTest
{
    public static class BarCodeValidator
    {

        //Set BarCode
        public static void setBarCode(ShiftRun shiftRun, string barCode)
        {

            if (!barCode.Equals(""))
            {
                shiftRun.barCode = barCode;
            }
            else {
                shiftRun.barCode = "";
            }        
        }
    }
}
