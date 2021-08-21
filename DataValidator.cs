using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;


namespace PortMainScaleTest
{
    class DataValidator
    {
        //public double validWaight = 0;

        public double validateWeight(string dataFromScale)
        {
            double weightToReturn = 0;
            try
            {

                if (dataFromScale.StartsWith("E")) // Validation for data from scale like "E10356579" (Old Scale)
                {
                    dataFromScale = dataFromScale.Substring(2, 5);
                    dataFromScale = dataFromScale.Insert(4, ".");
                    weightToReturn = Convert.ToDouble(dataFromScale);
                    //validWaight = 100.5;
                }

                if (dataFromScale.StartsWith("G")) // Validation for data from scale like "G3000423.5g6;" (New ISHIDA scale)
                {
                    int index = dataFromScale.IndexOf('.'); // New Validation way by finding a '.' For new ISHIDA scale 
                    dataFromScale = dataFromScale.Substring(index - 4, 6); // get 4 digits in front of '.' and 1 after 

                    //dataFromScale = dataFromScale.Substring(4, 6); // originally was this way  
                    dataFromScale = dataFromScale.Replace(".", string.Empty);
                    dataFromScale = dataFromScale.Insert(4, ".");
                    weightToReturn = Convert.ToDouble(dataFromScale);
                    //validWaight = 100.5;
                }

                if (dataFromScale.StartsWith("+")) // Validation for data from scale like "+00755.4g" (Manual Scale)
                {

                    dataFromScale = dataFromScale.Substring(2, 6);
                    dataFromScale = dataFromScale.Replace(".", string.Empty);

                    dataFromScale = dataFromScale.Insert(4, ".");
                    weightToReturn = Convert.ToDouble(dataFromScale);
                    //validWaight = 200.5;
                }

            }
            catch (Exception e)
            {
                Logger.ERROR("Exception thrown in DataValidator. RAW data from Scale: " + dataFromScale + "       Will be returned: "+ weightToReturn);
                return weightToReturn;
            }

            //MessageBox.Show(validWaight.ToString(), "Message", MessageBoxButtons.OK, MessageBoxIcon.Error);
            return weightToReturn;
        }
    }
}
