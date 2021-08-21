using System;
using System.CodeDom;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace PortMainScaleTest
{
    class UnitsPerHour
    {
        static int ONE_HOUR = 60;
        static int countTemp = 0;
        static int secondsTems = 0;
        static bool nextHour = false;


        // If timer reach 60 min set to next hour
        public void setNextHour(bool toSet)
        {
            nextHour = toSet;
        }


        // Calculate UPH overall
        public int unitsPerHourOveral(int totalSecondsWorking, int count)
        {
            int result = 0;

            if(count > 0 && totalSecondsWorking > 0)
            result = count * 3600 / totalSecondsWorking;

            return result;
        
        }


        // Calculate UPH for last hour
        public int unitsPerHourLastHour(int totalSecondsWorking, int count)
        {
            int result = 0;


            // Result for first hour
            if (totalSecondsWorking > 0 && (totalSecondsWorking / 60) < ONE_HOUR)
            {

                countTemp = count;

                if (count > 0)
                    result = count * 3600 / totalSecondsWorking;

                //Console.WriteLine("First hour Units p/hr statement colculation Result: "+result);
                return result;
            }

            // Result for next hours
            if ((totalSecondsWorking / 60) >= ONE_HOUR)
            {
                //Console.WriteLine("Next hour! Units p/hr statement");

                if (nextHour == true) // If 120 min then gonna start a new hour
                {
                    countTemp = count;
                    secondsTems = totalSecondsWorking-1;
                    nextHour = false;
                }


                //if ((totalSecondsWorking / 60) % ONE_HOUR != 0)
                //{
                    result = (count - countTemp) * 3600 / (totalSecondsWorking - secondsTems);
                    //nextHour = false;
                    //Console.WriteLine("Next hour! Calculating!! Nice Units p/hr statement" + count + " - " + countTemp +" * 3600 / " + totalSecondsWorking + " - " +secondsTems + " = " + result);
                //}
                
                return result;
                
            }


            //Console.WriteLine("This should NOT appearrs!!! :( ");
            return 0;
        }

    }
}
