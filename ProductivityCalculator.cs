using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;

namespace PortMainScaleTest
{
    public static class ProductivityCalculator
    {

        private const int COFFE_BREAK = 900; // 15 min in seconds
        private const int LUNCH = 1800; // 30 min in seconds

        // CH1 
        const string CH1_KB1_AM = "09:30";
        const string CH1_L_AM = "12:30";
        const string CH1_KB2_AM = "14:45";

        const string CH1_KB1_PM = "17:15";
        const string CH1_L_PM = "19:30";
        const string CH1_KB2_PM = "22:00";

        // CH2 
        const string CH2_KB1_AM = "09:15";
        const string CH2_L_AM = "12:00";
        const string CH2_KB2_AM = "14:00";

        const string CH2_KB1_PM = "17:45";
        const string CH2_L_PM = "20:00";
        const string CH2_KB2_PM = "23:15";

        // CH3 
        const string CH3_KB1_AM = "09:00";
        const string CH3_L_AM = "11:30";
        const string CH3_KB2_AM = "13:45";

        const string CH3_KB1_PM = "17:00";
        const string CH3_L_PM = "19:00";
        const string CH3_KB2_PM = "21:45";

        // KW1 
        const string KW1_KB1_AM = "09:30";
        const string KW1_L_AM = "12:30";
        const string KW1_KB2_AM = "15:15";

        const string KW1_KB1_PM = "17:45";
        const string KW1_L_PM = "20:00";
        const string KW1_KB2_PM = "22:15";




        static TimeSpan FIRST_COFFEE_TIME_AM;
        static TimeSpan LUNCH_TIME_AM;
        static TimeSpan SECOND_COFFEE_TIME_AM;

        static TimeSpan FIRST_COFFEE_TIME_PM;
        static TimeSpan LUNCH_TIME_PM;
        static TimeSpan SECOND_COFFEE_TIME_PM;

        public static void calculateProductivity(ShiftRun shiftRun) {

            int totalIdleWithoutBrakes = shiftRun.totalBreakTimeInSeconds; // Total seconds lane not run
            string pl = shiftRun.Location + shiftRun.PackLineNumber; // Exp "KW1"  

            calculateBrakeTimesForLane(pl); // Calculate brake times for different PL number

            TimeSpan NOW = DateTime.Now.TimeOfDay;
            TimeSpan START_TIME = shiftRun.startTime.TimeOfDay;

            double totlRunDuration = (NOW - START_TIME).TotalSeconds;

            // For AM Shift
            if (START_TIME < FIRST_COFFEE_TIME_AM && NOW > FIRST_COFFEE_TIME_AM) {
                totalIdleWithoutBrakes -= COFFE_BREAK;
            }
            if (START_TIME < LUNCH_TIME_AM && NOW > LUNCH_TIME_AM) {
                totalIdleWithoutBrakes -= LUNCH;
            }
            if (START_TIME < SECOND_COFFEE_TIME_AM && NOW > SECOND_COFFEE_TIME_AM)
            {
                totalIdleWithoutBrakes -= COFFE_BREAK;
            }

            //For PM Shift
            if (START_TIME < FIRST_COFFEE_TIME_PM && NOW > FIRST_COFFEE_TIME_PM)
            {
                totalIdleWithoutBrakes -= COFFE_BREAK;
            }
            if (START_TIME < LUNCH_TIME_PM && NOW > LUNCH_TIME_PM)
            {
                totalIdleWithoutBrakes -= LUNCH;
            }
            if (START_TIME < SECOND_COFFEE_TIME_PM && NOW > SECOND_COFFEE_TIME_PM)
            {
                totalIdleWithoutBrakes -= COFFE_BREAK;
            }

            shiftRun.productivityRun = 100 - (totalIdleWithoutBrakes * 100) / totlRunDuration;
        }


        static void calculateBrakeTimesForLane(String pl) {

            if (pl.Equals("CH1")) {
                setTimesForBrakes(CH1_KB1_AM, CH1_L_AM, CH1_KB2_AM, CH1_KB1_PM, CH1_L_PM, CH1_KB2_PM);
            }
            if (pl.Equals("CH2"))
            {
                setTimesForBrakes(CH2_KB1_AM, CH2_L_AM, CH2_KB2_AM, CH2_KB1_PM, CH2_L_PM, CH2_KB2_PM);
            }
            if (pl.Equals("CH3"))
            {
                setTimesForBrakes(CH3_KB1_AM, CH3_L_AM, CH3_KB2_AM, CH3_KB1_PM, CH3_L_PM, CH3_KB2_PM);
            }
            if (pl.Equals("KW1"))
            {
                setTimesForBrakes(KW1_KB1_AM, KW1_L_AM, KW1_KB2_AM, KW1_KB1_PM, KW1_L_PM, KW1_KB2_PM);
            }


            void setTimesForBrakes(String coffee1_AM, String lunch_AM, String coffee2_AM, String coffee1_PM, String lunch_PM, String coffee2_PM) {
                FIRST_COFFEE_TIME_AM = TimeSpan.Parse(coffee1_AM);
                LUNCH_TIME_AM = TimeSpan.Parse(lunch_AM);
                SECOND_COFFEE_TIME_AM = TimeSpan.Parse(coffee2_AM);

                FIRST_COFFEE_TIME_PM = TimeSpan.Parse(coffee1_PM);
                LUNCH_TIME_PM = TimeSpan.Parse(lunch_PM);
                SECOND_COFFEE_TIME_PM = TimeSpan.Parse(coffee2_PM);
            }

        } 
    }
}
