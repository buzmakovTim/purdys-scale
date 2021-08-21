using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace PortMainScaleTest
{
    class TimerRunning
    {

        static DateTime timeStart;
        static DateTime breakStart;
        static int breakTimeInSecond = 0;
        static int totalseconds = 0;
        static int totalBreakTimeInSeconds = 0;




        public DateTime startTimer()
        {
            timeStart = DateTime.Now;

            breakTimeInSecond = 0;
            totalseconds = 0;
            totalBreakTimeInSeconds = 0;

            return timeStart;

        }

        public DateTime getStartintTime()
        {
            return timeStart;
        }

        public DateTime breakTimer()
        {
            breakStart = DateTime.Now;

            return breakStart;
        }


        public int runninTimeInSeconds()
        {

            //if (!isBreak) // If not a break
            //{
                TimeSpan diff = DateTime.Now - timeStart;

                totalseconds = Convert.ToInt32(diff.TotalSeconds) - breakTimeInSecond;


                
            //}
            //if(isBreak) // If Break
            //{
            //    TimeSpan diff = DateTime.Now - timeStart;

            //    breakTimeInSeconds = Convert.ToInt32(diff.TotalSeconds) - totalseconds;

            //    return breakTimeInSeconds;
            //}

            return totalseconds;
        }


        public int breakTimeInSeconds()
        {

                TimeSpan diff = DateTime.Now - timeStart;

                breakTimeInSecond = Convert.ToInt32(diff.TotalSeconds) - totalseconds;

            TimeSpan diff2 = DateTime.Now - breakStart;
            totalBreakTimeInSeconds = breakTimeInSecond;


            return Convert.ToInt32(diff2.TotalSeconds);
        }



        public void setIsBreak(bool b)
        {
            if(b == true)
            breakTimer();
        }

        public int getRunInSeconds()
        {
            return totalseconds;
        }

        public int getBreakInSeconds()
        {
            return breakTimeInSecond;
        }

        public int getTotalBreakTimeInSeconds()
        {
            return totalBreakTimeInSeconds;
        }
    }
}
