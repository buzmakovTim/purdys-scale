using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.IO;

namespace PortMainScaleTest
{
    public static class Logger
    {

        //private static string fileName = "LOGS_" + DateTime.Now.ToString("MMMM dd yyyy") + ".txt";
        //static string[] path = { @"SCALE_LOGS\", fileName };
        //static string fullPath = Path.Combine(path);

        //DEBUG: Additional information about application behavior for cases when that information is necessary to diagnose problems
        //INFO: Application events for general purposes
        //WARN: Application events that may be an indication of a problem
        //ERROR: Typically logged in the catch block a try/catch block, includes the exception and contextual data
        //FATAL: A critical error that results in the termination of an application
        //TRACE: Used to mark the entry and exit of functions, for purposes of performance profiling


        public static void DEBUG(string Message)
        {
             string fileName = "LOGS_" + DateTime.Now.ToString("MMMM dd yyyy") + ".txt";
             string[] path = { @"SCALE_LOGS\", fileName };
             string fullPath = Path.Combine(path);


            System.IO.Directory.CreateDirectory("SCALE_LOGS");
            //using (System.IO.StreamWriter w = System.IO.File.AppendText("LOGS_" + DateTime.Now.ToString("MMMM dd yyyy") + ".txt"))
            using (System.IO.StreamWriter w = System.IO.File.AppendText(fullPath))
            {
                w.Write("\n---------------------------------------------");
                w.Write("\nDEBUG : ");
                w.Write("{0} {1}", DateTime.Now.ToLongTimeString(), DateTime.Now.ToLongDateString());
                w.Write(" : {0}", Message);
                w.Write("\n---------------------------------------------");
            }
        }

        public static void INFO(string Message)
        {
            string fileName = "LOGS_" + DateTime.Now.ToString("MMMM dd yyyy") + ".txt";
            string[] path = { @"SCALE_LOGS\", fileName };
            string fullPath = Path.Combine(path);

            System.IO.Directory.CreateDirectory("SCALE_LOGS");
            using (System.IO.StreamWriter w = System.IO.File.AppendText(fullPath))
            {
                w.Write("\nINFO : ");
                w.Write("{0} {1}", DateTime.Now.ToLongTimeString(), DateTime.Now.ToLongDateString());
                w.Write(" : {0}", Message);
                
            }
        }

        public static void WARN(string Message)
        {
            string fileName = "LOGS_" + DateTime.Now.ToString("MMMM dd yyyy") + ".txt";
            string[] path = { @"SCALE_LOGS\", fileName };
            string fullPath = Path.Combine(path);

            System.IO.Directory.CreateDirectory("SCALE_LOGS");
            using (System.IO.StreamWriter w = System.IO.File.AppendText(fullPath))
            {
                w.Write("\r\nWARNING : ");
                w.Write("{0} {1}", DateTime.Now.ToLongTimeString(), DateTime.Now.ToLongDateString());
                w.Write(" : {0}", Message);
                
            }
        }

        public static void ERROR(string Message)
        {
            string fileName = "LOGS_" + DateTime.Now.ToString("MMMM dd yyyy") + ".txt";
            string[] path = { @"SCALE_LOGS\", fileName };
            string fullPath = Path.Combine(path);

            System.IO.Directory.CreateDirectory("SCALE_LOGS");
            using (System.IO.StreamWriter w = System.IO.File.AppendText(fullPath))
            {
                w.Write("\n=====");
                w.Write("\r\nERROR : ");
                w.Write("{0} {1}", DateTime.Now.ToLongTimeString(), DateTime.Now.ToLongDateString());
                w.Write(" : {0}", Message);
                w.Write("\n=====");

            }
        }

        public static void FATAL(string Message)
        {
            string fileName = "LOGS_" + DateTime.Now.ToString("MMMM dd yyyy") + ".txt";
            string[] path = { @"SCALE_LOGS\", fileName };
            string fullPath = Path.Combine(path);

            System.IO.Directory.CreateDirectory("SCALE_LOGS");
            using (System.IO.StreamWriter w = System.IO.File.AppendText(fullPath))
            {
                w.Write("\n===============================================================");
                w.Write("\r\nFATAL : ");
                w.Write("{0} {1}", DateTime.Now.ToLongTimeString(), DateTime.Now.ToLongDateString());
                w.Write(" : {0}", Message);
                w.Write("\n===============================================================");
            }
        }

    }
}
