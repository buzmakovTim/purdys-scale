using System;
using System.Collections.Generic;
using System.Linq;
using System.Security.Permissions;
using System.Text;
using System.Threading.Tasks;

namespace PortMainScaleTest
{
    public class ShiftRun
    {
        
        public string DataFromPLScale { get; set; }
        public string DataFromManualScale { get; set; }
        public int PlCount { get; set; }
        public int ManualCount { get; set; }

        public string Sku { get; set; }

        public int StaffNumberRequired { get; set; }
        public int StaffNumberActual { get; set; }

        public string Shift { get; set; }
        public double LessWeight { get; set; }
        public double TargetWeight { get; set; }
        public double HeavyWeight { get; set; }

        public int LessCount { get; set; }
        public int TargetCount { get; set; }
        public int HeavyCount { get; set; }
        public int PackLineNumber { get; set; }

        public string Location { get; set; }

        public double AverageWeight { get; set; }

        public double AverageWeightDynamic { get; set; }
        public double AverageWeightCount { get; set; }
        public double AverageWeightAdjusted { get; set; }
        public double AverageWeightLess { get; set; }
        public double AverageWeightTarget { get; set; }
        public double AverageWeightHeavy { get; set; }

        public bool Running { get; set; }

        public bool saveToNEWformat { get; set; }

        public bool timsTesting { get; set; }

        public bool Warning { get; set; }

        public bool emailToQAsent { get; set; }
        public double kgGivingAway { get; set; }
        public double percentageGivingAway { get; set; }

        public double productivityRun { get; set; } // Productivity RUN Percentage
        public int productivityTarget { get; set; } // Productivity Target UPH units per hour
        public int productivityActual { get; set; } // Productivity Actual UPH units per hour

        public int runningEfficiency { get; set; } // Running Efficiency Actual Percentage
        public int expectedEfficiency { get; set; } // Expected Efficiency Percentage

        public int runningTimeInSeconds { get; set; } // Total running time in Seconds

        public int totalBreakTimeInSeconds { get; set; } // Total break time in Seconds
        public bool isBreak { get; set; } // Is Break 

        public int errorCountPL { get; set; } // total Error from PL scale
        public int errorCountManual { get; set; } // total Error from Manual scale

        public bool boxOverSize { get; set; } // For Oversize boxes

        public bool isTimerON { get; set; } // Timer is ON

        public bool isDelaySaving { get; set; } // If Daly Report file open will be delay for saving

        public bool isDailyReportSent { get; set; } // True is Day Report has been sent
        public bool isGenerateAndSend { get; set; } // If true we can generate and sent the Daily report for testing

        //For Reporting Start
        public bool autoGenerateReport { get; set; } // True if for auto reporting
        public int sendReportAtHour { get; set; } // Send report at this Hour
        public int sendReportAtMinute { get; set; } // Send report at this Minute
        public DateTime startTime { get; set; } // When Run starts


        //For Bar Code 
        public string BarCode { get; set; } // Bar code
        public bool isBarCodeMatch { get; set; } // Is Bar code Match?
        public int barCodeCheckAtCount { get; set; } // Check when count
                                                
    
    }
}
