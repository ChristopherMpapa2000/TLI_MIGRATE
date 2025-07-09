using System;
using System.Collections.Generic;
using System.Configuration;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace MigrateTLI.Item
{
    class ItemConfig
    {
        public static string dbConnectionString
        {
            get
            {
                var ServarName = ConfigurationManager.AppSettings["ServarName"];
                var Database = ConfigurationManager.AppSettings["Database"];
                var Username_database = ConfigurationManager.AppSettings["Username_database"];
                var Password_database = ConfigurationManager.AppSettings["Password_database"];
                var dbConnectionString = $"data source={ServarName};initial catalog={Database};persist security info=True;user id={Username_database};password={Password_database};Connection Timeout=200";

                if (!string.IsNullOrEmpty(dbConnectionString))
                {
                    return dbConnectionString;
                }
                return string.Empty;
            }
        }
        public static string LogFile
        {
            get
            {
                var LogFile = System.Configuration.ConfigurationSettings.AppSettings["LogFile"];
                if (!string.IsNullOrEmpty(LogFile))
                {
                    return (LogFile);
                }
                return string.Empty;
            }
        }
        public static string Admin
        {
            get
            {
                var Admin = System.Configuration.ConfigurationSettings.AppSettings["Admin"];
                if (!string.IsNullOrEmpty(Admin))
                {
                    return (Admin);
                }
                return string.Empty;
            }
        }
        public static int TemplateID
        {
            get
            {
                var iTemplateID = ConfigurationManager.AppSettings["Template_Darnew"];
                if (!string.IsNullOrEmpty(iTemplateID))
                {
                    return Convert.ToInt32(iTemplateID);
                }
                return 0;
            }
        }
        public static string PathFileExcel
        {
            get
            {
                var PathFileExcel = ConfigurationManager.AppSettings["PathFileExcel"];

                if (!string.IsNullOrEmpty(PathFileExcel))
                {
                    return PathFileExcel;
                }
                return string.Empty;
            }
        }
        public static string SheetName_Migrate
        {
            get
            {
                var SheetName_Migrate = ConfigurationManager.AppSettings["SheetName_Migrate"];

                if (!string.IsNullOrEmpty(SheetName_Migrate))
                {
                    return SheetName_Migrate;
                }
                return string.Empty;
            }
        }
        public static string SourcePath
        {
            get
            {
                var SourcePath = ConfigurationManager.AppSettings["SourcePath"];

                if (!string.IsNullOrEmpty(SourcePath))
                {
                    return SourcePath;
                }
                return string.Empty;
            }
        }
        public static string TargetPath
        {
            get
            {
                var TargetPath = ConfigurationManager.AppSettings["TargetPath"];

                if (!string.IsNullOrEmpty(TargetPath))
                {
                    return TargetPath;
                }
                return string.Empty;
            }
        }
        public static string Status
        {
            get
            {
                var Status = ConfigurationManager.AppSettings["Status"];

                if (!string.IsNullOrEmpty(Status))
                {
                    return Status;
                }
                return string.Empty;
            }
        }
        public static string WaitForApprove
        {
            get
            {
                var WaitForApprove = ConfigurationManager.AppSettings["WaitForApprove"];

                if (!string.IsNullOrEmpty(WaitForApprove))
                {
                    return WaitForApprove;
                }
                return string.Empty;
            }
        }
        public static int company
        {
            get
            {
                var company = ConfigurationManager.AppSettings["company"];
                if (!string.IsNullOrEmpty(company))
                {
                    return Convert.ToInt32(company);
                }
                return 0;
            }
        }
        public static int DigitCon
        {
            get
            {
                var Digit_ControlRunning = ConfigurationManager.AppSettings["Digit_ControlRunning"];
                if (!string.IsNullOrEmpty(Digit_ControlRunning))
                {
                    return Convert.ToInt32(Digit_ControlRunning);
                }
                return 0;
            }
        }
    }
}
