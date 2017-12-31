using System;
using System.IO;

namespace AddIn
{
    /// <summary>
    /// LogHelper:Record Log
    /// Author:Marco
    /// </summary>
    public static class LogHelper
    {
        private static string LOG_PATH = "C:\\Outlook\\Log\\" + DateTime.Now.ToString("yyyyMMdd") + ".txt";
        private static string LOG_DIRECTORY = "C:\\Outlook\\Log";
        private static string CASE_LOG = "C:\\Outlook\\Log\\CaseLog.txt";
        public static void WriteCase(string subject)
        {
            CheckDirectory();
            File.AppendAllText(CASE_LOG, DateTime.Now.ToString("yyyy-MM-dd HH:mm:ss") + " " + subject + " +++++++++maybe lost" + Environment.NewLine);
        }
        public static void Write(LogType logType,string context)
        {
            CheckDirectory();
            File.AppendAllText(LOG_PATH,DateTime.Now.ToString("yyyy-MM-dd HH:mm:ss") + " " + logType.ToString() + ":" + context + Environment.NewLine);
        }
        public static void Write(Exception ex)
        {
            CheckDirectory();
            File.AppendAllText(LOG_PATH,string.Format(DateTime.Now.ToString("yyyy-MM-dd HH:mm:ss") + " " + LogType.Error.ToString() + 
                "Exception:StackTrack->{0};Message->{1}", (ex.StackTrace ?? "null").ToString(),
                (ex.Message ?? "null").ToString()) + Environment.NewLine);
        }
        public static void Write(Exception ex,string content)
        {
            CheckDirectory();
            File.AppendAllText(LOG_PATH, string.Format(DateTime.Now.ToString("yyyy-MM-dd HH:mm:ss") + " " + LogType.Error.ToString() +
                "Exception:StackTrack->{0};Message->{1}", (ex.StackTrace ?? "null").ToString(),
                (ex.Message ?? "null").ToString()) + Environment.NewLine + content + Environment.NewLine);
        }
        private static void CheckDirectory()
        {
            if (!Directory.Exists(LOG_DIRECTORY))
            {
                Directory.CreateDirectory(LOG_DIRECTORY);
            }
        }
    }
    public enum LogType
    {
        Debug = 0,
        Info = 1,
        Warn = 2,
        Error = 3,
        Fatal = 4
    }
}
