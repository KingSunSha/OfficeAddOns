using System;
using System.IO;
using System.Collections.Generic;
using System.ComponentModel;

namespace EditablePivot.BaseClasses
{

    public enum LogLevel
    {
        [Description("ERROR")]
        ERROR = 1,   // Fatal error 
        [Description("WARNING")]
        WARNING = 2, // Warning, typically an error but not fatal
        [Description("INFO")]
        INFO = 3,    // Just for information
        [Description("DEBUG")]
        DEBUG = 4    // Used for debugging
    }

    public class LogMessage
    {
        public DateTime logTime { get; set; }
        public string logText { get; set; }
        public LogLevel logLevel { get; set; }

        private LogMessage() { }

        public LogMessage(string text, LogLevel lvl = LogLevel.INFO)
        {
            logTime = DateTime.Now;
            logText = text;
            logLevel = lvl;
        }
    }

    public class LogMessageList : List<LogMessage>
    {
        private string m_logType = "";
        public string LogType { get { return m_logType; } set { m_logType = value; } }

        private string m_logName = "";
        public string LogName { get { return m_logName; } set { m_logName = value; } }

        private int error_count = 0;
        public int errors { get { return error_count; } }

        private int warning_count = 0;
        public int warnings { get { return warning_count; } }

        public LogMessageList() { }

        public LogMessageList(string logType, string logName)
        {
            this.LogType = logType;
            this.LogName = logName;
        }

        public void AddStringList(StringList lst, LogLevel lvl = LogLevel.INFO)
        {
            foreach (string text in lst)
            {
                this.Add(new LogMessage(text, lvl));
                if (lvl == LogLevel.ERROR)
                    error_count++;
                else if (lvl == LogLevel.WARNING)
                    warning_count++;
            }
        }

        public void Add(LogMessage msg)
        {
            base.Add(msg);
            if (msg.logLevel == LogLevel.ERROR)
                error_count++;
            else if (msg.logLevel == LogLevel.WARNING)
                warning_count++;
        }
    }

    public class Logging
    {
        // conditional constant for debug
        //#Const _debug = True

        private static FileStream fileStream;

        private static StreamWriter streamWriter;
        public static string logMsg = "";

        public static bool debugMode = false;
        public static void Log(string txt)
        {
            if (debugMode)
            {
                logMsg = logMsg + DateTime.Now.ToString() + "\t" + txt + "\n";
            }
        }

        private static void OpenFile(string logName)
        {
            string strPath = null;

            strPath = System.Environment.GetFolderPath(Environment.SpecialFolder.MyDocuments) + "\\" + logName + ".log";
            if (System.IO.File.Exists(strPath))
            {
                fileStream = new FileStream(strPath, FileMode.Append, FileAccess.Write);
            }
            else
            {
                fileStream = new FileStream(strPath, FileMode.Create, FileAccess.Write);
            }
            streamWriter = new StreamWriter(fileStream);
        }

        private static void CloseFile()
        {
            streamWriter.Close();
            fileStream.Close();
        }

        public static void WriteLog(string logName, string strComments)
        {
            OpenFile(logName);
            streamWriter.WriteLine(strComments);
            CloseFile();
        }

        public static void debugLog(string txt)
        {
            Log(txt);
        }
    }

}
