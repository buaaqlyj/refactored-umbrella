using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Reflection;
using System.Text;
using System.Windows.Forms;

namespace Statistics.Log.LogWriter
{
    public class LocalLogWriter: ILogWriter
    {
        private static StreamWriter logFile = null;
        private static int exceptionNum = 0;
        private static int dataerrorNum = 0;
        private static string currentFile = "";
        private static string runningPath = "";
        private static bool firstTime = true;

        private static Stack<string> stk = null;
        private static LocalLogWriter logHelper = new LocalLogWriter();

        private LocalLogWriter()
        {
            
        }

        #region Delegate
        public delegate void TextBoxWriteInvoke(string str);
        public static event TextBoxWriteInvoke tbwi;

        public static void AddLog(string pre, string ex, bool sw)
        {
            string temp = @"【" + pre + @"】" + ex;
            if (sw)
            {
                Log_WriteLine(temp);
            }
            tbwi(temp + Environment.NewLine);
        }

        public static void AddLog(string ex, bool sw)
        {
            if (sw)
            {
                Log_WriteLine(ex);
            }
            tbwi(ex + Environment.NewLine);
        }

        public static void AddDataError(string ex, bool log)
        {
            dataerrorNum++;
            if (exceptionNum + dataerrorNum == 1)
            {
                AddLog(@"信息18", "原文件名：" + currentFile, true);
            }
            AddLog(@"错误19", "  第" + dataerrorNum + "个数据错误：" + ex, log);
        }

        public static void AddException(string ex, bool log)
        {
            exceptionNum++;
            if (exceptionNum + dataerrorNum == 1)
            {
                AddLog(@"信息20", "原文件名：" + currentFile, true);
            }
            AddLog(@"错误01", "  第" + exceptionNum + "个格式错误：" + ex, log);
        }

        #endregion

        public static void Log_WriteLine(string log)
        {
            if (Ready)
            {
                logFile.WriteLine(log);
                logFile.Flush();
            }
        }

        public static void StartNew(string fileName)
        {
            currentFile = fileName;
            ResetError();
        }

        public static void ResetError()
        {
            exceptionNum = 0;
            dataerrorNum = 0;
            runningPath = "";
        }

        public static void Dispose()
        {
            try
            {
                logFile.Close();
            }
            catch
            {

            }
        }

        #region ILogHelper

        public ILogWriter getInstance()
        {
            return (ILogWriter)logHelper;
        }

        public void initial()
        {
            DirectoryInfo di = new DirectoryInfo(Application.StartupPath + @"\日志");
            if (!di.Exists)
            {
                di.Create();
            }
            logFile = new StreamWriter(Application.StartupPath + @"\日志\" + DateTime.Now.ToString(@"yyyyMMdd-HH-mm-ss") + @".txt");
        }

        public void record(string message, LogLevel logLevel)
        {
            throw new NotImplementedException();
        }
        #endregion
        #region Property

        public static bool Ready
        {
            get
            {
                return logFile != null;
            }
        }
        #endregion

    }


}
