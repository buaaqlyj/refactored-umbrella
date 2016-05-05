using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Text;

using Statistics.Log;

namespace Statistics.Log
{
    public static class LogHelper
    {
        //private static LogForNet logWriter = null;
        private static StreamWriter logFile = null;

        //Job_Related : small
        private static int exceptionCount = 0;
        private static int dataerrorCount = 0;
        private static string currentFile = "";

        //Task_Related : big
        private static List<string> problemFilesList;
        private static int doneCount;
        private static int totalCount;
        private static double progressStep;
        private static double progressValue;

        //Testing Stack
        public static Stack<string> State = new Stack<string>();

        #region Public Interface
        public static void Initial(string filename)
        {
            //LogHelper.logWriter = logWriter;
            LogHelper.logFile = new StreamWriter(filename);
            //if (LogHelper.logWriter != null)
            //{
            //    LogHelper.logWriter.Initial();
            //}
            if (problemFilesList == null)
            {
                problemFilesList = new List<string>();
            }
        }

        public static void Close()
        {
            logFile.Close();
            logFile.Dispose();
        }
        #endregion

        #region Job_Related
        public static void AddException(string ex, bool log)
        {
            exceptionCount++;
            if (IsFirstError)
            {
                AddLog(@"信息20", "原文件名：" + currentFile, true);
            }
            AddLog(@"错误01", "  第" + ExceptionCount.ToString() + "个格式错误：" + ex, log);
        }

        public static void AddDataError(string ex, bool log)
        {
            dataerrorCount++;
            if (IsFirstError)
            {
                AddLog(@"信息18", "原文件名：" + currentFile, true);
            }
            AddLog(@"错误19", "  第" + DataErrorCount.ToString() + "个数据错误：" + ex, log);
        }

        public static void ResetError()
        {
            exceptionCount = 0;
            dataerrorCount = 0;
            UpdateFileNameDisplay("");
        }

        public static void StartNewJob(string filename)
        {
            ResetError();
        }

        public static void ReportSuspiciousFiles(Dictionary<string, int> suspFiles, string filename)
        {
            AddException("发现" + suspFiles.Count + "个可疑匹配项，暂不作任何处理", true);
            AddLog("可疑15", "    原文件名：" + filename, true);
            foreach (KeyValuePair<string, int> item in suspFiles)
            {
                AddLog("可疑16", "      文件名：" + item.Key, true);
                AddLog("可疑17", "        可疑指数：" + item.Value + " %", true);
            }
        }

        public static bool HasError
        {
            get
            {
                return HasException || HasDataError;
            }
        }

        public static bool HasException
        {
            get
            {
                return exceptionCount > 0;
            }
        }

        public static bool HasDataError
        {
            get
            {
                return dataerrorCount > 0;
            }
        }

        public static bool IsFirstError
        {
            get
            {
                return 1 == exceptionCount + dataerrorCount;
            }
        }

        public static int DataErrorCount
        {
            get
            {
                return dataerrorCount;
            }
        }

        public static int ExceptionCount
        {
            get
            {
                return exceptionCount;
            }
        }
        #endregion

        #region Task_Related
        public static void StartNewTask(int count)
        {
            if (count < 1) throw new Exception("产生的新任务为空！");
            ResetProblemFiles();
            LogHelper.totalCount = count;
            LogHelper.doneCount = 0;
            LogHelper.progressValue = 0;
            LogHelper.progressStep = 100.0 / (double)count;
            UpdateProgress(LogHelper.progressValue);
        }

        public static void FinishOneJob()
        {
            LogHelper.doneCount++;
            LogHelper.progressValue += LogHelper.progressStep;
            UpdateProgress(LogHelper.progressValue);
            if (LogHelper.HasDataError)
            {
                LogHelper.AddLog(@"***************************************************************", true);
            }
            ResetError();
        }

        public static void AddProblemFilesAndReset(string file)
        {
            problemFilesList.Add(file);
            AddLog(@"***************************************************************", true);
            ResetError();
        }

        public static int ProblemFilesCount
        {
            get
            {
                return problemFilesList.Count;
            }
        }

        public static List<string> ProblemFiles
        {
            get
            {
                return problemFilesList;
            }
        }

        public static int DoneCount
        {
            get
            {
                return doneCount;
            }
        }

        public static int TotalCount
        {
            get
            {
                return totalCount;
            }
        }
        #endregion

        #region LogEvent
        public static event AddLogHandler AddLogEvent;
        public delegate void AddLogHandler(string ex);

        public static void AddLog(string pre, string ex, bool sw)
        {
            string temp = @"【" + pre + @"】" + ex;
            if (sw)
            {
                WriteLogItem(temp);
            }
            AddLogEvent(temp);
        }

        public static void AddLog(string ex, bool sw)
        {
            if (sw)
            {
                WriteLogItem(ex);
            }
            AddLogEvent(ex);
        }

        private static void WriteLogItem(string log)
        {
            if (logFile != null)
            {
                logFile.WriteLine(log);
                logFile.Flush();
            }
        }
        #endregion

        #region ProgressEvent
        public static event UpdateProgressHandler UpdateProgressEvent;
        public delegate void UpdateProgressHandler(double progress);

        private static void UpdateProgress(double progress)
        {
            UpdateProgressEvent(progress);
        }
        #endregion

        #region UpdateFileNameDisplayEvent
        public static event UpdateFileNameDisplayHandler UpdateFileNameDisplayEvent;
        public delegate void UpdateFileNameDisplayHandler(string text);

        private static void UpdateFileNameDisplay(string text)
        {
            UpdateFileNameDisplayEvent(text);
        }
        #endregion

        #region Private Member

        private static void ResetProblemFiles()
        {
            problemFilesList.Clear();
        }

        #endregion

        #region LogEvent
        public static event AddDataErrorHandler AddDataErrorEvent;
        public delegate void AddDataErrorHandler(string ex, bool log);
        public static event AddExceptionHandler AddExceptionEvent;
        public delegate void AddExceptionHandler(string ex, bool log);
        public static event AddLogWithPreHandler AddLogWithPreEvent;
        public delegate void AddLogWithPreHandler(string pre, string ex, bool sw);
        public static event AddLogHandler AddLogEvent;
        public delegate void AddLogHandler(string ex, bool sw);

        public static void AddLog(string pre, string ex, bool sw)
        {
            AddLogWithPreEvent(pre, ex, sw);
        }

        public static void AddLog(string ex, bool sw)
        {
            AddLogEvent(ex, sw);
        }

        public static void AddDataError(string ex, bool log)
        {
            AddDataErrorEvent(ex, log);
        }

        public static void AddException(string ex, bool log)
        {
            AddExceptionEvent(ex, log);
        }

        #endregion
    }
}
