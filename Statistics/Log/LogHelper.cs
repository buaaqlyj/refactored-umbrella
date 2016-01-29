using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;

using Statistics.Log.LogWriter;

namespace Statistics.Log
{
    public static class LogHelper
    {
        private static ILogWriter _logWriter = null;

        public static void initial(ILogWriter logWriter)
        {
            _logWriter = logWriter;
        }

        #region Task_Related

        #endregion

        #region Job_Related

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
