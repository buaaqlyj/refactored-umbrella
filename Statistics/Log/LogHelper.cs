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
    }
}
