using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Reflection;
using System.Xml;
using System.Text;

using log4net;
using log4net.Config;

namespace Statistics.Log.LogWriter
{
    public class LogForNet: ILogWriter
    {
        private static LogForNet log4Net = new LogForNet();
        
        private LogForNet()
        {
            record(log4net.LogManager.GetLogger(MethodInfo.GetCurrentMethod().DeclaringType), "", LogLevel.Fatal);
        }

        #region ILogWriter 成员

        public ILogWriter getInstance()
        {
            return (ILogWriter)log4Net;
        }

        public void initial()
        {
            var logCfg = new FileInfo(AppDomain.CurrentDomain.BaseDirectory + "log4net.config");
            XmlConfigurator.ConfigureAndWatch(logCfg);
        }

        public void record(ILog logger, string message, LogLevel logLevel)
        {
            switch (logLevel)
            {
                case LogLevel.Debug:
                    logger.Debug(message, new Exception(message));
                    break;
            }
        }

        #endregion




        public void record(string message, LogLevel logLevel)
        {
            throw new NotImplementedException();
        }

        public void record(Exception exception, LogLevel logLevel)
        {
            throw new NotImplementedException();
        }
    }
}
