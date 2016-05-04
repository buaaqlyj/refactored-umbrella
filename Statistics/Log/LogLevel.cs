using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;

namespace Statistics.Log
{
    public enum LogLevel
    {
        Debug = -1, //仅供查错时需要的信息：流程
        Info = 0, //交互信息：进度
        Warn = 1, //有错但不要紧：警告信息
        Error = 2, //出错信息：终止当前Job
        Fatal = 3 //严重出错信息：终止当前Task
    }
}
