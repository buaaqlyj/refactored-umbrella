using System;
using System.Collections.Generic;
using System.ComponentModel;

namespace Statistics.ProjectModel
{
    public enum FixType
    {
        [Description("不修正")]
        DontFix = 0,
        [Description("修正")]
        NeedFix = 1,
        [Description("自修正")]
        SelfFix = 2
    }
}
