using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;

namespace Statistics.Criterion
{
    public abstract class DoseCriterion : Criterion
    {
        protected DoseCriterion(int index, int column, string value, string voltage)
            :base(index, column, value, voltage, "长期稳定性超差")
        {
 
        }

        public string HalfValueLayer
        {
            get
            {
                return base.Value;
            }
        }
    }
}
