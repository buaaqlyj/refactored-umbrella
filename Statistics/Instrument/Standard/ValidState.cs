using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;

namespace Statistics.Instrument.Standard
{

    public class ValidState
    {
        private int _index;
        private string _message;

        public static ValidState OK = new ValidState(0, "正常");
        public static ValidState WillExpireSoon = new ValidState(1, "即将过期");
        public static ValidState Expired = new ValidState(2, "过期");

        private ValidState(int index, string message)
        {
            _index = index;
            _message = message;
        }

        public string State
        {
            get
            {
                return _message;
            }
        }

        public int Index
        {
            get
            {
                return _index;
            }
        }
    }
}
