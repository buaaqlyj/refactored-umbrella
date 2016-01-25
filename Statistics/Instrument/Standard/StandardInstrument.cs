using System;
using System.Collections.Generic;
using System.Globalization;
using System.Linq;
using System.Text;

namespace Statistics.Instrument.Standard
{
    public class StandardInstrument
    {
        private string _name;
        private DateTime _dateTime;
        private ValidState _valid;

        private static CultureInfo provider = new CultureInfo("zh-Hans");
        private static DateTime Today = DateTime.Today;
        private static DateTime TwoWeeksLater = DateTime.Today.AddDays(14);
        
        public StandardInstrument(string name, string date)
        {
            _name = name;
            _dateTime = DateTime.ParseExact(date, "yyyy-MM-dd", provider);
            if (_dateTime.CompareTo(TwoWeeksLater) > 0)
            {
                _valid = ValidState.OK;
            }
            else if (_dateTime.CompareTo(Today) > 0)
            {
                _valid = ValidState.WillExpireSoon;
            }
            else
            {
                _valid = ValidState.Expired;
            }
        }

        public string Name
        {
            get
            {
                return _name;
            }
        }

        public string Date
        {
            get
            {
                return _dateTime.ToString("yyyy-MM-dd");
            }
        }

        public string State
        {
            get
            {
                return _valid.State;
            }
        }
    }
}
