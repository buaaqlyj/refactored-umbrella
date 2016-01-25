using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Text;

namespace Statistics.Instrument.Tested
{

    public class TestedInstrument
    {
        private string _instrumentTypeName = "";
        private static string[] _existTypeCT = null;
        private static string[] _existTypeDose = null;
        private static string[] _existTypeKV = null;
        private static List<string> _existTypes = null;

        public TestedInstrument(string name)
        {
            _instrumentTypeName = name;
        }

        public static void InitialTypes(string directoryName)
        {
            _existTypeCT = ReadInTypesFromFile(directoryName + @"\CT仪器.txt");
            _existTypeKV = ReadInTypesFromFile(directoryName + @"\KV仪器.txt");
            _existTypeDose = ReadInTypesFromFile(directoryName + @"\Dose仪器.txt");
            if (_existTypes == null)
            {
                _existTypes = new List<string>();
            }
            foreach (string item in _existTypeCT)
            {
                if (!_existTypes.Contains(item))
                {
                    _existTypes.Add(item);
                }
            }
            foreach (string item in _existTypeKV)
            {
                if (!_existTypes.Contains(item))
                {
                    _existTypes.Add(item);
                }
            }
            foreach (string item in _existTypeDose)
            {
                if (!_existTypes.Contains(item))
                {
                    _existTypes.Add(item);
                }
            }
        }

        public static bool IsEqualTo(string strType1, string strType2)
        {
            if (strType1.ToLower() == strType2.ToLower())
            {
                return true;
            }
            else if (strType1.ToLower().StartsWith("solidose") && strType2.ToLower().StartsWith("solidose"))
            {
                return true;
            }
            else if (strType1.ToLower().StartsWith("piranha") && strType2.ToLower().StartsWith("piranha"))
            {
                return true;
            }
            else if (strType1.ToLower().StartsWith("35050a") && strType2.ToLower().StartsWith("35050a"))
            {
                return true;
            }
            else if (strType1.ToLower().StartsWith("pmx") && strType2.ToLower().StartsWith("pmx"))
            {
                return true;
            }
            else
            {
                return false;
            }
        }

        private static string[] ReadInTypesFromFile(string filename)
        {
            if (File.Exists(filename))
            {
                string text = DataUtility.DataUtility.ReadInText(filename, "#", ",");
                text.Replace(@"，", ",");
                while (text.Contains(" ,"))
                {
                    text.Replace(" ,", ",");
                }
                while (text.Contains(", "))
                {
                    text.Replace(", ", ",");
                }
                return text.Split(new char[] { ',' }, StringSplitOptions.RemoveEmptyEntries);
            }
            else
            {
                //TODO: strict logging method
                //(@"没有找到Dose仪器类型文件，处理结果可能有误", @"提示", MessageBoxButtons.OK);
                return new string[] { "Barracuda", "Piranha", "Xi", "Solidose", "Diadose", "Cobia", "TNT12000", "4000M+", "Nero8000", "magicmax", "RS-2000", "T6580", "CONNy", "unidos", "9096" };
            }
        }

        public static string[] CTTypes
        {
            get
            {
                return _existTypeCT;
            }
        }

        public static string[] KVTypes
        {
            get
            {
                return _existTypeKV;
            }
        }

        public static string[] DoseTypes
        {
            get
            {
                return _existTypeDose;
            }
        }

        public static string[] AllTypes
        {
            get
            {
                return _existTypes.ToArray();
            }
        }
    }
}
