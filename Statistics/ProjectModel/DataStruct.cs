using System;
using System.Collections;
using System.Collections.Generic;
using System.Linq;
using System.Text;

namespace Statistics.ProjectModel
{
    public class DataStruct
    {
        private DataRange dr;
        private double da;
        private Distance dt;
        private double dd;

        public DataStruct(DataRange dataRange, double data, double standardData)
        {
            DataRanges = dataRange;
            Data = data;
            dd = standardData;
            if (dd < 0.0035 && dd > 0.0001)
            {
                dt = Distance.d1;
            }
            else if (dd > 0.0035 && dd < 0.008)
            {
                dt = Distance.d1_5;
            }
            else
            {
                dt = Distance.Unknown;
            }
        }

        public DataStruct(double data, double standardData, string range, int pattern)
        {
            da = data;
            dd = standardData;
            if (dd < 0.0035 && dd > 0.0001)
            {
                dt = Distance.d1_5;
            }
            else if (dd > 0.0034 && dd < 0.008)
            {
                dt = Distance.d1;
            }
            else
            {
                dt = Distance.Unknown;
            }
            switch (pattern)
            {
                case 0:
                    //Dose
                    if (da > 1000)
                    {
                        dr = DataRange.uGy;
                    }
                    else if (da > 100)
                    {
                        dr = DataRange.mR;
                    }
                    else if (da > 1.1 && da < 10)
                    {
                        if (range.Trim().ToLower().EndsWith("mgy"))
                        {
                            dr = DataRange.mGy;
                        }
                        else if (range.Trim().ToLower().EndsWith("cgycm"))
                        {
                            dr = DataRange.cGycm;
                        }
                        else if (range.Trim().ToLower().EndsWith("rcm"))
                        {
                            dr = DataRange.Rcm;
                        }
                    }
                    else if (da > 0.1 && da < 1.2)
                    {
                        if (range.Trim().ToLower().EndsWith("r"))
                        {
                            dr = DataRange.R;
                        }
                        else if (range.Trim().ToLower().EndsWith("cgy"))
                        {
                            dr = DataRange.cGy;
                        }
                    }
                    else if (da > 0.01 && da < 0.1)
                    {
                        dr = DataRange.Gycm;
                    }
                    else
                    {
                        dr = DataRange.Unknown;
                    }
                    break;
                case 1:
                    //CT
                    if (da > 1000)
                    {
                        dr = DataRange.uGy;
                    }
                    else if (da > 100)
                    {
                        dr = DataRange.mR;
                    }
                    else if (da > 1 && da < 10)
                    {
                        dr = DataRange.mGy;
                    }
                    else if (da > 10 && da < 90)
                    {
                        dr = DataRange.mGycm;
                    }
                    else if (range.Trim().ToLower().EndsWith("r"))
                    {
                        dr = DataRange.R;
                    }
                    else if (range.Trim().ToLower().EndsWith("cgy"))
                    {
                        dr = DataRange.cGy;
                    }
                    else
                    {
                        dr = DataRange.Unknown;
                    }
                    break;
                case 2:
                    //TODO: kv的单位判断
                    break;
            }

        }

        public static DataStruct CalDataRange(List<DataStruct> dataStructList, string rangeText, int pattern, out bool success)
        {
            int count1 = 0, count2 = 0;
            double daav = 0, diav = 0;
            foreach (DataStruct item in dataStructList)
            {
                if (item != null)
                {
                    if (item.DataRanges != DataRange.Unknown)
                    {
                        count1++;
                        daav += (double)item.Data;
                    }
                    if (item.Distance != Distance.Unknown)
                    {
                        count2++;
                        diav += (double)item.DistanceData;
                    }
                }
            }
            if (count1 > 0)
            {
                daav /= count1;
            }
            if (count2 > 0)
            {
                diav /= count2;
            }
            success = true;
            this(daav, diav, rangeText, pattern);
            return new DataStruct(daav, diav, rangeText, pattern);
        }

        public DataRange DataRanges
        {
            get
            {
                return dr;
            }
            set
            {
                dr = value;
            }
        }

        public double Data
        {
            get
            {
                return da;
            }
            set
            {
                da = value;
            }
        }

        public Distance Distance
        {
            get
            {
                return dt;
            }
            set
            {
                dt = value;
            }
        }

        public double DistanceData
        {
            set
            {
                dd = value;
            }
            get
            {
                return dd;
            }
        }
    }

    /// <summary>
    /// Unknown
    /// R     0.1 < V < 0.8
    /// cGy   0.8 < V < 1.2
    /// mGy   1.1 < V < 10.0
    /// Rcm   1.1 < V < 10.0 R/Rcm
    /// cGycm 1.1 < V < 10.0
    /// mGycm 10  < V < 90
    /// mR    100 < V < 1000
    /// uGy   1000< V < 10000
    /// Gycm  0.01 < V < 0.1
    /// </summary>
    public enum DataRange { Unknown = 0, cGy = 1, R = 2, mGy = 3, mR = 4, uGy = 5, mGycm = 6, Rcm = 7, cGycm = 8, Gycm = 10 };

    public enum Distance { Unknown = 0, d1 = 1, d1_5 = 2 };
}
