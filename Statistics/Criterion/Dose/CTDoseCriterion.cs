using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;

namespace Statistics.Criterion.Dose
{
    public class CTDoseCriterion : DoseCriterion
    {
        public static CTDoseCriterion Null = new CTDoseCriterion(0, 0, @"0.00", @"0kV");
        public static CTDoseCriterion RQR2_40 = new CTDoseCriterion(1, 4, @"1.42", @"40kV");
        public static CTDoseCriterion RQR3_50 = new CTDoseCriterion(2, 6, @"1.78", @"50kV");
        public static CTDoseCriterion RQR4_60 = new CTDoseCriterion(3, 8, @"2.19", @"60kV");
        public static CTDoseCriterion RQR5_70 = new CTDoseCriterion(4, 10, @"2.58", @"70kV");
        public static CTDoseCriterion RQR6_80 = new CTDoseCriterion(5, 12, @"3.01", @"80kV");
        public static CTDoseCriterion RQR7_90 = new CTDoseCriterion(6, 14, @"3.48", @"90kV");
        public static CTDoseCriterion RQR8_100 = new CTDoseCriterion(7, 16, @"3.97", @"100kV");
        public static CTDoseCriterion RQR9_120 = new CTDoseCriterion(8, 18, @"5.00", @"120kV");
        public static CTDoseCriterion RQR_140 = new CTDoseCriterion(9, 20, @"/", @"140kV");
        public static CTDoseCriterion RQR10_150 = new CTDoseCriterion(10, 22, @"6.57", @"150kV");

        private CTDoseCriterion(int index, int column, string halfLayer, string voltage)
            : base(index, column, halfLayer, voltage)
        {

        }
    }
}
