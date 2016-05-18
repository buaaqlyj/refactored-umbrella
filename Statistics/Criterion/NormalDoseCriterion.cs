using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;

namespace Statistics.Criterion
{
    public class NormalDoseCriterion : DoseCriterion
    {
        public static NormalDoseCriterion Null = new NormalDoseCriterion(0, 0, @"0.00", @"0kV");
        public static NormalDoseCriterion RQR2_40 = new NormalDoseCriterion(1, 4, @"1.42", @"40kV");
        public static NormalDoseCriterion RQR3_50 = new NormalDoseCriterion(2, 6, @"1.78", @"50kV");
        public static NormalDoseCriterion RQR4_60 = new NormalDoseCriterion(3, 8, @"2.19", @"60kV");
        public static NormalDoseCriterion RQR5_70 = new NormalDoseCriterion(4, 10, @"2.58", @"70kV");
        public static NormalDoseCriterion RQR6_80 = new NormalDoseCriterion(5, 12, @"3.01", @"80kV");
        public static NormalDoseCriterion RQR7_90 = new NormalDoseCriterion(6, 14, @"3.48", @"90kV");
        public static NormalDoseCriterion RQR8_100 = new NormalDoseCriterion(7, 16, @"3.97", @"100kV");
        public static NormalDoseCriterion RQR9_120 = new NormalDoseCriterion(8, 18, @"5.00", @"120kV");
        public static NormalDoseCriterion RQR_140 = new NormalDoseCriterion(9, 20, @"/", @"140kV");
        public static NormalDoseCriterion RQR10_150 = new NormalDoseCriterion(10, 22, @"6.57", @"150kV");

        private NormalDoseCriterion(int index, int column, string halfLayer, string voltage)
            : base(index, column, halfLayer, voltage)
        {
            
        }
    }
}
