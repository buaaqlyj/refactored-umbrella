using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;

namespace Statistics.Criterion.KV
{
    public class KVCriterion : Criterion
    {
        private double _th1 = 0.0;
        private double _th2 = 0.0;

        public static KVCriterion Null = new KVCriterion(0, 0, @"0.00", @"0kV", "空规范", 0, 0);
        public static KVCriterion RQR2_40 = new KVCriterion(1, 4, @"39.83", @"40kV", "相对固有误差过大", 1.0, 1.0);
        public static KVCriterion RQR3_50 = new KVCriterion(2, 7, @"49.48", @"50kV", "相对固有误差过大", 1.0, 1.0);
        public static KVCriterion RQR4_60 = new KVCriterion(3, 10, @"59.29", @"60kV", "固有误差过大", 0.02, 0.02);
        public static KVCriterion RQR5_70 = new KVCriterion(4, 13, @"69.15", @"70kV", "固有误差过大", 0.02, 0.02);
        public static KVCriterion RQR6_80 = new KVCriterion(5, 16, @"78.99", @"80kV", "固有误差过大", 0.02, 0.02);
        public static KVCriterion RQR7_90 = new KVCriterion(6, 19, @"89.02", @"90kV", "固有误差过大", 0.02, 0.02);
        public static KVCriterion RQR8_100 = new KVCriterion(7, 22, @"99.15", @"100kV", "固有误差过大", 0.02, 0.02);
        public static KVCriterion RQR9_120 = new KVCriterion(8, 25, @"118.84", @"120kV", "固有误差过大", 0.02, 0.02);
        public static KVCriterion RQR_140 = new KVCriterion(9, 28, @"138.90", @"140kV", "固有误差过大", 0.02, 0.02);
        public static KVCriterion RQR10_150 = new KVCriterion(10, 31, @"148.62", @"150kV", "固有误差过大", 0.02, 0.02);

        private KVCriterion(int index, int column, string ppv, string voltage, string testingItem, double th1, double th2)
            : base(index, column, ppv, voltage, testingItem)
        {
            _th1 = th1;
            _th2 = th2;
        }

        public string PPV
        {
            get
            {
                return base.Value;
            }
        }

        public double Threshold1
        {
            get
            {
                return _th1;
            }
        }

        public double Threshold2
        {
            get
            {
                return _th2;
            }
        }
    }
}
