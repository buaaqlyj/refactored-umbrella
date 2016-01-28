using System;
using System.Collections.Generic;
using System.Windows.Forms;
using System.Linq;
using System.Text;

namespace Statistics
{
    public static class FormOperator
    {
        public static bool MessageBox_Show_YesNo(string text, string title)
        {
            return MessageBox.Show(text, title, MessageBoxButtons.YesNo) == DialogResult.Yes;
        }
    }
}
