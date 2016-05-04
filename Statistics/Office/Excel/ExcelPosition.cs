using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;

namespace Statistics.Office.Excel
{
    public class ExcelPosition
    {
        private int _row = 0;
        private int _col = 0;
        private string _pos = "";
        private bool _valid = false;

        public ExcelPosition(int row, int col)
        {
            _row = row;
            _col = col;
            _valid = ChangeNumberToString(row, col, out _pos);
        }

        public ExcelPosition(string position)
        {
            _pos = position;
            _valid = ChangeStringToNumber(position, out _row, out _col);
        }

        #region Class Member
        public static bool ChangeNumberToString(int rowIndex, int columnIndex, out string position)
        {
            if (rowIndex > 0 && columnIndex > 0)
            {
                int dividedValue = columnIndex;
                int remainder = 0;
                string pos = "";
                while (dividedValue > 26)
                {
                    remainder = dividedValue % 26;
                    dividedValue /= 26;
                    pos = ((char)(remainder / 26 + 64)).ToString() + pos;
                }
                position = ((char)(dividedValue + 64)).ToString() + pos + rowIndex.ToString();
                return true;
            }
            position = "A1";
            throw new Exception("Invalid arguments! Row = " + rowIndex + ", Col = " + columnIndex + ".", null);
        }

        public static bool ChangeStringToNumber(string positionString, out int row, out int col)
        {
            char[] charArr = positionString.Trim().ToUpper().ToCharArray();
            if (charArr.Length > 0)
            {
                int stage = 0;
                int index = 0;
                row = 1;
                col = 1;
                for (int i = 0; i < charArr.Length; i++)
                {
                    switch (stage)
                    {
                        case 0:
                            //第一次跳入
                            if (DataUtility.DataUtility.IsCharUpperLetter(charArr[i], out index))
                            {
                                stage = 1;
                                col = index;
                            }
                            else
                            {
                                stage = 3;
                            }
                            break;
                        case 1:
                            //字母阶段
                            if (DataUtility.DataUtility.IsCharNumber(charArr[i], out index))
                            {
                                stage = 2;
                                row = index;
                            }
                            else if (DataUtility.DataUtility.IsCharUpperLetter(charArr[i], out index))
                            {
                                col = col * 26 + index;
                            }
                            break;
                        case 2:
                            //数字阶段
                            if (DataUtility.DataUtility.IsCharNumber(charArr[i], out index))
                            {
                                row = row * 10 + index;
                            }
                            else
                            {
                                stage = 3;
                            }
                            break;
                    }
                    if (stage > 2)
                    {
                        throw new ArgumentException("Excel位置不能识别: " + positionString, "positionString");
                    }
                }
                return true;
            }
            else
            {
                row = 1;
                col = 1;

                throw new ArgumentException("Excel位置为空，不能识别: " + positionString, "positionString");
            }
        }
        #endregion

        #region Property
        public int RowIndex
        {
            get
            {
                return _row;
            }
        }

        public int ColumnIndex
        {
            get
            {
                return _col;
            }
        }

        public string PositionString
        {
            get
            {
                return _pos;
            }
        }

        public bool IsValid
        {
            get
            {
                return _valid;
            }
        }
        #endregion

        #region Public Member
        public ExcelPosition GoDown()
        {
            return new ExcelPosition(_row + 1, _col);
        }
        #endregion
    }
}
