using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;

namespace Statistics.Criterion
{
    /// <summary>
    /// 一、什么时候使用抽象类和接口？
    /// 1.如果预计要创建组建的多个版本，则创建抽象类；抽象类提供简单的方法控制组件版本
    /// 2.如果创建的功能将在大范围的全异对象间使用，则使用接口；如果要设计小而简练的功能块，则使用接口；
    /// 3.如果要设计大的功能单元，则使用抽象类。如果要在组建的所有实现间提供通用的已实现功能，则使用抽象类。
    /// 4.抽象类主要用于关系密切的对象；而接口适合为不相关的类提供通用功能。
    /// 
    /// 二、实例
    /// 1.飞机会飞，鸟会飞，他们都继承了同一个接口“飞”；但是F22属于飞机抽象类，鸽子属于鸟抽象类；
    /// 2.铁门木门都是门抽象类，门不可以被实例化，只能给出具体的木门或铁门（多态）；而且只能是单继承于门，不能是窗或其他的；一个门可以有锁（接口）也可以有门铃（多实现）。门抽象类定义了是什么，接口锁定义了你能做什么，而且一个接口最好只能做一件事，不能要求锁也能发挥门铃的作用（接口污染）。
    /// 
    /// 接口：
    /// 显式实现的接口必须由接口调用（可以隐藏具体类名）
    /// 隐式实现的接口可以使用接口调用和类名调用
    /// 
    /// Criterion
    /// 序号Index，列号Column，管电压Voltage，检测项TestingItem，取值Value
    /// </summary>
    
    public abstract class Criterion
    {
        //序号
        private int _index = 0;
        //列号
        private int _column = 0;
        //取值
        private string _value = "";
        //管电压
        private string _voltage = "";
        //检测项
        private string _testingItem = "";

        protected Criterion(int index, int column, string value, string voltage, string testingItem)
        {
            _index = index;
            _column = column;
            _value = value;
            _voltage = voltage;
            _testingItem = testingItem;
        }

        public int Index
        {
            get
            {
                return _index;
            }
        }

        public int Column
        {
            get
            {
                return _column;
            }
        }

        public string Voltage
        {
            get
            {
                return _voltage;
            }
        }

        public string TestingItem
        {
            get
            {
                return _testingItem;
            }
        }
        //对KV来说是PPV，对Dose来说是半值层，对CT来说是半值层
        public string Value 
        {
            get
            {
                return _value;
            }
        }
    }
}