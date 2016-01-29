using System;
using System.Collections.Generic;
using System.Windows.Forms;

namespace Statistics
{
    static class Program
    {
        /// <summary>
        /// 应用程序的主入口点。
        /// </summary>
        [STAThread]
        static void Main()
        {
            Application.EnableVisualStyles();
            Application.SetCompatibleTextRenderingDefault(false);
            //Application.Run(new MainForm());
            //myDog.RunDogTesting();
            bool checkSuperDog = true;
            SuperDog.SuperDogSeries myDog = new SuperDog.SuperDogSeries();
            if (!checkSuperDog || myDog.DogFlag)
            {
                Application.Run(new MainForm());
            }
            else
                MessageBox.Show("Error:     " + myDog.Status.ToString() + "\nSuperDog disabled!");
        }
    }
}
