using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace _1CSimpleUpdater
{
    class Program
    {
        static void Main(string[] args)
        {
            Console.OutputEncoding = Encoding.Unicode;

            try
            {
                AppSettings.LoadSettings();
                AppSettings.CheckSettings();
            }
            catch (Exception E)
            {
                Common.LogException(E);
                Environment.Exit(1);
                return;
            }
        }
    }
}
