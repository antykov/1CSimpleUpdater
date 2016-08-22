using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.IO;
using System.Diagnostics;

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

                if (Platform1C.InstalledPlatforms.Count == 0)
                    throw new Exception("Не найдено ни одной установленной платформы 1С!");

                if (ConfUpdate1C.ConfUpdates.Count == 0)
                    throw new Exception("Не найдено ни одного обновления конфигураций!");
            }
            catch (Exception E)
            {
                Common.LogException(E);
                Environment.Exit(1);
                return;
            }

            foreach (var base1C in AppSettings.settings.Bases1C)
                Base1C.UpdateBase(base1C);

            Common.Log("\nPress any key...");
            Console.ReadKey();
        }
    }
}
