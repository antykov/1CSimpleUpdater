using Microsoft.Win32;
using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace _1CSimpleUpdater
{
    static class Common
    {
        public static long GetVersionAsLong(string version, char separator = '.')
        {
            string[] split = version.Split(separator);
            long lVersion = 0, lVersionPart = 0;
            for (int i = 0; i < split.Length; i++)
                if (Int64.TryParse(split[i], out lVersionPart))
                    lVersion += lVersionPart * (long)Math.Pow(10, 4 * (split.Length - i - 1));

            return lVersion;
        }

        public static string RemovePathInvalidChars(string path, string replaceString = "")
        {
            string result = path;

            char[] invalidChars = Path.GetInvalidPathChars();
            foreach (var c in invalidChars)
            {
                result = result.Replace(c.ToString(), replaceString);
            }

            return result;
        }

        public static void LogException(Exception E, string info = "")
        {
            Console.ForegroundColor = ConsoleColor.Red;
            if (info.Trim().Length == 0)
                Console.WriteLine(E.Message);
            else
                Console.WriteLine($"{info}: {E.Message}");
            Exception inner = E.InnerException;
            while (inner != null)
            {
                Console.WriteLine($"    --> {inner.Message}");
                inner = inner.InnerException;
            }
        }

        public static void Log(string message, ConsoleColor color = ConsoleColor.White, bool newLine = true, int emptyLineLength = 0)
        {
            Console.ForegroundColor = color;
            if (emptyLineLength > 0)
                Console.Write($"\r{new String(' ', emptyLineLength)}");
            if (newLine)
                Console.WriteLine(message);
            else
                Console.Write($"\r{message}");
            if (emptyLineLength > 0)
                Console.Write($"\r");
        }
    }
}
