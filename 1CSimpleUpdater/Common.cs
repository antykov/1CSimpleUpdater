using Microsoft.Win32;
using System;
using System.Collections.Generic;
using System.Diagnostics;
using System.IO;
using System.Linq;
using System.Management;
using System.Text;
using System.Threading.Tasks;

namespace _1CSimpleUpdater
{
    static class Common
    {

        public static long GetVersionAsLong(string version, char separator = '.')
        {
            string[] split = version.Split(separator);
            long lVersion = 0;
            for (int i = 0; i < split.Length; i++)
            {
                long lVersionPart;
                long.TryParse(split[i], out lVersionPart);
                lVersion += lVersionPart * (long)Math.Pow(10, 4 * (split.Length - i - 1));
            }

            return lVersion;
        }

        public static bool CompareMajorMinorVersions(string version1, string version2)
        {
            return String.Join("", version1.Split('.').Take(2).ToArray<string>()) == String.Join("", version2.Split('.').Take(2).ToArray<string>());
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

        public static int StartProcessWithArguments(string fileName, string arguments, bool isUseCMD = true)
        {
            try
            {
                Process process = new Process();

                string cmdFilePath = "";
                if (isUseCMD)
                {
                    cmdFilePath = Path.ChangeExtension(System.Reflection.Assembly.GetEntryAssembly().Location, "cmd");
                    File.WriteAllText(cmdFilePath, $"start \"\" /wait \"{fileName}\" {arguments}", Encoding.GetEncoding("cp866"));
                    process.StartInfo.FileName = cmdFilePath;
                } else
                {
                    process.StartInfo.FileName = fileName;
                    process.StartInfo.Arguments = arguments;
                }

                process.StartInfo.UseShellExecute = false;
                process.StartInfo.RedirectStandardOutput = true;
                process.StartInfo.RedirectStandardError = true;
                process.OutputDataReceived += (s, e) => Common.Log(e.Data);
                process.ErrorDataReceived += (s, e) => Common.Log(e.Data, ConsoleColor.Red);
                process.Start();
                process.WaitForExit();

                if (isUseCMD)
                    File.Delete(cmdFilePath);

                return process.ExitCode;
            }
            catch (Exception e)
            {
                Common.LogException(e, $"Запуск {fileName} {arguments}");
                return -1;
            }
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
