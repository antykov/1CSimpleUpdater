using Microsoft.Win32;
using System;
using System.Collections.Generic;
using System.Diagnostics;
using System.IO;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace _1CSimpleUpdater
{
    [Serializable]
    public class InstalledPlatformInfo
    {
        public string Description;
        public string PlatformVersion;
        public string InstallLocation;
        public string ApplicationPath;
        public string ComConnectorVersion;
        public string ComConnectorRegistryPath;
        public Type ComConnector;
    }

    public static class Platform1C
    {
        public static readonly string SessionLockCode = "1CSimpleUpdater";
        public static readonly string PlatformOutFilePath = Path.Combine(Environment.CurrentDirectory, "1cv8_out.log");
        public static SortedList<long, InstalledPlatformInfo> InstalledPlatforms;

        static Platform1C()
        {
            InstalledPlatforms = new SortedList<long, InstalledPlatformInfo>();

            using (RegistryKey key = Registry.LocalMachine.OpenSubKey(@"SOFTWARE\Microsoft\Windows\CurrentVersion\Uninstall"))
            {
                var query = key.GetSubKeyNames().Select(s => key.OpenSubKey(s)).Where(w =>
                {
                    if (w.GetValueNames().Contains("DisplayName"))
                    {
                        return (w.GetValue("DisplayName").ToString().IndexOf("1C") != -1);
                    }

                    else
                        return (false);
                });

                foreach (var item in query)
                {
                    string installLocation = Path.Combine(item.GetValue("InstallLocation").ToString(), "bin");
                    string appPath = Path.Combine(installLocation, "1cv8.exe");
                    if (!File.Exists(appPath))
                        continue;

                    string displayVersion = item.GetValue("DisplayVersion").ToString();

                    InstalledPlatforms.Add(Common.GetVersionAsLong(displayVersion), new InstalledPlatformInfo
                    {
                        Description = item.GetValue("DisplayName").ToString(),
                        PlatformVersion = displayVersion,
                        InstallLocation = installLocation,
                        ApplicationPath = appPath,
                        ComConnectorVersion = Platform1C.GetComConnectorVersion(displayVersion),
                        ComConnectorRegistryPath = GetComConnectorInprocServerRegistryPath(Platform1C.GetComConnectorVersion(displayVersion)),
                        ComConnector = Type.GetTypeFromProgID(Platform1C.GetComConnectorVersion(displayVersion))
                    });
                }
            }

        }

        public static string GetComConnectorVersion(string version)
        {
            int outN;
            int[] versionParts = version.Split('.').Select(s => { if (int.TryParse(s, out outN)) return (outN); else return (-1); }).Where(w => w != -1).ToArray();

            if (versionParts.Length == 4)
                return $"V{versionParts[0]}{versionParts[1]}.COMConnector";

            return "";
        }

        public static string GetComConnectorInprocServerRegistryPath(string connectorVersion)
        {
            if (String.IsNullOrWhiteSpace(connectorVersion))
                return "";

            using (RegistryKey connectorCLSID = Registry.ClassesRoot.OpenSubKey(connectorVersion + "\\CLSID"))
            {
                if (connectorCLSID == null)
                    return "";

                string result = "" + (Environment.Is64BitOperatingSystem ? "Wow6432Node" : "") + "\\CLSID\\" + connectorCLSID.GetValue("") + "\\InprocServer32";
                using (RegistryKey connectorInprocServer = Registry.ClassesRoot.OpenSubKey(result))
                {
                    if (connectorInprocServer == null)
                        return "";

                    return result;
                }
            }
        }

        public static void CheckComConnectorInprocServerVersion(InstalledPlatformInfo platformInfo, Base1CSettings baseSettings)
        {
            string comcntrPath = Path.Combine(platformInfo.InstallLocation, "comcntr.dll");
            if (!File.Exists(comcntrPath) || String.IsNullOrWhiteSpace(platformInfo.ComConnectorRegistryPath))
                return;

            try
            {
                using (RegistryKey key = Registry.ClassesRoot.OpenSubKey(platformInfo.ComConnectorRegistryPath, true))
                {
                    if (key.GetValue("").ToString().ToUpper() != comcntrPath.ToUpper())
                    {
                        key.SetValue("", comcntrPath);
                    }
                }
            }
            catch (Exception e)
            {
                Common.LogException(e);
            }
        }

        public static InstalledPlatformInfo GetAppropriatePlatformInfo(string neededVersion)
        {
            if (InstalledPlatforms.Count == 0)
                return null;

            if (neededVersion.Length == 0)
                return InstalledPlatforms.ElementAt(InstalledPlatforms.Count - 1).Value;

            int outN;
            int[] versionParts = neededVersion.Split('.').Select(s => { if (int.TryParse(s, out outN)) return (outN); else return (-1); }).Where(w => w != -1).ToArray();
            if (versionParts.Length == 2)
            {
                long lowVersion = Common.GetVersionAsLong($"{versionParts[0]}.{versionParts[1]}.0.0");
                long highVersion = Common.GetVersionAsLong($"{versionParts[0]}.{versionParts[1] + 1}.0.0");
                try
                {
                    var maxVersion = InstalledPlatforms.Select(s => s).Where(w => w.Key >= lowVersion && w.Key < highVersion).Max(l => l.Key);
                    return InstalledPlatforms[maxVersion];
                } catch { }
            } else if (versionParts.Length == 4)
            {
                long versionL = Common.GetVersionAsLong(neededVersion);
                if (InstalledPlatforms.ContainsKey(versionL))
                    return InstalledPlatforms[versionL];
            }

            return null;
        }

        public static string GetArgumentForBaseConnection(Base1CSettings baseSettings, string mode)
        {
            StringBuilder args = new StringBuilder();
            args.Append($" {mode} /IBConnectionString \"{baseSettings.IBConnectionString.Replace("\"", "\"\"")}\"");
            if (!String.IsNullOrWhiteSpace(baseSettings.Login))
                args.Append($" /N \"{baseSettings.Login}\"");
            if (!String.IsNullOrWhiteSpace(baseSettings.Password))
                args.Append($" /P \"{baseSettings.Password}\"");
            args.Append(" /DisableStartupMessages /DisableStartupDialogs");
            args.Append($" /UC {Platform1C.SessionLockCode}");
            args.Append($" /Out \"{Platform1C.PlatformOutFilePath}\"");

            return args.ToString();
        }

        public static string GetArgumentForBaseBackup(Base1CSettings baseSettings, string backupFileName)
        {
            StringBuilder args = new StringBuilder();
            args.Append(GetArgumentForBaseConnection(baseSettings, "DESIGNER"));
            args.Append($" /DumpIB \"{backupFileName}\"");

            return args.ToString();
        }

        public static string GetArgumentForBaseConfUpdate(Base1CSettings baseSettings, string cfuFileName)
        {
            StringBuilder args = new StringBuilder();
            args.Append(GetArgumentForBaseConnection(baseSettings, "DESIGNER"));
            args.Append($" /UpdateCfg \"{cfuFileName}\"");
            args.Append($" /UpdateDBCfg -Dynamic-");

            return args.ToString();
        }
    }
}
