using Microsoft.Win32;
using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace _1CSimpleUpdater
{
    public class InstalledPlatformInfo
    {
        public string Description;
        public string PlatformVersion;
        public string ApplicationPath;
        public string ComConnectorVersion;
        public string ComConnectorRegistryPath;
        public Type ComConnector;
    }

    public static class Platform1C
    {
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
                    if (!File.Exists(Path.Combine(installLocation, "1cv8.exe")))
                        continue;

                    string displayVersion = item.GetValue("DisplayVersion").ToString();

                    InstalledPlatforms.Add(Common.GetVersionAsLong(displayVersion), new InstalledPlatformInfo
                    {
                        Description = item.GetValue("DisplayName").ToString(),
                        PlatformVersion = displayVersion,
                        ApplicationPath = installLocation,
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
            string comcntrPath = Path.Combine(platformInfo.ApplicationPath, "comcntr.dll");
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

    }
}
