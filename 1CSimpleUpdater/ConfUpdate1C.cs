using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace _1CSimpleUpdater
{
    public class ConfUpdateInfo
    {
        public string Name;
        public string Version;
        public string[] FromVersions;
        public string UpdateFilePath;
    }

    public static class ConfUpdate1C
    {
        public static Dictionary<string, SortedList<long, ConfUpdateInfo>> ConfUpdates;

        static ConfUpdate1C()
        {
            ConfUpdates = new Dictionary<string, SortedList<long, ConfUpdateInfo>>();

            FindConfUpdates(new DirectoryInfo(AppSettings.settings.TemplatesDirectory));
        }

        public static void FindConfUpdates(DirectoryInfo dirInfo)
        {
            if (dirInfo.GetFiles().Where(w => w.Name.ToUpper() == "1CV8.MFT" || w.Name.ToUpper() == "UPDINFO.TXT" || w.Name.ToUpper() == "1CV8.CFU").Count() == 3)
            {
                string mftText = File.ReadAllText(Path.Combine(dirInfo.FullName, "1cv8.mft"));
                string updinfoText = File.ReadAllText(Path.Combine(dirInfo.FullName, "UpdInfo.txt"));

                ConfUpdateInfo confUpdateInfo = new ConfUpdateInfo();
                try
                {
                    confUpdateInfo.Name = 
                        mftText.Split('\n')
                            .Select(s => s.Split('='))
                            .Where(w => w.Length == 2 && w[0] == "Name")
                                .Select(s => s[1]).ToArray()[0].Trim();
                    confUpdateInfo.Version = 
                        updinfoText.Split('\n')
                            .Select(s => s.Split('='))
                            .Where(w => w.Length == 2 && w[0] == "Version")
                                .Select(s => s[1]).ToArray()[0].Trim();
                    confUpdateInfo.FromVersions =
                        updinfoText.Split('\n')
                            .Select(s => s.Split('='))
                            .Where(w => w.Length == 2 && w[0] == "FromVersions")
                                .Select(s => s[1].Split(';')).ToArray()[0]
                                    .Where(w => !String.IsNullOrWhiteSpace(w)).ToArray<string>();
                    confUpdateInfo.UpdateFilePath = Path.Combine(dirInfo.FullName, "1cv8.cfu");

                    if (!ConfUpdates.Keys.Contains(confUpdateInfo.Name))
                        ConfUpdates[confUpdateInfo.Name] = new SortedList<long, ConfUpdateInfo>();

                    ConfUpdates[confUpdateInfo.Name].Add(Common.GetVersionAsLong(confUpdateInfo.Version), confUpdateInfo);
                }
                catch { }
            }
            else
            {
                foreach (var subdirInfo in dirInfo.GetDirectories())
                {
                    FindConfUpdates(subdirInfo);
                }
            }
        }

        public static SortedList<long, ConfUpdateInfo> GetUpdatesSequence(Base1CInfo baseInfo)
        {
            if (!ConfUpdates.Keys.Contains(baseInfo.ConfName))
                throw new Exception($"Не удалось найти ни одного обновления для конфигурации {baseInfo.ConfName}!");

            SortedList<long, ConfUpdateInfo> result = new SortedList<long, ConfUpdateInfo>();
            ConfUpdateInfo confUpdate;

            var queryLastVersion =
                ConfUpdates[baseInfo.ConfName]
                    .Select(s => s.Value)
                    .Where(w => Common.CompareMajorMinorVersions(w.Version, baseInfo.ConfVersion));
            if (queryLastVersion.Count() == 0)
                return null;
            confUpdate = queryLastVersion.Last();
            if (confUpdate.Version == baseInfo.ConfVersion)
            {
                result.Add(Common.GetVersionAsLong(baseInfo.ConfVersion), confUpdate);
                return result;
            }

            var query = 
                ConfUpdates[baseInfo.ConfName]
                    .Select(s => s.Value)
                    .Where(w => w.FromVersions.Contains(baseInfo.ConfVersion) && Common.CompareMajorMinorVersions(w.Version, baseInfo.ConfVersion));
            if (query.Count() == 0)
                return null;

            confUpdate = query.ToArray()[0];
            result.Add(Common.GetVersionAsLong(confUpdate.Version), confUpdate);

            while (true)
            {
                query = 
                    ConfUpdates[baseInfo.ConfName]
                        .Select(s => s.Value)
                        .Where(w => w.FromVersions.Contains(confUpdate.Version) && Common.CompareMajorMinorVersions(w.Version, confUpdate.Version))
                        .OrderBy(o => Common.GetVersionAsLong(o.Version));
                if (query.Count() == 0)
                    break;

                confUpdate = query.Last<ConfUpdateInfo>();
                result.Add(Common.GetVersionAsLong(confUpdate.Version), confUpdate);
            }

            return result;
        }
    }
}
