using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace _1CSimpleUpdater
{
    public class ConfUpdateInfo
    {
        public string Name;
        public string Version;
        public string FromVersion;
        public string[] FromVersions;
        public string UpdateFilePath;
    }

    public static class ConfUpdate1C
    {
        public static Dictionary<string, ConfUpdateInfo> ConfUpdates;

        static ConfUpdate1C()
        {
            ConfUpdates = new Dictionary<string, ConfUpdateInfo>();


        }
    }
}
