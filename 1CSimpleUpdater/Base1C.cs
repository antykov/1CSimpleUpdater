using System;
using System.Collections.Generic;
using System.Diagnostics;
using System.IO;
using System.Linq;
using System.Runtime.InteropServices;
using System.Text;
using System.Threading;
using System.Threading.Tasks;

namespace _1CSimpleUpdater
{
    [Serializable]
    public class Base1CInfo
    {
        public string ConfName;
        public string ConfVersion;
        public int SessionsCount;
        public SortedList<long, ConfUpdateInfo> UpdatesSequence;
        public string BackupFilePath;
        public InstalledPlatformInfo PlatformInfo;
    }

    public static class Base1C
    {
        public static Base1CInfo GetBase1CInfo(Base1CSettings baseSettings)
        {
            try
            {
                InstalledPlatformInfo platformInfo = Platform1C.GetAppropriatePlatformInfo(baseSettings.PlatformVersion);
                if (platformInfo == null)
                    throw new Exception($"Не удалось найти подходящую версию платформы для конфигурации {baseSettings.Description}, требуемая версия = {baseSettings.PlatformVersion}!");
                if (platformInfo.ComConnector == null)
                    throw new Exception($"Для платформы {platformInfo.PlatformVersion} не найден ComConnector!");
                Platform1C.CheckComConnectorInprocServerVersion(platformInfo, baseSettings);

                Base1CInfo baseInfo = Common.CallFunctionAtComConnection1CDomain("GetBase1CInfo", baseSettings, platformInfo.ComConnectorVersion);
                if (baseInfo != null)
                    baseInfo.PlatformInfo = platformInfo;

                return baseInfo;
            } catch (Exception e)
            {
                Common.LogException(e);
                return null;
            }
        }

        public static bool CheckBaseUpdateNecessity(Base1CInfo baseInfo)
        {
            try
            {
                baseInfo.UpdatesSequence = ConfUpdate1C.GetUpdatesSequence(baseInfo);
                if (baseInfo.UpdatesSequence == null || baseInfo.UpdatesSequence.Count == 0)
                    throw new Exception("Не удалось найти подходящие обновления!");
                if (baseInfo.UpdatesSequence.Count == 1 && baseInfo.UpdatesSequence.ElementAt(0).Value.Version == baseInfo.ConfVersion)
                {
                    Common.Log("Версия конфигурации является актуальной!", ConsoleColor.Green);
                    return false;
                }

                return true;
            }
            catch (Exception e)
            {
                Common.LogException(e);
                return false;
            }
        }

        public static bool PrepareBaseForUpdate(Base1CSettings baseSettings, Base1CInfo baseInfo)
        {
            if (baseSettings.IsServerIB)
            {
                if (Common.CallFunctionAtComConnection1CDomain("CheckServerBaseLockPossibilityAndLock", baseSettings, baseInfo.PlatformInfo.ComConnectorVersion))
                    baseInfo.SessionsCount = 0;
            }

            if (baseInfo.SessionsCount != 0)
            {
                Common.Log($"Невозможно получить монопольный доступ к базе, завершите работу пользователей и повторите попытку!", ConsoleColor.Red);
                return false;
            }

            if (!baseSettings.IsServerIB)
            {
                try
                {
                    string[] strings = {
                        "{1",
                        DateTime.Now.ToString("yyyyMMddHHmmss"),
                        DateTime.Now.AddHours(6).ToString("yyyyMMddHHmmss"),
                        "\"Установлена блокировка соединений для обновления конфигурации!\"",
                        $"\"{Platform1C.SessionLockCode}\"",
                        "\"\"}"
                    };

                    using (var fs = new FileStream(Path.Combine(baseSettings.FileIBPath, "1Cv8.cdn"), FileMode.Create))
                    {
                        using (var writer = new StreamWriter(fs))
                        {
                            writer.Write(String.Join(",", strings));
                        }
                    }
                }
                catch (Exception e)
                {
                    Common.LogException(e, "Попытка установки блокировки соединений");
                }
            }

            return true;
        }

        public static void RestoreBaseAfterUpdate(Base1CSettings baseSettings, Base1CInfo baseInfo)
        {
            try
            {
                if (baseSettings.IsServerIB)
                {
                    Common.CallFunctionAtComConnection1CDomain("UnlockServerBase", baseSettings, baseInfo.PlatformInfo.ComConnectorVersion);
                }
                else
                {
                    File.Delete(Path.Combine(baseSettings.FileIBPath, "1Cv8.cdn"));
                }
            }
            catch (Exception e)
            {
                Common.LogException(e, "ОШИБКА ПРИ ПОДГОТОВКЕ БАЗЫ К РАБОТЕ (снятие блокировки соединений / включение регламентных заданий)");
            }
        }

        public static void MakeBaseBackup(Base1CSettings baseSettings, Base1CInfo baseInfo)
        {
            if (baseSettings.BackupsCount == 0)
                return;

            Common.Log("Создание резервной копии...");

            string descr = "{" + $"{Common.RemovePathInvalidChars(baseSettings.Description, "_").Replace(' ', '_')}" + "}";
            List<string> files = Directory.GetFiles(AppSettings.settings.BackupsDirectory, $"???????????????????_{descr}.dt").OrderBy(o => File.GetCreationTime(o)).ToList<string>();
            while (files.Count >= baseSettings.BackupsCount)
            {
                File.Delete(files[0]);
                files.RemoveAt(0);
            }

            baseInfo.BackupFilePath = Path.Combine(AppSettings.settings.BackupsDirectory, $"{DateTime.Now.ToString("yyyy-MM-dd-HH-mm-ss")}_{descr}.dt");
            Common.StartProcessWithArguments(baseInfo.PlatformInfo.ApplicationPath, Platform1C.GetArgumentForBaseBackup(baseSettings, baseInfo.BackupFilePath));

            if (!File.Exists(baseInfo.BackupFilePath))
                throw new Exception($"Возможно произошла ошибка при создании резервной копии. Файл {baseInfo.BackupFilePath} не обнаружен!");
        }

        public static bool CheckOutFileAfterConfUpdateOrRaiseException()
        {
            if (!File.Exists(Platform1C.PlatformOutFilePath))
                throw new Exception("Возможно произошла ошибка при обновлении, т.к. отсутствует файл с результатами работы конфигуратора 1С!");

            string outText = File.ReadAllText(Platform1C.PlatformOutFilePath, Encoding.GetEncoding("windows-1251"));
            if (String.IsNullOrWhiteSpace(outText))
                throw new Exception("Возможно произошла ошибка при обновлении, т.к. файл с результатами работы конфигуратора 1С пуст!");

            string[] splitOutText = outText.Split('\n');
            if (splitOutText.Where(w => w.IndexOf("Файл не содержит доступных обновлений") != -1).Count() > 0)
            {
                Common.Log("Произошла ошибка: файл не содержит доступных обновлений!\nБудет произведена попытка обновления на следующий релиз, т.к. возможно просто не обновлена конфигурация БД!", ConsoleColor.Red);
                return false;
            }
            if (splitOutText.Where(w => w.IndexOf("Обновление конфигурации успешно завершено") != -1).Count() != 2)
            {
                Common.Log("\n");
                foreach (var str in splitOutText)
                    Common.Log(str, ConsoleColor.Red);
                throw new Exception($"\nВозможно произошла ошибка при обновлении!");
            }

            return true;
        }

        public static bool UpdateBaseToNextRelease(Base1CSettings baseSettings, Base1CInfo baseInfo, ConfUpdateInfo updateInfo)
        {
            Common.Log($"Обновление на релиз {updateInfo.Version}...");

            int exitCode = Common.StartProcessWithArguments(baseInfo.PlatformInfo.ApplicationPath, Platform1C.GetArgumentForBaseConfUpdate(baseSettings, updateInfo.UpdateFilePath));
            if (exitCode < 0)
                throw new Exception($"Возможно произошла ошибка при обновлении, т.к. код результата выполнения = {exitCode}");

            return CheckOutFileAfterConfUpdateOrRaiseException();
        }

        public static void UpdateBaseInUserMode(Base1CSettings baseSettings, Base1CInfo baseInfo)
        {
            Common.Log("Попытка обновления ИБ в пользовательском режиме...");

            Exception result = Common.CallFunctionAtComConnection1CDomain("UpdateBaseInUserMode", baseSettings, baseInfo.PlatformInfo.ComConnectorVersion);
            if (result != null)
                throw result;
        }

        public static void UpdateBase(Base1CSettings baseSettings)
        {
            Common.Log("\n", ConsoleColor.White, false, 0, false);
            Common.Log($"////////////////////////////////////////////////////////////////////////////////", ConsoleColor.Yellow);
            Common.Log($"// ОБНОВЛЕНИЕ ИБ: {baseSettings.Description}", ConsoleColor.Yellow);

            Base1CInfo baseInfo = Base1C.GetBase1CInfo(baseSettings);
            if (baseInfo == null)
                return;

            Common.Log($"Конфигурация: {baseInfo.ConfName}, версия: {baseInfo.ConfVersion}", ConsoleColor.Green);

            if (!CheckBaseUpdateNecessity(baseInfo))
                return;

            try
            {
                if (!PrepareBaseForUpdate(baseSettings, baseInfo))
                    return;

                MakeBaseBackup(baseSettings, baseInfo);

                bool lastUpdateResult = false;
                foreach (var confUpdateInfo in baseInfo.UpdatesSequence)
                {
                    lastUpdateResult = UpdateBaseToNextRelease(baseSettings, baseInfo, confUpdateInfo.Value);
                    if (baseSettings.RunUserModeAfterEveryUpdate)
                    {
                        if (lastUpdateResult)
                            UpdateBaseInUserMode(baseSettings, baseInfo);
                        else
                            throw new Exception("Невозможно запустить обновление ИБ в пользовательском режиме из-за предыдущих ошибок!");
                    }
                }
                    
                if (lastUpdateResult && !baseSettings.RunUserModeAfterEveryUpdate)
                    UpdateBaseInUserMode(baseSettings, baseInfo);

                Common.Log("Обновление успешно завершено!", ConsoleColor.Green);
            }
            catch (Exception e)
            {
                Common.LogException(e);
            }
            finally
            {
                RestoreBaseAfterUpdate(baseSettings, baseInfo);
            }
        }
    }
}
