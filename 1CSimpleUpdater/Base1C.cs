using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Runtime.InteropServices;
using System.Text;
using System.Threading.Tasks;

namespace _1CSimpleUpdater
{
    class Base1CInfo
    {
        public string ConfName;
        public string ConfVersion;
        public int SessionsCount;
    }

    static class Base1C
    {
        static dynamic Get1CConnection(Base1CSettings baseSettings)
        {
            InstalledPlatformInfo platformInfo = Platform1C.GetAppropriatePlatformInfo(baseSettings.PlatformVersion);
            if (platformInfo == null)
                throw new Exception($"Не удалось найти подходящую версию платформы для конфигурации {baseSettings.Description}, требуемая версия = {baseSettings.PlatformVersion}!");
            if (platformInfo.ComConnector == null)
                throw new Exception($"Для платформы {platformInfo.PlatformVersion} не найден ComConnector!");
            Platform1C.CheckComConnectorInprocServerVersion(platformInfo, baseSettings);

            dynamic connectorInstance = Activator.CreateInstance(platformInfo.ComConnector);
            string connectionString = baseSettings.IBConnectionString;
            if (baseSettings.Login.Length > 0)
            {
                connectionString += $"Usr=\"{baseSettings.Login}\";";
                if (baseSettings.Password.Length > 0)
                    connectionString += $"Pwd=\"{baseSettings.Password}\";";
            }

            try
            {
                dynamic connection = connectorInstance.Connect(connectionString);
                return connection;
            }
            catch (Exception e)
            {
                Common.LogException(e);
                return null;
            }
        }

        public static Base1CInfo GetBase1CInfo(Base1CSettings base1CSettings)
        {
            var connection = Get1CConnection(base1CSettings);
            if (connection == null)
                return null;

            try
            {
                Base1CInfo info = new Base1CInfo();
                info.ConfName = connection.Метаданные.Имя;
                info.ConfVersion = connection.Метаданные.Версия;
                info.SessionsCount = connection.ПолучитьСоединенияИнформационнойБазы().Количество() - 1;

                return info;
            } catch (Exception e)
            {
                Common.LogException(e);
            }
            finally
            {
                Marshal.FinalReleaseComObject(connection);
            }
            
            return null;
        }

        public static bool CheckBaseUpdateNecessity(Base1CInfo baseInfo)
        {

        }

        public static bool PrepareBaseForUpdate(Base1CSettings baseSettings, Base1CInfo baseInfo)
        {
            if (baseSettings.IsServerIB)
            {
                throw new NotImplementedException();
            }

            if (baseInfo.SessionsCount > 0)
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
                        "\"1CSimpleUpdater\"",
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
                    throw new NotImplementedException();
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

        public static void MakeBaseBackup(Base1CSettings baseSettings)
        {

        }

        public static void UpdateBase(Base1CSettings baseSettings)
        {
            Common.Log($"\n////////////////////////////////////////////////////////////////////////////////", ConsoleColor.Yellow);
            Common.Log($"// ОБНОВЛЕНИЕ БАЗЫ: {baseSettings.Description}\n", ConsoleColor.Yellow);

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

                if (baseSettings.BackupsCount > 0)
                    MakeBaseBackup(baseSettings);
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
