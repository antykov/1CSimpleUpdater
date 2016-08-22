using _1CSimpleUpdater;
using System;
using System.Collections.Generic;
using System.Runtime.InteropServices;

namespace ComConnection1C
{
    public class ComConnection1C : MarshalByRefObject
    {
        dynamic GetComConnection(Base1CSettings baseSettings, string comConnectorVersion)
        {
            dynamic connectorInstance = Activator.CreateInstance(Type.GetTypeFromProgID(comConnectorVersion));
            string connectionString = baseSettings.IBConnectionString;
            if (baseSettings.Login.Length > 0)
            {
                connectionString += $"Usr=\"{baseSettings.Login}\";";
                if (baseSettings.Password.Length > 0)
                    connectionString += $"Pwd=\"{baseSettings.Password}\";";
            }
            connectionString += $"UC=\"{Platform1C.SessionLockCode}\";";

            return connectorInstance.Connect(connectionString);
        }

        public Base1CInfo GetBase1CInfo(Base1CSettings baseSettings, string comConnectorVersion)
        {
            dynamic connection = GetComConnection(baseSettings, comConnectorVersion);
            if (connection == null)
                return null;

            try
            {
                Base1CInfo info = new Base1CInfo();
                info.ConfName = connection.Метаданные.Имя;
                info.ConfVersion = connection.Метаданные.Версия;
                info.SessionsCount = connection.ПолучитьСоединенияИнформационнойБазы().Количество() - 1;
                info.UpdatesSequence = null;
                info.BackupFilePath = "";
                info.PlatformInfo = null;

                return info;
            }
            catch (Exception e)
            {
                LogException(e);
                return null;
            }
            finally
            {
                Marshal.FinalReleaseComObject(connection);
            }
        }

        public Exception UpdateBaseInUserMode(Base1CSettings baseSettings, string comConnectorVersion)
        {
            dynamic connection = GetComConnection(baseSettings, comConnectorVersion);
            if (connection == null)
                return new Exception("Не удалось установить COM-соединение!");

            try
            {
                string updateResult = "";

                string version = connection.Метаданные.Версия;
                if (connection.Метаданные.Имя == "ЗарплатаИУправлениеПерсоналом" && version.Substring(0, 3) == "2.5")
                {
                    connection.ОбновлениеИнформационнойБазыЗК.ВыполнитьОбновлениеИнформационнойБазы();
                } else if (connection.Метаданные.Имя == "БухгалтерияПредприятия" && version.Substring(0, 3) == "2.0")
                {
                    connection.ОбновлениеИнформационнойБазы.ВыполнитьОбновлениеИнформационнойБазы();
                } else if (connection.Метаданные.Имя == "УправлениеТорговлей" && version.Substring(0, 4) == "10.3")
                {
                    updateResult = connection.ОбновлениеИнформационнойБазы.ВыполнитьОбновлениеИнформационнойБазы();
                } else
                {
                    connection.ОбновлениеИнформационнойБазыСлужебный.УстановитьЗапускОбновленияИнформационнойБазы(true);
                    updateResult = connection.ОбновлениеИнформационнойБазы.ВыполнитьОбновлениеИнформационнойБазы();
                }

                if (!String.IsNullOrWhiteSpace(updateResult) && updateResult != "Успешно")
                    throw new Exception($"Результат обновления: {updateResult}");

                return null;
            }
            catch (Exception e)
            {
                return e;
            }
            finally
            {
                Marshal.FinalReleaseComObject(connection);
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

        public void LogException(Exception E, string info = "")
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
    }
}
