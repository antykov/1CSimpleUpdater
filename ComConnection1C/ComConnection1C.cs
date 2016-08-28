using _1CSimpleUpdater;
using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Runtime.InteropServices;
using System.Runtime.Remoting.Lifetime;
using System.Threading;

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

        Tuple<dynamic, dynamic> GetComConnectionToCluster(Base1CSettings baseSettings, string comConnectorVersion, out string ibsrvr, out string ibname)
        {
            ibname = "";
            ibsrvr = "";
            string ibport = "";

            try
            {
                string[] srvrparts = GetParsedIBConnectionString(baseSettings.IBConnectionString)["SRVR"].Split(':').Select(s => s.Trim()).ToArray();
                ibsrvr = srvrparts[0];
                ibport = "";
                if (srvrparts.Length == 2)
                {
                    int port;
                    if (Int32.TryParse(srvrparts[1], out port))
                        ibport = $":{(port - 1)}";
                }
                ibname = GetParsedIBConnectionString(baseSettings.IBConnectionString)["REF"].ToUpper();
            }
            catch (Exception e)
            {
                LogException(e, $"Не удалось получить путь к серверу из строки подключения {baseSettings.IBConnectionString}");
                return null;
            }

            dynamic connectorInstance = Activator.CreateInstance(Type.GetTypeFromProgID(comConnectorVersion));
            dynamic agentConnection = connectorInstance.ConnectAgent($"{ibsrvr}{ibport}");
            if (agentConnection == null)
                return null;

            try
            {
                dynamic clusters = agentConnection.GetClusters();
                if (clusters.Length == 0)
                    throw new Exception($"Отсутствуют кластеры 1С на сервере {ibsrvr}!");
                dynamic cluster = clusters[0];

                agentConnection.Authenticate(cluster, baseSettings.ClusterAdministratorLogin, baseSettings.ClusterAdministratorPassword);

                dynamic processes = agentConnection.GetWorkingProcesses(cluster);
                if (processes.Length == 0)
                    throw new Exception($"Отсутствуют рабочие процессы в кластере 1С на сервере {ibsrvr}!");
                dynamic process = connectorInstance.ConnectWorkingProcess($"{ibsrvr}:{processes[0].MainPort}");

                process.AddAuthentication(baseSettings.Login, baseSettings.Password);

                dynamic infobases = process.GetInfoBases();
                dynamic infobase = null;
                for (int i = 0; i < infobases.Length; i++)
                    if (infobases[i].Name.ToUpper() == ibname)
                    {
                        infobase = infobases[i];
                        break;
                    }
                if (infobase == null)
                    throw new Exception($"Не удалось найти ИБ {ibname} в кластере 1С на сервере {ibsrvr}!");

                return new Tuple<dynamic, dynamic>(process, infobase);
            }
            catch (Exception e)
            {
                LogException(e);
                return null;
            }
        }

        public bool UnlockServerBase(Base1CSettings baseSettings, string comConnectorVersion)
        {
            return UnlockServerBase(baseSettings, comConnectorVersion, null);
        }

        public bool UnlockServerBase(Base1CSettings baseSettings, string comConnectorVersion, Tuple<dynamic, dynamic> existing)
        {
            dynamic process = null, infobase = null;

            if (existing == null)
            {
                string ibsrvr = "", ibname = "";
                Tuple<dynamic, dynamic> tuple = GetComConnectionToCluster(baseSettings, comConnectorVersion, out ibsrvr, out ibname);
                if (tuple == null)
                    return false;
                process = tuple.Item1;
                infobase = tuple.Item2;
            } else
            {
                process = existing.Item1;
                infobase = existing.Item2;
            }

            try
            {
                infobase.ConnectDenied = false;
                infobase.SessionsDenied = false;
                infobase.ScheduledJobsDenied = !baseSettings.EnableScheduledJobs;
                infobase.DeniedFrom = DateTime.MinValue.ToString("yyyy-MM-dd HH:mm:ss");
                infobase.DeniedTo = DateTime.MinValue.ToString("yyyy-MM-dd HH:mm:ss");
                infobase.DeniedMessage = "";
                infobase.PermissionCode = "";
                process.UpdateInfoBase(infobase);

                return true;
            }
            catch (Exception e)
            {
                LogException(e, "Не удалось восстановить доступ ИБ для работы!!!");
                return false;
            }
        }

        public bool CheckServerBaseLockPossibilityAndLock(Base1CSettings baseSettings, string comConnectorVersion)
        {
            List<string> AppIdsTerminate = new List<string> { "DESIGNER", "BACKGROUNDJOB" };
            List<string> AppIdsIgnore = new List<string> { "SRVRCONSOLE", "COMCONSOLE" };

            string ibsrvr = "", ibname = "";

            Tuple<dynamic, dynamic> tuple = GetComConnectionToCluster(baseSettings, comConnectorVersion, out ibsrvr, out ibname);
            if (tuple == null)
                return false;

            try
            {
                dynamic process = tuple.Item1, infobase = tuple.Item2;

                dynamic connections = process.GetInfoBaseConnections(infobase);
                foreach (var connection in connections)
                    if (AppIdsTerminate.IndexOf(connection.AppId.ToUpper()) == -1 && AppIdsIgnore.IndexOf(connection.AppId.ToUpper()) == -1)
                        return false;

                infobase.ConnectDenied = true;
                infobase.SessionsDenied = true;
                infobase.ScheduledJobsDenied = true;
                infobase.DeniedFrom = DateTime.Now.ToString("yyyy-MM-dd HH:mm:ss");
                infobase.DeniedTo = DateTime.Now.AddHours(6).ToString("yyyy-MM-dd HH:mm:ss");
                infobase.DeniedMessage = "Установлена блокировка соединений для обновления конфигурации!";
                infobase.PermissionCode = Platform1C.SessionLockCode;
                process.UpdateInfoBase(infobase);

                foreach (var connection in connections)
                    if (AppIdsTerminate.IndexOf(connection.AppId.ToUpper()) != -1)
                    {
                        Log($"Отключение соединения (AppID={connection.AppId}, HostName={connection.HostName}, UserName={connection.UserName})...");
                        process.Disconnect(connection);
                    }

                connections = process.GetInfoBaseConnections(infobase);
                foreach (var connection in connections)
                    if ((AppIdsTerminate.IndexOf(connection.AppId.ToUpper()) == -1 && AppIdsIgnore.IndexOf(connection.AppId.ToUpper()) == -1) ||
                        (AppIdsTerminate.IndexOf(connection.AppId.ToUpper()) != -1))
                        throw new Exception("Не удалось завершить сеансы работы с ИБ!");

                return true;
            }
            catch (Exception e)
            {
                LogException(e);

                UnlockServerBase(baseSettings, comConnectorVersion, tuple);

                return false;
            }
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
                info.SessionsCount = (baseSettings.IsServerIB) ? -1 : connection.ПолучитьСоединенияИнформационнойБазы().Количество() - 1;
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
                    connection.ОбновлениеИнформационнойБазыСлужебный.ЗаписатьПодтверждениеЛегальностиПолученияОбновлений();
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

        public Dictionary<string, string> GetParsedIBConnectionString(string connectionString)
        {
            try
            {
                return connectionString.Split(';')
                    .Where(w => !String.IsNullOrWhiteSpace(w))
                        .Select(s =>
                        {
                            string[] arr = s.Split('=');
                            if (arr.Length == 2)
                                return new { key = arr[0].ToUpper(), value = arr[1].Replace("\"", "") };
                            else
                                return new { key = "", value = "" };
                        })
                        .ToDictionary(v => v.key, v => v.value);
            }
            catch (Exception e)
            {
                throw new Exception($"Не удалось распарсить строку подключения к серверу {connectionString}", e);
            }
        }

        public void Log(string message, ConsoleColor color = ConsoleColor.White, bool newLine = true, int emptyLineLength = 0, bool logToFile = true)
        {
            if (logToFile)
                File.AppendAllText(AppSettings.settings.LogFileName, $"{DateTime.Now.ToString()} {message}\n");

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
            {
                Console.WriteLine(E.Message);
                File.AppendAllText(AppSettings.settings.LogFileName, $"{DateTime.Now.ToString()} [ERROR] {E.Message}\n");
            }
            else
            {
                Console.WriteLine($"{info}: {E.Message}");
                File.AppendAllText(AppSettings.settings.LogFileName, $"{DateTime.Now.ToString()} [ERROR] {info}: {E.Message}\n");
            }
                
            Exception inner = E.InnerException;
            while (inner != null)
            {
                Console.WriteLine($"    --> {inner.Message}");
                File.AppendAllText(AppSettings.settings.LogFileName, $"{DateTime.Now.ToString()} [ERROR]    --> {inner.Message}\n");

                inner = inner.InnerException;
            }
        }
    }
}
