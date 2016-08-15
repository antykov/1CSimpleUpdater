using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Xml;
using System.Xml.Serialization;

namespace _1CSimpleUpdater
{
    public class Base1CSettings
    {
        public string Description;
        public string ConnectionString;
        public string PlatformVersion;
        public string Login;
        public string Password;


    }

    public class Settings
    {
        [XmlElement]
        public string BackupsDirectory;
        [XmlElement]
        public string TemplatesDirectory;
        [XmlArray("Bases1C"), XmlArrayItem("Base1CSettings")]
        public List<Base1CSettings> Bases1C;

        public Settings()
        {
            Bases1C = new List<Base1CSettings>();
        }
    }

    public static class AppSettings
    {
        public static Settings settings = new Settings();

        public static void LoadSettings()
        {
            string settingsFilePath = Path.Combine(Environment.CurrentDirectory, "settings.xml");
            if (!File.Exists(settingsFilePath))
            {
                CreateTemplateSettingsFile(settingsFilePath);
                throw new Exception("Создан файл настроек. Заполните настройки и запустите заново!");
            }

            XmlSerializer xmlSerializer = new XmlSerializer(typeof(Settings));
            using (FileStream xmlStream = new FileStream(settingsFilePath, FileMode.Open))
            {
                settings = (Settings)xmlSerializer.Deserialize(xmlStream);
            }
        }

        public static void CheckSettings()
        {
            if (!Directory.Exists(settings.TemplatesDirectory))
                throw new Exception($"Указан несуществующий каталог с обновлениями: {settings.TemplatesDirectory}!");

            if (settings.Bases1C.Count == 0)
                throw new Exception($"Не заполнен список баз для обновлений!");

            if (settings.BackupsDirectory.Length > 0)
                if (!Directory.Exists(settings.BackupsDirectory))
                    Directory.CreateDirectory(settings.BackupsDirectory);
        }

        private static void CreateTemplateSettingsFile(string settingsFilePath)
        {
            XmlSerializer xmlSerializer = new XmlSerializer(typeof(Settings));

            XmlWriterSettings xmlWriterSettings = new XmlWriterSettings();
            xmlWriterSettings.Indent = true;
            xmlWriterSettings.IndentChars = ("\t");
            xmlWriterSettings.OmitXmlDeclaration = true;

            Settings templateSettings = new Settings
            {
                BackupsDirectory = "Каталог для резервных копий (необязательно)",
                TemplatesDirectory = @"Каталог с шаблонами обновлений (...\tmplts)"
            };
            templateSettings.Bases1C.Add(new Base1CSettings
            {
                Description = "Название (Типовая бухгалтерия)",
                ConnectionString = @"Строка подключения (File=""D:\WORK\_Типовые_\_Типовая_БП2_"";, Srvr=""localhost"";Ref=""Accounting"";)",
                PlatformVersion = "Версия платформы (пусто = последняя, 8.2 = последняя из 8.2, 8.2.19.63 = конкретный релиз)",
                Login = "Логин",
                Password = "Пароль"
            });

            using (XmlWriter xmlWriter = XmlWriter.Create(settingsFilePath, xmlWriterSettings))
            {
                xmlSerializer.Serialize(xmlWriter, templateSettings);
            }
        }
    }
}
