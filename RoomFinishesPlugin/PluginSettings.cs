using System;
using System.IO;
using System.Xml.Serialization;

namespace RevitPlugin
{
    // Класс для хранения настроек
    public class PluginSettings
    {
        // Значения по умолчанию
        public string WallAreaParam { get; set; } = "Площадь стен (отделка)";
        public string SkirtingParam { get; set; } = "Длина плинтуса";
        public string HeightParam { get; set; } = "Высота (плагин)";

        // Путь к файлу настроек: C:\Users\Name\AppData\Roaming\RevitRoomPlugin\settings.xml
        private static string _folder = Path.Combine(Environment.GetFolderPath(Environment.SpecialFolder.ApplicationData), "RevitRoomPlugin");
        private static string _path = Path.Combine(_folder, "settings.xml");

        // Метод загрузки
        public static PluginSettings Load()
        {
            try
            {
                if (File.Exists(_path))
                {
                    XmlSerializer ser = new XmlSerializer(typeof(PluginSettings));
                    using (StreamReader sr = new StreamReader(_path))
                    {
                        return (PluginSettings)ser.Deserialize(sr);
                    }
                }
            }
            catch { }

            // Если файла нет или ошибка - возвращаем стандартные
            return new PluginSettings();
        }

        // Метод сохранения
        public void Save()
        {
            try
            {
                if (!Directory.Exists(_folder)) Directory.CreateDirectory(_folder);

                XmlSerializer ser = new XmlSerializer(typeof(PluginSettings));
                using (StreamWriter sw = new StreamWriter(_path))
                {
                    ser.Serialize(sw, this);
                }
            }
            catch { }
        }
    }
}