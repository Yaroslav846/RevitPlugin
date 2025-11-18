using Autodesk.Revit.UI;
using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Reflection;
using System.Windows.Media.Imaging;

namespace RevitPlugin
{
    public class App : IExternalApplication
    {
        // Метод для извлечения картинки из ресурсов DLL
        private BitmapImage GetEmbeddedImage(Assembly a, string fileName)
        {
            // Ищем ресурс, имя которого заканчивается на fileName
            // В вашем случае он найдет "RevitPlugin.icon16.png"
            string resourceName = a.GetManifestResourceNames()
                .FirstOrDefault(x => x.EndsWith(fileName));

            if (resourceName == null) return null;

            using (Stream s = a.GetManifestResourceStream(resourceName))
            {
                BitmapImage img = new BitmapImage();
                img.BeginInit();
                img.StreamSource = s;
                img.CacheOption = BitmapCacheOption.OnLoad; // Важно: грузим в память сразу
                img.EndInit();
                return img;
            }
        }

        public Result OnStartup(UIControlledApplication application)
        {
            CreateRibbonPanel(application);
            return Result.Succeeded;
        }

        public Result OnShutdown(UIControlledApplication application)
        {
            return Result.Succeeded;
        }

        private void CreateRibbonPanel(UIControlledApplication application)
        {
            string tabName = "ОТДЕЛКА";
            string panelName = "Помещения";

            // 1. Создаем вкладку
            try { application.CreateRibbonTab(tabName); } catch (Exception) { }

            // 2. Создаем или находим панель
            RibbonPanel panel = null;
            List<RibbonPanel> panels = application.GetRibbonPanels(tabName);
            foreach (RibbonPanel p in panels)
            {
                if (p.Name == panelName)
                {
                    panel = p;
                    break;
                }
            }

            if (panel == null)
            {
                panel = application.CreateRibbonPanel(tabName, panelName);
            }

            // 3. Создаем кнопку
            string assemblyPath = Assembly.GetExecutingAssembly().Location;

            PushButtonData btnData = new PushButtonData(
                "RoomFinishesBtn",
                "Отделка\nПомещений",
                assemblyPath,
                "RevitPlugin.MainCommand"
            );

            // 4. Загружаем иконки (САМАЯ ВАЖНАЯ ЧАСТЬ)
            Assembly assembly = Assembly.GetExecutingAssembly();

            // Получаем картинки через наш метод
            // Имена совпадают с теми, что были на вашем скриншоте (концовки)
            btnData.Image = GetEmbeddedImage(assembly, "icon16.png");
            btnData.LargeImage = GetEmbeddedImage(assembly, "icon32.png");

            btnData.ToolTip = "Расчет площади стен и плинтусов с вычетом проемов";

            // 5. Добавляем кнопку на панель
            panel.AddItem(btnData);
        }
    }
}