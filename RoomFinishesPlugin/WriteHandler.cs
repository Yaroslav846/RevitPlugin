using Autodesk.Revit.DB;
using Autodesk.Revit.DB.Architecture;
using Autodesk.Revit.UI;
using RevitPlugin.RoomFinishesPlugin;
using System;
using System.Collections.Generic;
using System.Windows;

namespace RevitPlugin
{
    public class WriteHandler : IExternalEventHandler
    {
        private RoomWindow _window;

        public WriteHandler(RoomWindow window)
        {
            _window = window;
        }

        public void Execute(UIApplication app)
        {
            Document doc = app.ActiveUIDocument.Document;

            // 1. ЗАГРУЖАЕМ НАСТРОЙКИ ПЕРЕД ЗАПИСЬЮ
            // Чтобы использовать актуальные имена параметров
            PluginSettings settings = PluginSettings.Load();
            string pWallName = settings.WallAreaParam;
            string pSkirtName = settings.SkirtingParam;
            string pHeightName = settings.HeightParam;

            try
            {
                using (Transaction t = new Transaction(doc, "Запись отделки"))
                {
                    t.Start();
                    doc.Regenerate();

                    List<RoomData> freshData = RoomProcessor.CalculateRooms(doc);
                    int updatedCount = 0;
                    List<string> errors = new List<string>();

                    foreach (var item in freshData)
                    {
                        Room room = doc.GetElement(item.RoomId) as Room;
                        if (room == null) continue;

                        // 2. Ищем параметры по именам из настроек
                        Parameter pWall = room.LookupParameter(pWallName);
                        Parameter pSkirt = room.LookupParameter(pSkirtName);
                        Parameter pHeight = room.LookupParameter(pHeightName);

                        // Собираем ошибки (каких параметров не хватает)
                        if (pWall == null && !string.IsNullOrWhiteSpace(pWallName))
                        {
                            if (!errors.Contains(pWallName)) errors.Add(pWallName);
                        }
                        if (pSkirt == null && !string.IsNullOrWhiteSpace(pSkirtName))
                        {
                            if (!errors.Contains(pSkirtName)) errors.Add(pSkirtName);
                        }

                        bool anyUpdated = false;

                        if (pWall != null) anyUpdated |= SetParameterValue(pWall, item.WallArea, UnitTypeId.SquareMeters);
                        if (pSkirt != null) anyUpdated |= SetParameterValue(pSkirt, item.SkirtingLength, UnitTypeId.Meters);
                        if (pHeight != null) anyUpdated |= SetParameterValue(pHeight, item.Height, UnitTypeId.Meters);

                        if (anyUpdated) updatedCount++;
                    }

                    t.Commit();
                    _window.UpdateRoomData(freshData);

                    if (errors.Count > 0)
                    {
                        MessageBox.Show("Не найдены параметры:\n" + string.Join("\n", errors) +
                                        "\n\nПроверьте их наличие в проекте или измените имена в настройках (кнопка ⚙).",
                                        "Ошибка параметров", MessageBoxButton.OK, MessageBoxImage.Warning);
                    }
                    else
                    {
                        MessageBox.Show($"Успешно записано: {updatedCount}");
                    }
                }
            }
            catch (Exception ex)
            {
                MessageBox.Show("Ошибка: " + ex.Message);
                try { _window.Dispatcher.Invoke(() => _window.UnlockUI()); } catch { }
            }
        }

        private bool SetParameterValue(Parameter param, double value, ForgeTypeId unitType)
        {
            if (param.IsReadOnly) return false;

            if (param.StorageType == StorageType.Double)
            {
                double internalValue = UnitUtils.ConvertToInternalUnits(value, unitType);
                return param.Set(internalValue);
            }
            else if (param.StorageType == StorageType.String)
            {
                return param.Set(value.ToString("F2"));
            }
            return false;
        }

        public string GetName() => "WriteDataToRevitHandler";
    }
}