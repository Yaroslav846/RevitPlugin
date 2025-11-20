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

        // Названия параметров (должны совпадать с Revit буква в букву)
        private const string P_WALL_AREA = "Площадь стен (отделка)";
        private const string P_SKIRT_LENGTH = "Длина плинтуса";
        private const string P_HEIGHT = "Высота (плагин)";

        public WriteHandler(RoomWindow window)
        {
            _window = window;
        }

        public void Execute(UIApplication app)
        {
            Document doc = app.ActiveUIDocument.Document;

            try
            {
                using (Transaction t = new Transaction(doc, "Запись отделки"))
                {
                    t.Start();
                    doc.Regenerate();

                    // 1. Считаем актуальные данные
                    List<RoomData> freshData = RoomProcessor.CalculateRooms(doc);

                    int updatedCount = 0;
                    List<string> errors = new List<string>();

                    foreach (var item in freshData)
                    {
                        Room room = doc.GetElement(item.RoomId) as Room;
                        if (room == null) continue;

                        // 2. Пытаемся найти параметры
                        Parameter pWall = room.LookupParameter(P_WALL_AREA);
                        Parameter pSkirt = room.LookupParameter(P_SKIRT_LENGTH);
                        Parameter pHeight = room.LookupParameter(P_HEIGHT);

                        // Если параметры не найдены - собираем ошибку и пропускаем (чтобы не спамить окнами)
                        if (pWall == null)
                        {
                            if (!errors.Contains(P_WALL_AREA)) errors.Add(P_WALL_AREA);
                        }
                        if (pSkirt == null)
                        {
                            if (!errors.Contains(P_SKIRT_LENGTH)) errors.Add(P_SKIRT_LENGTH);
                        }

                        // 3. Записываем данные (Умная запись)
                        bool anyUpdated = false;

                        if (pWall != null) anyUpdated |= SetParameterValue(pWall, item.WallArea, UnitTypeId.SquareMeters);
                        if (pSkirt != null) anyUpdated |= SetParameterValue(pSkirt, item.SkirtingLength, UnitTypeId.Meters);
                        if (pHeight != null) anyUpdated |= SetParameterValue(pHeight, item.Height, UnitTypeId.Meters);

                        if (anyUpdated) updatedCount++;
                    }

                    t.Commit();

                    // 4. Обновляем таблицу
                    _window.UpdateRoomData(freshData);

                    // 5. Отчет
                    if (errors.Count > 0)
                    {
                        MessageBox.Show("ВНИМАНИЕ! Данные не записаны, так как в проекте не найдены параметры:\n\n" +
                                        string.Join("\n", errors) +
                                        "\n\nДобавьте их через: Управление -> Параметры проекта -> Для категории 'Помещения'.",
                                        "Ошибка параметров", MessageBoxButton.OK, MessageBoxImage.Warning);
                    }
                    else
                    {
                        MessageBox.Show($"Успешно записано помещений: {updatedCount}");
                    }
                }
            }
            catch (Exception ex)
            {
                MessageBox.Show("Критическая ошибка: " + ex.Message);
                try { _window.Dispatcher.Invoke(() => _window.UnlockUI()); } catch { }
            }
        }

        // Универсальный метод записи (поддерживает и Число, и Текст)
        private bool SetParameterValue(Parameter param, double value, ForgeTypeId unitType)
        {
            if (param.IsReadOnly) return false;

            // Если параметр числовой (Площадь или Длина)
            if (param.StorageType == StorageType.Double)
            {
                double internalValue = UnitUtils.ConvertToInternalUnits(value, unitType);
                return param.Set(internalValue);
            }
            // Если параметр Текстовый (пользователь ошибся типом, но мы все равно запишем)
            else if (param.StorageType == StorageType.String)
            {
                return param.Set(value.ToString("F2"));
            }

            return false;
        }

        public string GetName()
        {
            return "WriteDataToRevitHandler";
        }
    }
}