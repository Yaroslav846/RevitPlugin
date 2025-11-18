using System;
using System.Collections.Generic;
using System.Linq;
using Autodesk.Revit.DB;
using Autodesk.Revit.DB.Architecture;

namespace RevitPlugin
{
    // Класс данных (обновленный)
    public class RoomData
    {
        public string RoomName { get; set; }
        public string RoomNumber { get; set; }

        // Новые поля
        public double Height { get; set; }
        public int OpeningsCount { get; set; }

        public double SkirtingLength { get; set; }
        public double WallArea { get; set; }
        public ElementId RoomId { get; set; }
    }

    public static class RoomProcessor
    {
        public static List<RoomData> CalculateRooms(Document doc)
        {
            List<RoomData> results = new List<RoomData>();

            // 1. Собираем комнаты
            var rooms = new FilteredElementCollector(doc)
                .OfCategory(BuiltInCategory.OST_Rooms)
                .WhereElementIsNotElementType()
                .Cast<Room>()
                .Where(r => r.Area > 0)
                .ToList();

            // 2. Собираем двери и окна
            var allDoors = new FilteredElementCollector(doc)
                .OfCategory(BuiltInCategory.OST_Doors)
                .WhereElementIsNotElementType()
                .Cast<FamilyInstance>()
                .ToList();

            var allWindows = new FilteredElementCollector(doc)
                .OfCategory(BuiltInCategory.OST_Windows)
                .WhereElementIsNotElementType()
                .Cast<FamilyInstance>()
                .ToList();

            // Настройки границ (по чистовой отделке)
            SpatialElementBoundaryOptions options = new SpatialElementBoundaryOptions();
            options.SpatialElementBoundaryLocation = SpatialElementBoundaryLocation.Finish;

            foreach (Room room in rooms)
            {
                // --- A. ПЕРИМЕТР ---
                double cleanPerimeterFeet = 0;
                IList<IList<BoundarySegment>> loops = room.GetBoundarySegments(options);

                if (loops != null)
                {
                    foreach (var loop in loops)
                    {
                        foreach (var seg in loop)
                        {
                            Element boundaryElement = doc.GetElement(seg.ElementId);
                            // Если это стена - добавляем длину
                            if (boundaryElement is Wall)
                            {
                                cleanPerimeterFeet += seg.GetCurve().Length;
                            }
                        }
                    }
                }

                // --- B. ВЫСОТА (Переменная heightFeet объявляется здесь) ---
                double heightFeet = room.get_Parameter(BuiltInParameter.ROOM_HEIGHT).AsDouble();

                // --- C. ПРОЕМЫ ---
                var roomDoors = allDoors.Where(d => IsInstanceRelatedToRoom(d, room)).ToList();
                var roomWindows = allWindows.Where(w => IsInstanceRelatedToRoom(w, room)).ToList();

                // Считаем количество (для новой колонки)
                int totalOpenings = roomDoors.Count + roomWindows.Count;

                double doorsWidthFeet = 0;
                double doorsAreaFeet = 0;
                double windowsAreaFeet = 0;

                foreach (var door in roomDoors)
                {
                    double w = door.Symbol.get_Parameter(BuiltInParameter.DOOR_WIDTH)?.AsDouble() ?? 0;
                    double h = door.Symbol.get_Parameter(BuiltInParameter.DOOR_HEIGHT)?.AsDouble() ?? 0;
                    doorsWidthFeet += w;
                    doorsAreaFeet += (w * h);
                }

                foreach (var win in roomWindows)
                {
                    double w = win.Symbol.get_Parameter(BuiltInParameter.WINDOW_WIDTH)?.AsDouble() ?? 0;
                    double h = win.Symbol.get_Parameter(BuiltInParameter.WINDOW_HEIGHT)?.AsDouble() ?? 0;
                    windowsAreaFeet += (w * h);
                }

                // --- D. КОНВЕРТАЦИЯ ЕДИНИЦ ---
                double perimMeters = UnitUtils.ConvertFromInternalUnits(cleanPerimeterFeet, UnitTypeId.Meters);

                // Переменная heightMeters (для таблицы)
                double heightMeters = UnitUtils.ConvertFromInternalUnits(heightFeet, UnitTypeId.Meters);

                double doorW_Meters = UnitUtils.ConvertFromInternalUnits(doorsWidthFeet, UnitTypeId.Meters);

                double doorArea_Meters = UnitUtils.ConvertFromInternalUnits(doorsAreaFeet, UnitTypeId.SquareMeters);
                double winArea_Meters = UnitUtils.ConvertFromInternalUnits(windowsAreaFeet, UnitTypeId.SquareMeters);

                // --- E. МАТЕМАТИКА (Переменные skirtingFinal и wallsNet объявляются здесь) ---

                // Плинтус
                double skirtingFinal = perimMeters - doorW_Meters;

                // Стены
                // Сначала переводим площадь стен (брутто) в метры
                // Важно: Считаем Gross в футах, потом конвертируем, ИЛИ берем метры * метры.
                // Правильнее: метры * метры
                double wallsGrossMeters = perimMeters * heightMeters;

                double wallsNet = wallsGrossMeters - doorArea_Meters - winArea_Meters;

                // --- F. СОЗДАНИЕ ОБЪЕКТА ---
                results.Add(new RoomData
                {
                    RoomName = room.get_Parameter(BuiltInParameter.ROOM_NAME).AsString(),
                    RoomNumber = room.get_Parameter(BuiltInParameter.ROOM_NUMBER).AsString(),

                    // Новые колонки
                    Height = Math.Round(heightMeters, 2),
                    OpeningsCount = totalOpenings,

                    // Результаты расчетов (с защитой от отрицательных чисел)
                    SkirtingLength = Math.Round(Math.Max(0, skirtingFinal), 2),
                    WallArea = Math.Round(Math.Max(0, wallsNet), 2),

                    RoomId = room.Id
                });
            }

            return results;
        }

        // Вспомогательный метод
        private static bool IsInstanceRelatedToRoom(FamilyInstance fi, Room room)
        {
            if (fi.Room?.Id == room.Id) return true;
            if (fi.FromRoom?.Id == room.Id) return true;
            if (fi.ToRoom?.Id == room.Id) return true;
            return false;
        }
    }
}