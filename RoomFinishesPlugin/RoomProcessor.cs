using Autodesk.Revit.DB;
using Autodesk.Revit.DB.Architecture;
using System;
using System.Collections.Generic;
using System.Linq;

namespace RevitPlugin
{
    public class RoomData
    {
        public string RoomName { get; set; }
        public string RoomNumber { get; set; }
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

            var rooms = new FilteredElementCollector(doc)
                .OfCategory(BuiltInCategory.OST_Rooms)
                .WhereElementIsNotElementType()
                .Cast<Room>()
                .Where(r => r.Area > 0)
                .ToList();

            SpatialElementGeometryCalculator calculator = new SpatialElementGeometryCalculator(doc);
            SpatialElementBoundaryOptions boundaryOpts = new SpatialElementBoundaryOptions();
            boundaryOpts.SpatialElementBoundaryLocation = SpatialElementBoundaryLocation.Finish;

            foreach (Room room in rooms)
            {
                // 1. Плинтус
                double cleanPerimeterFeet = 0;
                IList<IList<BoundarySegment>> loops = room.GetBoundarySegments(boundaryOpts);
                if (loops != null)
                {
                    foreach (var loop in loops)
                    {
                        foreach (var seg in loop)
                        {
                            Element boundaryElement = doc.GetElement(seg.ElementId);
                            if (boundaryElement is Wall)
                                cleanPerimeterFeet += seg.GetCurve().Length;
                        }
                    }
                }

                // 2. Площадь стен
                double netWallAreaFeet = 0;

                // ИСПОЛЬЗУЕМ HASHSET ДЛЯ УНИКАЛЬНОСТИ
                // Это решает проблему, когда одна дверь считается несколько раз,
                // если стена имеет сложную геометрию (несколько граней).
                HashSet<ElementId> uniqueOpenings = new HashSet<ElementId>();
                HashSet<ElementId> subtractedOpenings = new HashSet<ElementId>(); // Чтобы не вычитать площадь дважды

                try
                {
                    SpatialElementGeometryResults geomResults = calculator.CalculateSpatialElementGeometry(room);
                    Solid roomSolid = geomResults.GetGeometry();

                    foreach (Face face in roomSolid.Faces)
                    {
                        IList<SpatialElementBoundarySubface> subfaces = geomResults.GetBoundaryFaceInfo(face);
                        foreach (var subFace in subfaces)
                        {
                            Element boundaryEl = doc.GetElement(subFace.SpatialBoundaryElement.HostElementId);
                            if (boundaryEl is Wall wall)
                            {
                                // Передаем наши коллекции для учета уникальности
                                netWallAreaFeet += CalculateFaceAreaWithOpenings(doc, face, wall, room, uniqueOpenings, subtractedOpenings);
                            }
                        }
                    }
                }
                catch
                {
                    double h = room.Volume > 0 ? room.Volume / room.Area : room.get_Parameter(BuiltInParameter.ROOM_HEIGHT).AsDouble();
                    netWallAreaFeet = cleanPerimeterFeet * h;
                }

                // 3. Ширина дверей (для плинтуса)
                double doorsWidthFeet = GetDoorsTotalWidth(doc, room);

                // 4. Финал
                double perimMeters = UnitUtils.ConvertFromInternalUnits(cleanPerimeterFeet, UnitTypeId.Meters);
                double doorW_Meters = UnitUtils.ConvertFromInternalUnits(doorsWidthFeet, UnitTypeId.Meters);
                double wallAreaMeters = UnitUtils.ConvertFromInternalUnits(netWallAreaFeet, UnitTypeId.SquareMeters);
                double avgHeight = (room.Volume > 0 && room.Area > 0)
                    ? UnitUtils.ConvertFromInternalUnits(room.Volume / room.Area, UnitTypeId.Meters)
                    : 0;

                results.Add(new RoomData
                {
                    RoomName = room.get_Parameter(BuiltInParameter.ROOM_NAME).AsString(),
                    RoomNumber = room.get_Parameter(BuiltInParameter.ROOM_NUMBER).AsString(),
                    Height = Math.Round(avgHeight, 2),
                    // Теперь берем количество из уникального списка
                    OpeningsCount = uniqueOpenings.Count,
                    SkirtingLength = Math.Round(Math.Max(0, perimMeters - doorW_Meters), 2),
                    WallArea = Math.Round(Math.Max(0, wallAreaMeters), 2),
                    RoomId = room.Id
                });
            }
            return results;
        }

        private static double CalculateFaceAreaWithOpenings(
            Document doc,
            Face face,
            Wall wall,
            Room room,
            HashSet<ElementId> uniqueOpeningsRegistry,
            HashSet<ElementId> subtractedOpeningsRegistry)
        {
            double currentArea = face.Area;
            bool hasHoles = face.EdgeLoops.Size > 1;

            var openings = FindOpeningsInWall(doc, wall, room);

            foreach (var op in openings)
            {
                // 1. Регистрируем проем для подсчета количества (HashSet сам исключит дубликаты)
                uniqueOpeningsRegistry.Add(op.Id);

                // 2. Если Revit не вычел дырку сам (hasHoles == false), вычитаем принудительно.
                // НО! Проверяем, не вычитали ли мы этот проем уже на другой грани этой же стены.
                if (!hasHoles && !subtractedOpeningsRegistry.Contains(op.Id))
                {
                    double w = GetTrueWidth(op);
                    double h = GetTrueHeight(op);

                    currentArea -= (w * h);

                    // Запоминаем, что площадь этой двери мы уже вычли
                    subtractedOpeningsRegistry.Add(op.Id);
                }
            }

            return Math.Max(0, currentArea);
        }

        // === ЛОГИКА ПОЛУЧЕНИЯ РАЗМЕРОВ ===

        private static double GetTrueWidth(FamilyInstance fi)
        {
            double w = GetParamValue(fi, BuiltInParameter.DOOR_WIDTH) ??
                       GetParamValue(fi, BuiltInParameter.WINDOW_WIDTH) ??
                       GetParamValue(fi, BuiltInParameter.FAMILY_ROUGH_WIDTH_PARAM) ?? 0;

            if (w < 0.01) w = GetGeometrySize(fi).X;
            return w;
        }

        private static double GetTrueHeight(FamilyInstance fi)
        {
            double h = GetParamValue(fi, BuiltInParameter.DOOR_HEIGHT) ??
                       GetParamValue(fi, BuiltInParameter.WINDOW_HEIGHT) ??
                       GetParamValue(fi, BuiltInParameter.FAMILY_ROUGH_HEIGHT_PARAM) ?? 0;

            if (h < 0.01) h = GetGeometrySize(fi).Z;
            return h;
        }

        private static double? GetParamValue(FamilyInstance fi, BuiltInParameter param)
        {
            double? val = fi.get_Parameter(param)?.AsDouble();
            if (val.HasValue && val.Value > 0.001) return val;

            if (fi.Symbol != null)
            {
                val = fi.Symbol.get_Parameter(param)?.AsDouble();
                if (val.HasValue && val.Value > 0.001) return val;
            }
            return null;
        }

        private static XYZ GetGeometrySize(FamilyInstance fi)
        {
            BoundingBoxXYZ bb = fi.get_BoundingBox(null);
            if (bb == null) return XYZ.Zero;
            return new XYZ(
                Math.Abs(bb.Max.X - bb.Min.X),
                Math.Abs(bb.Max.Y - bb.Min.Y),
                Math.Abs(bb.Max.Z - bb.Min.Z)
            );
        }

        // === ВСПОМОГАТЕЛЬНЫЕ МЕТОДЫ ===

        private static List<FamilyInstance> FindOpeningsInWall(Document doc, Wall wall, Room room)
        {
            var openings = new FilteredElementCollector(doc)
                .OfClass(typeof(FamilyInstance))
                .WhereElementIsNotElementType()
                .Cast<FamilyInstance>()
                .Where(fi =>
                    fi.Category != null && fi.Host != null && fi.Host.Id == wall.Id &&
                    (fi.Category.Id.Value == (long)BuiltInCategory.OST_Doors ||
                     fi.Category.Id.Value == (long)BuiltInCategory.OST_Windows))
                .ToList();

            return openings.Where(fi => IsInstanceRelatedToRoom(fi, room)).ToList();
        }

        private static bool IsInstanceRelatedToRoom(FamilyInstance fi, Room room)
        {
            if (fi.Room?.Id == room.Id) return true;
            if (fi.FromRoom?.Id == room.Id) return true;
            if (fi.ToRoom?.Id == room.Id) return true;
            return false;
        }

        private static double GetDoorsTotalWidth(Document doc, Room room)
        {
            var doors = new FilteredElementCollector(doc)
               .OfCategory(BuiltInCategory.OST_Doors)
               .WhereElementIsNotElementType()
               .Cast<FamilyInstance>()
               .Where(d => IsInstanceRelatedToRoom(d, room));

            double width = 0;
            foreach (var d in doors) width += GetTrueWidth(d);
            return width;
        }
    }
}