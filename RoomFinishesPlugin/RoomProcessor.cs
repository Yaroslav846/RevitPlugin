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
        public string LevelName { get; set; } // НОВОЕ: Имя уровня
        public double Height { get; set; }
        public int OpeningsCount { get; set; }
        public double SkirtingLength { get; set; }
        public double WallArea { get; set; }
        public ElementId RoomId { get; set; }

        public int SortNumber
        {
            get
            {
                if (string.IsNullOrEmpty(RoomNumber)) return 0;
                if (int.TryParse(RoomNumber, out int n)) return n;
                var digits = new string(RoomNumber.TakeWhile(char.IsDigit).ToArray());
                if (int.TryParse(digits, out int n2)) return n2;
                return 0;
            }
        }
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
                HashSet<ElementId> uniqueOpenings = new HashSet<ElementId>();

                try
                {
                    SpatialElementGeometryResults geomResults = calculator.CalculateSpatialElementGeometry(room);
                    Solid roomSolid = geomResults.GetGeometry();

                    foreach (Face face in roomSolid.Faces)
                    {
                        IList<SpatialElementBoundarySubface> subfaces = geomResults.GetBoundaryFaceInfo(face);
                        foreach (var subFace in subfaces)
                        {
                            Element boundaryEl = subFace.SpatialBoundaryElement.GetHostElement();
                            if (boundaryEl is Wall wall)
                            {
                                netWallAreaFeet += CalculateWallArea(doc, subFace, wall, room, uniqueOpenings);
                            }
                            else if (GeometryUtils.IsColumn(boundaryEl))
                            {
                                netWallAreaFeet += GetColumnFaceArea(doc, subFace, boundaryEl as FamilyInstance, room, roomSolid);
                            }
                        }
                    }
                }
                catch
                {
                    double h = room.Volume > 0 ? room.Volume / room.Area : room.get_Parameter(BuiltInParameter.ROOM_HEIGHT).AsDouble();
                    netWallAreaFeet = cleanPerimeterFeet * h;
                }

                // 3. Ширина дверей
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
                    // Получаем имя уровня
                    LevelName = room.Level != null ? room.Level.Name : "Неизвестно",
                    Height = Math.Round(avgHeight, 2),
                    OpeningsCount = uniqueOpenings.Count,
                    SkirtingLength = Math.Round(Math.Max(0, perimMeters - doorW_Meters), 2),
                    WallArea = Math.Round(Math.Max(0, wallAreaMeters), 2),
                    RoomId = room.Id
                });
            }
            return results;
        }

        private static double CalculateWallArea(Document doc, SpatialElementBoundarySubface subFace, Wall wall, Room room, HashSet<ElementId> uniqueOpeningsRegistry)
        {
            Face subfaceFace = subFace.GetSubface();
            if (subfaceFace == null || subfaceFace.Area < 1e-6) return 0;

            XYZ normal = subfaceFace.ComputeNormal(new UV(0.5, 0.5));
            Solid subfaceSolid = GeometryCreationUtilities.CreateExtrusionGeometry(
                subfaceFace.GetEdgesAsCurveLoops(), normal, 1.0);

            var openings = GeometryUtils.FindOpeningsInWall(doc, wall, room);
            foreach (var op in openings)
            {
                uniqueOpeningsRegistry.Add(op.Id);
                Solid openingSolid = GeometryUtils.GetElementSolid(op);
                if (openingSolid != null)
                {
                    try
                    {
                        subfaceSolid = BooleanOperationsUtils.ExecuteBooleanOperation(subfaceSolid, openingSolid, BooleanOperationsType.Difference);
                    }
                    catch { }
                }
            }

            if (subfaceSolid == null) return 0;

            double finalArea = 0;
            foreach (Face f in subfaceSolid.Faces)
            {
                try
                {
                    XYZ faceNormal = f.ComputeNormal(new UV(0.5, 0.5));
                    if (faceNormal.IsAlmostEqualTo(normal))
                    {
                        finalArea += f.Area;
                    }
                }
                catch { }
            }
            return finalArea;
        }

        private static double GetTrueWidth(FamilyInstance fi)
        {
            double w = GetParamValue(fi, BuiltInParameter.DOOR_WIDTH) ??
                       GetParamValue(fi, BuiltInParameter.WINDOW_WIDTH) ??
                       GetParamValue(fi, BuiltInParameter.FAMILY_ROUGH_WIDTH_PARAM) ?? 0;

            if (w < 0.01) w = GetGeometrySize(fi).X;
            return w;
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
            return new XYZ(Math.Abs(bb.Max.X - bb.Min.X), Math.Abs(bb.Max.Y - bb.Min.Y), Math.Abs(bb.Max.Z - bb.Min.Z));
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

        private static double GetColumnFaceArea(Document doc, SpatialElementBoundarySubface subFace, FamilyInstance column, Room room, Solid roomSolid)
        {
            Face subfaceFace = subFace.GetSubface();
            if (subfaceFace == null || subfaceFace.Area < 1e-6) return 0;

            Solid visibleSolid = GeometryUtils.GetVisibleColumnSolid(subfaceFace, column, doc, roomSolid);
            if (visibleSolid == null) return 0;

            XYZ normal = subfaceFace.ComputeNormal(new UV(0.5, 0.5));
            double finalArea = 0;

            foreach (Face f in visibleSolid.Faces)
            {
                try
                {
                    XYZ faceNormal = f.ComputeNormal(new UV(0.5, 0.5));
                    if (faceNormal.IsAlmostEqualTo(normal))
                    {
                        finalArea += f.Area;
                    }
                }
                catch { }
            }
            return finalArea;
        }
    }
}
