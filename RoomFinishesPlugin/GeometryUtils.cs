using Autodesk.Revit.DB;
using System.Collections.Generic;
using System.Diagnostics;
using System.Linq;

namespace RevitPlugin
{
    public static class GeometryUtils
    {
        public static bool IsColumn(Element element)
        {
            if (element is FamilyInstance fi)
            {
                return fi.Category.Id.Value == (long)BuiltInCategory.OST_Columns ||
                       fi.Category.Id.Value == (long)BuiltInCategory.OST_StructuralColumns;
            }
            return false;
        }

        public static Solid GetVisibleColumnSolid(Face columnFace, FamilyInstance column, Document doc, Solid roomSolid)
        {
            Solid columnFaceSolid = CreateSingleSideExtrusion(columnFace, 1.0);
            if (columnFaceSolid == null || columnFaceSolid.Volume < 1.0E-9)
            {
                return null;
            }

            var intersectingWalls = GetIntersectingWalls(column, doc, roomSolid);
            Solid finalSolid = columnFaceSolid;
            Options geomOptions = new Options { DetailLevel = ViewDetailLevel.Fine };

            foreach (Wall wall in intersectingWalls)
            {
                Solid wallSolid = GetElementSolid(wall, geomOptions);
                if (wallSolid == null || wallSolid.Volume < 1.0E-9) continue;

                try
                {
                    Solid resultSolid = BooleanOperationsUtils.ExecuteBooleanOperation(finalSolid, wallSolid, BooleanOperationsType.Difference);
                    if (resultSolid != null && resultSolid.Volume > 1.0E-9)
                    {
                        finalSolid = resultSolid;
                    }
                    else
                    {
                        return null;
                    }
                }
                catch (System.Exception ex)
                {
                    Debug.WriteLine($"Error subtracting wall from column solid: {ex.Message}");
                    return null;
                }
            }
            return finalSolid;
        }

        private static Solid CreateSingleSideExtrusion(Face face, double thickness)
        {
            try
            {
                IList<CurveLoop> loops = face.GetEdgesAsCurveLoops();
                if (loops == null || loops.Count == 0) return null;

                XYZ normal = face.ComputeNormal(new UV(0.5, 0.5));
                if (normal == null || normal.IsZeroLength()) return null;

                return GeometryCreationUtilities.CreateExtrusionGeometry(loops, normal.Negate(), thickness);
            }
            catch (System.Exception ex)
            {
                Debug.WriteLine($"Error creating single side extrusion: {ex.Message}");
                return null;
            }
        }

        private static List<Wall> GetIntersectingWalls(FamilyInstance column, Document doc, Solid roomSolid)
        {
            var columnBox = column.get_BoundingBox(null);
            if (columnBox == null) return new List<Wall>();

            var filter = new BoundingBoxIntersectsFilter(new Outline(columnBox.Min, columnBox.Max));

            var candidateWalls = new FilteredElementCollector(doc)
                .OfClass(typeof(Wall))
                .WhereElementIsNotElementType()
                .WherePasses(filter)
                .Cast<Wall>()
                .ToList();

            var finalWalls = new List<Wall>();
            foreach (var wall in candidateWalls)
            {
                Solid wallSolid = GetElementSolid(wall);
                if (wallSolid == null) continue;

                try
                {
                    var intersection = BooleanOperationsUtils.ExecuteBooleanOperation(wallSolid, roomSolid, BooleanOperationsType.Intersect);
                    if (intersection != null && intersection.Volume > 1e-9)
                    {
                        finalWalls.Add(wall);
                    }
                }
                catch (System.Exception ex)
                {
                    Debug.WriteLine($"Error intersecting wall with room solid: {ex.Message}");
                }
            }
            return finalWalls;
        }

        public static Solid GetElementSolid(Element elem, Options opts = null)
        {
            if (opts == null)
            {
                opts = new Options { DetailLevel = ViewDetailLevel.Fine, ComputeReferences = true };
            }

            try
            {
                GeometryElement geom = elem.get_Geometry(opts);
                if (geom == null) return null;

                Solid mainSolid = null;
                double maxVolume = 0;

                foreach (var obj in geom)
                {
                    if (obj is Solid s && s.Volume > maxVolume)
                    {
                        mainSolid = s;
                        maxVolume = s.Volume;
                    }
                    else if (obj is GeometryInstance gi)
                    {
                        foreach (var instObj in gi.GetInstanceGeometry())
                        {
                            if (instObj is Solid instS && instS.Volume > maxVolume)
                            {
                                mainSolid = instS;
                                maxVolume = instS.Volume;
                            }
                        }
                    }
                }
                return mainSolid;
            }
            catch (System.Exception ex)
            {
                Debug.WriteLine($"Error getting element solid: {ex.Message}");
                return null;
            }
        }
        public static List<Solid> GetRoomOpeningSolids(Autodesk.Revit.DB.Architecture.Room room, Document doc)
        {
            var solids = new List<Solid>();
            var boundaryWalls = GetBoundaryWalls(room);

            foreach (var wall in boundaryWalls)
            {
                var openings = FindOpeningsInWall(doc, wall, room);
                foreach (var opening in openings)
                {
                    var openingSolid = GetElementSolid(opening);
                    if (openingSolid != null && openingSolid.Volume > 0)
                    {
                        solids.Add(openingSolid);
                    }
                }
            }
            return solids;
        }

        private static List<Wall> GetBoundaryWalls(Autodesk.Revit.DB.Architecture.Room room)
        {
            var walls = new List<Wall>();
            var boundarySegments = room.GetBoundarySegments(new SpatialElementBoundaryOptions());
            var doc = room.Document;

            foreach (var loop in boundarySegments)
            {
                foreach (var segment in loop)
                {
                    var wall = doc.GetElement(segment.ElementId) as Wall;
                    if (wall != null)
                    {
                        walls.Add(wall);
                    }
                }
            }
            return walls.Distinct().ToList();
        }

        public static List<FamilyInstance> FindOpeningsInWall(Document doc, Wall wall, Autodesk.Revit.DB.Architecture.Room room)
        {
            var openings = new FilteredElementCollector(doc)
                .OfClass(typeof(FamilyInstance))
                .WhereElementIsNotElementType()
                .Cast<FamilyInstance>()
                .Where(fi => fi.Category != null && fi.Host != null && fi.Host.Id == wall.Id &&
                    (fi.Category.Id.Value == (long)BuiltInCategory.OST_Doors ||
                     fi.Category.Id.Value == (long)BuiltInCategory.OST_Windows))
                .ToList();
            return openings.Where(fi => IsInstanceRelatedToRoom(fi, room)).ToList();
        }

        private static bool IsInstanceRelatedToRoom(FamilyInstance fi, Autodesk.Revit.DB.Architecture.Room room)
        {
            if (fi.Room?.Id == room.Id) return true;
            if (fi.FromRoom?.Id == room.Id) return true;
            if (fi.ToRoom?.Id == room.Id) return true;
            return false;
        }
    }
}
