using Autodesk.Revit.DB;
using System.Collections.Generic;
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
                catch
                {
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
            catch
            {
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
                catch { }
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
            catch
            {
                return null;
            }
        }
    }
}
