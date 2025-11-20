using Autodesk.Revit.DB;
using Autodesk.Revit.DB.Architecture;
using Autodesk.Revit.UI;
using RevitPlugin.RoomFinishesPlugin;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Threading.Tasks;
using System.Windows;

namespace RevitPlugin
{
    public partial class RoomWindow : Window
    {
        private Document _doc;
        private UIDocument _uidoc;
        private List<RoomData> _data;

        private ExternalEvent _highlightEvent;
        private ExternalEvent _writeEvent;
        private ExternalEvent _calculateEvent;

        private HighlightHandler _highlightHandler;
        private WriteHandler _writeHandler;
        private CalculateHandler _calculateHandler;

        private const string HighlightElementName = "Plugin_RoomHighlight_Temp";

        public RoomWindow(Document doc, UIDocument uidoc)
        {
            InitializeComponent();
            _doc = doc;
            _uidoc = uidoc;

            _highlightHandler = new HighlightHandler(this);
            _writeHandler = new WriteHandler(this);
            _calculateHandler = new CalculateHandler(this);

            _highlightEvent = ExternalEvent.Create(_highlightHandler);
            _writeEvent = ExternalEvent.Create(_writeHandler);
            _calculateEvent = ExternalEvent.Create(_calculateHandler);

            if (DataStorage.CachedRooms != null && DataStorage.CachedRooms.Count > 0)
            {
                _data = DataStorage.CachedRooms;
                dgRooms.ItemsSource = _data;
                SetActionButtonsEnabled(true);
                txtStatus.Text = $"Загружено: {_data.Count} комнат";
                pbStatus.Value = 100;
            }
        }

        private void Window_MouseLeftButtonDown(object sender, System.Windows.Input.MouseButtonEventArgs e) => DragMove();
        private void btnClose_Click(object sender, RoutedEventArgs e) => Close();
        private void SetActionButtonsEnabled(bool enabled) { btnHighlight.IsEnabled = enabled; btnWrite.IsEnabled = enabled; }

        public void UnlockUI()
        {
            SetActionButtonsEnabled(true);
            btnCalculate.IsEnabled = true;
            pbStatus.IsIndeterminate = false;
            txtStatus.Text = "Готов (после ошибки)";
        }

        // --- РАСЧЕТ (Кнопка "Рассчитать все") ---
        private void btnCalculate_Click(object sender, RoutedEventArgs e)
        {
            SetActionButtonsEnabled(false);
            btnCalculate.IsEnabled = false;
            dgRooms.ItemsSource = null;

            txtStatus.Text = "Ожидание Revit (расчет)...";
            pbStatus.IsIndeterminate = true;

            _calculateEvent.Raise();
        }

        // --- ЗАПИСЬ (Кнопка "Записать в Revit") ---
        private void btnWrite_Click(object sender, RoutedEventArgs e)
        {
            // ИСПРАВЛЕНО: Убрана строка _writeHandler.RoomDataList = _data;
            // Теперь WriteHandler сам пересчитывает актуальные данные внутри транзакции.

            SetActionButtonsEnabled(false);
            txtStatus.Text = "Пересчет и запись данных...";
            pbStatus.IsIndeterminate = true;

            _writeEvent.Raise();
        }

        // --- ПОДСВЕТКА (Кнопка "Показать в 3D") ---
        private void btnHighlight_Click(object sender, RoutedEventArgs e)
        {
            // Для подсветки данные передаем как обычно (если в HighlightHandler это свойство осталось)
            // Если HighlightHandler не менялся, эта строка нужна.
            if (_highlightHandler != null)
            {
                _highlightHandler.RoomDataList = _data;
            }
            _highlightEvent.Raise();
        }

        // Метод обновления данных (вызывается и из CalculateHandler, и из WriteHandler)
        public void UpdateRoomData(List<RoomData> newData)
        {
            this.Dispatcher.Invoke(() =>
            {
                _data = newData;
                DataStorage.CachedRooms = new List<RoomData>(newData);

                dgRooms.ItemsSource = null;
                dgRooms.ItemsSource = _data;

                SetActionButtonsEnabled(true);
                btnCalculate.IsEnabled = true;
                pbStatus.IsIndeterminate = false;
                pbStatus.Value = 100;
                txtStatus.Text = $"Готово! Обработано комнат: {_data.Count}";
            });
        }

        public void HighlightRooms()
        {
            using (Transaction t = new Transaction(_doc, "Подсветка"))
            {
                t.Start();
                CleanUpOldHighlights();

                OverrideGraphicSettings settings = new OverrideGraphicSettings();
                settings.SetSurfaceForegroundPatternColor(new Color(0, 255, 0));
                settings.SetSurfaceTransparency(40);
                var solidFill = new FilteredElementCollector(_doc).OfClass(typeof(FillPatternElement))
                    .Cast<FillPatternElement>().FirstOrDefault(f => f.GetFillPattern().IsSolidFill);
                if (solidFill != null) settings.SetSurfaceForegroundPatternId(solidFill.Id);

                SpatialElementGeometryCalculator calculator = new SpatialElementGeometryCalculator(_doc);

                foreach (var item in _data)
                {
                    Room room = _doc.GetElement(item.RoomId) as Room;
                    if (room == null) continue;

                    List<GeometryObject> facesToHighlight = new List<GeometryObject>();

                    try
                    {
                        SpatialElementGeometryResults results = calculator.CalculateSpatialElementGeometry(room);
                        Solid roomSolid = results.GetGeometry();

                        foreach (Face face in roomSolid.Faces)
                        {
                            foreach (var subface in results.GetBoundaryFaceInfo(face))
                            {
                                Element host = _doc.GetElement(subface.SpatialBoundaryElement.HostElementId);
                                if (host is Wall wall)
                                {
                                    bool revitFoundHoles = face.EdgeLoops.Size > 1;
                                    bool hasOpenings = WallHasOpenings(wall, room);

                                    if (!revitFoundHoles && hasOpenings)
                                    {
                                        Mesh wallMesh = GetMeshFromWallGeometry(wall, face);
                                        if (wallMesh != null) facesToHighlight.Add(wallMesh);
                                    }
                                    else
                                    {
                                        Mesh m = face.Triangulate();
                                        if (m != null) facesToHighlight.Add(m);
                                    }
                                    break;
                                }
                            }
                        }
                    }
                    catch { continue; }

                    if (facesToHighlight.Count > 0)
                    {
                        try
                        {
                            DirectShape ds = DirectShape.CreateElement(_doc, new ElementId(BuiltInCategory.OST_GenericModel));
                            ds.SetShape(facesToHighlight);
                            ds.Name = HighlightElementName;
                            _doc.ActiveView.SetElementOverrides(ds.Id, settings);
                        }
                        catch { }
                    }
                }
                t.Commit();
            }
            _uidoc.RefreshActiveView();
        }

        private Mesh GetMeshFromWallGeometry(Wall wall, Face roomFace)
        {
            GeometryElement wallGeom = wall.get_Geometry(new Options { DetailLevel = ViewDetailLevel.Fine });
            if (wallGeom == null) return null;

            Mesh bestMesh = null;
            double minDistance = double.MaxValue;

            foreach (var obj in wallGeom)
            {
                if (obj is Solid solid && solid.Volume > 0)
                {
                    foreach (Face wallFace in solid.Faces)
                    {
                        BoundingBoxUV bb = roomFace.GetBoundingBox();
                        UV centerUV = (bb.Min + bb.Max) / 2.0;
                        XYZ centerRoom = roomFace.Evaluate(centerUV);

                        IntersectionResult res = wallFace.Project(centerRoom);

                        if (res != null && res.Distance < 0.1)
                        {
                            if (wallFace.EdgeLoops.Size > 1) return wallFace.Triangulate();
                            if (res.Distance < minDistance)
                            {
                                minDistance = res.Distance;
                                bestMesh = wallFace.Triangulate();
                            }
                        }
                    }
                }
            }
            return bestMesh;
        }

        private bool WallHasOpenings(Wall wall, Room room)
        {
            return new FilteredElementCollector(_doc)
                .OfClass(typeof(FamilyInstance))
                .WhereElementIsNotElementType()
                .Cast<FamilyInstance>()
                .Any(fi => fi.Category != null && fi.Host != null && fi.Host.Id == wall.Id &&
                          (fi.Category.Id.Value == (long)BuiltInCategory.OST_Doors ||
                           fi.Category.Id.Value == (long)BuiltInCategory.OST_Windows));
        }

        private void CleanUpOldHighlights()
        {
            var ids = new FilteredElementCollector(_doc).OfClass(typeof(DirectShape))
                .WhereElementIsNotElementType().Where(e => e.Name == HighlightElementName).Select(e => e.Id).ToList();
            if (ids.Any()) _doc.Delete(ids);
        }
    }
}