using Autodesk.Revit.DB;
using Autodesk.Revit.DB.Architecture;
using Autodesk.Revit.UI;
using RevitPlugin.RoomFinishesPlugin;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Threading.Tasks;
using System.Windows;
using System.Windows.Data; // Важно для ICollectionView
using System.ComponentModel; // Важно для сортировки

namespace RevitPlugin
{
    public partial class RoomWindow : Window
    {
        private Document _doc;
        private UIDocument _uidoc;
        private List<RoomData> _data;
        private ICollectionView _roomsView; // Вид для фильтрации и сортировки

        private ExternalEvent _highlightEvent;
        private ExternalEvent _writeEvent;
        private ExternalEvent _calculateEvent;

        private HighlightHandler _highlightHandler;
        private WriteHandler _writeHandler;
        private CalculateHandler _calculateHandler;

        private const string HighlightElementName = "Plugin_RoomHighlight_Temp";

        private bool _isClearMode = false;

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
                // Используем метод UpdateRoomData, чтобы сразу настроить поиск
                UpdateRoomData(DataStorage.CachedRooms);
            }
        }

        private void Window_MouseLeftButtonDown(object sender, System.Windows.Input.MouseButtonEventArgs e) => DragMove();
        private void btnClose_Click(object sender, RoutedEventArgs e) => Close();

        private void SetActionButtonsEnabled(bool enabled)
        {
            btnHighlight.IsEnabled = enabled;
            btnWrite.IsEnabled = enabled;
            if (btnClear != null) btnClear.IsEnabled = enabled;
        }

        public void UnlockUI()
        {
            SetActionButtonsEnabled(true);
            btnCalculate.IsEnabled = true;
            pbStatus.IsIndeterminate = false;
            txtStatus.Text = "Готов (после ошибки)";
        }

        private void btnCalculate_Click(object sender, RoutedEventArgs e)
        {
            SetActionButtonsEnabled(false);
            btnCalculate.IsEnabled = false;
            dgRooms.ItemsSource = null;
            txtStatus.Text = "Ожидание Revit (расчет)...";
            pbStatus.IsIndeterminate = true;
            _calculateEvent.Raise();
        }

        private void btnWrite_Click(object sender, RoutedEventArgs e)
        {
            SetActionButtonsEnabled(false);
            txtStatus.Text = "Пересчет и запись данных...";
            pbStatus.IsIndeterminate = true;
            _writeEvent.Raise();
        }

        private void btnHighlight_Click(object sender, RoutedEventArgs e)
        {
            _isClearMode = false;
            if (_highlightHandler != null) _highlightHandler.RoomDataList = GetFilteredData();
            _highlightEvent.Raise();
        }

        private void btnClear_Click(object sender, RoutedEventArgs e)
        {
            _isClearMode = true;
            if (_highlightHandler != null) _highlightHandler.RoomDataList = GetFilteredData();
            _highlightEvent.Raise();
        }

        // Вспомогательный метод для получения только тех комнат, что видны при поиске
        private List<RoomData> GetFilteredData()
        {
            if (_roomsView != null)
            {
                return _roomsView.Cast<RoomData>().ToList();
            }
            return _data;
        }

        // --- НОВОЕ: ЛОГИКА ПОИСКА ---
        private void txtSearch_TextChanged(object sender, System.Windows.Controls.TextChangedEventArgs e)
        {
            if (_roomsView != null)
            {
                _roomsView.Refresh(); // Применяет фильтр заново
            }
        }

        private bool FilterRooms(object item)
        {
            if (string.IsNullOrEmpty(txtSearch.Text)) return true;

            var room = item as RoomData;
            if (room == null) return false;

            // Ищем по имени или номеру (без учета регистра)
            string searchText = txtSearch.Text;
            return (room.RoomName != null && room.RoomName.IndexOf(searchText, StringComparison.OrdinalIgnoreCase) >= 0) ||
                   (room.RoomNumber != null && room.RoomNumber.IndexOf(searchText, StringComparison.OrdinalIgnoreCase) >= 0);
        }

        public void UpdateRoomData(List<RoomData> newData)
        {
            this.Dispatcher.Invoke(() =>
            {
                _data = newData;
                DataStorage.CachedRooms = new List<RoomData>(newData);

                // Создаем "Представление" (View) для таблицы, которое поддерживает фильтрацию
                _roomsView = CollectionViewSource.GetDefaultView(_data);
                _roomsView.Filter = FilterRooms; // Подключаем наш фильтр

                dgRooms.ItemsSource = _roomsView;

                SetActionButtonsEnabled(true);
                btnCalculate.IsEnabled = true;
                pbStatus.IsIndeterminate = false;
                pbStatus.Value = 100;
                txtStatus.Text = $"Готово! Обработано комнат: {_data.Count}";
            });
        }

        public void HighlightRooms()
        {
            string transName = _isClearMode ? "Сброс подсветки" : "Подсветка отделки";

            using (Transaction t = new Transaction(_doc, transName))
            {
                t.Start();
                CleanUpOldHighlights();

                if (!_isClearMode)
                {
                    OverrideGraphicSettings settings = new OverrideGraphicSettings();
                    settings.SetSurfaceForegroundPatternColor(new Color(0, 255, 0));
                    settings.SetSurfaceTransparency(30);

                    var solidFill = new FilteredElementCollector(_doc).OfClass(typeof(FillPatternElement))
                        .Cast<FillPatternElement>().FirstOrDefault(f => f.GetFillPattern().IsSolidFill);
                    if (solidFill != null) settings.SetSurfaceForegroundPatternId(solidFill.Id);

                    SpatialElementGeometryCalculator calculator = new SpatialElementGeometryCalculator(_doc);
                    Options wallGeomOpts = new Options { DetailLevel = ViewDetailLevel.Fine, ComputeReferences = true };

                    // ВАЖНО: Берем только отфильтрованные данные для подсветки
                    var visibleRooms = GetFilteredData();

                    foreach (var item in visibleRooms)
                    {
                        Room room = _doc.GetElement(item.RoomId) as Room;
                        if (room == null) continue;

                        List<GeometryObject> finalGeometries = new List<GeometryObject>();

                        try
                        {
                            SpatialElementGeometryResults results = calculator.CalculateSpatialElementGeometry(room);
                            Solid roomRawSolid = results.GetGeometry();

                            foreach (Face face in roomRawSolid.Faces)
                            {
                                foreach (var subface in results.GetBoundaryFaceInfo(face))
                                {
                                    Element host = _doc.GetElement(subface.SpatialBoundaryElement.HostElementId);

                                    if (host is Wall wall)
                                    {
                                        bool preciseGeoCreated = false;

                                        try
                                        {
                                            XYZ normal = face.ComputeNormal(new UV(0.5, 0.5));
                                            XYZ extrudeDir = normal.Negate();
                                            Solid roomSensorSolid = CreateExtrusion(face, extrudeDir, 0.15);
                                            Solid wallRealSolid = GetElementSolid(wall, wallGeomOpts);

                                            if (roomSensorSolid != null && wallRealSolid != null)
                                            {
                                                Solid intersection = BooleanOperationsUtils.ExecuteBooleanOperation(
                                                    roomSensorSolid, wallRealSolid, BooleanOperationsType.Intersect);

                                                if (intersection != null && intersection.Volume > 0)
                                                {
                                                    finalGeometries.Add(intersection);
                                                    preciseGeoCreated = true;
                                                }
                                            }
                                        }
                                        catch
                                        {
                                            preciseGeoCreated = false;
                                        }

                                        if (!preciseGeoCreated)
                                        {
                                            Mesh m = face.Triangulate();
                                            if (m != null) finalGeometries.Add(m);
                                        }
                                        break;
                                    }
                                }
                            }
                        }
                        catch { continue; }

                        if (finalGeometries.Count > 0)
                        {
                            try
                            {
                                DirectShape ds = DirectShape.CreateElement(_doc, new ElementId(BuiltInCategory.OST_GenericModel));
                                ds.SetShape(finalGeometries);
                                ds.Name = HighlightElementName;
                                _doc.ActiveView.SetElementOverrides(ds.Id, settings);
                            }
                            catch { }
                        }
                    }
                }

                t.Commit();
            }
            _uidoc.RefreshActiveView();
        }

        private Solid CreateExtrusion(Face face, XYZ direction, double thickness)
        {
            try
            {
                List<CurveLoop> loops = new List<CurveLoop>();
                foreach (EdgeArray ea in face.EdgeLoops)
                {
                    CurveLoop loop = new CurveLoop();
                    foreach (Edge e in ea) loop.Append(e.AsCurve());
                    loops.Add(loop);
                }

                if (direction.IsZeroLength()) return null;
                if (loops.Count == 0 || !loops[0].IsOpen() == false) return null;

                return GeometryCreationUtilities.CreateExtrusionGeometry(loops, direction, thickness);
            }
            catch
            {
                return null;
            }
        }

        private Solid GetElementSolid(Element elem, Options opts)
        {
            try
            {
                GeometryElement geom = elem.get_Geometry(opts);
                if (geom == null) return null;

                Solid mainSolid = null;
                double maxVolume = 0;

                foreach (var obj in geom)
                {
                    if (obj is Solid s && s.Volume > 0)
                    {
                        if (s.Volume > maxVolume) { maxVolume = s.Volume; mainSolid = s; }
                    }
                    else if (obj is GeometryInstance gi)
                    {
                        foreach (var instObj in gi.GetInstanceGeometry())
                        {
                            if (instObj is Solid instS && instS.Volume > 0)
                            {
                                if (instS.Volume > maxVolume) { maxVolume = instS.Volume; mainSolid = instS; }
                            }
                        }
                    }
                }
                return mainSolid;
            }
            catch { return null; }
        }

        private void CleanUpOldHighlights()
        {
            var ids = new FilteredElementCollector(_doc).OfClass(typeof(DirectShape))
                .WhereElementIsNotElementType().Where(e => e.Name == HighlightElementName).Select(e => e.Id).ToList();
            if (ids.Any()) _doc.Delete(ids);
        }
    }
}