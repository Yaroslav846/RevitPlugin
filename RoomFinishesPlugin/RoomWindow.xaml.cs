using Autodesk.Revit.DB;
using Autodesk.Revit.DB.Architecture;
using Autodesk.Revit.UI;
using RevitPlugin.RoomFinishesPlugin;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Threading.Tasks;
using System.Windows;
using System.Windows.Controls; // Для TreeViewItem
using System.Windows.Data;
using System.ComponentModel;

namespace RevitPlugin
{
    public partial class RoomWindow : Window
    {
        private Document _doc;
        private UIDocument _uidoc;
        private List<RoomData> _data;
        private ICollectionView _roomsView;

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
                UpdateRoomData(DataStorage.CachedRooms);
            }
        }

        private void Window_MouseLeftButtonDown(object sender, System.Windows.Input.MouseButtonEventArgs e) => DragMove();
        private void btnClose_Click(object sender, RoutedEventArgs e) => Close();

        private void btnSettings_Click(object sender, RoutedEventArgs e)
        {
            SettingsWindow settingsWin = new SettingsWindow();
            settingsWin.Owner = this;
            settingsWin.ShowDialog();
        }

        private void SetActionButtonsEnabled(bool enabled)
        {
            btnHighlightWithOpenings.IsEnabled = enabled;
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

        private void btnHighlightWithOpenings_Click(object sender, RoutedEventArgs e)
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

        private List<RoomData> GetFilteredData()
        {
            if (_roomsView != null)
            {
                return _roomsView.Cast<RoomData>().ToList();
            }
            return _data;
        }

        // --- ЛОГИКА ФИЛЬТРАЦИИ ---

        private void txtSearch_TextChanged(object sender, System.Windows.Controls.TextChangedEventArgs e)
        {
            if (_roomsView != null) _roomsView.Refresh();
        }

        // НОВОЕ: Обработка выбора в дереве
        private void tvLevels_SelectedItemChanged(object sender, RoutedPropertyChangedEventArgs<object> e)
        {
            if (_roomsView != null) _roomsView.Refresh();
        }

        private bool FilterRooms(object item)
        {
            var room = item as RoomData;
            if (room == null) return false;

            // 1. Проверка по поиску
            bool searchMatch = true;
            if (!string.IsNullOrEmpty(txtSearch.Text))
            {
                string searchText = txtSearch.Text;
                searchMatch = (room.RoomName != null && room.RoomName.IndexOf(searchText, StringComparison.OrdinalIgnoreCase) >= 0) ||
                              (room.RoomNumber != null && room.RoomNumber.IndexOf(searchText, StringComparison.OrdinalIgnoreCase) >= 0);
            }

            // 2. Проверка по дереву уровней
            bool levelMatch = true;
            if (tvLevels.SelectedItem is TreeViewItem selectedNode)
            {
                // Заголовок может быть строкой или UI-элементом, приводим к строке
                string levelName = selectedNode.Header.ToString();

                if (levelName != "Все этажи")
                {
                    levelMatch = (room.LevelName == levelName);
                }
            }

            return searchMatch && levelMatch;
        }

        public void UpdateRoomData(List<RoomData> newData)
        {
            this.Dispatcher.Invoke(() =>
            {
                _data = newData;
                DataStorage.CachedRooms = new List<RoomData>(newData);

                // ЗАПОЛНЕНИЕ ДЕРЕВА
                PopulateTreeView();

                _roomsView = CollectionViewSource.GetDefaultView(_data);
                _roomsView.Filter = FilterRooms;

                dgRooms.ItemsSource = _roomsView;

                SetActionButtonsEnabled(true);
                btnCalculate.IsEnabled = true;
                pbStatus.IsIndeterminate = false;
                pbStatus.Value = 100;
                txtStatus.Text = $"Готово! Обработано комнат: {_data.Count}";
            });
        }

        private void PopulateTreeView()
        {
            tvLevels.Items.Clear();

            // Корневой элемент
            TreeViewItem rootItem = new TreeViewItem { Header = "Все этажи", IsExpanded = true };

            // Собираем уникальные уровни
            var levels = _data.Select(r => r.LevelName).Distinct().OrderBy(l => l).ToList();

            foreach (string lvl in levels)
            {
                rootItem.Items.Add(new TreeViewItem { Header = lvl });
            }

            tvLevels.Items.Add(rootItem);

            // Выбираем корень по умолчанию (чтобы показать все)
            rootItem.IsSelected = true;
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

                    var visibleRooms = GetFilteredData();

                    foreach (var item in visibleRooms)
                    {
                        Room room = _doc.GetElement(item.RoomId) as Room;
                        if (room == null) continue;

                        Solid roomSolid = GeometryUtils.GetRoomSolidWithOpeningsSubtracted(room, _doc);

                        if (roomSolid != null)
                        {
                            try
                            {
                                DirectShape ds = DirectShape.CreateElement(_doc, new ElementId(BuiltInCategory.OST_GenericModel));
                                ds.SetShape(new List<GeometryObject> { roomSolid });
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

        private void CleanUpOldHighlights()
        {
            var ids = new FilteredElementCollector(_doc).OfClass(typeof(DirectShape))
                .WhereElementIsNotElementType().Where(e => e.Name == HighlightElementName).Select(e => e.Id).ToList();
            if (ids.Any()) _doc.Delete(ids);
        }
    }
}
