using Autodesk.Revit.DB;
using Autodesk.Revit.DB.Architecture;
using Autodesk.Revit.UI;
using RevitPlugin.RoomFinishesPlugin;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Windows;

namespace RevitPlugin
{
    public partial class RoomWindow : Window
    {
        private Document _doc;
        private UIDocument _uidoc;
        private List<RoomData> _data;

        // Объекты External Event для безопасного вызова Revit API
        private ExternalEvent _highlightEvent;
        private ExternalEvent _writeEvent;
        private HighlightHandler _highlightHandler;
        private WriteHandler _writeHandler;

        // Чтобы окно можно было таскать мышкой
        private void Window_MouseLeftButtonDown(object sender, System.Windows.Input.MouseButtonEventArgs e)
        {
            try
            {
                this.DragMove();
            }
            catch { }
        }

        // Кнопка "Крестик"
        private void btnClose_Click(object sender, RoutedEventArgs e)
        {
            this.Close();
        }

        public RoomWindow(Document doc, UIDocument uidoc)
        {
            InitializeComponent();
            _doc = doc;
            _uidoc = uidoc;

            // 1. Инициализация обработчиков
            _highlightHandler = new HighlightHandler(this);
            _writeHandler = new WriteHandler(this);
            _highlightEvent = ExternalEvent.Create(_highlightHandler);
            _writeEvent = ExternalEvent.Create(_writeHandler);

            // Проверяем: если на складе уже что-то лежит
            if (DataStorage.CachedRooms != null && DataStorage.CachedRooms.Count > 0)
            {
                // Забираем данные со склада
                _data = DataStorage.CachedRooms;
                
                // Отображаем в таблице
                dgRooms.ItemsSource = _data;
                
                // Активируем кнопки (ведь данные уже есть)
                SetActionButtonsEnabled(true);
            }
        }

        // Вспомогательный метод для управления кнопками
        private void SetActionButtonsEnabled(bool enabled)
        {
            btnHighlight.IsEnabled = enabled;
            btnWrite.IsEnabled = enabled;
        }

        // --- ОБРАБОТЧИК: Кнопка "Рассчитать" ---
        private void btnCalculate_Click(object sender, RoutedEventArgs e)
        {
            SetActionButtonsEnabled(false);

            try
            {
                // Считаем заново
                List<RoomData> newData = RoomProcessor.CalculateRooms(_doc);

                // --- НОВАЯ ЛОГИКА: СОХРАНЕНИЕ ---

                // 1. Обновляем локальную переменную
                _data = newData;

                // 2. Сохраняем на "Вечный склад"
                DataStorage.CachedRooms = newData;

                // 3. Показываем
                dgRooms.ItemsSource = _data;

                // Включаем кнопки действий
                SetActionButtonsEnabled(true);
                MessageBox.Show("Расчет успешно завершен!", "Готово", MessageBoxButton.OK, MessageBoxImage.Information);

            }
            catch (Exception ex)
            {
                MessageBox.Show("Ошибка расчета: " + ex.Message, "Ошибка", MessageBoxButton.OK, MessageBoxImage.Error);
            }
        }

        // --- ОБРАБОТЧИКИ КНОПОК ДЕЙСТВИЙ (Вызывают ExternalEvent) ---

        private void btnWrite_Click(object sender, RoutedEventArgs e)
        {
            // Передаем данные в обработчик и сигнализируем Revit о готовности
            _writeHandler.RoomDataList = _data;
            _writeEvent.Raise();
        }

        private void btnHighlight_Click(object sender, RoutedEventArgs e)
        {
            _highlightHandler.RoomDataList = _data;
            _highlightEvent.Raise();
        }

        // --- МЕТОДЫ, КОТОРЫЕ ВЫЗЫВАЮТСЯ ИЗ ExternalEvent (Безопасно) ---

        // Вызывается из WriteHandler.cs
        public void WriteDataToRevit()
        {
            using (Transaction t = new Transaction(_doc, "Запись отделки"))
            {
                t.Start();
                try
                {
                    foreach (var item in _data)
                    {
                        Room room = _doc.GetElement(item.RoomId) as Room;
                        if (room == null) continue;

                        // ВАЖНО: Замените эти строки на ваши реальные параметры!
                        Parameter pWall = room.LookupParameter("Площадь стен (отделка)");
                        Parameter pSkirt = room.LookupParameter("Длина плинтуса");

                        // 1. Исправляем ПЛОЩАДЬ (Метры кв. -> Футы кв.)
                        if (pWall != null)
                        {
                            // Если параметр типа "Площадь" (Area), Revit ждет футы
                            double areaInFeet = UnitUtils.ConvertToInternalUnits(item.WallArea, UnitTypeId.SquareMeters);
                            pWall.Set(areaInFeet);
                        }
                        // Если параметр текстовый (Text), то можно писать pWall.Set(item.WallArea.ToString());

                        // 2. Исправляем ДЛИНУ (Метры -> Футы)
                        if (pSkirt != null)
                        {
                            // Если параметр типа "Длина" (Length)
                            double lengthInFeet = UnitUtils.ConvertToInternalUnits(item.SkirtingLength, UnitTypeId.Meters);
                            pSkirt.Set(lengthInFeet);
                        }
                    }
                    t.Commit();
                    MessageBox.Show("Данные успешно записаны в Revit!", "Успех", MessageBoxButton.OK, MessageBoxImage.Information);
                }
                catch (Exception ex)
                {
                    t.RollBack();
                    MessageBox.Show("Ошибка при записи в Revit: " + ex.Message);
                }
            }
        }

        // Вызывается из HighlightHandler.cs
        public void HighlightRooms()
        {
            using (Transaction t = new Transaction(_doc, "Подсветка"))
            {
                t.Start();

                // (Логика подсветки без изменений)
                Color color = new Color(0, 255, 0);
                OverrideGraphicSettings settings = new OverrideGraphicSettings();
                settings.SetSurfaceForegroundPatternColor(color);

                FillPatternElement solidFill = new FilteredElementCollector(_doc)
                    .OfClass(typeof(FillPatternElement))
                    .Cast<FillPatternElement>()
                    .FirstOrDefault(fp => fp.GetFillPattern().IsSolidFill);

                if (solidFill != null)
                    settings.SetSurfaceForegroundPatternId(solidFill.Id);

                foreach (var item in _data)
                {
                    Room room = _doc.GetElement(item.RoomId) as Room;
                    if (room == null) continue;

                    SpatialElementBoundaryOptions opts = new SpatialElementBoundaryOptions();
                    IList<IList<BoundarySegment>> loops = room.GetBoundarySegments(opts);
                    if (loops == null) continue;

                    foreach (var segmentList in loops)
                    {
                        foreach (var segment in segmentList)
                        {
                            Element wall = _doc.GetElement(segment.ElementId);
                            if (wall is Wall)
                            {
                                _doc.ActiveView.SetElementOverrides(wall.Id, settings);
                            }
                        }
                    }
                }
                t.Commit();
            }
            _uidoc.RefreshActiveView();
        }
    }
}