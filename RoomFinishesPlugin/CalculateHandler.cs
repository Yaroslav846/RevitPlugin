using Autodesk.Revit.DB;
using Autodesk.Revit.UI;
using RevitPlugin;
using RevitPlugin.RoomFinishesPlugin; // Добавили для доступа к DataStorage
using System;
using System.Collections.Generic;
using System.Windows;

namespace RevitPlugin
{
    public class CalculateHandler : IExternalEventHandler
    {
        private RoomWindow _window;

        public CalculateHandler(RoomWindow window)
        {
            _window = window;
        }

        public void Execute(UIApplication app)
        {
            Document doc = app.ActiveUIDocument.Document;

            try
            {
                // 1. ПРИНУДИТЕЛЬНАЯ ОЧИСТКА ПАМЯТИ
                // Гарантируем, что старые данные не повлияют на новый расчет
                DataStorage.CachedRooms = null;

                // 2. Используем транзакцию с Commit
                // RollBack часто отменяет не только изменения, но и пересчет кэша геометрии Revit (Spatial Calculator).
                // Commit "зафиксирует" актуальное состояние геометрии для этого сеанса.
                using (Transaction t = new Transaction(doc, "Расчет отделки (Плагин)"))
                {
                    t.Start();

                    // Обновляем модель, чтобы Revit увидел новые двери/окна
                    doc.Regenerate();

                    // Считаем данные
                    List<RoomData> data = RoomProcessor.CalculateRooms(doc);

                    // Передаем данные окну
                    _window.UpdateRoomData(data);

                    // ВАЖНО: Фиксируем транзакцию. 
                    // Так как мы не меняли параметры элементов, модель не испортится,
                    // но геометрия точно обновится.
                    t.Commit();
                }
            }
            catch (Exception ex)
            {
                MessageBox.Show("Ошибка при расчете:\n" + ex.ToString());

                try
                {
                    _window.Dispatcher.Invoke(() => {
                        // Если упало - разблокируем интерфейс
                        _window.UnlockUI();
                    });
                }
                catch { }
            }
        }

        public string GetName()
        {
            return "CalculateRoomsHandler";
        }
    }
}