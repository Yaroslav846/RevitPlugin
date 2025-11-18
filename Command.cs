// Command.cs
using Autodesk.Revit.Attributes;
using Autodesk.Revit.DB;
using Autodesk.Revit.UI;
using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Windows.Media.Imaging;

namespace RevitPlugin
{
    // Важно: Manual транзакция и ручная регенерация для немодальных окон
    [Transaction(TransactionMode.Manual)]
    [Regeneration(RegenerationOption.Manual)]
    public class MainCommand : IExternalCommand
    {
        public static string AddInPath = typeof(MainCommand).Assembly.Location; // Путь к нашей сборке

        public Result Execute(ExternalCommandData commandData, ref string message, ElementSet elements)
        {

            // Проверяем, не открыто ли уже окно
            if (System.Windows.Application.Current.Windows.OfType<RoomWindow>().Count() > 0)
            {
                System.Windows.Application.Current.Windows.OfType<RoomWindow>().First().Activate();
                return Result.Succeeded;
            }

            RoomWindow window = new RoomWindow(
                commandData.Application.ActiveUIDocument.Document,
                commandData.Application.ActiveUIDocument
            );

            window.Show(); // Non-modal window

            return Result.Succeeded;
        }
    }
}