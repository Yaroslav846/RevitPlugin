using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace RevitPlugin.RoomFinishesPlugin
{
    // Статический класс - живет пока запущен Revit
    public static class DataStorage
    {
        // Здесь будут лежать наши данные
        public static List<RoomData> CachedRooms { get; set; }
    }
}
