using Autodesk.Revit.UI;
using RevitPlugin;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace RevitPlugin
{
    public class HighlightHandler : IExternalEventHandler
    {
        private RoomWindow _window;
        public List<RoomData> RoomDataList { get; set; } // Для передачи данных

        public HighlightHandler(RoomWindow window)
        {
            _window = window;
        }

        public void Execute(UIApplication app)
        {
            // Здесь мы находимся в потоке Revit и можем вызывать API
            _window.HighlightRooms();
        }

        public string GetName()
        {
            return "HighlightRoomsHandler";
        }
    }
}
