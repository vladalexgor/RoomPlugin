using Autodesk.Revit.Attributes;
using Autodesk.Revit.DB;
using Autodesk.Revit.DB.Architecture;
using Autodesk.Revit.UI;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace RoomPlugin
{
    [TransactionAttribute(TransactionMode.Manual)]
    public class RoomPlugin : IExternalCommand
    {
        public Result Execute(ExternalCommandData commandData, ref string message, ElementSet elements)
        {            
            //Не забыть писать комментарии к коду
            Document doc = commandData.Application.ActiveUIDocument.Document;
            Level level1, level2;
            GetLevels(doc, out level1, out level2);

            Phase phase = new FilteredElementCollector(doc)
                .OfClass(typeof(Phase))
                .OfCategory(BuiltInCategory.OST_Phases)
                .OfType<Phase>()
                .Where(x => x.Name.Equals("Стадия 1"))
                .FirstOrDefault();

            
            FamilySymbol familySymbol = new FilteredElementCollector(doc)
                .OfClass(typeof(FamilySymbol))
                .OfCategory(BuiltInCategory.OST_RoomTags)
                .OfType<FamilySymbol>()
                .Where(x => x.FamilyName.Equals("ГВА_МаркаПомещения"))
                .FirstOrDefault();

            if (familySymbol == null)
            {
                TaskDialog.Show("Ошибка", "Не найдено семейство \"ГВА_МаркаПомещения\"");
                return Result.Cancelled;
            }

            Transaction transaction1 = new Transaction(doc);
            transaction1.Start("Активация семейства \"ГВА_МаркаПомещения\"");
            if (!familySymbol.IsActive)
                familySymbol.Activate();
            transaction1.Commit();

            Transaction transaction2 = new Transaction(doc, "Создание помещений и марок помещений");
            transaction2.Start();

            doc.Create.NewRooms2(level1, phase);
            doc.Create.NewRooms2(level2, phase);

            List<Room> rooms = new FilteredElementCollector(doc)
                .OfCategory(BuiltInCategory.OST_Rooms)
                .OfType<Room>()
                .ToList();

            foreach (Room room in rooms)
            {
                if (room.Level.Name == level1.Name)
                {
                    room.get_Parameter(BuiltInParameter.ALL_MODEL_INSTANCE_COMMENTS).Set("1");
                }
                else if (room.Level.Name == level2.Name)
                {
                    room.get_Parameter(BuiltInParameter.ALL_MODEL_INSTANCE_COMMENTS).Set("2");
                }
            }

            transaction2.Commit();

            return Result.Succeeded;
        }
        private static void GetLevels(Document doc, out Level level1, out Level level2)
        {
            List<Level> listLevel = new FilteredElementCollector(doc)
                            .OfClass(typeof(Level))
                            .OfType<Level>()
                            .ToList();

            level1 = listLevel
                .Where(x => x.Name.Equals("Уровень 1"))
                .FirstOrDefault();
            level2 = listLevel
                .Where(x => x.Name.Equals("Уровень 2"))
                .FirstOrDefault();
        }
    }
}
