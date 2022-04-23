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
            try
            {
                Document doc = commandData.Application.ActiveUIDocument.Document;
                Level level1, level2;
                List<RoomTag> roomTags = new List<RoomTag>();
                List<Room> rooms = new List<Room>();
                List<Room> roomsWithoutTag = new List<Room>();
                List<Room> roomsWithTag = new List<Room>();
                GetLevels(doc, out level1, out level2);

                //Выбор стадии
                Phase phase = new FilteredElementCollector(doc)
                    .OfClass(typeof(Phase))
                    .OfCategory(BuiltInCategory.OST_Phases)
                    .OfType<Phase>()
                    .Where(x => x.Name.Equals("Стадия 1"))
                    .FirstOrDefault();

                //Выбор созданной марки помещения
                var roomTagType = new FilteredElementCollector(doc)
                    .OfCategory(BuiltInCategory.OST_RoomTags)
                    .OfType<RoomTagType>()
                    .Where(x => x.FamilyName.Equals("ГВА_МаркаПомещения"))
                    .FirstOrDefault();

                //Проверка наличия созданной марки помещения в проекте
                if (roomTagType == null)
                {
                    TaskDialog.Show("Ошибка", "Не найдено семейство \"ГВА_МаркаПомещения\"");
                    return Result.Cancelled;
                }

                //Активация семейства созданной марки помещения
                Transaction transaction1 = new Transaction(doc);
                transaction1.Start("Активация семейства \"ГВА_МаркаПомещения\"");
                if (!roomTagType.IsActive)
                    roomTagType.Activate();
                transaction1.Commit();

                Transaction transaction2 = new Transaction(doc, "Создание помещений и расстановка марок помещений");
                transaction2.Start();

                doc.Create.NewRooms2(level1, phase);
                doc.Create.NewRooms2(level2, phase);

                //Получение списка помещений
                rooms = GetRooms(doc);

                //Запись для каждого помещения в параметр "Комментарии" значения уровня для вывода в марку помещения
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

                //Получение списка марок помещений
                roomTags = GetRoomTags(doc);
                //Приравнивание списка помещений без марок к списку всех помещений
                roomsWithoutTag = rooms;
                //Получение списка помещений без марок, исключая из списка помещения с марками
                foreach (RoomTag roomTag in roomTags)
                {
                    ElementId roomId = roomTag.Room.Id;
                    roomsWithoutTag.RemoveAll(room => room.Id == roomId);
                }
                //Расстановка марок помещений в помещениях без марок
                foreach (Room roomWithoutTag in roomsWithoutTag)
                {
                    LocationPoint locPoint = (LocationPoint)roomWithoutTag.Location;
                    UV point = new UV(locPoint.Point.X, locPoint.Point.Y);
                    doc.Create.NewRoomTag(new LinkElementId(roomWithoutTag.Id), point, null);
                }
                //Обновление списка всех марок помещений
                roomTags = GetRoomTags(doc);
                //Присвоение типа марки помещения, которая нами была создана и выбрана
                foreach (RoomTag roomTag in roomTags)
                {
                    roomTag.RoomTagType = roomTagType;
                }

                transaction2.Commit();
        }
            catch (Exception ex)
            {
                message = ex.Message;
                return Result.Failed;
            }

            return Result.Succeeded;
        }
        //Получение помещений из модели
        public static List<Room> GetRooms(Document doc)
        {
            List<Room> rooms = new FilteredElementCollector(doc)
                .OfCategory(BuiltInCategory.OST_Rooms)
                .OfType<Room>()
                .ToList();
            return rooms;
        }
        //Получение марок помещений из модели
        public static List<RoomTag> GetRoomTags(Document doc)
        {
            List<RoomTag> roomTag = new FilteredElementCollector(doc)
                            .OfCategory(BuiltInCategory.OST_RoomTags)
                            .WhereElementIsNotElementType()
                            .OfType<RoomTag>()
                            .ToList();
            return roomTag;
        }
        //Получение уровней из модели
        public static void GetLevels(Document doc, out Level level1, out Level level2)
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
