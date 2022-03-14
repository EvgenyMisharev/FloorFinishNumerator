using Autodesk.Revit.DB;
using Autodesk.Revit.DB.Architecture;
using Autodesk.Revit.UI;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading;
using System.Threading.Tasks;

namespace FinishNumerator
{
    [Autodesk.Revit.Attributes.Transaction(Autodesk.Revit.Attributes.TransactionMode.Manual)]
    class FinishNumeratorCommand : IExternalCommand
    {
        FinishNumeratorProgressBarWPF finishNumeratorProgressBarWPF;
        public Result Execute(ExternalCommandData commandData, ref string message, ElementSet elements)
        {
            Document doc = commandData.Application.ActiveUIDocument.Document;

            FinishNumeratorWPF finishNumeratorWPF = new FinishNumeratorWPF();
            finishNumeratorWPF.ShowDialog();
            if (finishNumeratorWPF.DialogResult != true)
            {
                return Result.Cancelled;
            }

            string finishNumberingSelectedName = finishNumeratorWPF.FinishNumberingSelectedName;

            if (finishNumberingSelectedName == "rbt_EndToEndThroughoutTheProject")
            {
                List<Room> roomList = new FilteredElementCollector(doc)
                    .OfClass(typeof(SpatialElement))
                    .WhereElementIsNotElementType()
                    .Where(r => r.GetType() == typeof(Room))
                    .Cast<Room>()
                    .Where(r => r.Area > 0)
                    .OrderBy(r => (doc.GetElement(r.LevelId) as Level).Elevation)
                    .ToList();

                using (Transaction t = new Transaction(doc))
                {
                    t.Start("Нумерация отделки");
                    //Типы полов для формы
                    List<FloorType> floorTypesList = new FilteredElementCollector(doc)
                        .OfClass(typeof(FloorType))
                        .Where(f => f.Category.Id.IntegerValue.Equals((int)BuiltInCategory.OST_Floors))
                        .Where(f => f.get_Parameter(BuiltInParameter.ALL_MODEL_MODEL) != null)
                        .Where(f => f.get_Parameter(BuiltInParameter.ALL_MODEL_MODEL).AsString() == "Пол"
                        || f.get_Parameter(BuiltInParameter.ALL_MODEL_MODEL).AsString() == "Полы")
                        .Cast<FloorType>()
                        .OrderBy(f => f.Name, new AlphanumComparatorFastString())
                        .ToList();

                    Thread newWindowThread = new Thread(new ThreadStart(ThreadStartingPoint));
                    newWindowThread.SetApartmentState(ApartmentState.STA);
                    newWindowThread.IsBackground = true;
                    newWindowThread.Start();
                    int step = 0;
                    Thread.Sleep(100);
                    finishNumeratorProgressBarWPF.pb_FloorCreatorProgressBar.Dispatcher.Invoke(() => finishNumeratorProgressBarWPF.pb_FloorCreatorProgressBar.Minimum = 0);
                    finishNumeratorProgressBarWPF.pb_FloorCreatorProgressBar.Dispatcher.Invoke(() => finishNumeratorProgressBarWPF.pb_FloorCreatorProgressBar.Maximum = floorTypesList.Count);

                    foreach (FloorType floorType in floorTypesList)
                    {
                        step++;
                        finishNumeratorProgressBarWPF.pb_FloorCreatorProgressBar.Dispatcher.Invoke(() => finishNumeratorProgressBarWPF.pb_FloorCreatorProgressBar.Value = step);
                        finishNumeratorProgressBarWPF.pb_FloorCreatorProgressBar.Dispatcher.Invoke(() => finishNumeratorProgressBarWPF.label_ItemName.Content = floorType.Name);

                        List<Floor> floorList = new FilteredElementCollector(doc)
                           .OfClass(typeof(Floor))
                           .Cast<Floor>()
                           .Where(f => f.FloorType.get_Parameter(BuiltInParameter.ALL_MODEL_MODEL) != null)
                           .Where(f => f.FloorType.get_Parameter(BuiltInParameter.ALL_MODEL_MODEL).AsString() == "Пол"
                           || f.FloorType.get_Parameter(BuiltInParameter.ALL_MODEL_MODEL).AsString() == "Полы")
                           .Where(f => f.FloorType.Id == floorType.Id)
                           .ToList();
                        if (floorList.Count == 0) continue;


                        //Очистка параметра "Помещение_Список номеров"
                        if (floorList.First().LookupParameter("АР_НомераПомещенийПоТипуПола") == null)
                        {
                            TaskDialog.Show("Revit", "У пола отсутствует параметр экземпляра \"АР_НомераПомещенийПоТипуПола\"");
                            finishNumeratorProgressBarWPF.Dispatcher.Invoke(() => finishNumeratorProgressBarWPF.Close());
                            return Result.Cancelled;
                        }

                        foreach (Floor floor in floorList)
                        {
                            floor.LookupParameter("АР_НомераПомещенийПоТипуПола").Set("");
                        }

                        List<string> roomNumbersList = new List<string>();
                        foreach (Floor floor in floorList)
                        {
                            Solid floorSolid = null;
                            GeometryElement geomFloorElement = floor.get_Geometry(new Options());
                            foreach (GeometryObject geomObj in geomFloorElement)
                            {
                                floorSolid = geomObj as Solid;
                                if (floorSolid != null) break;
                            }
                            if(floorSolid != null)
                            {
                                floorSolid = SolidUtils.CreateTransformed(floorSolid, Transform.CreateTranslation(new XYZ(0, 0, 500 / 304.8)));
                            }

                            foreach (Room room in roomList)
                            {
                                Solid roomSolid = null;
                                GeometryElement geomRoomElement = room.get_Geometry(new Options());
                                foreach (GeometryObject geomObj in geomRoomElement)
                                {
                                    roomSolid = geomObj as Solid;
                                    if (roomSolid != null) break;
                                }
                                if(roomSolid != null)
                                {
                                    Solid intersection = null;
                                    try
                                    {
                                        intersection = BooleanOperationsUtils.ExecuteBooleanOperation(floorSolid, roomSolid, BooleanOperationsType.Intersect);
                                    }
                                    catch
                                    {
                                        //ПРОПИСАТЬ ЛОГИКУ СБОРА ОШИБОК!!!!
                                    }
                                    if (intersection != null && intersection.Volume != 0)
                                    {
                                        roomNumbersList.Add(room.Number);
                                    }
                                }
                            }
                        }
                        roomNumbersList.Sort(new AlphanumComparatorFastString());
                        string roomNumbersByFloorType = null;
                        foreach (string roomNumber in roomNumbersList)
                        {
                            if (roomNumbersByFloorType == null)
                            {
                                roomNumbersByFloorType += roomNumber;
                            }
                            else
                            {
                                roomNumbersByFloorType += (", " + roomNumber);
                            }
                        }
                        foreach (Floor floor in floorList)
                        {
                            floor.LookupParameter("АР_НомераПомещенийПоТипуПола").Set(roomNumbersByFloorType);
                        }
                    }
                    finishNumeratorProgressBarWPF.Dispatcher.Invoke(() => finishNumeratorProgressBarWPF.Close());
                    t.Commit();
                }
            }
            else if (finishNumberingSelectedName == "rbt_SeparatedByLevels")
            {
                List<Level> levelList = new FilteredElementCollector(doc)
                   .OfClass(typeof(Level))
                   .WhereElementIsNotElementType()
                   .Cast<Level>()
                   .OrderBy(l => l.Elevation)
                   .ToList();

                Thread newWindowThread = new Thread(new ThreadStart(ThreadStartingPoint));
                newWindowThread.SetApartmentState(ApartmentState.STA);
                newWindowThread.IsBackground = true;
                newWindowThread.Start();
                int step = 0;
                Thread.Sleep(100);
                finishNumeratorProgressBarWPF.pb_FloorCreatorProgressBar.Dispatcher.Invoke(() => finishNumeratorProgressBarWPF.pb_FloorCreatorProgressBar.Minimum = 0);
                finishNumeratorProgressBarWPF.pb_FloorCreatorProgressBar.Dispatcher.Invoke(() => finishNumeratorProgressBarWPF.pb_FloorCreatorProgressBar.Maximum = levelList.Count);
                using (Transaction t = new Transaction(doc))
                {
                    t.Start("Нумерация отделки");
                    foreach (Level lv in levelList)
                    {
                        step++;
                        finishNumeratorProgressBarWPF.pb_FloorCreatorProgressBar.Dispatcher.Invoke(() => finishNumeratorProgressBarWPF.pb_FloorCreatorProgressBar.Value = step);
                        finishNumeratorProgressBarWPF.pb_FloorCreatorProgressBar.Dispatcher.Invoke(() => finishNumeratorProgressBarWPF.label_ItemName.Content = lv.Name);

                        List<Room> roomList = new FilteredElementCollector(doc)
                            .OfClass(typeof(SpatialElement))
                            .WhereElementIsNotElementType()
                            .Where(r => r.GetType() == typeof(Room))
                            .Cast<Room>()
                            .Where(r => r.Area > 0)
                            .Where(r => r.LevelId == lv.Id)
                            .ToList();

                        //Типы полов для формы
                        List<FloorType> floorTypesList = new FilteredElementCollector(doc)
                            .OfClass(typeof(FloorType))
                            .Where(f => f.Category.Id.IntegerValue.Equals((int)BuiltInCategory.OST_Floors))
                            .Where(f => f.get_Parameter(BuiltInParameter.ALL_MODEL_MODEL) != null)
                            .Where(f => f.get_Parameter(BuiltInParameter.ALL_MODEL_MODEL).AsString() == "Пол"
                            || f.get_Parameter(BuiltInParameter.ALL_MODEL_MODEL).AsString() == "Полы")
                            .Cast<FloorType>()
                            .OrderBy(f => f.Name, new AlphanumComparatorFastString())
                            .ToList();

                        foreach (FloorType floorType in floorTypesList)
                        {
                            List<Floor> floorList = new FilteredElementCollector(doc)
                               .OfClass(typeof(Floor))
                               .Cast<Floor>()
                               .Where(f => f.FloorType.get_Parameter(BuiltInParameter.ALL_MODEL_MODEL) != null)
                               .Where(f => f.FloorType.get_Parameter(BuiltInParameter.ALL_MODEL_MODEL).AsString() == "Пол"
                               || f.FloorType.get_Parameter(BuiltInParameter.ALL_MODEL_MODEL).AsString() == "Полы")
                               .Where(f => f.FloorType.Id == floorType.Id)
                               .Where(f => f.LevelId == lv.Id)
                               .ToList();
                            if (floorList.Count == 0) continue;


                            //Очистка параметра "Помещение_Список номеров"
                            if (floorList.First().LookupParameter("АР_НомераПомещенийПоТипуПола") == null)
                            {
                                TaskDialog.Show("Revit", "У пола отсутствует параметр экземпляра \"АР_НомераПомещенийПоТипуПола\"");
                                finishNumeratorProgressBarWPF.Dispatcher.Invoke(() => finishNumeratorProgressBarWPF.Close());
                                return Result.Cancelled;
                            }

                            foreach (Floor floor in floorList)
                            {
                                floor.LookupParameter("АР_НомераПомещенийПоТипуПола").Set("");
                            }

                            List<string> roomNumbersList = new List<string>();
                            foreach (Floor floor in floorList)
                            {
                                Solid floorSolid = null;
                                GeometryElement geomFloorElement = floor.get_Geometry(new Options());
                                foreach (GeometryObject geomObj in geomFloorElement)
                                {
                                    floorSolid = geomObj as Solid;
                                    if (floorSolid != null) break;
                                }
                                if (floorSolid != null)
                                {
                                    floorSolid = SolidUtils.CreateTransformed(floorSolid, Transform.CreateTranslation(new XYZ(0, 0, 500 / 304.8)));
                                }

                                foreach (Room room in roomList)
                                {
                                    Solid roomSolid = null;
                                    GeometryElement geomRoomElement = room.get_Geometry(new Options());
                                    foreach (GeometryObject geomObj in geomRoomElement)
                                    {
                                        roomSolid = geomObj as Solid;
                                        if (roomSolid != null) break;
                                    }
                                    if (roomSolid != null)
                                    {
                                        Solid intersection = null;
                                        try
                                        {
                                            intersection = BooleanOperationsUtils.ExecuteBooleanOperation(floorSolid, roomSolid, BooleanOperationsType.Intersect);
                                        }
                                        catch
                                        {
                                            //ПРОПИСАТЬ ЛОГИКУ СБОРА ОШИБОК!!!!
                                        }
                                        if (intersection != null && intersection.Volume != 0)
                                        {
                                            roomNumbersList.Add(room.Number);
                                        }
                                    }
                                }
                            }
                            roomNumbersList.Sort(new AlphanumComparatorFastString());
                            string roomNumbersByFloorType = null;
                            foreach (string roomNumber in roomNumbersList)
                            {
                                if (roomNumbersByFloorType == null)
                                {
                                    roomNumbersByFloorType += roomNumber;
                                }
                                else
                                {
                                    roomNumbersByFloorType += (", " + roomNumber);
                                }
                            }
                            foreach (Floor floor in floorList)
                            {
                                floor.LookupParameter("АР_НомераПомещенийПоТипуПола").Set(roomNumbersByFloorType);
                            }
                        }
                    }
                    finishNumeratorProgressBarWPF.Dispatcher.Invoke(() => finishNumeratorProgressBarWPF.Close());
                    t.Commit();
                }
            }

            return Result.Succeeded;
        }
        private void ThreadStartingPoint()
        {
            finishNumeratorProgressBarWPF = new FinishNumeratorProgressBarWPF();
            finishNumeratorProgressBarWPF.Show();
            System.Windows.Threading.Dispatcher.Run();
        }
    }
}
