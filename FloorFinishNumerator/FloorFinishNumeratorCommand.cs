using Autodesk.Revit.DB;
using Autodesk.Revit.DB.Architecture;
using Autodesk.Revit.UI;
using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Reflection;
using System.Text.RegularExpressions;
using System.Threading;

namespace FloorFinishNumerator
{
    [Autodesk.Revit.Attributes.Transaction(Autodesk.Revit.Attributes.TransactionMode.Manual)]
    class FloorFinishNumeratorCommand : IExternalCommand
    {
        FloorFinishNumeratorProgressBarWPF floorFinishNumeratorProgressBarWPF;
        public Result Execute(ExternalCommandData commandData, ref string message, ElementSet elements)
        {
            try
            {
                GetPluginStartInfo();
            }
            catch { }

            Document doc = commandData.Application.ActiveUIDocument.Document;

            Guid arRoomBookNumberGUID = new Guid("22868552-0e64-49b2-b8d9-9a2534bf0e14");
            Guid arRoomBookNameGUID = new Guid("b59a22a9-7890-45bd-9f93-a186341eef58");

            FloorFinishNumeratorWPF floorFinishNumeratorWPF = new FloorFinishNumeratorWPF();
            floorFinishNumeratorWPF.ShowDialog();
            if (floorFinishNumeratorWPF.DialogResult != true)
            {
                return Result.Cancelled;
            }

            string floorFinishNumberingSelectedName = floorFinishNumeratorWPF.FloorFinishNumberingSelectedName;
            bool fillRoomBookParameters = floorFinishNumeratorWPF.FillRoomBookParameters;

            if (floorFinishNumberingSelectedName == "rbt_EndToEndThroughoutTheProject")
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
                    t.Start("Нумерация пола");
                    //Типы полов 
                    List<FloorType> floorTypesList = new FilteredElementCollector(doc)
                        .OfCategory(BuiltInCategory.OST_Floors)
                        .WhereElementIsElementType()
                        .Where(f => f.get_Parameter(BuiltInParameter.ALL_MODEL_MODEL) != null)
                        .Where(f => f.get_Parameter(BuiltInParameter.ALL_MODEL_MODEL).AsString() == "Пол"
                        || f.get_Parameter(BuiltInParameter.ALL_MODEL_MODEL).AsString() == "Полы")
                        .Cast<FloorType>()
                        .OrderBy(f => PadNumbers(f.Name))
                        .ToList();

                    Thread newWindowThread = new Thread(new ThreadStart(ThreadStartingPoint));
                    newWindowThread.SetApartmentState(ApartmentState.STA);
                    newWindowThread.IsBackground = true;
                    newWindowThread.Start();
                    int step = 0;
                    Thread.Sleep(100);
                    floorFinishNumeratorProgressBarWPF.pb_FloorFinishNumeratorProgressBar.Dispatcher.Invoke(() => floorFinishNumeratorProgressBarWPF.pb_FloorFinishNumeratorProgressBar.Minimum = 0);
                    floorFinishNumeratorProgressBarWPF.pb_FloorFinishNumeratorProgressBar.Dispatcher.Invoke(() => floorFinishNumeratorProgressBarWPF.pb_FloorFinishNumeratorProgressBar.Maximum = floorTypesList.Count);

                    foreach (FloorType floorType in floorTypesList)
                    {
                        step++;
                        floorFinishNumeratorProgressBarWPF.pb_FloorFinishNumeratorProgressBar.Dispatcher.Invoke(() => floorFinishNumeratorProgressBarWPF.pb_FloorFinishNumeratorProgressBar.Value = step);
                        floorFinishNumeratorProgressBarWPF.pb_FloorFinishNumeratorProgressBar.Dispatcher.Invoke(() => floorFinishNumeratorProgressBarWPF.label_ItemName.Content = floorType.Name);

                        List<Floor> floorList = new FilteredElementCollector(doc)
                           .OfClass(typeof(Floor))
                           .Cast<Floor>()
                           .Where(f => f.FloorType.get_Parameter(BuiltInParameter.ALL_MODEL_MODEL) != null)
                           .Where(f => f.FloorType.get_Parameter(BuiltInParameter.ALL_MODEL_MODEL).AsString() == "Пол"
                           || f.FloorType.get_Parameter(BuiltInParameter.ALL_MODEL_MODEL).AsString() == "Полы")
                           .Where(f => f.FloorType.Id == floorType.Id)
                           .ToList();
                        if (floorList.Count == 0) continue;


                        //Очистка параметра "АР_НомераПомещенийПоТипуПола" и "АР_ИменаПомещенийПоТипуПола"
                        if (floorList.First().LookupParameter("АР_НомераПомещенийПоТипуПола") == null)
                        {
                            TaskDialog.Show("Revit", "У пола отсутствует параметр экземпляра \"АР_НомераПомещенийПоТипуПола\"");
                            floorFinishNumeratorProgressBarWPF.Dispatcher.Invoke(() => floorFinishNumeratorProgressBarWPF.Close());
                            return Result.Cancelled;
                        }

                        //Очистка параметра "АР_RoomBook_Номер" и "АР_RoomBook_Имя"
                        if (fillRoomBookParameters)
                        {
                            if (floorList.First().get_Parameter(arRoomBookNumberGUID) == null)
                            {
                                TaskDialog.Show("Revit", "У пола отсутствует параметр \"АР_RoomBook_Номер\"");
                                floorFinishNumeratorProgressBarWPF.Dispatcher.Invoke(() => floorFinishNumeratorProgressBarWPF.Close());
                                return Result.Cancelled;
                            }
                            if (floorList.First().get_Parameter(arRoomBookNameGUID) == null)
                            {
                                TaskDialog.Show("Revit", "У пола отсутствует параметр \"АР_RoomBook_Имя\"");
                                floorFinishNumeratorProgressBarWPF.Dispatcher.Invoke(() => floorFinishNumeratorProgressBarWPF.Close());
                                return Result.Cancelled;
                            }
                        }

                        foreach (Floor floor in floorList)
                        {
                            floor.LookupParameter("АР_НомераПомещенийПоТипуПола").Set("");
                            floor.LookupParameter("АР_ИменаПомещенийПоТипуПола").Set("");

                            if (fillRoomBookParameters)
                            {
                                floor.get_Parameter(arRoomBookNumberGUID).Set("");
                                floor.get_Parameter(arRoomBookNameGUID).Set("");
                            }
                        }

                        List<string> roomNumbersList = new List<string>();
                        List<string> roomNamesList = new List<string>();
                        Options opt = new Options();
                        opt.DetailLevel = ViewDetailLevel.Fine;
                        foreach (Floor floor in floorList)
                        {
                            Solid floorSolid = null;
                            GeometryElement geomFloorElement = floor.get_Geometry(opt);
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
                                GeometryElement geomRoomElement = room.get_Geometry(opt);
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
                                        XYZ pointForIntersect = null;
                                        FaceArray floorFaceArray = floorSolid.Faces;
                                        foreach (object planarFace in floorFaceArray)
                                        {
                                            if (planarFace is PlanarFace && (planarFace as PlanarFace).FaceNormal.IsAlmostEqualTo(XYZ.BasisZ))
                                            {
                                                List<CurveLoop> curveLoopList = (planarFace as PlanarFace).GetEdgesAsCurveLoops().ToList();
                                                if (curveLoopList.Count != 0)
                                                {
                                                    CurveLoop curveLoop = curveLoopList.First();
                                                    if (curveLoop != null)
                                                    {
                                                        Curve c = curveLoop.First();
                                                        pointForIntersect = c.GetEndPoint(0);
                                                    }
                                                }
                                            }
                                        }
                                        if (pointForIntersect == null) continue;
                                        Curve curve = Line.CreateBound(pointForIntersect, pointForIntersect + (500 / 304.8) * XYZ.BasisZ) as Curve;
                                        SolidCurveIntersection curveIntersection = roomSolid.IntersectWithCurve(curve, new SolidCurveIntersectionOptions());
                                        if (curveIntersection.SegmentCount > 0)
                                        {
                                            if (fillRoomBookParameters)
                                            {
                                                if (floor.get_Parameter(arRoomBookNumberGUID) != null)
                                                {
                                                    floor.get_Parameter(arRoomBookNumberGUID).Set(room.Number);
                                                }
                                                if (floor.get_Parameter(arRoomBookNameGUID) != null)
                                                {
                                                    floor.get_Parameter(arRoomBookNameGUID).Set(room.get_Parameter(BuiltInParameter.ROOM_NAME).AsString());
                                                }
                                            }

                                            if (roomNumbersList.Find(elem => elem == room.Number) == null)
                                            {
                                                roomNumbersList.Add(room.Number);
                                                roomNamesList.Add(room.get_Parameter(BuiltInParameter.ROOM_NAME).AsString());
                                                continue;
                                            }
                                        }
                                    }
                                    if (intersection != null && intersection.Volume != 0)
                                    {
                                        if (fillRoomBookParameters)
                                        {
                                            if (floor.get_Parameter(arRoomBookNumberGUID) != null)
                                            {
                                                floor.get_Parameter(arRoomBookNumberGUID).Set(room.Number);
                                            }
                                            if (floor.get_Parameter(arRoomBookNameGUID) != null)
                                            {
                                                floor.get_Parameter(arRoomBookNameGUID).Set(room.get_Parameter(BuiltInParameter.ROOM_NAME).AsString());
                                            }
                                        }

                                        if (roomNumbersList.Find(elem => elem == room.Number) == null)
                                        {
                                            roomNumbersList.Add(room.Number);
                                            roomNamesList.Add(room.get_Parameter(BuiltInParameter.ROOM_NAME).AsString());
                                        }
                                    }
                                    else
                                    {
                                        XYZ pointForIntersect = null;
                                        FaceArray floorFaceArray = floorSolid.Faces;
                                        foreach (object planarFace in floorFaceArray)
                                        {
                                            if (planarFace is PlanarFace && (planarFace as PlanarFace).FaceNormal.IsAlmostEqualTo(XYZ.BasisZ))
                                            {
                                                List<CurveLoop> curveLoopList = (planarFace as PlanarFace).GetEdgesAsCurveLoops().ToList();
                                                if (curveLoopList.Count != 0)
                                                {
                                                    CurveLoop curveLoop = curveLoopList.First();
                                                    if (curveLoop != null)
                                                    {
                                                        Curve c = curveLoop.First();
                                                        pointForIntersect = c.GetEndPoint(0);
                                                    }
                                                }
                                            }
                                        }
                                        if (pointForIntersect == null) continue;
                                        Curve curve = Line.CreateBound(pointForIntersect, pointForIntersect + (500 / 304.8) * XYZ.BasisZ) as Curve;
                                        SolidCurveIntersection curveIntersection = roomSolid.IntersectWithCurve(curve, new SolidCurveIntersectionOptions());
                                        if (curveIntersection.SegmentCount > 0)
                                        {
                                            if (fillRoomBookParameters)
                                            {
                                                if (floor.get_Parameter(arRoomBookNumberGUID) != null)
                                                {
                                                    floor.get_Parameter(arRoomBookNumberGUID).Set(room.Number);
                                                }
                                                if (floor.get_Parameter(arRoomBookNameGUID) != null)
                                                {
                                                    floor.get_Parameter(arRoomBookNameGUID).Set(room.get_Parameter(BuiltInParameter.ROOM_NAME).AsString());
                                                }
                                            }

                                            if (roomNumbersList.Find(elem => elem == room.Number) == null)
                                            {
                                                roomNumbersList.Add(room.Number);
                                                roomNamesList.Add(room.get_Parameter(BuiltInParameter.ROOM_NAME).AsString());
                                                continue;
                                            }
                                        }
                                    }
                                }
                            }
                        }
                        roomNumbersList.OrderBy(e => PadNumbers(e));
                        roomNamesList = roomNamesList.Distinct().ToList();
                        roomNamesList.OrderBy(e => PadNumbers(e));

                        string roomNumbersByFloorType = null;
                        string roomNamesByFloorType = null;
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

                        foreach (string roomName in roomNamesList)
                        {
                            if (roomNamesByFloorType == null)
                            {
                                roomNamesByFloorType += roomName;
                            }
                            else
                            {
                                roomNamesByFloorType += (", " + roomName);
                            }
                        }

                        foreach (Floor floor in floorList)
                        {
                            floor.LookupParameter("АР_НомераПомещенийПоТипуПола").Set(roomNumbersByFloorType);
                        }

                        foreach (Floor floor in floorList)
                        {
                            floor.LookupParameter("АР_ИменаПомещенийПоТипуПола").Set(roomNamesByFloorType);
                        }
                    }
                    floorFinishNumeratorProgressBarWPF.Dispatcher.Invoke(() => floorFinishNumeratorProgressBarWPF.Close());
                    t.Commit();
                }
            }
            else if (floorFinishNumberingSelectedName == "rbt_SeparatedByLevels")
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
                floorFinishNumeratorProgressBarWPF.pb_FloorFinishNumeratorProgressBar.Dispatcher.Invoke(() => floorFinishNumeratorProgressBarWPF.pb_FloorFinishNumeratorProgressBar.Minimum = 0);
                floorFinishNumeratorProgressBarWPF.pb_FloorFinishNumeratorProgressBar.Dispatcher.Invoke(() => floorFinishNumeratorProgressBarWPF.pb_FloorFinishNumeratorProgressBar.Maximum = levelList.Count);
                using (Transaction t = new Transaction(doc))
                {
                    t.Start("Нумерация пола");
                    foreach (Level lv in levelList)
                    {
                        step++;
                        floorFinishNumeratorProgressBarWPF.pb_FloorFinishNumeratorProgressBar.Dispatcher.Invoke(() => floorFinishNumeratorProgressBarWPF.pb_FloorFinishNumeratorProgressBar.Value = step);
                        floorFinishNumeratorProgressBarWPF.pb_FloorFinishNumeratorProgressBar.Dispatcher.Invoke(() => floorFinishNumeratorProgressBarWPF.label_ItemName.Content = lv.Name);

                        List<Room> roomList = new FilteredElementCollector(doc)
                            .OfClass(typeof(SpatialElement))
                            .WhereElementIsNotElementType()
                            .Where(r => r.GetType() == typeof(Room))
                            .Cast<Room>()
                            .Where(r => r.Area > 0)
                            .Where(r => r.LevelId == lv.Id)
                            .ToList();

                        //Типы полов 
                        List<FloorType> floorTypesList = new FilteredElementCollector(doc)
                            .OfClass(typeof(FloorType))
                            .Where(f => f.Category.Id.IntegerValue.Equals((int)BuiltInCategory.OST_Floors))
                            .Where(f => f.get_Parameter(BuiltInParameter.ALL_MODEL_MODEL) != null)
                            .Where(f => f.get_Parameter(BuiltInParameter.ALL_MODEL_MODEL).AsString() == "Пол"
                            || f.get_Parameter(BuiltInParameter.ALL_MODEL_MODEL).AsString() == "Полы")
                            .Cast<FloorType>()
                            .OrderBy(f => PadNumbers(f.Name))
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


                            //Очистка параметра "АР_НомераПомещенийПоТипуПола" и "АР_ИменаПомещенийПоТипуПола"
                            if (floorList.First().LookupParameter("АР_НомераПомещенийПоТипуПола") == null)
                            {
                                TaskDialog.Show("Revit", "У пола отсутствует параметр экземпляра \"АР_НомераПомещенийПоТипуПола\"");
                                floorFinishNumeratorProgressBarWPF.Dispatcher.Invoke(() => floorFinishNumeratorProgressBarWPF.Close());
                                return Result.Cancelled;
                            }

                            foreach (Floor floor in floorList)
                            {
                                floor.LookupParameter("АР_НомераПомещенийПоТипуПола").Set("");
                                floor.LookupParameter("АР_ИменаПомещенийПоТипуПола").Set("");
                            }

                            List<string> roomNumbersList = new List<string>();
                            List<string> roomNamesList = new List<string>();
                            Options opt = new Options();
                            opt.DetailLevel = ViewDetailLevel.Fine;
                            foreach (Floor floor in floorList)
                            {
                                Solid floorSolid = null;
                                GeometryElement geomFloorElement = floor.get_Geometry(opt);
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
                                    GeometryElement geomRoomElement = room.get_Geometry(opt);
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
                                            XYZ pointForIntersect = null;
                                            FaceArray floorFaceArray = floorSolid.Faces;
                                            foreach (object planarFace in floorFaceArray)
                                            {
                                                if (planarFace is PlanarFace && (planarFace as PlanarFace).FaceNormal.IsAlmostEqualTo(XYZ.BasisZ))
                                                {
                                                    List<CurveLoop> curveLoopList = (planarFace as PlanarFace).GetEdgesAsCurveLoops().ToList();
                                                    if (curveLoopList.Count != 0)
                                                    {
                                                        CurveLoop curveLoop = curveLoopList.First();
                                                        if (curveLoop != null)
                                                        {
                                                            Curve c = curveLoop.First();
                                                            pointForIntersect = c.GetEndPoint(0);
                                                        }
                                                    }
                                                }
                                            }
                                            if (pointForIntersect == null) continue;
                                            Curve curve = Line.CreateBound(pointForIntersect, pointForIntersect + (500 / 304.8) * XYZ.BasisZ) as Curve;
                                            SolidCurveIntersection curveIntersection = roomSolid.IntersectWithCurve(curve, new SolidCurveIntersectionOptions());
                                            if (curveIntersection.SegmentCount > 0)
                                            {
                                                if (fillRoomBookParameters)
                                                {
                                                    if (floor.get_Parameter(arRoomBookNumberGUID) != null)
                                                    {
                                                        floor.get_Parameter(arRoomBookNumberGUID).Set(room.Number);
                                                    }
                                                    if (floor.get_Parameter(arRoomBookNameGUID) != null)
                                                    {
                                                        floor.get_Parameter(arRoomBookNameGUID).Set(room.get_Parameter(BuiltInParameter.ROOM_NAME).AsString());
                                                    }
                                                }

                                                if (roomNumbersList.Find(elem => elem == room.Number) == null)
                                                {
                                                    roomNumbersList.Add(room.Number);
                                                    roomNamesList.Add(room.get_Parameter(BuiltInParameter.ROOM_NAME).AsString());
                                                    continue;
                                                }
                                            }
                                        }
                                        if (intersection != null && intersection.Volume != 0)
                                        {
                                            if (fillRoomBookParameters)
                                            {
                                                if (floor.get_Parameter(arRoomBookNumberGUID) != null)
                                                {
                                                    floor.get_Parameter(arRoomBookNumberGUID).Set(room.Number);
                                                }
                                                if (floor.get_Parameter(arRoomBookNameGUID) != null)
                                                {
                                                    floor.get_Parameter(arRoomBookNameGUID).Set(room.get_Parameter(BuiltInParameter.ROOM_NAME).AsString());
                                                }
                                            }

                                            if (roomNumbersList.Find(elem => elem == room.Number) == null)
                                            {
                                                roomNumbersList.Add(room.Number);
                                                roomNamesList.Add(room.get_Parameter(BuiltInParameter.ROOM_NAME).AsString());
                                            }
                                        }
                                        else
                                        {
                                            XYZ pointForIntersect = null;
                                            FaceArray floorFaceArray = floorSolid.Faces;
                                            foreach (object planarFace in floorFaceArray)
                                            {
                                                if (planarFace is PlanarFace && (planarFace as PlanarFace).FaceNormal.IsAlmostEqualTo(XYZ.BasisZ))
                                                {
                                                    List<CurveLoop> curveLoopList = (planarFace as PlanarFace).GetEdgesAsCurveLoops().ToList();
                                                    if (curveLoopList.Count != 0)
                                                    {
                                                        CurveLoop curveLoop = curveLoopList.First();
                                                        if (curveLoop != null)
                                                        {
                                                            Curve c = curveLoop.First();
                                                            pointForIntersect = c.GetEndPoint(0);
                                                        }
                                                    }
                                                }
                                            }
                                            if (pointForIntersect == null) continue;
                                            Curve curve = Line.CreateBound(pointForIntersect, pointForIntersect + (500 / 304.8) * XYZ.BasisZ) as Curve;
                                            SolidCurveIntersection curveIntersection = roomSolid.IntersectWithCurve(curve, new SolidCurveIntersectionOptions());
                                            if (curveIntersection.SegmentCount > 0)
                                            {
                                                if (fillRoomBookParameters)
                                                {
                                                    if (floor.get_Parameter(arRoomBookNumberGUID) != null)
                                                    {
                                                        floor.get_Parameter(arRoomBookNumberGUID).Set(room.Number);
                                                    }
                                                    if (floor.get_Parameter(arRoomBookNameGUID) != null)
                                                    {
                                                        floor.get_Parameter(arRoomBookNameGUID).Set(room.get_Parameter(BuiltInParameter.ROOM_NAME).AsString());
                                                    }
                                                }

                                                if (roomNumbersList.Find(elem => elem == room.Number) == null)
                                                {
                                                    roomNumbersList.Add(room.Number);
                                                    roomNamesList.Add(room.get_Parameter(BuiltInParameter.ROOM_NAME).AsString());
                                                    continue;
                                                }
                                            }
                                        }
                                    }
                                }
                            }
                            roomNumbersList.OrderBy(e => PadNumbers(e));
                            roomNamesList = roomNamesList.Distinct().ToList();
                            roomNamesList.OrderBy(e => PadNumbers(e));

                            string roomNumbersByFloorType = null;
                            string roomNamesByFloorType = null;
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

                            foreach (string roomName in roomNamesList)
                            {
                                if (roomNamesByFloorType == null)
                                {
                                    roomNamesByFloorType += roomName;
                                }
                                else
                                {
                                    roomNamesByFloorType += (", " + roomName);
                                }
                            }

                            foreach (Floor floor in floorList)
                            {
                                floor.LookupParameter("АР_НомераПомещенийПоТипуПола").Set(roomNumbersByFloorType);
                            }

                            foreach (Floor floor in floorList)
                            {
                                floor.LookupParameter("АР_ИменаПомещенийПоТипуПола").Set(roomNamesByFloorType);
                            }
                        }
                    }
                    floorFinishNumeratorProgressBarWPF.Dispatcher.Invoke(() => floorFinishNumeratorProgressBarWPF.Close());
                    t.Commit();
                }
            }

            return Result.Succeeded;
        }
        private void ThreadStartingPoint()
        {
            floorFinishNumeratorProgressBarWPF = new FloorFinishNumeratorProgressBarWPF();
            floorFinishNumeratorProgressBarWPF.Show();
            System.Windows.Threading.Dispatcher.Run();
        }
        public static string PadNumbers(string input)
        {
            return Regex.Replace(input, "[0-9]+", match => match.Value.PadLeft(10, '0'));
        }
        private static void GetPluginStartInfo()
        {
            // Получаем сборку, в которой выполняется текущий код
            Assembly thisAssembly = Assembly.GetExecutingAssembly();
            string assemblyName = "FloorFinishNumerator";
            string assemblyNameRus = "Нумератор пола";
            string assemblyFolderPath = Path.GetDirectoryName(thisAssembly.Location);

            int lastBackslashIndex = assemblyFolderPath.LastIndexOf("\\");
            string dllPath = assemblyFolderPath.Substring(0, lastBackslashIndex + 1) + "PluginInfoCollector\\PluginInfoCollector.dll";

            Assembly assembly = Assembly.LoadFrom(dllPath);
            Type type = assembly.GetType("PluginInfoCollector.InfoCollector");
            var constructor = type.GetConstructor(new Type[] { typeof(string), typeof(string) });

            if (type != null)
            {
                // Создание экземпляра класса
                object instance = Activator.CreateInstance(type, new object[] { assemblyName, assemblyNameRus });
            }
        }
    }
}
