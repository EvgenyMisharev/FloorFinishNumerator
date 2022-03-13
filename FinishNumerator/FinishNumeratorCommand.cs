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

                //List<Floor> floorList = new FilteredElementCollector(doc)
                //   .OfClass(typeof(Floor))
                //   .Cast<Floor>()
                //   .Where(f => f.FloorType.get_Parameter(BuiltInParameter.ALL_MODEL_MODEL) != null)
                //   .Where(f => f.FloorType.get_Parameter(BuiltInParameter.ALL_MODEL_MODEL).AsString() == "Пол"
                //   || f.FloorType.get_Parameter(BuiltInParameter.ALL_MODEL_MODEL).AsString() == "Полы")
                //   .ToList();

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
                        .OrderBy(f => f.Name)
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


                    }
                    finishNumeratorProgressBarWPF.Dispatcher.Invoke(() => finishNumeratorProgressBarWPF.Close());



                   

                    t.Commit();
                }
            }
            else if (finishNumberingSelectedName == "rbt_SeparatedByLevels")
            {

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
