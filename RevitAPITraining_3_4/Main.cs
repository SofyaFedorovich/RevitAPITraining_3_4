using Autodesk.Revit.ApplicationServices;
using Autodesk.Revit.Attributes;
using Autodesk.Revit.DB;
using Autodesk.Revit.DB.Plumbing;
using Autodesk.Revit.UI;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace RevitAPITraining_3_4
{
    [Transaction(TransactionMode.Manual)]
    public class Main : IExternalCommand
    {
        public Result Execute(ExternalCommandData commandData, ref string message, ElementSet elements)
        {
            UIApplication uiapp = commandData.Application;
            UIDocument uidoc = uiapp.ActiveUIDocument;
            Document doc = uidoc.Document;

            var categorySet = new CategorySet();
            categorySet.Insert(Category.GetCategory(doc, BuiltInCategory.OST_PipeCurves));

            using (Transaction ts = new Transaction(doc, "Добавить параметр"))
            {
                ts.Start();
                CreateSharedParameter(uiapp.Application, doc, "Наименование",
                    categorySet, BuiltInParameterGroup.PG_TEXT, true); //Требуется создать общий параметр "Наименование"
                ts.Commit();
            }

            List<Pipe> pipes = new FilteredElementCollector(doc)
                .OfClass(typeof(Pipe))
                .Cast<Pipe>()
                .ToList();

            using (Transaction ts = new Transaction(doc, "Записать параметр"))
            {
                ts.Start();
                foreach (var pipe in pipes)
                {
                    double diamOut = pipe.get_Parameter(BuiltInParameter.RBS_PIPE_OUTER_DIAMETER).AsDouble();
                    double diamIn = pipe.get_Parameter(BuiltInParameter.RBS_PIPE_INNER_DIAM_PARAM).AsDouble();
                    double diamOut1 = UnitUtils.ConvertFromInternalUnits(diamOut, UnitTypeId.Millimeters);
                    double diamIn1 = UnitUtils.ConvertFromInternalUnits(diamIn, UnitTypeId.Millimeters);
                    Parameter param = pipe.LookupParameter("Наименование");
                    param.Set(String.Format("Труба ø{0}/ø{1}",  diamOut1, diamIn1));
                }
                ts.Commit();
            }
            TaskDialog.Show("Выполнено", "Параметр создан");


            return Result.Succeeded;
        }
        private void CreateSharedParameter(Application application,
            Document doc, string parameterName, CategorySet categorySet,
            BuiltInParameterGroup builthInParameterGroup, bool isInstance)
        {
            DefinitionFile definitionFile = application.OpenSharedParameterFile();
            if (definitionFile == null)
            {
                TaskDialog.Show("Ошибка", "Не найден файл общих параметров");
                return;
            }

            Definition definition = definitionFile.Groups
                .SelectMany(group => group.Definitions)
                .FirstOrDefault(def => def.Name.Equals(parameterName));
            if (definition == null)
            {
                TaskDialog.Show("Ошибка", "Не найден указанный параметр");
                return;
            }

            Binding binding = application.Create.NewTypeBinding(categorySet);
            if (isInstance)
                binding = application.Create.NewInstanceBinding(categorySet);

            BindingMap map = doc.ParameterBindings;
            map.Insert(definition, binding, builthInParameterGroup);
        }
    }
}
