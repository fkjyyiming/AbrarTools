using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using Autodesk.Revit.Attributes;
using Autodesk.Revit.DB;
using Autodesk.Revit.UI;
using Autodesk.Revit.UI.Selection;

[Transaction(TransactionMode.Manual)]
public class MarkFloorEleAuto : IExternalCommand
{
    // 共享参数文件名
    private const string SharedParameterFileName = "Shared Parameters_FloorEvelationPoints.txt";

    public Result Execute(ExternalCommandData commandData, ref string message, ElementSet elements)
    {
        UIApplication uiApp = commandData.Application;
        UIDocument uiDoc = uiApp.ActiveUIDocument;
        Document doc = uiDoc.Document;

        // 载入共享参数文件（关键步骤，需要事务处理）
        if (!LoadSharedParameters(doc))
        {
            return Result.Failed;
        }

        // 获取项目中的所有楼板
        FilteredElementCollector collector = new FilteredElementCollector(doc);
        ICollection<Element> floors = collector.OfClass(typeof(Floor)).ToElements();

        // 事务处理：批量写入参数
        using (Transaction trans = new Transaction(doc, "Set Floor Elevation Parameters"))
        {
            trans.Start();

            foreach (Floor floor in floors)
            {
                // 获取楼板顶面几何顶点
                List<XYZ> slabPoints = GetSlabTopSurfacePoints(floor);
                if (slabPoints.Count < 4)
                {
                    continue; // 如果没有足够的顶点，跳过当前楼板
                }

                // 计算基准高程
                double baseElevation = GetBaseElevation(floor, doc);
                //需要的是基于0的高程，所以不需要减去基准高程
                //Dictionary<XYZ, double> elevations = slabPoints.ToDictionary(
                //    pt => pt,
                //    pt => pt.Z - baseElevation);

                Dictionary<XYZ, double> elevations = slabPoints.ToDictionary(
                    pt => pt,
                    pt => pt.Z);

                // 按绝对偏差值降序排序，取前4个点
                var selectedPoints = elevations.OrderByDescending(kv => Math.Abs(kv.Value))
                                               .Take(4)
                                               .ToDictionary(kv => kv.Key, kv => kv.Value);

                // 为每个楼板写入参数
                for (int i = 0; i < selectedPoints.Count; i++)
                {
                    string paramName = $"SpotElevation_{i + 1}";
                    double elevationValue = selectedPoints.ElementAt(i).Value;

                    // 参数有效性检查
                    Parameter param = floor.LookupParameter(paramName);
                    if (param == null)
                    {
                        TaskDialog.Show("Error",
                            $"Parameter {paramName} not found. Check shared parameters.");
                        continue;
                    }

                    // 参数赋值（注意单位处理）
                    if (param.StorageType == StorageType.Double)
                    {
                        param.Set(elevationValue);
                    }
                }
            }

            trans.Commit();
        }

        TaskDialog.Show("Success", "Elevation data written successfully for all floors");
        return Result.Succeeded;
    }

    private bool LoadSharedParameters(Document doc)
    {
        // 获取DLL路径并构建共享参数文件路径
        string dllPath = System.Reflection.Assembly.GetExecutingAssembly().Location;
        string sharedParamFile = Path.Combine(Path.GetDirectoryName(dllPath), SharedParameterFileName);

        if (!File.Exists(sharedParamFile))
        {
            TaskDialog.Show("Error", $"Shared parameter file not found: {sharedParamFile}");
            return false;
        }

        // 备份原始共享参数文件路径（重要：确保环境恢复）
        string originalSharedParamFile = doc.Application.SharedParametersFilename;
        doc.Application.SharedParametersFilename = sharedParamFile;

        // 打开参数定义文件（注意：需要有效参数组结构）
        DefinitionFile defFile = doc.Application.OpenSharedParameterFile();
        if (defFile == null)
        {
            TaskDialog.Show("Error", "Failed to open shared parameter file");
            return false;
        }

        using (Transaction trans = new Transaction(doc, "Load Shared Parameters"))
        {
            trans.Start();

            // 获取参数绑定映射表
            BindingMap map = doc.ParameterBindings;

            DefinitionGroup group = defFile.Groups.get_Item("FloorEvelation");
            if (group == null)
            {
                TaskDialog.Show("Error", "Parameter group 'FloorEvelation' not found");
                return false;
            }

            // 循环创建四个参数绑定
            for (int i = 1; i <= 4; i++)
            {
                string paramName = $"SpotElevation_{i}";
                Definition paramDef = group.Definitions.get_Item(paramName);
                if (paramDef == null)
                {
                    TaskDialog.Show("Error", $"Parameter {paramName} not found");
                    return false;
                }

                // 检查并创建参数绑定（防止重复绑定）
                if (!map.Contains(paramDef))
                {
                    CategorySet categorySet = new CategorySet();
                    categorySet.Insert(doc.Settings.Categories.get_Item(BuiltInCategory.OST_Floors));

                    InstanceBinding binding = new InstanceBinding(categorySet);
                    map.Insert(paramDef, binding, BuiltInParameterGroup.PG_IDENTITY_DATA);
                }
            }

            trans.Commit();
        }

        // 恢复原始共享参数路径（重要：避免影响其他功能）
        doc.Application.SharedParametersFilename = originalSharedParamFile;

        return true;
    }

    private List<XYZ> GetSlabTopSurfacePoints(Floor floor)
    {
        List<XYZ> topPoints = new List<XYZ>();
        IList<Reference> topFaces = HostObjectUtils.GetTopFaces(floor);

        if (topFaces.Count == 0)
        {
            return topPoints;
        }

        foreach (Reference topFaceRef in topFaces)
        {
            Face topFace = floor.GetGeometryObjectFromReference(topFaceRef) as Face;
            if (topFace == null) continue;

            EdgeArray edges = topFace.EdgeLoops.get_Item(0);
            foreach (Edge edge in edges)
            {
                foreach (XYZ pt in edge.Tessellate())
                {
                    if (!topPoints.Any(p => p.IsAlmostEqualTo(pt, 0.001)))
                    {
                        topPoints.Add(pt);
                    }
                }
            }
        }

        return topPoints.OrderByDescending(p => p.Z).Take(4).ToList();
    }

    private double GetBaseElevation(Floor floor, Document doc)
    {
        Parameter levelParam = floor.LookupParameter("Level");
        if (levelParam != null && levelParam.AsElementId() != ElementId.InvalidElementId)
        {
            Level level = doc.GetElement(levelParam.AsElementId()) as Level;
            return level?.Elevation ?? 0;
        }
        return 0;
    }
}
