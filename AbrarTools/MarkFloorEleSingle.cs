using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using Autodesk.Revit.Attributes;
using Autodesk.Revit.DB;
using Autodesk.Revit.UI;
using Autodesk.Revit.UI.Selection;

[Transaction(TransactionMode.Manual)]
public class MarkFloorEleSingle : IExternalCommand
{
    // 共享参数文件名
    private const string SharedParameterFileName = "Shared Parameters_FloorEvelationPoints.txt";

    public Result Execute(ExternalCommandData commandData, ref string message, ElementSet elements)
    {
        UIApplication uiApp = commandData.Application;
        UIDocument uiDoc = uiApp.ActiveUIDocument;
        Document doc = uiDoc.Document;

        // 选择楼板（使用自定义选择过滤器）
        Reference pickedRef = uiDoc.Selection.PickObject(
            ObjectType.Element,
            new FloorSelectionFilter(),
            "Select a floor"); // 提示语改为英文
        Floor floor = doc.GetElement(pickedRef) as Floor;

        if (floor == null)
        {
            TaskDialog.Show("Error", "No floor selected"); // 错误提示改为英文
            return Result.Failed;
        }

        // 载入共享参数文件（关键步骤，需要事务处理）
        if (!LoadSharedParameters(doc))
        {
            return Result.Failed;
        }

        // 获取楼板顶面几何顶点（核心几何处理方法）
        List<XYZ> slabPoints = GetSlabTopSurfacePoints(floor);
        if (slabPoints.Count < 4)
        {
            TaskDialog.Show("Error", "Insufficient top surface points found"); // 错误提示改为英文
            return Result.Failed;
        }
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

        // 事务处理：写入参数（关键数据库操作）
        using (Transaction trans = new Transaction(doc, "Set Floor Elevation Parameters"))
        {
            trans.Start();

            for (int i = 0; i < selectedPoints.Count; i++)
            {
                string paramName = $"SpotElevation_{i + 1}";
                double elevationValue = selectedPoints.ElementAt(i).Value;

                // 参数有效性检查
                Parameter param = floor.LookupParameter(paramName);
                if (param == null)
                {
                    TaskDialog.Show("Error",
                        $"Parameter {paramName} not found. Check shared parameters."); // 错误提示改为英文
                    return Result.Failed;
                }

                // 参数赋值（注意单位处理）
                if (param.StorageType == StorageType.Double)
                {
                    param.Set(elevationValue);
                }
            }

            trans.Commit();
        }

        TaskDialog.Show("Success", "Elevation data written successfully"); // 完成提示改为英文
        return Result.Succeeded;
    }

    /// <summary>
    /// 选择过滤器实现：仅允许选择楼板元素
    /// （通过实现ISelectionFilter接口控制选择行为）
    /// </summary>
    public class FloorSelectionFilter : ISelectionFilter
    {
        public bool AllowElement(Element elem) => elem is Floor;
        public bool AllowReference(Reference reference, XYZ position) => false;
    }

    /// <summary>
    /// 加载共享参数文件并绑定到楼板类别
    /// （关键参数管理方法，包含以下步骤）：
    /// 1. 定位共享参数文件
    /// 2. 临时修改应用程序的共享参数文件路径
    /// 3. 创建参数绑定
    /// 4. 恢复原始共享参数文件路径
    /// </summary>
    /// <param name="doc">当前Revit文档</param>
    /// <returns>操作是否成功</returns>
    private bool LoadSharedParameters(Document doc)
    {
        // 获取DLL路径并构建共享参数文件路径
        string dllPath = System.Reflection.Assembly.GetExecutingAssembly().Location;
        string sharedParamFile = Path.Combine(
            Path.GetDirectoryName(dllPath),
            SharedParameterFileName);

        if (!File.Exists(sharedParamFile))
        {
            TaskDialog.Show("Error",
                $"Shared parameter file not found: {sharedParamFile}"); // 错误提示改为英文
            return false;
        }

        // 备份原始共享参数文件路径（重要：确保环境恢复）
        string originalSharedParamFile = doc.Application.SharedParametersFilename;
        doc.Application.SharedParametersFilename = sharedParamFile;

        // 打开参数定义文件（注意：需要有效参数组结构）
        DefinitionFile defFile = doc.Application.OpenSharedParameterFile();
        if (defFile == null)
        {
            TaskDialog.Show("Error", "Failed to open shared parameter file"); // 错误提示改为英文
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
                TaskDialog.Show("Error",
                    "Parameter group 'FloorEvelation' not found"); // 错误提示改为英文
                return false;
            }

            // 循环创建四个参数绑定
            for (int i = 1; i <= 4; i++)
            {
                string paramName = $"SpotElevation_{i}";
                Definition paramDef = group.Definitions.get_Item(paramName);
                if (paramDef == null)
                {
                    TaskDialog.Show("Error",
                        $"Parameter {paramName} not found"); // 错误提示改为英文
                    return false;
                }

                // 检查并创建参数绑定（防止重复绑定），绑定到楼板类别
                if (!map.Contains(paramDef))
                {
                    CategorySet categorySet = new CategorySet();
                    categorySet.Insert(doc.Settings.Categories.get_Item(
                        BuiltInCategory.OST_Floors));

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

    /// <summary>
    /// 获取楼板顶面几何顶点（核心几何处理方法）：
    /// 1. 使用HostObjectUtils获取顶面引用
    /// 2. 遍历所有顶面几何边
    /// 3. 通过Tessellate方法获取离散点
    /// 4. 去重后按Z值排序取最高4个点
    /// </summary>
    /// <param name="floor">目标楼板对象</param>
    /// <returns>顶面顶点列表（最多4个点）</returns>
    private List<XYZ> GetSlabTopSurfacePoints(Floor floor)
    {
        List<XYZ> topPoints = new List<XYZ>();
        IList<Reference> topFaces = HostObjectUtils.GetTopFaces(floor);

        if (topFaces.Count == 0)
        {
            TaskDialog.Show("Error", "No top faces found"); // 错误提示改为英文
            return topPoints;
        }

        foreach (Reference topFaceRef in topFaces)
        {
            Face topFace = floor.GetGeometryObjectFromReference(topFaceRef) as Face;
            if (topFace == null) continue;

            // 获取边环集合（通常第一个边环是外边界）
            EdgeArray edges = topFace.EdgeLoops.get_Item(0);
            foreach (Edge edge in edges)
            {
                // 离散曲线为点集（精度由Revit内部决定）
                foreach (XYZ pt in edge.Tessellate())
                {
                    // 去重处理（考虑浮点精度误差）
                    if (!topPoints.Any(p => p.IsAlmostEqualTo(pt, 0.001)))
                    {
                        topPoints.Add(pt);
                    }
                }
            }
        }

        // 按Z值降序排列并取前4个点（适用于大部分常规楼板）
        return topPoints.OrderByDescending(p => p.Z).Take(4).ToList();
    }

    /// <summary>
    /// 获取楼板关联的基准标高高程
    /// （注意：可能返回0如果标高参数无效）
    /// </summary>
    /// <param name="floor">目标楼板对象</param>
    /// <param name="doc">当前文档</param>
    /// <returns>基准高程值（单位：项目单位）</returns>
    private double GetBaseElevation(Floor floor, Document doc)
    {
        // 通过"Level"参数获取关联标高
        Parameter levelParam = floor.LookupParameter("Level");
        if (levelParam != null && levelParam.AsElementId() != ElementId.InvalidElementId)
        {
            Level level = doc.GetElement(levelParam.AsElementId()) as Level;
            return level?.Elevation ?? 0; // 安全访问操作符
        }
        return 0; // 默认返回0，可能需要错误处理
    }
}