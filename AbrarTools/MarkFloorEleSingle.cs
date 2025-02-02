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
    private const double Tolerance = 0.001;

    public Result Execute(ExternalCommandData commandData, ref string message, ElementSet elements)
    {
        UIApplication uiApp = commandData.Application;
        UIDocument uiDoc = uiApp.ActiveUIDocument;
        Document doc = uiDoc.Document;

        // 选择楼板（使用自定义选择过滤器）
        Reference pickedRef = uiDoc.Selection.PickObject(
            ObjectType.Element,
            new FloorFilter(),
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

        // 获取顶面点并处理排序
        List<XYZ> slabPoints = GetTopSurfacePoints(floor);
        if (slabPoints.Count < 4)
        {
            TaskDialog.Show("Error", "Insufficient top surface points found"); // 错误提示改为英文
            return Result.Failed;
        }

        // 处理顶点排序
        List<XYZ> sortedPoints = slabPoints.Count == 4 ?
            SortQuadPoints(slabPoints) :
            slabPoints.OrderByDescending(p => p.Z).Take(4).ToList();

        // 事务处理：写入参数（关键数据库操作）
        using (Transaction trans = new Transaction(doc, "Set Floor Elevation Parameters"))
        {
            trans.Start();

            for (int pointIndex = 0; pointIndex < 4; pointIndex++)
            {
                if (pointIndex >= sortedPoints.Count) break;

                XYZ point = sortedPoints[pointIndex];
                SetParameters(floor, pointIndex + 1, point);

            }

            trans.Commit();
        }

        TaskDialog.Show("Success", "Parameters updated successfully");
        return Result.Succeeded;
    }

    #region 核心算法
    /// <summary>
    /// 四点顺时针排序算法（基于顶面边和最高点位置）
    /// </summary>
    private List<XYZ> SortQuadPoints(List<XYZ> points)
    {
        // 1. 获取顶面并按边排序
        List<XYZ> sortedPoints = SortPointsByEdges(points);

        // 2. 找到最高点及其索引
        XYZ highestPoint = sortedPoints.OrderByDescending(p => p.Z).First();
        int highestPointIndex = sortedPoints.IndexOf(highestPoint);

        // 3. 将最高点之前的点移动到列表末尾
        if (highestPointIndex > 0)
        {
            List<XYZ> pointsBeforeHighest = sortedPoints.GetRange(0, highestPointIndex);
            sortedPoints.RemoveRange(0, highestPointIndex);
            sortedPoints.AddRange(pointsBeforeHighest);
        }

        // 4. 分割列表为最高点和剩余点
        XYZ pH = sortedPoints[0];
        List<XYZ> list1 = sortedPoints.Skip(1).ToList();

        // 5. 判断顺时针方向并调整剩余点顺序
        XYZ p1 = list1[0]; // list1 的第一个点
        XYZ p3 = list1[2]; // list1 的最后一个点

        // 将点投影到XY平面
        XYZ projected_pH = new XYZ(pH.X, pH.Y, 0);
        XYZ projected_p1 = new XYZ(p1.X, p1.Y, 0);
        XYZ projected_p3 = new XYZ(p3.X, p3.Y, 0);

        // 计算向量
        XYZ v1 = (projected_p1 - projected_pH).Normalize();
        XYZ v3 = (projected_p3 - projected_pH).Normalize();

        // 判断 v1 到 v3 是否为顺时针方向
        if (v1.CrossProduct(v3).Z > 0)
        {
            list1.Reverse(); // 逆转 list1
        }

        // 6. 组合结果
        return new List<XYZ> { pH }.Concat(list1).ToList();
    }

    /// <summary>
    /// 根据顶面的边对点进行排序
    /// </summary>
    private List<XYZ> SortPointsByEdges(List<XYZ> points)
    {
        // 假设 points 列表中的点已经是从 GetTopSurfacePoints 函数获取的顶面点

        // 用于存储排序后的点
        List<XYZ> sortedPoints = new List<XYZ>();

        // 用于存储尚未排序的点
        List<XYZ> unsortedPoints = new List<XYZ>(points);

        // 随机选择一个起始点，并将其添加到 sortedPoints 列表中
        XYZ currentPoint = unsortedPoints[0];
        sortedPoints.Add(currentPoint);
        unsortedPoints.RemoveAt(0);

        // 循环直到所有点都被排序
        while (unsortedPoints.Count > 0)
        {
            // 找到与当前点最近的点
            XYZ nextPoint = unsortedPoints.OrderBy(p => p.DistanceTo(currentPoint)).First();

            // 将找到的点添加到 sortedPoints 列表中，并从 unsortedPoints 列表中移除
            sortedPoints.Add(nextPoint);
            unsortedPoints.Remove(nextPoint);

            // 更新当前点
            currentPoint = nextPoint;
        }

        // 检查第一个点和最后一个点是否闭合，如果不闭合，则反转列表
        if (sortedPoints.First().DistanceTo(sortedPoints.Last()) > Tolerance)
        {
            sortedPoints.Reverse();
        }

        return sortedPoints;
    }
    #endregion

    #region 几何处理
    /// <summary>
    /// 获取楼板顶面顶点
    /// </summary>
    private List<XYZ> GetTopSurfacePoints(Floor floor)
    {
        List<XYZ> points = new List<XYZ>();
        foreach (Reference faceRef in HostObjectUtils.GetTopFaces(floor))
        {
            Face face = floor.GetGeometryObjectFromReference(faceRef) as Face;
            if (face == null) continue;

            foreach (EdgeArray loop in face.EdgeLoops)
            {
                foreach (Edge edge in loop)
                {
                    foreach (XYZ pt in edge.Tessellate())
                    {
                        if (!points.Any(p => p.IsAlmostEqualTo(pt, Tolerance)))
                        {
                            points.Add(pt);
                        }
                    }
                }
            }
        }
        return points.OrderByDescending(p => p.Z).ToList();
    }
    #endregion

    #region 参数管理
    /// <summary>
    /// 设置参数值
    /// </summary>
    private void SetParameters(Floor floor, int index, XYZ point)
    {
        // Z值参数（保留原始名称）
        SetParameterValue(floor, $"SpotElevation_{index}", point.Z);

        // X/Y值参数（新命名规则）
        SetParameterValue(floor, $"SpotCoordinate_N{index}", point.Y);
        SetParameterValue(floor, $"SpotCoordinate_E{index}", point.X);
    }

    private void SetParameterValue(Element elem, string paramName, double value)
    {
        Parameter param = elem.LookupParameter(paramName);
        if (param != null && param.StorageType == StorageType.Double)
        {
            param.Set(value);
        }
    }

    /// <summary>
    /// 加载共享参数文件
    /// </summary>
    private bool LoadSharedParameters(Document doc)
    {
        try
        {
            string assemblyPath = System.Reflection.Assembly.GetExecutingAssembly().Location;
            string paramFile = Path.Combine(Path.GetDirectoryName(assemblyPath), SharedParameterFileName);

            if (!File.Exists(paramFile))
            {
                TaskDialog.Show("Error", $"Shared parameter file missing: {paramFile}");
                return false;
            }

            // 备份原始参数文件路径
            string originalParamFile = doc.Application.SharedParametersFilename;
            doc.Application.SharedParametersFilename = paramFile;

            using (Transaction t = new Transaction(doc, "Load Parameters"))
            {
                t.Start();

                BindingMap bindings = doc.ParameterBindings;
                DefinitionFile file = doc.Application.OpenSharedParameterFile();

                DefinitionGroup group = file.Groups.get_Item("FloorEvelation")
                    ?? file.Groups.Create("FloorEvelation");

                // 创建所有需要的参数
                for (int i = 1; i <= 4; i++)
                {
                    CreateLengthParameter(group, bindings, doc, $"SpotElevation_{i}");

                }
                for (int i = 1; i <= 4; i++)
                {
                    CreateLengthParameter(group, bindings, doc, $"SpotCoordinate_N{i}");
                    CreateLengthParameter(group, bindings, doc, $"SpotCoordinate_E{i}");

                }

                t.Commit();
            }

            // 恢复原始参数文件
            doc.Application.SharedParametersFilename = originalParamFile;
            return true;
        }
        catch (Exception ex)
        {
            TaskDialog.Show("Parameter Error", $"Failed to load parameters: {ex.Message}");
            return false;
        }
    }

    private void CreateLengthParameter(DefinitionGroup group, BindingMap bindings, Document doc, string name)
    {
        ExternalDefinitionCreationOptions opt = new ExternalDefinitionCreationOptions(name, SpecTypeId.Length);

        Definition paramDef = group.Definitions.get_Item(name)
            ?? group.Definitions.Create(opt);

        if (!bindings.Contains(paramDef))
        {
            CategorySet categories = new CategorySet();
            categories.Insert(doc.Settings.Categories.get_Item(BuiltInCategory.OST_Floors));
            bindings.Insert(paramDef, new InstanceBinding(categories), BuiltInParameterGroup.PG_IDENTITY_DATA);
        }
    }
    #endregion

    #region 辅助类
    /// <summary>
    /// 楼板选择过滤器
    /// </summary>
    private class FloorFilter : ISelectionFilter
    {
        public bool AllowElement(Element elem) => elem is Floor;
        public bool AllowReference(Reference reference, XYZ position) => false;
    }
    #endregion

}