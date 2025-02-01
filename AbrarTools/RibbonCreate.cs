using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Reflection;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Media.Imaging;
using Autodesk.Revit.UI;

namespace AbrarTools
{
    internal class RibbonCreate : IExternalApplication
    {
        public Result OnShutdown(UIControlledApplication application)
        {
            return Autodesk.Revit.UI.Result.Succeeded;
        }

        public Result OnStartup(UIControlledApplication application)
        {
            var tabName = "AbrarTools";
            application.CreateRibbonTab(tabName);


            // 创建一个面板
            var panel_1 = application.CreateRibbonPanel(tabName, "FloorEve");

            //按钮1的创建：手动创建楼板四个点的标高
            var assemblyType_1 = new MarkFloorEleSingle().GetType();
            var location_1 = assemblyType_1.Assembly.Location;
            var className_1 = assemblyType_1.FullName;
            var pushButtonData_1 = new PushButtonData("SpotEvelation_1", "Manual", location_1, className_1);

            var imageSource_1 = Path.GetDirectoryName(Assembly.GetExecutingAssembly().Location) + @"\Images\SinF.png";
            //注意：和图片大小有关系，用32像素的最好，有时候显示的32像素的图片也会有问题，若按此方法无法显示图片，尝试更换图片来源
            pushButtonData_1.LargeImage = new BitmapImage(new Uri(imageSource_1));

            // 添加悬停提示
            pushButtonData_1.ToolTip = "Manually add elevation attributes to four points of the floor.\n (If the top surface has more than four points, \nselect the four points with the greatest deviation from the elevation datum.)";

            var pushButton1 = panel_1.AddItem(pushButtonData_1) as PushButton;

         


            //按钮2的创建：手动创建楼板四个点的标高
            var assemblyType_2 = new MarkFloorEleAuto().GetType();
            var location_2 = assemblyType_2.Assembly.Location;
            var className_2 = assemblyType_2.FullName;
            var pushButtonData_2 = new PushButtonData("SpotEvelation_2", "Automatic", location_2, className_2);

            var imageSource_2 = Path.GetDirectoryName(Assembly.GetExecutingAssembly().Location) + @"\Images\MulF.png";
            //注意：和图片大小有关系，用32像素的最好，有时候显示的32像素的图片也会有问题，若按此方法无法显示图片，尝试更换图片来源
            pushButtonData_2.LargeImage = new BitmapImage(new Uri(imageSource_2));

            // 添加悬停提示
            pushButtonData_2.ToolTip = "Automatically add elevation attributes to four points on all floors in the project.\n (If a slab's top surface has more than four points, \nselect the four points with the greatest deviation from the elevation datum.)";

            var pushButton2 = panel_1.AddItem(pushButtonData_2) as PushButton;


            return Autodesk.Revit.UI.Result.Succeeded;
        }
    }
}
