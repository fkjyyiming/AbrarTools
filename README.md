# 1 目的
应项目BIM经理的需求，计划基于Revit2023，帮忙开发一系列的软件，也算是对Revit API的一些联系  
中沙友好~  
领域：建筑工程、装配式建筑、BIM、Revit二次开发  
ps. **DeepSeek 永远滴神！**

# 2 功能
## 2.1 楼板高程点属性
项目部分楼板为异形楼板，即通过修改楼板控制点、边形成的不规则楼板。因此像快捷量取各个控制点的高度是比较耗时的工作。  
![image](https://github.com/user-attachments/assets/f19c5b2a-4015-4d78-8caf-f93a44eab3c6)   
     
通过开发，将点作为第一点，剩下的点以顺时针储存，以共享参数的值的形式添加到楼板中去，以方便查看。有手动逐个添加，以及自动全部添加两种方法。  
![image](https://github.com/user-attachments/assets/dfaf47c0-56d3-430d-9886-d66611ef3483)   
     
添加结果（因为项目都为四边的楼板，对于图示复杂形式，只取偏离值最大的四个点）  
![image](https://github.com/user-attachments/assets/918a9b08-bb7e-48fb-9c3d-4a2a9348dcf2)

     

练习关键点：

- Revit共享参数的管理及添加（绑定）
- HostObjectUtils.GetTopFaces(floor)获取板的顶面，注意若为异形楼板，该List可能有多个值
- 插件可视化，Ribbon\Tab\Panel\Button的设置，及狂踩LargeImage必须为32x32像素的坑


