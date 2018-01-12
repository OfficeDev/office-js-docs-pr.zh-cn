
# <a name="task-pane-add-ins-for-project"></a>Project 任务窗格加载项


Project Standard 2013 和 Project Professional 2013 都支持任务窗格加载项。你可以运行为 Word 2013 或 Excel 2013 开发的常规任务窗格加载项。还可以开发在 Project 中处理选择事件的自定义加载项，并将项目中的任务、资源、视图和其他单元格级别的数据与 SharePoint 列表、SharePoint 外接程序、Web 部件、Web 服务和企业应用程序相集成。

 >**注释**[Project 2013 SDK 下载](https://www.microsoft.com/en-us/download/details.aspx?id=30435%20)中包括的示例加载项演示如何使用 Project 加载项对象模型以及如何在 Project Server 2013 中使用用于报告数据的 OData 服务。在提取和安装 SDK 时，请查看  `\Samples\Apps\` 子目录。

有关 Office 外接程序的简介，请参阅 [Office 加载项平台概述](../../docs/overview/office-add-ins.md)。

## <a name="add-in-scenarios-for-project"></a>用于 Project 的加载项方案


项目经理可以使用 Project 任务窗格加载项来帮助执行项目管理活动。不必离开 Project 并打开其他应用程序来搜索常用信息，项目经理可以直接在 Project 内访问信息。根据选定的任务、资源、视图或甘特图单元格中的其他数据、任务使用状况视图或资源使用状况视图，任务窗格加载项中的内容可以是上下文相关的。


 >**注释**  通过 Project Professional 2013，您可以开发访问 Project Server 2013 的本地安装、Project Online 以及本地或联机 SharePoint 2013 的任务窗格加载项。Project Standard 2013 不支持与 Project Server 数据或与 Project Server 同步的 SharePoint 任务列表的直接集成。

Project 的加载项方案包括以下几种：


-  **项目排定**???查看会影响排定的相关项目中的数据。在 Project Server 2013 中，任务窗格加载项可以集成来自其他项目的相关数据。例如，您可以查看项目和里程碑日期的部门集合，或查看基于选定的自定义字段的其他项目中的指定数据。
    
-  **资源管理**???根据指定的技能查看 Project Server 2013 中的完整资源库或子集，包括成本数据和资源可用性，以帮助选择合适的资源。
    
-  **状态和审批**???在任务窗格加载项中使用 Web 应用程序更新或查看外部企业资源规划 (ERP) 应用程序、时间表系统或会计应用程序中的数据。或者，创建在 Project Web App 和 Project Professional 2013 中均可使用的自定义状态审批 Web 部件。
    
-  **团队通信**???从任务窗格加载项中在项目上下文中直接与团队成员和资源通信。或者，当您在项目中工作时，为自己轻松维护一组与上下文相关的注释。
    
-  **工作包**???在 SharePoint 库和在线模板集合中搜索指定类型的项目模板。例如，查找用于构造项目的模板并将其添加到您的 Project 模板集合中。
    
-  **相关项目**???查看与项目计划中的特定任务相关的元数据、文档和消息。例如，可以使用 Project Professional 2013 管理从 SharePoint 任务列表导入的项目，并且仍将该任务列表与项目中的更改同步。任务窗格加载项可以显示 Project 没有为 SharePoint 列表中的任务导入的其他字段或元数据。
    
-  **使用 Project Server 对象模型**???将选定任务的 GUID 用于 Project Server Interface (PSI) 中的方法或 Project Server 的客户端对象模型 (CSOM)。例如，用于加载项的 Web 应用程序可以读取并更新选定任务和资源的状态数据，或与外部时间表应用程序集成。
    
-  **获取报告数据**???使用表示状态传输 (REST)、JavaScript 或 LINQ 查询在 Project Web App 中用于报告表的 OData 服务中查找选定任务或资源的相关信息。使用 OData 服务的查询可以通过 Project Server 2013 的在线或本地安装来执行。
    
    例如，请参阅 [创建将 REST 与本地 Project Server OData 服务结合使用的 Project 外接程序](../project/create-a-project-add-in-that-uses-rest-with-an-on-premises-odata-service.md)。
    

## <a name="developing-project-add-ins"></a>开发 Project 加载项


用于 Project 外接程序的 JavaScript 库包括  **Office** 命名空间别名的扩展，使开发人员可以访问 Project 应用程序的属性以及项目中的任务、资源和视图。Project-15.js 文件中的 JavaScript 库扩展用于用 Visual Studio 2015 创建的 Project 外接程序中。Project 2013 SDK 下载中还提供了 Office.js、Office.debug.js、Project-15.js、Project-15.debug.js 和相关文件。

若要创建外接程序，可以使用简单的文本编辑器来创建 HTML 网页和相关的 JavaScript 文件、CSS 文件以及 REST 查询。除了 HTML 页或 Web 应用程序外，外接程序还需要 XML 清单文件以用于配置。项目可以使用包含 **type** 属性指定为 **TaskPaneExtension** 的清单文件。清单文件可由多个 Office 2013 客户端应用程序使用，或者可以创建一个特定于 Project 2013 的清单文件。有关详细信息，请参阅 [Office 外接程序平台概述](../../docs/overview/office-add-ins.md) 中的_开发基础_。

对于复杂的自定义应用程序，为了便于调试，建议你使用 Visual Studio 2015 为外接程序开发网站。Visual Studio 2015 包括用于外接程序项目的模板，你可以在其中选择外接程序的类型（任务窗格、内容或邮件）和主机应用程序（Project、Word、Excel 或 Outlook）。有关与 Project Online 中的数据集成的示例，请参阅 MSDN 上“Project 编程功能”博客中的[将 Project 任务窗格外接程序连接到 PWA](http://blogs.msdn.com/b/project_programmability/archive/2012/11/02/connecting-a-project-task-pane-app-to-pwa.aspx)。

在安装 Project 2013 SDK 下载时， `\Samples\Apps\` 子目录包括以下示例外接程序：


-  **Bing 搜索：**BingSearch.XML 清单文件指向用于移动设备的 Bing 搜索页。由于 Bing Web 应用程序在 Internet 中已存在，因此 Bing 搜索加载项不使用其他源代码文件或 Project 加载项对象模型。
    
-  **Project OM 测试：**JSOM_SimpleOMCalls.XML 清单文件和 JSOM_Call.html 文件一起构成了在 Project 2013 中测试对象模型和加载项功能的示例。HTML 文件引用 JSOM_Sample.js 文件，其中包含将 Office.js 文件和 Project-15.js 文件用于主要功能的 JavaScript 函数。SDK 下载包括所有必要的源代码文件和用于 Project OM 测试加载项的清单 XML 文件。 [使用文本编辑器创建 Project 2013 的第一个任务窗格加载项](../project/create-your-first-task-pane-add-in-for-project-by-using-a-text-editor.md)中介绍了 Project OM 测试示例的开发和安装。
    
-  **HelloProject_OData：**??这是一个用于 Visual Studio 的 Visual Studio 2008 解决方案，它可以对活动项目中的数据（如成本、工作和完成百分比）进行归纳总结，并将其与存储活动项目的 Project Web App 实例中的所有已发布项目的平均值进行比较。 **创建将 REST 与本地 Project Server OData 服务结合使用的 Project 加载项**中介绍了有关将 REST 协议与 Project Web App 中的  [ProjectData](../project/create-a-project-add-in-that-uses-rest-with-an-on-premises-odata-service.md) 服务结合使用的示例的开发、安装和测试。
    

### <a name="creating-an-add-in-manifest-file"></a>创建加载项清单文件


清单文件指定加载项网页或 Web 应用程序的 URL、加载项的类型（Project 任务窗格）、用于其他语言和区域设置的内容的可选 URL 以及其他属性。


### <a name="procedure-1-to-create-the-add-in-manifest-file-for-bing-search"></a>过程 1. 创建用于 Bing 搜索的加载项清单文件


- 在本地目录中创建一个 XML 文件。该 XML 文件包括  **OfficeApp** 元素和子元素， [Office 外接程序 XML 清单](../../docs/overview/add-in-manifests.md)中对其进行了介绍。例如，创建一个名为 BingSearch.xml 的文件，其中包含以下 XML。
    
```XML
   <?xml version="1.0" encoding="utf-8"?>
 <OfficeApp xmlns="http://schemas.microsoft.com/office/appforoffice/1.0" 
            xmlns:xsi="http://www.w3.org/2001/XMLSchema-instance" 
           xsi:type="TaskPaneApp">
   <Id>1234-5678</Id>
   <Version>15.0</Version>
   <ProviderName>Microsoft</ProviderName>
   <DefaultLocale>en-us</DefaultLocale>
   <DisplayName DefaultValue="Bing Search">
   </DisplayName>
   <Description DefaultValue="Search selected data on Bing">
   </Description>
   <IconUrl DefaultValue="http://officeimg.vo.msecnd.net/_layouts/images/general/office_logo.jpg">
   </IconUrl>
   <Capabilities>
     <Capability Name="Project"/>
   </Capabilities>
   <DefaultSettings>
     <SourceLocation DefaultValue="http://m.bing.com">
     </SourceLocation>
   </DefaultSettings>
   <Permissions>ReadWriteDocument</Permissions>
 </OfficeApp>
```

- 下面是加载项清单中的必需元素：
  - 在  **OfficeApp** 元素中， `xsi:type="TaskPaneApp"` 属性指定该加载项属于任务窗格类型。
  - **Id** 元素是 UUID，并且必须唯一。
  - **Version** 元素是加载项的版本。 **ProviderName** 元素是提供加载项的公司或开发人员的名称。 **DefaultLocale** 元素指定清单中字符串的默认区域设置。
  - **DisplayName** 元素是在 Project 2013 的功能区中的“**视图**”选项卡的“**任务窗格外接程序**”下拉列表中显示的名称。该值最多可以包含 32 个字符。
  - **Description** 元素包含用于默认区域设置的加载项说明。该值最多可以包含 2000 个字符。
  - **Capabilities** 元素包含一个或多个指定主机应用程序的 **Capability** 子元素。
  - **DefaultSettings** 元素包括 **SourceLocation** 元素，后者指定 HTML 文件在文件共享中的路径或加载项使用的网页的 URL。任务窗格加载项将忽略 **RequestedHeight** 元素和 **RequestedWidth** 元素。
  - **IconUrl** 元素为可选元素。它可为文件共享中的图标或 Web 应用程序中图标的 URL。
    
- （可选）添加具有用于其他区域设置的值的  **Override** 元素。例如，以下清单为 **DisplayName**、 **Description**、 **IconUrl** 和 **SourceLocation** 的法语值提供 **Override** 元素。
    
```XML
   <?xml version="1.0" encoding="utf-8"?>
 <OfficeApp xmlns="http://schemas.microsoft.com/office/appforoffice/1.0" 
            xmlns:xsi="http://www.w3.org/2001/XMLSchema-instance" 
           xsi:type="TaskPaneApp">
   <Id>1234-5678</Id>
   <Version>15.0</Version>
   <ProviderName>Microsoft</ProviderName>
   <DefaultLocale>en-us</DefaultLocale>
   <DisplayName DefaultValue="Bing Search">
     <Override Locale="fr-fr" Value="Bing Search"/>
   </DisplayName>
   <Description DefaultValue="Search selected data on Bing">
     <Override Locale="fr-fr" Value="Search selected data on Bing"></Override>
   </Description>
   <IconUrl DefaultValue="http://officeimg.vo.msecnd.net/_layouts/images/general/office_logo.jpg">
     <Override Locale="fr-fr" Value="http://officeimg.vo.msecnd.net/_layouts/images/general/office_logo.jpg"/>
   </IconUrl>
   <Capabilities>
     <Capability Name="Project"/>
   </Capabilities>
   <DefaultSettings>
     <SourceLocation DefaultValue="http://m.bing.com">
       <Override Locale="fr-fr" Value="http://m.bing.com"/>
     </SourceLocation>
   </DefaultSettings>
   <Permissions>ReadWriteDocument</Permissions>
 </OfficeApp>
```


## <a name="installing-project-add-ins"></a>安装 Project 加载项


在 Project 2013 中，可将加载项作为独立解决方案安装在文件共享上或专用的加载项目录中。还可在公共 Office 商店中查看和购买加载项。

在文件共享中可以有多个外接程序清单 XML 文件和子目录。可以通过使用 Project 2013 中的“**信任中心**“对话框中“**受信任的外接程序目录**”选项卡来添加或删除清单目录位置和目录。若要显示 Project 中的外接程序，清单中的 **SourceLocation** 元素必须指向现有网站或 HTML 源文件。


 >**注释**  必须安装 Internet Explorer 9（或更高版本），但它不一定是默认浏览器。Office 外接程序需要 Internet Explorer 9 中的组件。默认浏览器可为 Explorer 9、Safari 5.0.6、Firefox 5、Chrome 13 或上述浏览器之一的更高版本。

在过程 2 中，在安装 Project 2013 的本地计算机上安装 Bing 搜索加载项。但是，由于加载项基础架构不直接使用本地文件路径，如  `C:\Project\AppManifests`，因此您可以在本地计算机上创建网络共享。如果您喜欢，可以在远程计算机上创建文件共享。


### <a name="procedure-2-to-install-the-bing-search-add-in"></a>过程 2. 安装 Bing 搜索加载项


1. 为加载项清单创建本地目录。例如，创建  `C:\Project\AppManifests` 目录。
    
2. 将  `C:\Project\AppManifests` 目录共享为AppManifests，这样文件共享的网络路径就变为  `\\ServerName\AppManifests`。
    
3. 将 BingSearch.XML 清单文件复制到  `C:\Project\AppManifests` 目录。
    
4. 在 Project 2013 中，打开“**Project 选项**”对话框，选择“**信任中心**”，然后选择“**信任中心设置**”。
    
5. 在“**信任中心**”对话框的左窗格中，选择“**受信任的外接程序目录**”。
    
6. 在“**受信任的外接程序目录**”窗格（请参阅图 1）中的“**目录 URL**”文本框中添加 `\\ServerName\AppManifests` 路径，选择“**添加目录**”，然后选择“**确定**”。
    
     >**注意**  图 1 显示“**受信任的目录地址**”列表中专用目录的两个文件共享和一个假设 URL。只有一个文件共享可以成为默认文件共享，并且只有一个目录 URL 可以成为默认目录。例如，如果将 `\\Server2\AppManifests` 设置为默认，Project 会清除 `\\ServerName\AppManifests` 的“**默认**”复选框。如果更改默认选择，可以选择“**清除**”来删除已安装的外接程序，然后重启 Project。如果在 Project 处于打开状态时将外接程序添加到默认文件共享或 SharePoint 目录，则应该重启 Project。

    **图 1.使用信任中心添加外接程序清单的目录**

    ![使用信任中心添加应用程序清单](../../images/pj15_AgaveOverview_TrustCenter.PNG)

7. 在“**项目**”功能区，选择“**Office 外接程序**”下拉菜单，然后选择“**查看所有**”。在“**插入外接程序**”对话框中，选择“**共享文件夹**”（见图 2）。
    
    **图 2.启动文件共享上的外接程序**

    ![启动文件共享上的 Office 应用程序](../../images/pj15_AgaveOverview_StartAgaveApp.PNG)

8. 选择必应搜索外接程序，然后选择“**插入**”。
    
Bing 搜索加载项显示在任务窗格中，如图 3 所示。可以手动调整任务窗格的大小，并使用 Bing 搜索加载项。

**图 3.使用必应搜索应用程序**

![使用 Bing 搜索应用程序](../../images/pj15_AgaveOverview_BingSearch.gif)


## <a name="distributing-project-add-ins"></a>分发 Project 加载项


可通过文件共享、SharePoint 库中的加载项目录或 Office 商店中的 Project 外接程序来分发加载项。有关详细信息，请参阅 [发布 Office 外接程序](../publish/publish.md)。


## <a name="additional-resources"></a>其他资源



- [Office 外接程序平台概述](../../docs/overview/office-add-ins.md)
    
- [Office 外接程序 XML 清单](../../docs/overview/add-in-manifests.md)
    
- [适用于 Office 的 JavaScript API](../../reference/javascript-api-for-office.md)
    
- [使用文本编辑器创建 Project 2013 的第一个任务窗格外接程序](../project/create-your-first-task-pane-add-in-for-project-by-using-a-text-editor.md)
    
- [创建将 REST 与本地 Project Server OData 服务结合使用的 Project 外接程序](../project/create-a-project-add-in-that-uses-rest-with-an-on-premises-odata-service.md)
    
- [将 Project 任务窗格外接程序连接到 PWA](http://blogs.msdn.com/b/project_programmability/archive/2012/11/02/connecting-a-project-task-pane-app-to-pwa.aspx)
    
- [Project 2013 SDK 下载](https://www.microsoft.com/en-us/download/details.aspx?id=30435%20)
    
