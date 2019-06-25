---
title: Project 任务窗格加载项
description: ''
ms.date: 06/20/2019
localization_priority: Priority
ms.openlocfilehash: c04b2d3d2da3f258751585e0d30a9c726af373e6
ms.sourcegitcommit: 382e2735a1295da914f2bfc38883e518070cec61
ms.translationtype: HT
ms.contentlocale: zh-CN
ms.lasthandoff: 06/21/2019
ms.locfileid: "35128599"
---
# <a name="task-pane-add-ins-for-project"></a><span data-ttu-id="88c83-102">Project 任务窗格加载项</span><span class="sxs-lookup"><span data-stu-id="88c83-102">Task pane add-ins for Project</span></span>

<span data-ttu-id="88c83-103">Project Standard 2013 和 Project Professional 2013（版本 15.1 或更高版本）均包括对任务窗格加载项的支持。你可以运行专为 Word 或 Excel 开发的通用任务窗格加载项。</span><span class="sxs-lookup"><span data-stu-id="88c83-103">Project Standard 2013 and Project Professional 2013 (version 15.1 or higher) both include support for task pane add-ins. You can run general task pane add-ins that are developed for Word or Excel.</span></span> <span data-ttu-id="88c83-104">此外，你还可以开发自定义加载项，用于处理 Project 中的选择事件以及将项目中的任务、资源、视图和其他单元格级别的数据与 SharePoint 列表、SharePoint 加载项、Web 部件、Web 服务和企业应用程序相集成。</span><span class="sxs-lookup"><span data-stu-id="88c83-104">You can also develop custom add-ins that handle selection events in Project and integrate task, resource, view, and other cell-level data in a project with SharePoint lists, SharePoint Add-ins, Web Parts, web services, and enterprise applications.</span></span>

> [!NOTE]
> <span data-ttu-id="88c83-p102">[Project 2013 SDK 下载](https://www.microsoft.com/download/details.aspx?id=30435%20)中的示例加载项展示了如何使用 Project 加载项对象模型，以及如何在 Project Server 2013 中使用 OData 报表数据服务。提取和安装 SDK 时，请查看 `\Samples\Apps\` 子目录。</span><span class="sxs-lookup"><span data-stu-id="88c83-p102">The [Project 2013 SDK download](https://www.microsoft.com/download/details.aspx?id=30435%20) includes sample add-ins that show how to use the add-in object model for Project, and how to use the OData service for reporting data in Project Server 2013. When you extract and install the SDK, see the `\Samples\Apps\` subdirectory.</span></span>

<span data-ttu-id="88c83-107">有关 Office 加载项的简介，请参阅 [Office 加载项平台概述](../overview/office-add-ins.md)。</span><span class="sxs-lookup"><span data-stu-id="88c83-107">For an introduction to Office Add-ins, see [Office Add-ins platform overview](../overview/office-add-ins.md).</span></span>

## <a name="add-in-scenarios-for-project"></a><span data-ttu-id="88c83-108">用于 Project 的加载项方案</span><span class="sxs-lookup"><span data-stu-id="88c83-108">Add-in scenarios for Project</span></span>

<span data-ttu-id="88c83-p103">项目经理可以使用 Project 任务窗格加载项来帮助执行项目管理活动。不必离开 Project 并打开其他应用程序来搜索常用信息，项目经理可以直接在 Project 内访问信息。根据选定的任务、资源、视图或甘特图单元格中的其他数据、任务使用状况视图或资源使用状况视图，任务窗格加载项中的内容可以是上下文相关的。</span><span class="sxs-lookup"><span data-stu-id="88c83-p103">Project managers can use Project task pane add-ins to help with project management activities. Instead of leaving Project and opening another application to search for frequently used information, project managers can directly access the information within Project. The content in a task pane add-in can be context-sensitive, based on the selected task, resource, view, or other data in a cell in a Gantt chart, task usage view, or resource usage view.</span></span>

> [!NOTE]
> <span data-ttu-id="88c83-112">使用 Project Professional 2013，可以开发访问 Project 网页版、Project Server 2013 本地安装以及本地或联机 SharePoint 2013 的任务窗格加载项。</span><span class="sxs-lookup"><span data-stu-id="88c83-112">With Project Professional 2013, you can develop task pane add-ins that access on-premises installations of Project Server 2013, Project Online, and on-premises or online SharePoint 2013.Project Standard 2013 does not support direct integration with Project Server data or SharePoint task lists that are synchronized with Project Server.</span></span> <span data-ttu-id="88c83-113">Project Standard 2013 不支持与 Project Server 数据或与 Project Server 同步的 SharePoint 任务列表直接集成。</span><span class="sxs-lookup"><span data-stu-id="88c83-113">pj15stdshort does not support direct integration with Project Server data or SharePoint task lists that are synchronized with Project Server.</span></span>

<span data-ttu-id="88c83-114">Project 加载项方案包括以下几种：</span><span class="sxs-lookup"><span data-stu-id="88c83-114">Add-in scenarios for Project include the following:</span></span>

- <span data-ttu-id="88c83-p105">**项目排定**???查看会影响排定的相关项目中的数据。在 Project Server 2013 中，任务窗格加载项可以集成来自其他项目的相关数据。例如，您可以查看项目和里程碑日期的部门集合，或查看基于选定的自定义字段的其他项目中的指定数据。</span><span class="sxs-lookup"><span data-stu-id="88c83-p105">**Project scheduling** View data from related projects that can affect scheduling. A task pane add-in can integrate relevant data from other projects in Project Server 2013. For example, you can view the departmental collection of projects and milestone dates, or view specified data from other projects that are based on a selected custom field.</span></span>

- <span data-ttu-id="88c83-118">**资源管理**???根据指定的技能查看 Project Server 2013 中的完整资源库或子集，包括成本数据和资源可用性，以帮助选择合适的资源。</span><span class="sxs-lookup"><span data-stu-id="88c83-118">**Resource management** View the complete resource pool in Project Server 2013 or a subset based on specified skills, including cost data and resource availability, to help select appropriate resources.</span></span>

- <span data-ttu-id="88c83-p106">**状态和审批**???在任务窗格加载项中使用 Web 应用程序更新或查看外部企业资源规划 (ERP) 应用程序、时间表系统或会计应用程序中的数据。或者，创建在 Project Web App 和 Project Professional 2013 中均可使用的自定义状态审批 Web 部件。</span><span class="sxs-lookup"><span data-stu-id="88c83-p106">**Statusing and approvals** Use a web application in a task pane add-in to update or view data from an external enterprise resource planning (ERP) application, timesheet system, or accounting application. Or, create a custom status approval Web Part that can be used within both Project Web App and Project Professional 2013.</span></span>

- <span data-ttu-id="88c83-p107">**团队通信**???从任务窗格加载项中在项目上下文中直接与团队成员和资源通信。或者，当您在项目中工作时，为自己轻松维护一组与上下文相关的注释。</span><span class="sxs-lookup"><span data-stu-id="88c83-p107">**Team communication** Communicate with team members and resources directly from a task pane add-in, within the context of a project. Or, easily maintain a set of context-sensitive notes for yourself as you work in a project.</span></span>

- <span data-ttu-id="88c83-p108">**工作包**???在 SharePoint 库和在线模板集合中搜索指定类型的项目模板。例如，查找用于构造项目的模板并将其添加到您的 Project 模板集合中。</span><span class="sxs-lookup"><span data-stu-id="88c83-p108">**Work packages** Search for specified kinds of project templates within SharePoint libraries and online template collections. For example, find templates for construction projects and add them to your Project template collection.</span></span>

- <span data-ttu-id="88c83-p109">**相关项目**???查看与项目计划中的特定任务相关的元数据、文档和消息。例如，可以使用 Project Professional 2013 管理从 SharePoint 任务列表导入的项目，并且仍将该任务列表与项目中的更改同步。任务窗格加载项可以显示 Project 没有为 SharePoint 列表中的任务导入的其他字段或元数据。</span><span class="sxs-lookup"><span data-stu-id="88c83-p109">**Related items** View metadata, documents, and messages that are related to specific tasks in a project plan. For example, you can use Project Professional 2013 to manage a project that was imported from a SharePoint task list, and still synchronize the task list with changes in the project. A task pane add-in can show additional fields or metadata that Project did not import for tasks in the SharePoint list.</span></span>

- <span data-ttu-id="88c83-p110">**使用 Project Server 对象模型**???将选定任务的 GUID 用于 Project Server Interface (PSI) 中的方法或 Project Server 的客户端对象模型 (CSOM)。例如，用于加载项的 Web 应用程序可以读取并更新选定任务和资源的状态数据，或与外部时间表应用程序集成。</span><span class="sxs-lookup"><span data-stu-id="88c83-p110">**Use the Project Server object models** Use the GUID of a selected task with methods in the Project Server Interface (PSI) or the client-side object model (CSOM) of Project Server. For example, the web application for an add-in can read and update the statusing data of a selected task and resource, or integrate with an external timesheet application.</span></span>

- <span data-ttu-id="88c83-p111">**获取报告数据**???使用表示状态传输 (REST)、JavaScript 或 LINQ 查询在 Project Web App 中用于报告表的 OData 服务中查找选定任务或资源的相关信息。使用 OData 服务的查询可以通过 Project Server 2013 的在线或本地安装来执行。</span><span class="sxs-lookup"><span data-stu-id="88c83-p111">**Get reporting data** Use Representational State Transfer (REST), JavaScript, or LINQ queries to find related information for a selected task or resource in the OData service for reporting tables in Project Web App. Queries that use the OData service can be done with an online or an on-premises installation of Project Server 2013.</span></span>

    <span data-ttu-id="88c83-132">例如，请参阅 [创建将 REST 与本地 Project Server OData 服务结合使用的 Project 外接程序](../project/create-a-project-add-in-that-uses-rest-with-an-on-premises-odata-service.md)。</span><span class="sxs-lookup"><span data-stu-id="88c83-132">For example, see [Create a Project add-in that uses REST with an on-premises Project Server OData  service](../project/create-a-project-add-in-that-uses-rest-with-an-on-premises-odata-service.md).</span></span>

## <a name="developing-project-add-ins"></a><span data-ttu-id="88c83-133">开发 Project 加载项</span><span class="sxs-lookup"><span data-stu-id="88c83-133">Developing Project add-ins</span></span>

<span data-ttu-id="88c83-p112">用于 Project 外接程序的 JavaScript 库包括  **Office** 命名空间别名的扩展，使开发人员可以访问 Project 应用程序的属性以及项目中的任务、资源和视图。Project-15.js 文件中的 JavaScript 库扩展用于用 Visual Studio 2015 创建的 Project 外接程序中。Project 2013 SDK 下载中还提供了 Office.js、Office.debug.js、Project-15.js、Project-15.debug.js 和相关文件。</span><span class="sxs-lookup"><span data-stu-id="88c83-p112">The JavaScript library for Project add-ins includes extensions of the  **Office** namespace alias that enable developers to access properties of the Project application and tasks, resources, and views in a project. The JavaScript library extensions in the Project-15.js file are used in a Project add-in created with Visual Studio 2015. The Office.js, Office.debug.js, Project-15.js, Project-15.debug.js, and related files are also provided in the Project 2013 SDK download.</span></span>

<span data-ttu-id="88c83-p113">若要创建加载项，可以使用简单文本编辑器创建 HTML 网页和相关的 JavaScript 文件、CSS 文件及 REST 查询。除了 HTML 页面或 Web 应用程序外，加载项还需要一个用于配置的 XML 清单文件。Project 可以使用包括指定为  **TaskPaneExtension** 的 **type** 属性的清单文件。该清单文件可供多个 Office 2013 客户端应用程序使用，您也可以创建一个 Project 2013 专用的清单文件。有关详细信息，请参阅 [Office 加载项平台概述](../overview/office-add-ins.md)中的"开发基础知识"一节。</span><span class="sxs-lookup"><span data-stu-id="88c83-p113">To create an add-in, you can use a simple text editor to create an HTML webpage and related JavaScript files, CSS files, and REST queries. In addition to an HTML page or a web application, an add-in requires an XML manifest file for configuration. Project can use a manifest file that includes a  **type** attribute that is specified as **TaskPaneExtension**. The manifest file can be used by multiple Office 2013 client applications, or you can create a manifest file that is specific for Project 2013. For more information, see the  _Development basics_ section in [Office Add-ins platform overview](../overview/office-add-ins.md).</span></span>

<span data-ttu-id="88c83-142">在安装 Project 2013 SDK 下载时，`\Samples\Apps\` 子目录包括以下示例加载项：</span><span class="sxs-lookup"><span data-stu-id="88c83-142">When you install the Project 2013 SDK download, the  `\Samples\Apps\` subdirectory includes the following sample add-ins:</span></span>

- <span data-ttu-id="88c83-p114">**Bing 搜索：** BingSearch.XML 清单文件指向用于移动设备的 Bing 搜索页。由于 Bing Web 应用程序在 Internet 中已存在，因此 Bing 搜索加载项不使用其他源代码文件或 Project 加载项对象模型。</span><span class="sxs-lookup"><span data-stu-id="88c83-p114">**Bing Search:** The BingSearch.xml manifest file points to the Bing search page for mobile devices. Because the Bing web app already exists on the Internet, the Bing Search add-in does not use other source code files or the add-in object model for Project.</span></span>

- <span data-ttu-id="88c83-p115">**Project OM 测试：** JSOM_SimpleOMCalls.XML 清单文件和 JSOM_Call.html 文件一起构成了在 Project 2013 中测试对象模型和加载项功能的示例。HTML 文件引用 JSOM_Sample.js 文件，其中包含将 Office.js 文件和 Project-15.js 文件用于主要功能的 JavaScript 函数。SDK 下载包括所有必要的源代码文件和用于 Project OM 测试加载项的清单 XML 文件。 [使用文本编辑器创建 Project 2013 的第一个任务窗格加载项](../project/create-your-first-task-pane-add-in-for-project-by-using-a-text-editor.md)中介绍了 Project OM 测试示例的开发和安装。</span><span class="sxs-lookup"><span data-stu-id="88c83-p115">**Project OM Test:** The JSOM_SimpleOMCalls.xml manifest file and the JSOM_Call.html file are, together, an example that tests the object model and add-in functionality in Project 2013. The HTML file references the JSOM_Sample.js file, which has JavaScript functions that use the Office.js file and the Project-15.js file for the primary functionality. The SDK download includes all of the necessary source code files and the manifest XML file for the Project OM Test add-in. The development and installation of the Project OM Test sample is described in [Create your first task pane add-in for Project 2013 by using a text editor](../project/create-your-first-task-pane-add-in-for-project-by-using-a-text-editor.md).</span></span>

- <span data-ttu-id="88c83-p116">**HelloProject_OData：**??这是一个用于 Visual Studio 的 Visual Studio 2008 解决方案，它可以对活动项目中的数据（如成本、工作和完成百分比）进行归纳总结，并将其与存储活动项目的 Project Web App 实例中的所有已发布项目的平均值进行比较。 **创建将 REST 与本地 Project Server OData 服务结合使用的 Project 加载项**中介绍了有关将 REST 协议与 Project Web App 中的  [ProjectData](../project/create-a-project-add-in-that-uses-rest-with-an-on-premises-odata-service.md) 服务结合使用的示例的开发、安装和测试。</span><span class="sxs-lookup"><span data-stu-id="88c83-p116">**HelloProject_OData:** This is a Visual Studio solution for Project Professional 2013 that summarizes data from the active project, such as cost, work, and percent complete, and compares that with the average for all published projects in the Project Web App instance where the active project is stored. The development, installation, and testing of the sample, which uses the REST protocol with the **ProjectData** service in Project Web App, is described in [Create a Project add-in that uses REST with an on-premises Project Server OData service](../project/create-a-project-add-in-that-uses-rest-with-an-on-premises-odata-service.md).</span></span>

### <a name="creating-an-add-in-manifest-file"></a><span data-ttu-id="88c83-151">创建加载项清单文件</span><span class="sxs-lookup"><span data-stu-id="88c83-151">Creating an add-in manifest file</span></span>

<span data-ttu-id="88c83-152">清单文件指定加载项网页或 Web 应用程序的 URL、加载项的类型（Project 任务窗格）、用于其他语言和区域设置的内容的可选 URL 以及其他属性。</span><span class="sxs-lookup"><span data-stu-id="88c83-152">The manifest file specifies the URL of the add-in webpage or web application, the kind of add-in (task pane for Project), optional URLs of content for other languages and locales, and other properties.</span></span>

### <a name="procedure-1-to-create-the-add-in-manifest-file-for-bing-search"></a><span data-ttu-id="88c83-p117">过程 1. 创建用于 Bing 搜索的加载项清单文件</span><span class="sxs-lookup"><span data-stu-id="88c83-p117">Procedure 1. To create the add-in manifest file for Bing Search</span></span>

- <span data-ttu-id="88c83-p118">在本地目录中创建一个 XML 文件。该 XML 文件包括  **OfficeApp** 元素和子元素， [Office 外接程序 XML 清单](../develop/add-in-manifests.md)中对其进行了介绍。例如，创建一个名为 BingSearch.xml 的文件，其中包含以下 XML。</span><span class="sxs-lookup"><span data-stu-id="88c83-p118">Create an XML file in a local directory. The XML file includes the  **OfficeApp** element and child elements, which are described in the [Office Add-ins XML manifest](../develop/add-in-manifests.md). For example, create a file named BingSearch.xml that contains the following XML.</span></span>

    ```XML
    <?xml version="1.0" encoding="utf-8"?>
    <OfficeApp xmlns="http://schemas.microsoft.com/office/appforoffice/1.0"
                xmlns:xsi="https://www.w3.org/2001/XMLSchema-instance"
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

- <span data-ttu-id="88c83-158">下面是加载项清单中的必需元素：</span><span class="sxs-lookup"><span data-stu-id="88c83-158">Following are the required elements in the add-in manifest:</span></span>
  - <span data-ttu-id="88c83-159">在  **OfficeApp** 元素中， `xsi:type="TaskPaneApp"` 属性指定该加载项属于任务窗格类型。</span><span class="sxs-lookup"><span data-stu-id="88c83-159">In the  **OfficeApp** element, the `xsi:type="TaskPaneApp"` attribute specifies that the add-in is a task pane type.</span></span>
  - <span data-ttu-id="88c83-160">**Id** 元素是 UUID，并且必须唯一。</span><span class="sxs-lookup"><span data-stu-id="88c83-160">The  **Id** element is a UUID and must be unique.</span></span>
  - <span data-ttu-id="88c83-p119">**Version** 元素是加载项的版本。 **ProviderName** 元素是提供加载项的公司或开发人员的名称。 **DefaultLocale** 元素指定清单中字符串的默认区域设置。</span><span class="sxs-lookup"><span data-stu-id="88c83-p119">The  **Version** element is the version of the add-in. The **ProviderName** element is the name of the company or developer who provides the add-in. The **DefaultLocale** element specifies the default locale for the strings in the manifest.</span></span>
  - <span data-ttu-id="88c83-p120">**DisplayName** 元素是在 Project 2013 的功能区中的“**视图**”选项卡的“**任务窗格外接程序**”下拉列表中显示的名称。该值最多可以包含 32 个字符。</span><span class="sxs-lookup"><span data-stu-id="88c83-p120">The  **DisplayName** element is the name that shows in the **Task Pane Add-in** drop-down list in the **VIEW** tab of the ribbon in Project 2013. The value can contain up to 32 characters.</span></span>
  - <span data-ttu-id="88c83-p121">**Description** 元素包含用于默认区域设置的加载项说明。该值最多可以包含 2000 个字符。</span><span class="sxs-lookup"><span data-stu-id="88c83-p121">The  **Description** element contains the add-in description for the default locale. The value can contain up to 2000 characters.</span></span>
  - <span data-ttu-id="88c83-168">**Capabilities** 元素包含一个或多个指定主机应用程序的 **Capability** 子元素。</span><span class="sxs-lookup"><span data-stu-id="88c83-168">The  **Capabilities** element contains one or more **Capability** child elements that specify the host application.</span></span>
  - <span data-ttu-id="88c83-p122">**DefaultSettings** 元素包括 **SourceLocation** 元素，后者指定 HTML 文件在文件共享中的路径或加载项使用的网页的 URL。任务窗格加载项将忽略 **RequestedHeight** 元素和 **RequestedWidth** 元素。</span><span class="sxs-lookup"><span data-stu-id="88c83-p122">The  **DefaultSettings** element includes the **SourceLocation** element, which specifies the path of an HTML file on a file share or the URL of a webpage that the add-in uses. A task pane add-in ignores the **RequestedHeight** element and the **RequestedWidth** element.</span></span>
  - <span data-ttu-id="88c83-p123">**IconUrl** 元素为可选元素。它可为文件共享中的图标或 Web 应用程序中图标的 URL。</span><span class="sxs-lookup"><span data-stu-id="88c83-p123">The  **IconUrl** element is optional. It can be an icon on a file share or the URL of an icon in a web application.</span></span>

- <span data-ttu-id="88c83-p124">（可选）添加具有用于其他区域设置的值的  **Override** 元素。例如，以下清单为 **DisplayName**、 **Description**、 **IconUrl** 和 **SourceLocation** 的法语值提供 **Override** 元素。</span><span class="sxs-lookup"><span data-stu-id="88c83-p124">(Optional) Add  **Override** elements that have values for other locales. For example, the following manifest provides **Override** elements for French values of **DisplayName**,  **Description**,  **IconUrl**, and  **SourceLocation**.</span></span>

    ```XML
    <?xml version="1.0" encoding="utf-8"?>
    <OfficeApp xmlns="http://schemas.microsoft.com/office/appforoffice/1.0"
                xmlns:xsi="https://www.w3.org/2001/XMLSchema-instance"
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

## <a name="installing-project-add-ins"></a><span data-ttu-id="88c83-175">安装 Project 加载项</span><span class="sxs-lookup"><span data-stu-id="88c83-175">Installing Project add-ins</span></span>

<span data-ttu-id="88c83-p125">在 Project 2013 中，可以将加载项安装为文件共享或专用加载项目录中的独立解决方案。还可以在 AppSource 中评论和购买加载项。</span><span class="sxs-lookup"><span data-stu-id="88c83-p125">In Project 2013, you can install add-ins as stand-alone solutions on a file share, or in a private add-in catalog. You can also review and purchase add-ins in AppSource.</span></span>

<span data-ttu-id="88c83-p126">在文件共享中可以有多个外接程序清单 XML 文件和子目录。可以通过使用 Project 2013 中的“**信任中心**“对话框中“**受信任的外接程序目录**”选项卡来添加或删除清单目录位置和目录。若要显示 Project 中的外接程序，清单中的 **SourceLocation** 元素必须指向现有网站或 HTML 源文件。</span><span class="sxs-lookup"><span data-stu-id="88c83-p126">There can be multiple add-in manifest XML files and subdirectories in a file share. You can add or remove manifest directory locations and catalogs by using the  **Trusted Add-in Catalogs** tab in the **Trust Center** dialog box in Project 2013. To show an add-in in Project, the **SourceLocation** element in a manifest must point to an existing website or HTML source file.</span></span>

> [!NOTE]
> <span data-ttu-id="88c83-p127">必须安装 Internet Explorer 9（或更高版本），但它不一定是默认浏览器。Office 加载项需要 Internet Explorer 9 中的组件。默认浏览器可以是 Internet Explorer 9、Safari 5.0.6、Firefox 5、Chrome 13 或上述浏览器之一的更高版本。</span><span class="sxs-lookup"><span data-stu-id="88c83-p127">Internet Explorer 9 (or later) must be installed, but does not have to be the default browser. Office Add-ins require components in Internet Explorer 9. The default browser can be Internet Explorer 9, Safari 5.0.6, Firefox 5, Chrome 13, or a later version of one of these browsers.</span></span>

<span data-ttu-id="88c83-p128">在过程 2 中，在安装 Project 2013 的本地计算机上安装 Bing 搜索加载项。但是，由于加载项基础架构不直接使用本地文件路径，如  `C:\Project\AppManifests`，因此您可以在本地计算机上创建网络共享。如果您喜欢，可以在远程计算机上创建文件共享。</span><span class="sxs-lookup"><span data-stu-id="88c83-p128">In Procedure 2, the Bing Search add-in is installed on the local computer where Project 2013 is installed. However, because the add-in infrastructure does not directly use local file paths such as  `C:\Project\AppManifests`, you can create a network share on the local computer. If you prefer, you can create a file share on a remote computer.</span></span>

### <a name="procedure-2-to-install-the-bing-search-add-in"></a><span data-ttu-id="88c83-p129">过程 2. 安装 Bing 搜索加载项</span><span class="sxs-lookup"><span data-stu-id="88c83-p129">Procedure 2. To install the Bing Search add-in</span></span>

1. <span data-ttu-id="88c83-p130">为加载项清单创建本地目录。例如，创建  `C:\Project\AppManifests` 目录。</span><span class="sxs-lookup"><span data-stu-id="88c83-p130">Create a local directory for add-in manifests. For example, create the  `C:\Project\AppManifests` directory.</span></span>

2. <span data-ttu-id="88c83-191">将  `C:\Project\AppManifests` 目录共享为AppManifests，这样文件共享的网络路径就变为  `\\ServerName\AppManifests`。</span><span class="sxs-lookup"><span data-stu-id="88c83-191">Share the  `C:\Project\AppManifests` directory asAppManifests, so the network path to the file share becomes  `\\ServerName\AppManifests`.</span></span>

3. <span data-ttu-id="88c83-192">将 BingSearch.XML 清单文件复制到  `C:\Project\AppManifests` 目录。</span><span class="sxs-lookup"><span data-stu-id="88c83-192">Copy the BingSearch.xml manifest file to the  `C:\Project\AppManifests` directory.</span></span>

4. <span data-ttu-id="88c83-193">在 Project 2013 中，打开“**Project 选项**”对话框，选择“**信任中心**”，然后选择“**信任中心设置**”。</span><span class="sxs-lookup"><span data-stu-id="88c83-193">In Project 2013, open the  **Project Options** dialog box, choose **Trust Center**, and then choose  **Trust Center Settings**.</span></span>

5. <span data-ttu-id="88c83-194">在“**信任中心**”对话框的左窗格中，选择“**受信任的外接程序目录**”。</span><span class="sxs-lookup"><span data-stu-id="88c83-194">In the  **Trust Center** dialog box, in the left pane, choose **Trusted Add-in Catalogs**.</span></span>

6. <span data-ttu-id="88c83-195">在“**受信任的外接程序目录**”窗格（请参阅图 1）中的“**目录 URL**”文本框中添加 `\\ServerName\AppManifests` 路径，选择“**添加目录**”，然后选择“**确定**”。</span><span class="sxs-lookup"><span data-stu-id="88c83-195">In the  **Trusted Add-in Catalogs** pane (see Figure 1), add the `\\ServerName\AppManifests` path in the **Catalog Url** text box, choose **Add Catalog**, and then choose  **OK**.</span></span>

    > [!NOTE]
    > <span data-ttu-id="88c83-p131">图 1 展示了“受信任的目录地址”\*\*\*\* 列表中的两个文件共享和一个虚构的专用目录 URL。只有一个文件共享可以成为默认文件共享，并且只有一个目录 URL 可以成为默认目录。例如，如果将 `\\Server2\AppManifests` 设置为默认，Project 就会清除 `\\ServerName\AppManifests` 的“默认”\*\*\*\* 复选框。如果更改默认选择，可以选择“清除”\*\*\*\* 删除已安装的加载项，再重启 Project。如果在 Project 处于打开状态时，将加载项添加到默认文件共享或 SharePoint 目录，应重启 Project。</span><span class="sxs-lookup"><span data-stu-id="88c83-p131">Figure 1 shows two file shares and one hypothetical URL for a private catalog in the  **Trusted Catalog Address** list. Only one file share can be the default file share and only one catalog URL can be the default catalog. For example, if you set `\\Server2\AppManifests` as the default, Project clears the **Default** check box for `\\ServerName\AppManifests`.If you change the default selection, you can choose  **Clear** to remove installed add-ins, and then restart Project. If you add an add-in to the default file share or SharePoint catalog while Project is open, you should restart Project.</span></span>

    <span data-ttu-id="88c83-200">*图 1：使用信任中心添加加载项清单目录*</span><span class="sxs-lookup"><span data-stu-id="88c83-200">*Figure 1. Using the Trust Center to add catalogs of add-in manifests*</span></span>

    ![使用信任中心添加应用程序清单](../images/pj15-agave-overview-trust-centers.png)

7. <span data-ttu-id="88c83-p132">在“**项目**”功能区，选择“**Office 外接程序**”下拉菜单，然后选择“**查看所有**”。在“**插入外接程序**”对话框中，选择“**共享文件夹**”（见图 2）。</span><span class="sxs-lookup"><span data-stu-id="88c83-p132">On the  **Project** ribbon, choose the **Office Add-ins** drop-down menu, and then choose **See All**. In the  **Insert Add-in** dialog box, choose **SHARED FOLDER** (see Figure 2).</span></span>

    <span data-ttu-id="88c83-204">*图 2：启动文件共享上的加载项*</span><span class="sxs-lookup"><span data-stu-id="88c83-204">*Figure 2. Starting an add-in that is on a file share*</span></span>

    ![启动文件共享上的 Office 应用程序](../images/pj15-agave-overview-start-agave-apps.png)

8. <span data-ttu-id="88c83-206">选择必应搜索外接程序，然后选择“**插入**”。</span><span class="sxs-lookup"><span data-stu-id="88c83-206">Select the Bing Search add-in, and then choose  **Insert**.</span></span>

    <span data-ttu-id="88c83-p133">Bing 搜索加载项显示在任务窗格中，如图 3 所示。可以手动调整任务窗格的大小，并使用 Bing 搜索加载项。</span><span class="sxs-lookup"><span data-stu-id="88c83-p133">The Bing Search add-in shows in a task pane, as in Figure 3. You can manually resize the task pane, and use the Bing Search add-in.</span></span>

    <span data-ttu-id="88c83-209">*图 3：使用“必应搜索”加载项*</span><span class="sxs-lookup"><span data-stu-id="88c83-209">*Figure 3. Using the Bing Search add-in*</span></span>

    ![使用 Bing 搜索应用程序](../images/pj15-agave-overview-bing-search.png)

## <a name="distributing-project-add-ins"></a><span data-ttu-id="88c83-211">分发 Project 加载项</span><span class="sxs-lookup"><span data-stu-id="88c83-211">Distributing Project add-ins</span></span>

<span data-ttu-id="88c83-212">可通过文件共享、SharePoint 库中的应用程序目录或 AppSource 分发加载项。</span><span class="sxs-lookup"><span data-stu-id="88c83-212">You can distribute add-ins through a file share, an add-in catalog in a SharePoint library, or AppSource.</span></span> <span data-ttu-id="88c83-213">有关详细信息，请参阅[发布 Office 加载项](../publish/publish.md)。</span><span class="sxs-lookup"><span data-stu-id="88c83-213">For more information, see [Publish your Office Add-in](../publish/publish.md).</span></span>

## <a name="see-also"></a><span data-ttu-id="88c83-214">另请参阅</span><span class="sxs-lookup"><span data-stu-id="88c83-214">See also</span></span>

- [<span data-ttu-id="88c83-215">Office 加载项平台概述</span><span class="sxs-lookup"><span data-stu-id="88c83-215">Office Add-ins platform overview</span></span>](../overview/office-add-ins.md)
- [<span data-ttu-id="88c83-216">Office 加载项 XML 清单</span><span class="sxs-lookup"><span data-stu-id="88c83-216">Office Add-ins XML manifest</span></span>](../develop/add-in-manifests.md)
- [<span data-ttu-id="88c83-217">适用于 Office 的 JavaScript API</span><span class="sxs-lookup"><span data-stu-id="88c83-217">JavaScript API for Office</span></span>](/office/dev/add-ins/reference/javascript-api-for-office)
- [<span data-ttu-id="88c83-218">使用文本编辑器创建首个 Project 2013 任务窗格加载项</span><span class="sxs-lookup"><span data-stu-id="88c83-218">Create your first task pane add-in for Project 2013 by using a text editor</span></span>](create-your-first-task-pane-add-in-for-project-by-using-a-text-editor.md)
- [<span data-ttu-id="88c83-219">创建将 REST 与本地 Project Server OData 服务结合使用的 Project 加载项</span><span class="sxs-lookup"><span data-stu-id="88c83-219">Create a Project add-in that uses REST with an on-premises Project Server OData service</span></span>](create-a-project-add-in-that-uses-rest-with-an-on-premises-odata-service.md)
- [<span data-ttu-id="88c83-220">Project 2013 SDK 下载</span><span class="sxs-lookup"><span data-stu-id="88c83-220">Project 2013 SDK download</span></span>](https://www.microsoft.com/download/details.aspx?id=30435%20)
