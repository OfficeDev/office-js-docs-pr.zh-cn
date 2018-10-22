---
title: 在 Visual Studio 中创建和调试 Office 加载项
description: ''
ms.date: 10/01/2018
ms.openlocfilehash: 224a4781b894e9bf165d279c30ca16d18bea956d
ms.sourcegitcommit: c400a220783b03a739449e2d3ff00bbffe5ec7c1
ms.translationtype: HT
ms.contentlocale: zh-CN
ms.lasthandoff: 10/20/2018
ms.locfileid: "25681838"
---
# <a name="create-and-debug-office-add-ins-in-visual-studio"></a><span data-ttu-id="5c9ba-102">在 Visual Studio 中创建和调试 Office 加载项</span><span class="sxs-lookup"><span data-stu-id="5c9ba-102">Create and debug Office Add-ins in Visual Studio</span></span>

<span data-ttu-id="5c9ba-p101">本文介绍如何使用 Visual Studio 创建第一个 Office 加载项。本文中的步骤基于 Visual Studio 2017。如果使用的是 Visual Studio 的其他版本，操作步骤可能略有不同。</span><span class="sxs-lookup"><span data-stu-id="5c9ba-p101">This article describes how to use Visual Studio to create your first Office Add-in. The steps in this article based on Visual Studio 2015. If you're using another version of Visual Studio, the procedures might vary slightly.</span></span>

> [!NOTE]
> <span data-ttu-id="5c9ba-106">若要开始创建 OneNote 加载项，请参阅[生成首个 OneNote 加载项](../onenote/onenote-add-ins-getting-started.md)。</span><span class="sxs-lookup"><span data-stu-id="5c9ba-106">To get started with an add-in for OneNote, see [Build your first OneNote add-in](../onenote/onenote-add-ins-getting-started.md).</span></span>

## <a name="create-an-office-add-in-project-in-visual-studio"></a><span data-ttu-id="5c9ba-107">在 Visual Studio 中创建 Office 加载项项目</span><span class="sxs-lookup"><span data-stu-id="5c9ba-107">Create an Office Add-in project in Visual Studio</span></span>


<span data-ttu-id="5c9ba-p102">首先，请确保已安装 [Office 开发人员工具](https://www.visualstudio.com/features/office-tools-vs.aspx)和一版 Microsoft Office。可以加入 [Office 365 开发人员计划](https://developer.microsoft.com/office/dev-program)，也可以按照下面的说明操作，以获取[最新版](../develop/install-latest-office-version.md)。</span><span class="sxs-lookup"><span data-stu-id="5c9ba-p102">To get started, make sure you have the [Office Developer Tools](https://www.visualstudio.com/features/office-tools-vs.aspx) installed, and a version of Microsoft Office. You can join the [Office 365 Developer Program](https://developer.microsoft.com/office/dev-program), or follow these instructions to get the [latest version](../develop/install-latest-office-version.md).</span></span>

1. <span data-ttu-id="5c9ba-110">在 Visual Studio 菜单栏中，依次选择**文件** > **新建** > **项目**。</span><span class="sxs-lookup"><span data-stu-id="5c9ba-110">On the Visual Studio menu bar, choose  **File** > **New** > **Project**.</span></span>
2. <span data-ttu-id="5c9ba-111">在 **Visual C#** 或 **Visual Basic** 下的项目类型列表中，展开 **Office/SharePoint**，选择 **Web 加载项**，然后选择加载项项目之一。</span><span class="sxs-lookup"><span data-stu-id="5c9ba-111">In the list of project types under  **Visual C#** or **Visual Basic**, expand  **Office/SharePoint**, choose  **Web Add-ins**, and then select one of the Add-in projects.</span></span>
3. <span data-ttu-id="5c9ba-112">命名此项目，再选择**确定**以创建项目。</span><span class="sxs-lookup"><span data-stu-id="5c9ba-112">Name the project, and then choose  **OK** to create the project.</span></span>

<span data-ttu-id="5c9ba-113">在 Visual Studio 2017 中，选择**确定**后，以下加载项项目模板会有额外选择：</span><span class="sxs-lookup"><span data-stu-id="5c9ba-113">In Visual Studio 2017, the following add-in project templates have additional choices after you choose **OK**:</span></span>

<span data-ttu-id="5c9ba-114">**PowerPoint**</span><span class="sxs-lookup"><span data-stu-id="5c9ba-114">**PowerPoint**</span></span>
- <span data-ttu-id="5c9ba-115">你可以选择**将新功能添加到 PowerPoint**，这会创建任务窗格加载项。</span><span class="sxs-lookup"><span data-stu-id="5c9ba-115">You can choose to **Add new functionalities to PowerPoint** which creates a task pane add-in.</span></span>
- <span data-ttu-id="5c9ba-116">或者，可以选择**将内容插入 PowerPoint 幻灯片**，这会创建内容加载项。</span><span class="sxs-lookup"><span data-stu-id="5c9ba-116">Or you can choose to **Insert content into PowerPoint slides** which creates a content add-in.</span></span>

<span data-ttu-id="5c9ba-117">**Excel**</span><span class="sxs-lookup"><span data-stu-id="5c9ba-117">**Excel**</span></span> 
- <span data-ttu-id="5c9ba-118">你可以选择**将新功能添加到 Excel**，这会创建任务窗格加载项。</span><span class="sxs-lookup"><span data-stu-id="5c9ba-118">You can choose to **Add new functionalities to Excel** which creates a task pane add-in.</span></span>
- <span data-ttu-id="5c9ba-119">或者，可以选择**将内容插入 Excel 电子表格**，这会创建内容加载项。</span><span class="sxs-lookup"><span data-stu-id="5c9ba-119">Or you can choose to **Insert content into Excel spreadsheet** which creates a content add-in.</span></span>
    - <span data-ttu-id="5c9ba-120">如果创建内容加载项，你可以有**基本加载项**的额外选择，这会议最少起始代码创建内容加载项项目。</span><span class="sxs-lookup"><span data-stu-id="5c9ba-120">If you create a content add-in, you have an additional choice of **Basic Add-in** which creates a content add-in project with minimal starter code.</span></span>
    - <span data-ttu-id="5c9ba-121">或者，可以选择**文档可视化加载项**，这包括可视化并绑定到数据的起始代码。</span><span class="sxs-lookup"><span data-stu-id="5c9ba-121">Or you can choose a **Document Visualization Add-in** which includes starter code to visualize and bind to data.</span></span>

<span data-ttu-id="5c9ba-122">完成该向导后，Visual Studio 会为你创建包含两个项目的解决方案。</span><span class="sxs-lookup"><span data-stu-id="5c9ba-122">When you've completed the wizard, Visual Studio creates a solution for you that contains two projects.</span></span> <span data-ttu-id="5c9ba-123">你将看到默认 Home.html 页面打开。</span><span class="sxs-lookup"><span data-stu-id="5c9ba-123">You'll see the default Home.html page open.</span></span>

|<span data-ttu-id="5c9ba-124">**项目**</span><span class="sxs-lookup"><span data-stu-id="5c9ba-124">**Project**</span></span>|<span data-ttu-id="5c9ba-125">**描述**</span><span class="sxs-lookup"><span data-stu-id="5c9ba-125">**Description**</span></span>|
|:-----|:-----|
|<span data-ttu-id="5c9ba-126">加载项项目</span><span class="sxs-lookup"><span data-stu-id="5c9ba-126">Add-in project</span></span>|<span data-ttu-id="5c9ba-p104">仅包含一个 XML 清单文件，该文件包含描述你加载项的所有设置。这些设置可帮助 Office 主机确定应何时激活加载项，以及在何处显示加载项。Visual Studio 会为你生成此文件的内容，以便你能够立即运行项目并使用加载项。你可以通过使用清单编辑器来随时更改这些设置。</span><span class="sxs-lookup"><span data-stu-id="5c9ba-p104">Contains only an XML manifest file, which contains all the settings that describe your add-in. These settings help the Office host determine when your add-in should be activated and where the add-in should appear. Visual Studio generates the contents of this file for you so that you can run the project and use your add-in immediately. You change these settings any time by using the Manifest editor.</span></span>|
|<span data-ttu-id="5c9ba-131">Web 应用程序项目</span><span class="sxs-lookup"><span data-stu-id="5c9ba-131">Web application project</span></span>|<span data-ttu-id="5c9ba-132">包含加载项的内容页面，其中包括开发 Office 感知 HTML 和 JavaScript 页面所需的全部文件和文件引用。</span><span class="sxs-lookup"><span data-stu-id="5c9ba-132">Contains the content pages of your add-in, including all the files and file references that you need to develop Office-aware HTML and JavaScript pages.</span></span> <span data-ttu-id="5c9ba-133">在用户开发加载项期间，Visual Studio 在本地 IIS 服务器上托管 Web 应用。</span><span class="sxs-lookup"><span data-stu-id="5c9ba-133">While you develop your add-in, Visual Studio hosts the web application on your local IIS server.</span></span> <span data-ttu-id="5c9ba-134">准备好发布时，你必须找到承载此项目的服务器。</span><span class="sxs-lookup"><span data-stu-id="5c9ba-134">When you're ready to publish, you'll have to find a server to host this project.</span></span> <span data-ttu-id="5c9ba-135">如果要了解有关 ASP.NET Web 应用程序项目的更多信息，请参阅 ASP.NET Web 项目。[ ](http://msdn.microsoft.com/library/cdcd712f-96b0-4165-8b5d-9d0566650a28%28Office.15%29.aspx)</span><span class="sxs-lookup"><span data-stu-id="5c9ba-135">To learn more about ASP.NET web application projects, see [ASP.NET Web Projects](http://msdn.microsoft.com/library/cdcd712f-96b0-4165-8b5d-9d0566650a28%28Office.15%29.aspx).</span></span>|

## <a name="modify-your-add-in-settings"></a><span data-ttu-id="5c9ba-136">修改你的加载项设置</span><span class="sxs-lookup"><span data-stu-id="5c9ba-136">Modify your add-in settings</span></span>


<span data-ttu-id="5c9ba-137">若要修改加载项设置，请编辑项目的 XML 清单文件。</span><span class="sxs-lookup"><span data-stu-id="5c9ba-137">To modify the settings of your add-in, edit the XML manifest file of the project.</span></span> <span data-ttu-id="5c9ba-138">在**解决方案资源管理器**中，展开加载项项目节点、展开包含 XML 清单的文件夹并选择 XML 清单。</span><span class="sxs-lookup"><span data-stu-id="5c9ba-138">In  **Solution Explorer**, expand the add-in project node, expand the folder that contains the XML manifest, and choose the XML manifest.</span></span> <span data-ttu-id="5c9ba-139">你可以指向该文件中的任何元素以查看说明该元素的用途的工具提示。</span><span class="sxs-lookup"><span data-stu-id="5c9ba-139">You can point to any element in the file to view a tooltip that describes the purpose of the element.</span></span> <span data-ttu-id="5c9ba-140">有关清单文件的详细信息，请参阅 [Office 加载项 XML 清单](../develop/add-in-manifests.md)。</span><span class="sxs-lookup"><span data-stu-id="5c9ba-140">For more information about the manfiest file, see [Office Add-ins XML manifest](../develop/add-in-manifests.md).</span></span>


## <a name="develop-the-contents-of-your-add-in"></a><span data-ttu-id="5c9ba-141">开发加载项的内容</span><span class="sxs-lookup"><span data-stu-id="5c9ba-141">Develop the contents of your add-in</span></span>

<span data-ttu-id="5c9ba-142">加载项项目允许你修改描述加载项的设置，而 Web 应用程序提供加载项中显示的内容。</span><span class="sxs-lookup"><span data-stu-id="5c9ba-142">While the add-in project lets you modify the settings that describe your add-in, the web application provides the content that appears in the add-in.</span></span> 

<span data-ttu-id="5c9ba-143">Web 应用程序项目包含可以开始使用的默认 HTML 页面和 JavaScript 文件。</span><span class="sxs-lookup"><span data-stu-id="5c9ba-143">The web application project contains a default HTML page and JavaScript file that you can use to get started.</span></span> <span data-ttu-id="5c9ba-144">这些文件包含对其他 JavaScript 库的引用，包括适用于 Office 的 JavaScript API。</span><span class="sxs-lookup"><span data-stu-id="5c9ba-144">These files are convenient because they contain references to other JavaScript libraries including the JavaScript API for Office.</span></span> <span data-ttu-id="5c9ba-145">更新这些文件，并添加更多的 HTML 和 JavaScript 文件可以开发加载项。</span><span class="sxs-lookup"><span data-stu-id="5c9ba-145">You can develop your add-in by updating these files, and adding more HTML and JavaScript files.</span></span> <span data-ttu-id="5c9ba-146">下表介绍了默认 HTML 和 JavaScript 文件。</span><span class="sxs-lookup"><span data-stu-id="5c9ba-146">The following table describes default HTML and JavaScript files.</span></span>

> [!NOTE]
> <span data-ttu-id="5c9ba-147">根据所用项目模板的类型，下表中的文件可能位于 web 项目的根文件夹或 **Home** 文件夹。</span><span class="sxs-lookup"><span data-stu-id="5c9ba-147">The files in the table below may be in the root folder of the web project, or the **Home** folder depending on the type of project template you used.</span></span>

|<span data-ttu-id="5c9ba-148">**文件**</span><span class="sxs-lookup"><span data-stu-id="5c9ba-148">**File**</span></span>|<span data-ttu-id="5c9ba-149">**描述**</span><span class="sxs-lookup"><span data-stu-id="5c9ba-149">**Description**</span></span>|
|:-----|:-----|
|<span data-ttu-id="5c9ba-150">**Home.html**</span><span class="sxs-lookup"><span data-stu-id="5c9ba-150">**Home.html**</span></span>|<span data-ttu-id="5c9ba-151">加载项的默认 HTML 页面。</span><span class="sxs-lookup"><span data-stu-id="5c9ba-151">The default HTML page of the add-in.</span></span> <span data-ttu-id="5c9ba-152">在文档、电子邮件或约会项目中激活此页面时，它会显示为加载项内的第一个页面。</span><span class="sxs-lookup"><span data-stu-id="5c9ba-152">This page appears as the first page inside of the add-in when it is activated in a document, email message or appointment item.</span></span> <span data-ttu-id="5c9ba-153">此文件包含你开始使用需要的所有文件引用。</span><span class="sxs-lookup"><span data-stu-id="5c9ba-153">This file is convenient because it contains all of the file references that you need to get started.</span></span> <span data-ttu-id="5c9ba-154">你可以将 HTML 代码添加到此文件，开始开发加载项。</span><span class="sxs-lookup"><span data-stu-id="5c9ba-154">You can start developing your add-in by adding your HTML code to this file.</span></span>|
|<span data-ttu-id="5c9ba-155">**Home.js**</span><span class="sxs-lookup"><span data-stu-id="5c9ba-155">**Home.js**</span></span>|<span data-ttu-id="5c9ba-156">与 Home.html 页面关联的 JavaScript 文件。</span><span class="sxs-lookup"><span data-stu-id="5c9ba-156">The JavaScript file associated with the Home.html page.</span></span> <span data-ttu-id="5c9ba-157">你可以将特定于 Home.html 页面的行为的任何代码置于 Home.js 文件中。</span><span class="sxs-lookup"><span data-stu-id="5c9ba-157">You can place any code that is specific to the behavior of the Home.html page in the Home.js file.</span></span> <span data-ttu-id="5c9ba-158">Home.js 文件包含一些可帮你入门的示例代码。</span><span class="sxs-lookup"><span data-stu-id="5c9ba-158">The Home.js file contains some example code to get you started.</span></span>|
|<span data-ttu-id="5c9ba-159">**Home.css**</span><span class="sxs-lookup"><span data-stu-id="5c9ba-159">**Home.css**</span></span>|<span data-ttu-id="5c9ba-160">定义要应用到加载项的默认样式。</span><span class="sxs-lookup"><span data-stu-id="5c9ba-160">Defines the default styles to apply to your add-in.</span></span> <span data-ttu-id="5c9ba-161">我们建议为设计和样式使用 Office UI Fabric。</span><span class="sxs-lookup"><span data-stu-id="5c9ba-161">We recommend using the Office UI Fabric for design and styles.</span></span> <span data-ttu-id="5c9ba-162">有关详细信息，请参阅 [Office 加载项中的 Office UI Fabric](../design/office-ui-fabric.md)。</span><span class="sxs-lookup"><span data-stu-id="5c9ba-162">For information about Office UI Fabric JS, see [Use Office UI Fabric in Office Add-ins](../design/office-ui-fabric.md).</span></span>|

> [!NOTE]
> <span data-ttu-id="5c9ba-163">你无需使用这些文件。</span><span class="sxs-lookup"><span data-stu-id="5c9ba-163">You don't have to use these files.</span></span> <span data-ttu-id="5c9ba-164">你可以随意将其他文件添加到项目并改为使用这些。</span><span class="sxs-lookup"><span data-stu-id="5c9ba-164">Feel free to add other files to the project and use those instead.</span></span> <span data-ttu-id="5c9ba-165">如果你想让其他 HTML 文件显示为加载项的初始页，请打开清单编辑器，然后将 **SourceLocation** 属性指向文件名。</span><span class="sxs-lookup"><span data-stu-id="5c9ba-165">If you want another HTML file to appear as the initial page of the add-in, open the manifest editor, and then point the  **SourceLocation** property to the name of the file.</span></span>

## <a name="debug-your-add-in"></a><span data-ttu-id="5c9ba-166">调试加载项</span><span class="sxs-lookup"><span data-stu-id="5c9ba-166">Debug your add-in</span></span>

<span data-ttu-id="5c9ba-167">Visual Studio 提供生成和调试属性，以帮助调试加载项。</span><span class="sxs-lookup"><span data-stu-id="5c9ba-167">Visual Studio provides build and debug properties to assist with debugging your add-in.</span></span>

### <a name="review-the-build-and-debug-properties"></a><span data-ttu-id="5c9ba-168">查看生成和调试属性</span><span class="sxs-lookup"><span data-stu-id="5c9ba-168">Review the build and debug properties</span></span>

<span data-ttu-id="5c9ba-p112">在启动解决方案之前，请确认 Visual Studio 将打开你需要的主机应用程序。该信息连同与构建和调试加载项有关的其他几个属性一起显示在项目的属性页中。</span><span class="sxs-lookup"><span data-stu-id="5c9ba-p112">Before you start the solution, verify that Visual Studio will open the host application that you want. That information appears in the property pages of the project along with several other properties that relate to building and debugging the add-in.</span></span>

### <a name="to-open-the-property-pages-of-a-project"></a><span data-ttu-id="5c9ba-171">打开项目的属性页</span><span class="sxs-lookup"><span data-stu-id="5c9ba-171">To open the property pages of a project</span></span>

1. <span data-ttu-id="5c9ba-172">在**解决方案资源管理器**中，选择基本加载项项目（非 Web 项目）。</span><span class="sxs-lookup"><span data-stu-id="5c9ba-172">In  **Solution Explorer**, choose the basic add-in project (not the Web project).</span></span>    
2. <span data-ttu-id="5c9ba-173">在菜单栏上，依次选择**视图** >   **属性窗口**。</span><span class="sxs-lookup"><span data-stu-id="5c9ba-173">On the menu bar, choose  View,  Properties Window.</span></span>
    
<span data-ttu-id="5c9ba-174">下表介绍了项目的属性。</span><span class="sxs-lookup"><span data-stu-id="5c9ba-174">The following table describes the properties of the project.</span></span>



|<span data-ttu-id="5c9ba-175">**属性**</span><span class="sxs-lookup"><span data-stu-id="5c9ba-175">**Property**</span></span>|<span data-ttu-id="5c9ba-176">**描述**</span><span class="sxs-lookup"><span data-stu-id="5c9ba-176">**Description**</span></span>|
|:-----|:-----|
|<span data-ttu-id="5c9ba-177">**启动操作**</span><span class="sxs-lookup"><span data-stu-id="5c9ba-177">**Start Action**</span></span>|<span data-ttu-id="5c9ba-178">指定是否在 Office 桌面客户端或在指定浏览器的 Office Online 客户端调试加载项。</span><span class="sxs-lookup"><span data-stu-id="5c9ba-178">Specifies whether to debug your add-in in an Office desktop client or in an Office Online client in the specified browser.</span></span>|
|<span data-ttu-id="5c9ba-179">**启动文档**（仅限内容和任务窗格加载项）</span><span class="sxs-lookup"><span data-stu-id="5c9ba-179">**Start Document** (Content and task pane add-ins only)</span></span>|<span data-ttu-id="5c9ba-180">指定要在启动项目时打开的文档。</span><span class="sxs-lookup"><span data-stu-id="5c9ba-180">Specifies what document to open when you start the project.</span></span>|
|<span data-ttu-id="5c9ba-181">**Web 项目**</span><span class="sxs-lookup"><span data-stu-id="5c9ba-181">**Web Project**</span></span>|<span data-ttu-id="5c9ba-182">指定与加载项关联的 Web 项目的名称。</span><span class="sxs-lookup"><span data-stu-id="5c9ba-182">Specifies the name of the web project associated with the add-in.</span></span>|
|<span data-ttu-id="5c9ba-183">**电子邮件地址**（仅限 Outlook 加载项）</span><span class="sxs-lookup"><span data-stu-id="5c9ba-183">**Email Address** (Outlook add-ins only)</span></span>|<span data-ttu-id="5c9ba-184">指定 Exchange Server 或 Exchange Online 中你想用来测试你的 Outlook 加载项的用户帐户的电子邮件地址。</span><span class="sxs-lookup"><span data-stu-id="5c9ba-184">Specifies the email address of the user account in Exchange Server or Exchange Online that you want to test your Outlook add-in with.</span></span>|
|<span data-ttu-id="5c9ba-185">**EWS URL**（仅限 Outlook 加载项）</span><span class="sxs-lookup"><span data-stu-id="5c9ba-185">**EWS Url** (Outlook add-ins only)</span></span>|<span data-ttu-id="5c9ba-186">Exchange Web 服务 URL（例如：https://www.contoso.com/ews/exchange.aspx)。</span><span class="sxs-lookup"><span data-stu-id="5c9ba-186">Exchange Web service URL (For example: https://www.contoso.com/ews/exchange.aspx).</span></span> |
|<span data-ttu-id="5c9ba-187">**OWA URL**（仅限 Outlook 加载项）</span><span class="sxs-lookup"><span data-stu-id="5c9ba-187">**OWA Url** (Outlook add-ins only)</span></span>|<span data-ttu-id="5c9ba-188">Outlook Web App URL（例如，https://www.contoso.com/owa)。</span><span class="sxs-lookup"><span data-stu-id="5c9ba-188">Outlook Web App URL (For example: https://www.contoso.com/owa).</span></span>|
|<span data-ttu-id="5c9ba-189">**用户名**（仅限 Outlook 加载项）</span><span class="sxs-lookup"><span data-stu-id="5c9ba-189">**User name** (Outlook add-ins only)</span></span>|<span data-ttu-id="5c9ba-190">指定 Exchange Server 或 Exchange Online 中的用户帐户名称。</span><span class="sxs-lookup"><span data-stu-id="5c9ba-190">Specifies the name of your user account in Exchange Server or Exchange Online.</span></span>|
|<span data-ttu-id="5c9ba-191">**项目文件**</span><span class="sxs-lookup"><span data-stu-id="5c9ba-191">**Project File**</span></span>|<span data-ttu-id="5c9ba-192">指定包含生成、配置和有关项目的其他信息的文件名称。</span><span class="sxs-lookup"><span data-stu-id="5c9ba-192">Specifies the name of the file containing build, configuration, and other information about the project.</span></span>|
|<span data-ttu-id="5c9ba-193">**项目文件夹**</span><span class="sxs-lookup"><span data-stu-id="5c9ba-193">**Project Folder**</span></span>|<span data-ttu-id="5c9ba-194">项目文件的位置。</span><span class="sxs-lookup"><span data-stu-id="5c9ba-194">The location of the project file.</span></span>|

### <a name="use-an-existing-document-to-debug-the-add-in-content-and-task-pane-add-ins-only"></a><span data-ttu-id="5c9ba-195">使用现有文档调试加载项（仅限内容和任务窗格加载项）</span><span class="sxs-lookup"><span data-stu-id="5c9ba-195">Use an existing document to debug the add-in (content and task pane add-ins only)</span></span>

<span data-ttu-id="5c9ba-p113">你可以将文档添加到加载项项目。如果你有包含要用于加载项的测试数据的文档，Visual Studio 将在你启动项目时为你打开该文档。</span><span class="sxs-lookup"><span data-stu-id="5c9ba-p113">You can add documents to the add-in project. If you have a document that contains test data that you want to use with your add-in, Visual Studio opens that document for you when you start the project.</span></span>

### <a name="to-use-an-existing-document-to-debug-the-add-in"></a><span data-ttu-id="5c9ba-198">使用现有文档调试加载项</span><span class="sxs-lookup"><span data-stu-id="5c9ba-198">To use an existing document to debug the add-in</span></span>

1. <span data-ttu-id="5c9ba-199">在**解决方案资源管理器**中，选择加载项项目文件夹。</span><span class="sxs-lookup"><span data-stu-id="5c9ba-199">In  **Solution Explorer**, choose the add-in project folder.</span></span>
    
    > [!NOTE]
    > <span data-ttu-id="5c9ba-200">选择加载项项目，而不是 Web 应用程序项目。</span><span class="sxs-lookup"><span data-stu-id="5c9ba-200">Choose the add-in project and not the web application project.</span></span>

2. <span data-ttu-id="5c9ba-201">在**项目**菜单中，选择**添加现有项**。</span><span class="sxs-lookup"><span data-stu-id="5c9ba-201">On the  **Project** menu, choose **Add Existing Item**.</span></span>
    
3. <span data-ttu-id="5c9ba-202">在**添加现有项**对话框中，找到并选择要添加的文档。</span><span class="sxs-lookup"><span data-stu-id="5c9ba-202">In the  **Add Existing Item** dialog box, locate and select the document that you want to add.</span></span>
    
4. <span data-ttu-id="5c9ba-203">选择**添加**按钮以向你的项目添加文档。</span><span class="sxs-lookup"><span data-stu-id="5c9ba-203">Choose the  **Add** button to add the document to your project.</span></span>
    
5. <span data-ttu-id="5c9ba-204">在**解决方案资源管理器**中，选择加载项项目文件夹。</span><span class="sxs-lookup"><span data-stu-id="5c9ba-204">In  **Solution Explorer**, choose the add-in project folder.</span></span>
6. <span data-ttu-id="5c9ba-205">在菜单栏上，依次选择**视图** >  **属性窗口**。</span><span class="sxs-lookup"><span data-stu-id="5c9ba-205">On the menu bar, choose  View,  Properties Window.</span></span>
7. <span data-ttu-id="5c9ba-206">在属性窗口中，选择**启动文档**列表，然后选择添加到项目的文档。</span><span class="sxs-lookup"><span data-stu-id="5c9ba-206">In the  **Start Document** list, choose the document that you added to the project, and then choose the OK button to close the property pages.</span></span> <span data-ttu-id="5c9ba-207">现在，项目将配置为在现有的文档中启动加载项。</span><span class="sxs-lookup"><span data-stu-id="5c9ba-207">Now the project is configured to start your add-in in your existing document.</span></span>

### <a name="start-the-solution"></a><span data-ttu-id="5c9ba-208">启动解决方案</span><span class="sxs-lookup"><span data-stu-id="5c9ba-208">Start the solution</span></span>

<span data-ttu-id="5c9ba-209">选择**调试** > **启动调试**，从菜单栏启动解决方案。</span><span class="sxs-lookup"><span data-stu-id="5c9ba-209">Start the solution from the menu bar by choosing **Debug** > **Start Debugging**.</span></span> <span data-ttu-id="5c9ba-210">Visual Studio 将自动生成解决方案，并启动 Office 来承载你的加载项。</span><span class="sxs-lookup"><span data-stu-id="5c9ba-210">Visual Studio will automatically build the solution and start Office to host your add-in.</span></span>

<span data-ttu-id="5c9ba-211">当 Visual Studio 生成项目时，它将执行以下任务：</span><span class="sxs-lookup"><span data-stu-id="5c9ba-211">When Visual Studio builds the project it performs the following tasks:</span></span>

1. <span data-ttu-id="5c9ba-p116">创建 XML 清单文件的副本并将其添加到  _ProjectName_\Output 目录。主机应用程序将在你启动 Visual Studio 并调试加载项时使用此副本。</span><span class="sxs-lookup"><span data-stu-id="5c9ba-p116">Creates a copy of the XML manifest file and adds it to  _ProjectName_\Output directory. The host application consumes this copy when you start Visual Studio and debug the add-in.</span></span>
    
2. <span data-ttu-id="5c9ba-214">在计算机上创建一组允许加载项在主机应用程序中显示的注册表项。</span><span class="sxs-lookup"><span data-stu-id="5c9ba-214">Creates a set of registry entries on your computer that enable the add-in to appear in the host application.</span></span>
    
3. <span data-ttu-id="5c9ba-215">生成网络应用程序项目，然后将其部署到本地 IIS Web 服务器（http://localhost)</span><span class="sxs-lookup"><span data-stu-id="5c9ba-215">Builds the web application project, and then deploys it to the local IIS web server (http://localhost).</span></span> 
    
<span data-ttu-id="5c9ba-216">接下来，Visual Studio 会执行以下操作：</span><span class="sxs-lookup"><span data-stu-id="5c9ba-216">Next, Visual Studio does the following:</span></span>

1. <span data-ttu-id="5c9ba-217">修改 XML 显示文件的 [SourceLocation](https://docs.microsoft.com/office/dev/add-ins/reference/manifest/sourcelocation?view=office-js)元素，通过将 ～remoteAppUrl 标记替换为起始页的完全限定地址（例如，http://localhost/MyAgave.html)）。</span><span class="sxs-lookup"><span data-stu-id="5c9ba-217">Modifies the [SourceLocation](https://docs.microsoft.com/office/dev/add-ins/reference/manifest/sourcelocation?view=office-js) element of the XML manifest file by replacing the ~remoteAppUrl token with the fully qualified address of the start page (for example, http://localhost/MyAgave.html).</span></span>
    
2. <span data-ttu-id="5c9ba-218">在 IIS Express 中启动 Web 应用程序项目。</span><span class="sxs-lookup"><span data-stu-id="5c9ba-218">Starts the web application project in IIS Express.</span></span>
    
3. <span data-ttu-id="5c9ba-219">打开主机应用程序。</span><span class="sxs-lookup"><span data-stu-id="5c9ba-219">Opens the host application.</span></span> 
    
<span data-ttu-id="5c9ba-p117">生成项目时，Visual Studio 不会显示“**输出**”窗口中的验证错误。Visual Studio 报告“**错误列表**”窗口中出现的错误和警告。通过在代码和文本编辑器中显示不同颜色的波浪下划线（称为波浪线），Visual Studio 还报告验证错误。通过这些标志，你可以得知 Visual Studio 在代码中检测到的问题。有关详细信息，请参阅 [代码和文本编辑器](https://msdn.microsoft.com/library/se2f663y(v=vs.140).aspx)。有关如何启用或禁用验证的详细信息，请参阅：</span><span class="sxs-lookup"><span data-stu-id="5c9ba-p117">Visual Studio doesn't show validation errors in the  **OUTPUT** window when you build the project. Visual Studio reports errors and warnings in the **ERRORLIST** window as they occur. Visual Studio also reports validation errors by showing wavy underlines (known as squiggles) of different colors in the code and text editor. These marks notify you of problems that Visual Studio detected in your code. For more information, see [Code and Text Editor](https://msdn.microsoft.com/library/se2f663y(v=vs.140).aspx). For more information about how to enable or disable validation, see:</span></span> 

- [<span data-ttu-id="5c9ba-226">选项、文本编辑器、JavaScript 和 IntelliSense</span><span class="sxs-lookup"><span data-stu-id="5c9ba-226">Options, Text Editor, JavaScript, IntelliSense</span></span>](https://docs.microsoft.com/visualstudio/ide/reference/options-text-editor-javascript-intellisense?view=vs-2015)
    
- <span data-ttu-id="5c9ba-227">[操作方法：为 Visual Web Developer 中的 HTML 编辑设置验证选项](https://msdn.microsoft.com/library/0byxkfet(v=vs.100).aspx)</span><span class="sxs-lookup"><span data-stu-id="5c9ba-227">[How to: Set Validation Options for HTML Editing in Visual Web Developer](https://msdn.microsoft.com/library/0byxkfet(v=vs.100).aspx)</span></span>
    
- <span data-ttu-id="5c9ba-228">[有关 CSS，请参阅验证、CSS、文本编辑器和“选项”对话框](https://msdn.microsoft.com/library/se2f663y(v=vs.140).aspx)</span><span class="sxs-lookup"><span data-stu-id="5c9ba-228">[CSS, see Validation, CSS, Text Editor, Options Dialog Box](https://msdn.microsoft.com/library/se2f663y(v=vs.140).aspx)</span></span>
    
<span data-ttu-id="5c9ba-229">若要查看项目中 XML 清单文件的验证规则，请参阅 [Office 加载项 XML 清单](../develop/add-in-manifests.md)。</span><span class="sxs-lookup"><span data-stu-id="5c9ba-229">To review the validation rules of the XML manifest file in your project, see [Office Add-ins XML manifest](../develop/add-in-manifests.md).</span></span>

### <a name="show-an-add-in-in-excel-or-word-and-step-through-your-code"></a><span data-ttu-id="5c9ba-230">在 Excel 或 Word 中显示加载项并单步调试代码</span><span class="sxs-lookup"><span data-stu-id="5c9ba-230">Show an add-in in Excel, Word, or Project and step through your code</span></span>

<span data-ttu-id="5c9ba-231">如果你将加载项项目的“**启动文档**”属性设置为 Excel 或 Word，Visual Studio 会创建一个新文档，加载项会出现。</span><span class="sxs-lookup"><span data-stu-id="5c9ba-231">If you set the  **Start Document** property of the add-in project to Excel or Word, Visual Studio creates a new document and the add-in appears.</span></span> <span data-ttu-id="5c9ba-232">如果你将加载项项目的**启动文档**属性设置为使用现有文档，Visual Studio 会打开该文档，但是你必须手动插入加载项。</span><span class="sxs-lookup"><span data-stu-id="5c9ba-232">If you set the **Start Document** property of the add-in project to use an existing document, Visual Studio opens the document, but you have to insert the add-in manually.</span></span>

1. <span data-ttu-id="5c9ba-233">在 Excel 或 Word 中的**插入**选项卡上个，选择**我的加载项**下拉列表。</span><span class="sxs-lookup"><span data-stu-id="5c9ba-233">In Excel or Word, on the  **Insert** tab, choose the **My Add-ins** drop down list.</span></span> <span data-ttu-id="5c9ba-234">从下拉箭头而不是按钮本身选择列表，这将打开 **Office 加载项**对话框。</span><span class="sxs-lookup"><span data-stu-id="5c9ba-234">Choose the list from the drop-down arrow, not the button itself which opens the **Office Add-ins** dialog.</span></span>
2. <span data-ttu-id="5c9ba-235">在**开发人员加载项**下，选择加载项。</span><span class="sxs-lookup"><span data-stu-id="5c9ba-235">Under **Developer Add-ins**, choose your add-in.</span></span>

<span data-ttu-id="5c9ba-236">然后，在 Visual Studio 中，可以设置中断点并与你的加载项进行交互，逐行执行 HTML 或 JavaScript 文件中的文件。</span><span class="sxs-lookup"><span data-stu-id="5c9ba-236">In Visual Studio, you can then set break-points. Then, as you interact with your add-in and step through the code in your HTML, JavaScript, and C# or VB code files.</span></span>

### <a name="show-the-outlook-add-in-in-outlook-and-step-through-your-code"></a><span data-ttu-id="5c9ba-237">在 Outlook 中显示 Outlook 加载项并单步调试代码</span><span class="sxs-lookup"><span data-stu-id="5c9ba-237">Show the Outlook add-in in Outlook and step through your code</span></span>

<span data-ttu-id="5c9ba-238">若要在 Outlook 中查看加载项，请打开一个电子邮件或约会项目。</span><span class="sxs-lookup"><span data-stu-id="5c9ba-238">To view the add-in in Outlook, open an email message or appointment item.</span></span>

<span data-ttu-id="5c9ba-p120">只要满足激活条件，Outlook 便会为项目激活加载项。加载项栏显示在"检查器"窗口或阅读窗格的顶部，Outlook 加载项显示为加载项栏中的一个按钮。如果你的加载项有加载项命令，那么在默认选项卡或指定的自定义选项卡中将有一个按钮显示在功能区中，而该加载项将不会显示在加载项栏中。</span><span class="sxs-lookup"><span data-stu-id="5c9ba-p120">Outlook activates the add-in for the item as long as the activation criteria are met. The add-in bar appears at the top of the Inspector window or Reading Pane, and your Outlook add-in appears as a button in the add-in bar. If your add-in has an add-in command, a button will appear in the ribbon, either in the default tab or a specified custom tab, and the add-in will not appear in the add-in bar.</span></span>

<span data-ttu-id="5c9ba-242">若要查看 Outlook 加载项，请选择 Outlook 加载项的按钮。</span><span class="sxs-lookup"><span data-stu-id="5c9ba-242">To view your Outlook add-in, choose the button for your Outlook add-in.</span></span>

<span data-ttu-id="5c9ba-243">然后，在 Visual Studio 中，可以设置中断点并与你的加载项进行交互，逐行执行 HTML 或 JavaScript 文件中的文件。</span><span class="sxs-lookup"><span data-stu-id="5c9ba-243">In Visual Studio, you can then set break-points. Then, as you interact with your add-in and step through the code in your HTML, JavaScript, and C# or VB code files.</span></span>

<span data-ttu-id="5c9ba-p121">你还可以更改代码并在 Outlook 加载项中查看这些更改的效果，而不必关闭 Office 加载项并再次启动项目。在 Outlook 中，只需打开 Outlook 加载项的快捷菜单，然后选择**重新加载**即可。</span><span class="sxs-lookup"><span data-stu-id="5c9ba-p121">You can also change your code and review the effects of those changes in your Outlook add-in without having to close the Office Add-in and start the project again. In Outlook, just open the shortcut menu for the Outlook add-in, and then choose  **Reload**.</span></span>


### <a name="modify-code-and-continue-to-debug-the-add-in-without-having-to-start-the-project-again"></a><span data-ttu-id="5c9ba-246">修改代码并继续调试加载项，而不必再次启动项目</span><span class="sxs-lookup"><span data-stu-id="5c9ba-246">Modify code and continue to debug the add-in without having to start the project again</span></span>

<span data-ttu-id="5c9ba-247">你可以更改代码并在加载项中查看这些更改的效果，无需关闭主机应用程序并重新启动该项目。</span><span class="sxs-lookup"><span data-stu-id="5c9ba-247">You can change your code and review the effects of those changes in your add-in without having to close the host application and start the project again.</span></span> <span data-ttu-id="5c9ba-248">更改并保存代码后，打开加载项的快捷菜单，然后选择**重新加载**。</span><span class="sxs-lookup"><span data-stu-id="5c9ba-248">After you change your code, open the shortcut menu for the add-in, and then choose  **Reload**.</span></span>
    

## <a name="next-steps"></a><span data-ttu-id="5c9ba-249">后续步骤</span><span class="sxs-lookup"><span data-stu-id="5c9ba-249">Next steps</span></span>

- [<span data-ttu-id="5c9ba-250">部署和发布 Office 加载项</span><span class="sxs-lookup"><span data-stu-id="5c9ba-250">Deploy and publish your Office Add-in</span></span>](../publish/publish.md)
    
