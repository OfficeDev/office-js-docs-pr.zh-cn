---
title: 在 Visual Studio 中创建和调试 Office 外接程序
description: 使用 Visual Studio 在 Windows 上的 Office 桌面客户端中创建和调试 Office 加载项
ms.date: 06/20/2019
localization_priority: Priority
ms.openlocfilehash: 9cc4d50d9b61daa4b1f55f7dd4c1e1156f8d959c
ms.sourcegitcommit: 382e2735a1295da914f2bfc38883e518070cec61
ms.translationtype: HT
ms.contentlocale: zh-CN
ms.lasthandoff: 06/21/2019
ms.locfileid: "35128183"
---
# <a name="create-and-debug-office-add-ins-in-visual-studio"></a><span data-ttu-id="fd05a-103">在 Visual Studio 中创建和调试 Office 外接程序</span><span class="sxs-lookup"><span data-stu-id="fd05a-103">Create and debug Office Add-ins in Visual Studio</span></span>

<span data-ttu-id="fd05a-104">本文介绍如何使用 Visual Studio 2017 为 Excel、Word、PowerPoint 或 Outlook 创建 Office 外接程序，并在 Windows 上的 Office 桌面客户端中调试外接程序。</span><span class="sxs-lookup"><span data-stu-id="fd05a-104">This article describes how to use Visual Studio 2017 to create an Office Add-in for Excel, Word, PowerPoint, or Outlook and debug the add-in in the Office desktop client on Windows.</span></span> <span data-ttu-id="fd05a-105">如果使用的是 Visual Studio 的其他版本，操作步骤可能略有不同。</span><span class="sxs-lookup"><span data-stu-id="fd05a-105">If you're using another version of Visual Studio, the procedures might vary slightly.</span></span>

> [!NOTE]
> <span data-ttu-id="fd05a-106">Visual Studio 不支持为 OneNote 或 Project 创建 Office 外接程序，但你可以使用 [Office 外接程序的 Yeoman 生成器](https://github.com/OfficeDev/generator-office)来创建这些类型的外接程序。</span><span class="sxs-lookup"><span data-stu-id="fd05a-106">Visual Studio does not support creating Office Add-ins for OneNote or Project, but you can use the [Yeoman generator for Office Add-ins](https://github.com/OfficeDev/generator-office) to create these types of add-ins.</span></span>
> - <span data-ttu-id="fd05a-107">若要开始使用 OneNote 的外接程序，请参阅[生成首个 OneNote 外接程序](../quickstarts/onenote-quickstart.md)。</span><span class="sxs-lookup"><span data-stu-id="fd05a-107">To get started with an add-in for OneNote, see [Build your first OneNote add-in](../quickstarts/onenote-quickstart.md).</span></span>
>
> - <span data-ttu-id="fd05a-108">若要开始使用 Project 的外接程序，请参阅[生成首个 Project 外接程序](../quickstarts/project-quickstart.md)。</span><span class="sxs-lookup"><span data-stu-id="fd05a-108">To get started with an add-in for Project, see [Build your first Project add-in](../quickstarts/project-quickstart.md).</span></span>

## <a name="prerequisites"></a><span data-ttu-id="fd05a-109">先决条件</span><span class="sxs-lookup"><span data-stu-id="fd05a-109">Prerequisites</span></span>

- <span data-ttu-id="fd05a-110">安装了 **Office/SharePoint 开发**工作负载的 [Visual Studio 2017](https://www.visualstudio.com/vs/)</span><span class="sxs-lookup"><span data-stu-id="fd05a-110">[Visual Studio 2017](https://www.visualstudio.com/vs/) with the **Office/SharePoint development** workload installed</span></span>

    > [!TIP]
    > <span data-ttu-id="fd05a-111">如果之前已安装 Visual Studio 2017，请[使用 Visual Studio 安装程序](/visualstudio/install/modify-visual-studio)，以确保安装 **Office/SharePoint 开发**工作负载。</span><span class="sxs-lookup"><span data-stu-id="fd05a-111">If you've previously installed Visual Studio 2017, [use the Visual Studio Installer](/visualstudio/install/modify-visual-studio) to ensure that the **Office/SharePoint development** workload is installed.</span></span> <span data-ttu-id="fd05a-112">如果尚未安装此工作负载，请使用 Visual Studio 安装程序进行[安装](/visualstudio/install/modify-visual-studio?view=vs-2017#modify-workloads)。</span><span class="sxs-lookup"><span data-stu-id="fd05a-112">If this workload is not yet installed, use the Visual Studio Installer to [install it](/visualstudio/install/modify-visual-studio?view=vs-2017#modify-workloads).</span></span>

- <span data-ttu-id="fd05a-113">Office 2013 或更高版本</span><span class="sxs-lookup"><span data-stu-id="fd05a-113">Office 2013 or later</span></span>

    > [!TIP]
    > <span data-ttu-id="fd05a-114">如果你还没有 Office，则可加入 [Office 365 开发人员计划](https://developer.microsoft.com/office/dev-program)以获取 Office 365 订阅，或者你可以[注册免费 1 个月的试用版](https://products.office.com/zh-CN/try?legRedir=true&WT.intid1=ODC_ENUS_FX101785584_XT104056786&CorrelationId=64c762de-7a97-4dd1-bb96-e231d7485735)。</span><span class="sxs-lookup"><span data-stu-id="fd05a-114">If you don't already have Office, you can join the [Office 365 Developer Program](https://developer.microsoft.com/office/dev-program) to get an Office 365 subscription, or you can [register for a free 1-month trial](https://products.office.com/en-US/try?legRedir=true&WT.intid1=ODC_ENUS_FX101785584_XT104056786&CorrelationId=64c762de-7a97-4dd1-bb96-e231d7485735).</span></span>

## <a name="create-the-add-in-project-in-visual-studio"></a><span data-ttu-id="fd05a-115">在 Visual Studio 中创建外接程序项目</span><span class="sxs-lookup"><span data-stu-id="fd05a-115">Create the add-in project in Visual Studio</span></span>

<span data-ttu-id="fd05a-116">首先完成以下三个步骤，然后完成后续部分中与你正在创建的外接程序类型相对应的步骤。</span><span class="sxs-lookup"><span data-stu-id="fd05a-116">Start by completing these three steps, and then complete the steps in the following section that corresponds to the type of add-in you're creating.</span></span> 

1. <span data-ttu-id="fd05a-117">打开 Visual Studio，在 Visual Studio 菜单栏中，依次选择“**文件**” > “**新建**” > “**项目**”。</span><span class="sxs-lookup"><span data-stu-id="fd05a-117">Open Visual Studio and from the Visual Studio menu bar, choose  **File** > **New** > **Project**.</span></span>

2. <span data-ttu-id="fd05a-118">在 **Visual C#** 或 **Visual Basic** 下的项目类型列表中，展开 **Office/SharePoint**，选择“**外接程序**”，然后选择要创建的外接程序项目的类型。</span><span class="sxs-lookup"><span data-stu-id="fd05a-118">In the list of project types under **Visual C#** or **Visual Basic**, expand  **Office/SharePoint**, choose **Add-ins**, and then choose the type of add-in project you want to create.</span></span> 

3. <span data-ttu-id="fd05a-119">命名此项目，然后选择“**确定**”。</span><span class="sxs-lookup"><span data-stu-id="fd05a-119">Name the project, and then choose **OK**.</span></span>

### <a name="word-web-add-in-or-outlook-web-add-in"></a><span data-ttu-id="fd05a-120">Word Web 外接程序或 Outlook Web 外接程序</span><span class="sxs-lookup"><span data-stu-id="fd05a-120">Word Web Add-in or Outlook Web Add-in</span></span>

<span data-ttu-id="fd05a-121">如果你已选择创建 **Word Web 外接程序**或 **Outlook Web 外接程序**，Visual Studio 将创建一个解决方案，并在“**解决方案资源管理器**”中显示这两个项目。</span><span class="sxs-lookup"><span data-stu-id="fd05a-121">If you've chosen to create a **Word Web Add-in** or an **Outlook Web Add-in**, Visual Studio creates a solution and its two projects appear in **Solution Explorer**.</span></span> <span data-ttu-id="fd05a-122">接下来，你可以[浏览 Visual Studio 解决方案](#explore-the-visual-studio-solution)。</span><span class="sxs-lookup"><span data-stu-id="fd05a-122">Next, you can [explore the Visual Studio solution](#explore-the-visual-studio-solution).</span></span> 

### <a name="powerpoint-web-add-in"></a><span data-ttu-id="fd05a-123">PowerPoint Web 外接程序</span><span class="sxs-lookup"><span data-stu-id="fd05a-123">PowerPoint Web Add-in</span></span>

<span data-ttu-id="fd05a-124">如果你已选择创建 **PowerPoint Web 外接程序**，则会出现“**创建 Office 外接程序**”对话框。</span><span class="sxs-lookup"><span data-stu-id="fd05a-124">If you've chosen to create a **PowerPoint Web Add-in**, the **Create Office Add-in** dialog appears.</span></span> 

- <span data-ttu-id="fd05a-125">若要创建任务窗格外接程序，请选择“**向 PowerPoint 添加新功能**”，然后选择“**完成**”按钮以创建 Visual Studio 解决方案。</span><span class="sxs-lookup"><span data-stu-id="fd05a-125">To create a task pane add-in, select **Add new functionalities to PowerPoint** and then choose the **Finish** button to create the Visual Studio solution.</span></span>

- <span data-ttu-id="fd05a-126">若要创建内容外接程序，请选择“**向 PowerPoint 幻灯片插入内容**”，然后选择“**完成**”按钮以创建 Visual Studio 解决方案。</span><span class="sxs-lookup"><span data-stu-id="fd05a-126">To create a content add-in, select **Insert content into PowerPoint slides** and then choose the **Finish** button to create the Visual Studio solution.</span></span>

<span data-ttu-id="fd05a-127">接下来，你可以[浏览 Visual Studio 解决方案](#explore-the-visual-studio-solution)。</span><span class="sxs-lookup"><span data-stu-id="fd05a-127">Next, you can [explore the Visual Studio solution](#explore-the-visual-studio-solution).</span></span>

### <a name="excel-web-add-in"></a><span data-ttu-id="fd05a-128">Excel Web 外接程序</span><span class="sxs-lookup"><span data-stu-id="fd05a-128">Excel Web Add-in</span></span>

<span data-ttu-id="fd05a-129">如果你已选择创建 **Excel Web 外接程序**，则会出现“**创建 Office 外接程序**”对话框。</span><span class="sxs-lookup"><span data-stu-id="fd05a-129">If you've chosen to create an **Excel Web Add-in**, the **Create Office Add-in** dialog appears.</span></span> 

- <span data-ttu-id="fd05a-130">若要创建任务窗格外接程序，请选择“**向 Excel 添加新功能**”，然后选择“**完成**”按钮以创建 Visual Studio 解决方案。</span><span class="sxs-lookup"><span data-stu-id="fd05a-130">To create a task pane add-in, select **Add new functionalities to Excel** and then choose the **Finish** button to create the Visual Studio solution.</span></span>

- <span data-ttu-id="fd05a-131">若要创建内容外接程序，请选择“**向 Excel 电子表格插入内容**”，选择“**下一步**”按钮，选择以下选项之一，然后选择“**完成**”按钮以创建 Visual Studio 解决方案：</span><span class="sxs-lookup"><span data-stu-id="fd05a-131">To create a content add-in, select **Insert content into Excel spreadsheets**, choose the **Next** button, select one of the following options, and then choose the **Finish** button to create the Visual Studio solution:</span></span>

    - <span data-ttu-id="fd05a-132">**基本外接程序** - 使用最少的入门代码创建内容外接程序项目</span><span class="sxs-lookup"><span data-stu-id="fd05a-132">**Basic Add-in** - to create a content add-in project with minimal starter code</span></span>

    - <span data-ttu-id="fd05a-133">**文档可视化外接程序** - 使用入门代码创建内容外接程序项目，以实现可视化并绑定到数据</span><span class="sxs-lookup"><span data-stu-id="fd05a-133">**Document Visualization Add-in** - to create a content add-in project with starter code to visualize and bind to data</span></span>  

### <a name="explore-the-visual-studio-solution"></a><span data-ttu-id="fd05a-134">浏览 Visual Studio 解决方案</span><span class="sxs-lookup"><span data-stu-id="fd05a-134">Explore the Visual Studio solution</span></span>

[!include[Description of Visual Studio projects](../includes/quickstart-vs-solution.md)]

## <a name="modify-your-add-in-settings"></a><span data-ttu-id="fd05a-135">修改外接程序设置</span><span class="sxs-lookup"><span data-stu-id="fd05a-135">Modify your add-in settings</span></span>

<span data-ttu-id="fd05a-136">若要修改外接程序的设置，请编辑外接程序项目中的 XML 清单文件。</span><span class="sxs-lookup"><span data-stu-id="fd05a-136">To modify the settings of your add-in, edit the XML manifest file in the add-in project.</span></span> <span data-ttu-id="fd05a-137">在“**解决方案资源管理器**”中，展开外接程序项目节点，展开包含 XML 清单的文件夹并选择 XML 清单。</span><span class="sxs-lookup"><span data-stu-id="fd05a-137">In  **Solution Explorer**, expand the add-in project node, expand the folder that contains the XML manifest, and choose the XML manifest.</span></span> <span data-ttu-id="fd05a-138">你可以指向该文件中的任何元素以查看说明该元素用途的工具提示。</span><span class="sxs-lookup"><span data-stu-id="fd05a-138">You can point to any element in the file to view a tooltip that describes the purpose of the element.</span></span> <span data-ttu-id="fd05a-139">有关清单文件的详细信息，请参阅 [Office 外接程序 XML 清单](../develop/add-in-manifests.md)。</span><span class="sxs-lookup"><span data-stu-id="fd05a-139">For more information about the manifest file, see [Office Add-ins XML manifest](../develop/add-in-manifests.md).</span></span>

## <a name="develop-the-contents-of-your-add-in"></a><span data-ttu-id="fd05a-140">开发外接程序的内容</span><span class="sxs-lookup"><span data-stu-id="fd05a-140">Develop the contents of your add-in</span></span>

<span data-ttu-id="fd05a-141">加载项项目允许您修改描述加载项的设置，而 Web 应用程序提供加载项中显示的内容。</span><span class="sxs-lookup"><span data-stu-id="fd05a-141">While the add-in project lets you modify the settings that describe your add-in, the web application provides the content that appears in the add-in.</span></span> 

<span data-ttu-id="fd05a-142">Web 应用程序项目包含可用于实现入门的默认 HTML 文件、JavaScript 文件和 CSS 文件。</span><span class="sxs-lookup"><span data-stu-id="fd05a-142">The web application project contains a default HTML file, JavaScript file, and CSS file that you can use to get started.</span></span> <span data-ttu-id="fd05a-143">其中一些文件包含对其他 JavaScript 库的引用，包括适用于 Office 的 JavaScript API。</span><span class="sxs-lookup"><span data-stu-id="fd05a-143">Some of these files contain references to other JavaScript libraries including the JavaScript API for Office.</span></span> <span data-ttu-id="fd05a-144">你可以通过更新这些文件和/或添加更多 HTML 和 JavaScript 文件来开发外接程序。</span><span class="sxs-lookup"><span data-stu-id="fd05a-144">You can develop your add-in by updating these files and/or adding more HTML and JavaScript files.</span></span> <span data-ttu-id="fd05a-145">下表描述了创建 Visual Studio 解决方案时 Web 应用程序项目包含的默认文件。</span><span class="sxs-lookup"><span data-stu-id="fd05a-145">The following table describes the default files that the web application project contains when the Visual Studio solution is created.</span></span>

|<span data-ttu-id="fd05a-146">**文件名**</span><span class="sxs-lookup"><span data-stu-id="fd05a-146">**File name**</span></span>|<span data-ttu-id="fd05a-147">**说明**</span><span class="sxs-lookup"><span data-stu-id="fd05a-147">**Description**</span></span>|
|:-----|:-----|
|<span data-ttu-id="fd05a-148">**Home.html**</span><span class="sxs-lookup"><span data-stu-id="fd05a-148">**Home.html**</span></span><br/><span data-ttu-id="fd05a-149">（Excel、PowerPoint、Word）</span><span class="sxs-lookup"><span data-stu-id="fd05a-149">(Excel, PowerPoint, Word)</span></span><br/><br/><span data-ttu-id="fd05a-150">**MessageRead.html**</span><span class="sxs-lookup"><span data-stu-id="fd05a-150">**MessageRead.html**</span></span><br/><span data-ttu-id="fd05a-151">(Outlook)</span><span class="sxs-lookup"><span data-stu-id="fd05a-151">(Outlook)</span></span>|<span data-ttu-id="fd05a-152">外接程序的默认 HTML 页面。</span><span class="sxs-lookup"><span data-stu-id="fd05a-152">The default HTML page of the add-in.</span></span> <span data-ttu-id="fd05a-153">在文档、电子邮件或约会项目中激活该外接程序时，此页面将显示为外接程序内的第一个页面。</span><span class="sxs-lookup"><span data-stu-id="fd05a-153">This page appears as the first page inside of the add-in when it is activated in a document, email message, or appointment item.</span></span> <span data-ttu-id="fd05a-154">此文件包含入门所需的所有文件引用。</span><span class="sxs-lookup"><span data-stu-id="fd05a-154">This file contains all of the file references that you need to get started.</span></span> <span data-ttu-id="fd05a-155">你可以通过将 HTML 代码添加到此文件来开始开发外接程序。</span><span class="sxs-lookup"><span data-stu-id="fd05a-155">You can start developing your add-in by adding your HTML code to this file.</span></span>|
|<span data-ttu-id="fd05a-156">**Home.js**</span><span class="sxs-lookup"><span data-stu-id="fd05a-156">**Home.js**</span></span><br/><span data-ttu-id="fd05a-157">（Excel、PowerPoint、Word）</span><span class="sxs-lookup"><span data-stu-id="fd05a-157">(Excel, PowerPoint, Word)</span></span><br/><br/><span data-ttu-id="fd05a-158">**MessageRead.js**</span><span class="sxs-lookup"><span data-stu-id="fd05a-158">**MessageRead.js**</span></span><br/><span data-ttu-id="fd05a-159">(Outlook)</span><span class="sxs-lookup"><span data-stu-id="fd05a-159">(Outlook)</span></span>|<span data-ttu-id="fd05a-160">与 **Home.html** 页面（Excel、PowerPoint、Word）或 **MessageRead.html** 页面 (Outlook) 关联的 JavaScript 文件。</span><span class="sxs-lookup"><span data-stu-id="fd05a-160">The JavaScript file associated with the **Home.html** page (Excel, PowerPoint, Word) or the **MessageRead.html** page (Outlook).</span></span> <span data-ttu-id="fd05a-161">此文件应包含特定于 **Home.html** 页面（Excel、PowerPoint、Word）或 **MessageRead.html** 页面 (Outlook) 行为的任何代码。</span><span class="sxs-lookup"><span data-stu-id="fd05a-161">This file should contain any code that is specific to the behavior of the **Home.html** page (Excel, PowerPoint, Word) or the **MessageRead.html** page (Outlook).</span></span> <span data-ttu-id="fd05a-162">此文件包含一些可帮你入门的示例代码。</span><span class="sxs-lookup"><span data-stu-id="fd05a-162">This file contains some example code to get you started.</span></span>|
|<span data-ttu-id="fd05a-163">**Home.css**</span><span class="sxs-lookup"><span data-stu-id="fd05a-163">**Home.css**</span></span><br/><span data-ttu-id="fd05a-164">（Excel、PowerPoint、Word）</span><span class="sxs-lookup"><span data-stu-id="fd05a-164">(Excel, PowerPoint, Word)</span></span><br/><br/><span data-ttu-id="fd05a-165">**MessageRead.css**</span><span class="sxs-lookup"><span data-stu-id="fd05a-165">**MessageRead.css**</span></span><br/><span data-ttu-id="fd05a-166">(Outlook)</span><span class="sxs-lookup"><span data-stu-id="fd05a-166">(Outlook)</span></span>|<span data-ttu-id="fd05a-167">定义要应用于外接程序的默认样式。</span><span class="sxs-lookup"><span data-stu-id="fd05a-167">Defines the default styles to apply to your add-in.</span></span> <span data-ttu-id="fd05a-168">我们建议对设计和样式使用 Office UI Fabric。</span><span class="sxs-lookup"><span data-stu-id="fd05a-168">We recommend using the Office UI Fabric for design and styles.</span></span> <span data-ttu-id="fd05a-169">有关详细信息，请参阅 [Office 外接程序中的 Office UI Fabric](../design/office-ui-fabric.md)。</span><span class="sxs-lookup"><span data-stu-id="fd05a-169">For more information see [Office UI Fabric in Office Add-ins](../design/office-ui-fabric.md).</span></span>|

> [!NOTE]
> <span data-ttu-id="fd05a-170">你无需使用这些文件。</span><span class="sxs-lookup"><span data-stu-id="fd05a-170">You don't have to use these files.</span></span> <span data-ttu-id="fd05a-171">你可以随意将其他文件添加到项目并改为使用这些文件。</span><span class="sxs-lookup"><span data-stu-id="fd05a-171">Feel free to add other files to the project and use those instead.</span></span> <span data-ttu-id="fd05a-172">如果要将另一个 HTML 文件显示为外接程序的初始页面，请打开清单编辑器，然后将 **SourceLocation** 属性设置为该文件的名称。</span><span class="sxs-lookup"><span data-stu-id="fd05a-172">If you want another HTML file to appear as the initial page of the add-in, open the manifest editor, and then set the  **SourceLocation** property to the name of the file.</span></span>

## <a name="debug-your-add-in"></a><span data-ttu-id="fd05a-173">调试外接程序</span><span class="sxs-lookup"><span data-stu-id="fd05a-173">Debug your add-in</span></span>

<span data-ttu-id="fd05a-174">你可以使用 Visual Studio 在 Windows 上的 Office 桌面客户端中调试外接程序，如以下部分所述：</span><span class="sxs-lookup"><span data-stu-id="fd05a-174">You can use Visual Studio to debug your add-in in the Office desktop client on Windows, as described in the following sections:</span></span>

- [<span data-ttu-id="fd05a-175">查看生成和调试属性</span><span class="sxs-lookup"><span data-stu-id="fd05a-175">Review the build and debug properties</span></span>](#review-the-build-and-debug-properties)
- [<span data-ttu-id="fd05a-176">使用现有文档调试外接程序</span><span class="sxs-lookup"><span data-stu-id="fd05a-176">Use an existing document to debug the add-in</span></span>](#use-an-existing-document-to-debug-the-add-in)
- [<span data-ttu-id="fd05a-177">启动项目</span><span class="sxs-lookup"><span data-stu-id="fd05a-177">Start the project</span></span>](#start-the-project)
- [<span data-ttu-id="fd05a-178">调试 Excel、PowerPoint 或 Word 外接程序的代码</span><span class="sxs-lookup"><span data-stu-id="fd05a-178">Debug the code for an Excel, PowerPoint, or Word add-in</span></span>](#debug-the-code-for-an-excel-powerpoint-or-word-add-in)
- [<span data-ttu-id="fd05a-179">调试 Outlook 外接程序的代码</span><span class="sxs-lookup"><span data-stu-id="fd05a-179">Debug the code for an Outlook add-in</span></span>](#debug-the-code-for-an-outlook-add-in)

> [!NOTE]
> <span data-ttu-id="fd05a-180">无法使用 Visual Studio 在 Office 网页版或 Mac 版 Office 中调试加载项。</span><span class="sxs-lookup"><span data-stu-id="fd05a-180">You cannot use Visual Studio to debug Office Add-ins in Office Online or Office for Mac.</span></span> <span data-ttu-id="fd05a-181">若要了解如何在这些平台上进行调试，请参阅[在 Office 网页版中调试 Office 加载项](../testing/debug-add-ins-in-office-online.md)，或[调试 iPad 版和 Mac 版 Office 加载项](../testing/debug-office-add-ins-on-ipad-and-mac.md)</span><span class="sxs-lookup"><span data-stu-id="fd05a-181">For information about debugging on these platforms, see [Debug Office Add-ins in Office Online](../testing/debug-add-ins-in-office-online.md) or [Debug Office Add-ins on iPad and Mac](../testing/debug-office-add-ins-on-ipad-and-mac.md)</span></span>

### <a name="review-the-build-and-debug-properties"></a><span data-ttu-id="fd05a-182">查看生成和调试属性</span><span class="sxs-lookup"><span data-stu-id="fd05a-182">Review the build and debug properties</span></span>

<span data-ttu-id="fd05a-183">在开始调试之前，请检查每个项目的属性以确认 Visual Studio 将打开所需的主机应用程序，并已正确设置其他生成和调试属性。</span><span class="sxs-lookup"><span data-stu-id="fd05a-183">Before you start debugging, review the properties of each project to confirm that Visual Studio will open the desired host application and that other build and debug properties are set appropriately.</span></span>

#### <a name="add-in-project-properties"></a><span data-ttu-id="fd05a-184">外接程序项目属性</span><span class="sxs-lookup"><span data-stu-id="fd05a-184">Add-in project properties</span></span>

<span data-ttu-id="fd05a-185">打开外接程序项目的“**属性**”窗口以查看项目属性：</span><span class="sxs-lookup"><span data-stu-id="fd05a-185">Open the **Properties** window for the add-in project to review project properties:</span></span>

1. <span data-ttu-id="fd05a-186">在“**解决方案资源管理器**”中，选择外接程序项目（*而不是* Web 应用程序项目）。</span><span class="sxs-lookup"><span data-stu-id="fd05a-186">In  **Solution Explorer**, choose the add-in project (*not* the web application project).</span></span>

2. <span data-ttu-id="fd05a-187">在菜单栏中，依次选择“**视图**” >  “**属性窗口**”。</span><span class="sxs-lookup"><span data-stu-id="fd05a-187">From the menu bar, choose  **View** >  **Properties Window**.</span></span>

<span data-ttu-id="fd05a-188">下表介绍了外接程序项目的属性。</span><span class="sxs-lookup"><span data-stu-id="fd05a-188">The following table describes the properties of the add-in project.</span></span>

|<span data-ttu-id="fd05a-189">**属性**</span><span class="sxs-lookup"><span data-stu-id="fd05a-189">**Property**</span></span>|<span data-ttu-id="fd05a-190">**说明**</span><span class="sxs-lookup"><span data-stu-id="fd05a-190">**Description**</span></span>|
|:-----|:-----|
|<span data-ttu-id="fd05a-191">**启动操作**</span><span class="sxs-lookup"><span data-stu-id="fd05a-191">**Start Action**</span></span>|<span data-ttu-id="fd05a-192">指定外接程序的调试模式。</span><span class="sxs-lookup"><span data-stu-id="fd05a-192">Specifies the debug mode for your add-in.</span></span> <span data-ttu-id="fd05a-193">目前，Office 外接程序项目仅支持 **Office 桌面客户端**模式。</span><span class="sxs-lookup"><span data-stu-id="fd05a-193">Currently only **Office Desktop Client** mode is supported for Office Add-in projects.</span></span>|
|<span data-ttu-id="fd05a-194">**启动文档**</span><span class="sxs-lookup"><span data-stu-id="fd05a-194">**Start Document**</span></span><br/><span data-ttu-id="fd05a-195">（仅限 Excel、PowerPoint 和 Word 外接程序）</span><span class="sxs-lookup"><span data-stu-id="fd05a-195">(Excel, PowerPoint, and Word add-ins only)</span></span>|<span data-ttu-id="fd05a-196">指定要在启动项目时打开的文档。</span><span class="sxs-lookup"><span data-stu-id="fd05a-196">Specifies what document to open when you start the project.</span></span>|
|<span data-ttu-id="fd05a-197">**Web 项目**</span><span class="sxs-lookup"><span data-stu-id="fd05a-197">**Web Project**</span></span>|<span data-ttu-id="fd05a-198">指定与外接程序关联的 Web 项目的名称。</span><span class="sxs-lookup"><span data-stu-id="fd05a-198">Specifies the name of the web project associated with the add-in.</span></span>|
|<span data-ttu-id="fd05a-199">**电子邮件地址**</span><span class="sxs-lookup"><span data-stu-id="fd05a-199">**Email Address**</span></span><br/><span data-ttu-id="fd05a-200">（仅限 Outlook 外接程序）</span><span class="sxs-lookup"><span data-stu-id="fd05a-200">(Outlook add-ins only)</span></span>|<span data-ttu-id="fd05a-201">指定你想在 Exchange Server 或 Exchange Online 中用来测试 Outlook 外接程序的用户帐户的电子邮件地址。</span><span class="sxs-lookup"><span data-stu-id="fd05a-201">Specifies the email address of the user account in Exchange Server or Exchange Online that you want to use to test your Outlook add-in.</span></span>|
|<span data-ttu-id="fd05a-202">**EWS Url**</span><span class="sxs-lookup"><span data-stu-id="fd05a-202">**EWS Url**</span></span><br/><span data-ttu-id="fd05a-203">（仅限 Outlook 外接程序）</span><span class="sxs-lookup"><span data-stu-id="fd05a-203">(Outlook add-ins only)</span></span>|<span data-ttu-id="fd05a-204">Exchange Web 服务 URL（例如：`https://www.contoso.com/ews/exchange.aspx`）。</span><span class="sxs-lookup"><span data-stu-id="fd05a-204">Exchange Web service URL (For example: `https://www.contoso.com/ews/exchange.aspx`).</span></span> |
|<span data-ttu-id="fd05a-205">**OWA Url**</span><span class="sxs-lookup"><span data-stu-id="fd05a-205">**OWA Url**</span></span><br/><span data-ttu-id="fd05a-206">（仅限 Outlook 外接程序）</span><span class="sxs-lookup"><span data-stu-id="fd05a-206">(Outlook add-ins only)</span></span>|<span data-ttu-id="fd05a-207">Outlook 网页版 URL（例如：`https://www.contoso.com/owa`）。</span><span class="sxs-lookup"><span data-stu-id="fd05a-207">Outlook Web App URL (For example: `https://www.contoso.com/owa`).</span></span>|
|<span data-ttu-id="fd05a-208">**使用多重身份验证**</span><span class="sxs-lookup"><span data-stu-id="fd05a-208">**Use multi-factor auth**</span></span><br/><span data-ttu-id="fd05a-209">（仅限 Outlook 加载项）</span><span class="sxs-lookup"><span data-stu-id="fd05a-209">(Outlook add-ins only)</span></span>|<span data-ttu-id="fd05a-210">布尔值，指示是否应使用多重身份验证。</span><span class="sxs-lookup"><span data-stu-id="fd05a-210">Boolean value that indicates whether multi-factor authentication should be used.</span></span>|
|<span data-ttu-id="fd05a-211">**用户名**</span><span class="sxs-lookup"><span data-stu-id="fd05a-211">**User Name**</span></span><br/><span data-ttu-id="fd05a-212">（仅限 Outlook 外接程序）</span><span class="sxs-lookup"><span data-stu-id="fd05a-212">(Outlook add-ins only)</span></span>|<span data-ttu-id="fd05a-213">指定你想在 Exchange Server 或 Exchange Online 中用来测试 Outlook 外接程序的用户帐户的名称。</span><span class="sxs-lookup"><span data-stu-id="fd05a-213">Specifies the name of the user account in Exchange Server or Exchange Online that you want to use to test your Outlook add-in.</span></span>|
|<span data-ttu-id="fd05a-214">**项目文件**</span><span class="sxs-lookup"><span data-stu-id="fd05a-214">**Project File**</span></span>|<span data-ttu-id="fd05a-215">指定包含生成、配置和有关项目的其他信息的文件名称。</span><span class="sxs-lookup"><span data-stu-id="fd05a-215">Specifies the name of the file containing build, configuration, and other information about the project.</span></span>|
|<span data-ttu-id="fd05a-216">**项目文件夹**</span><span class="sxs-lookup"><span data-stu-id="fd05a-216">**Project Folder**</span></span>|<span data-ttu-id="fd05a-217">项目文件的位置。</span><span class="sxs-lookup"><span data-stu-id="fd05a-217">The location of the project file.</span></span>|

> [!NOTE]
> <span data-ttu-id="fd05a-218">对于 Outlook 外接程序，你可以选择在“**属性**”窗口中为一个或多个 *Outlook 外接程序*属性指定值，但这样做并不是必须的。</span><span class="sxs-lookup"><span data-stu-id="fd05a-218">For an Outlook add-in, you may choose to specify values for one or more of the *Outlook add-in only* properties in the **Properties** window, but doing so is not required.</span></span>

#### <a name="web-application-project-properties"></a><span data-ttu-id="fd05a-219">Web 应用程序项目属性</span><span class="sxs-lookup"><span data-stu-id="fd05a-219">Web application project properties</span></span>

<span data-ttu-id="fd05a-220">打开 Web 应用程序项目的“**属性**”窗口以查看项目属性：</span><span class="sxs-lookup"><span data-stu-id="fd05a-220">Open the **Properties** window for the web application project to review project properties:</span></span>

1. <span data-ttu-id="fd05a-221">在“**解决方案资源管理器**”中，选择 Web 应用程序项目。</span><span class="sxs-lookup"><span data-stu-id="fd05a-221">In  **Solution Explorer**, choose the web application project.</span></span>

2. <span data-ttu-id="fd05a-222">在菜单栏中，依次选择“**视图**” >  “**属性窗口**”。</span><span class="sxs-lookup"><span data-stu-id="fd05a-222">From the menu bar, choose  **View** >  **Properties Window**.</span></span>

<span data-ttu-id="fd05a-223">下表介绍了与 Office 外接程序项目最相关的 Web 应用程序项目的属性。</span><span class="sxs-lookup"><span data-stu-id="fd05a-223">The following table describes the properties of the web application project that are most relevant to Office Add-in projects.</span></span>

|<span data-ttu-id="fd05a-224">**属性**</span><span class="sxs-lookup"><span data-stu-id="fd05a-224">**Property**</span></span>|<span data-ttu-id="fd05a-225">**说明**</span><span class="sxs-lookup"><span data-stu-id="fd05a-225">**Description**</span></span>|
|:-----|:-----|
|<span data-ttu-id="fd05a-226">**SSL 已启用**</span><span class="sxs-lookup"><span data-stu-id="fd05a-226">**SSL Enabled**</span></span>|<span data-ttu-id="fd05a-227">指定是否在站点上启用 SSL。</span><span class="sxs-lookup"><span data-stu-id="fd05a-227">Specifies whether SSL is enabled on the site.</span></span> <span data-ttu-id="fd05a-228">对于 Office 外接程序项目，此属性应设置为 **True**。</span><span class="sxs-lookup"><span data-stu-id="fd05a-228">This property should be set to **True** for Office Add-in projects.</span></span>|
|<span data-ttu-id="fd05a-229">**SSL URL**</span><span class="sxs-lookup"><span data-stu-id="fd05a-229">**SSL URL**</span></span>|<span data-ttu-id="fd05a-230">指定站点的安全 HTTPS URL。</span><span class="sxs-lookup"><span data-stu-id="fd05a-230">Specifies the secure HTTPS URL for the site.</span></span> <span data-ttu-id="fd05a-231">只读。</span><span class="sxs-lookup"><span data-stu-id="fd05a-231">Read-only.</span></span>|
|<span data-ttu-id="fd05a-232">**URL**</span><span class="sxs-lookup"><span data-stu-id="fd05a-232">**URL**</span></span>|<span data-ttu-id="fd05a-233">指定站点的 HTTP URL。</span><span class="sxs-lookup"><span data-stu-id="fd05a-233">Specifies the HTTP URL for the site.</span></span> <span data-ttu-id="fd05a-234">只读。</span><span class="sxs-lookup"><span data-stu-id="fd05a-234">Read-only.</span></span>|
|<span data-ttu-id="fd05a-235">**项目文件**</span><span class="sxs-lookup"><span data-stu-id="fd05a-235">**Project File**</span></span>|<span data-ttu-id="fd05a-236">指定包含生成、配置和有关项目的其他信息的文件名称。</span><span class="sxs-lookup"><span data-stu-id="fd05a-236">Specifies the name of the file containing build, configuration, and other information about the project.</span></span>|
|<span data-ttu-id="fd05a-237">**项目文件夹**</span><span class="sxs-lookup"><span data-stu-id="fd05a-237">**Project Folder**</span></span>|<span data-ttu-id="fd05a-238">指定项目文件的位置。</span><span class="sxs-lookup"><span data-stu-id="fd05a-238">Specifies the location of the project file.</span></span> <span data-ttu-id="fd05a-239">只读。</span><span class="sxs-lookup"><span data-stu-id="fd05a-239">Read-only.</span></span> <span data-ttu-id="fd05a-240">Visual Studio 在运行时生成的清单文件将写入到此位置的 `bin\Debug\OfficeAppManifests` 文件夹中。</span><span class="sxs-lookup"><span data-stu-id="fd05a-240">The manifest file that Visual Studio generates at runtime is written to the `bin\Debug\OfficeAppManifests` folder in this location.</span></span>|

### <a name="use-an-existing-document-to-debug-the-add-in"></a><span data-ttu-id="fd05a-241">使用现有文档调试外接程序</span><span class="sxs-lookup"><span data-stu-id="fd05a-241">Use an existing document to debug the add-in</span></span>

<span data-ttu-id="fd05a-242">如果你有一个文档包含要在调试 Excel、PowerPoint 或 Word 外接程序时使用的测试数据，则可以将 Visual Studio 配置为在启动项目时打开该文档。</span><span class="sxs-lookup"><span data-stu-id="fd05a-242">If you have a document that contains test data you want to use while debugging your Excel, PowerPoint, or Word add-in, Visual Studio can be configured to open that document when you start the project.</span></span> <span data-ttu-id="fd05a-243">若要指定在调试外接程序时要使用的现有文档，请完成以下步骤。</span><span class="sxs-lookup"><span data-stu-id="fd05a-243">To specify an existing document to use while debugging the add-in, complete the following steps.</span></span>

1. <span data-ttu-id="fd05a-244">在“**解决方案资源管理器**”中，选择外接程序项目（*而不是* Web 应用程序项目）。</span><span class="sxs-lookup"><span data-stu-id="fd05a-244">In **Solution Explorer**, choose the add-in project (*not* the web application project).</span></span>

2. <span data-ttu-id="fd05a-245">从菜单栏中，选择“**项目**” > “**添加现有项**”。</span><span class="sxs-lookup"><span data-stu-id="fd05a-245">From the menu bar, choose **Project** > **Add Existing Item**.</span></span>

3. <span data-ttu-id="fd05a-246">在“**添加现有项**”对话框中，找到并选择要添加的文档。</span><span class="sxs-lookup"><span data-stu-id="fd05a-246">In the **Add Existing Item** dialog box, locate and select the document that you want to add.</span></span>

4. <span data-ttu-id="fd05a-247">选择“**添加**”按钮以将文档添加到项目中。</span><span class="sxs-lookup"><span data-stu-id="fd05a-247">Choose the **Add** button to add the document to your project.</span></span>

5. <span data-ttu-id="fd05a-248">在“**解决方案资源管理器**”中，选择外接程序项目（*而不是* Web 应用程序项目）。</span><span class="sxs-lookup"><span data-stu-id="fd05a-248">In **Solution Explorer**, choose the add-in project (*not* the web application project).</span></span>

6. <span data-ttu-id="fd05a-249">在菜单栏中，依次选择“**视图**” > “**属性窗口**”。</span><span class="sxs-lookup"><span data-stu-id="fd05a-249">From the menu bar, choose **View** > **Properties Window**.</span></span>

7. <span data-ttu-id="fd05a-250">在“**属性**”窗口中，选择“**启动文档**”列表，然后选择添加到项目中的文档。</span><span class="sxs-lookup"><span data-stu-id="fd05a-250">In the **Properties** window, choose the **Start Document** list, and then select the document that you added to the project.</span></span> <span data-ttu-id="fd05a-251">该项目现在配置为在该文档中启动外接程序。</span><span class="sxs-lookup"><span data-stu-id="fd05a-251">The project is now configured to start the add-in in that document.</span></span>

### <a name="start-the-project"></a><span data-ttu-id="fd05a-252">启动项目</span><span class="sxs-lookup"><span data-stu-id="fd05a-252">Start the project</span></span>

<span data-ttu-id="fd05a-253">从菜单栏中依次选择“**调试**” > “**开始调试**”，可启动项目。</span><span class="sxs-lookup"><span data-stu-id="fd05a-253">Start the project by choosing **Debug** > **Start Debugging** from the menu bar.</span></span> <span data-ttu-id="fd05a-254">Visual Studio 将自动生成解决方案并启动 Office 以托管外接程序。</span><span class="sxs-lookup"><span data-stu-id="fd05a-254">Visual Studio will automatically build the solution and start Office to host your add-in.</span></span>

> [!NOTE]
> <span data-ttu-id="fd05a-255">启动 Outlook 外接程序项目时，系统会提示你输入登录凭据。</span><span class="sxs-lookup"><span data-stu-id="fd05a-255">When you start an Outlook add-in project, you'll be prompted for login credentials.</span></span> <span data-ttu-id="fd05a-256">如果系统要求你重复登录，或者如果收到指示未经授权的错误，则可能会禁用 Office 365 租户上帐户的基本身份验证。</span><span class="sxs-lookup"><span data-stu-id="fd05a-256">If you're asked to log in repeatedly or if you receive an error that you are unauthorized, then Basic Auth may be disabled for accounts on your Office 365 tenant.</span></span> <span data-ttu-id="fd05a-257">在这种情况下，请尝试使用 Microsoft 帐户。</span><span class="sxs-lookup"><span data-stu-id="fd05a-257">In this case, try using a Microsoft account instead.</span></span> <span data-ttu-id="fd05a-258">可能还需要在“Outlook Web 加载项”项目属性对话框中将属性“使用多重身份验证”设置为 True。</span><span class="sxs-lookup"><span data-stu-id="fd05a-258">You may also need to set the property "Use multi-factor auth" to True in the Outlook Web Add-in project properties dialog.</span></span>

<span data-ttu-id="fd05a-259">当 Visual Studio 生成项目时，它执行以下任务：</span><span class="sxs-lookup"><span data-stu-id="fd05a-259">When Visual Studio builds the project it performs the following tasks:</span></span>

1. <span data-ttu-id="fd05a-260">创建 XML 清单文件的副本并将其添加到 `_ProjectName_\bin\Debug\OfficeAppManifests` 目录。</span><span class="sxs-lookup"><span data-stu-id="fd05a-260">Creates a copy of the XML manifest file and adds it to  `_ProjectName_\bin\Debug\OfficeAppManifests` directory.</span></span> <span data-ttu-id="fd05a-261">启动 Visual Studio 并调试外接程序时，主机应用程序将使用此副本。</span><span class="sxs-lookup"><span data-stu-id="fd05a-261">The host application consumes this copy when you start Visual Studio and debug the add-in.</span></span>

2. <span data-ttu-id="fd05a-262">在计算机上创建一组允许外接程序在主机应用程序中显示的注册表项。</span><span class="sxs-lookup"><span data-stu-id="fd05a-262">Creates a set of registry entries on your computer that enable the add-in to appear in the host application.</span></span>

3. <span data-ttu-id="fd05a-263">生成 Web 应用程序项目，然后将其部署到本地 IIS Web 服务器 (https://localhost))。</span><span class="sxs-lookup"><span data-stu-id="fd05a-263">Builds the web application project, and then deploys it to the local IIS web server (https://localhost).</span></span>

4. <span data-ttu-id="fd05a-264">如果这是你已部署到本地 IIS Web 服务器的第一个加载项项目，系统可能会提示你将自签名证书安装到当前用户的受信任的根证书存储中。</span><span class="sxs-lookup"><span data-stu-id="fd05a-264">If this is the first add-in project that you have deployed to local IIS web server, you may be prompted to install a Self-Signed Certificate to the current user's Trusted Root Certificate store.</span></span> <span data-ttu-id="fd05a-265">若要使 IIS Express 正确显示加载项内容，这是必需的操作。</span><span class="sxs-lookup"><span data-stu-id="fd05a-265">This is required for IIS Express to display the content of your add-in correctly.</span></span>


> [!NOTE]
> <span data-ttu-id="fd05a-266">在 Windows 10 上运行时，最新版本的 Office 可能会使用较新的 Web 控件来显示加载项内容。</span><span class="sxs-lookup"><span data-stu-id="fd05a-266">The latest version of Office may use a newer web control to display the add-in contents when running on Windows 10.</span></span> <span data-ttu-id="fd05a-267">如果是这种情况，Visual Studio 可能会提示你添加本地网络环回豁免。</span><span class="sxs-lookup"><span data-stu-id="fd05a-267">If this is the case, Visual Studio may prompt you to add a local network loopback exemption.</span></span> <span data-ttu-id="fd05a-268">在 Office 主机应用程序中，需要这样做才能使 Web 控件访问部署到本地 IIS Web 服务器的网站。</span><span class="sxs-lookup"><span data-stu-id="fd05a-268">This is required for the web control, in the Office host application, to be able to access the website deployed to the local IIS web server.</span></span> <span data-ttu-id="fd05a-269">还可以在 Visual Studio 中的“工具” > “选项” > “Office 工具(Web)” > “Web 加载项调试”下随时更改此设置\*\*\*\*\*\*\*\*\*\*\*\*\*\*\*\*。</span><span class="sxs-lookup"><span data-stu-id="fd05a-269">You can also change this setting anytime in Visual Studio under **Tools** > **Options** > **Office Tools (Web)** > **Web Add-In Debugging**.</span></span>


<span data-ttu-id="fd05a-270">接下来，Visual Studio 会执行以下操作：</span><span class="sxs-lookup"><span data-stu-id="fd05a-270">Next, Visual Studio does the following:</span></span>

1. <span data-ttu-id="fd05a-271">通过将 `~remoteAppUrl` 标记替换为起始页的完全限定地址（例如，`https://localhost:44302/Home.html`）来修改 XML 清单文件的 [SourceLocation](/office/dev/add-ins/reference/manifest/sourcelocation) 元素。</span><span class="sxs-lookup"><span data-stu-id="fd05a-271">Modifies the [SourceLocation](/office/dev/add-ins/reference/manifest/sourcelocation) element of the XML manifest file by replacing the `~remoteAppUrl` token with the fully qualified address of the start page (for example, `https://localhost:44302/Home.html`).</span></span>

2. <span data-ttu-id="fd05a-272">在 IIS Express 中启动 Web 应用程序项目。</span><span class="sxs-lookup"><span data-stu-id="fd05a-272">Starts the web application project in IIS Express.</span></span>

3. <span data-ttu-id="fd05a-273">打开主机应用程序。</span><span class="sxs-lookup"><span data-stu-id="fd05a-273">Opens the host application.</span></span>

<span data-ttu-id="fd05a-274">生成项目时，Visual Studio 不会显示“**输出**”窗口中的验证错误。</span><span class="sxs-lookup"><span data-stu-id="fd05a-274">Visual Studio doesn't show validation errors in the  **OUTPUT** window when you build the project.</span></span> <span data-ttu-id="fd05a-275">Visual Studio 报告“**错误列表**”窗口中出现的错误和警告。</span><span class="sxs-lookup"><span data-stu-id="fd05a-275">Visual Studio reports errors and warnings in the **ERRORLIST** window as they occur.</span></span> <span data-ttu-id="fd05a-276">通过在代码和文本编辑器中显示不同颜色的波浪下划线（称为波浪线），Visual Studio 还报告验证错误。</span><span class="sxs-lookup"><span data-stu-id="fd05a-276">Visual Studio also reports validation errors by showing wavy underlines (known as squiggles) of different colors in the code and text editor.</span></span> <span data-ttu-id="fd05a-277">通过这些标志，你可以得知 Visual Studio 在你的代码中检测到的问题。</span><span class="sxs-lookup"><span data-stu-id="fd05a-277">These marks notify you of problems that Visual Studio detected in your code.</span></span> <span data-ttu-id="fd05a-278">有关详细信息，请参阅[代码和文本编辑器](https://msdn.microsoft.com/library/se2f663y(v=vs.140).aspx)。</span><span class="sxs-lookup"><span data-stu-id="fd05a-278">For more information, see [Code and Text Editor](https://msdn.microsoft.com/library/se2f663y(v=vs.140).aspx).</span></span> <span data-ttu-id="fd05a-279">有关如何启用或禁用验证的详细信息，请参阅[选项、文本编辑器、JavaScript、IntelliSense](/visualstudio/ide/reference/options-text-editor-javascript-intellisense?view=vs-2017)。</span><span class="sxs-lookup"><span data-stu-id="fd05a-279">For more information about how to enable or disable validation, see [Options, Text Editor, JavaScript, IntelliSense](/visualstudio/ide/reference/options-text-editor-javascript-intellisense?view=vs-2017).</span></span>

<span data-ttu-id="fd05a-280">要查看项目中 XML 清单文件的验证规则，请参阅 [Office 外接程序 XML 清单](../develop/add-in-manifests.md)。</span><span class="sxs-lookup"><span data-stu-id="fd05a-280">To review the validation rules of the XML manifest file in your project, see [Office Add-ins XML manifest](../develop/add-in-manifests.md).</span></span>

### <a name="debug-the-code-for-an-excel-powerpoint-or-word-add-in"></a><span data-ttu-id="fd05a-281">调试 Excel、PowerPoint 或 Word 外接程序的代码</span><span class="sxs-lookup"><span data-stu-id="fd05a-281">Debug the code for an Excel, PowerPoint, or Word add-in</span></span>

<span data-ttu-id="fd05a-282">如果在[启动项目](#start-the-project)后，在主机应用程序（Excel、PowerPoint 或 Word）中显示的文档中看不到外接程序，请在主机应用程序中手动启动外接程序。</span><span class="sxs-lookup"><span data-stu-id="fd05a-282">If your add-in isn't visible within the document that's displayed in the host application (Excel, PowerPoint, or Word) after you've [started the project](#start-the-project), manually launch the add-in in the host application.</span></span> <span data-ttu-id="fd05a-283">例如，通过选择“**主页**”选项卡功能区中的“**显示任务窗格**”按钮来启动任务窗格外接程序。在 Excel、PowerPoint 或 Word 中显示外接程序后，你可以通过执行以下操作来调试代码：</span><span class="sxs-lookup"><span data-stu-id="fd05a-283">For example, launch your task pane add-in by choosing the **Show Taskpane** button in the ribbon of the **Home** tab. After your add-in is displayed in Excel, PowerPoint, or Word, you can debug your code by doing the following:</span></span>

1. <span data-ttu-id="fd05a-284">在 Excel、PowerPoint 或 Word 中，选择“**插入**”选项卡，然后选择“**我的外接程序**”右侧的向下箭头。</span><span class="sxs-lookup"><span data-stu-id="fd05a-284">In Excel, PowerPoint, or Word, choose the **Insert** tab and then choose the down-arrow located to the right of **My Add-ins**.</span></span>

    ![Windows 版 Excel 的“插入”功能区及突出显示的“我的加载项”箭头](../images/excel-cf-register-add-in-1b.png)

2. <span data-ttu-id="fd05a-286">在可用外接程序列表中，找到“**开发人员外接程序**”部分并选择你的外接程序进行注册。</span><span class="sxs-lookup"><span data-stu-id="fd05a-286">In the list of available add-ins, find the **Developer Add-ins** section and select the your add-in to register it.</span></span>

3. <span data-ttu-id="fd05a-287">在 Visual Studio 中，在代码中设置断点。</span><span class="sxs-lookup"><span data-stu-id="fd05a-287">In Visual Studio, set breakpoints in your code.</span></span>

4. <span data-ttu-id="fd05a-288">在 Excel、PowerPoint 或 Word 中，与外接程序进行交互。</span><span class="sxs-lookup"><span data-stu-id="fd05a-288">In Excel, PowerPoint, or Word, interact with your add-in.</span></span>

5. <span data-ttu-id="fd05a-289">在 Visual Studio 中命中断点时，根据需要逐步执行代码。</span><span class="sxs-lookup"><span data-stu-id="fd05a-289">As breakpoints are hit in Visual Studio, step through the code as needed.</span></span>

<span data-ttu-id="fd05a-290">你可以更改代码并在外接程序中查看这些更改的效果，而无需关闭主机应用程序并重新启动该项目。</span><span class="sxs-lookup"><span data-stu-id="fd05a-290">You can change your code and review the effects of those changes in your add-in without having to close the host application and restart the project.</span></span> <span data-ttu-id="fd05a-291">保存对代码的更改后，只需在主机应用程序中重新加载外接程序。</span><span class="sxs-lookup"><span data-stu-id="fd05a-291">After you save changes to your code, simply reload the add-in in the host application.</span></span> <span data-ttu-id="fd05a-292">例如，通过选择任务窗格的右上角来激活[个性菜单](../design/task-pane-add-ins.md#personality-menu)，然后选择“**重新加载**”，便可重新加载任务窗格外接程序。</span><span class="sxs-lookup"><span data-stu-id="fd05a-292">For example, reload a task pane add-in by choosing the top-right corner of the task pane to activate the [personality menu](../design/task-pane-add-ins.md#personality-menu) and then choose **Reload**.</span></span>

### <a name="debug-the-code-for-an-outlook-add-in"></a><span data-ttu-id="fd05a-293">调试 Outlook 外接程序的代码</span><span class="sxs-lookup"><span data-stu-id="fd05a-293">Debug the code for an Outlook add-in</span></span>

<span data-ttu-id="fd05a-294">在你已[启动项目](#start-the-project)，且 Visual Studio 启动 Outlook 来托管外接程序后，打开电子邮件或约会项目。</span><span class="sxs-lookup"><span data-stu-id="fd05a-294">After you've [started the project](#start-the-project) and Visual Studio launches Outlook to host your add-in, open an email message or appointment item.</span></span> 

<span data-ttu-id="fd05a-p126">只要满足激活条件，Outlook 便会为项目激活外接程序。外接程序栏显示在"检查器"窗口或阅读窗格的顶部，Outlook 外接程序显示为外接程序栏中的一个按钮。如果您的外接程序有外接程序命令，那么在默认选项卡或指定的自定义选项卡中将有一个按钮显示在功能区中，而该外接程序将不会显示在外接程序栏中。</span><span class="sxs-lookup"><span data-stu-id="fd05a-p126">Outlook activates the add-in for the item as long as the activation criteria are met. The add-in bar appears at the top of the Inspector window or Reading Pane, and your Outlook add-in appears as a button in the add-in bar. If your add-in has an add-in command, a button will appear in the ribbon, either in the default tab or a specified custom tab, and the add-in will not appear in the add-in bar.</span></span>

<span data-ttu-id="fd05a-298">若要查看 Outlook 外接程序，请选择对应 Outlook 外接程序的按钮。</span><span class="sxs-lookup"><span data-stu-id="fd05a-298">To view your Outlook add-in, choose the button for your Outlook add-in.</span></span> <span data-ttu-id="fd05a-299">在 Outlook 中显示外接程序后，你可以通过执行以下操作来调试代码：</span><span class="sxs-lookup"><span data-stu-id="fd05a-299">After your add-in is displayed in Outlook, you can debug your code by doing the following:</span></span>

1. <span data-ttu-id="fd05a-300">在 Visual Studio 中，在代码中设置断点。</span><span class="sxs-lookup"><span data-stu-id="fd05a-300">In Visual Studio, set breakpoints in your code.</span></span>

2. <span data-ttu-id="fd05a-301">在 Outlook 中，与外接程序进行交互。</span><span class="sxs-lookup"><span data-stu-id="fd05a-301">In Outlook, interact with your add-in.</span></span>

3. <span data-ttu-id="fd05a-302">在 Visual Studio 中命中断点时，根据需要逐步执行代码。</span><span class="sxs-lookup"><span data-stu-id="fd05a-302">As breakpoints are hit in Visual Studio, step through the code as needed.</span></span>

<span data-ttu-id="fd05a-303">你可以更改代码并在外接程序中查看这些更改的效果，而无需关闭 Outlook 并重新启动该项目。</span><span class="sxs-lookup"><span data-stu-id="fd05a-303">You can change your code and review the effects of those changes in your add-in without having to close Outlook and restart the project.</span></span> <span data-ttu-id="fd05a-304">保存对代码的更改后，只需打开外接程序的快捷菜单（在 Outlook 中），然后选择“**重新加载**”。</span><span class="sxs-lookup"><span data-stu-id="fd05a-304">After you save changes to your code, simply open the shortcut menu for the add-in (in Outlook), and then choose **Reload**.</span></span>

## <a name="next-steps"></a><span data-ttu-id="fd05a-305">后续步骤</span><span class="sxs-lookup"><span data-stu-id="fd05a-305">Next steps</span></span>

<span data-ttu-id="fd05a-306">在外接程序正常工作后，请参阅[部署和发布 Office 外接程序](../publish/publish.md)，以了解可用于将外接程序分发给用户的方法。</span><span class="sxs-lookup"><span data-stu-id="fd05a-306">After your add-in is working as desired, see [Deploy and publish your Office Add-in](../publish/publish.md) to learn about the ways you can distribute the add-in to users.</span></span>
