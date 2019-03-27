---
title: Excel、Word 和 PowerPoint 加载项命令
description: ''
ms.date: 03/19/2019
localization_priority: Priority
ms.openlocfilehash: e255e6e517f6292b7e7cb7df7b59476b21306911
ms.sourcegitcommit: a2950492a2337de3180b713f5693fe82dbdd6a17
ms.translationtype: HT
ms.contentlocale: zh-CN
ms.lasthandoff: 03/27/2019
ms.locfileid: "30872247"
---
# <a name="add-in-commands-for-excel-word-and-powerpoint"></a><span data-ttu-id="08796-102">Excel、Word 和 PowerPoint 加载项命令</span><span class="sxs-lookup"><span data-stu-id="08796-102">Add-in commands for Excel, Word, and PowerPoint</span></span>

<span data-ttu-id="08796-p101">外接程序命令是 UI 元素，可扩展 Office UI，并在外接程序中启动操作。使用外接程序命令，可以在功能区上添加按钮，也可以向上下文菜单添加项。当用户选择外接程序命令时，将启动操作，如运行 JavaScript 代码或在任务窗格中显示外接程序页面。外接程序命令可帮助用户查找和使用外接程序，从而提高外接程序的采用率和重用率以及客户保留率。</span><span class="sxs-lookup"><span data-stu-id="08796-p101">Add-in commands are UI elements that extend the Office UI and start actions in your add-in. You can use add-in commands to add a button on the ribbon or an item to a context menu. When users select an add-in command, they initiate actions such as running JavaScript code, or showing a page of the add-in in a task pane. Add-in commands help users find and use your add-in, which can help increase your add-in's adoption and reuse, and improve customer retention.</span></span>

<span data-ttu-id="08796-107">有关此功能的概述，请观看视频 [Office 功能区中的加载项命令](https://channel9.msdn.com/events/Build/2016/P551)。</span><span class="sxs-lookup"><span data-stu-id="08796-107">For an overview of the feature, see the video [Add-in Commands in the Office Ribbon](https://channel9.msdn.com/events/Build/2016/P551).</span></span>

> [!NOTE]
> <span data-ttu-id="08796-p102">SharePoint 目录不支持加载项命令。可以通过[集中部署](../publish/centralized-deployment.md)或 [AppSource](/office/dev/store/submit-to-the-office-store) 部署加载项命令，也可以使用[旁加载](../testing/create-a-network-shared-folder-catalog-for-task-pane-and-content-add-ins.md)部署加载项命令以供测试。</span><span class="sxs-lookup"><span data-stu-id="08796-p102">SharePoint catalogs do not support add-in commands. You can deploy add-in commands via [Centralized Deployment](../publish/centralized-deployment.md) or [AppSource](/office/dev/store/submit-to-the-office-store), or use [sideloading](../testing/create-a-network-shared-folder-catalog-for-task-pane-and-content-add-ins.md) to deploy your add-in command for testing.</span></span> 

<span data-ttu-id="08796-110">*图 1：在 Excel Desktop 中运行命令的加载项*</span><span class="sxs-lookup"><span data-stu-id="08796-110">*Figure 1. Add-in with commands running in Excel Desktop*</span></span>

![Excel 中的加载项命令屏幕截图](../images/add-in-commands-1.png)

<span data-ttu-id="08796-112">*图 2：在 Excel Online 中运行命令的加载项*</span><span class="sxs-lookup"><span data-stu-id="08796-112">*Figure 2. Add-in with commands running in Excel Online*</span></span>

![Excel Online 中的加载项命令屏幕截图](../images/add-in-commands-2.png)

## <a name="command-capabilities"></a><span data-ttu-id="08796-114">命令功能</span><span class="sxs-lookup"><span data-stu-id="08796-114">Command capabilities</span></span>

<span data-ttu-id="08796-115">目前支持下列命令功能。</span><span class="sxs-lookup"><span data-stu-id="08796-115">The following command capabilities are currently supported.</span></span>

> [!NOTE]
> <span data-ttu-id="08796-116">内容加载项暂不支持加载项命令。</span><span class="sxs-lookup"><span data-stu-id="08796-116">Content add-ins do not currently support add-in commands.</span></span>

<span data-ttu-id="08796-117">**扩展点**</span><span class="sxs-lookup"><span data-stu-id="08796-117">**Extension points**</span></span>

- <span data-ttu-id="08796-118">功能区选项卡 - 扩展内置选项卡或新建自定义选项卡。</span><span class="sxs-lookup"><span data-stu-id="08796-118">Ribbon tabs - Extend built-in tabs or create a new custom tab.</span></span>
- <span data-ttu-id="08796-119">上下文菜单 - 扩展选定上下文菜单。</span><span class="sxs-lookup"><span data-stu-id="08796-119">Context menus - Extend selected context menus.</span></span>

<span data-ttu-id="08796-120">**控件类型**</span><span class="sxs-lookup"><span data-stu-id="08796-120">**Control types**</span></span>

- <span data-ttu-id="08796-121">简单按钮 - 触发特定操作。</span><span class="sxs-lookup"><span data-stu-id="08796-121">Simple buttons - trigger specific actions.</span></span>
- <span data-ttu-id="08796-122">菜单 - 简单的下拉菜单，内含可触发操作的按钮。</span><span class="sxs-lookup"><span data-stu-id="08796-122">Menus - simple menu dropdown with buttons that trigger actions.</span></span>

<span data-ttu-id="08796-123">**操作**</span><span class="sxs-lookup"><span data-stu-id="08796-123">**Actions**</span></span>

- <span data-ttu-id="08796-124">ShowTaskpane - 显示一个或多个在其中加载自定义 HTML 页的窗格。</span><span class="sxs-lookup"><span data-stu-id="08796-124">ShowTaskpane - Displays one or multiple panes that load custom HTML pages inside them.</span></span>
- <span data-ttu-id="08796-p103">ExecuteFunction - 加载一个不可见的 HTML 页，然后在其中执行一个 JavaScript 函数。若要在你的函数（例如错误、进度或其他输入）中显示 UI，你可以使用 [displayDialog](/javascript/api/office/office.ui) API。</span><span class="sxs-lookup"><span data-stu-id="08796-p103">ExecuteFunction - Loads an invisible HTML page and then execute a JavaScript function within it. To show UI within your function (such as errors, progress, or additional input) you can use the [displayDialog](/javascript/api/office/office.ui) API.</span></span>  

## <a name="supported-platforms"></a><span data-ttu-id="08796-127">支持的平台</span><span class="sxs-lookup"><span data-stu-id="08796-127">Supported platforms</span></span>

<span data-ttu-id="08796-128">目前，以下平台支持加载项命令：</span><span class="sxs-lookup"><span data-stu-id="08796-128">Add-in commands are currently supported on the following platforms:</span></span>

- <span data-ttu-id="08796-129">Office 2016 for Windows 或更高版本（内部版本 16.0.6769+）</span><span class="sxs-lookup"><span data-stu-id="08796-129">Office 2016 or later for Windows (build 16.0.6769+)</span></span>
- <span data-ttu-id="08796-130">Office for Mac（内部版本 15.33+）</span><span class="sxs-lookup"><span data-stu-id="08796-130">Office for Mac (build 15.33+)</span></span>
- <span data-ttu-id="08796-131">Office Online</span><span class="sxs-lookup"><span data-stu-id="08796-131">Office Online</span></span>

<span data-ttu-id="08796-132">即将推出更多受支持的平台。</span><span class="sxs-lookup"><span data-stu-id="08796-132">More platforms are coming soon.</span></span>

## <a name="best-practices"></a><span data-ttu-id="08796-133">最佳做法</span><span class="sxs-lookup"><span data-stu-id="08796-133">Best practices</span></span>

<span data-ttu-id="08796-134">在开发外接程序命令时应用下面的最佳做法：</span><span class="sxs-lookup"><span data-stu-id="08796-134">Apply the following best practices when you develop add-in commands:</span></span>

- <span data-ttu-id="08796-p104">使用命令来表示会给用户带来明确具体结果的特定操作。不要在单个按钮中组合多个操作。</span><span class="sxs-lookup"><span data-stu-id="08796-p104">Use commands to represent a specific action with a clear and specific outcome for users. Do not combine multiple actions in a single button.</span></span>
- <span data-ttu-id="08796-p105">提供使您的外接程序中的常见任务执行效率更高的具体操作。尽量减少完成一个操作的步骤。</span><span class="sxs-lookup"><span data-stu-id="08796-p105">Provide granular actions that make common tasks within your add-in more efficient to perform. Minimize the number of steps an action takes to complete.</span></span>
- <span data-ttu-id="08796-139">关于命令在 Office 功能区中的位置：</span><span class="sxs-lookup"><span data-stu-id="08796-139">For the placement of your commands in the Office ribbon:</span></span>
    - <span data-ttu-id="08796-p106">将命令放置在现有的选项卡（插入、审阅等）上，如果提供的功能适合那个位置。例如，如果外接程序允许用户插入媒体，则将组添加到“插入”选项卡。请注意，并非所有选项卡都在所有的 Office 版本之间可用。有关详细信息，请参阅 [Office 外接程序 XML 清单](../develop/add-in-manifests.md)。</span><span class="sxs-lookup"><span data-stu-id="08796-p106">Place commands on an existing tab (Insert, Review, and so on) if the functionality provided fits there. For example, if your add-in enables users to insert media, add a group to the Insert tab. Note that not all tabs are available across all Office versions. For more information, see [Office Add-ins XML manifest](../develop/add-in-manifests.md).</span></span> 
    - <span data-ttu-id="08796-p107">如果功能不适合其他选项卡，并且你拥有少于 6 个的顶级命令，将命令放置在“开始”选项卡上。如果外接程序需要跨 Office 版本（如 Office Desktop 和 Office Online）运行，并且某个选项卡并非在所有版本中（例如，“设计”选项卡不存在于 Office Online 中）都提供，你也可以将命令添加到“开始”选项卡。</span><span class="sxs-lookup"><span data-stu-id="08796-p107">Place commands on the Home tab if the functionality doesn't fit on another tab, and you have fewer than six top-level commands. You can also add commands to the Home tab if your add-in needs to work across Office versions (such as Office Desktop and Office Online) and a tab is not available in all versions (for example, the Design tab doesn't exist in Office Online).</span></span>  
    - <span data-ttu-id="08796-145">如果你拥有 6 个以上的顶级命令命令，将命令放置在自定义选项卡上。</span><span class="sxs-lookup"><span data-stu-id="08796-145">Place commands on a custom tab if you have more than six top-level commands.</span></span>
    - <span data-ttu-id="08796-p108">对组进行命名以与外接程序的名称相匹配。如果你拥有多个组，则基于对应组中的命令提供的功能为每个组命名。</span><span class="sxs-lookup"><span data-stu-id="08796-p108">Name your group to match the name of your add-in. If you have multiple groups, name each group based on the functionality that the commands in that group provide.</span></span>
    - <span data-ttu-id="08796-148">请勿添加不必要的按钮，这样会增加加载项占用的空间。</span><span class="sxs-lookup"><span data-stu-id="08796-148">Do not add superfluous buttons to increase the real estate of your add-in.</span></span>

     > [!NOTE]
     > <span data-ttu-id="08796-149">占用过多空间的加载项可能无法通过 [AppSource 验证](/office/dev/store/validation-policies)。</span><span class="sxs-lookup"><span data-stu-id="08796-149">Add-ins that take up too much space might not pass [AppSource validation](/office/dev/store/validation-policies).</span></span>

- <span data-ttu-id="08796-150">对于所有图标，请遵循[图标设计准则](add-in-icons.md)。</span><span class="sxs-lookup"><span data-stu-id="08796-150">For all icons, follow the [icon design guidelines](add-in-icons.md).</span></span>
- <span data-ttu-id="08796-151">提供也可以在不支持命令的主机上运行的加载项的版本。</span><span class="sxs-lookup"><span data-stu-id="08796-151">Provide a version of your add-in that also works on hosts that do not support commands.</span></span> <span data-ttu-id="08796-152">单个加载项清单可以在命令感知（带有命令）和非命令感知（作为任务窗格）的主机中工作。</span><span class="sxs-lookup"><span data-stu-id="08796-152">A single add-in manifest can work in both command-aware (with commands) and non-command-aware (as a task pane) hosts.</span></span>

   <span data-ttu-id="08796-153">*图 3. Office 2013 中的任务窗格加载项，以及 Office 2016 中使用加载项命令的相同加载项*</span><span class="sxs-lookup"><span data-stu-id="08796-153">*Figure 3. Task pane add-in in Office 2013 and the same add-in using add-in commands in Office 2016*</span></span>

   ![显示 Office 2013 中的任务窗格加载项，以及 Office 2016 中使用加载项命令的相同加载项的屏幕截图](../images/office-task-pane-add-ins.png)


## <a name="next-steps"></a><span data-ttu-id="08796-155">后续步骤</span><span class="sxs-lookup"><span data-stu-id="08796-155">Next steps</span></span>

<span data-ttu-id="08796-156">加载项命令的最佳入门方式是参照 GitHub 上的 [Office 加载项命令示例](https://github.com/OfficeDev/Office-Add-in-Commands-Samples/)。</span><span class="sxs-lookup"><span data-stu-id="08796-156">The best way to get started using add-in commands is to take a look at the [Office Add-in commands samples](https://github.com/OfficeDev/Office-Add-in-Commands-Samples/) on GitHub.</span></span>

<span data-ttu-id="08796-157">若要详细了解如何在清单中指定加载项命令，请参阅[在清单中创建加载项命令](../develop/create-addin-commands.md)和 [VersionOverrides](/office/dev/add-ins/reference/manifest/versionoverrides) 参考内容。</span><span class="sxs-lookup"><span data-stu-id="08796-157">For more information about specifying add-in commands in your manifest, see [Create add-in commands in your manifest](../develop/create-addin-commands.md) and the [VersionOverrides](/office/dev/add-ins/reference/manifest/versionoverrides) reference content.</span></span>
