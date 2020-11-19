---
title: 加载项命令的基本概念
description: 了解如何将自定义功能区按钮和菜单项添加到 Office 作为 Office 加载项的一部分。
ms.date: 11/01/2020
localization_priority: Priority
ms.openlocfilehash: 3d7d99f05e9b02712a4f416b891d3be38875525b
ms.sourcegitcommit: 3189c4bd62dbe5950b19f28ac2c1314b6d304dca
ms.translationtype: HT
ms.contentlocale: zh-CN
ms.lasthandoff: 11/17/2020
ms.locfileid: "49087964"
---
# <a name="add-in-commands-for-excel-powerpoint-and-word"></a><span data-ttu-id="507b8-103">Excel、PowerPoint 和 Word 的加载项命令</span><span class="sxs-lookup"><span data-stu-id="507b8-103">Add-in commands for Excel, PowerPoint, and Word</span></span>

<span data-ttu-id="507b8-p101">外接程序命令是 UI 元素，可扩展 Office UI，并在外接程序中启动操作。使用外接程序命令，可以在功能区上添加按钮，也可以向上下文菜单添加项。当用户选择外接程序命令时，将启动操作，如运行 JavaScript 代码或在任务窗格中显示外接程序页面。外接程序命令可帮助用户查找和使用外接程序，从而提高外接程序的采用率和重用率以及客户保留率。</span><span class="sxs-lookup"><span data-stu-id="507b8-p101">Add-in commands are UI elements that extend the Office UI and start actions in your add-in. You can use add-in commands to add a button on the ribbon or an item to a context menu. When users select an add-in command, they initiate actions such as running JavaScript code, or showing a page of the add-in in a task pane. Add-in commands help users find and use your add-in, which can help increase your add-in's adoption and reuse, and improve customer retention.</span></span>

<span data-ttu-id="507b8-108">有关此功能的概述，请观看视频 [Office 应用功能区中的加载项命令](https://channel9.msdn.com/events/Build/2016/P551)。</span><span class="sxs-lookup"><span data-stu-id="507b8-108">For an overview of the feature, see the video [Add-in Commands in the Office app ribbon](https://channel9.msdn.com/events/Build/2016/P551).</span></span>

> [!NOTE]
> <span data-ttu-id="507b8-p102">SharePoint 目录不支持加载项命令。可以通过[集中部署](../publish/centralized-deployment.md)或 [AppSource](/office/dev/store/submit-to-appsource-via-partner-center) 部署加载项命令，也可以使用[旁加载](../testing/create-a-network-shared-folder-catalog-for-task-pane-and-content-add-ins.md)部署加载项命令以供测试。</span><span class="sxs-lookup"><span data-stu-id="507b8-p102">SharePoint catalogs do not support add-in commands. You can deploy add-in commands via [Centralized Deployment](../publish/centralized-deployment.md) or [AppSource](/office/dev/store/submit-to-appsource-via-partner-center), or use [sideloading](../testing/create-a-network-shared-folder-catalog-for-task-pane-and-content-add-ins.md) to deploy your add-in command for testing.</span></span>

> [!IMPORTANT]
> <span data-ttu-id="507b8-111">Outlook 中也支持加载项命令。</span><span class="sxs-lookup"><span data-stu-id="507b8-111">Add-in commands are also supported in Outlook.</span></span> <span data-ttu-id="507b8-112">有关详细信息，请参阅[适用于 Outlook 的加载项命令](../outlook/add-in-commands-for-outlook.md)。</span><span class="sxs-lookup"><span data-stu-id="507b8-112">For more information, see [Add-in commands for Outlook](../outlook/add-in-commands-for-outlook.md).</span></span>

<span data-ttu-id="507b8-113">*图 1：在 Excel Desktop 中运行命令的加载项*</span><span class="sxs-lookup"><span data-stu-id="507b8-113">*Figure 1. Add-in with commands running in Excel Desktop*</span></span>

![Excel 中的加载项命令屏幕截图](../images/add-in-commands-1.png)

<span data-ttu-id="507b8-115">*图 2：在 Excel 网页版中运行命令的加载项*</span><span class="sxs-lookup"><span data-stu-id="507b8-115">*Figure 2. Add-in with commands running in Excel on the web*</span></span>

![Excel 网页版中加载项命令的屏幕截图](../images/add-in-commands-2.png)

## <a name="command-capabilities"></a><span data-ttu-id="507b8-117">命令功能</span><span class="sxs-lookup"><span data-stu-id="507b8-117">Command capabilities</span></span>

<span data-ttu-id="507b8-118">目前支持下列命令功能。</span><span class="sxs-lookup"><span data-stu-id="507b8-118">The following command capabilities are currently supported.</span></span>

> [!NOTE]
> <span data-ttu-id="507b8-119">内容加载项暂不支持加载项命令。</span><span class="sxs-lookup"><span data-stu-id="507b8-119">Content add-ins do not currently support add-in commands.</span></span>

### <a name="extension-points"></a><span data-ttu-id="507b8-120">扩展点</span><span class="sxs-lookup"><span data-stu-id="507b8-120">Extension points</span></span>

- <span data-ttu-id="507b8-121">功能区选项卡 - 扩展内置选项卡或新建自定义选项卡。</span><span class="sxs-lookup"><span data-stu-id="507b8-121">Ribbon tabs - Extend built-in tabs or create a new custom tab.</span></span>
- <span data-ttu-id="507b8-122">上下文菜单 - 扩展所选上下文菜单。</span><span class="sxs-lookup"><span data-stu-id="507b8-122">Context menus - Extend selected context menus.</span></span>

### <a name="control-types"></a><span data-ttu-id="507b8-123">控件类型</span><span class="sxs-lookup"><span data-stu-id="507b8-123">Control types</span></span>

- <span data-ttu-id="507b8-124">简单按钮 - 触发特定操作。</span><span class="sxs-lookup"><span data-stu-id="507b8-124">Simple buttons - trigger specific actions.</span></span>
- <span data-ttu-id="507b8-125">菜单 - 简单的下拉菜单，内含可触发操作的按钮。</span><span class="sxs-lookup"><span data-stu-id="507b8-125">Menus - simple menu dropdown with buttons that trigger actions.</span></span>

### <a name="actions"></a><span data-ttu-id="507b8-126">操作</span><span class="sxs-lookup"><span data-stu-id="507b8-126">Actions</span></span>

- <span data-ttu-id="507b8-127">ShowTaskpane - 显示一个或多个在其中加载自定义 HTML 页的窗格。</span><span class="sxs-lookup"><span data-stu-id="507b8-127">ShowTaskpane - Displays one or multiple panes that load custom HTML pages inside them.</span></span>
- <span data-ttu-id="507b8-p104">ExecuteFunction - 加载一个不可见的 HTML 页，然后在其中执行一个 JavaScript 函数。若要在你的函数（例如错误、进度或其他输入）中显示 UI，你可以使用 [displayDialog](/javascript/api/office/office.ui) API。</span><span class="sxs-lookup"><span data-stu-id="507b8-p104">ExecuteFunction - Loads an invisible HTML page and then execute a JavaScript function within it. To show UI within your function (such as errors, progress, or additional input) you can use the [displayDialog](/javascript/api/office/office.ui) API.</span></span>  

### <a name="default-enabled-or-disabled-status"></a><span data-ttu-id="507b8-130">默认启用或禁用状态</span><span class="sxs-lookup"><span data-stu-id="507b8-130">Default Enabled or Disabled Status</span></span>

<span data-ttu-id="507b8-131">可指定在加载项启动时是启用还是禁用该命令，并以编程方式更改设置。</span><span class="sxs-lookup"><span data-stu-id="507b8-131">You can specify whether the command is enabled or disabled when your add-in launches, and programmatically change the setting.</span></span>

> [!NOTE]
> <span data-ttu-id="507b8-132">此功能并非在所有 Office 应用程序或方案中受到支持。</span><span class="sxs-lookup"><span data-stu-id="507b8-132">This feature is not supported in all Office applications or scenarios.</span></span> <span data-ttu-id="507b8-133">有关详细信息，请参阅[启用和禁用加载项命令](disable-add-in-commands.md)。</span><span class="sxs-lookup"><span data-stu-id="507b8-133">For more information, see [Enable and Disable Add-in Commands](disable-add-in-commands.md).</span></span>

### <a name="position-on-the-ribbon-preview"></a><span data-ttu-id="507b8-134">功能区上的位置（预览）</span><span class="sxs-lookup"><span data-stu-id="507b8-134">Position on the ribbon (preview)</span></span>

<span data-ttu-id="507b8-135">可以指定自定义选项卡在 Office 应用程序功能区上的显示位置，例如“在“主页”选项卡右侧”。</span><span class="sxs-lookup"><span data-stu-id="507b8-135">You can specify where a custom tab appears on the Office application's ribbon, such as "just to the right of the Home tab".</span></span>

> [!NOTE]
> <span data-ttu-id="507b8-136">并非所有 Office 应用程序或方案均支持此功能。</span><span class="sxs-lookup"><span data-stu-id="507b8-136">This feature is not supported in all Office applications or scenarios.</span></span> <span data-ttu-id="507b8-137">有关详细信息，请参阅[在功能区上定位自定义选项卡](custom-tab-placement.md)。</span><span class="sxs-lookup"><span data-stu-id="507b8-137">For more information, see [Position a custom tab on the ribbon](custom-tab-placement.md).</span></span>

### <a name="integration-of-built-in-office-buttons-preview"></a><span data-ttu-id="507b8-138">内置 Office 按钮集成（预览）</span><span class="sxs-lookup"><span data-stu-id="507b8-138">Integration of built-in Office buttons (preview)</span></span>

<span data-ttu-id="507b8-139">可将内置的 Office 功能区按钮插入到自定义命令组和自定义功能区选项卡中。</span><span class="sxs-lookup"><span data-stu-id="507b8-139">You can insert the built-in Office ribbon buttons into your custom command groups and custom ribbon tabs.</span></span>

> [!NOTE]
> <span data-ttu-id="507b8-140">并非所有 Office 应用程序或方案均支持此功能。</span><span class="sxs-lookup"><span data-stu-id="507b8-140">This feature is not supported in all Office applications or scenarios.</span></span> <span data-ttu-id="507b8-141">有关详细信息，请参阅[将内置 Office 按钮集成到自定义选项卡中](built-in-button-integration.md)。</span><span class="sxs-lookup"><span data-stu-id="507b8-141">For more information, see [Integrate built-in Office buttons into custom tabs](built-in-button-integration.md).</span></span>


## <a name="supported-platforms"></a><span data-ttu-id="507b8-142">支持的平台</span><span class="sxs-lookup"><span data-stu-id="507b8-142">Supported platforms</span></span>

<span data-ttu-id="507b8-143">目前，以下平台支持加载项命令，但先前[命令功能](#command-capabilities)的小节中指定的限制除外。</span><span class="sxs-lookup"><span data-stu-id="507b8-143">Add-in commands are currently supported on the following platforms, except for limitations specified in the subsections of [Command capabilities](#command-capabilities) earlier.</span></span>

- <span data-ttu-id="507b8-144">Windows 版 Office（内部版本 16.0.6769 及更高版本，关联至 Microsoft 365 订阅）</span><span class="sxs-lookup"><span data-stu-id="507b8-144">Office on Windows (build 16.0.6769+, connected to a Microsoft 365 subscription)</span></span>
- <span data-ttu-id="507b8-145">Windows 版 Office 2019</span><span class="sxs-lookup"><span data-stu-id="507b8-145">Office 2019 on Windows</span></span>
- <span data-ttu-id="507b8-146">Mac 版 Office（内部版本 15.33 及更高版本，关联至 Microsoft 365 订阅）</span><span class="sxs-lookup"><span data-stu-id="507b8-146">Office on Mac (build 15.33+, connected to a Microsoft 365 subscription)</span></span>
- <span data-ttu-id="507b8-147">Mac 版 Office 2019</span><span class="sxs-lookup"><span data-stu-id="507b8-147">Office 2019 on Mac</span></span>
- <span data-ttu-id="507b8-148">Office 网页版</span><span class="sxs-lookup"><span data-stu-id="507b8-148">Office on the web</span></span>

> [!NOTE]
> <span data-ttu-id="507b8-149">有关 Outlook 支持的信息，请参阅[适用于 Outlook 的加载项命令](../outlook/add-in-commands-for-outlook.md)。</span><span class="sxs-lookup"><span data-stu-id="507b8-149">For information about support in Outlook, see [Add-in commands for Outlook](../outlook/add-in-commands-for-outlook.md).</span></span>

## <a name="debugging"></a><span data-ttu-id="507b8-150">调试</span><span class="sxs-lookup"><span data-stu-id="507b8-150">Debugging</span></span>

<span data-ttu-id="507b8-151">必须在 Office 网页版中运行加载项命令，才能调试命令。</span><span class="sxs-lookup"><span data-stu-id="507b8-151">To debug an Add-in Command, you must run it in Office on the web.</span></span> <span data-ttu-id="507b8-152">有关详细信息，请参阅[在 Office 网页版中调试加载项](../testing/debug-add-ins-in-office-online.md)。</span><span class="sxs-lookup"><span data-stu-id="507b8-152">For details, see [Debug add-ins in Office on the web](../testing/debug-add-ins-in-office-online.md).</span></span>

## <a name="best-practices"></a><span data-ttu-id="507b8-153">最佳做法</span><span class="sxs-lookup"><span data-stu-id="507b8-153">Best practices</span></span>

<span data-ttu-id="507b8-154">在开发外接程序命令时应用下面的最佳做法：</span><span class="sxs-lookup"><span data-stu-id="507b8-154">Apply the following best practices when you develop add-in commands:</span></span>

- <span data-ttu-id="507b8-p109">使用命令来表示会给用户带来明确具体结果的特定操作。不要在单个按钮中组合多个操作。</span><span class="sxs-lookup"><span data-stu-id="507b8-p109">Use commands to represent a specific action with a clear and specific outcome for users. Do not combine multiple actions in a single button.</span></span>
- <span data-ttu-id="507b8-p110">提供使您的外接程序中的常见任务执行效率更高的具体操作。尽量减少完成一个操作的步骤。</span><span class="sxs-lookup"><span data-stu-id="507b8-p110">Provide granular actions that make common tasks within your add-in more efficient to perform. Minimize the number of steps an action takes to complete.</span></span>
- <span data-ttu-id="507b8-159">关于命令在 Office 应用功能区中的位置：</span><span class="sxs-lookup"><span data-stu-id="507b8-159">For the placement of your commands in the Office app ribbon:</span></span>
    - <span data-ttu-id="507b8-p111">将命令放置在现有的选项卡（插入、审阅等）上，如果提供的功能适合那个位置。例如，如果外接程序允许用户插入媒体，则将组添加到“插入”选项卡。请注意，并非所有选项卡都在所有的 Office 版本之间可用。有关详细信息，请参阅 [Office 外接程序 XML 清单](../develop/add-in-manifests.md)。</span><span class="sxs-lookup"><span data-stu-id="507b8-p111">Place commands on an existing tab (Insert, Review, and so on) if the functionality provided fits there. For example, if your add-in enables users to insert media, add a group to the Insert tab. Note that not all tabs are available across all Office versions. For more information, see [Office Add-ins XML manifest](../develop/add-in-manifests.md).</span></span>
    - <span data-ttu-id="507b8-p112">如果此功能不适合其他选项卡，且顶级命令少于 6 个，请将命令置于“开始”选项卡中。此外，如果加载项需要跨 Office 版本（如 Office 网页版或 Office 桌面版）运行，且并非所有版本都有相应选项卡（例如，Office 网页版中没有“设计”选项卡），也可以将命令添加到“开始”选项卡中。</span><span class="sxs-lookup"><span data-stu-id="507b8-p112">Place commands on the Home tab if the functionality doesn't fit on another tab, and you have fewer than six top-level commands. You can also add commands to the Home tab if your add-in needs to work across Office versions (such as Office on the web or desktop) and a tab is not available in all versions (for example, the Design tab doesn't exist in Office on the web).</span></span>  
    - <span data-ttu-id="507b8-165">如果你拥有 6 个以上的顶级命令命令，将命令放置在自定义选项卡上。</span><span class="sxs-lookup"><span data-stu-id="507b8-165">Place commands on a custom tab if you have more than six top-level commands.</span></span>
    - <span data-ttu-id="507b8-p113">对组进行命名以与外接程序的名称相匹配。如果你拥有多个组，则基于对应组中的命令提供的功能为每个组命名。</span><span class="sxs-lookup"><span data-stu-id="507b8-p113">Name your group to match the name of your add-in. If you have multiple groups, name each group based on the functionality that the commands in that group provide.</span></span>
    - <span data-ttu-id="507b8-168">请勿添加不必要的按钮，这样会增加加载项占用的空间。</span><span class="sxs-lookup"><span data-stu-id="507b8-168">Do not add superfluous buttons to increase the real estate of your add-in.</span></span>
    - <span data-ttu-id="507b8-169">请勿要将“自定义”选项卡置于“主页”选项卡左侧，也不要在打开文档时默认将其放在焦点上，除非加载项是用户与文档进行交互的主要方式。</span><span class="sxs-lookup"><span data-stu-id="507b8-169">Do not position a custom tab to the left of the Home tab, or give it focus by default when the document opens, unless your add-in is the primary way users will interact with the document.</span></span> <span data-ttu-id="507b8-170">过分强调加载项的不便，并惹恼用户和管理员。</span><span class="sxs-lookup"><span data-stu-id="507b8-170">Giving excessive prominence to your add-in inconveniences and annoys users and administrators.</span></span>
    - <span data-ttu-id="507b8-171">如果加载项是用户与文档进行交互的主要方式，而且你具有自定义的功能区选项卡，请考虑将用户经常需要的 Office 功能按钮集成到该选项卡中。</span><span class="sxs-lookup"><span data-stu-id="507b8-171">If your add-in is the primary way users interact with the document and you have a custom ribbon tab, consider integrating into the tab the buttons for the Office functions that users will frequently need.</span></span>

     > [!NOTE]
     > <span data-ttu-id="507b8-172">占用过多空间的加载项可能无法通过 [AppSource 验证](/legal/marketplace/certification-policies)。</span><span class="sxs-lookup"><span data-stu-id="507b8-172">Add-ins that take up too much space might not pass [AppSource validation](/legal/marketplace/certification-policies).</span></span>

- <span data-ttu-id="507b8-173">对于所有图标，请遵循[图标设计准则](add-in-icons.md)。</span><span class="sxs-lookup"><span data-stu-id="507b8-173">For all icons, follow the [icon design guidelines](add-in-icons.md).</span></span>
- <span data-ttu-id="507b8-174">提供也可以在不支持命令的 Office 应用程序上运行的加载项版本。</span><span class="sxs-lookup"><span data-stu-id="507b8-174">Provide a version of your add-in that also works on Office applications that do not support commands.</span></span> <span data-ttu-id="507b8-175">单个加载项清单可以在命令感知型（带有命令）和非命令感知（作为任务窗格）型应用程序中工作。</span><span class="sxs-lookup"><span data-stu-id="507b8-175">A single add-in manifest can work in both command-aware (with commands) and non-command-aware (as a task pane) applications.</span></span>

   <span data-ttu-id="507b8-176">*图 3. Office 2013 中的任务窗格加载项，以及 Office 2016 中使用加载项命令的相同加载项*</span><span class="sxs-lookup"><span data-stu-id="507b8-176">*Figure 3. Task pane add-in in Office 2013 and the same add-in using add-in commands in Office 2016*</span></span>

   ![显示 Office 2013 中的任务窗格加载项，以及 Office 2016 中使用加载项命令的相同加载项的屏幕截图](../images/office-task-pane-add-ins.png)


## <a name="next-steps"></a><span data-ttu-id="507b8-178">后续步骤</span><span class="sxs-lookup"><span data-stu-id="507b8-178">Next steps</span></span>

<span data-ttu-id="507b8-179">加载项命令的最佳入门方式是参照 GitHub 上的 [Office 加载项命令示例](https://github.com/OfficeDev/Office-Add-in-Commands-Samples/)。</span><span class="sxs-lookup"><span data-stu-id="507b8-179">The best way to get started using add-in commands is to take a look at the [Office Add-in commands samples](https://github.com/OfficeDev/Office-Add-in-Commands-Samples/) on GitHub.</span></span>

<span data-ttu-id="507b8-180">若要详细了解如何在清单中指定加载项命令，请参阅[在清单中创建加载项命令](../develop/create-addin-commands.md)和 [VersionOverrides](../reference/manifest/versionoverrides.md) 参考内容。</span><span class="sxs-lookup"><span data-stu-id="507b8-180">For more information about specifying add-in commands in your manifest, see [Create add-in commands in your manifest](../develop/create-addin-commands.md) and the [VersionOverrides](../reference/manifest/versionoverrides.md) reference content.</span></span>
