---
title: 加载项命令的基本概念
description: 了解如何将自定义功能区按钮和菜单项添加到 Office 作为 Office 加载项的一部分。
ms.date: 05/12/2020
localization_priority: Priority
ms.openlocfilehash: 2fe14a41c93b53164ab0fa3a7d25f5b9810b9c6a
ms.sourcegitcommit: 7ef14753dce598a5804dad8802df7aaafe046da7
ms.translationtype: HT
ms.contentlocale: zh-CN
ms.lasthandoff: 07/10/2020
ms.locfileid: "45093873"
---
# <a name="add-in-commands-for-excel-powerpoint-and-word"></a><span data-ttu-id="5d26b-103">Excel、PowerPoint 和 Word 的加载项命令</span><span class="sxs-lookup"><span data-stu-id="5d26b-103">Add-in commands for Excel, PowerPoint, and Word</span></span>

<span data-ttu-id="5d26b-104">Add-in commands are UI elements that extend the Office UI and start actions in your add-in.</span><span class="sxs-lookup"><span data-stu-id="5d26b-104">Add-in commands are UI elements that extend the Office UI and start actions in your add-in.</span></span> <span data-ttu-id="5d26b-105">You can use add-in commands to add a button on the ribbon or an item to a context menu.</span><span class="sxs-lookup"><span data-stu-id="5d26b-105">You can use add-in commands to add a button on the ribbon or an item to a context menu.</span></span> <span data-ttu-id="5d26b-106">When users select an add-in command, they initiate actions such as running JavaScript code, or showing a page of the add-in in a task pane.</span><span class="sxs-lookup"><span data-stu-id="5d26b-106">When users select an add-in command, they initiate actions such as running JavaScript code, or showing a page of the add-in in a task pane.</span></span> <span data-ttu-id="5d26b-107">Add-in commands help users find and use your add-in, which can help increase your add-in's adoption and reuse, and improve customer retention.</span><span class="sxs-lookup"><span data-stu-id="5d26b-107">Add-in commands help users find and use your add-in, which can help increase your add-in's adoption and reuse, and improve customer retention.</span></span>

<span data-ttu-id="5d26b-108">有关此功能的概述，请观看视频 [Office 应用功能区中的加载项命令](https://channel9.msdn.com/events/Build/2016/P551)。</span><span class="sxs-lookup"><span data-stu-id="5d26b-108">For an overview of the feature, see the video [Add-in Commands in the Office app ribbon](https://channel9.msdn.com/events/Build/2016/P551).</span></span>

> [!NOTE]
> <span data-ttu-id="5d26b-109">SharePoint catalogs do not support add-in commands.</span><span class="sxs-lookup"><span data-stu-id="5d26b-109">SharePoint catalogs do not support add-in commands.</span></span> <span data-ttu-id="5d26b-110">You can deploy add-in commands via [Centralized Deployment](../publish/centralized-deployment.md) or [AppSource](/office/dev/store/submit-to-appsource-via-partner-center), or use [sideloading](../testing/create-a-network-shared-folder-catalog-for-task-pane-and-content-add-ins.md) to deploy your add-in command for testing.</span><span class="sxs-lookup"><span data-stu-id="5d26b-110">You can deploy add-in commands via [Centralized Deployment](../publish/centralized-deployment.md) or [AppSource](/office/dev/store/submit-to-appsource-via-partner-center), or use [sideloading](../testing/create-a-network-shared-folder-catalog-for-task-pane-and-content-add-ins.md) to deploy your add-in command for testing.</span></span>

> [!IMPORTANT]
> <span data-ttu-id="5d26b-111">Outlook 中也支持加载项命令。</span><span class="sxs-lookup"><span data-stu-id="5d26b-111">Add-in commands are also supported in Outlook.</span></span> <span data-ttu-id="5d26b-112">有关详细信息，请参阅[适用于 Outlook 的加载项命令](../outlook/add-in-commands-for-outlook.md)。</span><span class="sxs-lookup"><span data-stu-id="5d26b-112">For more information, see [Add-in commands for Outlook](../outlook/add-in-commands-for-outlook.md).</span></span>

<span data-ttu-id="5d26b-113">*图 1：在 Excel Desktop 中运行命令的加载项*</span><span class="sxs-lookup"><span data-stu-id="5d26b-113">*Figure 1. Add-in with commands running in Excel Desktop*</span></span>

![Excel 中的加载项命令屏幕截图](../images/add-in-commands-1.png)

<span data-ttu-id="5d26b-115">*图 2：在 Excel 网页版中运行命令的加载项*</span><span class="sxs-lookup"><span data-stu-id="5d26b-115">*Figure 2. Add-in with commands running in Excel on the web*</span></span>

![Excel 网页版中加载项命令的屏幕截图](../images/add-in-commands-2.png)

## <a name="command-capabilities"></a><span data-ttu-id="5d26b-117">命令功能</span><span class="sxs-lookup"><span data-stu-id="5d26b-117">Command capabilities</span></span>

<span data-ttu-id="5d26b-118">目前支持下列命令功能。</span><span class="sxs-lookup"><span data-stu-id="5d26b-118">The following command capabilities are currently supported.</span></span>

> [!NOTE]
> <span data-ttu-id="5d26b-119">内容加载项暂不支持加载项命令。</span><span class="sxs-lookup"><span data-stu-id="5d26b-119">Content add-ins do not currently support add-in commands.</span></span>

### <a name="extension-points"></a><span data-ttu-id="5d26b-120">扩展点</span><span class="sxs-lookup"><span data-stu-id="5d26b-120">Extension points</span></span>

- <span data-ttu-id="5d26b-121">功能区选项卡 - 扩展内置选项卡或新建自定义选项卡。</span><span class="sxs-lookup"><span data-stu-id="5d26b-121">Ribbon tabs - Extend built-in tabs or create a new custom tab.</span></span>
- <span data-ttu-id="5d26b-122">上下文菜单 - 扩展所选上下文菜单。</span><span class="sxs-lookup"><span data-stu-id="5d26b-122">Context menus - Extend selected context menus.</span></span>

### <a name="control-types"></a><span data-ttu-id="5d26b-123">控件类型</span><span class="sxs-lookup"><span data-stu-id="5d26b-123">Control types</span></span>

- <span data-ttu-id="5d26b-124">简单按钮 - 触发特定操作。</span><span class="sxs-lookup"><span data-stu-id="5d26b-124">Simple buttons - trigger specific actions.</span></span>
- <span data-ttu-id="5d26b-125">菜单 - 简单的下拉菜单，内含可触发操作的按钮。</span><span class="sxs-lookup"><span data-stu-id="5d26b-125">Menus - simple menu dropdown with buttons that trigger actions.</span></span>

### <a name="actions"></a><span data-ttu-id="5d26b-126">操作</span><span class="sxs-lookup"><span data-stu-id="5d26b-126">Actions</span></span>

- <span data-ttu-id="5d26b-127">ShowTaskpane - 显示一个或多个在其中加载自定义 HTML 页的窗格。</span><span class="sxs-lookup"><span data-stu-id="5d26b-127">ShowTaskpane - Displays one or multiple panes that load custom HTML pages inside them.</span></span>
- <span data-ttu-id="5d26b-128">ExecuteFunction - Loads an invisible HTML page and then execute a JavaScript function within it.</span><span class="sxs-lookup"><span data-stu-id="5d26b-128">ExecuteFunction - Loads an invisible HTML page and then execute a JavaScript function within it.</span></span> <span data-ttu-id="5d26b-129">To show UI within your function (such as errors, progress, or additional input) you can use the [displayDialog](/javascript/api/office/office.ui) API.</span><span class="sxs-lookup"><span data-stu-id="5d26b-129">To show UI within your function (such as errors, progress, or additional input) you can use the [displayDialog](/javascript/api/office/office.ui) API.</span></span>  

### <a name="default-enabled-or-disabled-status-preview"></a><span data-ttu-id="5d26b-130">默认启用或禁用状态（预览版）</span><span class="sxs-lookup"><span data-stu-id="5d26b-130">Default Enabled or Disabled Status (preview)</span></span>

<span data-ttu-id="5d26b-131">可指定在加载项启动时是启用还是禁用该命令，并以编程方式更改设置。</span><span class="sxs-lookup"><span data-stu-id="5d26b-131">You can specify whether the command is enabled or disabled when your add-in launches, and programmatically change the setting.</span></span>

> [!NOTE]
> <span data-ttu-id="5d26b-132">此功能处于预览状态，并非在所有主机或方案中均受支持。</span><span class="sxs-lookup"><span data-stu-id="5d26b-132">This feature is in preview and is not supported in all hosts or scenarios.</span></span> <span data-ttu-id="5d26b-133">有关详细信息，请参阅[启用和禁用加载项命令](disable-add-in-commands.md)。</span><span class="sxs-lookup"><span data-stu-id="5d26b-133">For more information, see [Enable and Disable Add-in Commands](disable-add-in-commands.md).</span></span>

## <a name="supported-platforms"></a><span data-ttu-id="5d26b-134">支持的平台</span><span class="sxs-lookup"><span data-stu-id="5d26b-134">Supported platforms</span></span>

<span data-ttu-id="5d26b-135">目前，以下平台支持加载项命令。</span><span class="sxs-lookup"><span data-stu-id="5d26b-135">Add-in commands are currently supported on the following platforms.</span></span>

- <span data-ttu-id="5d26b-136">Windows 版 Office（内部版本 16.0.6769 及更高版本，关联至 Microsoft 365 订阅）</span><span class="sxs-lookup"><span data-stu-id="5d26b-136">Office on Windows (build 16.0.6769+, connected to Microsoft 365 subscription)</span></span>
- <span data-ttu-id="5d26b-137">Windows 版 Office 2019</span><span class="sxs-lookup"><span data-stu-id="5d26b-137">Office 2019 on Windows</span></span>
- <span data-ttu-id="5d26b-138">Mac 版 Office（内部版本 15.33 及更高版本，关联至 Microsoft 365 订阅）</span><span class="sxs-lookup"><span data-stu-id="5d26b-138">Office on Mac (build 15.33+, connected to Microsoft 365 subscription)</span></span>
- <span data-ttu-id="5d26b-139">Mac 版 Office 2019</span><span class="sxs-lookup"><span data-stu-id="5d26b-139">Office 2019 on Mac</span></span>
- <span data-ttu-id="5d26b-140">Office 网页版</span><span class="sxs-lookup"><span data-stu-id="5d26b-140">Office on the web</span></span>

> [!NOTE]
> <span data-ttu-id="5d26b-141">有关 Outlook 支持的信息，请参阅[适用于 Outlook 的加载项命令](../outlook/add-in-commands-for-outlook.md)。</span><span class="sxs-lookup"><span data-stu-id="5d26b-141">For information about support in Outlook, see [Add-in commands for Outlook](../outlook/add-in-commands-for-outlook.md).</span></span>

## <a name="debugging"></a><span data-ttu-id="5d26b-142">调试</span><span class="sxs-lookup"><span data-stu-id="5d26b-142">Debugging</span></span>

<span data-ttu-id="5d26b-143">必须在 Office 网页版中运行加载项命令，才能调试命令。</span><span class="sxs-lookup"><span data-stu-id="5d26b-143">To debug an Add-in Command, you must run it in Office on the web.</span></span> <span data-ttu-id="5d26b-144">有关详细信息，请参阅[在 Office 网页版中调试加载项](../testing/debug-add-ins-in-office-online.md)。</span><span class="sxs-lookup"><span data-stu-id="5d26b-144">For details, see [Debug add-ins in Office on the web](../testing/debug-add-ins-in-office-online.md).</span></span>

## <a name="best-practices"></a><span data-ttu-id="5d26b-145">最佳做法</span><span class="sxs-lookup"><span data-stu-id="5d26b-145">Best practices</span></span>

<span data-ttu-id="5d26b-146">在开发外接程序命令时应用下面的最佳做法：</span><span class="sxs-lookup"><span data-stu-id="5d26b-146">Apply the following best practices when you develop add-in commands:</span></span>

- <span data-ttu-id="5d26b-147">Use commands to represent a specific action with a clear and specific outcome for users.</span><span class="sxs-lookup"><span data-stu-id="5d26b-147">Use commands to represent a specific action with a clear and specific outcome for users.</span></span> <span data-ttu-id="5d26b-148">Do not combine multiple actions in a single button.</span><span class="sxs-lookup"><span data-stu-id="5d26b-148">Do not combine multiple actions in a single button.</span></span>
- <span data-ttu-id="5d26b-149">Provide granular actions that make common tasks within your add-in more efficient to perform.</span><span class="sxs-lookup"><span data-stu-id="5d26b-149">Provide granular actions that make common tasks within your add-in more efficient to perform.</span></span> <span data-ttu-id="5d26b-150">Minimize the number of steps an action takes to complete.</span><span class="sxs-lookup"><span data-stu-id="5d26b-150">Minimize the number of steps an action takes to complete.</span></span>
- <span data-ttu-id="5d26b-151">关于命令在 Office 应用功能区中的位置：</span><span class="sxs-lookup"><span data-stu-id="5d26b-151">For the placement of your commands in the Office app ribbon:</span></span>
    - <span data-ttu-id="5d26b-152">Place commands on an existing tab (Insert, Review, and so on) if the functionality provided fits there.</span><span class="sxs-lookup"><span data-stu-id="5d26b-152">Place commands on an existing tab (Insert, Review, and so on) if the functionality provided fits there.</span></span> <span data-ttu-id="5d26b-153">For example, if your add-in enables users to insert media, add a group to the Insert tab. Note that not all tabs are available across all Office versions.</span><span class="sxs-lookup"><span data-stu-id="5d26b-153">For example, if your add-in enables users to insert media, add a group to the Insert tab. Note that not all tabs are available across all Office versions.</span></span> <span data-ttu-id="5d26b-154">For more information, see [Office Add-ins XML manifest](../develop/add-in-manifests.md).</span><span class="sxs-lookup"><span data-stu-id="5d26b-154">For more information, see [Office Add-ins XML manifest](../develop/add-in-manifests.md).</span></span>
    - <span data-ttu-id="5d26b-155">Place commands on the Home tab if the functionality doesn't fit on another tab, and you have fewer than six top-level commands.</span><span class="sxs-lookup"><span data-stu-id="5d26b-155">Place commands on the Home tab if the functionality doesn't fit on another tab, and you have fewer than six top-level commands.</span></span> <span data-ttu-id="5d26b-156">You can also add commands to the Home tab if your add-in needs to work across Office versions (such as Office on the web or desktop) and a tab is not available in all versions (for example, the Design tab doesn't exist in Office on the web).</span><span class="sxs-lookup"><span data-stu-id="5d26b-156">You can also add commands to the Home tab if your add-in needs to work across Office versions (such as Office on the web or desktop) and a tab is not available in all versions (for example, the Design tab doesn't exist in Office on the web).</span></span>  
    - <span data-ttu-id="5d26b-157">如果你拥有 6 个以上的顶级命令命令，将命令放置在自定义选项卡上。</span><span class="sxs-lookup"><span data-stu-id="5d26b-157">Place commands on a custom tab if you have more than six top-level commands.</span></span>
    - <span data-ttu-id="5d26b-158">Name your group to match the name of your add-in.</span><span class="sxs-lookup"><span data-stu-id="5d26b-158">Name your group to match the name of your add-in.</span></span> <span data-ttu-id="5d26b-159">If you have multiple groups, name each group based on the functionality that the commands in that group provide.</span><span class="sxs-lookup"><span data-stu-id="5d26b-159">If you have multiple groups, name each group based on the functionality that the commands in that group provide.</span></span>
    - <span data-ttu-id="5d26b-160">请勿添加不必要的按钮，这样会增加加载项占用的空间。</span><span class="sxs-lookup"><span data-stu-id="5d26b-160">Do not add superfluous buttons to increase the real estate of your add-in.</span></span>

     > [!NOTE]
     > <span data-ttu-id="5d26b-161">占用过多空间的加载项可能无法通过 [AppSource 验证](/legal/marketplace/certification-policies)。</span><span class="sxs-lookup"><span data-stu-id="5d26b-161">Add-ins that take up too much space might not pass [AppSource validation](/legal/marketplace/certification-policies).</span></span>

- <span data-ttu-id="5d26b-162">对于所有图标，请遵循[图标设计准则](add-in-icons.md)。</span><span class="sxs-lookup"><span data-stu-id="5d26b-162">For all icons, follow the [icon design guidelines](add-in-icons.md).</span></span>
- <span data-ttu-id="5d26b-163">提供也可以在不支持命令的主机上运行的加载项的版本。</span><span class="sxs-lookup"><span data-stu-id="5d26b-163">Provide a version of your add-in that also works on hosts that do not support commands.</span></span> <span data-ttu-id="5d26b-164">单个加载项清单可以在命令感知（带有命令）和非命令感知（作为任务窗格）的主机中工作。</span><span class="sxs-lookup"><span data-stu-id="5d26b-164">A single add-in manifest can work in both command-aware (with commands) and non-command-aware (as a task pane) hosts.</span></span>

   <span data-ttu-id="5d26b-165">*图 3. Office 2013 中的任务窗格加载项，以及 Office 2016 中使用加载项命令的相同加载项*</span><span class="sxs-lookup"><span data-stu-id="5d26b-165">*Figure 3. Task pane add-in in Office 2013 and the same add-in using add-in commands in Office 2016*</span></span>

   ![显示 Office 2013 中的任务窗格加载项，以及 Office 2016 中使用加载项命令的相同加载项的屏幕截图](../images/office-task-pane-add-ins.png)


## <a name="next-steps"></a><span data-ttu-id="5d26b-167">后续步骤</span><span class="sxs-lookup"><span data-stu-id="5d26b-167">Next steps</span></span>

<span data-ttu-id="5d26b-168">加载项命令的最佳入门方式是参照 GitHub 上的 [Office 加载项命令示例](https://github.com/OfficeDev/Office-Add-in-Commands-Samples/)。</span><span class="sxs-lookup"><span data-stu-id="5d26b-168">The best way to get started using add-in commands is to take a look at the [Office Add-in commands samples](https://github.com/OfficeDev/Office-Add-in-Commands-Samples/) on GitHub.</span></span>

<span data-ttu-id="5d26b-169">若要详细了解如何在清单中指定加载项命令，请参阅[在清单中创建加载项命令](../develop/create-addin-commands.md)和 [VersionOverrides](../reference/manifest/versionoverrides.md) 参考内容。</span><span class="sxs-lookup"><span data-stu-id="5d26b-169">For more information about specifying add-in commands in your manifest, see [Create add-in commands in your manifest](../develop/create-addin-commands.md) and the [VersionOverrides](../reference/manifest/versionoverrides.md) reference content.</span></span>
