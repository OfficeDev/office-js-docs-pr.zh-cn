---
title: Office 加载项的 Office UI 元素
description: 获取 Office 外接程序中不同种类的 UI 元素的概述。
ms.date: 12/24/2019
localization_priority: Normal
ms.openlocfilehash: 5b9907924c674ed9db2294621123c394419d0c12
ms.sourcegitcommit: 7ef14753dce598a5804dad8802df7aaafe046da7
ms.translationtype: MT
ms.contentlocale: zh-CN
ms.lasthandoff: 07/10/2020
ms.locfileid: "45093761"
---
# <a name="office-ui-elements-for-office-add-ins"></a><span data-ttu-id="e9d74-103">Office 加载项的 Office UI 元素</span><span class="sxs-lookup"><span data-stu-id="e9d74-103">Office UI elements for Office Add-ins</span></span>

<span data-ttu-id="e9d74-104">You can use several types of UI elements to extend the Office UI, including add-in commands and HTML containers.</span><span class="sxs-lookup"><span data-stu-id="e9d74-104">You can use several types of UI elements to extend the Office UI, including add-in commands and HTML containers.</span></span> <span data-ttu-id="e9d74-105">These UI elements look like a natural extension of Office and work across platforms.</span><span class="sxs-lookup"><span data-stu-id="e9d74-105">These UI elements look like a natural extension of Office and work across platforms.</span></span> <span data-ttu-id="e9d74-106">You can insert your custom web-based code into any of these elements.</span><span class="sxs-lookup"><span data-stu-id="e9d74-106">You can insert your custom web-based code into any of these elements.</span></span>

<span data-ttu-id="e9d74-107">下图显示了可以创建的 Office UI 元素的类型。</span><span class="sxs-lookup"><span data-stu-id="e9d74-107">The following image shows the types of Office UI elements that you can create.</span></span>

![在 Office 文档的功能区、任务窗格和对话框上显示外接程序命令的图像](../images/add-in-ui-elements.png)

## <a name="add-in-commands"></a><span data-ttu-id="e9d74-109">加载项命令</span><span class="sxs-lookup"><span data-stu-id="e9d74-109">Add-in commands</span></span>

<span data-ttu-id="e9d74-110">使用[外接程序命令](add-in-commands.md)将入口点添加到你的外接程序中的 Office 应用功能区。</span><span class="sxs-lookup"><span data-stu-id="e9d74-110">Use [add-in commands](add-in-commands.md) to add entry points to your add-in to the Office app ribbon.</span></span> <span data-ttu-id="e9d74-111">命令通过运行 JavaScript 代码，或启动 HTML 容器开始在外接程序中操作。</span><span class="sxs-lookup"><span data-stu-id="e9d74-111">Commands start actions in your add-in either by running JavaScript code, or by launching an HTML container.</span></span> <span data-ttu-id="e9d74-112">可以创建以下两种类型的外接程序命令。</span><span class="sxs-lookup"><span data-stu-id="e9d74-112">You can create two types of add-in commands.</span></span>

|<span data-ttu-id="e9d74-113">**命令类型**</span><span class="sxs-lookup"><span data-stu-id="e9d74-113">**Command type**</span></span>|<span data-ttu-id="e9d74-114">**说明**</span><span class="sxs-lookup"><span data-stu-id="e9d74-114">**Description**</span></span>|
|:---------------|:--------------|
|<span data-ttu-id="e9d74-115">功能区按钮、菜单和选项卡</span><span class="sxs-lookup"><span data-stu-id="e9d74-115">Ribbon buttons, menus, and tabs</span></span>|<span data-ttu-id="e9d74-116">Use to add custom buttons, menus (dropdowns), or tabs to the default ribbon in Office.</span><span class="sxs-lookup"><span data-stu-id="e9d74-116">Use to add custom buttons, menus (dropdowns), or tabs to the default ribbon in Office.</span></span> <span data-ttu-id="e9d74-117">Use Buttons and menus to trigger an action in Office.</span><span class="sxs-lookup"><span data-stu-id="e9d74-117">Use Buttons and menus to trigger an action in Office.</span></span> <span data-ttu-id="e9d74-118">Use tabs to group and organize buttons and menus.</span><span class="sxs-lookup"><span data-stu-id="e9d74-118">Use tabs to group and organize buttons and menus.</span></span>|
|<span data-ttu-id="e9d74-119">上下文菜单</span><span class="sxs-lookup"><span data-stu-id="e9d74-119">Context menus</span></span>| <span data-ttu-id="e9d74-120">Use to extend the default context menu.</span><span class="sxs-lookup"><span data-stu-id="e9d74-120">Use to extend the default context menu.</span></span> <span data-ttu-id="e9d74-121">Context menus are displayed when users right-click text in an Office document or a table in Excel.</span><span class="sxs-lookup"><span data-stu-id="e9d74-121">Context menus are displayed when users right-click text in an Office document or a table in Excel.</span></span>| 

## <a name="html-containers"></a><span data-ttu-id="e9d74-122">HTML 容器</span><span class="sxs-lookup"><span data-stu-id="e9d74-122">HTML containers</span></span>

<span data-ttu-id="e9d74-123">Use HTML containers to embed HTML-based UI code within Office clients.</span><span class="sxs-lookup"><span data-stu-id="e9d74-123">Use HTML containers to embed HTML-based UI code within Office clients.</span></span> <span data-ttu-id="e9d74-124">These web pages can then reference the Office JavaScript API to interact with content in the document.</span><span class="sxs-lookup"><span data-stu-id="e9d74-124">These web pages can then reference the Office JavaScript API to interact with content in the document.</span></span> <span data-ttu-id="e9d74-125">You can create three types of HTML containers.</span><span class="sxs-lookup"><span data-stu-id="e9d74-125">You can create three types of HTML containers.</span></span>

|<span data-ttu-id="e9d74-126">**HTML 容器**</span><span class="sxs-lookup"><span data-stu-id="e9d74-126">**HTML container**</span></span>|<span data-ttu-id="e9d74-127">**说明**</span><span class="sxs-lookup"><span data-stu-id="e9d74-127">**Description**</span></span>|
|:-----------------|:--------------|
|[<span data-ttu-id="e9d74-128">任务窗格</span><span class="sxs-lookup"><span data-stu-id="e9d74-128">Task panes</span></span>](task-pane-add-ins.md)|<span data-ttu-id="e9d74-129">Display custom UI in the right pane of the Office document.</span><span class="sxs-lookup"><span data-stu-id="e9d74-129">Display custom UI in the right pane of the Office document.</span></span> <span data-ttu-id="e9d74-130">Use task panes to allow users to interact with your add-in side-by-side with the Office document.</span><span class="sxs-lookup"><span data-stu-id="e9d74-130">Use task panes to allow users to interact with your add-in side-by-side with the Office document.</span></span>|
|[<span data-ttu-id="e9d74-131">内容加载项</span><span class="sxs-lookup"><span data-stu-id="e9d74-131">Content add-ins</span></span>](content-add-ins.md)|<span data-ttu-id="e9d74-132">Display custom UI embedded within Office documents.</span><span class="sxs-lookup"><span data-stu-id="e9d74-132">Display custom UI embedded within Office documents.</span></span> <span data-ttu-id="e9d74-133">Use content add-ins to allow users to interact with your add-in directly within the Office document.</span><span class="sxs-lookup"><span data-stu-id="e9d74-133">Use content add-ins to allow users to interact with your add-in directly within the Office document.</span></span> <span data-ttu-id="e9d74-134">For example, you might want to show external content such as videos or data visualizations from other sources.</span><span class="sxs-lookup"><span data-stu-id="e9d74-134">For example, you might want to show external content such as videos or data visualizations from other sources.</span></span> |
|[<span data-ttu-id="e9d74-135">对话框</span><span class="sxs-lookup"><span data-stu-id="e9d74-135">Dialog boxes</span></span>](dialog-boxes.md)|<span data-ttu-id="e9d74-136">Display custom UI in a dialog box that overlays the Office document.</span><span class="sxs-lookup"><span data-stu-id="e9d74-136">Display custom UI in a dialog box that overlays the Office document.</span></span> <span data-ttu-id="e9d74-137">Use a dialog box for interactions that require focus and more real estate, and do not require a side-by-side interaction with the document.</span><span class="sxs-lookup"><span data-stu-id="e9d74-137">Use a dialog box for interactions that require focus and more real estate, and do not require a side-by-side interaction with the document.</span></span>|

## <a name="see-also"></a><span data-ttu-id="e9d74-138">另请参阅</span><span class="sxs-lookup"><span data-stu-id="e9d74-138">See also</span></span>

- [<span data-ttu-id="e9d74-139">Excel、Word 和 PowerPoint 加载项命令</span><span class="sxs-lookup"><span data-stu-id="e9d74-139">Add-in commands for Excel, Word, and PowerPoint</span></span>](add-in-commands.md)
- [<span data-ttu-id="e9d74-140">任务窗格</span><span class="sxs-lookup"><span data-stu-id="e9d74-140">Task panes</span></span>](task-pane-add-ins.md)
- [<span data-ttu-id="e9d74-141">内容外接程序</span><span class="sxs-lookup"><span data-stu-id="e9d74-141">Content add-ins</span></span>](content-add-ins.md)
- [<span data-ttu-id="e9d74-142">对话框</span><span class="sxs-lookup"><span data-stu-id="e9d74-142">Dialog boxes</span></span>](dialog-boxes.md)
