---
title: Excel?Word ? PowerPoint ???????
description: ''
ms.date: 01/23/2018
ms.openlocfilehash: 42a46bf88cc3f72f94ff5f9162a247d90b33e5c7
ms.sourcegitcommit: c72c35e8389c47a795afbac1b2bcf98c8e216d82
ms.translationtype: HT
ms.contentlocale: zh-CN
ms.lasthandoff: 05/23/2018
---
# <a name="add-in-commands-for-excel-word-and-powerpoint"></a><span data-ttu-id="c7f15-102">Excel?Word ? PowerPoint ?????</span><span class="sxs-lookup"><span data-stu-id="c7f15-102">Add-in commands for Excel, Word, and PowerPoint</span></span>

<span data-ttu-id="c7f15-p101">??????? UI ?????? Office UI????????????????????????????????????????????????????????????????????? JavaScript ???????????????????????????????????????????????????????????????</span><span class="sxs-lookup"><span data-stu-id="c7f15-p101">Add-in commands are UI elements that extend the Office UI and start actions in your add-in. You can use add-in commands to add a button on the ribbon or an item to a context menu. When users select an add-in command, they initiate actions such as running JavaScript code, or showing a page of the add-in in a task pane. Add-in commands help users find and use your add-in, which can help increase your add-in's adoption and reuse, and improve customer retention.</span></span>

<span data-ttu-id="c7f15-107">?????????????? [Office ??????????](https://channel9.msdn.com/events/Build/2016/P551)?</span><span class="sxs-lookup"><span data-stu-id="c7f15-107">For an overview of the feature, see the video [Add-in Commands in the Office Ribbon](https://channel9.msdn.com/events/Build/2016/P551).</span></span>

> [!NOTE]
> <span data-ttu-id="c7f15-p102">SharePoint ???????????????[????](../publish/centralized-deployment.md)? [AppSource](https://docs.microsoft.com/en-us/office/dev/store/submit-to-the-office-store) ?????????????[???](../testing/create-a-network-shared-folder-catalog-for-task-pane-and-content-add-ins.md)????????????</span><span class="sxs-lookup"><span data-stu-id="c7f15-p102">SharePoint catalogs do not support add-in commands. You can deploy add-in commands via [Centralized Deployment](../publish/centralized-deployment.md) or [AppSource](https://docs.microsoft.com/en-us/office/dev/store/submit-to-the-office-store), or use [sideloading](../testing/create-a-network-shared-folder-catalog-for-task-pane-and-content-add-ins.md) to deploy your add-in command for testing.</span></span> 

<span data-ttu-id="c7f15-110">*? 1?? Excel Desktop ?????????*</span><span class="sxs-lookup"><span data-stu-id="c7f15-110">*Figure 1. Add-in with commands running in Excel Desktop*</span></span>

![Excel ???????????](../images/add-in-commands-1.png)

<span data-ttu-id="c7f15-112">*? 2?? Excel Online ?????????*</span><span class="sxs-lookup"><span data-stu-id="c7f15-112">*Figure 2. Add-in with commands running in Excel Online*</span></span>

![Excel Online ???????????](../images/add-in-commands-2.png)

## <a name="command-capabilities"></a><span data-ttu-id="c7f15-114">????</span><span class="sxs-lookup"><span data-stu-id="c7f15-114">Command capabilities</span></span>
<span data-ttu-id="c7f15-115">???????????</span><span class="sxs-lookup"><span data-stu-id="c7f15-115">The following command capabilities are currently supported.</span></span>

> [!NOTE]
> <span data-ttu-id="c7f15-116">???????????????</span><span class="sxs-lookup"><span data-stu-id="c7f15-116">Content add-ins do not currently support add-in commands.</span></span>

<span data-ttu-id="c7f15-117">**???**</span><span class="sxs-lookup"><span data-stu-id="c7f15-117">**Extension points**</span></span>

- <span data-ttu-id="c7f15-118">?????? - ?????????????????</span><span class="sxs-lookup"><span data-stu-id="c7f15-118">Ribbon tabs - Extend built-in tabs or create a new custom tab.</span></span>
- <span data-ttu-id="c7f15-119">????? - ??????????</span><span class="sxs-lookup"><span data-stu-id="c7f15-119">Context menus - Extend selected context menus.</span></span> 

<span data-ttu-id="c7f15-120">**????**</span><span class="sxs-lookup"><span data-stu-id="c7f15-120">**Control types**</span></span>

- <span data-ttu-id="c7f15-121">???? - ???????</span><span class="sxs-lookup"><span data-stu-id="c7f15-121">Simple buttons - trigger specific actions.</span></span>
- <span data-ttu-id="c7f15-122">?? - ???????????????????</span><span class="sxs-lookup"><span data-stu-id="c7f15-122">Menus - simple menu dropdown with buttons that trigger actions.</span></span>

<span data-ttu-id="c7f15-123">**??**</span><span class="sxs-lookup"><span data-stu-id="c7f15-123">**Actions**</span></span>

- <span data-ttu-id="c7f15-124">ShowTaskpane - ??????????????? HTML ?????</span><span class="sxs-lookup"><span data-stu-id="c7f15-124">ShowTaskpane - Displays one or multiple panes that load custom HTML pages inside them.</span></span>
- <span data-ttu-id="c7f15-p103">ExecuteFunction - ???????? HTML ??????????? JavaScript ??????????????????????????? UI?????? [displayDialog](http://dev.office.com/reference/add-ins/shared/officeui) API?</span><span class="sxs-lookup"><span data-stu-id="c7f15-p103">ExecuteFunction - Loads an invisible HTML page and then execute a JavaScript function within it. To show UI within your function (such as errors, progress, or additional input) you can use the [displayDialog](http://dev.office.com/reference/add-ins/shared/officeui) API.</span></span>  

## <a name="supported-platforms"></a><span data-ttu-id="c7f15-127">?????</span><span class="sxs-lookup"><span data-stu-id="c7f15-127">Supported platforms</span></span>
<span data-ttu-id="c7f15-128">????????????????</span><span class="sxs-lookup"><span data-stu-id="c7f15-128">Add-in commands are currently supported on the following platforms:</span></span>

- <span data-ttu-id="c7f15-129">Office for Windows Desktop 2016????? 16.0.6769+?</span><span class="sxs-lookup"><span data-stu-id="c7f15-129">Office for Windows Desktop 2016 (build 16.0.6769+)</span></span>
- <span data-ttu-id="c7f15-130">Office for Mac????? 15.33+?</span><span class="sxs-lookup"><span data-stu-id="c7f15-130">Office for Mac (build 15.33+)</span></span>
- <span data-ttu-id="c7f15-131">Office Online</span><span class="sxs-lookup"><span data-stu-id="c7f15-131">Office Online</span></span> 

<span data-ttu-id="c7f15-132">?????????????</span><span class="sxs-lookup"><span data-stu-id="c7f15-132">More platforms are coming soon.</span></span>

## <a name="best-practices"></a><span data-ttu-id="c7f15-133">????</span><span class="sxs-lookup"><span data-stu-id="c7f15-133">Best practices</span></span>

<span data-ttu-id="c7f15-134">????????????????????</span><span class="sxs-lookup"><span data-stu-id="c7f15-134">Apply the following best practices when you develop add-in commands:</span></span>

- <span data-ttu-id="c7f15-p104">????????????????????????????????????????</span><span class="sxs-lookup"><span data-stu-id="c7f15-p104">Use commands to represent a specific action with a clear and specific outcome for users. Do not combine multiple actions in a single button.</span></span>
- <span data-ttu-id="c7f15-p105">?????????????????????????????????????????</span><span class="sxs-lookup"><span data-stu-id="c7f15-p105">Provide granular actions that make common tasks within your add-in more efficient to perform. Minimize the number of steps an action takes to complete.</span></span>
- <span data-ttu-id="c7f15-139">????? Office ????????</span><span class="sxs-lookup"><span data-stu-id="c7f15-139">For the placement of your commands in the Office ribbon:</span></span>
    - <span data-ttu-id="c7f15-p106">???????????????????????????????????????????????????????????????????????????????????? Office ????????????????? [Office ???? XML ??](../develop/add-in-manifests.md)?</span><span class="sxs-lookup"><span data-stu-id="c7f15-p106">Place commands on an existing tab (Insert, Review, and so on) if the functionality provided fits there. For example, if your add-in enables users to insert media, add a group to the Insert tab. Note that not all tabs are available across all Office versions. For more information, see [Office Add-ins XML manifest](../develop/add-in-manifests.md).</span></span> 
    - <span data-ttu-id="c7f15-p107">???????????????????? 6 ??????????????????????????????? Office ???? Office Desktop ? Office Online?????????????????????????????????? Office Online ????????????????????????</span><span class="sxs-lookup"><span data-stu-id="c7f15-p107">Place commands on the Home tab if the functionality doesn't fit on another tab, and you have fewer than six top-level commands. You can also add commands to the Home tab if your add-in needs to work across Office versions (such as Office Desktop and Office Online) and a tab is not available in all versions (for example, the Design tab doesn't exist in Office Online).</span></span>  
    - <span data-ttu-id="c7f15-145">????? 6 ?????????????????????????</span><span class="sxs-lookup"><span data-stu-id="c7f15-145">Place commands on a custom tab if you have more than six top-level commands.</span></span> 
    - <span data-ttu-id="c7f15-p108">??????????????????????????????????????????????????</span><span class="sxs-lookup"><span data-stu-id="c7f15-p108">Name your group to match the name of your add-in. If you have multiple groups, name each group based on the functionality that the commands in that group provide.</span></span>
    - <span data-ttu-id="c7f15-148">?????????????????????????</span><span class="sxs-lookup"><span data-stu-id="c7f15-148">Do not add superfluous buttons to increase the real estate of your add-in.</span></span>

     > [!NOTE]
     > <span data-ttu-id="c7f15-149">???????????????? [AppSource ??](https://docs.microsoft.com/en-us/office/dev/store/validation-policies)?</span><span class="sxs-lookup"><span data-stu-id="c7f15-149">Add-ins that take up too much space might not pass [AppSource validation](https://docs.microsoft.com/en-us/office/dev/store/validation-policies).</span></span>

- <span data-ttu-id="c7f15-150">??????????[??????](design-icons.md)?</span><span class="sxs-lookup"><span data-stu-id="c7f15-150">For all icons, follow the [icon design guidelines](design-icons.md).</span></span>
- <span data-ttu-id="c7f15-p109">????????????????????????????????????????????????????????????????????</span><span class="sxs-lookup"><span data-stu-id="c7f15-p109">Provide a version of your add-in that also works on hosts that do not support commands. A single add-in manifest can work in both command-aware (with commands) and non-command-aware (as a taskpane) hosts.</span></span>

   <span data-ttu-id="c7f15-153">*? 3?Office 2013 ???????????? Office 2016 ??????????????*</span><span class="sxs-lookup"><span data-stu-id="c7f15-153">*Figure 3. Task pane add-in in Office 2013 and the same add-in using add-in commands in Office 2016*</span></span>

   ![?? Office 2013 ???????????? Office 2016 ???????????????????](../images/office-task-pane-add-ins.png)


## <a name="next-steps"></a><span data-ttu-id="c7f15-155">????</span><span class="sxs-lookup"><span data-stu-id="c7f15-155">Next steps</span></span>

<span data-ttu-id="c7f15-156">??????????????? GitHub ?? [Office ???????](https://github.com/OfficeDev/Office-Add-in-Commands-Samples/)?</span><span class="sxs-lookup"><span data-stu-id="c7f15-156">The best way to get started using add-in commands is to take a look at the [Office Add-in commands samples](https://github.com/OfficeDev/Office-Add-in-Commands-Samples/) on GitHub.</span></span>

<span data-ttu-id="c7f15-157">???????????????????????[???????????](../develop/create-addin-commands.md)? [VersionOverrides](https://dev.office.com/reference/add-ins/manifest/versionoverrides) ?????</span><span class="sxs-lookup"><span data-stu-id="c7f15-157">For more information about specifying add-in commands in your manifest, see [Create add-in commands in your manifest](../develop/create-addin-commands.md) and the [VersionOverrides](https://dev.office.com/reference/add-ins/manifest/versionoverrides) reference content.</span></span>




