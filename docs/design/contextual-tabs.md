---
title: 在加载项中Office上下文选项卡
description: 了解如何将自定义上下文选项卡添加到Office外接程序。
ms.date: 07/15/2021
localization_priority: Normal
ms.openlocfilehash: a8eaffe0402601ee11a063d0df5670ff208be4fd
ms.sourcegitcommit: b20041962a7f921a8c40eb9ae55bc6992450b243
ms.translationtype: MT
ms.contentlocale: zh-CN
ms.lasthandoff: 07/16/2021
ms.locfileid: "53456227"
---
# <a name="create-custom-contextual-tabs-in-office-add-ins"></a><span data-ttu-id="38709-103">在加载项中Office上下文选项卡</span><span class="sxs-lookup"><span data-stu-id="38709-103">Create custom contextual tabs in Office Add-ins</span></span>

<span data-ttu-id="38709-104">上下文选项卡是活动功能区中隐藏的选项卡Office，当文档文档中发生指定事件时，该选项卡Office显示。</span><span class="sxs-lookup"><span data-stu-id="38709-104">A contextual tab is a hidden tab control in the Office ribbon that is displayed in the tab row when a specified event occurs in the Office document.</span></span> <span data-ttu-id="38709-105">例如 **，当选择** 表时，Excel功能区上出现的"表设计"选项卡。</span><span class="sxs-lookup"><span data-stu-id="38709-105">For example, the **Table Design** tab that appears on the Excel ribbon when a table is selected.</span></span> <span data-ttu-id="38709-106">在加载项中添加自定义Office，并创建更改可见性的事件处理程序，指定它们何时可见或隐藏。</span><span class="sxs-lookup"><span data-stu-id="38709-106">You include custom contextual tabs in your Office Add-in and specify when they are visible or hidden, by creating event handlers that change the visibility.</span></span> <span data-ttu-id="38709-107"> (但是，自定义上下文选项卡不会响应焦点更改。) </span><span class="sxs-lookup"><span data-stu-id="38709-107">(However, custom contextual tabs do not respond to focus changes.)</span></span>

> [!NOTE]
> <span data-ttu-id="38709-108">本文假定你熟悉以下文档。</span><span class="sxs-lookup"><span data-stu-id="38709-108">This article assumes that you are familiar with the following documentation.</span></span> <span data-ttu-id="38709-109">如果你最近未使用加载项命令（自定义菜单项和功能区按钮），请查看该文档。</span><span class="sxs-lookup"><span data-stu-id="38709-109">Please review it if you haven't worked with Add-in Commands (custom menu items and ribbon buttons) recently.</span></span>
>
> - [<span data-ttu-id="38709-110">加载项命令的基本概念</span><span class="sxs-lookup"><span data-stu-id="38709-110">Basic concepts for Add-in Commands</span></span>](add-in-commands.md)

> [!IMPORTANT]
> <span data-ttu-id="38709-111">自定义上下文选项卡当前仅在以下Excel且仅在以下平台和内部版本上受支持：</span><span class="sxs-lookup"><span data-stu-id="38709-111">Custom contextual tabs are currently only supported on Excel and only on these platforms and builds:</span></span>
>
> - <span data-ttu-id="38709-112">Excel订阅Windows (Microsoft 365版本) 版本 2102 (内部版本 13801.20294) 或更高版本。</span><span class="sxs-lookup"><span data-stu-id="38709-112">Excel on Windows (Microsoft 365 subscription only): Version 2102 (Build 13801.20294) or later.</span></span>
> - <span data-ttu-id="38709-113">Excel 网页版</span><span class="sxs-lookup"><span data-stu-id="38709-113">Excel on the web</span></span>

> [!NOTE]
> <span data-ttu-id="38709-114">自定义上下文选项卡仅适用于支持以下要求集的平台。</span><span class="sxs-lookup"><span data-stu-id="38709-114">Custom contextual tabs work only on platforms that support the following requirement sets.</span></span> <span data-ttu-id="38709-115">有关要求集以及如何使用它们，请参阅指定Office应用程序和 API[要求](../develop/specify-office-hosts-and-api-requirements.md)。</span><span class="sxs-lookup"><span data-stu-id="38709-115">For more about requirement sets and how to work with them, see [Specify Office applications and API requirements](../develop/specify-office-hosts-and-api-requirements.md).</span></span>
>
> - [<span data-ttu-id="38709-116">RibbonApi 1.2</span><span class="sxs-lookup"><span data-stu-id="38709-116">RibbonApi 1.2</span></span>](../reference/requirement-sets/ribbon-api-requirement-sets.md)
> - [<span data-ttu-id="38709-117">SharedRuntime 1.1</span><span class="sxs-lookup"><span data-stu-id="38709-117">SharedRuntime 1.1</span></span>](../reference/requirement-sets/shared-runtime-requirement-sets.md)
>
> <span data-ttu-id="38709-118">可以在代码中使用运行时检查来测试用户的主机和平台组合是否支持这些要求集，如指定 Office[应用程序和 API 要求中所述](../develop/specify-office-hosts-and-api-requirements.md#use-runtime-checks-in-your-javascript-code)。</span><span class="sxs-lookup"><span data-stu-id="38709-118">You can use the runtime checks in your code to test whether the user's host and platform combination supports these requirement sets as described in [Specify Office applications and API requirements](../develop/specify-office-hosts-and-api-requirements.md#use-runtime-checks-in-your-javascript-code).</span></span> <span data-ttu-id="38709-119"> (在清单中指定要求集的技术（也如该文章中所述）当前对 RibbonApi 1.2.) 不起作用。或者，当自定义上下文选项卡不受支持时，你可以实现备用 [UI 体验](#implement-an-alternate-ui-experience-when-custom-contextual-tabs-are-not-supported)。</span><span class="sxs-lookup"><span data-stu-id="38709-119">(The technique of specifying the requirement sets in the manifest, which is also described in that article, does not currently work for RibbonApi 1.2.) Alternatively, you can [implement an alternate UI experience when custom contextual tabs are not supported](#implement-an-alternate-ui-experience-when-custom-contextual-tabs-are-not-supported).</span></span>

## <a name="behavior-of-custom-contextual-tabs"></a><span data-ttu-id="38709-120">自定义上下文选项卡的行为</span><span class="sxs-lookup"><span data-stu-id="38709-120">Behavior of custom contextual tabs</span></span>

<span data-ttu-id="38709-121">自定义上下文选项卡的用户体验遵循内置上下文选项卡Office模式。</span><span class="sxs-lookup"><span data-stu-id="38709-121">The user experience for custom contextual tabs follows the pattern of built-in Office contextual tabs.</span></span> <span data-ttu-id="38709-122">以下是放置自定义上下文选项卡的基本原则。</span><span class="sxs-lookup"><span data-stu-id="38709-122">The following are the basic principles for the placement custom contextual tabs.</span></span>

- <span data-ttu-id="38709-123">当自定义上下文选项卡可见时，它将显示在功能区的右端。</span><span class="sxs-lookup"><span data-stu-id="38709-123">When a custom contextual tab is visible, it appears on the right end of the ribbon.</span></span>
- <span data-ttu-id="38709-124">如果同时显示外接程序中的一个或多个内置上下文选项卡和一个或多个自定义上下文选项卡，则自定义上下文选项卡始终位于所有内置上下文选项卡的右侧。</span><span class="sxs-lookup"><span data-stu-id="38709-124">If one or more built-in contextual tabs and one or more custom contextual tabs from add-ins are visible at the same time, the custom contextual tabs are always to the right of all of the built-in contextual tabs.</span></span>
- <span data-ttu-id="38709-125">如果您的外接程序具有多个上下文选项卡，并且存在多个上下文，并且存在多个上下文可见，则它们按照在外接程序中定义的顺序显示。</span><span class="sxs-lookup"><span data-stu-id="38709-125">If your add-in has more than one contextual tab and there are contexts in which more than one is visible, they appear in the order in which they are defined in your add-in.</span></span> <span data-ttu-id="38709-126"> (方向与语言的方向Office相同;即，在从左到右语言中为从左到右，但从右到左使用从右到左的语言。) 请参阅定义选项卡上出现的组和[](#define-the-groups-and-controls-that-appear-on-the-tab)控件，详细了解如何定义它们。</span><span class="sxs-lookup"><span data-stu-id="38709-126">(The direction is the same direction as the Office language; that is, is left-to-right in left-to-right languages, but right-to-left in right-to-left languages.) See [Define the groups and controls that appear on the tab](#define-the-groups-and-controls-that-appear-on-the-tab) for details about how you define them.</span></span>
- <span data-ttu-id="38709-127">如果多个加载项具有特定上下文中可见的上下文选项卡，则它们按加载项的启动顺序显示。</span><span class="sxs-lookup"><span data-stu-id="38709-127">If more than one add-in has a contextual tab that is visible in a specific context, then they appear in the order in which the add-ins were launched.</span></span>
- <span data-ttu-id="38709-128">与 *自定义* 核心选项卡不同，自定义上下文选项卡不会Office应用程序的功能区。</span><span class="sxs-lookup"><span data-stu-id="38709-128">Custom *contextual* tabs, unlike custom core tabs, are not added permanently to the Office application's ribbon.</span></span> <span data-ttu-id="38709-129">它们仅存在于Office运行加载项的文档中。</span><span class="sxs-lookup"><span data-stu-id="38709-129">They are present only in Office documents on which your add-in is running.</span></span>

## <a name="major-steps-for-including-a-contextual-tab-in-an-add-in"></a><span data-ttu-id="38709-130">在外接程序中添加上下文选项卡的主要步骤</span><span class="sxs-lookup"><span data-stu-id="38709-130">Major steps for including a contextual tab in an add-in</span></span>

<span data-ttu-id="38709-131">以下是在外接程序中添加自定义上下文选项卡的主要步骤。</span><span class="sxs-lookup"><span data-stu-id="38709-131">The following are the major steps for including a custom contextual tab in an add-in.</span></span>

1. <span data-ttu-id="38709-132">将外接程序配置为使用共享运行时。</span><span class="sxs-lookup"><span data-stu-id="38709-132">Configure the add-in to use a shared runtime.</span></span>
1. <span data-ttu-id="38709-133">定义选项卡及其上出现的组和控件。</span><span class="sxs-lookup"><span data-stu-id="38709-133">Define the tab and the groups and controls that appear on it.</span></span>
1. <span data-ttu-id="38709-134">使用"上下文"选项卡注册Office。</span><span class="sxs-lookup"><span data-stu-id="38709-134">Register the contextual tab with Office.</span></span>
1. <span data-ttu-id="38709-135">指定选项卡可见时的情况。</span><span class="sxs-lookup"><span data-stu-id="38709-135">Specify the circumstances when the tab will be visible.</span></span>

## <a name="configure-the-add-in-to-use-a-shared-runtime"></a><span data-ttu-id="38709-136">将外接程序配置为使用共享运行时</span><span class="sxs-lookup"><span data-stu-id="38709-136">Configure the add-in to use a shared runtime</span></span>

<span data-ttu-id="38709-137">添加自定义上下文选项卡要求加载项使用共享运行时。</span><span class="sxs-lookup"><span data-stu-id="38709-137">Adding custom contextual tabs requires your add-in to use the shared runtime.</span></span> <span data-ttu-id="38709-138">有关详细信息，请参阅 [将加载项配置为使用共享运行时](../develop/configure-your-add-in-to-use-a-shared-runtime.md)。</span><span class="sxs-lookup"><span data-stu-id="38709-138">For more information, see [Configure an add-in to use a shared runtime](../develop/configure-your-add-in-to-use-a-shared-runtime.md).</span></span>

## <a name="define-the-groups-and-controls-that-appear-on-the-tab"></a><span data-ttu-id="38709-139">定义显示在选项卡上的组和控件</span><span class="sxs-lookup"><span data-stu-id="38709-139">Define the groups and controls that appear on the tab</span></span>

<span data-ttu-id="38709-140">与在清单中用 XML 定义的自定义核心选项卡不同，自定义上下文选项卡是在运行时使用 JSON blob 定义的。</span><span class="sxs-lookup"><span data-stu-id="38709-140">Unlike custom core tabs, which are defined with XML in the manifest, custom contextual tabs are defined at runtime with a JSON blob.</span></span> <span data-ttu-id="38709-141">代码将 blob 解析为 JavaScript 对象，然后将该对象传递给[Office.ribbon.requestCreateControls](/javascript/api/office/office.ribbon?view=common-js&preserve-view=true#requestCreateControls-tabDefinition-)方法。</span><span class="sxs-lookup"><span data-stu-id="38709-141">Your code parses the blob into a JavaScript object, and then passes the object to the [Office.ribbon.requestCreateControls](/javascript/api/office/office.ribbon?view=common-js&preserve-view=true#requestCreateControls-tabDefinition-) method.</span></span> <span data-ttu-id="38709-142">自定义上下文选项卡仅存在于加载项当前运行的文档中。</span><span class="sxs-lookup"><span data-stu-id="38709-142">Custom contextual tabs are only present in documents on which your add-in is currently running.</span></span> <span data-ttu-id="38709-143">这不同于在安装加载项时添加到 Office 应用程序功能区中的自定义核心选项卡，当打开另一个文档时，这些选项卡仍保持显示状态。</span><span class="sxs-lookup"><span data-stu-id="38709-143">This is different from custom core tabs which are added to the Office application ribbon when the add-in is installed and remain present when another document is opened.</span></span> <span data-ttu-id="38709-144">此外 `requestCreateControls` ，方法只能在加载项会话中运行一次。</span><span class="sxs-lookup"><span data-stu-id="38709-144">Also, the `requestCreateControls` method may be run only once in a session of your add-in.</span></span> <span data-ttu-id="38709-145">如果再次调用它，将引发错误。</span><span class="sxs-lookup"><span data-stu-id="38709-145">If it is called again, an error is thrown.</span></span>

> [!NOTE]
> <span data-ttu-id="38709-146">JSON blob 的属性和子属性 (和键名称) 的结构大致与清单 XML 中 [CustomTab](../reference/manifest/customtab.md) 元素及其后代元素的结构平行。</span><span class="sxs-lookup"><span data-stu-id="38709-146">The structure of the JSON blob's properties and subproperties (and the key names) is roughly parallel to the structure of the [CustomTab](../reference/manifest/customtab.md) element and its descendant elements in the manifest XML.</span></span>

<span data-ttu-id="38709-147">我们将分步构造上下文选项卡 JSON blob 的示例。</span><span class="sxs-lookup"><span data-stu-id="38709-147">We'll construct an example of a contextual tabs JSON blob step-by-step.</span></span> <span data-ttu-id="38709-148">上下文选项卡 JSON 的完整架构位于[上的dynamic-ribbon.schema.js。](https://developer.microsoft.com/json-schemas/office-js/dynamic-ribbon.schema.json)</span><span class="sxs-lookup"><span data-stu-id="38709-148">The full schema for the contextual tab JSON is at [dynamic-ribbon.schema.json](https://developer.microsoft.com/json-schemas/office-js/dynamic-ribbon.schema.json).</span></span> <span data-ttu-id="38709-149">如果你使用 Visual Studio Code，可以使用此文件获取IntelliSense并验证 JSON。</span><span class="sxs-lookup"><span data-stu-id="38709-149">If you are working in Visual Studio Code, you can use this file to get IntelliSense and to validate your JSON.</span></span> <span data-ttu-id="38709-150">有关详细信息，请参阅使用[JSON 架构Visual Studio Code JSON 编辑 JSON。](https://code.visualstudio.com/docs/languages/json#_json-schemas-and-settings)</span><span class="sxs-lookup"><span data-stu-id="38709-150">For more information, see [Editing JSON with Visual Studio Code - JSON schemas and settings](https://code.visualstudio.com/docs/languages/json#_json-schemas-and-settings).</span></span>


1. <span data-ttu-id="38709-151">首先，创建包含名为 和 的两个数组属性的 JSON `actions` 字符串 `tabs` 。</span><span class="sxs-lookup"><span data-stu-id="38709-151">Begin by creating a JSON string with two array properties named `actions` and `tabs`.</span></span> <span data-ttu-id="38709-152">`actions`数组是上下文选项卡上的控件可以执行的所有函数的规范。数组 `tabs` 定义一个或多个上下文选项卡，*最多 20 个*。</span><span class="sxs-lookup"><span data-stu-id="38709-152">The `actions` array is a specification of all the functions that can be executed by controls on the contextual tab. The `tabs` array defines one or more contextual tabs, *up to a maximum of 20*.</span></span>

    ```json
    '{
      "actions": [

      ],
      "tabs": [

      ]
    }'
    ```

1. <span data-ttu-id="38709-153">这个上下文选项卡的简单示例将只有一个按钮，因此只有一个操作。</span><span class="sxs-lookup"><span data-stu-id="38709-153">This simple example of a contextual tab will have only a single button and, thus, only a single action.</span></span> <span data-ttu-id="38709-154">将以下内容添加为数组的唯一 `actions` 成员。</span><span class="sxs-lookup"><span data-stu-id="38709-154">Add the following as the only member of the `actions` array.</span></span> <span data-ttu-id="38709-155">关于此标记，请注意：</span><span class="sxs-lookup"><span data-stu-id="38709-155">About this markup, note:</span></span>

    - <span data-ttu-id="38709-156">`id`和 `type` 属性是必需的。</span><span class="sxs-lookup"><span data-stu-id="38709-156">The `id` and `type` properties are mandatory.</span></span>
    - <span data-ttu-id="38709-157">的值 `type` 可以是"ExecuteFunction"或"ShowTaskpane"。</span><span class="sxs-lookup"><span data-stu-id="38709-157">The value of `type` can be either "ExecuteFunction" or "ShowTaskpane".</span></span>
    - <span data-ttu-id="38709-158">`functionName`只有当 的值为 时，才使用 `type` 属性 `ExecuteFunction` 。</span><span class="sxs-lookup"><span data-stu-id="38709-158">The `functionName` property is only used when the value of `type` is `ExecuteFunction`.</span></span> <span data-ttu-id="38709-159">它是 FunctionFile 中定义的函数的名称。</span><span class="sxs-lookup"><span data-stu-id="38709-159">It is the name of a function defined in the FunctionFile.</span></span> <span data-ttu-id="38709-160">有关 FunctionFile 的信息，请参阅 [加载项命令的基本概念](add-in-commands.md)。</span><span class="sxs-lookup"><span data-stu-id="38709-160">For more information about the FunctionFile, see [Basic concepts for Add-in Commands](add-in-commands.md).</span></span>
    - <span data-ttu-id="38709-161">在稍后的步骤中，您将此操作映射到上下文选项卡上的按钮。</span><span class="sxs-lookup"><span data-stu-id="38709-161">In a later step, you will map this action to a button on the contextual tab.</span></span>

    ```json
    {
      "id": "executeWriteData",
      "type": "ExecuteFunction",
      "functionName": "writeData"
    }
   ```

1. <span data-ttu-id="38709-162">将以下内容添加为数组的唯一 `tabs` 成员。</span><span class="sxs-lookup"><span data-stu-id="38709-162">Add the following as the only member of the `tabs` array.</span></span> <span data-ttu-id="38709-163">关于此标记，请注意：</span><span class="sxs-lookup"><span data-stu-id="38709-163">About this markup, note:</span></span>

    - <span data-ttu-id="38709-164">`id` 属性是必需的。</span><span class="sxs-lookup"><span data-stu-id="38709-164">The `id` property is required.</span></span> <span data-ttu-id="38709-165">使用外接程序中所有上下文选项卡中唯一的简短描述性 ID。</span><span class="sxs-lookup"><span data-stu-id="38709-165">Use a brief, descriptive ID that is unique among all contextual tabs in your add-in.</span></span>
    - <span data-ttu-id="38709-166">`label` 属性是必需的。</span><span class="sxs-lookup"><span data-stu-id="38709-166">The `label` property is required.</span></span> <span data-ttu-id="38709-167">它是一个用户友好字符串，用作上下文选项卡的标签。</span><span class="sxs-lookup"><span data-stu-id="38709-167">It is a user-friendly string to serve as the label of the contextual tab.</span></span>
    - <span data-ttu-id="38709-168">`groups` 属性是必需的。</span><span class="sxs-lookup"><span data-stu-id="38709-168">The `groups` property is required.</span></span> <span data-ttu-id="38709-169">它定义将显示在选项卡上的控件组。它必须至少有一个成员且不超过 *20* 个。</span><span class="sxs-lookup"><span data-stu-id="38709-169">It defines the groups of controls that will appear on the tab. It must have at least one member *and no more than 20*.</span></span> <span data-ttu-id="38709-170"> (自定义上下文选项卡上可以具有的控件数量也有限制，这也会限制你拥有多少个组。</span><span class="sxs-lookup"><span data-stu-id="38709-170">(There are also limits on the number of controls that you can have on a custom contextual tab and that will also constrain how many groups that you have.</span></span> <span data-ttu-id="38709-171">有关详细信息，请参阅下一步。) </span><span class="sxs-lookup"><span data-stu-id="38709-171">See the next step for more information.)</span></span>

    > [!NOTE]
    > <span data-ttu-id="38709-172">tab 对象还可以具有一个可选属性，该属性指定在外接程序启动时选项卡 `visible` 是否立即可见。</span><span class="sxs-lookup"><span data-stu-id="38709-172">The tab object can also have an optional `visible` property that specifies whether the tab is visible immediately when the add-in starts up.</span></span> <span data-ttu-id="38709-173">由于上下文选项卡通常是隐藏的，直到用户事件触发其可见性 (例如用户在文档中选择某种类型的实体) 因此当不存在时，该属性默认为 `visible` `false` 。</span><span class="sxs-lookup"><span data-stu-id="38709-173">Since contextual tabs are normally hidden until a user event triggers their visibility (such as the user selecting an entity of some type in the document), the `visible` property defaults to `false` when not present.</span></span> <span data-ttu-id="38709-174">在稍后的部分中，我们将展示如何设置 属性 `true` 以响应事件。</span><span class="sxs-lookup"><span data-stu-id="38709-174">In a later section, we show how to set the property to `true` in response to an event.</span></span>

    ```json
    {
      "id": "CtxTab1",
      "label": "Contoso Data",
      "groups": [

      ]
    }
    ```

1. <span data-ttu-id="38709-175">在简单正在进行的示例中，上下文选项卡只有一个组。</span><span class="sxs-lookup"><span data-stu-id="38709-175">In the simple ongoing example, the contextual tab has only a single group.</span></span> <span data-ttu-id="38709-176">将以下内容添加为数组的唯一 `groups` 成员。</span><span class="sxs-lookup"><span data-stu-id="38709-176">Add the following as the only member of the `groups` array.</span></span> <span data-ttu-id="38709-177">关于此标记，请注意：</span><span class="sxs-lookup"><span data-stu-id="38709-177">About this markup, note:</span></span>

    - <span data-ttu-id="38709-178">所有属性都是必需的。</span><span class="sxs-lookup"><span data-stu-id="38709-178">All the properties are required.</span></span>
    - <span data-ttu-id="38709-179">属性在选项卡的所有组中必须是唯一的。 `id` 请使用简短的描述性 ID。</span><span class="sxs-lookup"><span data-stu-id="38709-179">The `id` property must be unique among all the groups in the tab. Use a brief, descriptive ID.</span></span>
    - <span data-ttu-id="38709-180">`label`是用作组标签的用户友好字符串。</span><span class="sxs-lookup"><span data-stu-id="38709-180">The `label` is a user-friendly string to serve as the label of the group.</span></span>
    - <span data-ttu-id="38709-181">属性的值是一组对象，这些对象根据功能区的大小以及应用程序窗口的功能区Office `icon` 图标。</span><span class="sxs-lookup"><span data-stu-id="38709-181">The `icon` property's value is an array of objects that specify the icons that the group will have on the ribbon depending on the size of the ribbon and the Office application window.</span></span>
    - <span data-ttu-id="38709-182">`controls`属性的值是指定组中按钮和菜单的对象数组。</span><span class="sxs-lookup"><span data-stu-id="38709-182">The `controls` property's value is an array of objects that specify the buttons and menus in the group.</span></span> <span data-ttu-id="38709-183">必须至少有一个。</span><span class="sxs-lookup"><span data-stu-id="38709-183">There must be at least one.</span></span>

    > [!IMPORTANT]
    > <span data-ttu-id="38709-184">*整个选项卡上的控件总数不能超过 20 个。*</span><span class="sxs-lookup"><span data-stu-id="38709-184">*The total number of controls on the whole tab can be no more than 20.*</span></span> <span data-ttu-id="38709-185">例如，可以有 3 个组，每个组有 6 个控件，第四个组有 2 个控件，但不能有 4 个组，每个组有 6 个控件。</span><span class="sxs-lookup"><span data-stu-id="38709-185">For example, you could have 3 groups with 6 controls each, and a fourth group with 2 controls, but you cannot have 4 groups with 6 controls each.</span></span>  

    ```json
    {
        "id": "CustomGroup111",
        "label": "Insertion",
        "icon": [

        ],
        "controls": [

        ]
    }
    ```

1. <span data-ttu-id="38709-186">每个组必须具有至少两个大小的图标：32x32 像素和 80x80 像素。</span><span class="sxs-lookup"><span data-stu-id="38709-186">Every group must have an icon of at least two sizes, 32x32 px and 80x80 px.</span></span> <span data-ttu-id="38709-187">（可选）还可以具有大小为 16x16 像素、20x20 像素、24x24 像素、40x40 像素、48x48 像素和 64x64 像素的图标。</span><span class="sxs-lookup"><span data-stu-id="38709-187">Optionally, you can also have icons of sizes 16x16 px, 20x20 px, 24x24 px, 40x40 px, 48x48 px, and 64x64 px.</span></span> <span data-ttu-id="38709-188">Office根据功能区的大小和应用程序窗口的大小决定Office图标。</span><span class="sxs-lookup"><span data-stu-id="38709-188">Office decides which icon to use based on the size of the ribbon and Office application window.</span></span> <span data-ttu-id="38709-189">将以下对象添加到图标数组。</span><span class="sxs-lookup"><span data-stu-id="38709-189">Add the following objects to the icon array.</span></span> <span data-ttu-id="38709-190"> (如果窗口和功能区的大小足以使组中至少有一个控件显示，则不显示任何组图标。</span><span class="sxs-lookup"><span data-stu-id="38709-190">(If the window and ribbon sizes are large enough for at least one of the *controls* on the group to appear, then no group icon at all appears.</span></span> <span data-ttu-id="38709-191">例如，在缩小和展开 Word 窗口时，观察 Word 功能区上的 **"** 样式"组) 关于此标记，请注意：</span><span class="sxs-lookup"><span data-stu-id="38709-191">For an example, watch the **Styles** group on the Word ribbon as you shrink and expand the Word window.) About this markup, note:</span></span>

    - <span data-ttu-id="38709-192">这两个属性都是必需的。</span><span class="sxs-lookup"><span data-stu-id="38709-192">Both the properties are required.</span></span>
    - <span data-ttu-id="38709-193">`size`属性度量单位是像素。</span><span class="sxs-lookup"><span data-stu-id="38709-193">The `size` property unit of measure is pixels.</span></span> <span data-ttu-id="38709-194">图标始终为正方形，因此数字同时为高度和宽度。</span><span class="sxs-lookup"><span data-stu-id="38709-194">Icons are always square, so the number is both the height and the width.</span></span>
    - <span data-ttu-id="38709-195">`sourceLocation`属性指定图标的完整 URL。</span><span class="sxs-lookup"><span data-stu-id="38709-195">The `sourceLocation` property specifies the full URL to the icon.</span></span>

    > [!IMPORTANT]
    > <span data-ttu-id="38709-196">与从开发环境移动到生产 (（如将域从 localhost 更改为 contoso.com) ）时，通常必须更改加载项清单中的 URL 一样，您还必须更改上下文选项卡 JSON 中的 URL。</span><span class="sxs-lookup"><span data-stu-id="38709-196">Just as you typically must change the URLs in the add-in's manifest when you move from development to production (such as changing the domain from localhost to contoso.com), you must also change the URLs in your contextual tabs JSON.</span></span>

    ```json
    {
        "size": 32,
        "sourceLocation": "https://cdn.contoso.com/addins/datainsertion/Images/Group32x32.png"
    },
    {
        "size": 80,
        "sourceLocation": "https://cdn.contoso.com/addins/datainsertion/Images/Group80x80.png"
    }
    ```

1. <span data-ttu-id="38709-197">在我们的简单正在进行的示例中，该组只有一个按钮。</span><span class="sxs-lookup"><span data-stu-id="38709-197">In our simple ongoing example, the group has only a single button.</span></span> <span data-ttu-id="38709-198">将以下对象添加为数组的唯一 `controls` 成员。</span><span class="sxs-lookup"><span data-stu-id="38709-198">Add the following object as the only member of the `controls` array.</span></span> <span data-ttu-id="38709-199">关于此标记，请注意：</span><span class="sxs-lookup"><span data-stu-id="38709-199">About this markup, note:</span></span>

    - <span data-ttu-id="38709-200">除 之外的所有属性 `enabled` 都是必需的。</span><span class="sxs-lookup"><span data-stu-id="38709-200">All the properties, except `enabled`, are required.</span></span>
    - <span data-ttu-id="38709-201">`type` 指定控件的类型。</span><span class="sxs-lookup"><span data-stu-id="38709-201">`type` specifies the type of control.</span></span> <span data-ttu-id="38709-202">值可以是"Button"、"Menu"或"MobileButton"。</span><span class="sxs-lookup"><span data-stu-id="38709-202">The values can be "Button", "Menu", or "MobileButton".</span></span>
    - <span data-ttu-id="38709-203">`id` 可以是最多 125 个字符。</span><span class="sxs-lookup"><span data-stu-id="38709-203">`id` can be up to 125 characters.</span></span> 
    - <span data-ttu-id="38709-204">`actionId` 必须是数组中定义的操作 `actions` ID。</span><span class="sxs-lookup"><span data-stu-id="38709-204">`actionId` must be the ID of an action defined in the `actions` array.</span></span> <span data-ttu-id="38709-205"> (请参阅本节的步骤 1。) </span><span class="sxs-lookup"><span data-stu-id="38709-205">(See step 1 of this section.)</span></span>
    - <span data-ttu-id="38709-206">`label` 是用作按钮标签的用户友好字符串。</span><span class="sxs-lookup"><span data-stu-id="38709-206">`label` is a user-friendly string to serve as the label of the button.</span></span>
    - <span data-ttu-id="38709-207">`superTip` 表示工具提示的丰富形式。</span><span class="sxs-lookup"><span data-stu-id="38709-207">`superTip` represents a rich form of tool tip.</span></span> <span data-ttu-id="38709-208">和 `title` `description` 属性都是必需的。</span><span class="sxs-lookup"><span data-stu-id="38709-208">Both the `title` and `description` properties are required.</span></span>
    - <span data-ttu-id="38709-209">`icon` 指定按钮的图标。</span><span class="sxs-lookup"><span data-stu-id="38709-209">`icon` specifies the icons for the button.</span></span> <span data-ttu-id="38709-210">前面有关组图标的备注也适用于此处。</span><span class="sxs-lookup"><span data-stu-id="38709-210">The previous remarks about the group icon apply here too.</span></span>
    - <span data-ttu-id="38709-211">`enabled` (可选) 指定在上下文选项卡启动时是否启用按钮。</span><span class="sxs-lookup"><span data-stu-id="38709-211">`enabled` (optional) specifies whether the button is enabled when the contextual tab appears starts up.</span></span> <span data-ttu-id="38709-212">如果不存在，则默认为 `true` 。</span><span class="sxs-lookup"><span data-stu-id="38709-212">The default if not present is `true`.</span></span> 

    ```json
    {
        "type": "Button",
        "id": "CtxBt112",
        "actionId": "executeWriteData",
        "enabled": false,
        "label": "Write Data",
        "superTip": {
            "title": "Data Insertion",
            "description": "Use this button to insert data into the document."
        },
        "icon": [
            {
                "size": 32,
                "sourceLocation": "https://cdn.contoso.com/addins/datainsertion/Images/WriteDataButton32x32.png"
            },
            {
                "size": 80,
                "sourceLocation": "https://cdn.contoso.com/addins/datainsertion/Images/WriteDataButton80x80.png"
            }
        ]
    }
    ```
 
<span data-ttu-id="38709-213">下面是 JSON blob 的完整示例。</span><span class="sxs-lookup"><span data-stu-id="38709-213">The following is the complete example of the JSON blob.</span></span>

```json
`{
  "actions": [
    {
      "id": "executeWriteData",
      "type": "ExecuteFunction",
      "functionName": "writeData"
    }
  ],
  "tabs": [
    {
      "id": "CtxTab1",
      "label": "Contoso Data",
      "groups": [
        {
          "id": "CustomGroup111",
          "label": "Insertion",
          "icon": [
            {
                "size": 32,
                "sourceLocation": "https://cdn.contoso.com/addins/datainsertion/Images/Group32x32.png"
            },
            {
                "size": 80,
                "sourceLocation": "https://cdn.contoso.com/addins/datainsertion/Images/Group80x80.png"
            }
          ],
          "controls": [
            {
                "type": "Button",
                "id": "CtxBt112",
                "actionId": "executeWriteData",
                "enabled": false,
                "label": "Write Data",
                "superTip": {
                    "title": "Data Insertion",
                    "description": "Use this button to insert data into the document."
                },
                "icon": [
                    {
                        "size": 32,
                        "sourceLocation": "https://cdn.contoso.com/addins/datainsertion/Images/WriteDataButton32x32.png"
                    },
                    {
                        "size": 80,
                        "sourceLocation": "https://cdn.contoso.com/addins/datainsertion/Images/WriteDataButton80x80.png"
                    }
                ]
            }
          ]
        }
      ]
    }
  ]
}`
```

## <a name="register-the-contextual-tab-with-office-with-requestcreatecontrols"></a><span data-ttu-id="38709-214">使用 requestCreateControls Office上下文选项卡</span><span class="sxs-lookup"><span data-stu-id="38709-214">Register the contextual tab with Office with requestCreateControls</span></span>

<span data-ttu-id="38709-215">上下文选项卡通过调用[Office.ribbon.requestCreateControls](/javascript/api/office/office.ribbon?view=common-js&preserve-view=true#requestCreateControls_tabDefinition_)方法Office注册上下文选项卡。</span><span class="sxs-lookup"><span data-stu-id="38709-215">The contextual tab is registered with Office by calling the [Office.ribbon.requestCreateControls](/javascript/api/office/office.ribbon?view=common-js&preserve-view=true#requestCreateControls_tabDefinition_) method.</span></span> <span data-ttu-id="38709-216">这通常在分配给 方法的函数中或在 方法 `Office.initialize` 中 `Office.onReady` 完成。</span><span class="sxs-lookup"><span data-stu-id="38709-216">This is typically done in either the function that is assigned to `Office.initialize` or with the `Office.onReady` method.</span></span> <span data-ttu-id="38709-217">有关这些方法和初始化外接程序的更多信息，请参阅[Initialize your Office Add-in](../develop/initialize-add-in.md)。</span><span class="sxs-lookup"><span data-stu-id="38709-217">For more about these methods and initializing the add-in, see [Initialize your Office Add-in](../develop/initialize-add-in.md).</span></span> <span data-ttu-id="38709-218">但是，您可以在初始化后随时调用 该方法。</span><span class="sxs-lookup"><span data-stu-id="38709-218">You can, however, call the method anytime after initialization.</span></span>

> [!IMPORTANT]
> <span data-ttu-id="38709-219">`requestCreateControls`在加载项的给定会话中，只能调用方法一次。</span><span class="sxs-lookup"><span data-stu-id="38709-219">The `requestCreateControls` method may be called only once in a given session of an add-in.</span></span> <span data-ttu-id="38709-220">如果再次调用，将引发错误。</span><span class="sxs-lookup"><span data-stu-id="38709-220">An error is thrown if it is called again.</span></span>

<span data-ttu-id="38709-221">示例如下。</span><span class="sxs-lookup"><span data-stu-id="38709-221">The following is an example.</span></span> <span data-ttu-id="38709-222">请注意，必须先使用 方法将 JSON 字符串转换为 JavaScript 对象，然后才能 `JSON.parse` 将其传递给 JavaScript 函数。</span><span class="sxs-lookup"><span data-stu-id="38709-222">Note that the JSON string must be converted to a JavaScript object with the `JSON.parse` method before it can be passed to a JavaScript function.</span></span>

```javascript
Office.onReady(async () => {
    const contextualTabJSON = ` ... `; // Assign the JSON string such as the one at the end of the preceding section.
    const contextualTab = JSON.parse(contextualTabJSON);
    await Office.ribbon.requestCreateControls(contextualTab);
});
```

## <a name="specify-the-contexts-when-the-tab-will-be-visible-with-requestupdate"></a><span data-ttu-id="38709-223">使用 requestUpdate 指定选项卡可见时上下文</span><span class="sxs-lookup"><span data-stu-id="38709-223">Specify the contexts when the tab will be visible with requestUpdate</span></span>

<span data-ttu-id="38709-224">通常，当用户启动的事件更改外接程序上下文时，应显示自定义上下文选项卡。</span><span class="sxs-lookup"><span data-stu-id="38709-224">Typically, a custom contextual tab should appear when a user-initiated event changes the add-in context.</span></span> <span data-ttu-id="38709-225">请考虑在激活工作簿的默认工作表上的图表时（并且仅在 (工作簿的默认工作表上显示Excel选项卡) 的情况。</span><span class="sxs-lookup"><span data-stu-id="38709-225">Consider a scenario in which the tab should be visible when, and only when, a chart (on the default worksheet of an Excel workbook) is activated.</span></span>

<span data-ttu-id="38709-226">首先分配处理程序。</span><span class="sxs-lookup"><span data-stu-id="38709-226">Begin by assigning handlers.</span></span> <span data-ttu-id="38709-227">此操作通常在 方法中完成，如以下示例所示，该示例将 (在稍后步骤) 中创建的处理程序分配给工作表中所有图表的 和 `Office.onReady` `onActivated` `onDeactivated` 事件。</span><span class="sxs-lookup"><span data-stu-id="38709-227">This is commonly done in the `Office.onReady` method as in the following example which assigns handlers (created in a later step) to the `onActivated` and `onDeactivated` events of all the charts in the worksheet.</span></span>

```javascript
Office.onReady(async () => {
    const contextualTabJSON = ` ... `; // Assign the JSON string.
    const contextualTab = JSON.parse(contextualTabJSON);
    await Office.ribbon.requestCreateControls(contextualTab);

    await Excel.run(context => {
        var charts = context.workbook.worksheets
            .getActiveWorksheet()
            .charts;
        charts.onActivated.add(showDataTab);
        charts.onDeactivated.add(hideDataTab);
        return context.sync();
    });
});
```

<span data-ttu-id="38709-228">接下来，定义处理程序。</span><span class="sxs-lookup"><span data-stu-id="38709-228">Next, define the handlers.</span></span> <span data-ttu-id="38709-229">下面是 的一个简单示例，但请参阅本文稍后介绍的处理 `showDataTab` [HostRestartNeeded](#handle-the-hostrestartneeded-error) 错误，了解函数的更可靠版本。</span><span class="sxs-lookup"><span data-stu-id="38709-229">The following is a simple example of a `showDataTab`, but see [Handling the HostRestartNeeded error](#handle-the-hostrestartneeded-error) later in this article for a more robust version of the function.</span></span> <span data-ttu-id="38709-230">关于此代码，请注意以下几点：</span><span class="sxs-lookup"><span data-stu-id="38709-230">About this code, note:</span></span>

- <span data-ttu-id="38709-231">Office 控制何时更新功能区的状态。</span><span class="sxs-lookup"><span data-stu-id="38709-231">Office controls when it updates the state of the ribbon.</span></span> <span data-ttu-id="38709-232">the [Office.ribbon.requestUpdate](/javascript/api/office/office.ribbon?view=common-js&preserve-view=true#requestupdate-input-) method queues a request to update.</span><span class="sxs-lookup"><span data-stu-id="38709-232">The  [Office.ribbon.requestUpdate](/javascript/api/office/office.ribbon?view=common-js&preserve-view=true#requestupdate-input-) method queues a request to update.</span></span> <span data-ttu-id="38709-233">方法将在请求排入队列后（而不是功能区实际更新时） `Promise` 解析对象。</span><span class="sxs-lookup"><span data-stu-id="38709-233">The method will resolve the `Promise` object as soon as it has queued the request, not when the ribbon actually updates.</span></span>
- <span data-ttu-id="38709-234">该方法的参数是 `requestUpdate` [RibbonUpdaterData](/javascript/api/office/office.ribbonupdaterdata) 对象， (1) 按照 *JSON* 中指定的 ID 指定选项卡， (2) 指定选项卡的可见性。</span><span class="sxs-lookup"><span data-stu-id="38709-234">The parameter for the `requestUpdate` method is a [RibbonUpdaterData](/javascript/api/office/office.ribbonupdaterdata) object that (1) specifies the tab by its ID *exactly as specified in the JSON* and (2) specifies visibility of the tab.</span></span>
- <span data-ttu-id="38709-235">如果你有多个自定义上下文选项卡应在同一上下文中可见，则只需向数组中添加其他选项卡 `tabs` 对象。</span><span class="sxs-lookup"><span data-stu-id="38709-235">If you have more than one custom contextual tab that should be visible in the same context, you simply add additional tab objects to the `tabs` array.</span></span>

```javascript
async function showDataTab() {
    await Office.ribbon.requestUpdate({
        tabs: [
            {
                id: "CtxTab1",
                visible: true
            }
        ]});
}
```

<span data-ttu-id="38709-236">隐藏选项卡的处理程序几乎完全相同，只是它将 `visible` 属性设置回 `false` 。</span><span class="sxs-lookup"><span data-stu-id="38709-236">The handler to hide the tab is nearly identical, except that it sets the `visible` property back to `false`.</span></span>

<span data-ttu-id="38709-237">JavaScript Office还提供多个 (接口) ，以便更轻松地构造 `RibbonUpdateData` 对象。</span><span class="sxs-lookup"><span data-stu-id="38709-237">The Office JavaScript library also provides several interfaces (types) to make it easier to construct the`RibbonUpdateData` object.</span></span> <span data-ttu-id="38709-238">以下是 `showDataTab` TypeScript 中的 函数，它使用这些类型。</span><span class="sxs-lookup"><span data-stu-id="38709-238">The following is the `showDataTab` function in TypeScript and it makes use of these types.</span></span>

```typescript
const showDataTab = async () => {
    const myContextualTab: Office.Tab = {id: "CtxTab1", visible: true};
    const ribbonUpdater: Office.RibbonUpdaterData = { tabs: [ myContextualTab ]};
    await Office.ribbon.requestUpdate(ribbonUpdater);
}
```

### <a name="toggle-tab-visibility-and-the-enabled-status-of-a-button-at-the-same-time"></a><span data-ttu-id="38709-239">切换选项卡可见性和按钮的启用状态</span><span class="sxs-lookup"><span data-stu-id="38709-239">Toggle tab visibility and the enabled status of a button at the same time</span></span>

<span data-ttu-id="38709-240">方法还用于切换自定义上下文选项卡或自定义核心选项卡上自定义按钮的启用或 `requestUpdate` 禁用状态。有关详细信息，请参阅启用 [和禁用加载项命令](disable-add-in-commands.md)。</span><span class="sxs-lookup"><span data-stu-id="38709-240">The `requestUpdate` method is also used to toggle the enabled or disabled status of a custom button on either a custom contextual tab or a custom core tab. For details about this, see [Enable and Disable Add-in Commands](disable-add-in-commands.md).</span></span> <span data-ttu-id="38709-241">在某些情况下，你可能希望同时更改选项卡的可见性和按钮的启用状态。</span><span class="sxs-lookup"><span data-stu-id="38709-241">There may be scenarios in which you want to change both the visibility of a tab and the enabled status of a button at the same time.</span></span> <span data-ttu-id="38709-242">为此，请调用 `requestUpdate` 。</span><span class="sxs-lookup"><span data-stu-id="38709-242">You do this with a single call of `requestUpdate`.</span></span> <span data-ttu-id="38709-243">下面是一个示例，其中核心选项卡上的按钮在显示上下文选项卡的同时处于启用状态。</span><span class="sxs-lookup"><span data-stu-id="38709-243">The following is an example in which a button on a core tab is enabled at the same time as a contextual tab is made visible.</span></span>

```javascript
function myContextChanges() {
    Office.ribbon.requestUpdate({
        tabs: [
            {
                id: "CtxTab1",
                visible: true
            },
            {
                id: "OfficeAppTab1",
                groups: [
                    {
                        id: "CustomGroup111",
                        controls: [
                            {
                                id: "MyButton",
                                enabled: true
                            }
                        ]
                    }
                ]
            ]}
        ]
    });
}
```

<span data-ttu-id="38709-244">在下面的示例中，启用的按钮位于要显示的上下文选项卡上。</span><span class="sxs-lookup"><span data-stu-id="38709-244">In the following example, the button that is enabled is on the very same contextual tab that is being made visible.</span></span>

```javascript
function myContextChanges() {
    Office.ribbon.requestUpdate({
        tabs: [
            {
                id: "CtxTab1",
                visible: true,
                groups: [
                    {
                        id: "CustomGroup111",
                        controls: [
                            {
                                id: "MyButton",
                                enabled: true
                           }
                       ]
                   }
               ]
            }
        ]
    });
}
```

## <a name="open-a-task-pane-from-contextual-tabs"></a><span data-ttu-id="38709-245">从上下文选项卡打开任务窗格</span><span class="sxs-lookup"><span data-stu-id="38709-245">Open a task pane from contextual tabs</span></span>

<span data-ttu-id="38709-246">若要从自定义上下文选项卡上的按钮打开任务窗格，在 JSON 中创建具有 `type` 的 操作 `ShowTaskpane` 。</span><span class="sxs-lookup"><span data-stu-id="38709-246">To open your task pane from a button on a custom contextual tab, create an action in the JSON with a `type` of `ShowTaskpane`.</span></span> <span data-ttu-id="38709-247">然后，定义一个 `actionId` 按钮，将 属性设置为 `id` 操作 。</span><span class="sxs-lookup"><span data-stu-id="38709-247">Then define a button with the `actionId` property set to the `id` of the action.</span></span> <span data-ttu-id="38709-248">这将打开清单中的 元素 `<Runtime>` 指定的默认任务窗格。</span><span class="sxs-lookup"><span data-stu-id="38709-248">This opens the default task pane specified by the `<Runtime>` element in your manifest.</span></span>

```json
`{
  "actions": [
    {
      "id": "openChartsTaskpane",
      "type": "ShowTaskpane",
      "title": "Work with Charts",
      "supportPinning": false
    }
  ],
  "tabs": [
    {
      // some tab properties omitted
      "groups": [
        {
          // some group properties omitted
          "controls": [
            {
                "type": "Button",
                "id": "CtxBt112",
                "actionId": "openChartsTaskpane",
                "enabled": false,
                "label": "Open Charts Taskpane",
                // some control properties omitted
            }
          ]
        }
      ]
    }
  ]
}`
```

<span data-ttu-id="38709-249">若要打开不是默认任务窗格的任何任务窗格，请指定操作 `sourceLocation` 定义中的属性。</span><span class="sxs-lookup"><span data-stu-id="38709-249">To open any task pane that is not the default task pane, specify a `sourceLocation` property in the definition of the action.</span></span> <span data-ttu-id="38709-250">在下面的示例中，从另一个按钮打开第二个任务窗格。</span><span class="sxs-lookup"><span data-stu-id="38709-250">In the following example, a second task pane is opened from a different button.</span></span>

> [!IMPORTANT]
>
> - <span data-ttu-id="38709-251">如果 `sourceLocation` 为 操作指定了 ，则任务窗格 *不使用* 共享运行时。</span><span class="sxs-lookup"><span data-stu-id="38709-251">When a `sourceLocation` is specified for the action, then the task pane does *not* use the shared runtime.</span></span> <span data-ttu-id="38709-252">它在新的 JavaScript 运行时中运行。</span><span class="sxs-lookup"><span data-stu-id="38709-252">It runs in a new JavaScript runtime.</span></span>
> - <span data-ttu-id="38709-253">共享运行时不能超过一个任务窗格，因此类型操作不能超过一个，因此 `ShowTaskpane` 不能忽略 `sourceLocation` 属性。</span><span class="sxs-lookup"><span data-stu-id="38709-253">No more than one task pane can use the shared runtime, so no more than one action of type `ShowTaskpane` can omit the `sourceLocation` property.</span></span>

```json
`{
  "actions": [
    {
      "id": "openChartsTaskpane",
      "type": "ShowTaskpane",
      "title": "Work with Charts",
      "supportPinning": false
    },
    {
      "id": "openTablesTaskpane",
      "type": "ShowTaskpane",
      "title": "Work with Tables",
      "supportPinning": false
      "sourceLocation": "https://MyDomain.com/myPage.html"
    }
  ],
  "tabs": [
    {
      // some tab properties omitted
      "groups": [
        {
          // some group properties omitted
          "controls": [
            {
                "type": "Button",
                "id": "CtxBt112",
                "actionId": "openChartsTaskpane",
                "enabled": false,
                "label": "Open Charts Taskpane",
                // some control properties omitted
            },
            {
                "type": "Button",
                "id": "CtxBt113",
                "actionId": "openTablesTaskpane",
                "enabled": false,
                "label": "Open Tables Taskpane",
                // some control properties omitted
            }
          ]
        }
      ]
    }
  ]
}`
```

## <a name="localize-the-json-text"></a><span data-ttu-id="38709-254">本地化 JSON 文本</span><span class="sxs-lookup"><span data-stu-id="38709-254">Localize the JSON text</span></span>

<span data-ttu-id="38709-255">传递给 的 JSON blob 的本地化方式与自定义核心选项卡的清单标记本地化 (如从清单控件本地化) `requestCreateControls` 中所述。 [](../develop/localization.md#control-localization-from-the-manifest)</span><span class="sxs-lookup"><span data-stu-id="38709-255">The JSON blob that is passed to `requestCreateControls` is not localized the same way that the manifest markup for custom core tabs is localized (which is described at [Control localization from the manifest](../develop/localization.md#control-localization-from-the-manifest)).</span></span> <span data-ttu-id="38709-256">相反，本地化必须在运行时针对每个区域设置使用不同的 JSON blob。</span><span class="sxs-lookup"><span data-stu-id="38709-256">Instead, the localization must occur at runtime using distinct JSON blobs for each locale.</span></span> <span data-ttu-id="38709-257">建议您使用一个 `switch` 语句来测试[Office.context.displayLanguage](/javascript/api/office/office.context#displayLanguage)属性。</span><span class="sxs-lookup"><span data-stu-id="38709-257">We suggest that you use a `switch` statement that tests the [Office.context.displayLanguage](/javascript/api/office/office.context#displayLanguage) property.</span></span> <span data-ttu-id="38709-258">示例如下。</span><span class="sxs-lookup"><span data-stu-id="38709-258">The following is an example.</span></span>

```javascript
function GetContextualTabsJsonSupportedLocale () {
    var displayLanguage = Office.context.displayLanguage;

        switch (displayLanguage) {
            case 'en-US':
                return `{
                    "actions": [
                        // actions omitted
                     ],
                    "tabs": [
                        {
                          "id": "CtxTab1",
                          "label": "Contoso Data",
                          "groups": [
                              // groups omitted
                          ]
                        }
                    ]
                }`;

            case 'fr-FR':
                return `{
                    "actions": [
                        // actions omitted 
                    ],
                    "tabs": [
                        {
                          "id": "CtxTab1",
                          "label": "Contoso Données",
                          "groups": [
                              // groups omitted
                          ]
                       }
                    ]
               }`;

            // Other cases omitted
       }
}
```

<span data-ttu-id="38709-259">然后，代码调用 函数，获取传递给 的本地化 `requestCreateControls` blob，如以下示例所示。</span><span class="sxs-lookup"><span data-stu-id="38709-259">Then your code calls the function to get the localized blob that is passed to `requestCreateControls`, as in the following example.</span></span>

```javascript
var contextualTabJSON = GetContextualTabsJsonSupportedLocale();
```

## <a name="best-practices-for-custom-contextual-tabs"></a><span data-ttu-id="38709-260">自定义上下文选项卡的最佳实践</span><span class="sxs-lookup"><span data-stu-id="38709-260">Best practices for custom contextual tabs</span></span>

### <a name="implement-an-alternate-ui-experience-when-custom-contextual-tabs-are-not-supported"></a><span data-ttu-id="38709-261">在不支持自定义上下文选项卡时实现备用 UI 体验</span><span class="sxs-lookup"><span data-stu-id="38709-261">Implement an alternate UI experience when custom contextual tabs are not supported</span></span>

<span data-ttu-id="38709-262">平台、应用程序Office和内部Office组合不支持 `requestCreateControls` 。</span><span class="sxs-lookup"><span data-stu-id="38709-262">Some combinations of platform, Office application, and Office build don't support `requestCreateControls`.</span></span> <span data-ttu-id="38709-263">您的外接程序应设计为为在这些组合之一上运行外接程序的用户提供备用体验。</span><span class="sxs-lookup"><span data-stu-id="38709-263">Your add-in should be designed to provide an alternate experience to users who are running the add-in on one of those combinations.</span></span> <span data-ttu-id="38709-264">以下各节介绍提供回退体验的两种方法。</span><span class="sxs-lookup"><span data-stu-id="38709-264">The following sections describe two ways of providing a fallback experience.</span></span>

#### <a name="use-noncontextual-tabs-or-controls"></a><span data-ttu-id="38709-265">使用非上下文选项卡或控件</span><span class="sxs-lookup"><span data-stu-id="38709-265">Use noncontextual tabs or controls</span></span>

<span data-ttu-id="38709-266">有一个清单元素 [OverriddenByRibbonApi](../reference/manifest/overriddenbyribbonapi.md)，旨在创建外接程序中的回退体验，该体验在外接程序在不支持自定义上下文选项卡的应用程序或平台上运行时实现自定义上下文选项卡。</span><span class="sxs-lookup"><span data-stu-id="38709-266">There is a manifest element, [OverriddenByRibbonApi](../reference/manifest/overriddenbyribbonapi.md), that is designed to create a fallback experience in an add-in that implements custom contextual tabs when the add-in is running on an application or platform that doesn't support custom contextual tabs.</span></span> 

<span data-ttu-id="38709-267">使用此元素的最简单策略是，在清单中定义一个或多个自定义核心选项卡 (即，与外接程序中的自定义上下文选项卡的功能区自定义项重复的非上下文自定义选项卡) 。</span><span class="sxs-lookup"><span data-stu-id="38709-267">The simplest strategy for using this element is that you define in the manifest one or more custom core tabs (that is, *noncontextual* custom tabs) that duplicate the ribbon customizations of the custom contextual tabs in your add-in.</span></span> <span data-ttu-id="38709-268">但添加 `<OverriddenByRibbonApi>true</OverriddenByRibbonApi>` 为 [CustomTab](../reference/manifest/customtab.md)的第一个子元素。</span><span class="sxs-lookup"><span data-stu-id="38709-268">But you add `<OverriddenByRibbonApi>true</OverriddenByRibbonApi>` as the first child element of the [CustomTab](../reference/manifest/customtab.md).</span></span> <span data-ttu-id="38709-269">这样做的效果如下：</span><span class="sxs-lookup"><span data-stu-id="38709-269">The effect of doing so is the following:</span></span>

- <span data-ttu-id="38709-270">如果外接程序在支持自定义上下文选项卡的应用程序和平台上运行，则自定义核心选项卡将不会显示在功能区上。</span><span class="sxs-lookup"><span data-stu-id="38709-270">If the add-in runs on an application and platform that support custom contextual tabs, then the custom core tab won't appear on the ribbon.</span></span> <span data-ttu-id="38709-271">相反，自定义上下文选项卡将在外接程序调用 方法时 `requestCreateControls` 创建。</span><span class="sxs-lookup"><span data-stu-id="38709-271">Instead, the custom contextual tab will be created when the add-in calls the `requestCreateControls` method.</span></span>
- <span data-ttu-id="38709-272">如果加载项在不支持的应用程序或平台上运行，则自定义核心 `requestCreateControls` 选项卡会显示在功能区上。</span><span class="sxs-lookup"><span data-stu-id="38709-272">If the add-in runs on an application or platform that *doesn't* support `requestCreateControls`, then the custom core tab does appear on the ribbon.</span></span>

<span data-ttu-id="38709-273">下面是此简单策略的一个示例。</span><span class="sxs-lookup"><span data-stu-id="38709-273">The following is an example of this simple strategy.</span></span>

```xml
<OfficeApp ...>
  ...
  <VersionOverrides ...>
    ...
    <Hosts>
      <Host ...>
        ...
        <DesktopFormFactor>
          <ExtensionPoint ...>
            <CustomTab ...>
              <OverriddenByRibbonApi>true</OverriddenByRibbonApi>
              ...
              <Group ...>
                ...
                <Control ... id="MyButton">
                  ...
                  <Action ...>
...
</OfficeApp>
```

<span data-ttu-id="38709-274">此简单策略使用自定义核心选项卡，该选项卡将自定义上下文选项卡与它的子组和控件镜像，但您可以使用更复杂的策略。</span><span class="sxs-lookup"><span data-stu-id="38709-274">This simple strategy uses a custom core tab that mirrors a custom contextual tab with it's child groups and controls, but you can use a more complex strategy.</span></span> <span data-ttu-id="38709-275">还可以将元素作为第一 (个) 子元素添加到 Group 和 Control 元素中 (按钮类型和菜单类型) 菜单 `<OverriddenByRibbonApi>` [](../reference/manifest/group.md)[](../reference/manifest/control.md)[](../reference/manifest/control.md#button-control)[](../reference/manifest/control.md#menu-dropdown-button-controls) `<Item>` 元素。</span><span class="sxs-lookup"><span data-stu-id="38709-275">The `<OverriddenByRibbonApi>` element can also be added as (the first) child element to the [Group](../reference/manifest/group.md) and [Control](../reference/manifest/control.md) elements (both [button type](../reference/manifest/control.md#button-control) and [menu type](../reference/manifest/control.md#menu-dropdown-button-controls)), and menu `<Item>` elements.</span></span> <span data-ttu-id="38709-276">这一事实使你能够将原本显示在上下文选项卡上的组和控件分布在各种自定义核心选项卡的各种组、按钮和菜单中。</span><span class="sxs-lookup"><span data-stu-id="38709-276">This fact enables you to distribute the groups and controls that would otherwise appear on the contextual tab among various groups, buttons, and menus in various custom core tabs.</span></span> <span data-ttu-id="38709-277">示例如下。</span><span class="sxs-lookup"><span data-stu-id="38709-277">The following is an example.</span></span> <span data-ttu-id="38709-278">请注意，仅在不支持自定义上下文选项卡时，"MyButton"将显示在自定义核心选项卡上。</span><span class="sxs-lookup"><span data-stu-id="38709-278">Note that "MyButton" will appear on the custom core tab only when custom contextual tabs are not supported.</span></span> <span data-ttu-id="38709-279">但是，无论自定义上下文选项卡是否受支持，都会显示父组和自定义核心选项卡。</span><span class="sxs-lookup"><span data-stu-id="38709-279">But the parent group and custom core tab will appear regardless of whether custom contextual tabs are supported.</span></span>

```xml
<OfficeApp ...>
  ...
  <VersionOverrides ...>
    ...
    <Hosts>
      <Host ...>
        ...
        <DesktopFormFactor>
          <ExtensionPoint ...>
            <CustomTab ...>              
              ...
              <Group ...>
                ...
                <Control ... id="MyButton">
                  <OverriddenByRibbonApi>true</OverriddenByRibbonApi>
                  ...
                  <Action ...>
...
</OfficeApp>
```

<span data-ttu-id="38709-280">有关更多示例，请参阅 [OverriddenByRibbonApi](../reference/manifest/overriddenbyribbonapi.md)。</span><span class="sxs-lookup"><span data-stu-id="38709-280">For more examples, see [OverriddenByRibbonApi](../reference/manifest/overriddenbyribbonapi.md).</span></span>

<span data-ttu-id="38709-281">当父选项卡、组或菜单标记为 时，它将不可见，并且当不支持自定义上下文选项卡时，将忽略其所有子 `<OverriddenByRibbonApi>true</OverriddenByRibbonApi>` 标记。</span><span class="sxs-lookup"><span data-stu-id="38709-281">When a parent tab, group, or menu is marked with `<OverriddenByRibbonApi>true</OverriddenByRibbonApi>`, then it isn't visible, and all of it's child markup is ignored, when custom contextual tabs aren't supported.</span></span> <span data-ttu-id="38709-282">因此，这些子元素中的任一元素是否具有 元素或其 `<OverriddenByRibbonApi>` 值是什么并不重要。</span><span class="sxs-lookup"><span data-stu-id="38709-282">So, it doesn't matter if any of those child elements have the `<OverriddenByRibbonApi>` element or what its value is.</span></span> <span data-ttu-id="38709-283">其含义是，如果某个菜单项、控件或组必须在所有上下文中可见，则不仅不应使用 标记它，而且其上级菜单、组和选项卡也必须不按此方式 `<OverriddenByRibbonApi>true</OverriddenByRibbonApi>` *进行标记*。</span><span class="sxs-lookup"><span data-stu-id="38709-283">The implication of this is that if a menu item, control, or group must be visible in all contexts, then not only should it not be marked with `<OverriddenByRibbonApi>true</OverriddenByRibbonApi>`, but *its ancestor menu, group, and tab must also not be marked this way*.</span></span>

> [!IMPORTANT]
> <span data-ttu-id="38709-284">请勿使用 *标记选项卡* 、组或菜单的所有子元素 `<OverriddenByRibbonApi>true</OverriddenByRibbonApi>` 。</span><span class="sxs-lookup"><span data-stu-id="38709-284">Don't mark *all* of the child elements of a tab, group, or menu with `<OverriddenByRibbonApi>true</OverriddenByRibbonApi>`.</span></span> <span data-ttu-id="38709-285">如果由于上一段给出的原因而标记父元素， `<OverriddenByRibbonApi>true</OverriddenByRibbonApi>` 则这一点没有意义。</span><span class="sxs-lookup"><span data-stu-id="38709-285">This is pointless if the parent element is marked with `<OverriddenByRibbonApi>true</OverriddenByRibbonApi>` for reasons given in the preceding paragraph.</span></span> <span data-ttu-id="38709-286">此外，如果在父选项卡上将 (或设置为) ，则无论自定义上下文选项卡是否受支持，父选项卡都会显示，但在支持自定义上下文选项卡时将为空 `<OverriddenByRibbonApi>` `false` 。</span><span class="sxs-lookup"><span data-stu-id="38709-286">Moreover, if you leave out the `<OverriddenByRibbonApi>` on the parent (or set it to `false`), then the parent will appear regardless of whether custom contextual tabs are supported, but it will be empty when they are supported.</span></span> <span data-ttu-id="38709-287">因此，如果支持自定义上下文选项卡时不应显示所有子元素，则使用 标记父元素，仅标记父元素 `<OverriddenByRibbonApi>true</OverriddenByRibbonApi>` 。</span><span class="sxs-lookup"><span data-stu-id="38709-287">So, if all the child elements shouldn't appear when custom contextual tabs are supported, mark the parent, and only the parent, with `<OverriddenByRibbonApi>true</OverriddenByRibbonApi>`.</span></span>

#### <a name="use-apis-that-show-or-hide-a-task-pane-in-specified-contexts"></a><span data-ttu-id="38709-288">使用在指定的上下文中显示或隐藏任务窗格的 API</span><span class="sxs-lookup"><span data-stu-id="38709-288">Use APIs that show or hide a task pane in specified contexts</span></span>

<span data-ttu-id="38709-289">作为替代方法，加载项可以使用与自定义上下文选项卡上控件功能重复的 UI 控件定义 `<OverriddenByRibbonApi>` 任务窗格。然后，使用[Office.addin.showAsTaskpane](/javascript/api/office/office.addin?view=common-js&preserve-view=true#showAsTaskpane__)和[Office.addin.hide](/javascript/api/office/office.addin?view=common-js&preserve-view=true#hide__)方法在支持上下文选项卡时（且仅在何时）显示任务窗格。</span><span class="sxs-lookup"><span data-stu-id="38709-289">As an alternative to `<OverriddenByRibbonApi>`, your add-in can define a task pane with UI controls that duplicate the functionality of the controls on a custom contextual tab. Then use the [Office.addin.showAsTaskpane](/javascript/api/office/office.addin?view=common-js&preserve-view=true#showAsTaskpane__) and [Office.addin.hide](/javascript/api/office/office.addin?view=common-js&preserve-view=true#hide__) methods to show the task pane when, and only when, the contextual tab would have been shown if it was supported.</span></span> <span data-ttu-id="38709-290">有关如何使用这些方法的详细信息，请参阅显示或隐藏加载项Office[窗格](../develop/show-hide-add-in.md)。</span><span class="sxs-lookup"><span data-stu-id="38709-290">For details on how to use these methods, see [Show or hide the task pane of your Office Add-in](../develop/show-hide-add-in.md).</span></span>

### <a name="handle-the-hostrestartneeded-error"></a><span data-ttu-id="38709-291">处理 HostRestartNeeded 错误</span><span class="sxs-lookup"><span data-stu-id="38709-291">Handle the HostRestartNeeded error</span></span>

<span data-ttu-id="38709-292">在某些情况下，Office 无法更新功能区，并将返回错误。</span><span class="sxs-lookup"><span data-stu-id="38709-292">In some scenarios, Office is unable to update the ribbon and will return an error.</span></span> <span data-ttu-id="38709-293">例如，如果升级了加载项，并且升级后的加载项具有一组不同的自定义加载项命令，则必须关闭并重新打开 Office 应用程序。</span><span class="sxs-lookup"><span data-stu-id="38709-293">For example, if the add-in is upgraded and the upgraded add-in has a different set of custom add-in commands, then the Office application must be closed and reopened.</span></span> <span data-ttu-id="38709-294">在此之前，`requestUpdate` 方法将返回错误 `HostRestartNeeded`。</span><span class="sxs-lookup"><span data-stu-id="38709-294">Until it is, the `requestUpdate` method will return the error `HostRestartNeeded`.</span></span> <span data-ttu-id="38709-295">代码应处理此错误。</span><span class="sxs-lookup"><span data-stu-id="38709-295">Your code should handle this error.</span></span> <span data-ttu-id="38709-296">下面是操作方法的示例。</span><span class="sxs-lookup"><span data-stu-id="38709-296">The following is an example of how.</span></span> <span data-ttu-id="38709-297">在此示例中，`reportError` 方法向用户显示错误。</span><span class="sxs-lookup"><span data-stu-id="38709-297">In this case, the `reportError` method displays the error to the user.</span></span>

```javascript
function showDataTab() {
    try {
        Office.ribbon.requestUpdate({
            tabs: [
                {
                    id: "CtxTab1",
                    visible: true
                }
            ]});
    }
    catch(error) {
        if (error.code == "HostRestartNeeded"){
            reportError("Contoso Awesome Add-in has been upgraded. Please save your work, then close and reopen the Office application.");
        }
    }
}
```
