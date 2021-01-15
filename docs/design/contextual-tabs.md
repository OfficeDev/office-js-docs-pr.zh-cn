---
title: 在 Office 加载项中创建自定义上下文选项卡
description: 了解如何将自定义上下文选项卡添加到 Office 外接程序。
ms.date: 01/11/2021
localization_priority: Normal
ms.openlocfilehash: 12286ef675a938e4abd8dd3caa90cd97586cb6d7
ms.sourcegitcommit: 6a378d2a3679757c5014808ae9da8ababbfe8b16
ms.translationtype: MT
ms.contentlocale: zh-CN
ms.lasthandoff: 01/15/2021
ms.locfileid: "49870635"
---
# <a name="create-custom-contextual-tabs-in-office-add-ins-preview"></a><span data-ttu-id="047f4-103">在 Office 加载项中创建自定义上下文选项卡（预览）</span><span class="sxs-lookup"><span data-stu-id="047f4-103">Create custom contextual tabs in Office Add-ins (preview)</span></span>

<span data-ttu-id="047f4-104">上下文选项卡是 Office 功能区中的隐藏选项卡控件，在 Office 文档中发生指定事件时显示在选项卡行中。</span><span class="sxs-lookup"><span data-stu-id="047f4-104">A contextual tab is a hidden tab control in the Office ribbon that is displayed in the tab row when a specified event occurs in the Office document.</span></span> <span data-ttu-id="047f4-105">例如 **，选择表** 时显示在 Excel 功能区上的"表设计"选项卡。</span><span class="sxs-lookup"><span data-stu-id="047f4-105">For example, the **Table Design** tab that appears on the Excel ribbon when a table is selected.</span></span> <span data-ttu-id="047f4-106">您可以通过创建更改可见性的事件处理程序，在 Office 外接程序中包括自定义上下文选项卡并指定它们何时可见或隐藏。</span><span class="sxs-lookup"><span data-stu-id="047f4-106">You can include custom contextual tabs in your Office add-in and specify when they are visible or hidden, by creating event handlers that change the visibility.</span></span> <span data-ttu-id="047f4-107"> (，自定义上下文选项卡不会响应焦点更改。) </span><span class="sxs-lookup"><span data-stu-id="047f4-107">(However, custom contextual tabs do not respond to focus changes.)</span></span>

> [!NOTE]
> <span data-ttu-id="047f4-108">本文假定你熟悉以下文档。</span><span class="sxs-lookup"><span data-stu-id="047f4-108">This article assumes that you are familiar with the following documentation.</span></span> <span data-ttu-id="047f4-109">如果你最近未使用加载项命令（自定义菜单项和功能区按钮），请查看该文档。</span><span class="sxs-lookup"><span data-stu-id="047f4-109">Please review it if you haven't worked with Add-in Commands (custom menu items and ribbon buttons) recently.</span></span>
>
> - [<span data-ttu-id="047f4-110">加载项命令的基本概念</span><span class="sxs-lookup"><span data-stu-id="047f4-110">Basic concepts for Add-in Commands</span></span>](add-in-commands.md)

> [!IMPORTANT]
> <span data-ttu-id="047f4-111">自定义上下文选项卡为预览。</span><span class="sxs-lookup"><span data-stu-id="047f4-111">Custom contextual tabs are in preview.</span></span> <span data-ttu-id="047f4-112">请在开发或测试环境中试验它们，但不要将其添加到生产外接程序。</span><span class="sxs-lookup"><span data-stu-id="047f4-112">Please experiment with them in a development or testing environment but don't add them to a production add-in.</span></span>
>
> <span data-ttu-id="047f4-113">自定义上下文选项卡当前仅在 Excel 上受支持，并且仅在以下平台和内部版本上受支持：</span><span class="sxs-lookup"><span data-stu-id="047f4-113">Custom contextual tabs are currently only supported on Excel and only on these platforms and builds:</span></span>
>
> - <span data-ttu-id="047f4-114">仅适用于 Windows (Microsoft 365 上的 Excel，而不是永久许可证) ：版本 2011 (内部版本 13426.20274) 。</span><span class="sxs-lookup"><span data-stu-id="047f4-114">Excel on Windows (Microsoft 365 only, not perpetual license): Version 2011 (Build 13426.20274).</span></span> <span data-ttu-id="047f4-115">你的 Microsoft 365 订阅可能需要位于当前频道 [ (预览版) ](https://insider.office.com/join/windows) 以前称为"每月频道 (定向) "或"预览体验成员慢"。</span><span class="sxs-lookup"><span data-stu-id="047f4-115">Your Microsoft 365 subscription may need to be on the [Current Channel (Preview)](https://insider.office.com/join/windows) formerly called "Monthly Channel (Targeted)" or "Insider Slow".</span></span>

> [!NOTE]
> <span data-ttu-id="047f4-116">自定义上下文选项卡仅适用于支持以下要求集的平台。</span><span class="sxs-lookup"><span data-stu-id="047f4-116">Custom contextual tabs work only on platforms that support the following requirement sets.</span></span> <span data-ttu-id="047f4-117">有关要求集以及如何使用它们，请参阅"指定 Office 应用程序和[API 要求"。](../develop/specify-office-hosts-and-api-requirements.md)</span><span class="sxs-lookup"><span data-stu-id="047f4-117">For more about requirement sets and how to work with them, see [Specify Office applications and API requirements](../develop/specify-office-hosts-and-api-requirements.md).</span></span>
>
> - [<span data-ttu-id="047f4-118">SharedRuntime 1.1</span><span class="sxs-lookup"><span data-stu-id="047f4-118">SharedRuntime 1.1</span></span>](../reference/requirement-sets/shared-runtime-requirement-sets.md)

## <a name="behavior-of-custom-contextual-tabs"></a><span data-ttu-id="047f4-119">自定义上下文选项卡的行为</span><span class="sxs-lookup"><span data-stu-id="047f4-119">Behavior of custom contextual tabs</span></span>

<span data-ttu-id="047f4-120">自定义上下文选项卡的用户体验遵循内置 Office 上下文选项卡的模式。</span><span class="sxs-lookup"><span data-stu-id="047f4-120">The user experience for custom contextual tabs follows the pattern of built-in Office contextual tabs.</span></span> <span data-ttu-id="047f4-121">以下是放置自定义上下文选项卡的基础知识：</span><span class="sxs-lookup"><span data-stu-id="047f4-121">The following are the basic principles for the placement custom contextual tabs:</span></span>

- <span data-ttu-id="047f4-122">当自定义上下文选项卡可见时，它将显示在功能区的右端。</span><span class="sxs-lookup"><span data-stu-id="047f4-122">When a custom contextual tab is visible, it appears on the right end of the ribbon.</span></span>
- <span data-ttu-id="047f4-123">如果加载项中的一个或多个内置上下文选项卡和一个或多个自定义上下文选项卡同时可见，则自定义上下文选项卡始终位于所有内置上下文选项卡的右侧。</span><span class="sxs-lookup"><span data-stu-id="047f4-123">If one or more built-in contextual tabs and one or more custom contextual tabs from add-ins are visible at the same time, the custom contextual tabs are always to the right of all of the built-in contextual tabs.</span></span>
- <span data-ttu-id="047f4-124">如果您的外接程序具有多个上下文选项卡，并且存在多个可见上下文，则它们按照在外接程序中定义的顺序显示。</span><span class="sxs-lookup"><span data-stu-id="047f4-124">If your add-in has more than one contextual tab and there are contexts in which more than one is visible, they appear in the order in which they are defined in your add-in.</span></span> <span data-ttu-id="047f4-125"> (方向与 Office 语言的方向相同;也就是说，使用从左到右的语言从左到右，但从右到左使用从右到左的语言。) 请参阅"定义选项卡上出现的组和控件[](#define-the-groups-and-controls-that-appear-on-the-tab)"，详细了解如何定义它们。</span><span class="sxs-lookup"><span data-stu-id="047f4-125">(The direction is the same direction as the Office language; that is, is left-to-right in left-to-right languages, but right-to-left in right-to-left languages.) See [Define the groups and controls that appear on the tab](#define-the-groups-and-controls-that-appear-on-the-tab) for details about how you define them.</span></span>
- <span data-ttu-id="047f4-126">如果多个加载项具有特定上下文中可见的上下文选项卡，则它们按加载项的启动顺序显示。</span><span class="sxs-lookup"><span data-stu-id="047f4-126">If more than one add-in has a contextual tab that is visible in a specific context, then they appear in the order in which the add-ins were launched.</span></span>
- <span data-ttu-id="047f4-127">自定义 *上下文* 选项卡与自定义核心选项卡不同，不会永久添加到 Office 应用程序的功能区。</span><span class="sxs-lookup"><span data-stu-id="047f4-127">Custom *contextual* tabs, unlike custom core tabs, are not added permanently to the Office application's ribbon.</span></span> <span data-ttu-id="047f4-128">它们仅存在于运行加载项的 Office 文档中。</span><span class="sxs-lookup"><span data-stu-id="047f4-128">They are present only in Office documents on which your add-in is running.</span></span>

## <a name="major-steps-for-including-a-contextual-tab-in-an-add-in"></a><span data-ttu-id="047f4-129">在加载项中添加上下文选项卡的主要步骤</span><span class="sxs-lookup"><span data-stu-id="047f4-129">Major steps for including a contextual tab in an add-in</span></span>

<span data-ttu-id="047f4-130">以下是在加载项中添加自定义上下文选项卡的主要步骤：</span><span class="sxs-lookup"><span data-stu-id="047f4-130">The following are the major steps for including a custom contextual tab in an add-in:</span></span>

1. <span data-ttu-id="047f4-131">将外接程序配置为使用共享运行时。</span><span class="sxs-lookup"><span data-stu-id="047f4-131">Configure the add-in to use a shared runtime.</span></span>
1. <span data-ttu-id="047f4-132">定义选项卡及其上出现的组和控件。</span><span class="sxs-lookup"><span data-stu-id="047f4-132">Define the tab and the groups and controls that appear on it.</span></span>
1. <span data-ttu-id="047f4-133">向 Office 注册上下文选项卡。</span><span class="sxs-lookup"><span data-stu-id="047f4-133">Register the contextual tab with Office.</span></span>
1. <span data-ttu-id="047f4-134">指定选项卡可见时的情况。</span><span class="sxs-lookup"><span data-stu-id="047f4-134">Specify the circumstances when the tab will be visible.</span></span>

## <a name="configure-the-add-in-to-use-a-shared-runtime"></a><span data-ttu-id="047f4-135">配置加载项以使用共享运行时</span><span class="sxs-lookup"><span data-stu-id="047f4-135">Configure the add-in to use a shared runtime</span></span>

<span data-ttu-id="047f4-136">添加自定义上下文选项卡需要加载项使用共享运行时。</span><span class="sxs-lookup"><span data-stu-id="047f4-136">Adding custom contextual tabs requires your add-in to use the shared runtime.</span></span> <span data-ttu-id="047f4-137">有关详细信息，请参阅 [配置外接程序以使用共享运行时](../develop/configure-your-add-in-to-use-a-shared-runtime.md)。</span><span class="sxs-lookup"><span data-stu-id="047f4-137">For more information, see [Configure an add-in to use a shared runtime](../develop/configure-your-add-in-to-use-a-shared-runtime.md).</span></span>

## <a name="define-the-groups-and-controls-that-appear-on-the-tab"></a><span data-ttu-id="047f4-138">定义显示在选项卡上的组和控件</span><span class="sxs-lookup"><span data-stu-id="047f4-138">Define the groups and controls that appear on the tab</span></span>

<span data-ttu-id="047f4-139">与使用清单中的 XML 定义的自定义核心选项卡不同，自定义上下文选项卡在运行时使用 JSON blob 定义。</span><span class="sxs-lookup"><span data-stu-id="047f4-139">Unlike custom core tabs, which are defined with XML in the manifest, custom contextual tabs are defined at runtime with a JSON blob.</span></span> <span data-ttu-id="047f4-140">代码将 blob 解析为 JavaScript 对象，然后将该对象传递给 [Office.ribbon.requestCreateControls](/javascript/api/office/office.ribbon?view=common-js&preserve-view=true#requestCreateControls-tabDefinition-) 方法。</span><span class="sxs-lookup"><span data-stu-id="047f4-140">Your code parses the blob into a JavaScript object, and then passes the object to the [Office.ribbon.requestCreateControls](/javascript/api/office/office.ribbon?view=common-js&preserve-view=true#requestCreateControls-tabDefinition-) method.</span></span> <span data-ttu-id="047f4-141">自定义上下文选项卡仅存在于加载项当前运行的文档中。</span><span class="sxs-lookup"><span data-stu-id="047f4-141">Custom contextual tabs are only present in documents on which your add-in is currently running.</span></span> <span data-ttu-id="047f4-142">这不同于在安装加载项时添加到 Office 应用程序功能区的自定义核心选项卡，当打开另一个文档时，这些选项卡仍保持显示状态。</span><span class="sxs-lookup"><span data-stu-id="047f4-142">This is different from custom core tabs which are added to the Office application ribbon when the add-in is installed and remain present when another document is opened.</span></span> <span data-ttu-id="047f4-143">此外 `requestCreateControls` ，此方法只能在加载项会话中运行一次。</span><span class="sxs-lookup"><span data-stu-id="047f4-143">Also, the `requestCreateControls` method can be run only once in a session of your add-in.</span></span> <span data-ttu-id="047f4-144">如果再次调用它，将引发错误。</span><span class="sxs-lookup"><span data-stu-id="047f4-144">If it is called again, an error is thrown.</span></span>

> [!NOTE]
> <span data-ttu-id="047f4-145">JSON blob 的属性和子属性 (和键名称) 的结构大致与清单 XML 中 [CustomTab](../reference/manifest/customtab.md) 元素及其后代元素的结构平行。</span><span class="sxs-lookup"><span data-stu-id="047f4-145">The structure of the JSON blob's properties and subproperties (and the key names) is roughly parallel to the structure of the [CustomTab](../reference/manifest/customtab.md) element and its descendant elements in the manifest XML.</span></span>

<span data-ttu-id="047f4-146">我们将分步构造上下文选项卡 JSON blob 的示例。</span><span class="sxs-lookup"><span data-stu-id="047f4-146">We'll construct an example of a contextual tabs JSON blob step-by-step.</span></span> <span data-ttu-id="047f4-147"> (上下文选项卡 JSON 的完整架构位于 [dynamic-ribbon.schema.js上](https://developer.microsoft.com/json-schemas/office-js/dynamic-ribbon.schema.json)。</span><span class="sxs-lookup"><span data-stu-id="047f4-147">(The full schema for the contextual tab JSON is at [dynamic-ribbon.schema.json](https://developer.microsoft.com/json-schemas/office-js/dynamic-ribbon.schema.json).</span></span> <span data-ttu-id="047f4-148">此链接在上下文选项卡的早期预览阶段可能无法运行。</span><span class="sxs-lookup"><span data-stu-id="047f4-148">This link may not be working in the early preview period for contextual tabs.</span></span> <span data-ttu-id="047f4-149">如果链接不工作，您可以在 .) 上的草稿 [dynamic-ribbon.schema.js](https://github.com/OfficeDev/testing-assets/tree/master/jsonschema/dynamic-ribbon.schema.json)找到架构的最新草稿（如果您使用 Visual Studio Code，您可以使用此文件获取 IntelliSense 并验证 JSON。</span><span class="sxs-lookup"><span data-stu-id="047f4-149">If the link is not working, you can find the latest draft of the schema at [draft dynamic-ribbon.schema.json](https://github.com/OfficeDev/testing-assets/tree/master/jsonschema/dynamic-ribbon.schema.json).) If you are working in Visual Studio Code, you can use this file to get IntelliSense and to validate your JSON.</span></span> <span data-ttu-id="047f4-150">有关详细信息，请参阅编辑 [JSON 和Visual Studio代码 - JSON 架构和设置](https://code.visualstudio.com/docs/languages/json#_json-schemas-and-settings)。</span><span class="sxs-lookup"><span data-stu-id="047f4-150">For more information, see [Editing JSON with Visual Studio Code - JSON schemas and settings](https://code.visualstudio.com/docs/languages/json#_json-schemas-and-settings).</span></span>


1. <span data-ttu-id="047f4-151">首先创建一个 JSON 字符串，该字符串具有名为 和 的两个 `actions` 数组属性 `tabs` 。</span><span class="sxs-lookup"><span data-stu-id="047f4-151">Begin by creating a JSON string with two array properties named `actions` and `tabs`.</span></span> <span data-ttu-id="047f4-152">该数组是上下文选项卡上的控件可以执行的所有函数 `actions` 的规范。数组 `tabs` 定义一个或多个上下文选项卡，最多 *10 个*。</span><span class="sxs-lookup"><span data-stu-id="047f4-152">The `actions` array is a specification of all the functions that can be executed by controls on the contextual tab. The `tabs` array defines one or more contextual tabs, *up to a maximum of 10*.</span></span>

    ```json
    '{
      "actions": [

      ],
      "tabs": [

      ]
    }'
    ```

1. <span data-ttu-id="047f4-153">上下文选项卡的这个简单示例将只有一个按钮，因此只有一个操作。</span><span class="sxs-lookup"><span data-stu-id="047f4-153">This simple example of a contextual tab will have only a single button and, thus, only a single action.</span></span> <span data-ttu-id="047f4-154">将以下内容添加为数组的唯一 `actions` 成员。</span><span class="sxs-lookup"><span data-stu-id="047f4-154">Add the following as the only member of the `actions` array.</span></span> <span data-ttu-id="047f4-155">关于此标记，请注意：</span><span class="sxs-lookup"><span data-stu-id="047f4-155">About this markup, note:</span></span>

    - <span data-ttu-id="047f4-156">和 `id` `type` 属性是必需的。</span><span class="sxs-lookup"><span data-stu-id="047f4-156">The `id` and `type` properties are mandatory.</span></span>
    - <span data-ttu-id="047f4-157">其值 `type` 可以是"ExecuteFunction"或"ShowTaskpane"。</span><span class="sxs-lookup"><span data-stu-id="047f4-157">The value of `type` can be either "ExecuteFunction" or "ShowTaskpane".</span></span>
    - <span data-ttu-id="047f4-158">该属性 `functionName` 仅在值为 `type` `ExecuteFunction` 时使用。</span><span class="sxs-lookup"><span data-stu-id="047f4-158">The `functionName` property is only used when the value of `type` is `ExecuteFunction`.</span></span> <span data-ttu-id="047f4-159">它是 FunctionFile 中定义的函数的名称。</span><span class="sxs-lookup"><span data-stu-id="047f4-159">It is the name of a function defined in the FunctionFile.</span></span> <span data-ttu-id="047f4-160">有关 FunctionFile 详细信息，请参阅 [外接程序命令的基本概念](add-in-commands.md)。</span><span class="sxs-lookup"><span data-stu-id="047f4-160">For more information about the FunctionFile, see [Basic concepts for Add-in Commands](add-in-commands.md).</span></span>
    - <span data-ttu-id="047f4-161">在稍后的步骤中，您将此操作映射到上下文选项卡上的按钮。</span><span class="sxs-lookup"><span data-stu-id="047f4-161">In a later step, you will map this action to a button on the contextual tab.</span></span>

    ```json
    {
      "id": "executeWriteData",
      "type": "ExecuteFunction",
      "functionName": "writeData"
    }
   ```

1. <span data-ttu-id="047f4-162">将以下内容添加为数组的唯一 `tabs` 成员。</span><span class="sxs-lookup"><span data-stu-id="047f4-162">Add the following as the only member of the `tabs` array.</span></span> <span data-ttu-id="047f4-163">关于此标记，请注意：</span><span class="sxs-lookup"><span data-stu-id="047f4-163">About this markup, note:</span></span>

    - <span data-ttu-id="047f4-164">`id` 属性是必需的。</span><span class="sxs-lookup"><span data-stu-id="047f4-164">The `id` property is required.</span></span> <span data-ttu-id="047f4-165">使用外接程序中所有上下文选项卡中唯一的简短描述性 ID。</span><span class="sxs-lookup"><span data-stu-id="047f4-165">Use a brief, descriptive ID that is unique among all contextual tabs in your add-in.</span></span>
    - <span data-ttu-id="047f4-166">`label` 属性是必需的。</span><span class="sxs-lookup"><span data-stu-id="047f4-166">The `label` property is required.</span></span> <span data-ttu-id="047f4-167">它是一个用户友好字符串，用作上下文选项卡的标签。</span><span class="sxs-lookup"><span data-stu-id="047f4-167">It is a user-friendly string to serve as the label of the contextual tab.</span></span>
    - <span data-ttu-id="047f4-168">`groups` 属性是必需的。</span><span class="sxs-lookup"><span data-stu-id="047f4-168">The `groups` property is required.</span></span> <span data-ttu-id="047f4-169">它定义将在选项卡上出现的控件组。它必须至少有一个成员，且 *不超过 20 个*。</span><span class="sxs-lookup"><span data-stu-id="047f4-169">It defines the groups of controls that will appear on the tab. It must have at least one member *and no more than 20*.</span></span> <span data-ttu-id="047f4-170"> (自定义上下文选项卡上可以具有的控件数量也具有一些限制，这也会限制你拥有多少个组。</span><span class="sxs-lookup"><span data-stu-id="047f4-170">(There are also limits on the number of controls that you can have on a custom contextual tab and that will also constrain how many groups that you have.</span></span> <span data-ttu-id="047f4-171">有关详细信息，请参阅下一步。) </span><span class="sxs-lookup"><span data-stu-id="047f4-171">See the next step for more information.)</span></span>

    > [!NOTE]
    > <span data-ttu-id="047f4-172">Tab 对象还可以具有一个可选属性，该属性指定在加载项启动时选项卡是否立即 `visible` 可见。</span><span class="sxs-lookup"><span data-stu-id="047f4-172">The tab object can also have an optional `visible` property that specifies whether the tab is visible immediately when the add-in starts up.</span></span> <span data-ttu-id="047f4-173">由于上下文选项卡通常是隐藏的，直到用户事件触发其可见性 (例如用户在文档中选择某种类型的实体) ，该属性默认为不存在时 `visible` `false` 。</span><span class="sxs-lookup"><span data-stu-id="047f4-173">Since contextual tabs are normally hidden until a user event triggers their visibility (such as the user selecting an entity of some type in the document), the `visible` property defaults to `false` when not present.</span></span> <span data-ttu-id="047f4-174">在稍后的部分中，我们将展示如何设置该属性 `true` 以响应事件。</span><span class="sxs-lookup"><span data-stu-id="047f4-174">In a later section, we show how to set the property to `true` in response to an event.</span></span>

    ```json
    {
      "id": "CtxTab1",
      "label": "Contoso Data",
      "groups": [

      ]
    }
    ```

1. <span data-ttu-id="047f4-175">在简单的正在进行的示例中，上下文选项卡只有一个组。</span><span class="sxs-lookup"><span data-stu-id="047f4-175">In the simple ongoing example, the contextual tab has only a single group.</span></span> <span data-ttu-id="047f4-176">将以下内容添加为数组的唯一 `groups` 成员。</span><span class="sxs-lookup"><span data-stu-id="047f4-176">Add the following as the only member of the `groups` array.</span></span> <span data-ttu-id="047f4-177">关于此标记，请注意：</span><span class="sxs-lookup"><span data-stu-id="047f4-177">About this markup, note:</span></span>

    - <span data-ttu-id="047f4-178">所有属性都是必需的。</span><span class="sxs-lookup"><span data-stu-id="047f4-178">All the properties are required.</span></span>
    - <span data-ttu-id="047f4-179">该属性在选项卡的所有组中必须是唯一的。 `id` 请使用简短的描述性 ID。</span><span class="sxs-lookup"><span data-stu-id="047f4-179">The `id` property must be unique among all the groups in the tab. Use a brief, descriptive ID.</span></span>
    - <span data-ttu-id="047f4-180">这是 `label` 一个用户友好字符串，用作组的标签。</span><span class="sxs-lookup"><span data-stu-id="047f4-180">The `label` is a user-friendly string to serve as the label of the group.</span></span>
    - <span data-ttu-id="047f4-181">该属性的值是一组对象，这些对象根据功能区的大小和 Office 应用程序窗口指定组将在功能区上具有 `icon` 的图标。</span><span class="sxs-lookup"><span data-stu-id="047f4-181">The `icon` property's value is an array of objects that specify the icons that the group will have on the ribbon depending on the size of the ribbon and the Office application window.</span></span>
    - <span data-ttu-id="047f4-182">`controls`该属性的值是指定组中按钮和菜单的对象数组。</span><span class="sxs-lookup"><span data-stu-id="047f4-182">The `controls` property's value is an array of objects that specify the buttons and menus in the group.</span></span> <span data-ttu-id="047f4-183">组中必须至少有一个且 *不超过 6 个*。</span><span class="sxs-lookup"><span data-stu-id="047f4-183">There must be at least one and *no more than 6 in a group*.</span></span>

    > [!IMPORTANT]
    > <span data-ttu-id="047f4-184">*整个选项卡上的控件总数不能超过 20 个。*</span><span class="sxs-lookup"><span data-stu-id="047f4-184">*The total number of controls on the whole tab can be no more than 20.*</span></span> <span data-ttu-id="047f4-185">例如，可以有 3 个组，每个组有 6 个控件，第四个组有 2 个控件，但不能有 4 个组，每个组有 6 个控件。</span><span class="sxs-lookup"><span data-stu-id="047f4-185">For example, you could have 3 groups with 6 controls each, and a fourth group with 2 controls, but you cannot have 4 groups with 6 controls each.</span></span>  

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

1. <span data-ttu-id="047f4-186">每个组都必须具有至少两个大小的图标：32x32 像素和 80x80 像素。</span><span class="sxs-lookup"><span data-stu-id="047f4-186">Every group must have an icon of at least two sizes, 32x32 px and 80x80 px.</span></span> <span data-ttu-id="047f4-187">（可选）还可以具有大小为 16x16 像素、20x20 像素、24x24 像素、40x40 像素、48x48 像素和 64x64 像素的图标。</span><span class="sxs-lookup"><span data-stu-id="047f4-187">Optionally, you can also have icons of sizes 16x16 px, 20x20 px, 24x24 px, 40x40 px, 48x48 px, and 64x64 px.</span></span> <span data-ttu-id="047f4-188">Office 根据功能区的大小和 Office 应用程序窗口决定使用哪个图标。</span><span class="sxs-lookup"><span data-stu-id="047f4-188">Office decides which icon to use based on the size of the ribbon and Office application window.</span></span> <span data-ttu-id="047f4-189">将以下对象添加到图标数组。</span><span class="sxs-lookup"><span data-stu-id="047f4-189">Add the following objects to the icon array.</span></span> <span data-ttu-id="047f4-190"> (如果窗口和功能区大小足够大，组中至少有一个控件可以显示，则不显示任何组图标。</span><span class="sxs-lookup"><span data-stu-id="047f4-190">(If the window and ribbon sizes are large enough for at least one of the *controls* on the group to appear, then no group icon at all appears.</span></span> <span data-ttu-id="047f4-191">例如，在缩小和展开 Word 窗口时，观察 Word 功能区上的 **Styles** 组。) 关于此标记，请注意：</span><span class="sxs-lookup"><span data-stu-id="047f4-191">For an example, watch the **Styles** group on the Word ribbon as you shrink and expand the Word window.) About this markup, note:</span></span>

    - <span data-ttu-id="047f4-192">这两个属性都是必需的。</span><span class="sxs-lookup"><span data-stu-id="047f4-192">Both the properties are required.</span></span>
    - <span data-ttu-id="047f4-193">`size`属性度量单位为像素。</span><span class="sxs-lookup"><span data-stu-id="047f4-193">The `size` property unit of measure is pixels.</span></span> <span data-ttu-id="047f4-194">图标始终为正方形，因此数字为高度和宽度。</span><span class="sxs-lookup"><span data-stu-id="047f4-194">Icons are always square, so the number is both the height and the width.</span></span>
    - <span data-ttu-id="047f4-195">该属性 `sourceLocation` 指定图标的完整 URL。</span><span class="sxs-lookup"><span data-stu-id="047f4-195">The `sourceLocation` property specifies the full URL to the icon.</span></span>

    > [!IMPORTANT]
    > <span data-ttu-id="047f4-196">与在从开发环境移动到生产 (（如将域从 localhost 更改为 contoso.com) ）时，通常必须更改加载项清单中的 URL 一样，您还必须更改上下文选项卡 JSON 中的 URL。</span><span class="sxs-lookup"><span data-stu-id="047f4-196">Just as you typically must change the URLs in the add-in's manifest when you move from development to production (such as changing the domain from localhost to contoso.com), you must also change the URLs in your contextual tabs JSON.</span></span>

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

1. <span data-ttu-id="047f4-197">在我们的简单正在进行的示例中，该组只有一个按钮。</span><span class="sxs-lookup"><span data-stu-id="047f4-197">In our simple ongoing example, the group has only a single button.</span></span> <span data-ttu-id="047f4-198">将以下对象添加为数组的唯一 `controls` 成员。</span><span class="sxs-lookup"><span data-stu-id="047f4-198">Add the following object as the only member of the `controls` array.</span></span> <span data-ttu-id="047f4-199">关于此标记，请注意：</span><span class="sxs-lookup"><span data-stu-id="047f4-199">About this markup, note:</span></span>

    - <span data-ttu-id="047f4-200">除属性外 `enabled` ，其他所有属性都是必需的。</span><span class="sxs-lookup"><span data-stu-id="047f4-200">All the properties, except `enabled`, are required.</span></span>
    - <span data-ttu-id="047f4-201">`type` 指定控件的类型。</span><span class="sxs-lookup"><span data-stu-id="047f4-201">`type` specifies the type of control.</span></span> <span data-ttu-id="047f4-202">值可以是"Button"、"Menu"或"MobileButton"。</span><span class="sxs-lookup"><span data-stu-id="047f4-202">The values can be "Button", "Menu", or "MobileButton".</span></span>
    - <span data-ttu-id="047f4-203">`id` 最多为 125 个字符。</span><span class="sxs-lookup"><span data-stu-id="047f4-203">`id` can be up to 125 characters.</span></span> 
    - <span data-ttu-id="047f4-204">`actionId` 必须是数组中定义的操作 `actions` ID。</span><span class="sxs-lookup"><span data-stu-id="047f4-204">`actionId` must be the ID of an action defined in the `actions` array.</span></span> <span data-ttu-id="047f4-205"> (请参阅本节的步骤 1.) </span><span class="sxs-lookup"><span data-stu-id="047f4-205">(See step 1 of this section.)</span></span>
    - <span data-ttu-id="047f4-206">`label` 是用作按钮标签的用户友好字符串。</span><span class="sxs-lookup"><span data-stu-id="047f4-206">`label` is a user-friendly string to serve as the label of the button.</span></span>
    - <span data-ttu-id="047f4-207">`superTip` 表示工具提示的丰富形式。</span><span class="sxs-lookup"><span data-stu-id="047f4-207">`superTip` represents a rich form of tool tip.</span></span> <span data-ttu-id="047f4-208">和 `title` `description` 属性都是必需的。</span><span class="sxs-lookup"><span data-stu-id="047f4-208">Both the `title` and `description` properties are required.</span></span>
    - <span data-ttu-id="047f4-209">`icon` 指定按钮的图标。</span><span class="sxs-lookup"><span data-stu-id="047f4-209">`icon` specifies the icons for the button.</span></span> <span data-ttu-id="047f4-210">前面有关组图标的备注也适用于此处。</span><span class="sxs-lookup"><span data-stu-id="047f4-210">The previous remarks about the group icon apply here too.</span></span>
    - <span data-ttu-id="047f4-211">`enabled` (可选) 指定在上下文选项卡启动时是否启用按钮。</span><span class="sxs-lookup"><span data-stu-id="047f4-211">`enabled` (optional) specifies whether the button is enabled when the contextual tab appears starts up.</span></span> <span data-ttu-id="047f4-212">如果不存在，则默认为 `true` 。</span><span class="sxs-lookup"><span data-stu-id="047f4-212">The default if not present is `true`.</span></span> 

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
 
<span data-ttu-id="047f4-213">以下是 JSON blob 的完整示例：</span><span class="sxs-lookup"><span data-stu-id="047f4-213">The following is the complete example of the JSON blob:</span></span>

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

## <a name="register-the-contextual-tab-with-office-with-requestcreatecontrols"></a><span data-ttu-id="047f4-214">使用 requestCreateControls 向 Office 注册上下文选项卡</span><span class="sxs-lookup"><span data-stu-id="047f4-214">Register the contextual tab with Office with requestCreateControls</span></span>

<span data-ttu-id="047f4-215">上下文选项卡通过调用[Office.ribbon.requestCreateControls 方法注册到 Office。](/javascript/api/office/office.ribbon?view=common-js&preserve-view=true#requestCreateControls_tabDefinition_)</span><span class="sxs-lookup"><span data-stu-id="047f4-215">The contextual tab is registered with Office by calling the [Office.ribbon.requestCreateControls](/javascript/api/office/office.ribbon?view=common-js&preserve-view=true#requestCreateControls_tabDefinition_) method.</span></span> <span data-ttu-id="047f4-216">这通常在分配给方法的函数中完成， `Office.initialize` 或随方法 `Office.onReady` 一起完成。</span><span class="sxs-lookup"><span data-stu-id="047f4-216">This is typically done in either the function that is assigned to `Office.initialize` or with the `Office.onReady` method.</span></span> <span data-ttu-id="047f4-217">有关这些方法和初始化外接程序的更多信息，请参阅["初始化 Office 外接程序"。](../develop/initialize-add-in.md)</span><span class="sxs-lookup"><span data-stu-id="047f4-217">For more about these methods and initializing the add-in, see [Initialize your Office Add-in](../develop/initialize-add-in.md).</span></span> <span data-ttu-id="047f4-218">但是，可以在初始化后随时调用该方法。</span><span class="sxs-lookup"><span data-stu-id="047f4-218">You can, however, call the method anytime after initialization.</span></span>

> [!IMPORTANT]
> <span data-ttu-id="047f4-219">`requestCreateControls`该方法只能在加载项的给定会话中调用一次。</span><span class="sxs-lookup"><span data-stu-id="047f4-219">The `requestCreateControls` method can be called only once in a given session of an add-in.</span></span> <span data-ttu-id="047f4-220">如果再次调用错误，将引发错误。</span><span class="sxs-lookup"><span data-stu-id="047f4-220">An error is thrown if it is called again.</span></span>

<span data-ttu-id="047f4-221">示例如下。</span><span class="sxs-lookup"><span data-stu-id="047f4-221">The following is an example.</span></span> <span data-ttu-id="047f4-222">请注意，必须先使用该方法将 JSON 字符串转换为 JavaScript 对象，然后才能将其 `JSON.parse` 传递给 JavaScript 函数。</span><span class="sxs-lookup"><span data-stu-id="047f4-222">Note that the JSON string must be converted to a JavaScript object with the `JSON.parse` method before it can be passed to a JavaScript function.</span></span>

```javascript
Office.onReady(async () => {
    const contextualTabJSON = ` ... `; // Assign the JSON string such as the one at the end of the preceding section.
    const contextualTab = JSON.parse(contextualTabJSON);
    await Office.ribbon.requestCreateControls(contextualTab);
});
```

## <a name="specify-the-contexts-when-the-tab-will-be-visible-with-requestupdate"></a><span data-ttu-id="047f4-223">使用 requestUpdate 指定选项卡何时可见</span><span class="sxs-lookup"><span data-stu-id="047f4-223">Specify the contexts when the tab will be visible with requestUpdate</span></span>

<span data-ttu-id="047f4-224">通常，当用户启动的事件更改加载项上下文时，应显示自定义上下文选项卡。</span><span class="sxs-lookup"><span data-stu-id="047f4-224">Typically, a custom contextual tab should appear when a user-initiated event changes the add-in context.</span></span> <span data-ttu-id="047f4-225">考虑在激活 Excel 工作簿的默认工作表上的图表 (且仅在激活时，选项卡) 可见。</span><span class="sxs-lookup"><span data-stu-id="047f4-225">Consider a scenario in which the tab should be visible when, and only when, a chart (on the default worksheet of an Excel workbook) is activated.</span></span>

<span data-ttu-id="047f4-226">首先分配处理程序。</span><span class="sxs-lookup"><span data-stu-id="047f4-226">Begin by assigning handlers.</span></span> <span data-ttu-id="047f4-227">此方法中通常完成此操作，如以下示例所示，该示例将 (步骤) 中创建的处理程序分配给工作表中所有图表的和 `Office.onReady` `onActivated` `onDeactivated` 事件。</span><span class="sxs-lookup"><span data-stu-id="047f4-227">This is commonly done in the `Office.onReady` method as in the following example which assigns handlers (created in a later step) to the `onActivated` and `onDeactivated` events of all the charts in the worksheet.</span></span>

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

<span data-ttu-id="047f4-228">接下来，定义处理程序。</span><span class="sxs-lookup"><span data-stu-id="047f4-228">Next, define the handlers.</span></span> <span data-ttu-id="047f4-229">下面是一个简单示例，但请参阅本文稍后介绍的"处理 `showDataTab` [HostRestartNeeded](#handling-the-hostrestartneeded-error) 错误"，了解函数的更可靠版本。</span><span class="sxs-lookup"><span data-stu-id="047f4-229">The following is a simple example of a `showDataTab`, but see [Handling the HostRestartNeeded error](#handling-the-hostrestartneeded-error) later in this article for a more robust version of the function.</span></span> <span data-ttu-id="047f4-230">关于此代码，请注意以下几点：</span><span class="sxs-lookup"><span data-stu-id="047f4-230">About this code, note:</span></span>

- <span data-ttu-id="047f4-231">Office 控制何时更新功能区的状态。</span><span class="sxs-lookup"><span data-stu-id="047f4-231">Office controls when it updates the state of the ribbon.</span></span> <span data-ttu-id="047f4-232">[Office.ribbon.requestUpdate](/javascript/api/office/office.ribbon?view=common-js&preserve-view=true#requestupdate-input-)方法将更新请求排成队列。</span><span class="sxs-lookup"><span data-stu-id="047f4-232">The  [Office.ribbon.requestUpdate](/javascript/api/office/office.ribbon?view=common-js&preserve-view=true#requestupdate-input-) method queues a request to update.</span></span> <span data-ttu-id="047f4-233">一旦将请求排入队列，该方法将解析该对象，而不是功能 `Promise` 区实际更新时。</span><span class="sxs-lookup"><span data-stu-id="047f4-233">The method will resolve the `Promise` object as soon as it has queued the request, not when the ribbon actually updates.</span></span>
- <span data-ttu-id="047f4-234">该方法的参数是 `requestUpdate` [RibbonUpdaterData](/javascript/api/office/office.ribbonupdaterdata) 对象， (1) 按其 ID 指定 *选项卡* ，而 (2) 指定选项卡的可见性。</span><span class="sxs-lookup"><span data-stu-id="047f4-234">The parameter for the `requestUpdate` method is a [RibbonUpdaterData](/javascript/api/office/office.ribbonupdaterdata) object that (1) specifies the tab by its ID *exactly as specified in the JSON* and (2) specifies visibility of the tab.</span></span>
- <span data-ttu-id="047f4-235">如果多个自定义上下文选项卡应在同一上下文中可见，只需向数组添加其他 `tabs` 选项卡对象。</span><span class="sxs-lookup"><span data-stu-id="047f4-235">If you have more than one custom contextual tab that should be visible in the same context, you simply add additional tab objects to the `tabs` array.</span></span>

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

<span data-ttu-id="047f4-236">隐藏选项卡的处理程序几乎完全相同，只是将 `visible` 该属性设置回 `false` 。</span><span class="sxs-lookup"><span data-stu-id="047f4-236">The handler to hide the tab is nearly identical, except that it sets the `visible` property back to `false`.</span></span>

<span data-ttu-id="047f4-237">Office JavaScript 库还提供了多个 (类型的) ，以便更轻松地构造 `RibbonUpdateData` 对象。</span><span class="sxs-lookup"><span data-stu-id="047f4-237">The Office JavaScript library also provides several interfaces (types) to make it easier to construct the`RibbonUpdateData` object.</span></span> <span data-ttu-id="047f4-238">以下是 `showDataTab` TypeScript 中的函数，它使用这些类型。</span><span class="sxs-lookup"><span data-stu-id="047f4-238">The following is the `showDataTab` function in TypeScript and it makes use of these types.</span></span>

```typescript
const showDataTab = async () => {
    const myContextualTab: Office.Tab = {id: "CtxTab1", visible: true};
    const ribbonUpdater: Office.RibbonUpdaterData = { tabs: [ myContextualTab ]};
    await Office.ribbon.requestUpdate(ribbonUpdater);
}
```

### <a name="toggle-tab-visibility-and-the-enabled-status-of-a-button-at-the-same-time"></a><span data-ttu-id="047f4-239">同时切换选项卡可见性和按钮的启用状态</span><span class="sxs-lookup"><span data-stu-id="047f4-239">Toggle tab visibility and the enabled status of a button at the same time</span></span>

<span data-ttu-id="047f4-240">该方法还用于切换自定义上下文选项卡或自定义核心选项卡上自定义按钮的启用 `requestUpdate` 或禁用状态。有关此内容的详细信息，请参阅["启用和禁用外接程序命令"。](disable-add-in-commands.md)</span><span class="sxs-lookup"><span data-stu-id="047f4-240">The `requestUpdate` method is also used to toggle the enabled or disabled status of a custom button on either a custom contextual tab or a custom core tab. For details about this, see [Enable and Disable Add-in Commands](disable-add-in-commands.md).</span></span> <span data-ttu-id="047f4-241">在某些情况下，你可能希望同时更改选项卡的可见性和按钮的启用状态。</span><span class="sxs-lookup"><span data-stu-id="047f4-241">There may be scenarios in which you want to change both the visibility of a tab and the enabled status of a button at the same time.</span></span> <span data-ttu-id="047f4-242">可以通过单个调用来此操作 `requestUpdate` 。</span><span class="sxs-lookup"><span data-stu-id="047f4-242">You can do this with a single call of `requestUpdate`.</span></span> <span data-ttu-id="047f4-243">下面是一个示例，在使上下文选项卡可见的同时启用核心选项卡上的按钮。</span><span class="sxs-lookup"><span data-stu-id="047f4-243">The following is an example in which a button on a core tab is enabled at the same time as a contextual tab is made visible.</span></span>

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

<span data-ttu-id="047f4-244">在下面的示例中，启用的按钮位于要显示上下文选项卡的同一个选项卡上。</span><span class="sxs-lookup"><span data-stu-id="047f4-244">In the following example, the button that is enabled is on the very same contextual tab that is being made visible.</span></span>

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

## <a name="localizing-the-json-blob"></a><span data-ttu-id="047f4-245">本地化 JSON blob</span><span class="sxs-lookup"><span data-stu-id="047f4-245">Localizing the JSON blob</span></span>

<span data-ttu-id="047f4-246">传递给的 JSON blob 的本地化方式与自定义核心选项卡的清单标记的本地化方式不同 (如清单控件本地化中所述 `requestCreateControls`) 。 [](../develop/localization.md#control-localization-from-the-manifest)</span><span class="sxs-lookup"><span data-stu-id="047f4-246">The JSON blob that is passed to `requestCreateControls` is not localized the same way that the manifest markup for custom core tabs is localized (which is described at [Control localization from the manifest](../develop/localization.md#control-localization-from-the-manifest)).</span></span> <span data-ttu-id="047f4-247">相反，本地化必须在运行时针对每个区域设置使用不同的 JSON blob。</span><span class="sxs-lookup"><span data-stu-id="047f4-247">Instead, the localization must occur at runtime using distinct JSON blobs for each locale.</span></span> <span data-ttu-id="047f4-248">建议您使用用于测试 `switch` [Office.context.displayLanguage](/javascript/api/office/office.context#displayLanguage) 属性的语句。</span><span class="sxs-lookup"><span data-stu-id="047f4-248">We suggest that you use a `switch` statement that tests the [Office.context.displayLanguage](/javascript/api/office/office.context#displayLanguage) property.</span></span> <span data-ttu-id="047f4-249">示例如下：</span><span class="sxs-lookup"><span data-stu-id="047f4-249">The following is an example:</span></span>

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

<span data-ttu-id="047f4-250">然后，代码调用该函数，获取传递给的本地化 `requestCreateControls` blob，如以下示例所示：</span><span class="sxs-lookup"><span data-stu-id="047f4-250">Then your code calls the function to get the localized blob that is passed to `requestCreateControls`, as in the following example:</span></span>

```javascript
var contextualTabJSON = GetContextualTabsJsonSupportedLocale();
```

## <a name="handling-the-hostrestartneeded-error"></a><span data-ttu-id="047f4-251">处理 HostRestartNeeded 错误</span><span class="sxs-lookup"><span data-stu-id="047f4-251">Handling the HostRestartNeeded error</span></span>

<span data-ttu-id="047f4-252">在某些情况下，Office 无法更新功能区，并将返回错误。</span><span class="sxs-lookup"><span data-stu-id="047f4-252">In some scenarios, Office is unable to update the ribbon and will return an error.</span></span> <span data-ttu-id="047f4-253">例如，如果升级了加载项，并且升级后的加载项具有一组不同的自定义加载项命令，则必须关闭并重新打开 Office 应用程序。</span><span class="sxs-lookup"><span data-stu-id="047f4-253">For example, if the add-in is upgraded and the upgraded add-in has a different set of custom add-in commands, then the Office application must be closed and reopened.</span></span> <span data-ttu-id="047f4-254">在此之前，`requestUpdate` 方法将返回错误 `HostRestartNeeded`。</span><span class="sxs-lookup"><span data-stu-id="047f4-254">Until it is, the `requestUpdate` method will return the error `HostRestartNeeded`.</span></span> <span data-ttu-id="047f4-255">以下是如何处理此错误的示例。</span><span class="sxs-lookup"><span data-stu-id="047f4-255">The following is an example of how to handle this error.</span></span> <span data-ttu-id="047f4-256">在此示例中，`reportError` 方法向用户显示错误。</span><span class="sxs-lookup"><span data-stu-id="047f4-256">In this case, the `reportError` method displays the error to the user.</span></span>

```javascript
function showDataTab() {
    try {
        await Office.ribbon.requestUpdate({
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
