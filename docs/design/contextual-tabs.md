---
title: 在 Office 外接程序中创建自定义上下文选项卡
description: 了解如何将自定义上下文选项卡添加到 Office 外接程序。
ms.date: 11/20/2020
localization_priority: Normal
ms.openlocfilehash: 49a773aca0651b88c972c24a4cde0aa1e300d5e7
ms.sourcegitcommit: 6619e07cdfa68f9fa985febd5f03caf7aee57d5e
ms.translationtype: MT
ms.contentlocale: zh-CN
ms.lasthandoff: 11/30/2020
ms.locfileid: "49505552"
---
# <a name="create-custom-contextual-tabs-in-office-add-ins-preview"></a><span data-ttu-id="fe783-103">在 Office 外接程序中创建自定义上下文选项卡 (预览) </span><span class="sxs-lookup"><span data-stu-id="fe783-103">Create custom contextual tabs in Office Add-ins (preview)</span></span>

<span data-ttu-id="fe783-104">上下文选项卡是 Office 功能区中的一个隐藏的选项卡控件，当 Office 文档中发生指定事件时，该选项卡将显示在 "选项卡" 行中。</span><span class="sxs-lookup"><span data-stu-id="fe783-104">A contextual tab is a hidden tab control in the Office ribbon that is displayed in the tab row when a specified event occurs in the Office document.</span></span> <span data-ttu-id="fe783-105">例如，选择表时 Excel 功能区上显示的 " **表设计** " 选项卡。</span><span class="sxs-lookup"><span data-stu-id="fe783-105">For example, the **Table Design** tab that appears on the Excel ribbon when a table is selected.</span></span> <span data-ttu-id="fe783-106">您可以在 Office 外接程序中包含自定义上下文选项卡，并通过创建更改可见性的事件处理程序来指定它们何时可见或隐藏。</span><span class="sxs-lookup"><span data-stu-id="fe783-106">You can include custom contextual tabs in your Office add-in and specify when they are visible or hidden, by creating event handlers that change the visibility.</span></span> <span data-ttu-id="fe783-107"> (但是，自定义上下文选项卡对焦点更改没有响应。 ) </span><span class="sxs-lookup"><span data-stu-id="fe783-107">(However, custom contextual tabs do not respond to focus changes.)</span></span>

> [!NOTE]
> <span data-ttu-id="fe783-108">本文假定你熟悉以下文档。</span><span class="sxs-lookup"><span data-stu-id="fe783-108">This article assumes that you are familiar with the following documentation.</span></span> <span data-ttu-id="fe783-109">如果你最近未使用加载项命令（自定义菜单项和功能区按钮），请查看该文档。</span><span class="sxs-lookup"><span data-stu-id="fe783-109">Please review it if you haven't worked with Add-in Commands (custom menu items and ribbon buttons) recently.</span></span>
>
> - [<span data-ttu-id="fe783-110">加载项命令的基本概念</span><span class="sxs-lookup"><span data-stu-id="fe783-110">Basic concepts for Add-in Commands</span></span>](add-in-commands.md)

> [!IMPORTANT]
> <span data-ttu-id="fe783-111">"自定义上下文选项卡" 处于预览阶段。</span><span class="sxs-lookup"><span data-stu-id="fe783-111">Custom contextual tabs are in preview.</span></span> <span data-ttu-id="fe783-112">请在开发或测试环境中试用它们，但不要将其添加到生产外接加载项中。</span><span class="sxs-lookup"><span data-stu-id="fe783-112">Please experiment with them in a development or testing environment but don't add them to a production add-in.</span></span>
>
> <span data-ttu-id="fe783-113">自定义上下文选项卡目前仅在 Excel 中受支持，并且仅在这些平台和生成上受支持：</span><span class="sxs-lookup"><span data-stu-id="fe783-113">Custom contextual tabs are currently only supported on Excel and only on these platforms and builds:</span></span>
>
> - <span data-ttu-id="fe783-114">Windows (的 Excel 仅适用于 Microsoft 365，而不是永久许可证) ：版本 2011 (内部版本 13426.20274) 。</span><span class="sxs-lookup"><span data-stu-id="fe783-114">Excel on Windows (Microsoft 365 only, not perpetual license): Version 2011 (Build 13426.20274).</span></span> <span data-ttu-id="fe783-115">您的 Microsoft 365 订阅可能需要在 [当前频道 (预览) ](https://insider.office.com/join/windows) 以前称为 "每月频道 (目标) " 或 "内幕慢速"。</span><span class="sxs-lookup"><span data-stu-id="fe783-115">Your Microsoft 365 subscription may need to be on the [Current Channel (Preview)](https://insider.office.com/join/windows) formerly called "Monthly Channel (Targeted)" or "Insider Slow".</span></span>

> [!NOTE]
> <span data-ttu-id="fe783-116">自定义上下文选项卡仅适用于支持以下要求集的平台。</span><span class="sxs-lookup"><span data-stu-id="fe783-116">Custom contextual tabs work only on platforms that support the following requirement sets.</span></span> <span data-ttu-id="fe783-117">有关要求集以及如何使用它们的详细信息，请参阅 [指定 Office 应用程序和 API 要求](../develop/specify-office-hosts-and-api-requirements.md)。</span><span class="sxs-lookup"><span data-stu-id="fe783-117">For more about requirement sets and how to work with them, see [Specify Office applications and API requirements](../develop/specify-office-hosts-and-api-requirements.md).</span></span>
>
> - [<span data-ttu-id="fe783-118">SharedRuntime 1。1</span><span class="sxs-lookup"><span data-stu-id="fe783-118">SharedRuntime 1.1</span></span>](../reference/requirement-sets/shared-runtime-requirement-sets.md)

## <a name="behavior-of-custom-contextual-tabs"></a><span data-ttu-id="fe783-119">自定义上下文选项卡的行为</span><span class="sxs-lookup"><span data-stu-id="fe783-119">Behavior of custom contextual tabs</span></span>

<span data-ttu-id="fe783-120">自定义上下文选项卡的用户体验遵循内置 Office 上下文选项卡的模式。</span><span class="sxs-lookup"><span data-stu-id="fe783-120">The user experience for custom contextual tabs follows the pattern of built-in Office contextual tabs.</span></span> <span data-ttu-id="fe783-121">以下是放置自定义上下文选项卡的基本原则：</span><span class="sxs-lookup"><span data-stu-id="fe783-121">The following are the basic principles for the placement custom contextual tabs:</span></span>

- <span data-ttu-id="fe783-122">当自定义上下文选项卡可见时，它将显示在功能区的右端。</span><span class="sxs-lookup"><span data-stu-id="fe783-122">When a custom contextual tab is visible, it appears on the right end of the ribbon.</span></span>
- <span data-ttu-id="fe783-123">如果一个或多个内置上下文选项卡以及外接程序中的一个或多个自定义上下文选项卡同时可见，则自定义上下文选项卡将始终位于所有内置上下文选项卡的右侧。</span><span class="sxs-lookup"><span data-stu-id="fe783-123">If one or more built-in contextual tabs and one or more custom contextual tabs from add-ins are visible at the same time, the custom contextual tabs are always to the right of all of the built-in contextual tabs.</span></span>
- <span data-ttu-id="fe783-124">如果你的外接程序有多个上下文选项卡，并且存在多个上下文选项卡，则它们将按其在外接程序中的定义顺序显示。</span><span class="sxs-lookup"><span data-stu-id="fe783-124">If your add-in has more than one contextual tab and there are contexts in which more than one is visible, they appear in the order in which they are defined in your add-in.</span></span> <span data-ttu-id="fe783-125"> (方向与 Office 语言的方向相同，则为 [ ](#define-the-groups-and-controls-that-appear-on-the-tab) ; 否则为  。也就是说，从左到右的语言按从左到右的语言，但从右到左的语言为从右到左。 ) 请参阅定义显示在选项卡上的组和控件，了解有关如何定义它们的详细信息。</span><span class="sxs-lookup"><span data-stu-id="fe783-125">(The direction is the same direction as the Office language; that is, is left-to-right in left-to-right languages, but right-to-left in right-to-left languages.) See [Define the groups and controls that appear on the tab](#define-the-groups-and-controls-that-appear-on-the-tab) for details about how you define them.</span></span>
- <span data-ttu-id="fe783-126">如果有多个加载项具有在特定上下文中可见的上下文选项卡，则它们将按加载项启动的顺序显示。</span><span class="sxs-lookup"><span data-stu-id="fe783-126">If more than one add-in has a contextual tab that is visible in a specific context, then they appear in the order in which the add-ins were launched.</span></span>
- <span data-ttu-id="fe783-127">与自定义核心选项卡不同，自定义 *上下文* 选项卡不会永久添加到 Office 应用程序的功能区。</span><span class="sxs-lookup"><span data-stu-id="fe783-127">Custom *contextual* tabs, unlike custom core tabs, are not added permanently to the Office application's ribbon.</span></span> <span data-ttu-id="fe783-128">它们仅存在于运行外接程序的 Office 文档中。</span><span class="sxs-lookup"><span data-stu-id="fe783-128">They are present only in Office documents on which your add-in is running.</span></span>

## <a name="major-steps-for-including-a-contextual-tab-in-an-add-in"></a><span data-ttu-id="fe783-129">在外接程序中包含上下文选项卡的主要步骤</span><span class="sxs-lookup"><span data-stu-id="fe783-129">Major steps for including a contextual tab in an add-in</span></span>

<span data-ttu-id="fe783-130">以下是在外接程序中包含自定义上下文选项卡的主要步骤：</span><span class="sxs-lookup"><span data-stu-id="fe783-130">The following are the major steps for including a custom contextual tab in an add-in:</span></span>

1. <span data-ttu-id="fe783-131">将加载项配置为使用共享运行时。</span><span class="sxs-lookup"><span data-stu-id="fe783-131">Configure the add-in to use a shared runtime.</span></span>
1. <span data-ttu-id="fe783-132">定义选项卡以及显示在其上的组和控件。</span><span class="sxs-lookup"><span data-stu-id="fe783-132">Define the tab and the groups and controls that appear on it.</span></span>
1. <span data-ttu-id="fe783-133">使用 Office 注册上下文选项卡。</span><span class="sxs-lookup"><span data-stu-id="fe783-133">Register the contextual tab with Office.</span></span>
1. <span data-ttu-id="fe783-134">指定选项卡将可见的情况。</span><span class="sxs-lookup"><span data-stu-id="fe783-134">Specify the circumstances when the tab will be visible.</span></span>

## <a name="configure-the-add-in-to-use-a-shared-runtime"></a><span data-ttu-id="fe783-135">将加载项配置为使用共享运行时</span><span class="sxs-lookup"><span data-stu-id="fe783-135">Configure the add-in to use a shared runtime</span></span>

<span data-ttu-id="fe783-136">添加自定义上下文选项卡需要您的外接程序使用共享运行时。</span><span class="sxs-lookup"><span data-stu-id="fe783-136">Adding custom contextual tabs requires your add-in to use the shared runtime.</span></span> <span data-ttu-id="fe783-137">有关详细信息，请参阅 [Configure a 外接程序以使用共享运行时](../excel/configure-your-add-in-to-use-a-shared-runtime.md)。</span><span class="sxs-lookup"><span data-stu-id="fe783-137">For more information, see [Configure an add-in to use a shared runtime](../excel/configure-your-add-in-to-use-a-shared-runtime.md).</span></span>

## <a name="define-the-groups-and-controls-that-appear-on-the-tab"></a><span data-ttu-id="fe783-138">定义选项卡上显示的组和控件</span><span class="sxs-lookup"><span data-stu-id="fe783-138">Define the groups and controls that appear on the tab</span></span>

<span data-ttu-id="fe783-139">与使用清单中的 XML 定义的自定义核心选项卡不同，自定义上下文选项卡是在运行时使用 JSON blob 定义的。</span><span class="sxs-lookup"><span data-stu-id="fe783-139">Unlike custom core tabs, which are defined with XML in the manifest, custom contextual tabs are defined at runtime with a JSON blob.</span></span> <span data-ttu-id="fe783-140">您的代码将 blob 解析为 JavaScript 对象，然后将该对象传递给 [requestCreateControls](/javascript/api/office/office.ribbon?view=common-js&preserve-view=true#requestCreateControls-tabDefinition-) 方法。</span><span class="sxs-lookup"><span data-stu-id="fe783-140">Your code parses the blob into a JavaScript object, and then passes the object to the [Office.ribbon.requestCreateControls](/javascript/api/office/office.ribbon?view=common-js&preserve-view=true#requestCreateControls-tabDefinition-) method.</span></span> <span data-ttu-id="fe783-141">自定义上下文选项卡仅存在于您的外接程序当前运行的文档中。</span><span class="sxs-lookup"><span data-stu-id="fe783-141">Custom contextual tabs are only present in documents on which your add-in is currently running.</span></span> <span data-ttu-id="fe783-142">这不同于自定义核心选项卡，在安装加载项时，这些选项卡会添加到 Office 应用程序功能区中，并且在打开另一个文档时仍然存在。</span><span class="sxs-lookup"><span data-stu-id="fe783-142">This is different from custom core tabs which are added to the Office application ribbon when the add-in is installed and remain present when another document is opened.</span></span> <span data-ttu-id="fe783-143">此外，该 `requestCreateControls` 方法在外接程序的会话中只能运行一次。</span><span class="sxs-lookup"><span data-stu-id="fe783-143">Also, the `requestCreateControls` method can be run only once in a session of your add-in.</span></span> <span data-ttu-id="fe783-144">如果再次调用该方法，则会引发错误。</span><span class="sxs-lookup"><span data-stu-id="fe783-144">If it is called again, an error is thrown.</span></span>

> [!NOTE]
> <span data-ttu-id="fe783-145">JSON blob 的 properties 和子属性的结构 (和键名称) 大致与清单 XML 中的 [CustomTab](../reference/manifest/customtab.md) 元素及其后代元素的结构平行。</span><span class="sxs-lookup"><span data-stu-id="fe783-145">The structure of the JSON blob's properties and subproperties (and the key names) is roughly parallel to the structure of the [CustomTab](../reference/manifest/customtab.md) element and its descendant elements in the manifest XML.</span></span>

<span data-ttu-id="fe783-146">我们将构造一个上下文选项卡 JSON blob 的示例。</span><span class="sxs-lookup"><span data-stu-id="fe783-146">We'll construct an example of a contextual tabs JSON blob step-by-step.</span></span> <span data-ttu-id="fe783-147"> (上下文选项卡 JSON 的完整架构位于 [dynamic-ribbon.schema.js的](https://developer.microsoft.com/json-schemas/office-js/dynamic-ribbon.schema.json)。</span><span class="sxs-lookup"><span data-stu-id="fe783-147">(The full schema for the contextual tab JSON is at [dynamic-ribbon.schema.json](https://developer.microsoft.com/json-schemas/office-js/dynamic-ribbon.schema.json).</span></span> <span data-ttu-id="fe783-148">此链接可能无法在上下文选项卡的早期预览周期中工作。</span><span class="sxs-lookup"><span data-stu-id="fe783-148">This link may not be working in the early preview period for contextual tabs.</span></span> <span data-ttu-id="fe783-149">如果链接未正常运行，则可以在 [草稿 dynamic-ribbon.schema.js上](https://github.com/OfficeDev/testing-assets/tree/master/jsonschema/dynamic-ribbon.schema.json)找到架构的最新草案。 ) 如果您在 Visual Studio Code 中工作，则可以使用此文件获取 IntelliSense 并验证您的 JSON。</span><span class="sxs-lookup"><span data-stu-id="fe783-149">If the link is not working, you can find the latest draft of the schema at [draft dynamic-ribbon.schema.json](https://github.com/OfficeDev/testing-assets/tree/master/jsonschema/dynamic-ribbon.schema.json).) If you are working in Visual Studio Code, you can use this file to get IntelliSense and to validate your JSON.</span></span> <span data-ttu-id="fe783-150">有关详细信息，请参阅 [使用 Visual Studio CODE JSON 架构和设置编辑 JSON](https://code.visualstudio.com/docs/languages/json#_json-schemas-and-settings)。</span><span class="sxs-lookup"><span data-stu-id="fe783-150">For more information, see [Editing JSON with Visual Studio Code - JSON schemas and settings](https://code.visualstudio.com/docs/languages/json#_json-schemas-and-settings).</span></span>


1. <span data-ttu-id="fe783-151">首先，创建一个具有两个名为和的数组属性的 JSON 字符串 `actions` `tabs` 。</span><span class="sxs-lookup"><span data-stu-id="fe783-151">Begin by creating a JSON string with two array properties named `actions` and `tabs`.</span></span> <span data-ttu-id="fe783-152">`actions`数组是上下文选项卡上的控件可以执行的所有函数的规范。`tabs`数组定义了一个或多个上下文选项卡，*最多可达 10* 个。</span><span class="sxs-lookup"><span data-stu-id="fe783-152">The `actions` array is a specification of all the functions that can be executed by controls on the contextual tab. The `tabs` array defines one or more contextual tabs, *up to a maximum of 10*.</span></span>

    ```json
    '{
      "actions": [

      ],
      "tabs": [

      ]
    }'
    ```

1. <span data-ttu-id="fe783-153">这一简单的上下文选项卡示例将仅包含一个按钮，因此仅有一个操作。</span><span class="sxs-lookup"><span data-stu-id="fe783-153">This simple example of a contextual tab will have only a single button and, thus, only a single action.</span></span> <span data-ttu-id="fe783-154">将以下项添加为数组中的唯一成员 `actions` 。</span><span class="sxs-lookup"><span data-stu-id="fe783-154">Add the following as the only member of the `actions` array.</span></span> <span data-ttu-id="fe783-155">有关此标记的信息，请注意：</span><span class="sxs-lookup"><span data-stu-id="fe783-155">About this markup, note:</span></span>

    - <span data-ttu-id="fe783-156">`id`和 `type` 属性是必需的。</span><span class="sxs-lookup"><span data-stu-id="fe783-156">The `id` and `type` properties are mandatory.</span></span>
    - <span data-ttu-id="fe783-157">的值 `type` 可以是 "ExecuteFunction" 或 "ShowTaskpane"。</span><span class="sxs-lookup"><span data-stu-id="fe783-157">The value of `type` can be either "ExecuteFunction" or "ShowTaskpane".</span></span>
    - <span data-ttu-id="fe783-158">`functionName`仅当的值为时，才使用 `type` 属性 `ExecuteFunction` 。</span><span class="sxs-lookup"><span data-stu-id="fe783-158">The `functionName` property is only used when the value of `type` is `ExecuteFunction`.</span></span> <span data-ttu-id="fe783-159">它是在 FunctionFile 中定义的函数的名称。</span><span class="sxs-lookup"><span data-stu-id="fe783-159">It is the name of a function defined in the FunctionFile.</span></span> <span data-ttu-id="fe783-160">有关 FunctionFile 的详细信息，请参阅 [外接程序命令的基本概念](add-in-commands.md)。</span><span class="sxs-lookup"><span data-stu-id="fe783-160">For more information about the FunctionFile, see [Basic concepts for Add-in Commands](add-in-commands.md).</span></span>
    - <span data-ttu-id="fe783-161">在后续步骤中，将此操作映射到 "上下文" 选项卡上的一个按钮。</span><span class="sxs-lookup"><span data-stu-id="fe783-161">In a later step, you will map this action to a button on the contextual tab.</span></span>

    ```json
    {
      "id": "executeWriteData",
      "type": "ExecuteFunction",
      "functionName": "writeData"
    }
   ```

1. <span data-ttu-id="fe783-162">将以下项添加为数组中的唯一成员 `tabs` 。</span><span class="sxs-lookup"><span data-stu-id="fe783-162">Add the following as the only member of the `tabs` array.</span></span> <span data-ttu-id="fe783-163">有关此标记的信息，请注意：</span><span class="sxs-lookup"><span data-stu-id="fe783-163">About this markup, note:</span></span>

    - <span data-ttu-id="fe783-164">`id` 属性是必需的。</span><span class="sxs-lookup"><span data-stu-id="fe783-164">The `id` property is required.</span></span> <span data-ttu-id="fe783-165">使用外接程序中的所有上下文选项卡中唯一的简短描述性 ID。</span><span class="sxs-lookup"><span data-stu-id="fe783-165">Use a brief, descriptive ID that is unique among all contextual tabs in your add-in.</span></span>
    - <span data-ttu-id="fe783-166">`label` 属性是必需的。</span><span class="sxs-lookup"><span data-stu-id="fe783-166">The `label` property is required.</span></span> <span data-ttu-id="fe783-167">它是一个用户友好的字符串，用作上下文选项卡的标签。</span><span class="sxs-lookup"><span data-stu-id="fe783-167">It is a user-friendly string to serve as the label of the contextual tab.</span></span>
    - <span data-ttu-id="fe783-168">`groups` 属性是必需的。</span><span class="sxs-lookup"><span data-stu-id="fe783-168">The `groups` property is required.</span></span> <span data-ttu-id="fe783-169">它定义将显示在选项卡上的控件组。它必须至少有一个成员 *，且不能超过20个*。</span><span class="sxs-lookup"><span data-stu-id="fe783-169">It defines the groups of controls that will appear on the tab. It must have at least one member *and no more than 20*.</span></span> <span data-ttu-id="fe783-170"> (还限制了在自定义上下文选项卡上可以拥有的控件数，同时也会限制您拥有的组数。</span><span class="sxs-lookup"><span data-stu-id="fe783-170">(There are also limits on the number of controls that you can have on a custom contextual tab and that will also constrain how many groups that you have.</span></span> <span data-ttu-id="fe783-171">有关详细信息，请参阅下一步。 ) </span><span class="sxs-lookup"><span data-stu-id="fe783-171">See the next step for more information.)</span></span>

    > [!NOTE]
    > <span data-ttu-id="fe783-172">Tab 对象还可以具有一个可选 `visible` 属性，该属性指定在加载项启动时选项卡是否立即可见。</span><span class="sxs-lookup"><span data-stu-id="fe783-172">The tab object can also have an optional `visible` property that specifies whether the tab is visible immediately when the add-in starts up.</span></span> <span data-ttu-id="fe783-173">由于上下文选项卡通常是隐藏的，直到用户事件触发其可见性 (例如，用户在文档中选择某种类型的实体) ，而该 `visible` 属性默认 `false` 情况下不显示时。</span><span class="sxs-lookup"><span data-stu-id="fe783-173">Since contextual tabs are normally hidden until a user event triggers their visibility (such as the user selecting an entity of some type in the document), the `visible` property defaults to `false` when not present.</span></span> <span data-ttu-id="fe783-174">在后面的部分中，我们将演示如何将属性设置为，以 `true` 响应事件。</span><span class="sxs-lookup"><span data-stu-id="fe783-174">In a later section, we show how to set the property to `true` in response to an event.</span></span>

    ```json
    {
      "id": "CtxTab1",
      "label": "Data",
      "groups": [

      ]
    }
    ```

1. <span data-ttu-id="fe783-175">在简单的后续示例中，上下文选项卡仅有一个组。</span><span class="sxs-lookup"><span data-stu-id="fe783-175">In the simple ongoing example, the contextual tab has only a single group.</span></span> <span data-ttu-id="fe783-176">将以下项添加为数组中的唯一成员 `groups` 。</span><span class="sxs-lookup"><span data-stu-id="fe783-176">Add the following as the only member of the `groups` array.</span></span> <span data-ttu-id="fe783-177">有关此标记的信息，请注意：</span><span class="sxs-lookup"><span data-stu-id="fe783-177">About this markup, note:</span></span>

    - <span data-ttu-id="fe783-178">所有属性都是必需的。</span><span class="sxs-lookup"><span data-stu-id="fe783-178">All the properties are required.</span></span>
    - <span data-ttu-id="fe783-179">该 `id` 属性在选项卡中的所有组中必须是唯一的。使用简短的描述性 ID。</span><span class="sxs-lookup"><span data-stu-id="fe783-179">The `id` property must be unique among all the groups in the tab. Use a brief, descriptive ID.</span></span>
    - <span data-ttu-id="fe783-180">`label`是用户友好的字符串，用作组的标签。</span><span class="sxs-lookup"><span data-stu-id="fe783-180">The `label` is a user-friendly string to serve as the label of the group.</span></span>
    - <span data-ttu-id="fe783-181">该 `icon` 属性的值是对象的数组，这些对象指定根据功能区和 Office 应用程序窗口的大小，组将在功能区上所具有的图标。</span><span class="sxs-lookup"><span data-stu-id="fe783-181">The `icon` property's value is an array of objects that specify the icons that the group will have on the ribbon depending on the size of the ribbon and the Office application window.</span></span>
    - <span data-ttu-id="fe783-182">该 `controls` 属性的值是指定组中的按钮和其他控件的对象的数组。</span><span class="sxs-lookup"><span data-stu-id="fe783-182">The `controls` property's value is an array of objects that specify the buttons and other controls in the group.</span></span> <span data-ttu-id="fe783-183">组中必须至少有一个和 *不超过6个*。</span><span class="sxs-lookup"><span data-stu-id="fe783-183">There must be at least one and *no more than 6 in a group*.</span></span>

    > [!IMPORTANT]
    > <span data-ttu-id="fe783-184">*"整个" 选项卡上的总控件数不能超过20。*</span><span class="sxs-lookup"><span data-stu-id="fe783-184">*The total number of controls on the whole tab can be no more than 20.*</span></span> <span data-ttu-id="fe783-185">例如，可以有3个组，每个组具有6个控件，第四组具有2个控件，但您不能有4个组，每个组都有6个控件。</span><span class="sxs-lookup"><span data-stu-id="fe783-185">For example, you could have 3 groups with 6 controls each, and a fourth group with 2 controls, but you cannot have 4 groups with 6 controls each.</span></span>  

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

1. <span data-ttu-id="fe783-186">每个组都必须有至少两个大小的图标： 32x32 px 和 80x80 px。</span><span class="sxs-lookup"><span data-stu-id="fe783-186">Every group must have an icon of at least two sizes, 32x32 px and 80x80 px.</span></span> <span data-ttu-id="fe783-187">（可选）还可以具有16x16、20x20、24x24、40x40、48x48 和64x64 大小的图标。</span><span class="sxs-lookup"><span data-stu-id="fe783-187">Optionally, you can also have icons of sizes 16x16, 20x20, 24x24, 40x40, 48x48 and 64x64.</span></span> <span data-ttu-id="fe783-188">Office 根据功能区和 Office 应用程序窗口的大小决定要使用哪个图标。</span><span class="sxs-lookup"><span data-stu-id="fe783-188">Office decides which icon to use based on the size of the ribbon and Office application window.</span></span> <span data-ttu-id="fe783-189">将以下对象添加到图标数组中。</span><span class="sxs-lookup"><span data-stu-id="fe783-189">Add the following objects to the icon array.</span></span> <span data-ttu-id="fe783-190"> (如果窗口和功能区大小足以满足组中的至少一个 *控件* 的显示，则不会显示任何组图标。</span><span class="sxs-lookup"><span data-stu-id="fe783-190">(If the window and ribbon sizes are large enough for at least one of the *controls* on the group to appear, then no group icon at all appears.</span></span> <span data-ttu-id="fe783-191">有关示例，请查看 Word 功能区上的 **样式** 组，将其缩小并展开 word 窗口。有关此标记的 ) ，请注意：</span><span class="sxs-lookup"><span data-stu-id="fe783-191">For an example, watch the **Styles** group on the Word ribbon as you shrink and expand the Word window.) About this markup, note:</span></span>

    - <span data-ttu-id="fe783-192">这两个属性都是必需的。</span><span class="sxs-lookup"><span data-stu-id="fe783-192">Both the properties are required.</span></span>
    - <span data-ttu-id="fe783-193">该 `size` 属性的度量单位为像素。</span><span class="sxs-lookup"><span data-stu-id="fe783-193">The `size` property unit of measure is pixels.</span></span> <span data-ttu-id="fe783-194">图标始终为方形，因此该数字同时为高度和宽度。</span><span class="sxs-lookup"><span data-stu-id="fe783-194">Icons are always square, so the number is both the height and the width.</span></span>
    - <span data-ttu-id="fe783-195">`sourceLocation`属性指定图标的完整 URL。</span><span class="sxs-lookup"><span data-stu-id="fe783-195">The `sourceLocation` property specifies the full URL to the icon.</span></span>

    > [!IMPORTANT]
    > <span data-ttu-id="fe783-196">在迁移到生产 (时，通常必须更改加载项清单中的 Url，如将域从 localhost 更改为 contoso.com) 中，则还必须更改上下文选项卡 JSON 中的 Url。</span><span class="sxs-lookup"><span data-stu-id="fe783-196">Just as you typically must change the URLs in the add-in's manifest when you move from development to production (such as changing the domain from localhost to contoso.com), you must also change the URLs in your contextual tabs JSON.</span></span>

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

1. <span data-ttu-id="fe783-197">在我们简单的示例中，组仅有一个按钮。</span><span class="sxs-lookup"><span data-stu-id="fe783-197">In our simple ongoing example, the group has only a single button.</span></span> <span data-ttu-id="fe783-198">将以下对象添加为数组的唯一成员 `controls` 。</span><span class="sxs-lookup"><span data-stu-id="fe783-198">Add the following object as the only member of the `controls` array.</span></span> <span data-ttu-id="fe783-199">有关此标记的信息，请注意：</span><span class="sxs-lookup"><span data-stu-id="fe783-199">About this markup, note:</span></span>

    - <span data-ttu-id="fe783-200">除之外的所有属性 `enabled` 都是必需的。</span><span class="sxs-lookup"><span data-stu-id="fe783-200">All the properties, except `enabled`, are required.</span></span>
    - <span data-ttu-id="fe783-201">`type` 指定控件的类型。</span><span class="sxs-lookup"><span data-stu-id="fe783-201">`type` specifies the type of control.</span></span> <span data-ttu-id="fe783-202">这些值可以是 "Button"、"Menu" 或 "MobileButton"。</span><span class="sxs-lookup"><span data-stu-id="fe783-202">The values can be "Button", "Menu", or "MobileButton".</span></span>
    - <span data-ttu-id="fe783-203">`id` 最大可以为125个字符。</span><span class="sxs-lookup"><span data-stu-id="fe783-203">`id` can be up to 125 characters.</span></span> 
    - <span data-ttu-id="fe783-204">`actionId` 必须是在数组中定义的操作的 ID `actions` 。</span><span class="sxs-lookup"><span data-stu-id="fe783-204">`actionId` must be the ID of an action defined in the `actions` array.</span></span> <span data-ttu-id="fe783-205"> (请参阅本部分的步骤1。 ) </span><span class="sxs-lookup"><span data-stu-id="fe783-205">(See step 1 of this section.)</span></span>
    - <span data-ttu-id="fe783-206">`label` 是用户友好的字符串，用作按钮的标签。</span><span class="sxs-lookup"><span data-stu-id="fe783-206">`label` is a user-friendly string to serve as the label of the button.</span></span>
    - <span data-ttu-id="fe783-207">`superTip` 代表一种丰富的工具提示形式。</span><span class="sxs-lookup"><span data-stu-id="fe783-207">`superTip` represents a rich form of tool tip.</span></span> <span data-ttu-id="fe783-208">`title`和 `description` 属性都是必需的。</span><span class="sxs-lookup"><span data-stu-id="fe783-208">Both the `title` and `description` properties are required.</span></span>
    - <span data-ttu-id="fe783-209">`icon` 指定按钮的图标。</span><span class="sxs-lookup"><span data-stu-id="fe783-209">`icon` specifies the icons for the button.</span></span> <span data-ttu-id="fe783-210">上面关于组图标的备注也适用于此处。</span><span class="sxs-lookup"><span data-stu-id="fe783-210">The previous remarks about the group icon apply here too.</span></span>
    - <span data-ttu-id="fe783-211">`enabled` (可选) 指定在启动上下文选项卡时是否启用按钮。</span><span class="sxs-lookup"><span data-stu-id="fe783-211">`enabled` (optional) specifies whether the button is enabled when the contextual tab appears starts up.</span></span> <span data-ttu-id="fe783-212">如果不存在，则为默认值 `true` 。</span><span class="sxs-lookup"><span data-stu-id="fe783-212">The default if not present is `true`.</span></span> 

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
 
<span data-ttu-id="fe783-213">下面是 JSON blob 的完整示例：</span><span class="sxs-lookup"><span data-stu-id="fe783-213">The following is the complete example of the JSON blob:</span></span>

```json
'{
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
      "label": "Data",
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
}'
```

## <a name="register-the-contextual-tab-with-office-with-requestcreatecontrols"></a><span data-ttu-id="fe783-214">使用 requestCreateControls 注册带有 Office 的上下文选项卡</span><span class="sxs-lookup"><span data-stu-id="fe783-214">Register the contextual tab with Office with requestCreateControls</span></span>

<span data-ttu-id="fe783-215">上下文选项卡通过调用 [requestCreateControls](/javascript/api/office/office.ribbon?view=common-js&preserve-view=true#requestCreateControls_tabDefinition_) 方法在 Office 中注册。</span><span class="sxs-lookup"><span data-stu-id="fe783-215">The contextual tab is registered with Office by calling the [Office.ribbon.requestCreateControls](/javascript/api/office/office.ribbon?view=common-js&preserve-view=true#requestCreateControls_tabDefinition_) method.</span></span> <span data-ttu-id="fe783-216">这通常是在分配给或方法的函数中完成的 `Office.initialize` `Office.onReady` 。</span><span class="sxs-lookup"><span data-stu-id="fe783-216">This is typically done in either the function that is assigned to `Office.initialize` or with the `Office.onReady` method.</span></span> <span data-ttu-id="fe783-217">有关这些方法和初始化外接程序的详细信息，请参阅 [初始化 Office 外接程序](../develop/initialize-add-in.md)。</span><span class="sxs-lookup"><span data-stu-id="fe783-217">For more about these methods and initializing the add-in, see [Initialize your Office Add-in](../develop/initialize-add-in.md).</span></span> <span data-ttu-id="fe783-218">不过，您可以在初始化后随时调用方法。</span><span class="sxs-lookup"><span data-stu-id="fe783-218">You can, however, call the method anytime after initialization.</span></span>

> [!IMPORTANT]
> <span data-ttu-id="fe783-219">`requestCreateControls`在外接程序的给定会话中，只能调用一次方法。</span><span class="sxs-lookup"><span data-stu-id="fe783-219">The `requestCreateControls` method can be called only once in a given session of an add-in.</span></span> <span data-ttu-id="fe783-220">如果再次调用，则会引发错误。</span><span class="sxs-lookup"><span data-stu-id="fe783-220">An error is thrown if it is called again.</span></span>

<span data-ttu-id="fe783-221">示例如下。</span><span class="sxs-lookup"><span data-stu-id="fe783-221">The following is an example.</span></span> <span data-ttu-id="fe783-222">请注意，必须使用方法将 JSON 字符串转换为 JavaScript 对象， `JSON.parse` 然后才能将其传递给 javascript 函数。</span><span class="sxs-lookup"><span data-stu-id="fe783-222">Note that the JSON string must be converted to a JavaScript object with the `JSON.parse` method before it can be passed to a JavaScript function.</span></span>

```javascript
Office.onReady(async () => {
    const contextualTabJSON = ' ... '; // Assign the JSON string such as the one at the end of the preceding section.
    const contextualTab = JSON.parse(contextualTabJSON);
    await Office.ribbon.requestCreateControls(contextualTab);
});
```

## <a name="specify-the-contexts-when-the-tab-will-be-visible-with-requestupdate"></a><span data-ttu-id="fe783-223">在选项卡将与 requestUpdate 一起显示时指定上下文</span><span class="sxs-lookup"><span data-stu-id="fe783-223">Specify the contexts when the tab will be visible with requestUpdate</span></span>

<span data-ttu-id="fe783-224">通常情况下，当用户启动的事件更改加载项上下文时，将显示自定义上下文选项卡。</span><span class="sxs-lookup"><span data-stu-id="fe783-224">Typically, a custom contextual tab should appear when a user-initiated event changes the add-in context.</span></span> <span data-ttu-id="fe783-225">考虑在激活 Excel 工作簿) 的默认工作表上的图表 (时，选项卡应可见的情况。</span><span class="sxs-lookup"><span data-stu-id="fe783-225">Consider a scenario in which the tab should be visible when, and only when, a chart (on the default worksheet of an Excel workbook) is activated.</span></span>

<span data-ttu-id="fe783-226">首先分配处理程序。</span><span class="sxs-lookup"><span data-stu-id="fe783-226">Begin by assigning handlers.</span></span> <span data-ttu-id="fe783-227">这通常是在方法中完成的 `Office.onReady` ，如以下示例所示，在后续步骤中 (创建的处理程序分配) 到 `onActivated` `onDeactivated` 工作表中的所有图表的和事件。</span><span class="sxs-lookup"><span data-stu-id="fe783-227">This is commonly done in the `Office.onReady` method as in the following example which assigns handlers (created in a later step) to the `onActivated` and `onDeactivated` events of all the charts in the worksheet.</span></span>

```javascript
Office.onReady(async () => {
    const contextualTabJSON = ' ... '; // Assign the JSON string.
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

<span data-ttu-id="fe783-228">接下来，定义处理程序。</span><span class="sxs-lookup"><span data-stu-id="fe783-228">Next, define the handlers.</span></span> <span data-ttu-id="fe783-229">下面是一个简单的示例 `showDataTab` ，但请参阅本文稍后部分的 [错误处理](#error-handling) ，以获取更强大的函数版本。</span><span class="sxs-lookup"><span data-stu-id="fe783-229">The following is a simple example of a `showDataTab`, but see [Error Handling](#error-handling) later in this article for a more robust version of the function.</span></span> <span data-ttu-id="fe783-230">关于此代码，请注意以下几点：</span><span class="sxs-lookup"><span data-stu-id="fe783-230">About this code, note:</span></span>

- <span data-ttu-id="fe783-231">Office 控制何时更新功能区的状态。</span><span class="sxs-lookup"><span data-stu-id="fe783-231">Office controls when it updates the state of the ribbon.</span></span> <span data-ttu-id="fe783-232">[RequestUpdate](/javascript/api/office/office.ribbon?view=common-js&preserve-view=true#requestupdate-input-)方法对要更新的请求进行排队。</span><span class="sxs-lookup"><span data-stu-id="fe783-232">The  [Office.ribbon.requestUpdate](/javascript/api/office/office.ribbon?view=common-js&preserve-view=true#requestupdate-input-) method queues a request to update.</span></span> <span data-ttu-id="fe783-233">该方法将在 `Promise` 对象排队请求（而不是功能区实际更新）时立即解析该对象。</span><span class="sxs-lookup"><span data-stu-id="fe783-233">The method will resolve the `Promise` object as soon as it has queued the request, not when the ribbon actually updates.</span></span>
- <span data-ttu-id="fe783-234">方法的参数 `requestUpdate` 是一个 [RibbonUpdaterData](/javascript/api/office/office.ribbonupdaterdata)对象，该对象 (1) 按它的 ID 指定选项 *exactly as specified in the JSON* 卡的 (ID。) 指定选项卡的可见性。</span><span class="sxs-lookup"><span data-stu-id="fe783-234">The parameter for the `requestUpdate` method is a [RibbonUpdaterData](/javascript/api/office/office.ribbonupdaterdata) object that (1) specifies the tab by its ID *exactly as specified in the JSON* and (2) specifies visibility of the tab.</span></span>
- <span data-ttu-id="fe783-235">如果有多个自定义上下文选项卡应在相同上下文中可见，则只需向该数组中添加其他选项卡对象 `tabs` 。</span><span class="sxs-lookup"><span data-stu-id="fe783-235">If you have more than one custom contextual tab that should be visible in the same context, you simply add additional tab objects to the `tabs` array.</span></span>

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

<span data-ttu-id="fe783-236">隐藏选项卡的处理程序几乎完全相同，不同之处在于它将 `visible` 属性重新设置为 `false` 。</span><span class="sxs-lookup"><span data-stu-id="fe783-236">The handler to hide the tab is nearly identical, except that it sets the `visible` property back to `false`.</span></span>

<span data-ttu-id="fe783-237">Office JavaScript 库还提供了多个 (类型) 接口，以便更轻松地构造 `RibbonUpdateData` 对象。</span><span class="sxs-lookup"><span data-stu-id="fe783-237">The Office JavaScript library also provides several interfaces (types) to make it easier to construct the`RibbonUpdateData` object.</span></span> <span data-ttu-id="fe783-238">以下是 `showDataTab` TypeScript 中的函数，它使用这些类型。</span><span class="sxs-lookup"><span data-stu-id="fe783-238">The following is the `showDataTab` function in TypeScript and it makes use of these types.</span></span>

```typescript
const showDataTab = async () => {
    const myContextualTab: Tab = {id: "CtxTab1", visible: true};
    const ribbonUpdater: RibbonUpdaterData = { tabs: [ myContextualTab ]};
    await Office.ribbon.requestUpdate(ribbonUpdater);
}
```

### <a name="toggle-tab-visibility-and-the-enabled-status-of-a-button-at-the-same-time"></a><span data-ttu-id="fe783-239">同时切换选项卡可见性和按钮的启用状态</span><span class="sxs-lookup"><span data-stu-id="fe783-239">Toggle tab visibility and the enabled status of a button at the same time</span></span>

<span data-ttu-id="fe783-240">此 `requestUpdate` 方法还用于切换自定义上下文选项卡或自定义 "核心" 选项卡上的自定义按钮的启用或禁用状态。有关此内容的详细信息，请参阅 [Enable And Disable 外接程序命令](disable-add-in-commands.md)。</span><span class="sxs-lookup"><span data-stu-id="fe783-240">The `requestUpdate` method is also used to toggle the enabled or disabled status of a custom button on either a custom contextual tab or a custom core tab. For details about this, see [Enable and Disable Add-in Commands](disable-add-in-commands.md).</span></span> <span data-ttu-id="fe783-241">在某些情况下，您可能希望同时更改选项卡的可见性和按钮的启用状态。</span><span class="sxs-lookup"><span data-stu-id="fe783-241">There may be scenarios in which you want to change both the visibility of a tab and the enabled status of a button at the same time.</span></span> <span data-ttu-id="fe783-242">您可以通过一次调用来执行此操作 `requestUpdate` 。</span><span class="sxs-lookup"><span data-stu-id="fe783-242">You can do this with a single call of `requestUpdate`.</span></span> <span data-ttu-id="fe783-243">下面的示例演示了启用 "核心" 选项卡上的按钮时，将显示上下文选项卡。</span><span class="sxs-lookup"><span data-stu-id="fe783-243">The following is an example in which a button on a core tab is enabled at the same time as a contextual tab is made visible.</span></span>

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
                controls: [
                {
                    id: "MyButton",
                    enabled: true
                }
            ]}
        ]});
}
```

<span data-ttu-id="fe783-244">在下面的示例中，启用的按钮在相同的上下文选项卡上，使其可见。</span><span class="sxs-lookup"><span data-stu-id="fe783-244">In the following example, the button that is enabled is on the very same contextual tab that is being made visible.</span></span>

```javascript
function myContextChanges() {
    Office.ribbon.requestUpdate({
        tabs: [
            {
                id: "CtxTab1",
                visible: true,
                controls: [
                    {
                        id: "MyButton",
                        enabled: true
                    }
                ]
            }
        ]});
}
```

## <a name="error-handling"></a><span data-ttu-id="fe783-245">错误处理</span><span class="sxs-lookup"><span data-stu-id="fe783-245">Error handling</span></span>

<span data-ttu-id="fe783-246">在某些情况下，Office 无法更新功能区，并将返回错误。</span><span class="sxs-lookup"><span data-stu-id="fe783-246">In some scenarios, Office is unable to update the ribbon and will return an error.</span></span> <span data-ttu-id="fe783-247">例如，如果升级了加载项，并且升级后的加载项具有一组不同的自定义加载项命令，则必须关闭并重新打开 Office 应用程序。</span><span class="sxs-lookup"><span data-stu-id="fe783-247">For example, if the add-in is upgraded and the upgraded add-in has a different set of custom add-in commands, then the Office application must be closed and reopened.</span></span> <span data-ttu-id="fe783-248">在此之前，`requestUpdate` 方法将返回错误 `HostRestartNeeded`。</span><span class="sxs-lookup"><span data-stu-id="fe783-248">Until it is, the `requestUpdate` method will return the error `HostRestartNeeded`.</span></span> <span data-ttu-id="fe783-249">以下是如何处理此错误的示例。</span><span class="sxs-lookup"><span data-stu-id="fe783-249">The following is an example of how to handle this error.</span></span> <span data-ttu-id="fe783-250">在此示例中，`reportError` 方法向用户显示错误。</span><span class="sxs-lookup"><span data-stu-id="fe783-250">In this case, the `reportError` method displays the error to the user.</span></span>

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
