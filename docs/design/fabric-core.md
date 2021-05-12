---
title: Office外接程序中的 Fabric Core
description: 大致了解如何在加载项中Office Fabric Core 和 Fabric UI 组件。
ms.date: 05/12/2021
localization_priority: Normal
ms.openlocfilehash: e93efaea55841cc3bb6fa79ea1d1bbcaa76a4d05
ms.sourcegitcommit: 30f6c620380075e3459cac748ca0c656427b384d
ms.translationtype: MT
ms.contentlocale: zh-CN
ms.lasthandoff: 05/12/2021
ms.locfileid: "52330198"
---
# <a name="fabric-core-in-office-add-ins"></a><span data-ttu-id="ab643-103">Office外接程序中的 Fabric Core</span><span class="sxs-lookup"><span data-stu-id="ab643-103">Fabric Core in Office Add-ins</span></span>

<span data-ttu-id="ab643-104">Fabric Core 是 CSS 类和 SASS mixin 的开源集合，旨在用于非 React *Office* 外接程序。Fabric Core 包含 Fluent UI 设计语言的基本元素，如图标、颜色、字样和网格。</span><span class="sxs-lookup"><span data-stu-id="ab643-104">Fabric Core is an open-source collection of CSS classes and SASS mixins that's *intended for use in non-React* Office Add-ins. Fabric Core contains basic elements of the Fluent UI design language such as icons, colors, typefaces, and grids.</span></span> <span data-ttu-id="ab643-105">Fabric Core 与框架无关，因此可用于任何单页应用程序或任何服务器端 Web UI 框架。</span><span class="sxs-lookup"><span data-stu-id="ab643-105">Fabric Core is framework independent, so it can be used with any single-page application or any server-side web UI framework.</span></span> <span data-ttu-id="ab643-106"> (出于历史原因，它被称为"Fabric Core"，而不是"Fluent Core"。) </span><span class="sxs-lookup"><span data-stu-id="ab643-106">(It's called "Fabric Core" instead of "Fluent Core" for historical reasons.)</span></span>

<span data-ttu-id="ab643-107">如果外接程序的 UI 不是基于React的，则您还可以使用一组非React组件。</span><span class="sxs-lookup"><span data-stu-id="ab643-107">If your add-in's UI is not React-based, you can also make use of a set of non-React components.</span></span> <span data-ttu-id="ab643-108">请参阅[使用Office UI Fabric JS 组件](#use-office-ui-fabric-js-components)。</span><span class="sxs-lookup"><span data-stu-id="ab643-108">See [Use Office UI Fabric JS components](#use-office-ui-fabric-js-components).</span></span>

> [!NOTE]
> <span data-ttu-id="ab643-109">本文介绍了 Fabric Core 在加载项Office的使用。但它还用于各种应用Microsoft 365扩展。</span><span class="sxs-lookup"><span data-stu-id="ab643-109">This article describes the use of Fabric Core in the context of Office Add-ins. But it's also used in a wide range of Microsoft 365 apps and extensions.</span></span> <span data-ttu-id="ab643-110">有关详细信息，请参阅[Fabric Core](https://developer.microsoft.com/fluentui#/get-started/web#fabric-core)和开源存储库Office UI Fabric [Core。](https://github.com/OfficeDev/office-ui-fabric-core)</span><span class="sxs-lookup"><span data-stu-id="ab643-110">For more information, see [Fabric Core](https://developer.microsoft.com/fluentui#/get-started/web#fabric-core) and the open source repo [Office UI Fabric Core](https://github.com/OfficeDev/office-ui-fabric-core).</span></span>

## <a name="use-fabric-core-icons-fonts-colors"></a><span data-ttu-id="ab643-111">使用 Fabric Core：图标、字体、颜色</span><span class="sxs-lookup"><span data-stu-id="ab643-111">Use Fabric Core: icons, fonts, colors</span></span>

<span data-ttu-id="ab643-112">开始使用 Fabric Core：</span><span class="sxs-lookup"><span data-stu-id="ab643-112">To get started using Fabric Core:</span></span>

1. <span data-ttu-id="ab643-113">向页面上的 HTML 添加 CDN 参考。</span><span class="sxs-lookup"><span data-stu-id="ab643-113">Add the CDN reference to the HTML on your page.</span></span>  

    ```html
    <link rel="stylesheet" href="https://static2.sharepointonline.com/files/fabric/office-ui-fabric-core/9.6.1/css/fabric.min.css">
    ```

2. <span data-ttu-id="ab643-114">使用 Fabric Core 图标和字体。</span><span class="sxs-lookup"><span data-stu-id="ab643-114">Use Fabric Core icons and fonts.</span></span>

    <span data-ttu-id="ab643-115">若要使用 Fabric Core 图标，请在你的页面上包括"i"元素，然后引用相应的类。</span><span class="sxs-lookup"><span data-stu-id="ab643-115">To use a Fabric Core icon, include the "i" element on your page, and then reference the appropriate classes.</span></span> <span data-ttu-id="ab643-116">可以通过更改字号来控制图标的大小。</span><span class="sxs-lookup"><span data-stu-id="ab643-116">You can control the size of the icon by changing the font size.</span></span> <span data-ttu-id="ab643-117">例如，下面的代码展示了如何制作使用 themePrimary (#0078d7) 颜色的超大表图标。</span><span class="sxs-lookup"><span data-stu-id="ab643-117">For example, the following code shows how to make an extra-large table icon that uses the themePrimary (#0078d7) color.</span></span>

    ```html
    <i class="ms-Icon ms-font-xl ms-Icon--Table ms-fontColor-themePrimary"></i>
    ```

    <span data-ttu-id="ab643-118">有关更详细的说明，请参阅 [Fluent UI 图标](https://developer.microsoft.com/fluentui#/styles/web/icons)。</span><span class="sxs-lookup"><span data-stu-id="ab643-118">For more detailed instructions, see [Fluent UI Icons](https://developer.microsoft.com/fluentui#/styles/web/icons).</span></span> <span data-ttu-id="ab643-119">若要查找 Fabric Core 中可用的更多图标，请使用该页面上的搜索功能。</span><span class="sxs-lookup"><span data-stu-id="ab643-119">To find more icons that are available in Fabric Core, use the search feature on that page.</span></span> <span data-ttu-id="ab643-120">找到要在外接程序中使用的图标后，请务必在图标名称前加上前缀 `ms-Icon--`。</span><span class="sxs-lookup"><span data-stu-id="ab643-120">When you find an icon to use in your add-in, be sure to prefix the icon name with `ms-Icon--`.</span></span>

    <span data-ttu-id="ab643-121">有关 Fabric Core 中可用的字体大小和颜色的信息，请参阅 Colors[](https://developer.microsoft.com/fluentui#/styles/web/typography)中的版式和颜色[目录](https://developer.microsoft.com/fluentui#/styles/web/colors)。</span><span class="sxs-lookup"><span data-stu-id="ab643-121">For information about font sizes and colors that are available in Fabric Core, see [Typography](https://developer.microsoft.com/fluentui#/styles/web/typography) and the **Colors** table of contents at [Colors](https://developer.microsoft.com/fluentui#/styles/web/colors).</span></span>

<span data-ttu-id="ab643-122">示例包含在本文稍后 [的示例中](#samples) 。</span><span class="sxs-lookup"><span data-stu-id="ab643-122">Examples are included in the [Samples](#samples) later in this article.</span></span>

## <a name="use-office-ui-fabric-js-components"></a><span data-ttu-id="ab643-123">使用 Office UI Fabric JS 组件</span><span class="sxs-lookup"><span data-stu-id="ab643-123">Use Office UI Fabric JS components</span></span>

<span data-ttu-id="ab643-124">具有非 REACT API 的外接程序还可使用[Office UI Fabric JS 中的](https://github.com/OfficeDev/office-ui-fabric-js)任意组件，包括按钮、对话框、选取器等。</span><span class="sxs-lookup"><span data-stu-id="ab643-124">Add-ins with non-React UIs can also use any of the many components from [Office UI Fabric JS](https://github.com/OfficeDev/office-ui-fabric-js), including buttons, dialogs, pickers, and more.</span></span> <span data-ttu-id="ab643-125">有关说明，请参阅存储库自述。</span><span class="sxs-lookup"><span data-stu-id="ab643-125">See the readme of the repo for instructions.</span></span>

<span data-ttu-id="ab643-126">示例包含在本文稍后 [的示例中](#samples) 。</span><span class="sxs-lookup"><span data-stu-id="ab643-126">Examples are included in the [Samples](#samples) later in this article.</span></span>

## <a name="samples"></a><span data-ttu-id="ab643-127">示例</span><span class="sxs-lookup"><span data-stu-id="ab643-127">Samples</span></span>

<span data-ttu-id="ab643-128">以下示例外接程序使用 Fabric Core 和/或 Office UI Fabric JS 组件。</span><span class="sxs-lookup"><span data-stu-id="ab643-128">The following sample add-ins use Fabric Core and/or Office UI Fabric JS components.</span></span> <span data-ttu-id="ab643-129">其中一些资源已存档，这意味着不再使用 Bug 或安全修补程序更新它们，但你仍可以使用它们了解如何使用 Fabric Core 和 Fabric UI 组件。</span><span class="sxs-lookup"><span data-stu-id="ab643-129">Some of these repos are archived, meaning that they are no longer being updated with bug or security fixes, but you can still use them to learn how to use Fabric Core and Fabric UI components.</span></span>

- [<span data-ttu-id="ab643-130">Excel外接程序 JavaScript SalesTracker</span><span class="sxs-lookup"><span data-stu-id="ab643-130">Excel Add-in JavaScript SalesTracker</span></span>](https://github.com/OfficeDev/Excel-Add-in-JavaScript-SalesTracker)
- [<span data-ttu-id="ab643-131">Excel外接程序 SalesLeads</span><span class="sxs-lookup"><span data-stu-id="ab643-131">Excel Add-in SalesLeads</span></span>](https://github.com/OfficeDev/Excel-Add-in-SalesLeads)
- [<span data-ttu-id="ab643-132">Excel加载项 WoodGrove 费用趋势</span><span class="sxs-lookup"><span data-stu-id="ab643-132">Excel Add-in WoodGrove Expense Trends</span></span>](https://github.com/OfficeDev/Excel-Add-in-WoodGrove-Expense-Trends)
- [<span data-ttu-id="ab643-133">Excel内容外接程序 Humongous Insurance</span><span class="sxs-lookup"><span data-stu-id="ab643-133">Excel Content Add-in Humongous Insurance</span></span>](https://github.com/OfficeDev/Excel-Content-Add-in-Humongous-Insurance)
- [<span data-ttu-id="ab643-134">Office外接程序 Fabric UI 示例</span><span class="sxs-lookup"><span data-stu-id="ab643-134">Office Add-in Fabric UI Sample</span></span>](https://github.com/OfficeDev/Office-Add-in-Fabric-UI-Sample)
- [<span data-ttu-id="ab643-135">Office-Add-in-UX-Design-Patterns-Code</span><span class="sxs-lookup"><span data-stu-id="ab643-135">Office-Add-in-UX-Design-Patterns-Code</span></span>](https://github.com/OfficeDev/Office-Add-in-UX-Design-Patterns-Code)
- [<span data-ttu-id="ab643-136">Outlook外接程序 GifMe</span><span class="sxs-lookup"><span data-stu-id="ab643-136">Outlook Add-in GifMe</span></span>](https://github.com/OfficeDev/Outlook-Add-in-GifMe)
- [<span data-ttu-id="ab643-137">PowerPoint外接程序 Microsoft Graph ASPNET 插入图</span><span class="sxs-lookup"><span data-stu-id="ab643-137">PowerPoint Add-in Microsoft Graph ASPNET InsertChart</span></span>](https://github.com/OfficeDev/PowerPoint-Add-in-Microsoft-Graph-ASPNET-InsertChart)
- [<span data-ttu-id="ab643-138">Word 外接程序 Angular2 StyleChecker</span><span class="sxs-lookup"><span data-stu-id="ab643-138">Word Add-in Angular2 StyleChecker</span></span>](https://github.com/OfficeDev/Word-Add-in-Angular2-StyleChecker)
- [<span data-ttu-id="ab643-139">Word 外接程序 JS 修订</span><span class="sxs-lookup"><span data-stu-id="ab643-139">Word Add-in JS Redact</span></span>](https://github.com/OfficeDev/Word-Add-in-JS-Redact)
- [<span data-ttu-id="ab643-140">Word 加载项 MarkdownConversion</span><span class="sxs-lookup"><span data-stu-id="ab643-140">Word Add-in MarkdownConversion</span></span>](https://github.com/OfficeDev/Word-Add-in-MarkdownConversion)
