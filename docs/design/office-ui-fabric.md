---
title: Office 加载项中的 Office UI Fabric
description: 概述如何在 Office 外接程序中使用 Office UI Fabric 组件。
ms.date: 10/29/2020
localization_priority: Normal
ms.openlocfilehash: c4a13c615fe63183f595e24895b9fe6054fdc05d
ms.sourcegitcommit: ca66ff7462bfdf4ed7ae04f43d1388c24de63bf9
ms.translationtype: MT
ms.contentlocale: zh-CN
ms.lasthandoff: 11/11/2020
ms.locfileid: "48996373"
---
# <a name="office-ui-fabric-in-office-add-ins"></a><span data-ttu-id="83595-103">Office 加载项中的 Office UI Fabric</span><span class="sxs-lookup"><span data-stu-id="83595-103">Office UI Fabric in Office Add-ins</span></span>

<span data-ttu-id="83595-p101">Office UI Fabric 是用于生成 Office 和 Office 365 用户体验的 JavaScript 前端框架。Fabric 提供了以视觉对象为中心的组件，可在 Office 外接程序中进行扩展、返工和使用。由于 Fabric 使用的是 Office 设计语言，因此 Fabric 的用户体验组件看起来像是 Office 的自然扩展。</span><span class="sxs-lookup"><span data-stu-id="83595-p101">Office UI Fabric is a JavaScript front-end framework for building user experiences for Office and Office 365. Fabric provides visuals-focused components that you can extend, rework, and use in your Office Add-in. Because Fabric uses the Office Design Language, Fabric's UX components look like a natural extension of Office.</span></span>

<span data-ttu-id="83595-p102">若要生成外接程序，我们建议使用 Office UI Fabric 生成用户体验。使用 Office UI Fabric 是可选的。</span><span class="sxs-lookup"><span data-stu-id="83595-p102">If you are building an add-in, we encourage you to use Office UI Fabric to create your user experience. Using Office UI Fabric is optional.</span></span>

<span data-ttu-id="83595-109">以下各节介绍如何开始使用 Fabric 以满足要求。</span><span class="sxs-lookup"><span data-stu-id="83595-109">The following sections explain how to get started using Fabric to meet your requirements.</span></span>

## <a name="use-fabric-core-icons-fonts-colors"></a><span data-ttu-id="83595-110">使用 Fabric Core：图标、字体、颜色</span><span class="sxs-lookup"><span data-stu-id="83595-110">Use Fabric Core: icons, fonts, colors</span></span>

<span data-ttu-id="83595-111">Fabric Core 包含设计语言的基本元素，如图标、颜色、类型和网格等。</span><span class="sxs-lookup"><span data-stu-id="83595-111">Fabric Core contains basic elements of the design language such as icons, colors, type, and grid.</span></span> <span data-ttu-id="83595-112"> Fabric Core 与框架无关。</span><span class="sxs-lookup"><span data-stu-id="83595-112">Fabric core is framework independent.</span></span> <span data-ttu-id="83595-113">Fabric Core 供 Fabric React 使用并且包含其中。</span><span class="sxs-lookup"><span data-stu-id="83595-113">Fabric Core is used by and included with Fabric React.</span></span>

<span data-ttu-id="83595-114">开始使用 Fabric Core：</span><span class="sxs-lookup"><span data-stu-id="83595-114">To get started using Fabric Core:</span></span>

1. <span data-ttu-id="83595-115">向页面上的 HTML 添加 CDN 参考。</span><span class="sxs-lookup"><span data-stu-id="83595-115">Add the CDN reference to the HTML on your page.</span></span>  

    ```html
    <link rel="stylesheet" href="https://static2.sharepointonline.com/files/fabric/office-ui-fabric-core/9.6.1/css/fabric.min.css">
    ```

2. <span data-ttu-id="83595-116">使用 Fabric 图标和字体。</span><span class="sxs-lookup"><span data-stu-id="83595-116">Use Fabric icons and fonts.</span></span>

    <span data-ttu-id="83595-p104">若要使用 Fabric 图标，在页面上包括“i”元素，然后引用适当的类。可以通过更改字号来控制图标的大小。例如，下面的代码展示了如何制作使用 themePrimary (#0078d7) 颜色的超大表图标。</span><span class="sxs-lookup"><span data-stu-id="83595-p104">To use a Fabric icon, include the "i" element on your page, and then reference the appropriate classes. You can control the size of the icon by changing the font size. For example, the following code shows how to make an extra-large table icon that uses the themePrimary (#0078d7) color.</span></span>

    ```html
    <i class="ms-Icon ms-font-xl ms-Icon--Table ms-fontColor-themePrimary"></i>
    ```

    <span data-ttu-id="83595-p105">若要查找 Office UI Fabric 中可用的更多图标，请在“[图标](https://developer.microsoft.com/fabric#/styles/icons)”页上使用搜索功能。找到要在外接程序中使用的图标后，请务必在图标名称前加上前缀 `ms-Icon--`。</span><span class="sxs-lookup"><span data-stu-id="83595-p105">To find more icons that are available in Office UI Fabric, use the search feature on the [Icons](https://developer.microsoft.com/fabric#/styles/icons) page. When you find an icon to use in your add-in, be sure to prefix the icon name with `ms-Icon--`.</span></span>

    <span data-ttu-id="83595-122">若要了解 Office UI Fabric 中可用的字号和颜色，请参阅[版式](https://developer.microsoft.com/fabric#/styles/typography)和[颜色](https://developer.microsoft.com/fabric#/styles/colors)。</span><span class="sxs-lookup"><span data-stu-id="83595-122">For information about font sizes and colors that are available in Office UI Fabric, see [Typography](https://developer.microsoft.com/fabric#/styles/typography) and [Colors](https://developer.microsoft.com/fabric#/styles/colors).</span></span>

## <a name="use-fabric-components"></a><span data-ttu-id="83595-123">使用 Fabric 组件</span><span class="sxs-lookup"><span data-stu-id="83595-123">Use Fabric Components</span></span>

<span data-ttu-id="83595-124">Fabric 提供了多种 UX 组件，可用于生成外接程序。</span><span class="sxs-lookup"><span data-stu-id="83595-124">Fabric provides a variety of UX components that you can use to build your add-in.</span></span> <span data-ttu-id="83595-125">我们不希望所有 fabric 组件都将由单个外接程序使用。</span><span class="sxs-lookup"><span data-stu-id="83595-125">We do not expect that all fabric components will be used by a single add-in.</span></span> <span data-ttu-id="83595-126">确定适用于您的方案和用户体验的最佳组件 (例如，可能很难在任务窗格) 中正确显示 [痕迹导航](https://developer.microsoft.com/fabric#/components/breadcrumb) 。</span><span class="sxs-lookup"><span data-stu-id="83595-126">Determine the best components for your scenario and user experience (for example, it may be hard to properly display a [Breadcrumb](https://developer.microsoft.com/fabric#/components/breadcrumb) in the task pane).</span></span>

<span data-ttu-id="83595-127">以下是我们建议用于外接程序的常见 [Fabric 响应 UX 组件](https://developer.microsoft.com/fluentui#/controls/web) 的列表：</span><span class="sxs-lookup"><span data-stu-id="83595-127">The following is a list of common [Fabric React UX components](https://developer.microsoft.com/fluentui#/controls/web) that we recommend for use in an add-in:</span></span>

- [<span data-ttu-id="83595-128">按钮</span><span class="sxs-lookup"><span data-stu-id="83595-128">Button</span></span>](https://developer.microsoft.com/fabric#/components/button)
- [<span data-ttu-id="83595-129">复选框</span><span class="sxs-lookup"><span data-stu-id="83595-129">Checkbox</span></span>](https://developer.microsoft.com/fabric#/components/checkbox)
- [<span data-ttu-id="83595-130">ChoiceGroup</span><span class="sxs-lookup"><span data-stu-id="83595-130">ChoiceGroup</span></span>](https://developer.microsoft.com/fabric#/components/choicegroup)
- [<span data-ttu-id="83595-131">下拉列表</span><span class="sxs-lookup"><span data-stu-id="83595-131">Dropdown</span></span>](https://developer.microsoft.com/fabric#/components/dropdown)
- [<span data-ttu-id="83595-132">标签</span><span class="sxs-lookup"><span data-stu-id="83595-132">Label</span></span>](https://developer.microsoft.com/fabric#/components/label)
- [<span data-ttu-id="83595-133">列表</span><span class="sxs-lookup"><span data-stu-id="83595-133">List</span></span>](https://developer.microsoft.com/fabric#/components/list)
- [<span data-ttu-id="83595-134">透视</span><span class="sxs-lookup"><span data-stu-id="83595-134">Pivot</span></span>](https://developer.microsoft.com/fabric#/components/pivot)
- [<span data-ttu-id="83595-135">TextField</span><span class="sxs-lookup"><span data-stu-id="83595-135">TextField</span></span>](https://developer.microsoft.com/fabric#/components/textfield)
- [<span data-ttu-id="83595-136">切换</span><span class="sxs-lookup"><span data-stu-id="83595-136">Toggle</span></span>](https://developer.microsoft.com/fabric#/components/toggle)

<span data-ttu-id="83595-p107">你可以使用不同的 JavaScript 框架（如 Angular 或 React）来生成外接程序。若要开始将 Fabric 组件与框架一起使用，请参阅以下资源。</span><span class="sxs-lookup"><span data-stu-id="83595-p107">You can use different JavaScript frameworks, such as Angular or React, to build your add-in. To get started using Fabric components with your framework, see the following resources.</span></span>

|<span data-ttu-id="83595-139">**框架**</span><span class="sxs-lookup"><span data-stu-id="83595-139">**Framework**</span></span>|<span data-ttu-id="83595-140">**示例**</span><span class="sxs-lookup"><span data-stu-id="83595-140">**Example**</span></span>|
|:------------|:----------|
|<span data-ttu-id="83595-141">**React**</span><span class="sxs-lookup"><span data-stu-id="83595-141">**React**</span></span>|[<span data-ttu-id="83595-142">在 Office 外接程序中使用 Office UI Fabric React</span><span class="sxs-lookup"><span data-stu-id="83595-142">Using Office UI Fabric React in Office Add-ins</span></span>](using-office-ui-fabric-react.md )|
|<span data-ttu-id="83595-143">**Angular**</span><span class="sxs-lookup"><span data-stu-id="83595-143">**Angular**</span></span>| [<span data-ttu-id="83595-144">考虑使用角2组件包装 Fabric 组件</span><span class="sxs-lookup"><span data-stu-id="83595-144">Consider wrapping Fabric components with Angular 2 components</span></span>](../develop/add-ins-with-angular2.md#consider-wrapping-fabric-components-with-angular-components)|
