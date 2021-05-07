---
title: Office 加载项中的 Office UI Fabric
description: 大致了解如何在加载项Office UI Fabric加载项Office组件。
ms.date: 05/03/2021
localization_priority: Normal
ms.openlocfilehash: 20f926913335197a65ac24e4ec30ed0106b81bae
ms.sourcegitcommit: 8fbc7c7eb47875bf022e402b13858695a8536ec5
ms.translationtype: MT
ms.contentlocale: zh-CN
ms.lasthandoff: 05/06/2021
ms.locfileid: "52253366"
---
# <a name="office-ui-fabric-in-office-add-ins"></a><span data-ttu-id="854b0-103">Office 加载项中的 Office UI Fabric</span><span class="sxs-lookup"><span data-stu-id="854b0-103">Office UI Fabric in Office Add-ins</span></span>

<span data-ttu-id="854b0-104">Office UI Fabric是一个 JavaScript 前端框架，用于生成适用于Office。</span><span class="sxs-lookup"><span data-stu-id="854b0-104">Office UI Fabric is a JavaScript front-end framework for building user experiences for Office.</span></span> <span data-ttu-id="854b0-105">Fabric 提供了以视觉对象为中心的组件，可在 Office 外接程序中进行扩展、返工和使用。</span><span class="sxs-lookup"><span data-stu-id="854b0-105">Fabric provides visuals-focused components that you can extend, rework, and use in your Office Add-in.</span></span> <span data-ttu-id="854b0-106">由于 Fabric 使用的是 Office 设计语言，因此 Fabric 的用户体验组件看起来像是 Office 的自然扩展。</span><span class="sxs-lookup"><span data-stu-id="854b0-106">Because Fabric uses the Office Design Language, Fabric's UX components look like a natural extension of Office.</span></span>

<span data-ttu-id="854b0-p102">若要生成外接程序，我们建议使用 Office UI Fabric 生成用户体验。使用 Office UI Fabric 是可选的。</span><span class="sxs-lookup"><span data-stu-id="854b0-p102">If you are building an add-in, we encourage you to use Office UI Fabric to create your user experience. Using Office UI Fabric is optional.</span></span>

<span data-ttu-id="854b0-109">以下各节介绍如何开始使用 Fabric 以满足要求。</span><span class="sxs-lookup"><span data-stu-id="854b0-109">The following sections explain how to get started using Fabric to meet your requirements.</span></span>

## <a name="use-fabric-core-icons-fonts-colors"></a><span data-ttu-id="854b0-110">使用 Fabric Core：图标、字体、颜色</span><span class="sxs-lookup"><span data-stu-id="854b0-110">Use Fabric Core: icons, fonts, colors</span></span>

<span data-ttu-id="854b0-111">Fabric Core 包含设计语言的基本元素，如图标、颜色、类型和网格等。</span><span class="sxs-lookup"><span data-stu-id="854b0-111">Fabric Core contains basic elements of the design language such as icons, colors, type, and grid.</span></span> <span data-ttu-id="854b0-112"> Fabric Core 与框架无关。</span><span class="sxs-lookup"><span data-stu-id="854b0-112">Fabric core is framework independent.</span></span> <span data-ttu-id="854b0-113">Fabric Core 供 Fabric React 使用并且包含其中。</span><span class="sxs-lookup"><span data-stu-id="854b0-113">Fabric Core is used by and included with Fabric React.</span></span>

<span data-ttu-id="854b0-114">开始使用 Fabric Core：</span><span class="sxs-lookup"><span data-stu-id="854b0-114">To get started using Fabric Core:</span></span>

1. <span data-ttu-id="854b0-115">向页面上的 HTML 添加 CDN 参考。</span><span class="sxs-lookup"><span data-stu-id="854b0-115">Add the CDN reference to the HTML on your page.</span></span>  

    ```html
    <link rel="stylesheet" href="https://static2.sharepointonline.com/files/fabric/office-ui-fabric-core/9.6.1/css/fabric.min.css">
    ```

2. <span data-ttu-id="854b0-116">使用 Fabric 图标和字体。</span><span class="sxs-lookup"><span data-stu-id="854b0-116">Use Fabric icons and fonts.</span></span>

    <span data-ttu-id="854b0-p104">若要使用 Fabric 图标，在页面上包括“i”元素，然后引用适当的类。可以通过更改字号来控制图标的大小。例如，下面的代码展示了如何制作使用 themePrimary (#0078d7) 颜色的超大表图标。</span><span class="sxs-lookup"><span data-stu-id="854b0-p104">To use a Fabric icon, include the "i" element on your page, and then reference the appropriate classes. You can control the size of the icon by changing the font size. For example, the following code shows how to make an extra-large table icon that uses the themePrimary (#0078d7) color.</span></span>

    ```html
    <i class="ms-Icon ms-font-xl ms-Icon--Table ms-fontColor-themePrimary"></i>
    ```

    <span data-ttu-id="854b0-p105">若要查找 Office UI Fabric 中可用的更多图标，请在“[图标](https://developer.microsoft.com/fabric#/styles/icons)”页上使用搜索功能。找到要在外接程序中使用的图标后，请务必在图标名称前加上前缀 `ms-Icon--`。</span><span class="sxs-lookup"><span data-stu-id="854b0-p105">To find more icons that are available in Office UI Fabric, use the search feature on the [Icons](https://developer.microsoft.com/fabric#/styles/icons) page. When you find an icon to use in your add-in, be sure to prefix the icon name with `ms-Icon--`.</span></span>

    <span data-ttu-id="854b0-122">若要了解 Office UI Fabric 中可用的字号和颜色，请参阅[版式](https://developer.microsoft.com/fabric#/styles/typography)和[颜色](https://developer.microsoft.com/fabric#/styles/colors)。</span><span class="sxs-lookup"><span data-stu-id="854b0-122">For information about font sizes and colors that are available in Office UI Fabric, see [Typography](https://developer.microsoft.com/fabric#/styles/typography) and [Colors](https://developer.microsoft.com/fabric#/styles/colors).</span></span>

## <a name="use-fabric-components"></a><span data-ttu-id="854b0-123">使用 Fabric 组件</span><span class="sxs-lookup"><span data-stu-id="854b0-123">Use Fabric Components</span></span>

<span data-ttu-id="854b0-124">Fabric 提供了各种可用于生成外接程序的 UX 组件。</span><span class="sxs-lookup"><span data-stu-id="854b0-124">Fabric provides a variety of UX components that you can use to build your add-in.</span></span> <span data-ttu-id="854b0-125">我们预计单个外接程序不会使用所有结构组件。</span><span class="sxs-lookup"><span data-stu-id="854b0-125">We do not expect that all fabric components will be used by a single add-in.</span></span> <span data-ttu-id="854b0-126">确定适用于您的方案和用户体验的最佳组件 (例如，可能很难在任务窗格中正确显示痕迹导航) 。 [](https://developer.microsoft.com/fabric#/components/breadcrumb)</span><span class="sxs-lookup"><span data-stu-id="854b0-126">Determine the best components for your scenario and user experience (for example, it may be hard to properly display a [Breadcrumb](https://developer.microsoft.com/fabric#/components/breadcrumb) in the task pane).</span></span>

<span data-ttu-id="854b0-127">以下是我们建议在外接程序React Fabric 和[UX](https://developer.microsoft.com/fluentui#/controls/web)组件的常见列表：</span><span class="sxs-lookup"><span data-stu-id="854b0-127">The following is a list of common [Fabric React UX components](https://developer.microsoft.com/fluentui#/controls/web) that we recommend for use in an add-in:</span></span>

- [<span data-ttu-id="854b0-128">按钮</span><span class="sxs-lookup"><span data-stu-id="854b0-128">Button</span></span>](https://developer.microsoft.com/fabric#/components/button)
- [<span data-ttu-id="854b0-129">复选框</span><span class="sxs-lookup"><span data-stu-id="854b0-129">Checkbox</span></span>](https://developer.microsoft.com/fabric#/components/checkbox)
- [<span data-ttu-id="854b0-130">ChoiceGroup</span><span class="sxs-lookup"><span data-stu-id="854b0-130">ChoiceGroup</span></span>](https://developer.microsoft.com/fabric#/components/choicegroup)
- [<span data-ttu-id="854b0-131">下拉列表</span><span class="sxs-lookup"><span data-stu-id="854b0-131">Dropdown</span></span>](https://developer.microsoft.com/fabric#/components/dropdown)
- [<span data-ttu-id="854b0-132">标签</span><span class="sxs-lookup"><span data-stu-id="854b0-132">Label</span></span>](https://developer.microsoft.com/fabric#/components/label)
- [<span data-ttu-id="854b0-133">列表</span><span class="sxs-lookup"><span data-stu-id="854b0-133">List</span></span>](https://developer.microsoft.com/fabric#/components/list)
- [<span data-ttu-id="854b0-134">透视</span><span class="sxs-lookup"><span data-stu-id="854b0-134">Pivot</span></span>](https://developer.microsoft.com/fabric#/components/pivot)
- [<span data-ttu-id="854b0-135">TextField</span><span class="sxs-lookup"><span data-stu-id="854b0-135">TextField</span></span>](https://developer.microsoft.com/fabric#/components/textfield)
- [<span data-ttu-id="854b0-136">切换</span><span class="sxs-lookup"><span data-stu-id="854b0-136">Toggle</span></span>](https://developer.microsoft.com/fabric#/components/toggle)

<span data-ttu-id="854b0-p107">你可以使用不同的 JavaScript 框架（如 Angular 或 React）来生成外接程序。若要开始将 Fabric 组件与框架一起使用，请参阅以下资源。</span><span class="sxs-lookup"><span data-stu-id="854b0-p107">You can use different JavaScript frameworks, such as Angular or React, to build your add-in. To get started using Fabric components with your framework, see the following resources.</span></span>

|<span data-ttu-id="854b0-139">**框架**</span><span class="sxs-lookup"><span data-stu-id="854b0-139">**Framework**</span></span>|<span data-ttu-id="854b0-140">**示例**</span><span class="sxs-lookup"><span data-stu-id="854b0-140">**Example**</span></span>|
|:------------|:----------|
|<span data-ttu-id="854b0-141">**React**</span><span class="sxs-lookup"><span data-stu-id="854b0-141">**React**</span></span>|[<span data-ttu-id="854b0-142">在 Office 外接程序中使用 Office UI Fabric React</span><span class="sxs-lookup"><span data-stu-id="854b0-142">Using Office UI Fabric React in Office Add-ins</span></span>](using-office-ui-fabric-react.md )|
