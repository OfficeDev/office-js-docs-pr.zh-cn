---
title: Office 加载项中的 Office UI Fabric
description: ''
ms.date: 12/04/2017
ms.openlocfilehash: b573f720ebe4f90f7d4dbfdb05693871b93a2258
ms.sourcegitcommit: 4de2a1b62ccaa8e51982e95537fc9f52c0c5e687
ms.translationtype: HT
ms.contentlocale: zh-CN
ms.lasthandoff: 08/10/2018
ms.locfileid: "22925190"
---
# <a name="office-ui-fabric-in-office-add-ins"></a><span data-ttu-id="50baa-102">Office 加载项中的 Office UI Fabric</span><span class="sxs-lookup"><span data-stu-id="50baa-102">Office UI Fabric in Office Add-ins</span></span> 

<span data-ttu-id="50baa-p101">Office UI Fabric 是用于生成 Office 和 Office 365 用户体验的 JavaScript 前端框架。Fabric 提供了以视觉对象为中心的组件，可在 Office 外接程序中进行扩展、返工和使用。由于 Fabric 使用的是 Office 设计语言，因此 Fabric 的用户体验组件看起来像是 Office 的自然扩展。</span><span class="sxs-lookup"><span data-stu-id="50baa-p101">Office UI Fabric is a JavaScript front-end framework for building user experiences for Office and Office 365. Fabric provides visuals-focused components that you can extend, rework, and use in your Office Add-in. Because Fabric uses the Office Design Language, Fabric's UX components look like a natural extension of Office.</span></span> 

<span data-ttu-id="50baa-p102">若要生成外接程序，我们建议使用 Office UI Fabric 生成用户体验。使用 Office UI Fabric 是可选的。</span><span class="sxs-lookup"><span data-stu-id="50baa-p102">If you are building an add-in, we encourage you to use Office UI Fabric to create your user experience. Using Office UI Fabric is optional.</span></span>

<span data-ttu-id="50baa-108">以下各节介绍如何开始使用 Fabric 以满足要求。</span><span class="sxs-lookup"><span data-stu-id="50baa-108">The following sections explain how to get started using Fabric to meet your requirements.</span></span> 

## <a name="use-fabric-core-icons-fonts-colors"></a><span data-ttu-id="50baa-109">使用 Fabric Core：图标、字体、颜色</span><span class="sxs-lookup"><span data-stu-id="50baa-109">Use Fabric Core: icons, fonts, colors</span></span>
<span data-ttu-id="50baa-p103">Fabric Core 包含设计语言的基本元素，如图标、颜色、类型和网格等。Fabric core 与框架无关。Fabric React 和 Fabric JS 都使用 Fabric Core。</span><span class="sxs-lookup"><span data-stu-id="50baa-p103">Fabric Core contains basic elements of the design language such as icons, colors, type, and grid. Fabric core is framework independent. Both Fabric React and Fabric JS use Fabric Core.</span></span>

<span data-ttu-id="50baa-113">开始使用 Fabric Core：</span><span class="sxs-lookup"><span data-stu-id="50baa-113">To get started using Fabric Core:</span></span>

1. <span data-ttu-id="50baa-114">向页面上的 HTML 添加 CDN 参考。</span><span class="sxs-lookup"><span data-stu-id="50baa-114">Add the CDN reference to the HTML on your page.</span></span>  

    ```html
    <link rel="stylesheet" href="https://static2.sharepointonline.com/files/fabric/office-ui-fabric-js/1.4.0/css/fabric.min.css">
    ```   
    
2. <span data-ttu-id="50baa-115">使用 Fabric 图标和字体。</span><span class="sxs-lookup"><span data-stu-id="50baa-115">Use Fabric icons and fonts.</span></span> 

    <span data-ttu-id="50baa-p104">若要使用 Fabric 图标，在页面上包括“i”元素，然后引用适当的类。可以通过更改字号来控制图标的大小。例如，下面的代码展示了如何制作使用 themePrimary (#0078d7) 颜色的超大表图标。</span><span class="sxs-lookup"><span data-stu-id="50baa-p104">To use a Fabric icon, include the "i" element on your page, and then reference the appropriate classes. You can control the size of the icon by changing the font size. For example, the following code shows how to make an extra-large table icon that uses the themePrimary (#0078d7) color.</span></span> 
   
    ```html
    <i class="ms-Icon ms-font-xl ms-Icon--Table ms-fontColor-themePrimary"></i>
    ```

    <span data-ttu-id="50baa-p105">若要查找 Office UI Fabric 中可用的更多图标，请在“[图标](https://dev.office.com/fabric#/styles/icons)”页上使用搜索功能。找到要在外接程序中使用的图标后，请务必在图标名称前加上前缀 `ms-Icon--`。</span><span class="sxs-lookup"><span data-stu-id="50baa-p105">To find more icons that are available in Office UI Fabric, use the search feature on the [Icons](https://dev.office.com/fabric#/styles/icons) page. When you find an icon to use in your add-in, be sure to prefix the icon name with `ms-Icon--`.</span></span> 

    <span data-ttu-id="50baa-121">若要了解 Office UI Fabric 中可用的字号和颜色，请参阅[版式](https://dev.office.com/fabric#/styles/typography)和[颜色](https://dev.office.com/fabric#/styles/colors)。</span><span class="sxs-lookup"><span data-stu-id="50baa-121">For information about font sizes and colors that are available in Office UI Fabric, see [Typography](https://dev.office.com/fabric#/styles/typography) and [Colors](https://dev.office.com/fabric#/styles/colors).</span></span>
 
## <a name="use-fabric-components"></a><span data-ttu-id="50baa-122">使用 Fabric 组件</span><span class="sxs-lookup"><span data-stu-id="50baa-122">Use Fabric Components</span></span> 
<span data-ttu-id="50baa-123">Fabric 提供了多种可用于生成外界程序的 UX 组件，包括以下类型的组件：</span><span class="sxs-lookup"><span data-stu-id="50baa-123">Fabric provides a variety of UX components that you can use to build your add-in, including the following types of components:</span></span>

- <span data-ttu-id="50baa-124">输入组件 - 如按钮、复选框和切换</span><span class="sxs-lookup"><span data-stu-id="50baa-124">Input components - for example, Button, Checkbox, and Toggle</span></span>
- <span data-ttu-id="50baa-125">导航组件 - 如透视、痕迹</span><span class="sxs-lookup"><span data-stu-id="50baa-125">Navigation components - for example, Pivot Breadcrumb</span></span>
- <span data-ttu-id="50baa-126">通知组件 - 例如，消息栏和标注</span><span class="sxs-lookup"><span data-stu-id="50baa-126">Notification components - for example, MessageBar and Callout</span></span>  

<span data-ttu-id="50baa-127">并非所有 Fabric 组件都推荐用于外接程序。以下是我们建议在外接程序中使用的 Fabric React UX 组件列表：</span><span class="sxs-lookup"><span data-stu-id="50baa-127">Not all Fabric components are recommended for use in add-ins. Here is a list of Fabric React UX components that we recommend for use in an add-in:</span></span>

- [<span data-ttu-id="50baa-128">痕迹导航</span><span class="sxs-lookup"><span data-stu-id="50baa-128">Breadcrumb</span></span>](https://developer.microsoft.com/fabric#/components/breadcrumb)
- [<span data-ttu-id="50baa-129">按钮</span><span class="sxs-lookup"><span data-stu-id="50baa-129">Button</span></span>](https://developer.microsoft.com/fabric#/components/button)
- [<span data-ttu-id="50baa-130">复选框</span><span class="sxs-lookup"><span data-stu-id="50baa-130">Checkbox</span></span>](https://developer.microsoft.com/fabric#/components/checkbox)
- [<span data-ttu-id="50baa-131">选择组</span><span class="sxs-lookup"><span data-stu-id="50baa-131">ChoiceGroup</span></span>](https://developer.microsoft.com/fabric#/components/choicegroup)
- [<span data-ttu-id="50baa-132">下拉列表</span><span class="sxs-lookup"><span data-stu-id="50baa-132">Dropdown</span></span>](https://developer.microsoft.com/fabric#/components/dropdown)
- [<span data-ttu-id="50baa-133">标签</span><span class="sxs-lookup"><span data-stu-id="50baa-133">Label</span></span>](https://developer.microsoft.com/fabric#/components/label)
- [<span data-ttu-id="50baa-134">列表</span><span class="sxs-lookup"><span data-stu-id="50baa-134">List</span></span>](https://developer.microsoft.com/fabric#/components/list)
- [<span data-ttu-id="50baa-135">透视</span><span class="sxs-lookup"><span data-stu-id="50baa-135">Pivot</span></span>](https://developer.microsoft.com/fabric#/components/pivot)
- [<span data-ttu-id="50baa-136">文本字段</span><span class="sxs-lookup"><span data-stu-id="50baa-136">TextField</span></span>](https://developer.microsoft.com/fabric#/components/textfield)
- [<span data-ttu-id="50baa-137">切换</span><span class="sxs-lookup"><span data-stu-id="50baa-137">Toggle</span></span>](https://developer.microsoft.com/fabric#/components/toggle)

<span data-ttu-id="50baa-p106">你可以使用不同的 JavaScript 框架（如 Angular 或 React）来生成外接程序。若要开始将 Fabric 组件与框架一起使用，请参阅以下资源。</span><span class="sxs-lookup"><span data-stu-id="50baa-p106">You can use different JavaScript frameworks, such as Angular or React, to build your add-in. To get started using Fabric components with your framework, see the following resources.</span></span>

|<span data-ttu-id="50baa-140">**框架**</span><span class="sxs-lookup"><span data-stu-id="50baa-140">**Framework**</span></span>|<span data-ttu-id="50baa-141">**示例**</span><span class="sxs-lookup"><span data-stu-id="50baa-141">**Example**</span></span>|
|:------------|:----------|
|<span data-ttu-id="50baa-142">**React**</span><span class="sxs-lookup"><span data-stu-id="50baa-142">**React**</span></span>|[<span data-ttu-id="50baa-143">在 Office 外接程序中使用 Office UI Fabric React</span><span class="sxs-lookup"><span data-stu-id="50baa-143">Using Office UI Fabric React in Office Add-ins</span></span>](using-office-ui-fabric-react.md )|
|<span data-ttu-id="50baa-144">**Angular**</span><span class="sxs-lookup"><span data-stu-id="50baa-144">**Angular**</span></span>| <span data-ttu-id="50baa-145">请参阅包含 Angular 1.5 指令的社区项目 [ngOfficeUIFabric](http://ngofficeuifabric.com/)，以及[考虑使用 Angular 2 组件包装 Fabric 组件](../develop/add-ins-with-angular2.md#consider-wrapping-fabric-components-with-angular-components)</span><span class="sxs-lookup"><span data-stu-id="50baa-145">See [ngOfficeUIFabric](http://ngofficeuifabric.com/), which is a community project with Angular 1.5 directives, and [Consider wrapping Fabric components with Angular 2 components](../develop/add-ins-with-angular2.md#consider-wrapping-fabric-components-with-angular-components)</span></span>|
