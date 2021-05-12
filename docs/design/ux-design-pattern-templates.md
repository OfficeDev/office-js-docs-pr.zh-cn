---
title: 适用于 Office 外接程序的 UX 设计模式
description: 大致了解加载项的 UI 设计Office，包括导航、身份验证、首次运行和品牌打造模式。
ms.date: 05/12/2021
localization_priority: Normal
ms.openlocfilehash: 8544b56b85a25d522c95546b42a78fe01a3c2586
ms.sourcegitcommit: 30f6c620380075e3459cac748ca0c656427b384d
ms.translationtype: MT
ms.contentlocale: zh-CN
ms.lasthandoff: 05/12/2021
ms.locfileid: "52330106"
---
# <a name="ux-design-patterns-for-office-add-ins"></a><span data-ttu-id="b93f9-103">适用于 Office 外接程序的 UX 设计模式</span><span class="sxs-lookup"><span data-stu-id="b93f9-103">UX design patterns for Office Add-ins</span></span>

<span data-ttu-id="b93f9-104">在设计 Office 外接程序的用户体验时，应为 Office 用户提供具有吸引力的体验并通过在默认 Office UI 内无缝接入来扩展整体 Office 体验。</span><span class="sxs-lookup"><span data-stu-id="b93f9-104">Designing the user experience for Office Add-ins should provide a compelling experience for Office users and extend the overall Office experience by fitting seamlessly within the default Office UI.</span></span>  

<span data-ttu-id="b93f9-105">我们的 UX 模式由组件组成。</span><span class="sxs-lookup"><span data-stu-id="b93f9-105">Our UX patterns are composed of components.</span></span> <span data-ttu-id="b93f9-106">组件是帮助客户与软件或服务元素进行交互的控件。</span><span class="sxs-lookup"><span data-stu-id="b93f9-106">Components are controls that help your customers interact with elements of your software or service.</span></span> <span data-ttu-id="b93f9-107">按钮、导航、和菜单是常见组件的示例，通常具有一致的样式和行为。</span><span class="sxs-lookup"><span data-stu-id="b93f9-107">Buttons, navigation, and menus are examples of common components that often have consistent styles and behaviors.</span></span>

<span data-ttu-id="b93f9-108">[Fluent UI React组件](using-office-ui-fabric-react.md)的外观和行为与 Office 的一部分类似，Office UI Fabric [JS 的中性框架组件一样](fabric-core.md)。</span><span class="sxs-lookup"><span data-stu-id="b93f9-108">[Fluent UI React components](using-office-ui-fabric-react.md) look and behave like a part of Office, as do the framework-neutral components of [Office UI Fabric JS](fabric-core.md).</span></span> <span data-ttu-id="b93f9-109">利用任一组组件与Office。</span><span class="sxs-lookup"><span data-stu-id="b93f9-109">Take advantage of either set of components to integrate with Office.</span></span> <span data-ttu-id="b93f9-110">或者，如果您的外接程序具有自己的预先不存在的组件语言，则无需放弃它。</span><span class="sxs-lookup"><span data-stu-id="b93f9-110">Alternatively, if your add-in has its own preexisting component language, you don't need to discard it.</span></span> <span data-ttu-id="b93f9-111">与 Office 集成的同时寻找保留该语言的机会。</span><span class="sxs-lookup"><span data-stu-id="b93f9-111">Look for opportunities to retain it while integrating with Office.</span></span> <span data-ttu-id="b93f9-112">想办法改变风格元素、消除冲突或采用可避免用户混淆的样式和行为。</span><span class="sxs-lookup"><span data-stu-id="b93f9-112">Consider ways to swap out stylistic elements, remove conflicts, or adopt styles and behaviors that remove user confusion.</span></span>

<span data-ttu-id="b93f9-113">提供的模式是基于常见客户方案和用户体验研究的最佳做法解决方案。</span><span class="sxs-lookup"><span data-stu-id="b93f9-113">The provided patterns are best practice solutions based on common customer scenarios and user experience research.</span></span> <span data-ttu-id="b93f9-114">它们旨在提供设计和开发外接程序的快速入口点，以及指导在 Microsoft 品牌元素和您自己的元素之间实现平衡。</span><span class="sxs-lookup"><span data-stu-id="b93f9-114">They are meant to provide both a quick entry point to designing and developing add-ins as well as guidance to achieve balance between Microsoft brand elements and your own.</span></span> <span data-ttu-id="b93f9-115">提供干净、现代用户体验，在 Microsoft Fluent UI 设计语言中的设计元素与合作伙伴的独特品牌标识之间实现平衡，这可帮助提高外接程序的用户保留率和采用率。</span><span class="sxs-lookup"><span data-stu-id="b93f9-115">Providing a clean, modern user experience that balances design elements from Microsoft's Fluent UI design language and the partner's unique brand identity may help increase user retention and adoption of your add-in.</span></span>

<span data-ttu-id="b93f9-116">使用 UX 模式模板来实现以下目的：</span><span class="sxs-lookup"><span data-stu-id="b93f9-116">Use the UX pattern templates to:</span></span>

* <span data-ttu-id="b93f9-117">将解决方案应用于常见的客户方案。</span><span class="sxs-lookup"><span data-stu-id="b93f9-117">Apply solutions to common customer scenarios.</span></span>
* <span data-ttu-id="b93f9-118">应用设计最佳实践。</span><span class="sxs-lookup"><span data-stu-id="b93f9-118">Apply design best practices.</span></span>
* <span data-ttu-id="b93f9-119">合并 [Fluent UI](https://developer.microsoft.com/fluentui#/get-started) 组件和样式。</span><span class="sxs-lookup"><span data-stu-id="b93f9-119">Incorporate [Fluent UI](https://developer.microsoft.com/fluentui#/get-started) components and styles.</span></span>
* <span data-ttu-id="b93f9-120">构建以可视方式与默认 Office UI 集成的外接程序。</span><span class="sxs-lookup"><span data-stu-id="b93f9-120">Build add-ins that visually integrate with the default Office UI.</span></span>
* <span data-ttu-id="b93f9-121">形成 UX 概念并将其可视化。</span><span class="sxs-lookup"><span data-stu-id="b93f9-121">Ideate and visualize UX.</span></span>

## <a name="getting-started"></a><span data-ttu-id="b93f9-122">入门</span><span class="sxs-lookup"><span data-stu-id="b93f9-122">Getting started</span></span>

<span data-ttu-id="b93f9-123">该模式按照外接程序中的常见按键操作或体验来进行组织。</span><span class="sxs-lookup"><span data-stu-id="b93f9-123">The patterns are organized by key actions or experiences that are common in an add-in.</span></span> <span data-ttu-id="b93f9-124">主要的组包括：</span><span class="sxs-lookup"><span data-stu-id="b93f9-124">The main groups are:</span></span>

* [<span data-ttu-id="b93f9-125">初次运行体验 (FRE)</span><span class="sxs-lookup"><span data-stu-id="b93f9-125">First run experience (FRE)</span></span>](../design/first-run-experience-patterns.md)
* [<span data-ttu-id="b93f9-126">身份验证</span><span class="sxs-lookup"><span data-stu-id="b93f9-126">Authentication</span></span>](../design/authentication-patterns.md)
* [<span data-ttu-id="b93f9-127">导航</span><span class="sxs-lookup"><span data-stu-id="b93f9-127">Navigation</span></span>](../design/navigation-patterns.md)
* [<span data-ttu-id="b93f9-128">品牌设计</span><span class="sxs-lookup"><span data-stu-id="b93f9-128">Branding Design</span></span>](../design/branding-patterns.md)

<span data-ttu-id="b93f9-129">浏览每个分组，了解如何使用最佳做法来设计外接程序。</span><span class="sxs-lookup"><span data-stu-id="b93f9-129">Browse each grouping to get an idea of how you can design your add-in using best practices.</span></span>

> [!NOTE]
> <span data-ttu-id="b93f9-130">本文档中显示的所有示例屏幕均按 **1366x768** 的分辨率进行设计和显示。</span><span class="sxs-lookup"><span data-stu-id="b93f9-130">The example screens shown throughout this documentation are designed and displayed at a resolution of **1366x768**.</span></span>

## <a name="see-also"></a><span data-ttu-id="b93f9-131">另请参阅</span><span class="sxs-lookup"><span data-stu-id="b93f9-131">See also</span></span>

* [<span data-ttu-id="b93f9-132">设计工具包</span><span class="sxs-lookup"><span data-stu-id="b93f9-132">Design tool kits</span></span>](design-toolkits.md)
* [<span data-ttu-id="b93f9-133">Fluent UI</span><span class="sxs-lookup"><span data-stu-id="b93f9-133">Fluent UI</span></span>](https://developer.microsoft.com/fluentui#)
* [<span data-ttu-id="b93f9-134">Office 加载项开发最佳做法</span><span class="sxs-lookup"><span data-stu-id="b93f9-134">Best practices for developing Office Add-ins</span></span>](../concepts/add-in-development-best-practices.md)
* [<span data-ttu-id="b93f9-135">Fluent UI React外接程序Office中的用户界面</span><span class="sxs-lookup"><span data-stu-id="b93f9-135">Fluent UI React in Office Add-ins</span></span>](using-office-ui-fabric-react.md)
