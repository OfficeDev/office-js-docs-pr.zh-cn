---
title: 适用于 Office 外接程序的 UX 设计模式
description: ''
ms.date: 06/27/2018
ms.openlocfilehash: 635fc27d18a2c671dd1ac5a521c9d0a920c154ed
ms.sourcegitcommit: 6f53df6f3ee91e084cd5160bb48afbbd49743b7e
ms.translationtype: HT
ms.contentlocale: zh-CN
ms.lasthandoff: 12/22/2018
ms.locfileid: "27432471"
---
# <a name="ux-design-patterns-for-office-add-ins"></a><span data-ttu-id="0f5fa-102">适用于 Office 外接程序的 UX 设计模式</span><span class="sxs-lookup"><span data-stu-id="0f5fa-102">UX design patterns for Office Add-ins</span></span>

<span data-ttu-id="0f5fa-103">在设计 Office 外接程序的用户体验时，应为 Office 用户提供具有吸引力的体验并通过在默认 Office UI 内无缝接入来扩展整体 Office 体验。</span><span class="sxs-lookup"><span data-stu-id="0f5fa-103">Designing the user experience for Office Add-ins should provide a compelling experience for Office users and extend the overall Office experience by fitting seamlessly within the default Office UI.</span></span>  

<span data-ttu-id="0f5fa-104">我们的 UX 模式由组件组成。</span><span class="sxs-lookup"><span data-stu-id="0f5fa-104">Our UX patterns are composed of components.</span></span> <span data-ttu-id="0f5fa-105">组件是帮助客户与软件或服务元素进行交互的控件。</span><span class="sxs-lookup"><span data-stu-id="0f5fa-105">Components are controls that help your customers interact with elements of your software or service.</span></span> <span data-ttu-id="0f5fa-106">按钮、导航、和菜单是常见组件的示例，通常具有一致的样式和行为。</span><span class="sxs-lookup"><span data-stu-id="0f5fa-106">Buttons, navigation, and menus are examples of common components that often have consistent styles and behaviors.</span></span>

<span data-ttu-id="0f5fa-107">Office UI Fabric 呈现外观和行为类似于 Office 部件的组件。</span><span class="sxs-lookup"><span data-stu-id="0f5fa-107">Office UI Fabric renders components that look and behave like a part of Office.</span></span> <span data-ttu-id="0f5fa-108">利用 Fabric 来轻松与 Office 集成。</span><span class="sxs-lookup"><span data-stu-id="0f5fa-108">Take advantage of Fabric to easily integrate with Office.</span></span> <span data-ttu-id="0f5fa-109">如果外接程序有其自己预先存在的组件语言，则不需要为支持 Fabric 而放弃它。</span><span class="sxs-lookup"><span data-stu-id="0f5fa-109">If your add-in has its own preexisting component language, you don't need to discard it in favor of Fabric.</span></span> <span data-ttu-id="0f5fa-110">与 Office 集成的同时寻找保留该语言的机会。</span><span class="sxs-lookup"><span data-stu-id="0f5fa-110">Look for opportunities to retain it while integrating with Office.</span></span> <span data-ttu-id="0f5fa-111">想办法改变风格元素、消除冲突或采用可避免用户混淆的样式和行为。</span><span class="sxs-lookup"><span data-stu-id="0f5fa-111">Consider ways to swap out stylistic elements, remove conflicts, or adopt styles and behaviors that remove user confusion.</span></span>

<span data-ttu-id="0f5fa-112">提供的模式是基于常见客户方案和用户体验研究的最佳做法解决方案。</span><span class="sxs-lookup"><span data-stu-id="0f5fa-112">The provided patterns are best practice solutions based on common customer scenarios and user experience research.</span></span> <span data-ttu-id="0f5fa-113">它们旨在提供设计和开发外接程序的快速切入点，以及提供在 Microsoft 和品牌元素之间实现平衡的指导。</span><span class="sxs-lookup"><span data-stu-id="0f5fa-113">They are meant to provide both a quick entry point to designing and developing add-ins as well as guidance to achieve balance between Microsoft and brand elements.</span></span> <span data-ttu-id="0f5fa-114">提供整洁的新式用户体验，并在 Microsoft Fabric 设计语言的设计元素与合作伙伴的独特品牌标识之间保持平衡，可能有助于提高外接程序的用户保留率和采用率。</span><span class="sxs-lookup"><span data-stu-id="0f5fa-114">Providing a clean, modern user experience that balances design elements from Microsoft's Fabric design language and the partner's unique brand identity may help increase user retention and adoption of your add-in.</span></span>

<span data-ttu-id="0f5fa-115">使用 UX 模式模板来实现以下目的：</span><span class="sxs-lookup"><span data-stu-id="0f5fa-115">Use the UX pattern templates to:</span></span>

* <span data-ttu-id="0f5fa-116">将解决方案应用于常见的客户方案。</span><span class="sxs-lookup"><span data-stu-id="0f5fa-116">Apply solutions to common customer scenarios.</span></span>
* <span data-ttu-id="0f5fa-117">应用设计最佳实践。</span><span class="sxs-lookup"><span data-stu-id="0f5fa-117">Apply design best practices.</span></span>
* <span data-ttu-id="0f5fa-118">纳入“[Office UI Fabric](https://developer.microsoft.com/fabric#/get-started)”组件和样式。</span><span class="sxs-lookup"><span data-stu-id="0f5fa-118">Incorporate [Office UI Fabric](https://developer.microsoft.com/fabric#/get-started) components and styles.</span></span>
* <span data-ttu-id="0f5fa-119">构建以可视方式与默认 Office UI 集成的外接程序。</span><span class="sxs-lookup"><span data-stu-id="0f5fa-119">Build add-ins that visually integrate with the default Office UI.</span></span>
* <span data-ttu-id="0f5fa-120">构想 UX 并将其可视化。</span><span class="sxs-lookup"><span data-stu-id="0f5fa-120">Ideate and visualize UX.</span></span>


## <a name="getting-started"></a><span data-ttu-id="0f5fa-121">入门</span><span class="sxs-lookup"><span data-stu-id="0f5fa-121">Getting started</span></span>

<span data-ttu-id="0f5fa-122">模式按照外接程序中的常见按键操作或体验来进行组织。</span><span class="sxs-lookup"><span data-stu-id="0f5fa-122">The patterns are organized by key actions or experiences that are common in an add-in.</span></span> <span data-ttu-id="0f5fa-123">主要的组包括：</span><span class="sxs-lookup"><span data-stu-id="0f5fa-123">The main groups are:</span></span>

* [<span data-ttu-id="0f5fa-124">初次运行体验 (FRE)</span><span class="sxs-lookup"><span data-stu-id="0f5fa-124">First run experience</span></span>](../design/first-run-experience-patterns.md)
* [<span data-ttu-id="0f5fa-125">身份验证</span><span class="sxs-lookup"><span data-stu-id="0f5fa-125">Authentication</span></span>](../design/authentication-patterns.md)
* [<span data-ttu-id="0f5fa-126">导航</span><span class="sxs-lookup"><span data-stu-id="0f5fa-126">Navigation</span></span>](../design/navigation-patterns.md)
* [<span data-ttu-id="0f5fa-127">品牌设计</span><span class="sxs-lookup"><span data-stu-id="0f5fa-127">Branding Design</span></span>](../design/branding-patterns.md)

<span data-ttu-id="0f5fa-128">浏览每个分组，了解如何使用最佳做法来设计外接程序。</span><span class="sxs-lookup"><span data-stu-id="0f5fa-128">Browse each grouping to get an idea of how you can design your add-in using best practices.</span></span>



><span data-ttu-id="0f5fa-129">注意：本文档中显示的所有示例屏幕均是以 **1366x768** 的分辨率来设计和显示的</span><span class="sxs-lookup"><span data-stu-id="0f5fa-129">NOTE: The example screens shown throughout this documentation are designed and displayed at a resolution of **1366x768**</span></span>




## <a name="see-also"></a><span data-ttu-id="0f5fa-130">另请参阅</span><span class="sxs-lookup"><span data-stu-id="0f5fa-130">See also</span></span>
* [<span data-ttu-id="0f5fa-131">设计工具包</span><span class="sxs-lookup"><span data-stu-id="0f5fa-131">Design toolkits</span></span>](design-toolkits.md)
* [<span data-ttu-id="0f5fa-132">Office UI Fabric</span><span class="sxs-lookup"><span data-stu-id="0f5fa-132">Office UI Fabric</span></span>](https://developer.microsoft.com/fabric)
* [<span data-ttu-id="0f5fa-133">开发 Office 外接程序的最佳做法</span><span class="sxs-lookup"><span data-stu-id="0f5fa-133">Best practices for developing Office Add-ins</span></span>](https://docs.microsoft.com/office/dev/add-ins/concepts/add-in-development-best-practices)
* [<span data-ttu-id="0f5fa-134">Fabric React 使用入门</span><span class="sxs-lookup"><span data-stu-id="0f5fa-134">Get started using Fabric React</span></span>](https://docs.microsoft.com/office/dev/add-ins/design/using-office-ui-fabric-react)
