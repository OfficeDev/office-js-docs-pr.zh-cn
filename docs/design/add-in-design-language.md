---
title: Office 加载项设计语言
description: 了解如何使加载项Office与加载项Office。
ms.date: 05/12/2021
localization_priority: Normal
ms.openlocfilehash: 4374eaad1e660728a347d19a012d345b642755af
ms.sourcegitcommit: 30f6c620380075e3459cac748ca0c656427b384d
ms.translationtype: MT
ms.contentlocale: zh-CN
ms.lasthandoff: 05/12/2021
ms.locfileid: "52330057"
---
# <a name="office-add-in-design-language"></a><span data-ttu-id="4a62f-103">Office 加载项设计语言</span><span class="sxs-lookup"><span data-stu-id="4a62f-103">Office Add-in design language</span></span>

<span data-ttu-id="4a62f-p101">Office 设计语言是一种简单明了的视觉对象系统，它可确保体验的一致性。它包含一组用于定义 Office 接口的可视化元素，包括：</span><span class="sxs-lookup"><span data-stu-id="4a62f-p101">The Office design language is a clean and simple visual system that ensures consistency across experiences. It contains a set of visual elements that define Office interfaces, including:</span></span>

- <span data-ttu-id="4a62f-106">标准字样</span><span class="sxs-lookup"><span data-stu-id="4a62f-106">A standard typeface</span></span>
- <span data-ttu-id="4a62f-107">公用调色板</span><span class="sxs-lookup"><span data-stu-id="4a62f-107">A common color palette</span></span>
- <span data-ttu-id="4a62f-108">一组版式大小和权重</span><span class="sxs-lookup"><span data-stu-id="4a62f-108">A set of typographic sizes and weights</span></span>
- <span data-ttu-id="4a62f-109">图标准则</span><span class="sxs-lookup"><span data-stu-id="4a62f-109">Icon guidelines</span></span>
- <span data-ttu-id="4a62f-110">共享图标资源</span><span class="sxs-lookup"><span data-stu-id="4a62f-110">Shared icon assets</span></span>
- <span data-ttu-id="4a62f-111">动画定义</span><span class="sxs-lookup"><span data-stu-id="4a62f-111">Animation definitions</span></span>
- <span data-ttu-id="4a62f-112">常见组件</span><span class="sxs-lookup"><span data-stu-id="4a62f-112">Common components</span></span>

<span data-ttu-id="4a62f-113">[Fluent UI](../design/add-in-design.md)是使用自定义设计语言构建的官方Office框架。</span><span class="sxs-lookup"><span data-stu-id="4a62f-113">[Fluent UI](../design/add-in-design.md) is the official front-end framework for building with the Office design language.</span></span> <span data-ttu-id="4a62f-114">可以选择使用 Fluent UI，但这是确保加载项感觉像加载项自然扩展的最快Office。</span><span class="sxs-lookup"><span data-stu-id="4a62f-114">Using Fluent UI is optional, but it is the fastest way to ensure that your add-ins feel like a natural extension of Office.</span></span> <span data-ttu-id="4a62f-115">利用 Fluent UI 来设计和构建对应用程序进行补充的Office。</span><span class="sxs-lookup"><span data-stu-id="4a62f-115">Take advantage of Fluent UI to design and build add-ins that complement Office.</span></span>

<span data-ttu-id="4a62f-p103">许多 Office 外接程序与先前存在的品牌相关联。你可以保留外接程序中的强大品牌及其视觉对象或组件语言。与 Office 集成的同时寻找保留自己的视觉对象语言的机会。寻找方法将 Office 颜色、版式、图标或其他样式元素置换为你自己品牌的元素。在插入客户熟悉的控件和组件时，寻找遵循通用外接程序布局或 UX 设计模式的方法。</span><span class="sxs-lookup"><span data-stu-id="4a62f-p103">Many Office Add-ins are associated with a preexisting brand. You can retain a strong brand and its visual or component language in your add-in. Look for opportunities to retain your own visual language while integrating with Office. Consider ways to swap out Office colors, typography, icons, or other stylistic elements with elements of your own brand. Consider ways to follow common add-in layouts or UX design patterns while inserting controls and components that are familiar to your customers.</span></span>

<span data-ttu-id="4a62f-p104">在 Office 内插入基于主要品牌的 HTML 的 UI 会对客户产生不一致性。找到一个能够在 Office 中无缝整合的平衡点，同时与你的服务或父品牌保持明确一致。如果外接程序不适合 Office，通常是因为样式元素发生冲突。例如，版式过大和网格关闭、颜色对比度鲜明或太过强烈，或者相比 Office 动画过多且行为有差异。控件或组件的外观和行为与 Office 标准相差甚远。</span><span class="sxs-lookup"><span data-stu-id="4a62f-p104">Inserting a heavily branded HTML-based UI inside of Office can create dissonance for customers. Find a balance that fits seamlessly in Office but also clearly aligns with your service or parent brand. When an add-in does not fit with Office, it's often because stylistic elements conflict. For example, typography is too large and off grid, colors are contrasting or particularly loud, or animations are superfluous and behave differently than Office. The appearance and behavior of controls or components veer too far from Office standards.</span></span>
