---
title: Office 外接程序的图标准则
description: 概述如何为外接程序命令设计图标以及新的和 Monoline 的设计样式。
ms.date: 12/09/2019
localization_priority: Normal
ms.openlocfilehash: b6a960b038b7e02f75101f589469db328465d6bc
ms.sourcegitcommit: be23b68eb661015508797333915b44381dd29bdb
ms.translationtype: MT
ms.contentlocale: zh-CN
ms.lasthandoff: 06/08/2020
ms.locfileid: "44607671"
---
# <a name="icons"></a><span data-ttu-id="93aa5-103">图标</span><span class="sxs-lookup"><span data-stu-id="93aa5-103">Icons</span></span>

<span data-ttu-id="93aa5-104">图标是行为或概念的可视化表示形式。</span><span class="sxs-lookup"><span data-stu-id="93aa5-104">Icons are the visual representation of a behavior or concept.</span></span> <span data-ttu-id="93aa5-105">它们通常用于为控件和命令添加含义。</span><span class="sxs-lookup"><span data-stu-id="93aa5-105">They are often used to add meaning to controls and commands.</span></span> <span data-ttu-id="93aa5-106">实际或符号化的视觉对象使用户能够以与标记帮助用户浏览其环境的相同方式浏览 UI。</span><span class="sxs-lookup"><span data-stu-id="93aa5-106">Visuals, either realistic or symbolic, enable the user to navigate the UI the same way signs help users navigate their environment.</span></span> <span data-ttu-id="93aa5-107">这些视觉对象应简单明了，并且只包含所需的详细信息，以使客户能够快速分析他们在选择控件时将会发生的操作。</span><span class="sxs-lookup"><span data-stu-id="93aa5-107">They should be simple, clear, and contain only the necessary details to enable customers to quickly parse what action will occur when they choose a control.</span></span>

<span data-ttu-id="93aa5-108">Office 功能区界面具有标准的视觉样式。</span><span class="sxs-lookup"><span data-stu-id="93aa5-108">Office ribbon interfaces have a standard visual style.</span></span> <span data-ttu-id="93aa5-109">这可以确保一致性并熟悉各个 Office 应用程序。</span><span class="sxs-lookup"><span data-stu-id="93aa5-109">This ensures consistency and familiarity across Office apps.</span></span> <span data-ttu-id="93aa5-110">这些准则将有助于你为解决方案设计一组适合作为 Office 固有部分的 PNG 资产。</span><span class="sxs-lookup"><span data-stu-id="93aa5-110">The guidelines will help you design a set of PNG assets for your solution that fit in as a natural part of Office.</span></span>

<span data-ttu-id="93aa5-p103">许多 HTML 容器包含带有插图的控件。使用 Office UI Fabric 的自定义字体在外接程序中呈现 Office 样式图标。Fabric 的图标字体包含很多针对可缩放的常见 Office 隐喻、颜色和样式的字形以满足你的需要。如果你有带自己图标集的现有视觉语言，则可在 HTML 画布中随意使用。构建自己带标准图标集的品牌的连续性是任何设计语言的重要组成部分。请注意避免与 Office 隐喻产生冲突导致客户混淆。</span><span class="sxs-lookup"><span data-stu-id="93aa5-p103">Many HTML containers contain controls with iconography. Use Office UI Fabric’s custom font to render Office styled icons in your add-in. Fabric’s icon font contains many glyphs for common Office metaphors that you can scale, color, and style to suit your needs. If you have an existing visual language with your own set of icons, feel free to use it in your HTML canvases. Building continuity with your own brand with a standard set of icons is an important part of any design language. Be careful to avoid creating confusion for customers by conflicting with Office metaphors.</span></span>

## <a name="design-icons-for-add-in-commands"></a><span data-ttu-id="93aa5-117">加载项命令的设计图标</span><span class="sxs-lookup"><span data-stu-id="93aa5-117">Design icons for add-in commands</span></span>

<span data-ttu-id="93aa5-118">[外接程序命令](add-in-commands.md)添加按钮、文本和 Office UI 图标。</span><span class="sxs-lookup"><span data-stu-id="93aa5-118">[Add-in commands](add-in-commands.md) add buttons, text, and icons to the Office UI.</span></span> <span data-ttu-id="93aa5-119">外接程序命令按钮应提供有意义的图标和标签，以便清楚地标识用户在使用命令时执行的操作。</span><span class="sxs-lookup"><span data-stu-id="93aa5-119">Your add-in command buttons should provide meaningful icons and labels that clearly identify the action the user is taking when they use a command.</span></span> <span data-ttu-id="93aa5-120">以下文章提供了样式和生产准则，可帮助您设计与 Office 无缝集成的图标。</span><span class="sxs-lookup"><span data-stu-id="93aa5-120">The following articles provide stylistic and production guidelines to help you design icons that integrate seamlessly with Office.</span></span>

- <span data-ttu-id="93aa5-121">有关 Office 365 的 Monoline 样式，请参阅[适用于 Office 外接程序的 Monoline 样式图标准则](add-in-icons-monoline.md)。</span><span class="sxs-lookup"><span data-stu-id="93aa5-121">For the Monoline style of Office 365, see [Monoline style icon guidelines for Office Add-ins](add-in-icons-monoline.md).</span></span>
- <span data-ttu-id="93aa5-122">有关非订阅 Office 2013 + 的全新样式，请参阅[适用于 Office 外接程序的新样式图标指南](add-in-icons-fresh.md)。</span><span class="sxs-lookup"><span data-stu-id="93aa5-122">For the Fresh style of non-subscription Office 2013+, see [Fresh style icon guidelines for Office Add-ins](add-in-icons-fresh.md).</span></span>

> [!NOTE]
> <span data-ttu-id="93aa5-123">您必须选择一个样式或另一个样式，并且您的外接程序将使用相同的图标，无论它是在 Office 365 还是非订阅办公室中运行。</span><span class="sxs-lookup"><span data-stu-id="93aa5-123">You must choose one style or the other and your add-in will use the same icons whether it is running in Office 365 or non-subscription Office.</span></span>

## <a name="see-also"></a><span data-ttu-id="93aa5-124">另请参阅</span><span class="sxs-lookup"><span data-stu-id="93aa5-124">See also</span></span>

- [<span data-ttu-id="93aa5-125">加载项开发最佳做法</span><span class="sxs-lookup"><span data-stu-id="93aa5-125">Add-in development best practices</span></span>](../concepts/add-in-development-best-practices.md)
- [<span data-ttu-id="93aa5-126">Excel、Word 和 PowerPoint 的加载项命令</span><span class="sxs-lookup"><span data-stu-id="93aa5-126">Add-in commands for Excel, Word, and PowerPoint</span></span>](../design/add-in-commands.md)
