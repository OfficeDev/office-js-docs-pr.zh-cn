---
title: Office 外接程序的布局准则
description: ''
ms.date: 06/27/2018
localization_priority: Normal
ms.openlocfilehash: 9570bf041cf1df70ab95af656decb3c458c0d480
ms.sourcegitcommit: 9e7b4daa8d76c710b9d9dd4ae2e3c45e8fe07127
ms.translationtype: MT
ms.contentlocale: zh-CN
ms.lasthandoff: 04/24/2019
ms.locfileid: "32447166"
---
# <a name="layout"></a><span data-ttu-id="4c681-102">布局</span><span class="sxs-lookup"><span data-stu-id="4c681-102">Layout</span></span>
<span data-ttu-id="4c681-p101">嵌入到 Office 中的每个 HTML 容器都将有一个布局。这些布局是外接程序的主屏幕。你将在其中创建使客户能够启动操作、修改设置、查看、滚动或导航内容的体验。设计在屏幕中具有一致布局的外接程序，以确保体验的连续性。如果你有客户熟悉使用的现有网站，请考虑重新使用现有网页中的布局。对它们进行调整以协调适应 Office HTML 容器。</span><span class="sxs-lookup"><span data-stu-id="4c681-p101">Each HTML container embedded in Office will have a layout. These layouts are the main screens of your add-in. In them you will create experiences that enable customers to initiate actions, modify settings, view, scroll, or navigate content. Design your add-in with a consistent layouts across screens to guarantee continuity of experience. If you have an existing website that your customers are familiar with using, consider reusing layouts from your existing web pages. Adapt them to fit harmoniously within Office HTML containers.</span></span>

<span data-ttu-id="4c681-p102">有关布局指南，请参阅[任务窗格](task-pane-add-ins.md)、[内容](content-add-ins.md)和[对话框](dialog-boxes.md)。若要详细了解如何将 Office UI Fabric 组件装配到通用布局和用户体验流，请参阅[用户体验设计模式模板](ux-design-pattern-templates.md)。</span><span class="sxs-lookup"><span data-stu-id="4c681-p102">For guidelines on layout, see [Task pane](task-pane-add-ins.md), [Content](content-add-ins.md), and [Dialog box](dialog-boxes.md). For more information about how to assemble Office UI Fabric components into common layouts and user experience flows, see [UX design patterns templates](ux-design-pattern-templates.md).</span></span>

<span data-ttu-id="4c681-111">请遵循下面的一般布局指南：</span><span class="sxs-lookup"><span data-stu-id="4c681-111">Apply the following general guidelines for layouts:</span></span>

*   <span data-ttu-id="4c681-p103">避免 HTML 容器上的边距过窄或过宽。20 像素是理想的默认值。</span><span class="sxs-lookup"><span data-stu-id="4c681-p103">Avoid narrow or wide margins on your HTML containers. 20 pixels is a great default.</span></span>
*   <span data-ttu-id="4c681-p104">有意对齐元素。额外缩进和新对齐点应该有助于可视化层次结构。</span><span class="sxs-lookup"><span data-stu-id="4c681-p104">Align elements intentionally. Extra indents and new points of alignment should aid visual hierarchy.</span></span>
*   <span data-ttu-id="4c681-p105">Office 接口在 4 像素网格上。旨在使元素之间的填充保持在 4 的倍数。</span><span class="sxs-lookup"><span data-stu-id="4c681-p105">Office interfaces are on a 4px grid. Aim to keep your padding between elements at multiples of 4.</span></span>
*   <span data-ttu-id="4c681-118">界面过于拥挤可能导致混乱，并抑制触控交互的易用性。</span><span class="sxs-lookup"><span data-stu-id="4c681-118">Overcrowding your interface can lead to confusion and inhibit ease of use with touch interactions.</span></span>
*   <span data-ttu-id="4c681-p106">在各个屏幕之间保持布局一致性。意外布局更改类似于视觉错误，这将导致对解决方案的信心和信任的缺失。</span><span class="sxs-lookup"><span data-stu-id="4c681-p106">Keep layouts consistent across screens. Unexpected layout changes look like visual bugs that contribute to a lack of confidence and trust with your solution.</span></span>
*   <span data-ttu-id="4c681-p107">遵循公用的布局模式。约定可帮助用户了解如何使用界面。</span><span class="sxs-lookup"><span data-stu-id="4c681-p107">Follow common layout patterns. Conventions help users understand how to use an interface.</span></span>
*   <span data-ttu-id="4c681-123">避免冗余元素，如品牌或命令。</span><span class="sxs-lookup"><span data-stu-id="4c681-123">Avoid redundant elements like branding or commands.</span></span>
*   <span data-ttu-id="4c681-124">整合控件和视图，以避免需要过多地移动鼠标。</span><span class="sxs-lookup"><span data-stu-id="4c681-124">Consolidate controls and views to avoid requiring too much mouse movement.</span></span>
*   <span data-ttu-id="4c681-125">创建适应 HTML 容器宽度和高度的响应式体验。</span><span class="sxs-lookup"><span data-stu-id="4c681-125">Create responsive experiences that adapt to HTML container widths and heights.</span></span>
