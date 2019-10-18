---
title: Excel JavaScript API 概述
description: ''
ms.date: 07/05/2019
ms.prod: excel
localization_priority: Priority
ms.openlocfilehash: e6064bf7e7dce6931079fc2d3eb262533da7edf3
ms.sourcegitcommit: c3673cc693fa7070e1b397922bd735ba3f9342f3
ms.translationtype: HT
ms.contentlocale: zh-CN
ms.lasthandoff: 07/05/2019
ms.locfileid: "35575630"
---
# <a name="excel-javascript-api-overview"></a><span data-ttu-id="34efd-102">Excel JavaScript API 概述</span><span class="sxs-lookup"><span data-stu-id="34efd-102">Excel JavaScript API overview</span></span>

<span data-ttu-id="34efd-103">Excel 加载项通过使用适用于 Office 的 JavaScript API 与 Excel 中的对象进行交互，适用于 Office 的 JavaScript API包括两个 JavaScript 对象模型：</span><span class="sxs-lookup"><span data-stu-id="34efd-103">An Excel add-in interacts with objects in Excel by using the JavaScript API for Office, which includes two JavaScript object models:</span></span>

* <span data-ttu-id="34efd-104">**Excel JavaScript API**：[Excel JavaScript API](/javascript/api/excel) 随 Office 2016 一起引入，提供了强类型的对象，可用于访问工作表、区域、表格、图表等。</span><span class="sxs-lookup"><span data-stu-id="34efd-104">**Excel JavaScript API**: Introduced with Office 2016, the [Excel JavaScript API](/javascript/api/excel) provides strongly-typed objects that you can use to access worksheets, ranges, tables, charts, and more.</span></span> 

* <span data-ttu-id="34efd-105">**通用 API**：[通用 API](/javascript/api/office) 随 Office 2013 引入，可用于访问多种类型的 Office 应用程序中常见的 UI、对话框和客户端设置等功能。</span><span class="sxs-lookup"><span data-stu-id="34efd-105">**Common APIs**: Introduced with Office 2013, the [Common API](/javascript/api/office) can be used to access features such as UI, dialogs, and client settings that are common across multiple types of host applications such as Word, Excel, and PowerPoint.</span></span>

<span data-ttu-id="34efd-106">文档的本部分着重介绍了 Excel JavaScript API，它可用于开发面向 Excel 网页版或 Excel 2016 或更高版本的加载项中的大部分功能。</span><span class="sxs-lookup"><span data-stu-id="34efd-106">This section of the documentation focuses on the Excel JavaScript API, which you'll use to develop the majority of functionality in add-ins that target Excel on the web or Excel 2016 or later.</span></span> <span data-ttu-id="34efd-107">有关通用 API 的信息，请参阅[适用于 Office 的 JavaScript API](../javascript-api-for-office.md)。</span><span class="sxs-lookup"><span data-stu-id="34efd-107">For more information about the distinction between host-specific APIs and Common APIs, see [JavaScript API for Office](../javascript-api-for-office.md).</span></span> 

## <a name="learn-programming-concepts"></a><span data-ttu-id="34efd-108">了解编程概念</span><span class="sxs-lookup"><span data-stu-id="34efd-108">Learn programming concepts</span></span>

<span data-ttu-id="34efd-109">有关重要编程概念的信息，请参阅以下文章：</span><span class="sxs-lookup"><span data-stu-id="34efd-109">See the following articles for information about important programming concepts:</span></span>
 
- [<span data-ttu-id="34efd-110">Excel JavaScript API 基本编程概念</span><span class="sxs-lookup"><span data-stu-id="34efd-110">Fundamental programming concepts with the Excel JavaScript API</span></span>](../../excel/excel-add-ins-core-concepts.md)

- [<span data-ttu-id="34efd-111">Excel JavaScript API 高级编程概念</span><span class="sxs-lookup"><span data-stu-id="34efd-111">Advanced programming concepts with the Excel JavaScript API</span></span>](../../excel/excel-add-ins-advanced-concepts.md)

## <a name="learn-about-api-capabilities"></a><span data-ttu-id="34efd-112">了解 API 功能</span><span class="sxs-lookup"><span data-stu-id="34efd-112">Learn about API capabilities</span></span>

<span data-ttu-id="34efd-113">阅读此文档部分中的其他文章，了解如何处理[事件](../../excel/excel-add-ins-events.md)、[图表](../../excel/excel-add-ins-charts.md)、[区域](../../excel/excel-add-ins-ranges.md)、[表格](../../excel/excel-add-ins-tables.md)、[工作表](../../excel/excel-add-ins-worksheets.md)等。</span><span class="sxs-lookup"><span data-stu-id="34efd-113">Use other articles in this section of the documentation to learn about working with [events](../../excel/excel-add-ins-events.md), [charts](../../excel/excel-add-ins-charts.md), [ranges](../../excel/excel-add-ins-ranges.md), [tables](../../excel/excel-add-ins-tables.md), [worksheets](../../excel/excel-add-ins-worksheets.md), and more.</span></span> <span data-ttu-id="34efd-114">在此部分中，你还可以找到有关 Excel JavaScript API 概念的指南，例如[使用 Excel 加载项共同创作](../../excel/co-authoring-in-excel-add-ins.md)、[数据验证](../../excel/excel-add-ins-data-validation.md)、[错误处理](../../excel/excel-add-ins-error-handling.md)和[性能优化](../../excel/performance.md)。</span><span class="sxs-lookup"><span data-stu-id="34efd-114">Also in this section, you'll find guidance about Excel JavaScript API concepts such as [coauthoring in Excel add-ins](../../excel/co-authoring-in-excel-add-ins.md), [data validation](../../excel/excel-add-ins-data-validation.md), [error handling](../../excel/excel-add-ins-error-handling.md), and [performance optimization](../../excel/performance.md).</span></span> <span data-ttu-id="34efd-115">有关可用文章的完整列表，请参阅目录。</span><span class="sxs-lookup"><span data-stu-id="34efd-115">See the table of contents of this reference for the complete list of members.</span></span>

<span data-ttu-id="34efd-116">有关使用 Excel JavaScript API 访问 Excel 中对象的实际经验，请完成 [Excel 加载项教程](../../tutorials/excel-tutorial.md)。</span><span class="sxs-lookup"><span data-stu-id="34efd-116">For hands-on experience using the Excel JavaScript API to access objects in Excel, complete the [Excel add-in tutorial](../../tutorials/excel-tutorial.md).</span></span> 

<span data-ttu-id="34efd-117">有关 Excel JavaScript API 对象模型的详细信息，请参阅 [Excel JavaScript API 参考文档](/javascript/api/excel)。</span><span class="sxs-lookup"><span data-stu-id="34efd-117">For detailed information about the Excel JavaScript API, see the [Excel JavaScript API reference documentation](/javascript/api/excel).</span></span>

## <a name="try-out-code-samples-in-script-lab"></a><span data-ttu-id="34efd-118">试用 Script Lab 中的代码示例</span><span class="sxs-lookup"><span data-stu-id="34efd-118">Try out code samples in Script Lab</span></span>

<span data-ttu-id="34efd-119">使用 [Script Lab](../../overview/explore-with-script-lab.md) 快速熟悉一系列展示如何使用 API 完成任务的内置示例。</span><span class="sxs-lookup"><span data-stu-id="34efd-119">Use [Script Lab](../../overview/explore-with-script-lab.md) to get started quickly with a collection of built-in samples that show how to complete tasks with the API.</span></span> <span data-ttu-id="34efd-120">你可以运行 Script Lab 中的示例来立即查看任务窗格或工作表中的结果，检查示例来了解 API 的工作原理，甚至可以使用示例来构建自己的加载项的原型。</span><span class="sxs-lookup"><span data-stu-id="34efd-120">You can run the samples in Script Lab to instantly see the result in the task pane or worksheet, examine the samples to learn how the API works, and even use samples to prototype your own add-in.</span></span>

## <a name="see-also"></a><span data-ttu-id="34efd-121">另请参阅</span><span class="sxs-lookup"><span data-stu-id="34efd-121">See also</span></span>

- [<span data-ttu-id="34efd-122">Excel 加载项文档</span><span class="sxs-lookup"><span data-stu-id="34efd-122">Excel add-ins documentation</span></span>](../../excel/index.md)
- [<span data-ttu-id="34efd-123">Excel 加载项概述</span><span class="sxs-lookup"><span data-stu-id="34efd-123">Excel add-ins overview</span></span>](../../excel/excel-add-ins-overview.md)
- [<span data-ttu-id="34efd-124">Excel JavaScript API 参考</span><span class="sxs-lookup"><span data-stu-id="34efd-124">Excel JavaScript API reference</span></span>](/javascript/api/excel)
- [<span data-ttu-id="34efd-125">Office 加载项主机和平台可用性</span><span class="sxs-lookup"><span data-stu-id="34efd-125">Office Add-in host and platform availability</span></span>](../../overview/office-add-in-availability.md)
- [<span data-ttu-id="34efd-126">API 开放性规范</span><span class="sxs-lookup"><span data-stu-id="34efd-126">API open specifications</span></span>](../openspec/openspec.md)
