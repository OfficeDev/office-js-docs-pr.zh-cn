---
title: Excel JavaScript API 概述
description: 详细了解 Excel Javascript API
ms.date: 02/19/2020
ms.prod: excel
localization_priority: Priority
ms.openlocfilehash: 336944f4315e9e5aea492ecbb8cd557a38edbb39
ms.sourcegitcommit: 6c381634c77d316f34747131860db0a0bced2529
ms.translationtype: HT
ms.contentlocale: zh-CN
ms.lasthandoff: 03/21/2020
ms.locfileid: "42891192"
---
# <a name="excel-javascript-api-overview"></a><span data-ttu-id="5832f-103">Excel JavaScript API 概述</span><span class="sxs-lookup"><span data-stu-id="5832f-103">Excel JavaScript API overview</span></span>

<span data-ttu-id="5832f-104">Excel 加载项通过使用 Office JavaScript API 与 Excel 中的对象进行交互，JavaScript API 包括两个 JavaScript 对象模型：</span><span class="sxs-lookup"><span data-stu-id="5832f-104">An Excel add-in interacts with objects in Excel by using the Office JavaScript API, which includes two JavaScript object models:</span></span>

* <span data-ttu-id="5832f-105">**Excel JavaScript API**：[Excel JavaScript API](/javascript/api/excel) 随 Office 2016 一起引入，提供了强类型的对象，可用于访问工作表、区域、表格、图表等。</span><span class="sxs-lookup"><span data-stu-id="5832f-105">**Excel JavaScript API**: Introduced with Office 2016, the [Excel JavaScript API](/javascript/api/excel) provides strongly-typed objects that you can use to access worksheets, ranges, tables, charts, and more.</span></span> 

* <span data-ttu-id="5832f-106">**通用 API**：[通用 API](/javascript/api/office) 随 Office 2013 引入，可用于访问多种类型的 Office 应用程序中常见的 UI、对话框和客户端设置等功能。</span><span class="sxs-lookup"><span data-stu-id="5832f-106">**Common APIs**: Introduced with Office 2013, the [Common API](/javascript/api/office) can be used to access features such as UI, dialogs, and client settings that are common across multiple types of Office applications.</span></span>

<span data-ttu-id="5832f-107">文档的本部分着重介绍了 Excel JavaScript API，它可用于开发面向 Excel 网页版或 Excel 2016 或更高版本的加载项中的大部分功能。</span><span class="sxs-lookup"><span data-stu-id="5832f-107">This section of the documentation focuses on the Excel JavaScript API, which you'll use to develop the majority of functionality in add-ins that target Excel on the web or Excel 2016 or later.</span></span> <span data-ttu-id="5832f-108">有关通用 API 的信息，请参阅[常见 JavaScript API 对象模型](../../develop/office-javascript-api-object-model.md)。</span><span class="sxs-lookup"><span data-stu-id="5832f-108">For information about the Common API, see [Common JavaScript API object model](../../develop/office-javascript-api-object-model.md).</span></span> 

## <a name="learn-programming-concepts"></a><span data-ttu-id="5832f-109">了解编程概念</span><span class="sxs-lookup"><span data-stu-id="5832f-109">Learn programming concepts</span></span>

<span data-ttu-id="5832f-110">有关重要编程概念的信息，请参阅以下文章：</span><span class="sxs-lookup"><span data-stu-id="5832f-110">See the following articles for information about important programming concepts:</span></span>
 
- [<span data-ttu-id="5832f-111">Excel JavaScript API 基本编程概念</span><span class="sxs-lookup"><span data-stu-id="5832f-111">Fundamental programming concepts with the Excel JavaScript API</span></span>](../../excel/excel-add-ins-core-concepts.md)

- [<span data-ttu-id="5832f-112">Excel JavaScript API 高级编程概念</span><span class="sxs-lookup"><span data-stu-id="5832f-112">Advanced programming concepts with the Excel JavaScript API</span></span>](../../excel/excel-add-ins-advanced-concepts.md)

## <a name="learn-about-api-capabilities"></a><span data-ttu-id="5832f-113">了解 API 功能</span><span class="sxs-lookup"><span data-stu-id="5832f-113">Learn about API capabilities</span></span>

<span data-ttu-id="5832f-114">阅读此文档部分中的其他文章，了解如何处理[事件](../../excel/excel-add-ins-events.md)、[图表](../../excel/excel-add-ins-charts.md)、[区域](../../excel/excel-add-ins-ranges.md)、[表格](../../excel/excel-add-ins-tables.md)、[工作表](../../excel/excel-add-ins-worksheets.md)等。</span><span class="sxs-lookup"><span data-stu-id="5832f-114">Use other articles in this section of the documentation to learn about working with [events](../../excel/excel-add-ins-events.md), [charts](../../excel/excel-add-ins-charts.md), [ranges](../../excel/excel-add-ins-ranges.md), [tables](../../excel/excel-add-ins-tables.md), [worksheets](../../excel/excel-add-ins-worksheets.md), and more.</span></span> <span data-ttu-id="5832f-115">在此部分中，你还可以找到有关 Excel JavaScript API 概念的指南，例如[使用 Excel 加载项共同创作](../../excel/co-authoring-in-excel-add-ins.md)、[数据验证](../../excel/excel-add-ins-data-validation.md)、[错误处理](../../excel/excel-add-ins-error-handling.md)和[性能优化](../../excel/performance.md)。</span><span class="sxs-lookup"><span data-stu-id="5832f-115">Also in this section, you'll find guidance about Excel JavaScript API concepts such as [coauthoring in Excel add-ins](../../excel/co-authoring-in-excel-add-ins.md), [data validation](../../excel/excel-add-ins-data-validation.md), [error handling](../../excel/excel-add-ins-error-handling.md), and [performance optimization](../../excel/performance.md).</span></span> <span data-ttu-id="5832f-116">有关可用文章的完整列表，请参阅目录。</span><span class="sxs-lookup"><span data-stu-id="5832f-116">See the table of contents for the complete list of available articles.</span></span>

<span data-ttu-id="5832f-117">有关使用 Excel JavaScript API 访问 Excel 中对象的实际经验，请完成 [Excel 加载项教程](../../tutorials/excel-tutorial.md)。</span><span class="sxs-lookup"><span data-stu-id="5832f-117">For hands-on experience using the Excel JavaScript API to access objects in Excel, complete the [Excel add-in tutorial](../../tutorials/excel-tutorial.md).</span></span> 

<span data-ttu-id="5832f-118">有关 Excel JavaScript API 对象模型的详细信息，请参阅 [Excel JavaScript API 参考文档](/javascript/api/excel)。</span><span class="sxs-lookup"><span data-stu-id="5832f-118">For detailed information about the Excel JavaScript API object model, see the [Excel JavaScript API reference documentation](/javascript/api/excel).</span></span>

## <a name="try-out-code-samples-in-script-lab"></a><span data-ttu-id="5832f-119">试用 Script Lab 中的代码示例</span><span class="sxs-lookup"><span data-stu-id="5832f-119">Try out code samples in Script Lab</span></span>

<span data-ttu-id="5832f-120">使用 [Script Lab](../../overview/explore-with-script-lab.md) 快速熟悉一系列展示如何使用 API 完成任务的内置示例。</span><span class="sxs-lookup"><span data-stu-id="5832f-120">Use [Script Lab](../../overview/explore-with-script-lab.md) to get started quickly with a collection of built-in samples that show how to complete tasks with the API.</span></span> <span data-ttu-id="5832f-121">你可以运行 Script Lab 中的示例来立即查看任务窗格或工作表中的结果，检查示例来了解 API 的工作原理，甚至可以使用示例来构建自己的加载项的原型。</span><span class="sxs-lookup"><span data-stu-id="5832f-121">You can run the samples in Script Lab to instantly see the result in the task pane or worksheet, examine the samples to learn how the API works, and even use samples to prototype your own add-in.</span></span>

## <a name="see-also"></a><span data-ttu-id="5832f-122">另请参阅</span><span class="sxs-lookup"><span data-stu-id="5832f-122">See also</span></span>

- [<span data-ttu-id="5832f-123">Excel 加载项文档</span><span class="sxs-lookup"><span data-stu-id="5832f-123">Excel add-ins documentation</span></span>](../../excel/index.md)
- [<span data-ttu-id="5832f-124">Excel 加载项概述</span><span class="sxs-lookup"><span data-stu-id="5832f-124">Excel add-ins overview</span></span>](../../excel/excel-add-ins-overview.md)
- [<span data-ttu-id="5832f-125">Excel JavaScript API 参考</span><span class="sxs-lookup"><span data-stu-id="5832f-125">Excel JavaScript API reference</span></span>](/javascript/api/excel)
- [<span data-ttu-id="5832f-126">Office 加载项主机和平台可用性</span><span class="sxs-lookup"><span data-stu-id="5832f-126">Office Add-in host and platform availability</span></span>](../../overview/office-add-in-availability.md)
