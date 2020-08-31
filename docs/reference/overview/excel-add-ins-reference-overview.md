---
title: Excel JavaScript API 概述
description: 详细了解 Excel Javascript API
ms.date: 07/28/2020
ms.prod: excel
localization_priority: Priority
ms.openlocfilehash: e589bd7ce814211759cc731d828e9c180339ea1f
ms.sourcegitcommit: 9609bd5b4982cdaa2ea7637709a78a45835ffb19
ms.translationtype: HT
ms.contentlocale: zh-CN
ms.lasthandoff: 08/28/2020
ms.locfileid: "47293658"
---
# <a name="excel-javascript-api-overview"></a><span data-ttu-id="ab5d1-103">Excel JavaScript API 概述</span><span class="sxs-lookup"><span data-stu-id="ab5d1-103">Excel JavaScript API overview</span></span>

<span data-ttu-id="ab5d1-104">Excel 加载项通过使用 Office JavaScript API 与 Excel 中的对象进行交互，JavaScript API 包括两个 JavaScript 对象模型：</span><span class="sxs-lookup"><span data-stu-id="ab5d1-104">An Excel add-in interacts with objects in Excel by using the Office JavaScript API, which includes two JavaScript object models:</span></span>

* <span data-ttu-id="ab5d1-105">**Excel JavaScript API**：下面是针对 Excel 的[应用程序特定 API](../../develop/application-specific-api-model.md)。</span><span class="sxs-lookup"><span data-stu-id="ab5d1-105">**Excel JavaScript API**: These are the [application-specific APIs](../../develop/application-specific-api-model.md) for Excel.</span></span> <span data-ttu-id="ab5d1-106">[Excel JavaScript API](/javascript/api/excel) 随 Office 2016 一起引入，提供了强类型的对象，可用于访问工作表、区域、表格、图表等。</span><span class="sxs-lookup"><span data-stu-id="ab5d1-106">Introduced with Office 2016, the [Excel JavaScript API](/javascript/api/excel) provides strongly-typed objects that you can use to access worksheets, ranges, tables, charts, and more.</span></span>

* <span data-ttu-id="ab5d1-107">**通用 API**：[通用 API](/javascript/api/office) 随 Office 2013 引入，可用于访问多种类型的 Office 应用程序中常见的 UI、对话框和客户端设置等功能。</span><span class="sxs-lookup"><span data-stu-id="ab5d1-107">**Common APIs**: Introduced with Office 2013, the [Common API](/javascript/api/office) can be used to access features such as UI, dialogs, and client settings that are common across multiple types of Office applications.</span></span>

<span data-ttu-id="ab5d1-108">文档的本部分着重介绍了 Excel JavaScript API，它可用于开发面向 Excel 网页版或 Excel 2016 或更高版本的加载项中的大部分功能。</span><span class="sxs-lookup"><span data-stu-id="ab5d1-108">This section of the documentation focuses on the Excel JavaScript API, which you'll use to develop the majority of functionality in add-ins that target Excel on the web or Excel 2016 or later.</span></span> <span data-ttu-id="ab5d1-109">有关通用 API 的信息，请参阅[常见 JavaScript API 对象模型](../../develop/office-javascript-api-object-model.md)。</span><span class="sxs-lookup"><span data-stu-id="ab5d1-109">For information about the Common API, see [Common JavaScript API object model](../../develop/office-javascript-api-object-model.md).</span></span>

## <a name="learn-programming-concepts"></a><span data-ttu-id="ab5d1-110">了解编程概念</span><span class="sxs-lookup"><span data-stu-id="ab5d1-110">Learn programming concepts</span></span>

<span data-ttu-id="ab5d1-111">有关重要编程概念的信息，请参阅 [Excel JavaScript API 基本编程概念](../../excel/excel-add-ins-core-concepts.md)。</span><span class="sxs-lookup"><span data-stu-id="ab5d1-111">See [Fundamental programming concepts with the Excel JavaScript API](../../excel/excel-add-ins-core-concepts.md) for information about important programming concepts.</span></span>

<span data-ttu-id="ab5d1-112">有关使用 Excel JavaScript API 访问 Excel 中对象的实际经验，请完成 [Excel 加载项教程](../../tutorials/excel-tutorial.md)。</span><span class="sxs-lookup"><span data-stu-id="ab5d1-112">For hands-on experience using the Excel JavaScript API to access objects in Excel, complete the [Excel add-in tutorial](../../tutorials/excel-tutorial.md).</span></span>

## <a name="learn-api-capabilities"></a><span data-ttu-id="ab5d1-113">了解 API 功能</span><span class="sxs-lookup"><span data-stu-id="ab5d1-113">Learn API capabilities</span></span>

<span data-ttu-id="ab5d1-114">每个主要的 Excel API 功能都有一篇文章，探讨该功能的作用以及相关的对象模型。</span><span class="sxs-lookup"><span data-stu-id="ab5d1-114">Each major Excel API feature has an article exploring what that feature can do and the relevant object model.</span></span>

* [<span data-ttu-id="ab5d1-115">图表</span><span class="sxs-lookup"><span data-stu-id="ab5d1-115">Charts</span></span>](../../excel/excel-add-ins-charts.md)
* [<span data-ttu-id="ab5d1-116">备注</span><span class="sxs-lookup"><span data-stu-id="ab5d1-116">Comments</span></span>](../../excel/excel-add-ins-comments.md)
* [<span data-ttu-id="ab5d1-117">条件格式</span><span class="sxs-lookup"><span data-stu-id="ab5d1-117">Conditional formatting</span></span>](../../excel/excel-add-ins-conditional-formatting.md)
* [<span data-ttu-id="ab5d1-118">自定义函数</span><span class="sxs-lookup"><span data-stu-id="ab5d1-118">Custom functions</span></span>](../../excel/custom-functions-overview.md)
* [<span data-ttu-id="ab5d1-119">数据验证</span><span class="sxs-lookup"><span data-stu-id="ab5d1-119">Data validation</span></span>](../../excel/excel-add-ins-data-validation.md)
* [<span data-ttu-id="ab5d1-120">事件</span><span class="sxs-lookup"><span data-stu-id="ab5d1-120">Events</span></span>](../../excel/excel-add-ins-events.md)
* [<span data-ttu-id="ab5d1-121">多个范围 (RangeArea)</span><span class="sxs-lookup"><span data-stu-id="ab5d1-121">Multiple ranges (RangeArea)</span></span>](../../excel/excel-add-ins-multiple-ranges.md)
* [<span data-ttu-id="ab5d1-122">数据透视表</span><span class="sxs-lookup"><span data-stu-id="ab5d1-122">PivotTables</span></span>](../../excel/excel-add-ins-pivottables.md)
* <span data-ttu-id="ab5d1-123">[范围](../../excel/excel-add-ins-ranges.md)和[高级范围 API](../../excel/excel-add-ins-ranges-advanced.md)</span><span class="sxs-lookup"><span data-stu-id="ab5d1-123">[Ranges](../../excel/excel-add-ins-ranges.md) and [Advanced Range APIs](../../excel/excel-add-ins-ranges-advanced.md)</span></span>
* [<span data-ttu-id="ab5d1-124">性状</span><span class="sxs-lookup"><span data-stu-id="ab5d1-124">Shapes</span></span>](../../excel/excel-add-ins-shapes.md)
* [<span data-ttu-id="ab5d1-125">表格</span><span class="sxs-lookup"><span data-stu-id="ab5d1-125">Tables</span></span>](../../excel/excel-add-ins-tables.md)
* [<span data-ttu-id="ab5d1-126">工作簿和应用程序级 API</span><span class="sxs-lookup"><span data-stu-id="ab5d1-126">Workbooks and Application-level APIs</span></span>](../../excel/excel-add-ins-workbooks.md)
* [<span data-ttu-id="ab5d1-127">工作表</span><span class="sxs-lookup"><span data-stu-id="ab5d1-127">Worksheets</span></span>](../../excel/excel-add-ins-worksheets.md)

<span data-ttu-id="ab5d1-128">有关 Excel JavaScript API 对象模型的详细信息，请参阅 [Excel JavaScript API 参考文档](/javascript/api/excel)。</span><span class="sxs-lookup"><span data-stu-id="ab5d1-128">For detailed information about the Excel JavaScript API object model, see the [Excel JavaScript API reference documentation](/javascript/api/excel).</span></span>

## <a name="try-out-code-samples-in-script-lab"></a><span data-ttu-id="ab5d1-129">试用 Script Lab 中的代码示例</span><span class="sxs-lookup"><span data-stu-id="ab5d1-129">Try out code samples in Script Lab</span></span>

<span data-ttu-id="ab5d1-130">使用 [Script Lab](../../overview/explore-with-script-lab.md) 快速熟悉一系列展示如何使用 API 完成任务的内置示例。</span><span class="sxs-lookup"><span data-stu-id="ab5d1-130">Use [Script Lab](../../overview/explore-with-script-lab.md) to get started quickly with a collection of built-in samples that show how to complete tasks with the API.</span></span> <span data-ttu-id="ab5d1-131">你可以运行 Script Lab 中的示例来立即查看任务窗格或工作表中的结果，检查示例来了解 API 的工作原理，甚至可以使用示例来构建自己的加载项的原型。</span><span class="sxs-lookup"><span data-stu-id="ab5d1-131">You can run the samples in Script Lab to instantly see the result in the task pane or worksheet, examine the samples to learn how the API works, and even use samples to prototype your own add-in.</span></span>

## <a name="see-also"></a><span data-ttu-id="ab5d1-132">另请参阅</span><span class="sxs-lookup"><span data-stu-id="ab5d1-132">See also</span></span>

* [<span data-ttu-id="ab5d1-133">Excel 加载项文档</span><span class="sxs-lookup"><span data-stu-id="ab5d1-133">Excel add-ins documentation</span></span>](../../excel/index.yml)
* [<span data-ttu-id="ab5d1-134">Excel 加载项概述</span><span class="sxs-lookup"><span data-stu-id="ab5d1-134">Excel add-ins overview</span></span>](../../excel/excel-add-ins-overview.md)
* [<span data-ttu-id="ab5d1-135">Excel JavaScript API 参考</span><span class="sxs-lookup"><span data-stu-id="ab5d1-135">Excel JavaScript API reference</span></span>](/javascript/api/excel)
* [<span data-ttu-id="ab5d1-136">Office 客户端应用程序和 Office 加载项的平台可用性</span><span class="sxs-lookup"><span data-stu-id="ab5d1-136">Office client application and platform availability for Office Add-ins</span></span>](../../overview/office-add-in-availability.md)
* [<span data-ttu-id="ab5d1-137">使用特定于应用程序的 API 模型</span><span class="sxs-lookup"><span data-stu-id="ab5d1-137">Using the application-specific API model</span></span>](../../develop/application-specific-api-model.md)
