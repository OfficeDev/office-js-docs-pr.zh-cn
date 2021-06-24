---
title: Office 加载项中的 Word JavaScript 对象模型
description: 了解特定于 Word 的 JavaScript 对象模型中最重要的类。
ms.date: 10/14/2020
localization_priority: Priority
ms.openlocfilehash: 43ca88e7899e2ff11748dc91d5c8a5059d8bb559
ms.sourcegitcommit: ee9e92a968e4ad23f1e371f00d4888e4203ab772
ms.translationtype: HT
ms.contentlocale: zh-CN
ms.lasthandoff: 06/23/2021
ms.locfileid: "53077230"
---
# <a name="word-javascript-object-model-in-office-add-ins"></a><span data-ttu-id="cd5e8-103">Office 加载项中的 Word JavaScript 对象模型</span><span class="sxs-lookup"><span data-stu-id="cd5e8-103">Word JavaScript object model in Office Add-ins</span></span>

<span data-ttu-id="cd5e8-104">本文介绍使用 [Word JavaScript API](../reference/overview/word-add-ins-reference-overview.md) 生成加载项的基本概念。它介绍了使用 API 的基本核心概念。</span><span class="sxs-lookup"><span data-stu-id="cd5e8-104">This article describes concepts that are fundamental to using the [Word JavaScript API](../reference/overview/word-add-ins-reference-overview.md) to build add-ins. It introduces core concepts that are fundamental to using the API.</span></span>

> [!IMPORTANT]
> <span data-ttu-id="cd5e8-105">请参阅[使用特定于应用程序的 API 模型](../develop/application-specific-api-model.md)，以了解 Word API 的异步性质以及它们如何与文档协同工作。</span><span class="sxs-lookup"><span data-stu-id="cd5e8-105">See [Using the application-specific API model](../develop/application-specific-api-model.md) to learn about the asynchronous nature of the Word APIs and how they work with the document.</span></span>

## <a name="officejs-apis-for-word"></a><span data-ttu-id="cd5e8-106">适用于 Word 的 Office.js API</span><span class="sxs-lookup"><span data-stu-id="cd5e8-106">Office.js APIs for Word</span></span>

<span data-ttu-id="cd5e8-107">Word 加载项通过使用 Office JavaScript API 与 Excel 中的对象进行交互，JavaScript API包括两个 JavaScript 对象模型：</span><span class="sxs-lookup"><span data-stu-id="cd5e8-107">A Word add-in interacts with objects in Excel by using the Office JavaScript API, which includes two JavaScript object models:</span></span>

* <span data-ttu-id="cd5e8-108">**Word JavaScript API**：[Word JavaScript API](../reference/overview/word-add-ins-reference-overview.md) 提供了强类型的对象，可用于访问文档、范围、表格、列表、格式等。</span><span class="sxs-lookup"><span data-stu-id="cd5e8-108">**Word JavaScript API**: The [Word JavaScript API](../reference/overview/word-add-ins-reference-overview.md) provides strongly-typed objects that you can use to access the document, ranges, tables, lists, formatting, and more.</span></span>

* <span data-ttu-id="cd5e8-109">**通用 API**：[通用 API](/javascript/api/office) 可用于访问在多种类型的 Office 应用程序中都很常见的 UI、对话框和客户端设置等功能。</span><span class="sxs-lookup"><span data-stu-id="cd5e8-109">**Common APIs**: The [Common API](/javascript/api/office) can be used to access features such as UI, dialogs, and client settings that are common across multiple types of Office applications.</span></span>

<span data-ttu-id="cd5e8-p101">你可能会使用 Word JavaScript API 开发面向 Word 的加载项中的大部分功能，同时还可以使用通用 API 中的对象。例如：</span><span class="sxs-lookup"><span data-stu-id="cd5e8-p101">While you'll likely use the Word JavaScript API to develop the majority of functionality in add-ins that target Word, you'll also use objects in the Common API. For example:</span></span>

* <span data-ttu-id="cd5e8-112">[Context](/javascript/api/office/office.context)：`Context` 对象表示加载项的运行时环境，并提供对 API 的关键对象的访问。</span><span class="sxs-lookup"><span data-stu-id="cd5e8-112">[Context](/javascript/api/office/office.context): The `Context` object represents the runtime environment of the add-in and provides access to key objects of the API.</span></span> <span data-ttu-id="cd5e8-113">它由文档配置详细信息（如 `contentLanguage` 和 `officeTheme`）组成，并提供有关加载项的运行时环境（如 `host` 和 `platform`）的信息。</span><span class="sxs-lookup"><span data-stu-id="cd5e8-113">It consists of document configuration details such as `contentLanguage` and `officeTheme` and also provides information about the add-in's runtime environment such as `host` and `platform`.</span></span> <span data-ttu-id="cd5e8-114">此外，它还提供了 `requirements.isSetSupported()` 方法，可用于检查运行加载项的 Excel 应用程序是否支持指定的要求集。</span><span class="sxs-lookup"><span data-stu-id="cd5e8-114">Additionally, it provides the `requirements.isSetSupported()` method, which you can use to check whether a specified requirement set is supported by the Excel application where the add-in is running.</span></span>
* <span data-ttu-id="cd5e8-115">[Document](/javascript/api/office/office.document)：`Document` 对象提供 `getFileAsync()` 方法，用于下载运行加载项的 Word 文件。</span><span class="sxs-lookup"><span data-stu-id="cd5e8-115">[Document](/javascript/api/office/office.document): The `Document` object provides the `getFileAsync()` method, which you can use to download the Word file where the add-in is running.</span></span>

![Word JS API 和通用 API 之间的差异。](../images/word-js-api-common-api.png)

## <a name="word-specific-object-model"></a><span data-ttu-id="cd5e8-117">特定于 Word 的对象模型</span><span class="sxs-lookup"><span data-stu-id="cd5e8-117">Word-specific object model</span></span>

<span data-ttu-id="cd5e8-118">若要了解 Word API，则必须了解文档的各个组件之间如何相互关联。</span><span class="sxs-lookup"><span data-stu-id="cd5e8-118">To understand the Word APIs, you must understand how the components of a document are related to one another.</span></span>

* <span data-ttu-id="cd5e8-119">**Document** 包含 **Section** 以及设置和自定义 XML 部件等文档级实体。</span><span class="sxs-lookup"><span data-stu-id="cd5e8-119">The **Document** contains the **Section** s, and document-level entities such as settings and custom XML parts.</span></span>
* <span data-ttu-id="cd5e8-120">**Section** 包含 **Body**。</span><span class="sxs-lookup"><span data-stu-id="cd5e8-120">A **Section** contains a **Body**.</span></span>
* <span data-ttu-id="cd5e8-121">通过 **Body** 可以访问 **Paragraph**、**ContentControl** 和 **Range** 等对象。</span><span class="sxs-lookup"><span data-stu-id="cd5e8-121">A **Body** gives access to **Paragraph** s, **ContentControl** s, and **Range** objects, among others.</span></span>
* <span data-ttu-id="cd5e8-122">**Range** 表示连续的内容区域，包括文本、空白区域、**Table** 和图像。</span><span class="sxs-lookup"><span data-stu-id="cd5e8-122">A **Range** represents a contiguous area of content, including text, white space, **Table** s, and images.</span></span> <span data-ttu-id="cd5e8-123">此外，它还包含大部分文本操作方法。</span><span class="sxs-lookup"><span data-stu-id="cd5e8-123">It also contains most of the text manipulation methods.</span></span>
* <span data-ttu-id="cd5e8-124">**List** 表示带标号或项目符号的列表中的文本。</span><span class="sxs-lookup"><span data-stu-id="cd5e8-124">A **List** represents text in a numbered or bulleted list.</span></span>

## <a name="see-also"></a><span data-ttu-id="cd5e8-125">另请参阅</span><span class="sxs-lookup"><span data-stu-id="cd5e8-125">See also</span></span>

- [<span data-ttu-id="cd5e8-126">Word JavaScript API 概述</span><span class="sxs-lookup"><span data-stu-id="cd5e8-126">Word JavaScript API overview</span></span>](../reference/overview/word-add-ins-reference-overview.md)
- [<span data-ttu-id="cd5e8-127">生成首个 Word 加载项</span><span class="sxs-lookup"><span data-stu-id="cd5e8-127">Build your first Word add-in</span></span>](../quickstarts/word-quickstart.md)
- [<span data-ttu-id="cd5e8-128">Word 加载项教程</span><span class="sxs-lookup"><span data-stu-id="cd5e8-128">Word add-in tutorial</span></span>](../tutorials/word-tutorial.md)
- [<span data-ttu-id="cd5e8-129">Word JavaScript API 参考</span><span class="sxs-lookup"><span data-stu-id="cd5e8-129">Word JavaScript API reference</span></span>](/javascript/api/word)
- [<span data-ttu-id="cd5e8-130">了解 Microsoft 365 开发人员计划</span><span class="sxs-lookup"><span data-stu-id="cd5e8-130">Learn about the Microsoft 365 Developer Program</span></span>](https://developer.microsoft.com/microsoft-365/dev-program)
