---
title: Word JavaScript API 概述
description: ''
ms.date: 07/05/2019
ms.prod: word
localization_priority: Priority
ms.openlocfilehash: fbc9e8293642d1ab8edf32d568a5dab7ef77a8f0
ms.sourcegitcommit: c3673cc693fa7070e1b397922bd735ba3f9342f3
ms.translationtype: HT
ms.contentlocale: zh-CN
ms.lasthandoff: 07/05/2019
ms.locfileid: "35575623"
---
# <a name="word-javascript-api-overview"></a><span data-ttu-id="268b0-102">Word JavaScript API 概述</span><span class="sxs-lookup"><span data-stu-id="268b0-102">Word JavaScript API overview</span></span>

<span data-ttu-id="268b0-103">Word 加载项通过使用 Office JavaScript API 与 Word 中的对象进行交互，其中包括两个 JavaScript 对象模型：</span><span class="sxs-lookup"><span data-stu-id="268b0-103">An Excel add-in interacts with objects in Excel by using the JavaScript API for Office, which includes two JavaScript object models:</span></span>

* <span data-ttu-id="268b0-104">**Word JavaScript API**：[Word JavaScript API](/javascript/api/word) 随 Office 2016 一起引入，提供了强类型的对象，可用于访问 Word 文档中的对象和元数据。</span><span class="sxs-lookup"><span data-stu-id="268b0-104">**Word JavaScript API**: Introduced with Office 2016, the [Word JavaScript API](/javascript/api/word) provides strongly-typed objects that you can use to access objects and metadata in a Word document.</span></span> 

* <span data-ttu-id="268b0-105">**通用 API**：[通用 API](/javascript/api/office) 随 Office 2013 引入，可用于访问多种类型的 Office 应用程序中常见的 UI、对话框和客户端设置等功能。</span><span class="sxs-lookup"><span data-stu-id="268b0-105">**Common APIs**: Introduced with Office 2013, the [Common API](/javascript/api/office) can be used to access features such as UI, dialogs, and client settings that are common across multiple types of host applications such as Word, Excel, and PowerPoint.</span></span>

<span data-ttu-id="268b0-106">此文档部分重点介绍了 Word JavaScript AP，你可以通过此 API 开发面向 Word 网页版或 Word 2016 或更高版本的加载项中的大部分功能。</span><span class="sxs-lookup"><span data-stu-id="268b0-106">This section of the documentation focuses on the Word JavaScript API, which you'll use to develop the majority of functionality in add-ins that target Word on the web or Word 2016 or later.</span></span> <span data-ttu-id="268b0-107">有关公共 API 的信息，请参阅 [Office JavaScript API](../javascript-api-for-office.md)。</span><span class="sxs-lookup"><span data-stu-id="268b0-107">For more information about the distinction between host-specific APIs and Common APIs, see [JavaScript API for Office](../javascript-api-for-office.md).</span></span> 

## <a name="learn-programming-concepts"></a><span data-ttu-id="268b0-108">了解编程概念</span><span class="sxs-lookup"><span data-stu-id="268b0-108">Learn programming concepts</span></span>

<span data-ttu-id="268b0-109">有关重要编程概念的信息，请参阅 [Word JavaScript API 基本编程概念](../../word/word-add-ins-core-concepts.md)。</span><span class="sxs-lookup"><span data-stu-id="268b0-109">See [Fundamental programming concepts with the Word JavaScript API](../../word/word-add-ins-core-concepts.md) for information about important programming concepts.</span></span>
 
## <a name="learn-about-api-capabilities"></a><span data-ttu-id="268b0-110">了解 API 功能</span><span class="sxs-lookup"><span data-stu-id="268b0-110">Learn about API capabilities</span></span>

<span data-ttu-id="268b0-111">阅读此文档部分中的其他文章，了解如何[通过加载项获取文档](../../word/get-the-whole-document-from-an-add-in-for-word.md)、[使用搜索选项查找 Word 加载项中的文本](../../word/search-option-guidance.md)等。</span><span class="sxs-lookup"><span data-stu-id="268b0-111">Use other articles in this section of the documentation to learn how to [get the whole document from an add-in](../../word/get-the-whole-document-from-an-add-in-for-word.md), [use search options to find text in your Word add-in](../../word/search-option-guidance.md), and more.</span></span> <span data-ttu-id="268b0-112">有关可用文章的完整列表，请参阅目录。</span><span class="sxs-lookup"><span data-stu-id="268b0-112">See the table of contents of this reference for the complete list of members.</span></span>

<span data-ttu-id="268b0-113">有关使用 Word JavaScript API 访问 Word 中的对象的实践体验，请完成 [Word 加载项教程](../../tutorials/word-tutorial.md)。</span><span class="sxs-lookup"><span data-stu-id="268b0-113">For hands-on experience using the Word JavaScript API to access objects in Word, complete the [Word add-in tutorial](../../tutorials/word-tutorial.md).</span></span> 

<span data-ttu-id="268b0-114">有关 Word JavaScript API 对象模型的详细信息，请参阅 [Word JavaScript API 参考文档](/javascript/api/word)。</span><span class="sxs-lookup"><span data-stu-id="268b0-114">For detailed information about the Word JavaScript API, see the [Word JavaScript API reference documentation](/javascript/api/word).</span></span>

## <a name="try-out-code-samples-in-script-lab"></a><span data-ttu-id="268b0-115">试用 Script Lab 中的代码示例</span><span class="sxs-lookup"><span data-stu-id="268b0-115">Try out code samples in Script Lab</span></span>

<span data-ttu-id="268b0-116">使用 [Script Lab](../../overview/explore-with-script-lab.md) 快速熟悉一系列展示如何使用 API 完成任务的内置示例。</span><span class="sxs-lookup"><span data-stu-id="268b0-116">Use [Script Lab](../../overview/explore-with-script-lab.md) to get started quickly with a collection of built-in samples that show how to complete tasks with the API.</span></span> <span data-ttu-id="268b0-117">你可以运行 Script Lab 中的示例来立即查看任务窗格或文档中的结果，检查示例来了解 API 的工作原理，甚至可以使用示例来构建自己的加载项的原型。</span><span class="sxs-lookup"><span data-stu-id="268b0-117">You can run the samples in Script Lab to instantly see the result in the task pane or document, examine the samples to learn how the API works, and even use samples to prototype your own add-in.</span></span>

## <a name="see-also"></a><span data-ttu-id="268b0-118">另请参阅</span><span class="sxs-lookup"><span data-stu-id="268b0-118">See also</span></span>

- [<span data-ttu-id="268b0-119">Word 加载项文档</span><span class="sxs-lookup"><span data-stu-id="268b0-119">Word add-ins documentation</span></span>](../../word/index.md)
- [<span data-ttu-id="268b0-120">Word 加载项概述</span><span class="sxs-lookup"><span data-stu-id="268b0-120">Word add-ins overview</span></span>](../../word/word-add-ins-programming-overview.md)
- [<span data-ttu-id="268b0-121">Word JavaScript API 参考</span><span class="sxs-lookup"><span data-stu-id="268b0-121">Word JavaScript API reference</span></span>](/javascript/api/word)
- [<span data-ttu-id="268b0-122">Office 加载项主机和平台可用性</span><span class="sxs-lookup"><span data-stu-id="268b0-122">Office Add-in host and platform availability</span></span>](../../overview/office-add-in-availability.md)
- [<span data-ttu-id="268b0-123">API 开放性规范</span><span class="sxs-lookup"><span data-stu-id="268b0-123">API open specifications</span></span>](../openspec/openspec.md)
