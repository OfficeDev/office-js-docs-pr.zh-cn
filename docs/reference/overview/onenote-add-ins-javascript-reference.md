---
title: OneNote JavaScript API 概述
description: 详细了解 OneNote JavaScript API
ms.date: 02/19/2020
ms.prod: onenote
localization_priority: Priority
ms.openlocfilehash: 70c3bca323084630f1926b501900bca26cf54304
ms.sourcegitcommit: fa4e81fcf41b1c39d5516edf078f3ffdbd4a3997
ms.translationtype: HT
ms.contentlocale: zh-CN
ms.lasthandoff: 03/17/2020
ms.locfileid: "42719917"
---
# <a name="onenote-javascript-api-overview"></a><span data-ttu-id="9f23d-103">OneNote JavaScript API 概述</span><span class="sxs-lookup"><span data-stu-id="9f23d-103">OneNote JavaScript API overview</span></span>

<span data-ttu-id="9f23d-104">OneNote 加载项通过使用 Office JavaScript API 与 OneNote web 版中的对象进行交互，其中包括两个 JavaScript 对象模型：</span><span class="sxs-lookup"><span data-stu-id="9f23d-104">A OneNote add-in interacts with objects in OneNote on the web by using the Office JavaScript API, which includes two JavaScript object models:</span></span>

* <span data-ttu-id="9f23d-105">**OneNote JavaScript API**：[OneNote JavaScript API](/javascript/api/onenote) 随 Office 2016 一起引入，提供了强类型的对象，可用于访问 OneNote web 版中的对象。</span><span class="sxs-lookup"><span data-stu-id="9f23d-105">**OneNote JavaScript API**: Introduced with Office 2016, the [OneNote JavaScript API](/javascript/api/onenote) provides strongly-typed objects that you can use to access objects in OneNote on the web.</span></span> 

* <span data-ttu-id="9f23d-106">**通用 API**：[通用 API](/javascript/api/office) 随 Office 2013 引入，可用于访问多种类型的 Office 应用程序中常见的 UI、对话框和客户端设置等功能。</span><span class="sxs-lookup"><span data-stu-id="9f23d-106">**Common APIs**: Introduced with Office 2013, the [Common API](/javascript/api/office) can be used to access features such as UI, dialogs, and client settings that are common across multiple types of Office applications.</span></span>

<span data-ttu-id="9f23d-107">此文档部分重点介绍了 OneNote JavaScript AP，你可通过此 API 开发面向 OneNote web 版的加载项中的大部分功能。</span><span class="sxs-lookup"><span data-stu-id="9f23d-107">This section of the documentation focuses on the OneNote JavaScript API, which you'll use to develop the majority of functionality in add-ins that target OneNote on the web.</span></span> <span data-ttu-id="9f23d-108">有关通用 API 的信息，请参阅[常见 JavaScript API 对象模型](../../develop/office-javascript-api-object-model.md)。</span><span class="sxs-lookup"><span data-stu-id="9f23d-108">For information about the Common API, see [Common JavaScript API object model](../../develop/office-javascript-api-object-model.md).</span></span> 

## <a name="learn-programming-concepts"></a><span data-ttu-id="9f23d-109">了解编程概念</span><span class="sxs-lookup"><span data-stu-id="9f23d-109">Learn programming concepts</span></span>

<span data-ttu-id="9f23d-110">有关重要编程概念的信息，请参阅以下文章：</span><span class="sxs-lookup"><span data-stu-id="9f23d-110">See the following articles for information about important programming concepts:</span></span>

- [<span data-ttu-id="9f23d-111">OneNote JavaScript API 编程概述</span><span class="sxs-lookup"><span data-stu-id="9f23d-111">OneNote JavaScript API programming overview</span></span>](../../onenote/onenote-add-ins-programming-overview.md)

- [<span data-ttu-id="9f23d-112">使用 OneNote 页面内容</span><span class="sxs-lookup"><span data-stu-id="9f23d-112">Work with OneNote page content</span></span>](../../onenote/onenote-add-ins-page-content.md)

## <a name="learn-about-api-capabilities"></a><span data-ttu-id="9f23d-113">了解 API 功能</span><span class="sxs-lookup"><span data-stu-id="9f23d-113">Learn about API capabilities</span></span>

<span data-ttu-id="9f23d-114">有关使用 OneNote JavaScript API 与 OneNote web 版中的内容进行交互的实践体验，请完成 [OneNote 加载项快速入门](../../quickstarts/onenote-quickstart.md)。</span><span class="sxs-lookup"><span data-stu-id="9f23d-114">For hands-on experience using the OneNote JavaScript API to interact with content in OneNote on the web, complete the [OneNote add-in quick start](../../quickstarts/onenote-quickstart.md).</span></span> 

<span data-ttu-id="9f23d-115">有关 OneNote JavaScript API 对象模型的详细信息，请参阅 [OneNote JavaScript API 参考文档](/javascript/api/onenote)。</span><span class="sxs-lookup"><span data-stu-id="9f23d-115">For detailed information about the OneNote JavaScript API object model, see the [OneNote JavaScript API reference documentation](/javascript/api/onenote).</span></span>

## <a name="see-also"></a><span data-ttu-id="9f23d-116">另请参阅</span><span class="sxs-lookup"><span data-stu-id="9f23d-116">See also</span></span>

- [<span data-ttu-id="9f23d-117">OneNote 加载项文档</span><span class="sxs-lookup"><span data-stu-id="9f23d-117">OneNote add-ins documentation</span></span>](../../onenote/index.md)
- [<span data-ttu-id="9f23d-118">OneNote 加载项概述</span><span class="sxs-lookup"><span data-stu-id="9f23d-118">OneNote add-ins overview</span></span>](../../onenote/onenote-add-ins-programming-overview.md)
- [<span data-ttu-id="9f23d-119">OneNote JavaScript API 参考</span><span class="sxs-lookup"><span data-stu-id="9f23d-119">OneNote JavaScript API reference</span></span>](/javascript/api/onenote)
- [<span data-ttu-id="9f23d-120">Office 加载项主机和平台可用性</span><span class="sxs-lookup"><span data-stu-id="9f23d-120">Office Add-in host and platform availability</span></span>](../../overview/office-add-in-availability.md)

