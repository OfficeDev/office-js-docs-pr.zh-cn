---
title: OneNote JavaScript API 概述
description: 详细了解 OneNote JavaScript API
ms.date: 07/28/2020
ms.prod: onenote
localization_priority: Priority
ms.openlocfilehash: 08e98e81e46ca62178235454d3ba44f35be2eec2
ms.sourcegitcommit: 9609bd5b4982cdaa2ea7637709a78a45835ffb19
ms.translationtype: HT
ms.contentlocale: zh-CN
ms.lasthandoff: 08/28/2020
ms.locfileid: "47293630"
---
# <a name="onenote-javascript-api-overview"></a><span data-ttu-id="32445-103">OneNote JavaScript API 概述</span><span class="sxs-lookup"><span data-stu-id="32445-103">OneNote JavaScript API overview</span></span>

<span data-ttu-id="32445-104">OneNote 加载项通过使用 Office JavaScript API 与 OneNote web 版中的对象进行交互，其中包括两个 JavaScript 对象模型：</span><span class="sxs-lookup"><span data-stu-id="32445-104">A OneNote add-in interacts with objects in OneNote on the web by using the Office JavaScript API, which includes two JavaScript object models:</span></span>

* <span data-ttu-id="32445-105">**OneNote JavaScript API**：下面是针对 OneNote 的[应用程序专用 API](../../develop/application-specific-api-model.md)。</span><span class="sxs-lookup"><span data-stu-id="32445-105">**OneNote JavaScript API**: These are the [application-specific APIs](../../develop/application-specific-api-model.md) for OneNote.</span></span> <span data-ttu-id="32445-106">[OneNote JavaScript API](/javascript/api/onenote) 是随 Office 2016 一起推出的，它提供了强类型的对象，可用于访问 OneNote 网页版中的对象。</span><span class="sxs-lookup"><span data-stu-id="32445-106">Introduced with Office 2016, the [OneNote JavaScript API](/javascript/api/onenote) provides strongly-typed objects that you can use to access objects in OneNote on the web.</span></span>

* <span data-ttu-id="32445-107">**通用 API**：[通用 API](/javascript/api/office) 随 Office 2013 引入，可用于访问多种类型的 Office 应用程序中常见的 UI、对话框和客户端设置等功能。</span><span class="sxs-lookup"><span data-stu-id="32445-107">**Common APIs**: Introduced with Office 2013, the [Common API](/javascript/api/office) can be used to access features such as UI, dialogs, and client settings that are common across multiple types of Office applications.</span></span>

<span data-ttu-id="32445-108">此文档部分重点介绍了 OneNote JavaScript AP，你可通过此 API 开发面向 OneNote web 版的加载项中的大部分功能。</span><span class="sxs-lookup"><span data-stu-id="32445-108">This section of the documentation focuses on the OneNote JavaScript API, which you'll use to develop the majority of functionality in add-ins that target OneNote on the web.</span></span> <span data-ttu-id="32445-109">有关通用 API 的信息，请参阅[常见 JavaScript API 对象模型](../../develop/office-javascript-api-object-model.md)。</span><span class="sxs-lookup"><span data-stu-id="32445-109">For information about the Common API, see [Common JavaScript API object model](../../develop/office-javascript-api-object-model.md).</span></span>

## <a name="learn-programming-concepts"></a><span data-ttu-id="32445-110">了解编程概念</span><span class="sxs-lookup"><span data-stu-id="32445-110">Learn programming concepts</span></span>

<span data-ttu-id="32445-111">有关重要编程概念的信息，请参阅以下文章：</span><span class="sxs-lookup"><span data-stu-id="32445-111">See the following articles for information about important programming concepts:</span></span>

* [<span data-ttu-id="32445-112">OneNote JavaScript API 编程概述</span><span class="sxs-lookup"><span data-stu-id="32445-112">OneNote JavaScript API programming overview</span></span>](../../onenote/onenote-add-ins-programming-overview.md)
* [<span data-ttu-id="32445-113">使用 OneNote 页面内容</span><span class="sxs-lookup"><span data-stu-id="32445-113">Work with OneNote page content</span></span>](../../onenote/onenote-add-ins-page-content.md)

## <a name="learn-about-api-capabilities"></a><span data-ttu-id="32445-114">了解 API 功能</span><span class="sxs-lookup"><span data-stu-id="32445-114">Learn about API capabilities</span></span>

<span data-ttu-id="32445-115">有关使用 OneNote JavaScript API 与 OneNote web 版中的内容进行交互的实践体验，请完成 [OneNote 加载项快速入门](../../quickstarts/onenote-quickstart.md)。</span><span class="sxs-lookup"><span data-stu-id="32445-115">For hands-on experience using the OneNote JavaScript API to interact with content in OneNote on the web, complete the [OneNote add-in quick start](../../quickstarts/onenote-quickstart.md).</span></span>

<span data-ttu-id="32445-116">有关 OneNote JavaScript API 对象模型的详细信息，请参阅 [OneNote JavaScript API 参考文档](/javascript/api/onenote)。</span><span class="sxs-lookup"><span data-stu-id="32445-116">For detailed information about the OneNote JavaScript API object model, see the [OneNote JavaScript API reference documentation](/javascript/api/onenote).</span></span>

## <a name="see-also"></a><span data-ttu-id="32445-117">另请参阅</span><span class="sxs-lookup"><span data-stu-id="32445-117">See also</span></span>

* [<span data-ttu-id="32445-118">OneNote 加载项文档</span><span class="sxs-lookup"><span data-stu-id="32445-118">OneNote add-ins documentation</span></span>](../../onenote/index.yml)
* [<span data-ttu-id="32445-119">OneNote 加载项概述</span><span class="sxs-lookup"><span data-stu-id="32445-119">OneNote add-ins overview</span></span>](../../onenote/onenote-add-ins-programming-overview.md)
* [<span data-ttu-id="32445-120">OneNote JavaScript API 参考</span><span class="sxs-lookup"><span data-stu-id="32445-120">OneNote JavaScript API reference</span></span>](/javascript/api/onenote)
* [<span data-ttu-id="32445-121">Office 客户端应用程序和平台的 Office 加载项可用性</span><span class="sxs-lookup"><span data-stu-id="32445-121">Office client application and platform availability for Office Add-ins</span></span>](../../overview/office-add-in-availability.md)
