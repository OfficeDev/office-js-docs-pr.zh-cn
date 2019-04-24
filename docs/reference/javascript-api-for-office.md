---
title: 适用于 Office 的 JavaScript API
description: ''
ms.date: 03/19/2019
localization_priority: Priority
ms.openlocfilehash: c8b33bbf9d0107786c0272410c59b1a3fe998cba
ms.sourcegitcommit: 9e7b4daa8d76c710b9d9dd4ae2e3c45e8fe07127
ms.translationtype: HT
ms.contentlocale: zh-CN
ms.lasthandoff: 04/24/2019
ms.locfileid: "32450714"
---
# <a name="javascript-api-for-office"></a><span data-ttu-id="5a88d-102">适用于 Office 的 JavaScript API</span><span class="sxs-lookup"><span data-stu-id="5a88d-102">JavaScript API for Office</span></span>

<span data-ttu-id="5a88d-103">借助适用于 Office 的 JavaScript API，您可以创建可与 Office 主机应用程序中的对象模型进行交互的 Web 应用程序。</span><span class="sxs-lookup"><span data-stu-id="5a88d-103">The JavaScript API for Office enables you to create web applications that interact with the object models in Office host applications.</span></span> <span data-ttu-id="5a88d-104">你的应用程序将引用 office.js 库中，该库是一个脚本加载程序。</span><span class="sxs-lookup"><span data-stu-id="5a88d-104">Your application will reference the office.js library, which is a script loader.</span></span> <span data-ttu-id="5a88d-105">Office.js 库加载适用于正在运行外接程序的 Office 应用程序的对象模型。</span><span class="sxs-lookup"><span data-stu-id="5a88d-105">The office.js library loads the object models that are applicable to the Office application that is running the add-in.</span></span> <span data-ttu-id="5a88d-106">你可以使用以下 JavaScript 对象模型：</span><span class="sxs-lookup"><span data-stu-id="5a88d-106">You can use the following JavaScript object models:</span></span>

- <span data-ttu-id="5a88d-107">**公用 API** - 与 **Office 2013** 一起引入的 API。</span><span class="sxs-lookup"><span data-stu-id="5a88d-107">**Common APIs** - APIs that were introduced with **Office 2013**.</span></span> <span data-ttu-id="5a88d-108">这为**所有 Office 主机应用程序**加载 API，并将外接程序应用程序与 Office 客户端应用程序相连接。</span><span class="sxs-lookup"><span data-stu-id="5a88d-108">This is loaded for **all Office host applications** and connects your add-in application with the Office client application.</span></span> <span data-ttu-id="5a88d-109">对象模型包含特定于 Office 客户端的 API 以及适用于多个 Office 客户端主机应用程序的 API。</span><span class="sxs-lookup"><span data-stu-id="5a88d-109">The object model contains APIs that are specific to Office clients, and APIs that are applicable to multiple Office client host applications.</span></span> <span data-ttu-id="5a88d-110">所有这些内容位于**通用 API** 下。</span><span class="sxs-lookup"><span data-stu-id="5a88d-110">All of this content is under **Common API**.</span></span> <span data-ttu-id="5a88d-111">此对象模型使用回调。</span><span class="sxs-lookup"><span data-stu-id="5a88d-111">This object model uses callbacks.</span></span> 

  <span data-ttu-id="5a88d-112">**Outlook** 还使用通用 API 语法。</span><span class="sxs-lookup"><span data-stu-id="5a88d-112">**Outlook** also uses the Common API syntax.</span></span> <span data-ttu-id="5a88d-113">代码中的别名 Office 下的全部内容包含可以用于编写与 Office 文档、工作簿、演示文稿、邮件项以及 Office 加载项中的项目中的内容交互的脚本的对象。如果加载项面向 Office 2013 及更高版本，则必须使用这些通用 API。</span><span class="sxs-lookup"><span data-stu-id="5a88d-113">Everything under the alias Office contains objects you can use to write scripts that interact with content in Office documents, worksheets, presentations, mail items, and projects from your Office Add-ins. You must use these Common APIs if your add-in will target Office 2013 and later.</span></span> <span data-ttu-id="5a88d-114">此对象模型使用回调。</span><span class="sxs-lookup"><span data-stu-id="5a88d-114">This object model uses callbacks.</span></span>

- <span data-ttu-id="5a88d-115">**特定于主机的 API** - 与 **Office 2016** 一起引入的 API。</span><span class="sxs-lookup"><span data-stu-id="5a88d-115">**Host-specific APIs** - APIs that were introduced with **Office 2016**.</span></span> <span data-ttu-id="5a88d-116">此对象模型提供特定于主机的强类型对象，这些对象对应于使用 Office 客户端时所看到的熟悉对象，并表示 Office JavaScript API 的未来。</span><span class="sxs-lookup"><span data-stu-id="5a88d-116">This object model provides host-specific strongly-typed objects that correspond to familiar objects that you see when you use Office clients, and represents the future of Office JavaScript APIs.</span></span> <span data-ttu-id="5a88d-117">特定于主机的 API 目前包括 Word JavaScript API 和 Excel JavaScript API。</span><span class="sxs-lookup"><span data-stu-id="5a88d-117">The host-specific APIs currently include the Word JavaScript API and the Excel JavaScript API.</span></span>

## <a name="supported-host-applications"></a><span data-ttu-id="5a88d-118">支持的主机应用程序</span><span class="sxs-lookup"><span data-stu-id="5a88d-118">Supported host applications</span></span>

- [<span data-ttu-id="5a88d-119">Excel</span><span class="sxs-lookup"><span data-stu-id="5a88d-119">Excel</span></span>](overview/excel-add-ins-reference-overview.md)
- [<span data-ttu-id="5a88d-120">OneNote</span><span class="sxs-lookup"><span data-stu-id="5a88d-120">OneNote</span></span>](overview/onenote-add-ins-javascript-reference.md)
- [<span data-ttu-id="5a88d-121">Outlook</span><span class="sxs-lookup"><span data-stu-id="5a88d-121">Outlook</span></span>](requirement-sets/outlook-api-requirement-sets.md)
- [<span data-ttu-id="5a88d-122">Visio</span><span class="sxs-lookup"><span data-stu-id="5a88d-122">Visio</span></span>](overview/visio-javascript-reference-overview.md)
- [<span data-ttu-id="5a88d-123">Word</span><span class="sxs-lookup"><span data-stu-id="5a88d-123">Word</span></span>](overview/word-add-ins-reference-overview.md)
- [<span data-ttu-id="5a88d-124">通用 API</span><span class="sxs-lookup"><span data-stu-id="5a88d-124">Common API</span></span>](requirement-sets/office-add-in-requirement-sets.md)

> [!NOTE] 
> <span data-ttu-id="5a88d-125">[PowerPoint 和 Project](requirement-sets/powerpoint-and-project-note.md) 支持通过 JavaScript API 创建的加载项。</span><span class="sxs-lookup"><span data-stu-id="5a88d-125">[PowerPoint and Project](requirement-sets/powerpoint-and-project-note.md) support add-ins made with the JavaScript API.</span></span> <span data-ttu-id="5a88d-126">但是，它们当前没有特定于主机的 API。</span><span class="sxs-lookup"><span data-stu-id="5a88d-126">However, they currently do not have host-specific APIs.</span></span> <span data-ttu-id="5a88d-127">你可以通过通用 API 与这些主机交互。</span><span class="sxs-lookup"><span data-stu-id="5a88d-127">You interact with these hosts through the Common API.</span></span>

<span data-ttu-id="5a88d-128">了解有关[支持的主机和其他要求](../concepts/requirements-for-running-office-add-ins.md)的详细信息。</span><span class="sxs-lookup"><span data-stu-id="5a88d-128">Learn more about [supported hosts and other requirements](../concepts/requirements-for-running-office-add-ins.md).</span></span>

## <a name="open-api-specifications"></a><span data-ttu-id="5a88d-129">开放 API 规范</span><span class="sxs-lookup"><span data-stu-id="5a88d-129">Open API specifications</span></span>

<span data-ttu-id="5a88d-p106">在我们设计和开发新的 API 以用于 Office 外接程序时，我们将使它们适用于[开放 API 规范](openspec.md)页的反馈。了解管道中的新增功能，并提供您对我们的设计规范的宝贵意见。</span><span class="sxs-lookup"><span data-stu-id="5a88d-p106">As we design and develop new APIs for Office Add-ins, we'll make them available for your feedback on our [Open API specifications](openspec.md) page. Find out what new features are in the pipeline, and provide your input on our design specifications.</span></span>

## <a name="see-also"></a><span data-ttu-id="5a88d-132">另请参阅</span><span class="sxs-lookup"><span data-stu-id="5a88d-132">See also</span></span>

- [<span data-ttu-id="5a88d-133">Office JavaScript API 参考</span><span class="sxs-lookup"><span data-stu-id="5a88d-133">Office JavaScript API reference</span></span>](/javascript/api/overview/office)
