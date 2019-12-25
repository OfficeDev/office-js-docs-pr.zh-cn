---
title: 开发 Office 加载项
description: Office 加载项开发简介。
ms.date: 12/24/2019
localization_priority: Priority
ms.openlocfilehash: 731226883e2bdea4b68d0720042010a0f0117098
ms.sourcegitcommit: 350f5c6954dec3e9384e2030cd3265aaba7ae904
ms.translationtype: HT
ms.contentlocale: zh-CN
ms.lasthandoff: 12/23/2019
ms.locfileid: "40851686"
---
# <a name="develop-office-add-ins"></a><span data-ttu-id="9e5fb-103">开发 Office 加载项</span><span class="sxs-lookup"><span data-stu-id="9e5fb-103">Develop Office Add-ins with Angular</span></span>

> [!TIP]
> <span data-ttu-id="9e5fb-104">阅读本文之前，请查看[构建 Office 加载项](../overview/office-add-ins-fundamentals.md)。</span><span class="sxs-lookup"><span data-stu-id="9e5fb-104">Please review [Building Office Add-ins](../overview/office-add-ins-fundamentals.md) before reading this article.</span></span>

<span data-ttu-id="9e5fb-105">所有 Office 加载项均基于 Office 加载项平台构建。</span><span class="sxs-lookup"><span data-stu-id="9e5fb-105">All Office Add-ins are built upon the Office Add-ins platform.</span></span> <span data-ttu-id="9e5fb-106">它们共享一个可实现某些功能的公共框架。</span><span class="sxs-lookup"><span data-stu-id="9e5fb-106">They share a common framework through which certain capabilities can be implemented.</span></span> <span data-ttu-id="9e5fb-107">无论构建任何加载项，你都需要了解主机和平台可用性、Office JavaScript API 编程模式、如何在清单文件中指定加载项的设置和功能等重要概念。</span><span class="sxs-lookup"><span data-stu-id="9e5fb-107">For any add-in you build, you'll need to understand important concepts like host and platform availability, Office JavaScript API programming patterns, how to specify an add-in's settings and capabilities in the manifest file, and more.</span></span> <span data-ttu-id="9e5fb-108">本文档的“**核心概念**” > “**开发**”部分在此介绍了这类核心开发概念。</span><span class="sxs-lookup"><span data-stu-id="9e5fb-108">Core development concepts like these are covered here in the **Core concepts** > **Develop** section of the documentation.</span></span> <span data-ttu-id="9e5fb-109">在浏览与所构建的加载项（例如 [Excel](../excel/index.md)）相对应的主机特定文档之前，请先查看此处的信息。</span><span class="sxs-lookup"><span data-stu-id="9e5fb-109">Review the information here before exploring the host-specific documentation that corresponds to the add-in you're building (for example, [Excel](../excel/index.md)).</span></span>

> [!NOTE]
> <span data-ttu-id="9e5fb-110">本文档的“**核心概念**” > “**开发**” > “**操作方法**”部分包含侧重于具体开发概念或任务的文章。</span><span class="sxs-lookup"><span data-stu-id="9e5fb-110">The **Core concepts** > **Develop** > **How to** section of this documentation contains articles focused on specific development concepts or tasks.</span></span> <span data-ttu-id="9e5fb-111">例如，你将在此处找到诸如[使用 Visual Studio Code 开发加载项](develop-add-ins-vscode.md)、[随文档自动打开任务窗格](automatically-open-a-task-pane-with-a-document.md)、[创建加载项命令](create-addin-commands.md)以及[打开对话框](dialog-api-in-office-add-ins.md)等任务的信息。</span><span class="sxs-lookup"><span data-stu-id="9e5fb-111">For example, you'll find information there about tasks like [developing add-ins with Visual Studio Code](develop-add-ins-vscode.md), [automatically opening a task pane with a document](automatically-open-a-task-pane-with-a-document.md), [creating add-in commands](create-addin-commands.md), and [opening a dialog box](dialog-api-in-office-add-ins.md).</span></span>

## <a name="next-steps"></a><span data-ttu-id="9e5fb-112">后续步骤</span><span class="sxs-lookup"><span data-stu-id="9e5fb-112">Next steps</span></span>

<span data-ttu-id="9e5fb-113">在熟悉此处介绍的核心概念之后，请浏览与所构建的加载项（例如 [Excel](../excel/index.md)）相对应的主机特定文档。</span><span class="sxs-lookup"><span data-stu-id="9e5fb-113">After you're familiar with the core concepts covered here, explore the host-specific documentation that corresponds to the add-in you're building (for example, [Excel](../excel/index.md)).</span></span> <span data-ttu-id="9e5fb-114">文档中每个主机特定的部分都包含关于为特定 Office 主机构建加载项的具体信息。</span><span class="sxs-lookup"><span data-stu-id="9e5fb-114">Each host-specific section of the documentation contains information specifically about building add-ins for a certain Office host.</span></span>

## <a name="see-also"></a><span data-ttu-id="9e5fb-115">另请参阅</span><span class="sxs-lookup"><span data-stu-id="9e5fb-115">See also</span></span>

- [<span data-ttu-id="9e5fb-116">Office 加载项平台概述</span><span class="sxs-lookup"><span data-stu-id="9e5fb-116">Office Add-ins platform overview</span></span>](../overview/office-add-ins.md)
- [<span data-ttu-id="9e5fb-117">构建 Office 加载项</span><span class="sxs-lookup"><span data-stu-id="9e5fb-117">Building Office Add-ins using Office.js book</span></span>](../overview/office-add-ins-fundamentals.md)
- [<span data-ttu-id="9e5fb-118">Office 加载项的核心概念</span><span class="sxs-lookup"><span data-stu-id="9e5fb-118">Core concepts for Office Add-ins</span></span>](../overview/core-concepts-office-add-ins.md)
- [<span data-ttu-id="9e5fb-119">设计 Office 加载项</span><span class="sxs-lookup"><span data-stu-id="9e5fb-119">Design Office Add-ins</span></span>](../design/add-in-design.md)
- [<span data-ttu-id="9e5fb-120">测试和调试 Office 加载项</span><span class="sxs-lookup"><span data-stu-id="9e5fb-120">Test and debug Office Add-ins</span></span>](../testing/test-debug-office-add-ins.md)
- <span data-ttu-id="9e5fb-121">[发布 Office 加载项](../publish/publish.md)</span><span class="sxs-lookup"><span data-stu-id="9e5fb-121">The [Publish Office Add-ins](../publish/publish.md) wizard appears.</span></span>