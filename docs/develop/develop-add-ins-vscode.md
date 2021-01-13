---
title: 使用 Visual Studio Code 开发 Office 加载项
description: 如何使用 Visual Studio Code 开发 Office 加载项。
ms.date: 10/14/2020
localization_priority: Priority
ms.openlocfilehash: 3f2a95e8ba08fa91b5f10707459ea87ab6c857e3
ms.sourcegitcommit: d28392721958555d6edea48cea000470bd27fcf7
ms.translationtype: HT
ms.contentlocale: zh-CN
ms.lasthandoff: 01/13/2021
ms.locfileid: "49839948"
---
# <a name="develop-office-add-ins-with-visual-studio-code"></a><span data-ttu-id="309e9-103">使用 Visual Studio Code 开发 Office 加载项</span><span class="sxs-lookup"><span data-stu-id="309e9-103">Develop Office Add-ins with Visual Studio Code</span></span>

<span data-ttu-id="309e9-104">本文介绍如何使用 [Visual Studio Code (VS Code)](https://code.visualstudio.com) 开发 Office 加载项。</span><span class="sxs-lookup"><span data-stu-id="309e9-104">This article describes how to use [Visual Studio Code (VS Code)](https://code.visualstudio.com) to develop an Office Add-in.</span></span>

> [!NOTE]
> <span data-ttu-id="309e9-105">要了解如何使用 Visual Studio 创建 Office 加载项，请参阅[使用 Visual Studio 开发 Office 加载项](develop-add-ins-visual-studio.md)。</span><span class="sxs-lookup"><span data-stu-id="309e9-105">For information about using Visual Studio to create an Office Add-in, see [Develop Office Add-ins with Visual Studio](develop-add-ins-visual-studio.md).</span></span>

## <a name="prerequisites"></a><span data-ttu-id="309e9-106">先决条件</span><span class="sxs-lookup"><span data-stu-id="309e9-106">Prerequisites</span></span>

- [<span data-ttu-id="309e9-107">Visual Studio Code</span><span class="sxs-lookup"><span data-stu-id="309e9-107">Visual Studio Code</span></span>](https://code.visualstudio.com/)

[!include[Yeoman generator prerequisites](../includes/quickstart-yo-prerequisites.md)]

## <a name="create-the-add-in-project-using-the-yeoman-generator"></a><span data-ttu-id="309e9-108">使用 Yeoman 生成器创建加载项项目</span><span class="sxs-lookup"><span data-stu-id="309e9-108">Create the add-in project using the Yeoman generator</span></span>

<span data-ttu-id="309e9-109">如果你正在将 VS Code 用作集成开发环境 (IDE)，则应使用[适用于 Office 加载项的 Yeoman 生成器](https://github.com/OfficeDev/generator-office)来创建 Office 加载项项目。Yeoman 生成器会创建一个 Node.js 项目，它可通过 VS Code 或任何其他编辑器进行管理。</span><span class="sxs-lookup"><span data-stu-id="309e9-109">If you're using VS Code as your integrated development environment (IDE), you should create the Office Add-in project with the [Yeoman generator for Office Add-ins](https://github.com/OfficeDev/generator-office). The Yeoman generator creates a Node.js project that can be managed with VS Code or any other editor.</span></span> 

<span data-ttu-id="309e9-110">要使用 Yeoman 生成器创建 Office 加载项，请按照 [5 分钟快速入门](../index.yml)中与你要创建的加载项类型相对应的说明进行操作。</span><span class="sxs-lookup"><span data-stu-id="309e9-110">To create an Office Add-in with the Yeoman generator, follow instructions in the [5-minute quick start](../index.yml) that corresponds to the type of add-in you'd like to create.</span></span>

## <a name="develop-the-add-in-using-vs-code"></a><span data-ttu-id="309e9-111">使用 VS Code 开发加载项</span><span class="sxs-lookup"><span data-stu-id="309e9-111">Develop the add-in using VS Code</span></span>

<span data-ttu-id="309e9-112">在 Yeoman 生成器完成加载项项目的创建后，请使用 VS Code 打开项目的根文件夹。</span><span class="sxs-lookup"><span data-stu-id="309e9-112">When the Yeoman generator finishes creating the add-in project, open the root folder of the project with VS Code.</span></span> 

> [!TIP]
> <span data-ttu-id="309e9-113">在 Windows 上，可通过命令行导航到项目的根目录，然后输入 `code .`在 VS Code 中打开该文件夹。</span><span class="sxs-lookup"><span data-stu-id="309e9-113">On Windows, you can navigate to the root directory of the project via the command line and then enter `code .` to open that folder in VS Code.</span></span> <span data-ttu-id="309e9-114">在 Mac 上，需要先[将 `code` 命令添加到路径中](https://code.visualstudio.com/docs/setup/mac#_launching-from-the-command-line)，然后才可使用该命令在 VS Code 中打开项目文件夹。</span><span class="sxs-lookup"><span data-stu-id="309e9-114">On Mac, you'll need to [add the `code` command to the path](https://code.visualstudio.com/docs/setup/mac#_launching-from-the-command-line) before you can use that command to open the project folder in VS Code.</span></span>

<span data-ttu-id="309e9-115">Yeoman 生成器会创建一个功能受限的基本加载项。</span><span class="sxs-lookup"><span data-stu-id="309e9-115">The Yeoman generator creates a basic add-in with limited functionality.</span></span> <span data-ttu-id="309e9-116">你可通过在 VS Code 中编辑[清单](add-in-manifests.md)HTML、JavaScript/TypeScript 和 CSS 文件，自定义该加载项。</span><span class="sxs-lookup"><span data-stu-id="309e9-116">You can customize the add-in by editing the [manifest](add-in-manifests.md), HTML, JavaScript or TypeScript, and CSS files in VS Code.</span></span> <span data-ttu-id="309e9-117">要简要了解 Yeoman 生成器创建的加载项项目中的项目结构和文件，请查看 [5 分钟快速入门](../index.yml)中与你创建的加载项类型相对应的 Yeoman 生成器指南。</span><span class="sxs-lookup"><span data-stu-id="309e9-117">For a high-level description of the project structure and files in the add-in project that the Yeoman generator creates, see the Yeoman generator guidance within the [5-minute quick start](../index.yml) that corresponds to the type of add-in you've created.</span></span>

## <a name="test-and-debug-the-add-in"></a><span data-ttu-id="309e9-118">测试和调试加载项</span><span class="sxs-lookup"><span data-stu-id="309e9-118">Test and debug the add-in</span></span>

<span data-ttu-id="309e9-119">用于测试、调试和故障排除 Office 加载项的方法因平台而异。</span><span class="sxs-lookup"><span data-stu-id="309e9-119">Methods for testing, debugging, and troubleshooting Office Add-ins vary by platform.</span></span> <span data-ttu-id="309e9-120">有关详细信息，请参阅 [测试和调试 Office 加载项](../testing/test-debug-office-add-ins.md)。</span><span class="sxs-lookup"><span data-stu-id="309e9-120">For more information, see [Test and debug Office Add-ins](../testing/test-debug-office-add-ins.md).</span></span>

## <a name="publish-the-add-in"></a><span data-ttu-id="309e9-121">发布加载项</span><span class="sxs-lookup"><span data-stu-id="309e9-121">Publish the add-in</span></span>

[!include[instructions for publishing an Office Add-in](../includes/publish-add-in.md)]

## <a name="see-also"></a><span data-ttu-id="309e9-122">另请参阅</span><span class="sxs-lookup"><span data-stu-id="309e9-122">See also</span></span>

- [<span data-ttu-id="309e9-123">Office 加载项的核心概念</span><span class="sxs-lookup"><span data-stu-id="309e9-123">Core concepts for Office Add-ins</span></span>](../overview/core-concepts-office-add-ins.md)
- [<span data-ttu-id="309e9-124">开发 Office 加载项</span><span class="sxs-lookup"><span data-stu-id="309e9-124">Develop Office Add-ins</span></span>](../develop/develop-overview.md)
- [<span data-ttu-id="309e9-125">设计 Office 加载项</span><span class="sxs-lookup"><span data-stu-id="309e9-125">Design Office Add-ins</span></span>](../design/add-in-design.md)
- [<span data-ttu-id="309e9-126">测试和调试 Office 加载项</span><span class="sxs-lookup"><span data-stu-id="309e9-126">Test and debug Office Add-ins</span></span>](../testing/test-debug-office-add-ins.md)
- [<span data-ttu-id="309e9-127">发布 Office 加载项</span><span class="sxs-lookup"><span data-stu-id="309e9-127">Publish Office Add-ins</span></span>](../publish/publish.md)