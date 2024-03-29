---
title: VSTO 加载项开发人员指南
description: 资深 VSTO 加载项开发人员了解 Office Web 加载项资源的建议路径。
ms.date: 10/14/2020
ms.custom: scenarios:getting-started
ms.localizationpriority: high
ms.openlocfilehash: bc27177c67028e57030c9baed6b416d0c57c77d1
ms.sourcegitcommit: 3abcf7046446e7b02679c79d9054843088312200
ms.translationtype: MT
ms.contentlocale: zh-CN
ms.lasthandoff: 11/02/2022
ms.locfileid: "68810125"
---
# <a name="vsto-add-in-developers-guide"></a>VSTO 加载项开发人员指南

因此，你已经为在 Windows 上运行的 Office 应用程序制作了一些 VSTO 加载项，现在你正在探索扩展将在 Windows、Mac 和 Office 套件的 Web 浏览器版本上运行的 Office 的新方法：Office Web 外接程序。

Your understanding of the object models for the Excel, Word, and the other Office applications will be a huge help because the object models in Office Web Add-ins follow similar patterns. But there are going to be some challenges:

- 你将使用其他语言（JavaScript 或 TypeScript）而不是 C＃或 Visual Basic .NET。 （还有一种方法，如下所述，可以重复使用 Web 加载项中存在的代码。）
- Office Web 加载项的部署方式不同于 VSTO 加载项。
- Office Web 加载项是在 Office 应用程序中嵌入的简化浏览器窗口中运行的 Web 应用程序，因此需要对 Web 应用程序以及如何在Web服务器或云帐户上托管有基本的了解。 

出于以上原因，本文的大部分内容都向完整的 Office 扩展初学者介绍了我们的学习路径：[入门指南](learning-path-beginner.md)。 我们添加了一些其他学习资源，以帮助 VSTO 加载项开发人员利用他们的经验，并帮助他们重用现有代码。

## <a name="step-0-prerequisites"></a>步骤 0：先决条件

- Office Web 加载项（也称为 Office 加载项）本质上是嵌入在 Office 中的 Web 应用程序。 因此，你首先应该对 Web 应用程序以及如何在 Web 上托管它们有基本的了解。 Internet、书籍和在线课程提供了有关它的大量信息。 如果你根本不了解 Web 应用程序，那么一个很好的开始方法是在 必应上搜索“什么是 Web 应用程序？”。
- 创建 Office 加载项将使用的主要编程语言是 JavaScript 或 TypeScript。 可将 TypeScript 视为 JavaScript 的强类型版本。 如果你不熟悉这两种语言，但是你有使用 VBA、VB.Net、C# 的经验，则你可能会发现 TypeScript 更容易学习。 此外，Internet、书籍和在线课程提供了有关这些语言的大量信息。

## <a name="step-1-begin-with-fundamentals"></a>步骤 1：从基础知识开始

我们知道你渴望开始编码，但是在打开 IDE 或代码编辑器之前，你应该先阅读一些有关 Office 加载项的信息。

- [Office 加载项平台概述](office-add-ins.md)：了解什么是 Office Web 加载项以及它们与扩展 Office（如 VSTO 加载项）的旧方法有何区别。
- [开发 Office 加载项](../develop/develop-overview.md)：获取 Office 加载项的开发和生命周期概述，包括工具、创建加载项 UI 以及使用 JavaScript API 与 Office 文档交互。

这些文章中有许多链接，但是如果你正在过渡至 Office Web 加载项的初学者，我们建议你在阅读完后返回此处并继续下一部分。

## <a name="step-2-install-tools-and-create-your-first-add-in"></a>步骤 2：安装工具并创建首个加载项

现在，你已有了大致的了解，下面需要深入了解其中一个快速入门。 出于学习平台的目的，我们推荐使用 Excel 快速入门。 一个版本基于 Visual Studio，另一个版本基于 Node.js 和 Visual Studio Code。 如果正在从 VSTO 加载项转换，可能会发现 Visual Studio 版本更易于使用。

- [Visual Studio](../quickstarts/excel-quickstart-jquery.md?tabs=visualstudio)
- [Node.js 和 Visual Studio Code](../quickstarts/excel-quickstart-jquery.md?tabs=yeomangenerator)

## <a name="step-3-code"></a>步骤 3：代码

你无法通过阅读车主手册学会开车，因此请从此 [Excel 教程](../tutorials/excel-tutorial.md)开始编码吧。 你将使用 Office JavaScript 库和加载项清单中的一些 XML。 无需记住任何内容，因为在后面的步骤中，你将获得关于这两者的更多背景知识。

## <a name="step-4-understand-the-javascript-library"></a>步骤 4：了解 JavaScript 库

通过本教程获取 Office JavaScript 库的大图，请参阅 Microsoft Learn 培训： [了解 Office JavaScript API](/training/modules/intro-office-add-ins/3-apis)。

然后，使用 [Script Lab 工具](explore-with-script-lab.md)（一种用于运行和探索 API 的沙箱）来探索 Office JavaScript API。

### <a name="special-resource-for-vsto-add-in-developers"></a>适用于 VSTO 加载项开发人员的特殊支援

这里将介绍如何查看示例加载项、[Excel 加载项 JavaScript SalesTracker](https://github.com/OfficeDev/Excel-Add-in-JavaScript-SalesTracker)。 创建的目的是为了突出显示 VSTO 加载项和 Office Web 加载项之间的异同，并且示例的自述文件指出了比较的重点。

## <a name="step-5-understand-the-manifest"></a>步骤 5：了解清单

在 [Office 加载项 XML 清单](../develop/add-in-manifests.md)中了解 web 加载项清单的用途以及有关其 XML 标记的简介。

## <a name="step-6-for-vsto-developers-only-reuse-your-vsto-code"></a>步骤 6（仅适用于 VSTO 开发人员）：重复使用 VSTO 代码

可以在 Office Web 加载项中重复使用某些 VSTO 加载项代码，方法是将其移到服务器上 Web 应用程序的后端，然后将其作为 Web API 供 JavaScript 或 TypeScript 使用。 有关指南，参见 [教程：使用共享代码库在 VSTO 加载项与 Office 加载项之间共享代码](../tutorials/migrate-vsto-to-office-add-in-shared-code-library-tutorial.md)。

## <a name="next-steps"></a>后续步骤

恭喜你完成了 VSTO 加载项的 Office Web 加载项学习之路！ 以下是进一步探索我们的文档的一些建议：

- 其他 Office 应用程序的教程或快速入门：

  - [OneNote 快速入门](../quickstarts/onenote-quickstart.md)
  - [Outlook 教程](/outlook/add-ins/addin-tutorial)
  - [PowerPoint 教程](../tutorials/powerpoint-tutorial.md)
  - [Project 快速入门](../quickstarts/project-quickstart.md)
  - [Word 教程](../tutorials/word-tutorial.md)

- 其他重要主题：

  - [开发 Office 加载项](../develop/develop-overview.md)
  - [Office 加载项开发最佳做法](../concepts/add-in-development-best-practices.md)
  - [设计 Office 加载项](../design/add-in-design.md)
  - [测试和调试 Office 加载项](../testing/test-debug-office-add-ins.md)
  - [部署和发布 Office 加载项](../publish/publish.md)
  - [资源](../resources/resources-links-help.md)
  - [了解 Microsoft 365 开发人员计划](https://developer.microsoft.com/microsoft-365/dev-program)
