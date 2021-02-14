---
title: 初学者指南
description: 通过 Office 加载项的学习资源为初学者提供指导的推荐路径。
ms.date: 10/14/2020
ms.custom: scenarios:getting-started
localization_priority: Priority
ms.openlocfilehash: 154d7b5e1a9e135ea583ae6b1afa4ac9e95e9c69
ms.sourcegitcommit: fefc279b85e37463413b6b0e84c880d9ed5d7ac3
ms.translationtype: HT
ms.contentlocale: zh-CN
ms.lasthandoff: 02/12/2021
ms.locfileid: "50234224"
---
# <a name="beginners-guide"></a>初学者指南

想要开始构建自己的跨平台 Office 扩展？ 以下步骤显示了需要先阅读的内容、要安装的工具以及要完成的推荐教程。

> [!NOTE]
> 如果你已熟知如何创建适用于 Office 的 VSTO 加载项，建议直接转到 [VSTO 加载项开发人员指南](learning-path-transition.md)（该文章是本文中信息的超集）。

## <a name="step-0-prerequisites"></a>步骤 0：先决条件

- Office 加载项本质上是嵌入在 Office 中的 Web 应用程序。 因此，你首先应该对 Web 应用程序以及如何在 Web 上托管它们有基本的了解。 Internet、书籍和在线课程提供了有关它的大量信息。 如果你根本不了解 Web 应用程序，那么一个很好的开始方法是在 必应上搜索“什么是 Web 应用程序？”。
- 创建 Office 加载项时将使用的主要编程语言是 JavaScript 或 TypeScript。 可将 TypeScript 视为 JavaScript 的强类型版本。 如果你不熟悉这两种语言，但是你有使用 VBA、VB.Net、C# 的经验，则你可能会发现 TypeScript 更容易学习。 此外，Internet、书籍和在线课程提供了有关这些语言的大量信息。

## <a name="step-1-begin-with-fundamentals"></a>步骤 1：从基础知识开始

我们知道你渴望开始编码，但是在打开 IDE 或代码编辑器之前，你应该先阅读一些有关 Office 加载项的信息。

- [Office 加载项平台概述](office-add-ins.md)：了解什么是 Office Web 加载项以及它们与扩展 Office（如 VSTO 加载项）的旧方法有何区别。
- [开发 Office 加载项](../develop/develop-overview.md)：获取 Office 加载项的开发和生命周期概述，包括工具、创建加载项 UI 以及使用 JavaScript API 与 Office 文档交互。

这些文章中有许多链接，但是如果你是 Office 加载项的初学者，我们建议你在阅读完后返回此处并继续下一部分。

## <a name="step-2-install-tools-and-create-your-first-add-in"></a>步骤 2：安装工具并创建首个加载项

现在，你已有了大致的了解，下面需要深入了解其中一个快速入门。 出于学习平台的目的，我们推荐使用 Excel 快速入门。 我们提供基于 Visual Studio 的版本以及基于 Node.js 和 Visual Studio Code 的版本。

- [Visual Studio](../quickstarts/excel-quickstart-jquery.md?tabs=visualstudio)
- [Node.js 和 Visual Studio Code](../quickstarts/excel-quickstart-jquery.md?tabs=yeomangenerator)

## <a name="step-3-code"></a>步骤 3：代码

你无法通过阅读车主手册学会开车，因此请从此 [Excel 教程](../tutorials/excel-tutorial.md)开始编码吧。 你将使用 Office JavaScript 库和加载项清单中的一些 XML。 无需记住任何内容，因为在后面的步骤中，你将获得关于这两者的更多背景知识。

## <a name="step-4-understand-the-javascript-library"></a>步骤 4：了解 JavaScript 库

首先，通过来自 Microsoft Learn 的本教程大致了解 Office JavaScript 库：[了解 Office JavaScript API](/learn/modules/understand-office-javascript-apis/index)。

然后，使用我们的 [Script Lab 工具](explore-with-script-lab.md)（一种用于运行和探索 API 的沙箱）来探索 Office JavaScript API。

## <a name="step-5-understand-the-manifest"></a>步骤 5：了解清单

在 [Office 加载项 XML 清单](../develop/add-in-manifests.md)中了解加载项清单的用途以及有关其 XML 标记的简介。

## <a name="next-steps"></a>后续步骤

恭喜你完成了初学者的 Office 加载项学习之路！ 以下是进一步探索我们的文档的一些建议：

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