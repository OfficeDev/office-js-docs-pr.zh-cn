---
title: 开发 Office 加载项
description: Office 加载项开发简介。
ms.date: 10/14/2020
localization_priority: Priority
ms.openlocfilehash: 03a693e29f5aa52156a2ebfc7b6ed782323bb2b7
ms.sourcegitcommit: ee9e92a968e4ad23f1e371f00d4888e4203ab772
ms.translationtype: HT
ms.contentlocale: zh-CN
ms.lasthandoff: 06/23/2021
ms.locfileid: "53076116"
---
# <a name="develop-office-add-ins"></a>开发 Office 加载项

> [!TIP]
> 阅读本文之前，请查看 [Office 加载项平台概述](../overview/office-add-ins.md)。

所有 Office 加载项均基于 Office 加载项平台构建。 无论构建任何加载项，你都需要了解应用程序和平台可用性、Office JavaScript API 编程模式、如何在清单文件中指定加载项的设置和功能、如何设计 UI 和用户体验等重要概念。 本文档的“**开发生命周期**” > “**开发**”部分在此介绍了这类核心开发概念。 在浏览与所构建的加载项（例如 [Excel](../excel/index.yml)）相对应的应用程序特定文档之前，请先查看此处的信息。

## <a name="creating-an-office-add-in"></a>创建 Office 加载项

你可通过适用于 Office 加载项的 Yeoman 生成器或 Visual Studio 来创建 Office 加载项。

### <a name="yeoman-generator-for-office-add-ins"></a>适用于 Office 加载项的 Yeoman 生成器

[](https://github.com/officedev/generator-office)可用来创建 Node.js Office 加载项项目，而后者可通过 Visual Studio Code 或任何其他编辑器进行管理。 该生成器可创建适合下述任一应用的 Office 加载项：

- Excel
- OneNote
- Outlook
- PowerPoint
- Project
- Word
- Excel 自定义函数

你可选择使用 HTML、CSS 和 JavaScript 创建该项目，也可使用 Angular 或 React 进行创建。 此外，无论选择哪种框架，都可在 JavaScript 和 Typescript 之间进行选择。 有关使用 Yeoman 生成器创建加载项的详细信息，请参阅[使用 Visual Studio Code 开发 Office 加载项](../develop/develop-add-ins-vscode.md)。

### <a name="visual-studio"></a>Visual Studio

Visual Studio 可用于创建适用于 Excel、Outlook、Word 和 PowerPoint 的 Office 加载项。 Office 加载项项目是作为 Visual Studio 解决方案的一部分创建的，它使用 HTML、CSS 和 JavaScript。 有关使用 Visual Studio 创建加载项的详细信息，请参阅[使用 Visual Studio 开发 Office 加载项](../develop/develop-add-ins-visual-studio.md)。

[!include[Yeoman vs Visual Studio comparison](../includes/yeoman-generator-recommendation.md)]

## <a name="understanding-the-two-parts-of-an-office-add-in"></a>了解 Office 加载项的两个部分

Office 加载项由两部分组成：

- 加载项清单（XML 文件），它定义了加载项的设置和功能。

- Web 应用程序，它定义了加载项组件的 UI 和功能，例如任务窗格、内容加载项和对话框。

这些 Web 应用程序可使用 Office JavaScript API 来与其中正在运行加载项的 Office 文档中的内容进行交互。你的加载项还可执行 Web 应用程序通常可实现的其他操作，例如调用外部 Web 服务和简化用户身份验证等等。

### <a name="defining-an-add-ins-settings-and-capabilities"></a>定义加载项的设置和功能

Office 加载项的清单是一个 XML 文件，它定义了加载项的设置和功能。 你需配置清单来指定如下内容：

- 描述加载项的元数据（例如 ID、版本、说明、显示名称和默认区域设置）。
- 将在其中运行加载项的 Office 应用程序。
- 加载项所需的权限。
- 加载项与 Office 集成的方式，包括与加载项创建的自定义选项卡和功能区按钮等自定义 UI 的集成。
- 加载项对品牌和命令图标使用的图像的位置。
- 加载项的尺寸（例如内容加载项的尺寸、Outlook 加载项请求的高度）。
- 指定何时在消息或约会上下文中激活加载项的规则（仅限 Outlook 加载项）。

有关清单的详细信息，请参阅 [Office 加载项 XML 清单](add-in-manifests.md)。

### <a name="interacting-with-content-in-an-office-document"></a>与 Office 文档中的内容交互

Office 加载项可使用 Office JavaScript API 来与其中在运行加载项的 Office 文档中的内容进行交互。

#### <a name="accessing-the-office-javascript-api-library"></a>访问 Office JavaScript API 库

[!include[information about accessing the Office JS API library](../includes/office-js-access-library.md)]

#### <a name="api-models"></a>API 模型

[!include[information about the Office JS API models](../includes/office-js-api-models.md)]

#### <a name="api-requirement-sets"></a>API 要求集

[!include[information about the Office JS API requirement sets](../includes/office-js-requirement-sets.md)]

#### <a name="exploring-apis-with-script-lab"></a>使用 Script Lab 了解 API

Script Lab 是一款加载项，在 Excel 或 Word 等 Office 程序中工作时，你可用它来了解 Office JavaScript API 和运行代码片段。 该工具通过 [AppSource](https://appsource.microsoft.com/product/office/WA104380862) 免费提供，随附在你的开发工具包中，在你建立希望加载项中拥有的功能原型和验证该功能时非常有用。 在 Script Lab 中，你可访问内置示例库以快速试用 API，甚至还可将示例用作你自己的代码的起点。

下面时长一分钟的视频展示了 Script Lab 的实际运行情况。

[![显示 Script Lab 在 Excel、Word 和 PowerPoint 中运行的短视频。](../images/screenshot-wide-youtube.png 'Script Lab 预览视频。')](https://aka.ms/scriptlabvideo)

有关 Script Lab 的详细信息，请参阅[使用 Script Lab 了解 Office JavaScript API](../overview/explore-with-script-lab.md)。

## <a name="extending-the-office-ui"></a>扩展 Office UI

Office 加载项可使用加载项命令和 HTML 容器（如任务窗格、内容加载项或对话框）来扩展 Office UI。

- [加载项命令](../design/add-in-commands.md)可用于向 Office 中的默认功能区添加自定义选项卡、按钮和菜单，或者扩展当用户右键单击 Office 文档中的文本或 Excel 中的对象时显示的默认上下文菜单。 当用户选择加载项命令时，他们将启动该加载项命令指定的任务，例如运行 JavaScript 代码、打开任务窗格或启动对话框。

- [任务窗格](../design/task-pane-add-ins.md)、[内容加载项](../design/content-add-ins.md)和[对话框](../design/dialog-boxes.md)等 HTML 容器可用于显示自定义 UI 和探索 Office 应用程序中的附加功能。 每个任务窗格、内容加载项或对话框的内容和功能派生自你指定的网页。 这些网页可使用 Office JavaScript API 来与其中正在运行加载项的 Office 文档中的内容进行交互，还可执行网页通常可实现的其他操作，例如调用外部 Web 服务和简化用户身份验证等等。

下图显示功能区中有一个加载项命令、文档右侧有一个任务窗格，且文档上方有一个对话框或内容加载项。

![显示 Office 文档中的功能区内加载项命令、任务窗格、对话框/内容加载项的图表。](../images/add-in-ui-elements.png)

要详细了解如何扩展 Office UI 和设计加载项的 UX，请参阅 [Office 加载项的 Office UI 元素](../design/interface-elements.md)。

## <a name="next-steps"></a>后续步骤

本文概述了创建 Office 加载项的不同方法、介绍了加载项扩展 Office UI 的方法，描述了 API 集,介绍了 Script Lab（一种用来于了解 Office JavaScript API 和建立加载项功能原型的宝贵工具）。现在，你了解这一介绍性信息，请考虑沿着以下学习路径继续你的 Office 加载项之旅。

### <a name="create-an-office-add-in"></a>创建 Office 加载项

可完成 [5 分钟快速入门](../index.yml)，快速创建适合 Excel、OneNote、Outlook、PowerPoint、Project 或 Word 的基本加载项。 如果你之前已完成快速入门，并且想要创建更复杂一些的加载项，请尝试本[教程](../index.yml)。

### <a name="learn-more"></a>了解详细信息

查看此文档，详细了解如何开发、测试和发布 Office 加载项。

> [!TIP]
> 对于你构建的任何加载项，都可查看本文档的[开发生命周期](../overview/core-concepts-office-add-ins.md)部分中的信息，还可查看与你要构建的加载项类型（例如 [Excel](../excel/index.yml)）相对应的应用程序特定部分中的信息。

## <a name="see-also"></a>另请参阅

- [Office 加载项平台概述](../overview/office-add-ins.md)
- [了解 Microsoft 365 开发人员计划](https://developer.microsoft.com/microsoft-365/dev-program)
- [设计 Office 加载项](../design/add-in-design.md)
- [测试和调试 Office 加载项](../testing/test-debug-office-add-ins.md)
- [发布 Office 加载项](../publish/publish.md)
