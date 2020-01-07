---
title: 构建 Office 加载项
description: Office 加载项开发简介。
ms.date: 12/24/2019
localization_priority: Priority
ms.openlocfilehash: e0deeebb3a1c8761217a9fe33a3ef04a945b2cff
ms.sourcegitcommit: d5ac9284d1e96dc91a9168d7641e44d88535e1a7
ms.translationtype: HT
ms.contentlocale: zh-CN
ms.lasthandoff: 12/31/2019
ms.locfileid: "40915019"
---
# <a name="building-office-add-ins"></a>构建 Office 加载项

> [!TIP]
> 阅读本文之前，请查看 [Office 加载项平台概述](office-add-ins.md)。

Office 加载项可扩展 Office 应用程序的 UI 和功能，并与 Office 文档中的内容交互。 你将使用熟悉的 Web 技术创建 Office 加载项来扩展 Word、Excel、PowerPoint、OneNote、Project 或 Outlook 并与之交互。 你构建的加载项可跨多个平台在 Office 中运行，包括 Windows、Mac、iPad 和在浏览器中。 本文简要介绍了如何开发 Office 加载项。

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

[!include[Yeoman vs Visual Studio comparision](../includes/yeoman-generator-recommendation.md)]

## <a name="exploring-apis-with-script-lab"></a>使用 Script Lab 了解 API

Script Lab 是一款加载项，在 Excel 或 Word 等 Office 程序中工作时，你可用它来了解 Office JavaScript API 和运行代码片段。 该工具通过 [AppSource](https://appsource.microsoft.com/product/office/WA104380862) 免费提供，随附在你的开发工具包中，在你建立希望加载项中拥有的功能原型和验证该功能时非常有用。 在 Script Lab 中，你可访问内置示例库以快速试用 API，甚至还可将示例用作你自己的代码的起点。 

下面时长一分钟的视频展示了 Script Lab 的实际运行情况。

[![展示 Script Lab 在 Excel、Word 和 PowerPoint 中运行的预览视频。](../images/screenshot-wide-youtube.png 'Script Lab 预览视频')](https://aka.ms/scriptlabvideo)

有关 Script Lab 的详细信息，请参阅[使用 Script Lab 了解 Office JavaScript API](../overview/explore-with-script-lab.md)。

## <a name="extending-the-office-ui"></a>扩展 Office UI

Office 加载项可使用加载项命令和 HTML 容器（如任务窗格、内容加载项或对话框）来扩展 Office UI。

- [加载项命令](../design/add-in-commands.md)可用于向 Office 中的默认功能区添加自定义选项卡、按钮和菜单，或者扩展当用户右键单击 Office 文档中的文本或 Excel 中的对象时显示的默认上下文菜单。 当用户选择加载项命令时，他们将启动该加载项命令指定的任务，例如运行 JavaScript 代码、打开任务窗格或启动对话框。

- [任务窗格](../design/task-pane-add-ins.md)、[内容加载项](../design/content-add-ins.md)和[对话框](../design/dialog-boxes.md)等 HTML 容器可用于显示自定义 UI 和探索 Office 应用程序中的附加功能。 每个任务窗格、内容加载项或对话框的内容和功能派生自你指定的网页。 这些网页可使用 Office JavaScript API 来与其中正在运行加载项的 Office 文档中的内容进行交互，还可执行网页通常可实现的其他操作，例如调用外部 Web 服务和简化用户身份验证等等。

下图显示功能区中有一个加载项命令、文档右侧有一个任务窗格，且文档上方有一个对话框或内容加载项。

![显示 Office 文档中的功能区内加载项命令、任务窗格和对话框的图像](../images/add-in-ui-elements.png)

要详细了解如何扩展 Office UI，请参阅 [Office 加载项的 Office UI 元素](../design/interface-elements.md)。

## <a name="core-development-concepts"></a>核心开发概念 

Office 加载项由两部分组成：

- 加载项清单（XML 文件），它定义了加载项的设置和功能。

- Web 应用程序，它定义了加载项组件的 UI 和功能，例如任务窗格、内容加载项和对话框。

Web 应用程序使用 Office JavaScript API 来与其中在运行加载项的 Office 文档中的内容进行交互。 你的加载项还可执行 Web 应用程序通常可实现的其他操作，例如调用外部 Web 服务和简化用户身份验证等等。

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

#### <a name="accessing-the-office-javascript-library"></a>访问 Office JavaScript 库

可通过 Office JS 内容交付网络 (CDN) 访问 Office JavaScript 库：`https://appsforoffice.microsoft.com/lib/1/hosted/Office.js` 要在任何加载项的网页中使用 Office JavaScript API，必须在页面的 `<head>` 标记中的 `<script>` 标记内引用 CDN。

```html
<head>
    ...
    <script src="https://appsforoffice.microsoft.com/lib/1/hosted/Office.js" type="text/javascript"></script>
</head>
```

> [!NOTE]
> 要使用预览版 API，请参考 CDN 上的 Office JavaScript 库预览版：https://appsforoffice.microsoft.com/lib/beta/hosted/office.js。

要详细了解如何访问 Office JavaScript 库（包括如何获取 IntelliSense），请参阅[通过 JavaScript API for Office 的内容交付网络 (CDN) 引用该库](../develop/referencing-the-javascript-api-for-office-library-from-its-cdn.md)。

#### <a name="api-models"></a>API 模型

Office JavaScript API 包含两种不同的模型：

- **主机特定的** API 提供了强类型对象，它可用于与特定 Office 应用程序的本机对象进行交互。 例如，可使用 Excel JavaScript API 来访问工作表、区域、表格和图表等。 主机特定的 API 当前可用于 [Excel](../reference/overview/excel-add-ins-reference-overview.md)、[Word](../reference/overview/word-add-ins-reference-overview.md) 和 [OneNote](../reference/overview/onenote-add-ins-javascript-reference.md)。 此 API 模型使用的是[承诺](https://developer.mozilla.org/docs/Web/JavaScript/Reference/Global_Objects/Promise)，你可用它在你发送给 Office 主机的每个请求中指定多个操作。 通过此方式批量处理操作，可大幅提升 Web 应用程序上的 Office 中的性能。 主机特定的 API 是随 Office 2016 引入的，不可用于与 Office 2013 进行交互。

- **通用** API 可用于访问在多种类型的 Office 应用程序中都很常见的 UI、对话框和客户端设置等功能。 此 API 模型使用的是[回调](https://developer.mozilla.org/docs/Glossary/Callback_function)，其中你仅可在发送给 Office 主机的每个请求中指定一个操作。 通用 API 是随 Office 2013 引入的，可用于与 Office 2013 或更高版本进行交互。 要详细了解通用 API 对象模型（其中包括用于与 Outlook 和 PowerPoint 交互的 API），请参阅 [Office JavaScript API 对象模型](../develop/office-javascript-api-object-model.md)。

> [!NOTE]
> Excel 自定义函数在排列了计算执行优先级的唯一运行时中运行，因此使用的编程模型略有不同。 有关详细信息，请参阅[自定义函数体系结构](../excel/custom-functions-architecture.md)。

有关 Office JavaScript API 的详细信息，请参阅[了解 JavaScript API for Office](../develop/understanding-the-javascript-api-for-office.md)。

#### <a name="api-requirement-sets"></a>API 要求集

[要求集](../develop/office-versions-and-requirement-sets.md)是指各组已命名的 API 成员。 要求集可特定于 Office 主机，例如 `ExcelApi 1.7` 要求集（一组仅可在 Excel 中使用的 API），也可常用于多台主机，例如 `DialogApi 1.1` 要求集（一组可在支持对话框 API 的任何 Office 应用程序中使用的 API）。

加载项可使用要求集来确定 Office 主机是否支持需要使用的 API 成员。 有关详细信息，请参阅[指定 Office 主机和 API 要求](../develop/specify-office-hosts-and-api-requirements.md)。

要求集支持因 Office 主机、版本和平台而异。 要详细了解每个 Office 应用程序支持的平台、要求集和通用 API，请参阅 [Office 加载项主机和平台可用性](office-add-in-availability.md)。

## <a name="testing-and-debugging-an-office-add-in"></a>测试和调试 Office 加载项

开发加载项时，可使用一种名为_旁加载_的技术在本地测试它。 加载项的旁加载过程因平台而异，在某些情况下，也因产品而异。 同样地，加载项的调试流程也因平台和产品而异。 有关测试和调试的详细信息，请参阅[测试和调试 Office 加载项](../testing/test-debug-office-add-ins.md)。

## <a name="publishing-an-office-add-in"></a>发布 Office 加载项

当准备好与他人共享加载项时，可使用最符合你的目标的部署方法实现这一点。 例如，若要将加载项部署给组织内部用户，可使用集中式部署或在 SharePoint 应用目录中发布加载项。 如果想要公开共享加载项供任何人获取，可在 AppSource 中发布加载项。 有关发布的详细信息，请参阅[部署和发布 Office 加载项](../publish/publish.md)。

## <a name="next-steps"></a>后续步骤

本文概述了创建 Office 加载项的不同方法、介绍了 Script Lab（一种用来了解 Office JavaScript API 和建立加载项功能原型的宝贵工具），还描述了重要的 Office 加载项开发、测试和发布概念。 现在，你了解这一介绍性信息，请考虑沿着以下学习路径继续你的 Office 加载项之旅。

### <a name="create-an-office-add-in"></a>创建 Office 加载项

可完成 [5 分钟快速入门](../index.md)，快速创建适合 Excel、OneNote、Outlook、PowerPoint、Project 或 Word 的基本加载项。 如果你之前已完成快速入门，并且想要创建更复杂一些的加载项，请尝试本[教程](../index.md)。

### <a name="explore-the-apis-with-script-lab"></a>使用 Script Lab 了解 API

了解 [Script Lab](explore-with-script-lab.md) 中的内置示例库，熟悉 Office JavaScript API 的功能。

### <a name="learn-more"></a>了解详细信息

查看此文档，详细了解如何开发、测试和发布 Office 加载项。

> [!TIP]
> 对于你构建的任何加载项，都可查看本文档的[核心概念](core-concepts-office-add-ins.md)部分中的信息，还可查看与你要构建的加载项类型（例如 [Excel](../excel/index.md)）相对应的主机特定部分中的信息。
>
> ![显示目录的图像](../images/top-level-toc.png)

## <a name="see-also"></a>另请参阅 

- [Office 加载项平台概述](office-add-ins.md)
- [Office 加载项的核心概念](../overview/core-concepts-office-add-ins.md)
- [开发 Office 加载项](../develop/develop-overview.md)
- [使用 Visual Studio Code 开发 Office 加载项](../develop/develop-add-ins-vscode.md)
- [使用 Visual Studio 开发 Office 加载项](../develop/develop-add-ins-visual-studio.md)
- [设计 Office 加载项](../design/add-in-design.md)
- [测试和调试 Office 加载项](../testing/test-debug-office-add-ins.md)
- [发布 Office 加载项](../publish/publish.md)