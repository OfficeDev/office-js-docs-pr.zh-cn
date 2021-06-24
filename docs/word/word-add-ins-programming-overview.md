---
title: Word 加载项概述
description: 了解 Word 加载项的基本知识。
ms.date: 10/14/2020
ms.topic: conceptual
ms.custom: scenarios:getting-started
localization_priority: Priority
ms.openlocfilehash: c4abde797ac25b049e3d77acad59f7e2263005aa
ms.sourcegitcommit: ee9e92a968e4ad23f1e371f00d4888e4203ab772
ms.translationtype: HT
ms.contentlocale: zh-CN
ms.lasthandoff: 06/23/2021
ms.locfileid: "53075542"
---
# <a name="word-add-ins-overview"></a>Word 加载项概述

要创建解决方案来扩展 Word 功能？例如，涉及自动文档程序集的解决方案？或从其他数据源绑定到并访问 Word 文档中数据的解决方案？可以使用 Office 加载项平台，其中包含 Word JavaScript API 和Office JavaScript API，可用于扩展在 Windows 桌面设备、Mac 或云中运行的 Word 客户端。

Word 外接程序是 [Office 外接程序平台](../overview/office-add-ins.md)上众多开发选项中的一项。外接程序命令可用于扩展 Word 用户界面并启动运行 JavaScript 并与 Word 文档中内容交互的任务窗格。在浏览器中可以运行的任何代码均可在 Word 外接程序中运行。与 Word 文档内容进行交互的外接程序可创建作用于 Word 对象的请求并同步对象状态。

[!INCLUDE [publish policies note](../includes/note-publish-policies.md)]

下图中的示例展示了在任务窗格中运行的 Word 加载项。

*图 1：在 Word 的任务窗格中运行的加载项*

![在 Word 的任务窗格中运行的加载项。](../images/word-add-in-show-host-client.png)

Word 外接程序 (1) 可以将请求发送到 Word 文档 (2) 可以使用 JavaScript 来访问段落对象和更新、删除或移动段落。例如，下面的代码演示如何将一个新句子附加到该段落。

```js
Word.run(function (context) {
    var paragraphs = context.document.getSelection().paragraphs;
    paragraphs.load();
    return context.sync().then(function () {
        paragraphs.items[0].insertText(' New sentence in the paragraph.',
                                       Word.InsertLocation.end);
    }).then(context.sync);
});

```

若要托管 Word 加载项，可以使用任何 Web 服务器技术（如 ASP.NET、NodeJS 或 Python）。可以使用常用的客户端框架（Ember、Backbone、Angular、React），也可以坚持使用 VanillaJS 开发解决方案，并能使用 Azure 等服务[验证](../develop/overview-authn-authz.md)和托管应用。

通过 Word JavaScript API 可使应用程序访问 Word 文档中的对象和元数据。这些 API 可用于创建面向以下应用程序的外接程序：

* Windows 版 Word 2013 或更高版本
* Word 网页版
* Mac 版 Word 2016 或更高版本
* iPad 版 Word

加载项只需编写一次，即可跨多个平台在所有版本 Word 中运行。有关详细信息，请参阅 [Office 客户端应用程序和加载项平台可用性](../overview/office-add-in-availability.md)。

## <a name="javascript-apis-for-word"></a>适用于 Word 的 JavaScript API

有两组 JavaScript API 可用于与 Word 文档中的对象和元数据进行交互。 第一组是在 Office 2013 中引入的[通用 API](/javascript/api/office)。 通用 API 中的许多对象可以在由两个或多个 Office 客户端托管的加载项中使用。 此 API 广泛使用回调。

第二组是 [Word JavaScript API](/javascript/api/word)。这是与 Word 2016 年一起引入的[应用程序特定 API 模型](../develop/application-specific-api-model.md)。它是强类型对象模型，可用于创建面向 Mac 版和 Windows 版 Word 2016 的 Word 加载项。此对象模型使用承诺模式，并提供对特定于 Word 的对象（如[正文](/javascript/api/word/word.body)、[内容控件](/javascript/api/word/word.contentcontrol)、[内联图片](/javascript/api/word/word.inlinepicture)和[段落](/javascript/api/word/word.paragraph)）的访问权限。Word JavaScript API 包括 TypeScript 定义和 vsdoc 文件，这样，你便可以在 IDE 中获得代码提示。

目前，所有 Word 客户端均支持共享 Office  JavaScript API，大多数客户端支持 Word JavaScript API。有关受支持的客户端的详细信息，请参阅[ Office 客户端应用程序和 Office 加载项的平台可用性](../overview/office-add-in-availability.md)。

我们建议从 Word JavaScript API 开始，因为对象模型更易于使用。如果需要执行以下操作，请使用 Word JavaScript API：

* 访问 Word 文档中的对象。

在需要执行以下操作时，使用共享的 Office JavaScript API：

* 面向 Word 2013。
* 执行应用程序的初始操作。
* 检查支持的要求集。
* 访问文档的元数据、设置和环境信息。
* 绑定到文档中的部分并捕获事件。
* 使用自定义 XML 部件。
* 打开一个对话框。

## <a name="next-steps"></a>后续步骤

准备好创建首个 Word 加载项了吗？请参阅[构建首个 Word 加载项](../quickstarts/word-quickstart.md)。请使用[加载项清单](../develop/add-in-manifests.md)描述加载项的托管位置和显示方式，并定义权限和其他信息。

若要了解如何设计世界一流的 Word 外接程序来为用户打造具有吸引力的体验，请参阅[设计指南](../design/add-in-design.md)和[最佳实践](../concepts/add-in-development-best-practices.md)。

开发加载项后，可以将它[发布](../publish/publish.md)到网络共享、应用目录或 AppSource。

## <a name="see-also"></a>另请参阅

* [开发 Office 加载项](../develop/develop-overview.md)
* [了解 Microsoft 365 开发人员计划](https://developer.microsoft.com/microsoft-365/dev-program)
* [Office 加载项平台概述](../overview/office-add-ins.md)
* [Word JavaScript API 参考](../reference/overview/word-add-ins-reference-overview.md)