---
title: 使用 Script Lab 探索 Office JavaScript API
description: 使用 Script Lab 探索 Office JS API 和原型功能。
ms.date: 08/05/2021
ms.topic: conceptual
ms.custom: scenarios:getting-started
ms.localizationpriority: high
ms.openlocfilehash: 2a8d7f94eab26bc1cb831f95a30442e8e08afe70
ms.sourcegitcommit: 33824aa3995a2e0bcc6d8e67ada46f296c224642
ms.translationtype: HT
ms.contentlocale: zh-CN
ms.lasthandoff: 01/12/2022
ms.locfileid: "61765932"
---
# <a name="explore-office-javascript-api-using-script-lab"></a>使用 Script Lab 探索 Office JavaScript API

借助 [Script Lab](https://appsource.microsoft.com/product/office/WA104380862) 和[适用于 Outlook 的 Script Lab](https://appsource.microsoft.com/product/office/WA200001603) 加载项（可从 AppSource 免费获取），你可以在使用 Excel 或 Outlook 等 Office 程序时探索 Office JavaScript API。 Script Lab 是一项方便的工具，可将其作为原型添加到开发工具包，并在你自己的加载项中验证你想使用的功能。

## <a name="what-is-script-lab"></a>什么是 Script Lab？

Script Lab 是一款面向具有以下需求的用户的工具：希望了解如何在 Excel、Outlook、Word 和 PowerPoint 中开发使用 Office JavaScript API 的 Office 加载项。 它提供 IntelliSense，让你可以看到可用的内容；并且它是基于 Monaco 框架构建的（Visual Studio Code 也使用该框架）。 通过 Script Lab，可访问示例库以快速试用各项功能，也由示例开始编写自己的代码。 甚至可以通过 Script Lab 试用预览 API。

到目前为止听起来不错？观看以下片长一分钟的视频，在操作中了解 Script Lab。

[![展示 Script Lab 在 Excel、Word 和 PowerPoint 中运行的预览视频。](../images/screenshot-wide-youtube.png 'Script Lab 预览视频')](https://aka.ms/scriptlabvideo)

## <a name="key-features"></a>关键功能

Script Lab 提供许多功能，可帮助你探索 Office JavaScript API 和原型加载项功能。

### <a name="explore-samples"></a>浏览示例

通过一系列展示如何使用 API 完成任务的内置示例快速入门。 可以运行示例来立即查看任务窗格或文档中的结果，检查示例来了解 API 的工作原理，甚至可以使用示例来构建自己的加载项的原型。

![示例。](../images/script-lab-samples.jpg)

### <a name="code-and-style"></a>代码和样式

除了用于调用 Office JS API 的 JavaScript 或 TypeScript 代码之外，每个代码段还包含用于定义任务窗格内容的 HTML 标记和用于定义任务窗格外观的 CSS。 在为自己的加载项确定任务窗格设计原型时，可以自定义该 HTML 标记 和 CSS，对元素放置和样式设计进行试验。

> [!TIP]
> 若要在代码段中调用预览 API，需更新该代码段的库，令其使用 beta CDN (`https://appsforoffice.microsoft.com/lib/beta/hosted/office.js`) 和预览类型定义 `@types/office-js-preview`。 此外，仅当注册 [Office 预览体验计划](https://insider.office.com)后、运行 Office 预览体验计划版本时，才能访问某些预览 API。

### <a name="save-and-share-snippets"></a>保存和共享代码段

默认情况下，在 Script Lab 中打开的代码段将保存到浏览器缓存中。 若要永久保存代码段，可将其导出到 [GitHub gist](https://gist.github.com)。 可创建机密 gist 来保存自己专用的代码段，或创建公用 gist 以便与他人共享。

![共享选项。](../images/script-lab-share.jpg)

### <a name="import-snippets"></a>导入代码段

可通过指定存用于储代码段 YAML 的公共 [GitHub gist](https://gist.github.com) URL，或通过在代码段的完整 YAML 中粘贴，将代码段导入到 Script Lab。 当其他人通过发布到 GitHub gist 或提供 YAML 来与你共享其代码段时，此功能可能很有用。

![导入代码段选项。](../images/script-lab-import-snippet.jpg)

## <a name="supported-clients"></a>支持的客户端

以下客户端上的 Excel、Word 和 PowerPoint 支持 Script Lab。

- Microsoft 365 Office 订阅
- Mac 上的 Office 2016 或更高版本
- Office 网页版

适用于 Outlook 的 Script Lab 在以下客户端上可用。

- Microsoft 365 Office 订阅
- Mac 版 Outlook 2016 或更高版本
- 使用 Chrome、Microsoft Edge 或 Safari 浏览器时的 Outlook 网页版

有关适用于 Outlook 的 Script Lab 的更多详细信息，请参阅相关[博客文章](https://developer.microsoft.com/outlook/blogs/script-lab-now-supports-outlook/)。

> [!IMPORTANT]
> 脚本实验室不再适用于平台和使用 Internet Explorer 托管加载项的 Office 版本组合。这包括通过 Office 2019 一次性购买的 Office 版本。 有关详细信息，请参阅 [Office 加载项使用的浏览器](../concepts/browsers-used-by-office-web-add-ins.md)。

## <a name="next-steps"></a>后续步骤

若要在 Excel、Word 或 PowerPoint 中使用 Script Lab，请从 AppSource 安装 [Script Lab 加载项](https://appsource.microsoft.com/product/office/WA104380862)。

若要使用适用于 Outlook 的 Script Lab，请从 AppSource 安装 [适用于 Outlook 的 Script Lab 加载项](https://appsource.microsoft.com/product/office/wa200001603)。

欢迎将新代码段发布到 [office-js-snippets](https://github.com/OfficeDev/office-js-snippets#office-js-snippets) GitHub 存储库，以扩充 Script Lab 中的示例库。

准备好创建你的首个 Office 加载项时，请尝试使用 [Excel](../quickstarts/excel-quickstart-jquery.md)、[Outlook](../quickstarts/outlook-quickstart.md)、[Word](../quickstarts/word-quickstart.md)、[OneNote](../quickstarts/onenote-quickstart.md)、[PowerPoint](../quickstarts/powerpoint-quickstart.md) 或 [Project](../quickstarts/project-quickstart.md) 快速入门。

## <a name="see-also"></a>另请参阅

- [获取适用于 Excel、Word 或 Powerpoint 的 Script Lab](https://appsource.microsoft.com/product/office/WA104380862)
- [获取适用于 Outlook 的 Script Lab](https://appsource.microsoft.com/product/office/wa200001603)
- [详细了解 Script Lab](https://github.com/OfficeDev/script-lab#script-lab-a-microsoft-garage-project)
- [加入 Microsoft 365 开发人员计划](https://developer.microsoft.com/office/dev-program)
- [开发 Office 加载项](../develop/develop-overview.md)
- [了解 Microsoft 365 开发人员计划](https://developer.microsoft.com/microsoft-365/dev-program)
