---
title: 使用 Script Lab 探索 Office JavaScript API
description: 使用脚本实验室浏览 Office JS API 并建立原型功能。
ms.date: 07/05/2019
ms.topic: conceptual
ms.custom: scenarios:getting-started
localization_priority: Normal
ms.openlocfilehash: 3212aec08cdf4e0185ae5856ae522b1d81e28ea1
ms.sourcegitcommit: 212c810f3480a750df779777c570159a7f76054a
ms.translationtype: MT
ms.contentlocale: zh-CN
ms.lasthandoff: 01/17/2020
ms.locfileid: "41216971"
---
# <a name="explore-office-javascript-api-using-script-lab"></a>使用 Script Lab 探索 Office JavaScript API

[脚本实验室外接程序](https://appsource.microsoft.com/product/office/WA104380862)（从 AppSource 中免费获取）使您能够在使用 office 程序（如 Excel 或 Word）时浏览 OFFICE JavaScript API。 当您在外接程序中原型和验证所需功能时，脚本实验室是一个方便的工具，可将其添加到开发工具包中。

## <a name="what-is-script-lab"></a>什么是脚本实验室？

脚本实验室是任何希望了解如何使用 Excel、Word 或 PowerPoint 中的 Office JavaScript API 开发 Office 外接程序的工具。 它提供了智能感知功能，以便您可以查看在摩纳哥框架（由 Visual Studio Code 使用的相同框架）中构建的可用功能。 通过脚本实验室，您可以访问示例库以快速试用功能，也可以将示例用作您自己的代码的起始点。 您甚至可以使用脚本实验室尝试预览 Api。

我到目前为止听起来正常吗？ 查看此一分钟视频可查看脚本实验室的实际效果。

[![展示 Script Lab 在 Excel、Word 和 PowerPoint 中运行的预览视频。](../images/screenshot-wide-youtube.png 'Script Lab 预览视频')](https://aka.ms/scriptlabvideo)

## <a name="key-features"></a>关键功能

脚本实验室提供了许多功能，可帮助您探索 Office JavaScript API 和原型加载项功能。

### <a name="explore-samples"></a>浏览示例

使用内置示例代码段集合快速入门，其中展示了如何使用 API 完成任务。 您可以运行示例来即时查看任务窗格或文档中的结果，检查示例以了解 API 的工作原理，甚至使用示例来原型自己的外接程序。

![示例](../images/script-lab-samples.jpg)

### <a name="code-and-style"></a>代码和样式

除了调用 Office JS API 的 JavaScript 或 TypeScript 代码外，每个代码段还包含用于定义任务窗格外观的任务窗格和 CSS 内容的 HTML 标记。 您可以自定义 HTML 标记和 CSS 以在为自己的外接程序设置任务窗格设计原型时体验元素的放置和样式。

> [!TIP]
> 若要在代码段中调用预览 Api，您需要更新代码段的库以使用 beta CDN （`https://appsforoffice.microsoft.com/lib/beta/hosted/office.js`）和预览类型定义。 `@types/office-js-preview` 此外，某些预览 Api 仅当你注册[Office 预览体验计划](https://products.office.com/office-insider)并运行内部版本的 office 时才可访问。

### <a name="save-and-share-snippets"></a>保存和共享代码段

默认情况下，在脚本实验室中打开的代码段将保存到您的浏览器缓存中。 若要永久保存代码段，可以将其导出到[GitHub gist](https://gist.github.com)。 创建一个机密 gist 以仅用于您自己使用的代码段，或者创建一个公用 gist （如果您计划与其他人共享它）。

![共享选项](../images/script-lab-share.jpg)

### <a name="import-snippets"></a>导入代码段

您可以通过指定存储代码段 YAML 的公共[GitHub gist](https://gist.github.com)的 URL 或在代码段的完整 YAML 中粘贴，将代码段导入脚本实验室。 如果其他人已通过将代码段发布到 GitHub gist 或提供代码段的 YAML，则此功能可能对您共享其代码段的方案有用。

![导入代码段选项](../images/script-lab-import-snippet.jpg)

## <a name="supported-clients"></a>支持的客户端

以下客户端上的 Excel、Word 和 PowerPoint 支持脚本实验室。

- Windows 上的 Office 2013 或更高版本
- Mac 上的 Office 2016 或更高版本
- Office 网页版

## <a name="next-steps"></a>后续步骤

若要在 Excel、Word 或 PowerPoint 中使用脚本实验室，请从 AppSource 安装[脚本实验室加载项](https://appsource.microsoft.com/product/office/WA104380862)。 

欢迎您通过将新代码片段发布到[office js](https://github.com/OfficeDev/office-js-snippets#office-js-snippets)的 GitHub 存储库来扩展脚本实验室中的示例库。

当您准备好创建第一个 Office 加载项时，请试用[Excel](../quickstarts/excel-quickstart-jquery.md)、 [Outlook](/outlook/add-ins/quick-start?context=office/dev/add-ins/context)、 [Word](../quickstarts/word-quickstart.md)、 [OneNote](../quickstarts/onenote-quickstart.md)、 [PowerPoint](../quickstarts/powerpoint-quickstart.md)或[Project](../quickstarts/project-quickstart.md)的快速入门。

## <a name="see-also"></a>另请参阅

- [获取脚本实验室](https://appsource.microsoft.com/product/office/WA104380862)
- [了解有关脚本实验室的详细信息](https://github.com/OfficeDev/script-lab#script-lab-a-microsoft-garage-project)
- [加入 Office 365 开发人员计划](https://developer.microsoft.com/office/dev-program)
- [构建 Office 加载项](../overview/office-add-ins-fundamentals.md)
