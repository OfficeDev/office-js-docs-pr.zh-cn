---
title: 参考 Office JavaScript API 库
description: 了解如何在加载项中引用 Office JavaScript API 库和类型定义。
ms.date: 02/18/2021
localization_priority: Normal
ms.openlocfilehash: 346a34c0cbc31b5e569a5106dcd2bc01593b114a
ms.sourcegitcommit: e7009c565b18c607fe0868db2e26e250ad308dce
ms.translationtype: MT
ms.contentlocale: zh-CN
ms.lasthandoff: 03/05/2021
ms.locfileid: "50505190"
---
# <a name="referencing-the-office-javascript-api-library"></a>参考 Office JavaScript API 库

[Office JavaScript API](../reference/javascript-api-for-office.md)库提供加载项可用于与 Office 应用程序交互的 API。 引用库的最简单方法就是通过添加以下标记 (CDN) 内容传送 `<script>` `<head>` 网络：  

```html
<head>
    ...
    <script src="https://appsforoffice.microsoft.com/lib/1/hosted/office.js" type="text/javascript"></script>
</head>
```

这将在外接程序首次加载时下载并缓存 Office JavaScript API 文件，以确保它使用指定版本的 Office.js 及其关联文件的最最新实现。

> [!IMPORTANT]
> 您必须从页面部分内引用 Office JavaScript API，以确保 API 在任何正文元素之前 `<head>` 完全初始化。

## <a name="api-versioning-and-backward-compatibility"></a>API 版本控制与向后兼容性

在之前的 HTML 代码段中，CDN URL 中的前面部分指定版本 1 中的最新增量 `/1/` `office.js` Office.js。 由于 Office JavaScript API 保持向后兼容性，因此最新版本将继续支持在版本 1 中之前引入的 API 成员。 如果需要升级现有项目，请参阅更新 Office [JavaScript API 的版本和清单架构文件](update-your-javascript-api-for-office-and-manifest-schema-version.md)。 

如果计划从 AppSource 发布 Office 加载项，必须使用此 CDN 引用。本地引用仅适用于内部、开发和调试应用场景。

> [!NOTE]
> 要使用预览版 API，请参考 CDN 上的 Office JavaScript API 库预览版：`https://appsforoffice.microsoft.com/lib/beta/hosted/office.js`。

## <a name="enabling-intellisense-for-a-typescript-project"></a>为 TypeScript IntelliSense启用项目

除了如前面所述引用 Office JavaScript API，您还可以使用 [DefinitelyTyped](https://github.com/DefinitelyTyped/DefinitelyTyped/tree/master/types/office-js)中的类型定义为 TypeScript 外接程序项目启用 IntelliSense。 为此，请从项目文件夹的根目录 (启用节点的系统提示符或 git bash) 运行以下命令。 必须安装 [Node.js](https://nodejs.org)（包括 npm）。

```command&nbsp;line
npm install --save-dev @types/office-js
```

## <a name="preview-apis"></a>预览 API

新的 JavaScript API 首先在"预览"中引入，稍后在经过充分测试且需要用户反馈后成为特定编号要求集的一部分。

[!INCLUDE [Information about using preview APIs](../includes/using-preview-apis-host.md)]

## <a name="see-also"></a>另请参阅

- [了解 Office JavaScript API](understanding-the-javascript-api-for-office.md)
- [Office JavaScript API](../reference/javascript-api-for-office.md)
