---
title: 参考 Office JavaScript API 库
description: 了解如何在加载项中引用 Office JavaScript API 库和类型定义。
ms.date: 02/18/2021
ms.localizationpriority: medium
ms.openlocfilehash: 38121fe3d3df0a86fef3e2c8e3a58399640f1e2a
ms.sourcegitcommit: 4ba5f750358c139c93eb2170ff2c97322dfb50df
ms.translationtype: MT
ms.contentlocale: zh-CN
ms.lasthandoff: 07/06/2022
ms.locfileid: "66660114"
---
# <a name="referencing-the-office-javascript-api-library"></a>参考 Office JavaScript API 库

[Office JavaScript API](../reference/javascript-api-for-office.md) 库提供外接程序可用于与 Office 应用程序交互的 API。 引用库的最简单方法是在 HTML 页面的一节中`<head>`添加以下`<script>`标记来使用内容分发网络 (CDN) 。

```html
<head>
    ...
    <script src="https://appsforoffice.microsoft.com/lib/1/hosted/office.js" type="text/javascript"></script>
</head>
```

这会在加载项第一次加载时下载并缓存 Office JavaScript API 文件，以确保它使用针对指定版本的Office.js及其关联文件的最新实现。

> [!IMPORTANT]
> 必须从页面部分内 `<head>` 引用 Office JavaScript API，以确保在任何正文元素之前完全初始化 API。

## <a name="api-versioning-and-backward-compatibility"></a>API 版本控制和向后兼容性

在上一个 HTML 代码片段中， `/1/` CDN URL 前面 `office.js` 指定Office.js版本 1 中的最新增量版本。 由于 Office JavaScript API 保持向后兼容性，因此最新版本将继续支持版本 1 中前面引入的 API 成员。 如果需要升级现有项目，请参阅 [更新 Office JavaScript API 和清单架构文件的版本](update-your-javascript-api-for-office-and-manifest-schema-version.md)。 

如果计划从 AppSource 发布 Office 加载项，必须使用此 CDN 引用。本地引用仅适用于内部、开发和调试应用场景。

> [!NOTE]
> 要使用预览版 API，请参考 CDN 上的 Office JavaScript API 库预览版：`https://appsforoffice.microsoft.com/lib/beta/hosted/office.js`。

## <a name="enabling-intellisense-for-a-typescript-project"></a>为 TypeScript 项目启用 IntelliSense

除了如前所述引用 Office JavaScript API 外，还可以使用 [DefinitelyTyped](https://github.com/DefinitelyTyped/DefinitelyTyped/tree/master/types/office-js) 中的类型定义启用 IntelliSense for TypeScript 外接程序项目。 为此，请在启用节点的系统提示符 (或 git bash 窗口中运行以下命令，) 项目文件夹的根目录。 必须安装 [Node.js](https://nodejs.org)（包括 npm）。

```command&nbsp;line
npm install --save-dev @types/office-js
```

## <a name="preview-apis"></a>预览 API

新的 JavaScript API 首先在“预览”中引入，然后在进行足够的测试并获取用户反馈后成为特定编号要求集的一部分。

[!INCLUDE [Information about using preview APIs](../includes/using-preview-apis-host.md)]

## <a name="see-also"></a>另请参阅

- [了解 Office JavaScript API](understanding-the-javascript-api-for-office.md)
- [Office JavaScript API](../reference/javascript-api-for-office.md)
