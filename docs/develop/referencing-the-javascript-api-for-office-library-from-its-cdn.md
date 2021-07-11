---
title: 参考 Office JavaScript API 库
description: 了解如何在外接程序Office JavaScript API 库和类型定义。
ms.date: 02/18/2021
localization_priority: Normal
ms.openlocfilehash: 04f97412c07cb39f5b2f753c3ce14e56e87c3de5
ms.sourcegitcommit: 883f71d395b19ccfc6874a0d5942a7016eb49e2c
ms.translationtype: MT
ms.contentlocale: zh-CN
ms.lasthandoff: 07/09/2021
ms.locfileid: "53349754"
---
# <a name="referencing-the-office-javascript-api-library"></a>参考 Office JavaScript API 库

Office [JavaScript API](../reference/javascript-api-for-office.md)库提供了外接程序可用于与应用程序应用程序交互Office API。 引用库的最简单方法就是使用内容交付网络 (CDN) HTML 页面的 部分添加 `<script>` `<head>` 以下标记。

```html
<head>
    ...
    <script src="https://appsforoffice.microsoft.com/lib/1/hosted/office.js" type="text/javascript"></script>
</head>
```

这将在外接程序首次加载时下载并缓存 Office JavaScript API 文件，以确保其对指定版本使用 Office.js 及其关联文件最新的实现。

> [!IMPORTANT]
> 你必须从页面Office引用 JavaScript API，以确保 API 在任意 body 元素之前 `<head>` 完全初始化。

## <a name="api-versioning-and-backward-compatibility"></a>API 版本控制与向后兼容性

在前面的 HTML 代码段中，CDN URL 中的 前面的 指定版本 1 中的最新增量 `/1/` `office.js` Office.js。 由于 Office JavaScript API 保持向后兼容性，因此最新版本将继续支持版本 1 中之前引入的 API 成员。 如果需要升级现有项目，请参阅[更新 javaScript API Office清单架构文件的版本](update-your-javascript-api-for-office-and-manifest-schema-version.md)。 

如果计划从 AppSource 发布 Office 加载项，必须使用此 CDN 引用。本地引用仅适用于内部、开发和调试应用场景。

> [!NOTE]
> 要使用预览版 API，请参考 CDN 上的 Office JavaScript API 库预览版：`https://appsforoffice.microsoft.com/lib/beta/hosted/office.js`。

## <a name="enabling-intellisense-for-a-typescript-project"></a>为IntelliSense启用项目

除了如前面Office引用 JavaScript API 外，您还可以使用[DefinitelyTyped](https://github.com/DefinitelyTyped/DefinitelyTyped/tree/master/types/office-js)中的类型定义为 TypeScript 外接程序项目启用 IntelliSense。 为此，请从项目文件夹的根目录 (启用节点的系统提示符或 git bash) 运行以下命令。 必须安装 [Node.js](https://nodejs.org)（包括 npm）。

```command&nbsp;line
npm install --save-dev @types/office-js
```

## <a name="preview-apis"></a>预览 API

新的 JavaScript API 首先在"预览版"中引入，之后在经过充分测试且需要用户反馈后成为特定编号要求集的一部分。

[!INCLUDE [Information about using preview APIs](../includes/using-preview-apis-host.md)]

## <a name="see-also"></a>另请参阅

- [了解 Office JavaScript API](understanding-the-javascript-api-for-office.md)
- [Office JavaScript API](../reference/javascript-api-for-office.md)
