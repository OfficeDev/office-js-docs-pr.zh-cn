---
title: 参考 Office JavaScript API 库
description: 了解如何在外接程序中引用 Office JavaScript API 库和类型定义。
ms.date: 06/23/2020
localization_priority: Normal
ms.openlocfilehash: 64dd08329b7bbc8c249bd270a431b6cbe93ec52c
ms.sourcegitcommit: 9609bd5b4982cdaa2ea7637709a78a45835ffb19
ms.translationtype: MT
ms.contentlocale: zh-CN
ms.lasthandoff: 08/28/2020
ms.locfileid: "47293182"
---
# <a name="referencing-the-office-javascript-api-library"></a>参考 Office JavaScript API 库

[Office JAVASCRIPT API](../reference/javascript-api-for-office.md)库提供你的外接程序可用于与 Office 应用程序进行交互的 api。 引用库的最简单方法是使用内容传递网络 (CDN) ，方法是 `<script>` 在 HTML 页面的部分中添加以下标记 `<head>` ：  

```html
<head>
    ...
    <script src="https://appsforoffice.microsoft.com/lib/1/hosted/office.js" type="text/javascript"></script>
</head>
```

这将在您首次加载加载项时下载并缓存 Office JavaScript API 文件，以确保它使用的是最新的 Office.js 实现，并将其与指定版本相关联的文件。

> [!IMPORTANT]
> 您必须从页面的部分中引用 Office JavaScript API， `<head>` 以确保 API 在任何 body 元素之前完全初始化。 Office 应用程序要求外接程序在激活5秒内初始化。 如果外接程序未在此阈值内激活，则会被声明为无响应，并且用户会看到错误消息。

## <a name="api-versioning-and-backward-compatibility"></a>API 版本控制和向后兼容性

在上面的 HTML 代码段中，CDN URL 中的 " `/1/` 在 `office.js` Office.js 的第1版中指定最新的增量释放。 由于 Office JavaScript API 保持向后兼容性，最新版本将继续支持之前在版本1中引入的 API 成员。 如果需要升级现有项目，请参阅 [更新 Office JAVASCRIPT API 和清单架构文件的版本](update-your-javascript-api-for-office-and-manifest-schema-version.md)。 

如果计划从 AppSource 发布 Office 加载项，必须使用此 CDN 引用。本地引用仅适用于内部、开发和调试应用场景。

> [!NOTE]
> 要使用预览版 API，请参考 CDN 上的 Office JavaScript API 库预览版：`https://appsforoffice.microsoft.com/lib/beta/hosted/office.js`。

## <a name="enabling-intellisense-for-a-typescript-project"></a>为 TypeScript 项目启用 IntelliSense

除了参照前面所述的 Office JavaScript API 之外，还可以使用 [jquery.typescript.definitelytyped](https://github.com/DefinitelyTyped/DefinitelyTyped/tree/master/types/office-js)中的类型定义为 TypeScript 加载项项目启用 IntelliSense。 若要执行此操作，请在已启用节点的系统提示符 (或 git bash 窗口中) 从项目文件夹的根目录中运行以下命令。 必须安装 [Node.js](https://nodejs.org)（包括 npm）。

```command&nbsp;line
npm install --save-dev @types/office-js
```

## <a name="preview-apis"></a>预览 Api

新的 JavaScript Api 首先在 "预览" 中引入，并在进行充分的测试并需要用户反馈之后成为特定编号的要求集的一部分。

[!INCLUDE [Information about using preview APIs](../includes/using-preview-apis-host.md)]

## <a name="see-also"></a>另请参阅

- [了解 Office JavaScript API](understanding-the-javascript-api-for-office.md)
- [Office JavaScript API](../reference/javascript-api-for-office.md)
