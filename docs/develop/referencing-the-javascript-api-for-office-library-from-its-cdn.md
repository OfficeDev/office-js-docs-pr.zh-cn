---
title: 参考 Office JavaScript API 库
description: 了解如何在外接程序中引用 Office JavaScript API 库和类型定义。
ms.date: 02/27/2020
localization_priority: Normal
ms.openlocfilehash: 5e26d5b0454a6833c593ff60c1577d24583dcc51
ms.sourcegitcommit: 4079903c3cc45b7d8c041509a44e9fc38da399b1
ms.translationtype: MT
ms.contentlocale: zh-CN
ms.lasthandoff: 03/11/2020
ms.locfileid: "42596716"
---
# <a name="referencing-the-office-javascript-api-library"></a>参考 Office JavaScript API 库

[Office JAVASCRIPT API](../reference/javascript-api-for-office.md)库提供你的外接程序可用于与 Office 主机进行交互的 api。 若要引用库，最简单的方法是使用内容传送网络（CDN），方法是在`<script>` HTML 页面的`<head>`部分中添加以下标记：  

```html
<head>
    ...
    <script src="https://appsforoffice.microsoft.com/lib/1/hosted/office.js" type="text/javascript"></script>
</head>
```

这将在首次加载加载项时下载并缓存 Office JavaScript API 文件，以确保它使用的是最新的 Office js 实现并将其与指定版本关联的文件关联起来。

> [!IMPORTANT]
> 您必须从页面的`<head>`部分中引用 OFFICE JavaScript API，以确保 API 在任何 body 元素之前完全初始化。 Office 主机要求外接程序在激活后的 5 秒内进行初始化。 如果外接程序未在此阈值内激活，则会被声明为无响应，并且用户会看到错误消息。

## <a name="api-versioning-and-backward-compatibility"></a>API 版本控制和向后兼容性

在上面的 HTML 代码段中`/1/` ，CDN URL `office.js`中的 "在第1版" 中指定的最新增量发布是在 Office .js 的第1版中。 由于 Office JavaScript API 保持向后兼容性，最新版本将继续支持之前在版本1中引入的 API 成员。 如果需要升级现有项目，请参阅[更新 Office JAVASCRIPT API 和清单架构文件的版本](update-your-javascript-api-for-office-and-manifest-schema-version.md)。 

如果计划从 AppSource 发布 Office 加载项，必须使用此 CDN 引用。本地引用仅适用于内部、开发和调试应用场景。

> [!NOTE]
> 要使用预览版 API，请参考 CDN 上的 Office JavaScript API 库预览版：`https://appsforoffice.microsoft.com/lib/beta/hosted/office.js`。

## <a name="enabling-intellisense-for-a-typescript-project"></a>为 TypeScript 项目启用 Intellisense

除了参照前面所述的 Office JavaScript API 之外，还可以使用[jquery.typescript.definitelytyped](https://github.com/DefinitelyTyped/DefinitelyTyped/tree/master/types/office-js)中的类型定义为 TypeScript 加载项项目启用 Intellisense。 若要执行此操作，请在已启用节点的系统提示符（或 git bash 窗口）中，从您的项目文件夹的根处运行以下命令。 必须安装 [Node.js](https://nodejs.org)（包括 npm）。

```command&nbsp;line
npm install --save-dev @types/office-js
```

> [!NOTE]
> 若要为预览 Api 启用 Intellisense，请在项目文件夹的根目录中运行以下命令，以使用[jquery.typescript.definitelytyped](https://github.com/DefinitelyTyped/DefinitelyTyped/tree/master/types/office-js-preview)中的预览类型定义： 
>
> `npm install --save-dev @types/office-js-preview`

## <a name="see-also"></a>另请参阅

- [了解 Office JavaScript API](understanding-the-javascript-api-for-office.md)
- [Office JavaScript API](../reference/javascript-api-for-office.md)
