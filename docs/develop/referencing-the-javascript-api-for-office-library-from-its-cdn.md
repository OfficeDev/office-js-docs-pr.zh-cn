---
title: 从内容传送网络 (CDN) 引用适用于 Office 的 JavaScript API 库
description: ''
ms.date: 03/19/2019
localization_priority: Priority
ms.openlocfilehash: 6b9512d5d0969e185902d7ab9d3227e820c4d0dc
ms.sourcegitcommit: 528577145b2cf0a42bc64c56145d661c4d019fb8
ms.translationtype: HT
ms.contentlocale: zh-CN
ms.lasthandoff: 10/02/2019
ms.locfileid: "37353816"
---
# <a name="referencing-the-javascript-api-for-office-library-from-its-content-delivery-network-cdn"></a>从内容传送网络 (CDN) 引用适用于 Office 的 JavaScript API 库

> [!NOTE]
> 如果想要使用 TypeScript 获取 IntelliSense，除了本文中所述的步骤之外，还需要在项目文件夹根目录的节点支持系统提示框（或 Git Bash 窗口）中运行以下命令。 必须安装 [Node.js](https://nodejs.org)（包括 npm）。
> 
> ```command&nbsp;line
> npm install --save-dev @types/office-js
> ```

[适用于 Office 的 JavaScript API](/office/dev/add-ins/reference/javascript-api-for-office) 库包含 Office.js 文件和关联的主机应用专用 .js 文件（如 Excel-15.js 和 Outlook-15.js）。 


引用 API 的最简单方法是，通过向页面的 `<head>` 标记添加以下 `<script>` 来使用我们的 CDN：  

```html
<script src="https://appsforoffice.microsoft.com/lib/1/hosted/office.js" type="text/javascript"></script>
```

在 CDN URL 中，`office.js` 前面的 `/1/` 指定 Office.js 第 1 版中的最新增量版本。由于适用于 Office 的 JavaScript API 保留向后兼容性，因此最新版本将继续支持之前在第 1 版中引入的 API 成员。如果需要升级现有项目，请参阅[更新适用于 Office 的 JavaScript API 的版本和清单架构文件](update-your-javascript-api-for-office-and-manifest-schema-version.md)。 

如果计划从 AppSource 发布 Office 加载项，必须使用此 CDN 引用。本地引用仅适用于内部、开发和调试应用场景。

> [!IMPORTANT]
> 开发任何 Office 主机应用的加载项时，请从页面的 `<head>` 部分引用适用于 Office 的 JavaScript API。这样可确保 API 先于所有正文元素完全初始化。Office 主机要求，加载项必须在激活后的 5 秒内初始化。如果加载项未在此阈值内激活，则会被声明为无响应，并且用户会看到错误消息。

## <a name="see-also"></a>另请参阅

- [了解适用于 Office 的 JavaScript API](understanding-the-javascript-api-for-office.md)
- [适用于 Office 的 JavaScript API](/office/dev/add-ins/reference/javascript-api-for-office)
