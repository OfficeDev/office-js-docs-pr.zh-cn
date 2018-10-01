---
title: 从内容传送网络 (CDN) 引用适用于 Office 的 JavaScript API 库
description: ''
ms.date: 01/23/2018
ms.openlocfilehash: 422cbd947dde09a8cd19559db9a86ddacd5e2dba
ms.sourcegitcommit: fdf7f4d686700edd6e6b04b2ea1bd43e59d4a03a
ms.translationtype: HT
ms.contentlocale: zh-CN
ms.lasthandoff: 09/28/2018
ms.locfileid: "25348091"
---
# <a name="referencing-the-javascript-api-for-office-library-from-its-content-delivery-network-cdn"></a>从内容传送网络 (CDN) 引用适用于 Office 的 JavaScript API 库

> [!NOTE]
> 除了本文中描述的步骤之外，如果要使用 TypeScript，然后获取 Intellisense，您需要在项目文件夹的根目录下在启用 Node 的系统提示符（或 git bash 窗口）中运行以下命令。 您必须安装 [Node.js](https://nodejs.org) （其中包括 npm）。
> 
> ```
> npm install --save-dev @types/office-js
> ```

[适用于 Office 的 JavaScript API](https://docs.microsoft.com/javascript/office/javascript-api-for-office?view=office-js) 库包含 Office.js 文件和关联的主机应用专用 .js 文件（如 Excel-15.js 和 Outlook-15.js）。 


引用 API 的最简单方法是，通过向页面的 `<head>` 标记添加以下 `<script>` 来使用我们的 CDN：  

```html
<script src="https://appsforoffice.microsoft.com/lib/1/hosted/Office.js" type="text/javascript"></script>
```

在 CDN URL 中，`office.js` 前面的 `/1/` 指定 Office.js 第 1 版中的最新增量版本。由于适用于 Office 的 JavaScript API 保留向后兼容性，因此最新版本将继续支持之前在第 1 版中引入的 API 成员。如果需要升级现有项目，请参阅[更新适用于 Office 的 JavaScript API 的版本和清单架构文件](update-your-javascript-api-for-office-and-manifest-schema-version.md)。 

如果计划从 AppSource 发布 Office 加载项，必须使用此 CDN 引用。本地引用仅适用于内部、开发和调试应用场景。

> [!IMPORTANT]
>  开发任何 Office 主机应用的加载项时，请从页面的 `<head>` 部分引用适用于 Office 的 JavaScript API。这样可确保 API 先于所有正文元素完全初始化。Office 主机要求，加载项必须在激活后的 5 秒内初始化。如果加载项未在此阈值内激活，则会被声明为无响应，并且用户会看到错误消息。       

## <a name="see-also"></a>另请参阅

- [了解适用于 Office 的 JavaScript API](understanding-the-javascript-api-for-office.md)    
- [适用于 Office 的 JavaScript API](https://docs.microsoft.com/javascript/office/javascript-api-for-office?view=office-js)
    
